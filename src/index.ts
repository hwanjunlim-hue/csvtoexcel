import { Worker } from "@notionhq/workers";
import { j } from "@notionhq/workers/schema-builder";
import ExcelJS from "exceljs";
import iconv from "iconv-lite";
import { createRequire } from "module";
import type { Client } from "@notionhq/client";

// CJS 모듈 로드
const require = createRequire(import.meta.url);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const XLSXChart = require("xlsx-chart") as any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const JSZip = require("jszip") as any;

const xlsxChart = new XLSXChart();
const XLSX_MIME =
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const worker = new Worker();
export default worker;

// ─────────────────────────────────────────────────────────────────────────────
// CSV 파싱
// ─────────────────────────────────────────────────────────────────────────────

const REQUIRED_COLS = [
	"전표번호", "회계일자", "계정코드", "계정과목",
	"차변금액", "대변금액", "거래처코드", "거래처명", "적요", "부서",
] as const;
type ColName = (typeof REQUIRED_COLS)[number];
type Row = Record<ColName, string> & {
	_date: Date; _code: number; _debit: number; _credit: number; _codeStr: string;
};

function decodeCsv(buf: Buffer): string {
	const utf8 = buf.toString("utf8");
	return utf8.includes("�") ? iconv.decode(buf, "cp949") : utf8;
}

function parseLine(line: string): string[] {
	const res: string[] = [];
	let cur = "", inQ = false;
	for (const ch of line) {
		if (ch === '"') inQ = !inQ;
		else if (ch === "," && !inQ) { res.push(cur.trim()); cur = ""; }
		else cur += ch;
	}
	res.push(cur.trim());
	return res;
}

function parseRows(buf: Buffer): Row[] {
	const text = decodeCsv(buf);
	const lines = text.split(/\r?\n/).filter((l) => l.trim());
	const headers = parseLine(lines[0]);
	const missing = REQUIRED_COLS.filter((c) => !headers.includes(c));
	if (missing.length) throw new Error(`필수 컬럼 없음: ${missing.join(", ")}`);

	return lines.slice(1).flatMap((line) => {
		const vals = parseLine(line);
		const raw = Object.fromEntries(headers.map((h, i) => [h, vals[i] ?? ""])) as Record<ColName, string>;
		const date = new Date(raw["회계일자"].trim());
		if (isNaN(date.getTime())) return [];
		const toAmt = (s: string) => { const n = parseInt(s.replace(/[,\s]/g, ""), 10); return isNaN(n) ? 0 : n; };
		return [{
			...raw,
			_date: date,
			_code: parseInt(raw["계정코드"].replace(/[,\s]/g, ""), 10) || 0,
			_debit: toAmt(raw["차변금액"]),
			_credit: toAmt(raw["대변금액"]),
			_codeStr: raw["계정코드"].trim(),
		}] satisfies Row[];
	});
}

// ─────────────────────────────────────────────────────────────────────────────
// 집계
// ─────────────────────────────────────────────────────────────────────────────

const isRev = (r: Row) => r._codeStr.startsWith("4");
const isExp = (r: Row) => r._codeStr.startsWith("5") || r._codeStr.startsWith("8");

function groupBy(rows: Row[], key: (r: Row) => string, val: (r: Row) => number): [string, number][] {
	const m = new Map<string, number>();
	for (const r of rows) m.set(key(r), (m.get(key(r)) ?? 0) + val(r));
	return [...m].sort((a, b) => b[1] - a[1]);
}

function detectMonth(rows: Row[]): string {
	const m = new Map<string, number>();
	for (const r of rows) {
		const k = `${r._date.getFullYear()}${String(r._date.getMonth() + 1).padStart(2, "0")}`;
		m.set(k, (m.get(k) ?? 0) + 1);
	}
	return [...m].sort((a, b) => b[1] - a[1])[0]?.[0] ?? new Date().toISOString().slice(0, 7).replace("-", "");
}

// ─────────────────────────────────────────────────────────────────────────────
// xlsx-chart 네이티브 차트 생성
// ─────────────────────────────────────────────────────────────────────────────

function xChart(type: string, label: string, entries: [string, number][]): Promise<Buffer> {
	const fields = entries.map(([k]) => k);
	const dataRow = Object.fromEntries(entries.map(([k, v]) => [k, v]));
	return new Promise((res, rej) =>
		xlsxChart.generate(
			{ chart: type, titles: [label], fields, data: { [label]: dataRow } },
			(e: Error | null, b: Buffer) => (e ? rej(e) : res(b)),
		),
	);
}

async function chartXmlFrom(buf: Buffer): Promise<string> {
	const zip = await JSZip.loadAsync(buf);
	return zip.files["xl/charts/chart1.xml"].async("string") as Promise<string>;
}

function addRedBar(xml: string): string {
	// 손익 막대(idx=2)에 빨간 채우기 주입
	const dPt = "<c:dPt><c:idx val=\"2\"/><c:invertIfNegative val=\"0\"/>" +
		"<c:spPr><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></c:spPr></c:dPt>";
	return xml.replace("</c:tx>", `</c:tx>${dPt}`);
}

function sanitize(xml: string): string {
	// 없는 시트("Table!") 참조 제거 → 캐시 데이터로 렌더링
	return xml.replace(/<c:f>[^<]*<\/c:f>/g, "");
}

// ─────────────────────────────────────────────────────────────────────────────
// XML 헬퍼 (drawing 주입용)
// ─────────────────────────────────────────────────────────────────────────────

const CHART_COL_S = 3;   // D열 (0-indexed)
const CHART_COL_E = 17;  // R열
const CHART_ROWS = 21;   // 차트 높이 (행 수)

function makeDrawingXml(count: number): string {
	const anchors = Array.from({ length: count }, (_, i) => {
		const rs = i * CHART_ROWS, re = rs + CHART_ROWS;
		return `<xdr:twoCellAnchor>` +
			`<xdr:from><xdr:col>${CHART_COL_S}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${rs}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
			`<xdr:to><xdr:col>${CHART_COL_E}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${re}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
			`<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>` +
			`<xdr:cNvPr id="${i + 1}" name="차트 ${i + 1}"/><xdr:cNvGraphicFramePr/>` +
			`</xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
			`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
			`<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${i + 1}"/>` +
			`</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
	}).join("");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ` +
		`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors}</xdr:wsDr>`;
}

function makeDrawingRels(count: number): string {
	const rels = Array.from({ length: count }, (_, i) =>
		`<Relationship Id="rId${i + 1}" ` +
		`Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" ` +
		`Target="../charts/chart${i + 1}.xml"/>`
	).join("");
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

function makeSheet2Rels(): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>` +
		`</Relationships>`;
}

function patchContentTypes(xml: string, count: number): string {
	const chartOvrs = Array.from({ length: count }, (_, i) =>
		`<Override PartName="/xl/charts/chart${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`
	).join("");
	const drawOvr = `<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;
	return xml.replace("</Types>", `${chartOvrs}${drawOvr}</Types>`);
}

// ─────────────────────────────────────────────────────────────────────────────
// 마스터 워크북 생성 (ExcelJS) + 차트 주입 (JSZip)
// ─────────────────────────────────────────────────────────────────────────────

const HDR_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } };
const SUB_FILL: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCE6F1" } };

function writeTable(ws: ExcelJS.Worksheet, r: number, title: string, headers: string[], data: [string, number][]): number {
	ws.getCell(r, 1).value = title;
	ws.getCell(r, 1).font = { bold: true, color: { argb: "FF1F3864" }, size: 11 };
	r++;
	headers.forEach((h, ci) => {
		const cell = ws.getCell(r, ci + 1);
		cell.value = h; cell.fill = SUB_FILL; cell.font = { bold: true };
	});
	r++;
	data.forEach(([label, amount]) => {
		ws.getCell(r, 1).value = label;
		const ac = ws.getCell(r, 2);
		ac.value = amount; ac.numFmt = "#,##0";
		r++;
	});
	return r + 1;
}

async function buildExcel(rows: Row[]): Promise<{ buffer: Buffer; stats: object }> {
	const rev = rows.filter(isRev);
	const exp = rows.filter(isExp);

	const revAcct = groupBy(rev, (r) => r["계정과목"], (r) => r._credit);
	const revDept = groupBy(rev, (r) => r["부서"], (r) => r._credit);
	const revVendor = groupBy(rev, (r) => r["거래처명"], (r) => r._credit).slice(0, 10);
	const expAcct = groupBy(exp, (r) => r["계정과목"], (r) => r._debit);
	const revDate = groupBy(rev, (r) => r._date.toISOString().slice(0, 10), (r) => r._credit)
		.sort((a, b) => a[0].localeCompare(b[0]));

	const totalRevenue = revAcct.reduce((s, [, v]) => s + v, 0);
	const totalExpense = expAcct.reduce((s, [, v]) => s + v, 0);
	const profit = totalRevenue - totalExpense;
	const month = detectMonth(rows);

	// 검증
	if (revDept.reduce((s, [, v]) => s + v, 0) !== totalRevenue) throw new Error("부서별 매출 합계 불일치");
	if (revDate.reduce((s, [, v]) => s + v, 0) !== totalRevenue) throw new Error("일자별 매출 합계 불일치");

	// ── 차트 생성 (xlsx-chart) ────────────────────────────────────────────
	type ChartDef = { type: string; label: string; entries: [string, number][]; postProcess?: (xml: string) => string };
	const pnlEntries: [string, number][] = [["총매출", totalRevenue], ["총비용", totalExpense], ["손익", profit]];

	const defs: (ChartDef | null)[] = [
		revAcct.length ? { type: "pie", label: "매출 구성", entries: revAcct } : null,
		revDept.length ? { type: "column", label: "부서별 매출", entries: revDept } : null,
		revVendor.length ? { type: "bar", label: "거래처별 매출 TOP10", entries: revVendor } : null,
		expAcct.length ? { type: "pie", label: "비용 구성", entries: expAcct } : null,
		{ type: "column", label: "매출 vs 비용(손익)", entries: pnlEntries, postProcess: profit < 0 ? addRedBar : undefined },
		revDate.length ? { type: "line", label: "일자별 매출 추이", entries: revDate } : null,
	];

	const created: string[] = [];
	const skipped: string[] = [];
	const skipLabels = ["매출 구성 - 4xxx 없음", "부서별 매출 - 부서 없음", "거래처별 TOP10 - 데이터 없음", "비용 구성 - 5/8xxx 없음", "", "일자별 추이 - 날짜 없음"];

	const chartXmls: string[] = [];
	await Promise.all(defs.map(async (def, i) => {
		if (!def) {
			if (skipLabels[i]) skipped.push(skipLabels[i]);
			return;
		}
		try {
			const buf = await xChart(def.type, def.label, def.entries);
			let xml = await chartXmlFrom(buf);
			xml = sanitize(xml);
			if (def.postProcess) xml = def.postProcess(xml);
			chartXmls.push(xml);
			created.push(`${def.label} (${def.type})`);
		} catch (e) {
			skipped.push(`${def.label} - 생성 실패: ${e instanceof Error ? e.message : e}`);
		}
	}));

	// ── 마스터 워크북 (ExcelJS) ───────────────────────────────────────────
	const wb = new ExcelJS.Workbook();

	// 원본데이터 시트
	const wsRaw = wb.addWorksheet("원본데이터");
	REQUIRED_COLS.forEach((col, ci) => {
		const cell = wsRaw.getCell(1, ci + 1);
		cell.value = col; cell.fill = HDR_FILL;
		cell.font = { color: { argb: "FFFFFFFF" }, bold: true };
	});
	rows.forEach((r, ri) => {
		REQUIRED_COLS.forEach((col, ci) => {
			const cell = wsRaw.getCell(ri + 2, ci + 1);
			if (col === "차변금액") { cell.value = r._debit; cell.numFmt = "#,##0"; }
			else if (col === "대변금액") { cell.value = r._credit; cell.numFmt = "#,##0"; }
			else if (col === "회계일자") { cell.value = r._date; cell.numFmt = "yyyy-mm-dd"; }
			else cell.value = r[col];
		});
	});
	wsRaw.columns.forEach((c) => { c.width = 15; });

	// 집계 시트 (데이터 테이블)
	const ws = wb.addWorksheet("집계");
	ws.getColumn(1).width = 24;
	ws.getColumn(2).width = 16;
	let tr = 1;
	tr = writeTable(ws, tr, "① 매출 구성", ["계정과목", "금액"], revAcct);
	tr = writeTable(ws, tr, "② 부서별 매출", ["부서", "금액"], revDept);
	tr = writeTable(ws, tr, "③ 거래처별 매출 TOP10", ["거래처명", "금액"], revVendor);
	tr = writeTable(ws, tr, "④ 비용 구성", ["계정과목", "금액"], expAcct);
	tr = writeTable(ws, tr, "⑤ 매출 vs 비용(손익)", ["항목", "금액"], pnlEntries);
	writeTable(ws, tr, "⑥ 일자별 매출 추이", ["일자", "금액"], revDate);

	// 마스터 워크북 버퍼
	const masterBuf = Buffer.from(await wb.xlsx.writeBuffer());

	// ── 차트 주입 (JSZip) ────────────────────────────────────────────────
	const masterZip = await JSZip.loadAsync(masterBuf);

	// chart XML 파일 추가
	chartXmls.forEach((xml, i) => masterZip.file(`xl/charts/chart${i + 1}.xml`, xml));

	// drawing XML + rels
	masterZip.file("xl/drawings/drawing1.xml", makeDrawingXml(chartXmls.length));
	masterZip.file("xl/drawings/_rels/drawing1.xml.rels", makeDrawingRels(chartXmls.length));

	// sheet2.xml에 <drawing> 요소 삽입
	const sheet2Xml = await masterZip.files["xl/worksheets/sheet2.xml"].async("string") as string;
	masterZip.file(
		"xl/worksheets/sheet2.xml",
		sheet2Xml.replace("</worksheet>", '<drawing r:id="rId1"/></worksheet>'),
	);

	// sheet2 relationships
	masterZip.file("xl/worksheets/_rels/sheet2.xml.rels", makeSheet2Rels());

	// [Content_Types].xml 업데이트
	const ctXml = await masterZip.files["[Content_Types].xml"].async("string") as string;
	masterZip.file("[Content_Types].xml", patchContentTypes(ctXml, chartXmls.length));

	const finalBuf = Buffer.from(
		await masterZip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" })
	);

	return {
		buffer: finalBuf,
		stats: { totalRevenue, totalExpense, profit, month, chartsCreated: created, chartsSkipped: skipped },
	};
}

// ─────────────────────────────────────────────────────────────────────────────
// Notion 유틸리티
// ─────────────────────────────────────────────────────────────────────────────

type NotionFileProp = {
	type: "files";
	files: Array<{ name: string; type: "file" | "external"; file?: { url: string }; external?: { url: string } }>;
};

function getFileUrl(prop: NotionFileProp | undefined): string | null {
	if (!prop?.files?.length) return null;
	const f = prop.files[0];
	return f.type === "file" ? (f.file?.url ?? null) : (f.external?.url ?? null);
}

function getPageTitle(properties: Record<string, unknown>): string {
	for (const val of Object.values(properties)) {
		const p = val as Record<string, unknown>;
		if (p?.type === "title" && Array.isArray(p.title) && p.title.length > 0)
			return (p.title as Array<{ plain_text?: string }>).map((t) => t.plain_text ?? "").join("").trim();
	}
	return "재무리포트";
}

async function postComment(notion: Client, pageId: string, text: string): Promise<void> {
	try {
		await notion.comments.create({ parent: { page_id: pageId }, rich_text: [{ type: "text", text: { content: text } }] });
	} catch (e) {
		console.error("[ERROR] 코멘트 작성 실패:", e);
	}
}

// ─────────────────────────────────────────────────────────────────────────────
// 자동화 정의
// ─────────────────────────────────────────────────────────────────────────────

worker.tool("csvToExcel", {
	title: "CSV → Excel 재무 리포트",
	description: '회계 전표 CSV를 분석하여 네이티브 차트 6종이 포함된 Excel 재무 리포트를 생성하고 "엑셀 파일" 속성에 첨부합니다.',
	schema: j.object({ pageId: j.string().describe("대상 Notion 페이지 ID") }),

	execute: async ({ pageId }, { notion }) => {

		const page = await notion.pages.retrieve({ page_id: pageId });
		const props = (page as { properties: Record<string, unknown> }).properties;

		const csvUrl = getFileUrl(props["CSV 파일"] as NotionFileProp | undefined);
		if (!csvUrl) {
			await postComment(notion, pageId, "⚠️ 오류: 'CSV 파일' 속성에 첨부 파일이 없습니다.");
			return "CSV 파일 없음";
		}

		const csvResp = await fetch(csvUrl);
		if (!csvResp.ok) throw new Error(`CSV 다운로드 실패: ${csvResp.status}`);
		const csvBuf = Buffer.from(await csvResp.arrayBuffer());
		const pageTitle = getPageTitle(props);

		try {
			const rows = parseRows(csvBuf);
			const { buffer: xlsxBuf, stats } = await buildExcel(rows);
			const s = stats as {
				totalRevenue: number; totalExpense: number; profit: number;
				month: string; chartsCreated: string[]; chartsSkipped: string[];
			};

			const safeTitle = pageTitle.replace(/[^\w가-힣]/g, "_").slice(0, 30);
			const filename = `${safeTitle}_재무리포트_${s.month}.xlsx`;

			const upload = await notion.fileUploads.create({ mode: "single_part", filename, content_type: XLSX_MIME });
			await notion.fileUploads.send({
				file_upload_id: upload.id,
				file: { data: new Blob([new Uint8Array(xlsxBuf)], { type: XLSX_MIME }), filename },
			});
			await (notion.pages.update as (args: unknown) => Promise<unknown>)({
				page_id: pageId,
				properties: {
					"엑셀 파일": { files: [{ type: "file_upload", name: filename, file_upload: { id: upload.id } }] },
				},
			});

			const fmt = (n: number) => n.toLocaleString("ko-KR");
			const summary = [
				`✅ Excel 재무 리포트 생성 완료: ${filename}`,
				`📊 총매출: ${fmt(s.totalRevenue)}원`,
				`📉 총비용: ${fmt(s.totalExpense)}원`,
				`💰 손익: ${fmt(s.profit)}원 ${s.profit < 0 ? "🔴 적자" : "🟢 흑자"}`,
				`📈 차트 (${s.chartsCreated.length}종): ${s.chartsCreated.join(" | ")}`,
				...(s.chartsSkipped.length > 0 ? [`⚠️ 생략: ${s.chartsSkipped.join(" | ")}`] : []),
			].join("\n");
			await postComment(notion, pageId, summary);
			console.log("[INFO]", summary);
			return summary;
		} catch (err: unknown) {
			const msg = err instanceof Error ? err.message : String(err);
			await postComment(notion, pageId, `⚠️ 처리 중 오류 발생:\n${msg}`);
			throw err;
		}
	},
});
