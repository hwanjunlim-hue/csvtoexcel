/**
 * Excel 변환 대기 중인 Notion 페이지 ID를 stdout에 출력 (1줄 1 ID)
 * 조건: 상태 = "Excel" AND 엑셀 파일 속성이 비어 있음
 *
 * 사용: npx tsx src/scan.ts
 */
import dotenv from "dotenv";

dotenv.config();

const DB_ID = "38ea3147-c621-80d0-b171-ccda5918becd";
const TOKEN = process.env.NOTION_API_TOKEN;
if (!TOKEN) { console.error("NOTION_API_TOKEN 없음"); process.exit(1); }

const HEADERS = {
	Authorization: `Bearer ${TOKEN}`,
	"Notion-Version": "2022-06-28",
	"Content-Type": "application/json",
};

type Page = {
	id: string;
	properties: Record<string, { type: string; files?: unknown[] }>;
};

type QueryResponse = {
	results: Page[];
	has_more: boolean;
	next_cursor: string | null;
};

let cursor: string | null = null;
do {
	const body: Record<string, unknown> = {
		filter: { property: "상태", status: { equals: "Excel" } },
		page_size: 100,
	};
	if (cursor) body.start_cursor = cursor;

	const res = await fetch(`https://api.notion.com/v1/databases/${DB_ID}/query`, {
		method: "POST",
		headers: HEADERS,
		body: JSON.stringify(body),
	});

	if (!res.ok) {
		console.error(`Notion API 오류: ${res.status} ${await res.text()}`);
		process.exit(1);
	}

	const data = (await res.json()) as QueryResponse;

	for (const page of data.results) {
		const excelProp = page.properties["엑셀 파일"];
		if (!excelProp?.files?.length) {
			process.stdout.write(page.id + "\n");
		}
	}

	cursor = data.has_more ? data.next_cursor : null;
} while (cursor);
