import fs from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputPath = fileURLToPath(new URL("./test_sales.xlsx", import.meta.url));
const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Продажи");
const headers = ["Дата", "Регион", "Магазин", "Продукт", "Категория", "Продажи", "Транзакция"];
const regions = ["Центр", "Северо-Запад", "Юг", "Урал"];
const products = [
  ["Кофе", "Напитки"], ["Чай", "Напитки"], ["Сок", "Напитки"],
  ["Шоколад", "Сладости"], ["Печенье", "Сладости"],
  ["Сэндвич", "Готовая еда"], ["Выпечка", "Готовая еда"],
];
const start = new Date("2025-07-23T00:00:00Z");
const rows = [];
for (let day = 0; day < 365; day += 1) {
  for (let transaction = 0; transaction < 12; transaction += 1) {
    const date = new Date(start);
    date.setUTCDate(start.getUTCDate() + day);
    const region = regions[(day + transaction) % regions.length];
    const storeNumber = ((transaction + day * 3) % 16) + 1;
    const [product, category] = products[(day * 2 + transaction) % products.length];
    const seasonality = 1 + Math.sin(day / 32) * 0.16;
    const sales = Math.round((850 + ((day * 71 + transaction * 193) % 5900)) * seasonality);
    rows.push([date, region, `Магазин ${String(storeNumber).padStart(2, "0")}`, product, category, sales, `TX-${String(day + 1).padStart(3, "0")}-${String(transaction + 1).padStart(2, "0")}`]);
  }
}

sheet.getRangeByIndexes(0, 0, rows.length + 1, headers.length).values = [headers, ...rows];
sheet.getRange(`A1:G${rows.length + 1}`).format.font = { name: "Aptos", size: 10 };
sheet.getRange("A1:G1").format = { fill: "#1D6B50", font: { bold: true, color: "#FFFFFF", name: "Aptos", size: 10 }, rowHeight: 24 };
sheet.getRange(`A2:A${rows.length + 1}`).format.numberFormat = "yyyy-mm-dd";
sheet.getRange(`F2:F${rows.length + 1}`).format.numberFormat = "#,##0\" ₽\"";
sheet.getRange(`A1:G${rows.length + 1}`).format.borders = { preset: "insideHorizontal", style: "thin", color: "#E6E2D8" };
sheet.getRange("A:G").format.columnWidth = 17;
sheet.getRange("D:E").format.columnWidth = 19;
sheet.getRange("G:G").format.columnWidth = 20;
sheet.freezePanes.freezeRows(1);
sheet.tables.add(`A1:G${rows.length + 1}`, true, "SalesTable").style = "TableStyleMedium4";
sheet.showGridLines = false;

const info = workbook.worksheets.add("Описание");
info.getRange("A1:D1").merge();
info.getRange("A1").values = [["Тестовые данные для дашборда продаж"]];
info.getRange("A1:D1").format = { fill: "#13231F", font: { bold: true, color: "#FFFFFF", size: 16 }, rowHeight: 34 };
info.getRange("A3:B8").values = [
  ["Период", "23.07.2025–22.07.2026"], ["Строк", rows.length], ["Регионов", regions.length],
  ["Магазинов", 16], ["Продуктов", products.length], ["Назначение", "Проверка импорта, KPI, фильтров и графиков"],
];
info.getRange("A3:A8").format = { fill: "#B9E6CD", font: { bold: true } };
info.getRange("A3:B8").format.borders = { preset: "insideHorizontal", style: "thin", color: "#DEDAD0" };
info.getRange("A:B").format.columnWidth = 34;
info.getRange("B8").format.wrapText = true;
info.showGridLines = false;

const inspection = await workbook.inspect({ kind: "table", range: "Продажи!A1:G6", include: "values,formulas", tableMaxRows: 6, tableMaxCols: 7 });
console.log(inspection.ndjson);
const errors = await workbook.inspect({ kind: "match", searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A", options: { useRegex: true, maxResults: 50 }, summary: "formula errors" });
console.log(errors.ndjson);
const preview = await workbook.render({ sheetName: "Продажи", range: "A1:G18", scale: 1.3, format: "png" });
await fs.writeFile("test_sales_preview.png", new Uint8Array(await preview.arrayBuffer()));
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`Saved ${rows.length} rows to ${outputPath}`);
