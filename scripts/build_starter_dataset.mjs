import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputDir = path.join(projectRoot, "outputs", "knime-baslangic");
const outputPath = `${outputDir}/mini_satislar.xlsx`;
const previewDir = `${outputDir}/previews`;
const fontFamily = "Arial";

const cities = ["İstanbul", "Ankara", "İzmir", "Bursa", "Antalya"];
const customerTypes = ["Bireysel", "Kurumsal"];
const channels = ["Web", "Mağaza", "Pazaryeri"];
const catalog = [
  ["Teknoloji", "Kablosuz Kulaklık", 1450],
  ["Teknoloji", "Mekanik Klavye", 2100],
  ["Teknoloji", "Monitör", 6800],
  ["Teknoloji", "Webcam", 1800],
  ["Ofis", "Not Defteri", 120],
  ["Ofis", "Masa Lambası", 950],
  ["Ofis", "Ergonomik Sandalye", 6200],
  ["Ofis", "Kalem Seti", 180],
  ["Ev & Yaşam", "Kahve Makinesi", 4200],
  ["Ev & Yaşam", "Termos", 780],
  ["Ev & Yaşam", "Hava Temizleyici", 7500],
  ["Ev & Yaşam", "Masa Organizer", 420],
];
const discounts = [0, 0.05, 0.1, 0.15];
const blankDiscountIndexes = new Set([5, 18, 31, 42]);

const rows = [];
for (let i = 0; i < 48; i += 1) {
  const [category, product, price] = catalog[(i * 5 + 2) % catalog.length];
  const orderDate = new Date(Date.UTC(2026, 0, 2 + i * 2));
  const status = i % 13 === 0 ? "İptal" : i % 10 === 0 ? "İade" : "Tamamlandı";
  rows.push([
    `S-${1001 + i}`,
    orderDate,
    cities[(i * 2 + 1) % cities.length],
    category,
    product,
    1 + ((i * 3) % 5),
    price,
    blankDiscountIndexes.has(i) ? null : discounts[i % discounts.length],
    customerTypes[i % customerTypes.length],
    channels[(i + 1) % channels.length],
    status,
  ]);
}

// KNIME'da Duplicate Row Filter pratiği için bir satır bilinçli olarak yinelenir.
rows.push([...rows[21]]);

const workbook = Workbook.create();
const dataSheet = workbook.worksheets.add("Siparisler");
const dictionarySheet = workbook.worksheets.add("Veri_Sozlugu");

const headers = [
  "siparis_id",
  "tarih",
  "sehir",
  "kategori",
  "urun",
  "adet",
  "birim_fiyat",
  "indirim_orani",
  "musteri_tipi",
  "kanal",
  "durum",
];

dataSheet.getRange("A1:K1").values = [headers];
dataSheet.getRange(`A2:K${rows.length + 1}`).values = rows;
dataSheet.showGridLines = false;
dataSheet.freezePanes.freezeRows(1);
dataSheet.getRange(`A1:K${rows.length + 1}`).format.font = { name: fontFamily, size: 10, color: "#1F2937" };
dataSheet.getRange("A1:K1").format = {
  fill: "#1F4E78",
  font: { name: fontFamily, size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  wrapText: false,
  borders: { preset: "outside", style: "thin", color: "#163A5A" },
};
dataSheet.getRange("B2:B50").setNumberFormat("yyyy-mm-dd");
dataSheet.getRange("F2:F50").setNumberFormat("0");
dataSheet.getRange("G2:G50").setNumberFormat("#,##0 \"₺\"");
dataSheet.getRange("H2:H50").setNumberFormat("0.0%");
dataSheet.getRange("A2:A50").format.horizontalAlignment = "left";
dataSheet.getRange("B2:B50").format.horizontalAlignment = "center";
dataSheet.getRange("F2:H50").format.horizontalAlignment = "right";
dataSheet.getRange("A1:K50").format.verticalAlignment = "center";
dataSheet.getRange("A1:K50").format.autofitColumns();
dataSheet.getRange("A1:K50").format.autofitRows();
dataSheet.getRange("A:A").format.columnWidth = 14;
dataSheet.getRange("B:B").format.columnWidth = 13;
dataSheet.getRange("C:C").format.columnWidth = 13;
dataSheet.getRange("D:D").format.columnWidth = 15;
dataSheet.getRange("E:E").format.columnWidth = 23;
dataSheet.getRange("F:F").format.columnWidth = 9;
dataSheet.getRange("G:G").format.columnWidth = 15;
dataSheet.getRange("H:H").format.columnWidth = 16;
dataSheet.getRange("I:I").format.columnWidth = 16;
dataSheet.getRange("J:J").format.columnWidth = 14;
dataSheet.getRange("K:K").format.columnWidth = 15;
const ordersTable = dataSheet.tables.add(`A1:K${rows.length + 1}`, true, "SiparislerTable");
ordersTable.style = "TableStyleMedium2";

const dictionaryRows = [
  ["siparis_id", "Siparişin benzersiz kimliği", "Metin"],
  ["tarih", "Sipariş tarihi", "Tarih"],
  ["sehir", "Teslimat şehri", "Metin"],
  ["kategori", "Ürün kategorisi", "Metin"],
  ["urun", "Ürün adı", "Metin"],
  ["adet", "Satılan ürün adedi", "Tam sayı"],
  ["birim_fiyat", "Ürünün birim fiyatı (TL)", "Sayı"],
  ["indirim_orani", "Sipariş indirimi; bazı hücreler eğitim amacıyla boştur", "Ondalık"],
  ["musteri_tipi", "Bireysel veya kurumsal müşteri", "Metin"],
  ["kanal", "Satış kanalı", "Metin"],
  ["durum", "Tamamlandı, İade veya İptal", "Metin"],
];

dictionarySheet.getRange("A1:C1").values = [["kolon", "aciklama", "veri_tipi"]];
dictionarySheet.getRange("A2:C12").values = dictionaryRows;
dictionarySheet.getRange("A15:B18").values = [
  ["Eğitim notu", "Bu veri seti sentetiktir ve yalnızca KNIME pratiği içindir."],
  ["Ham satır sayısı", rows.length],
  ["Bilinçli veri sorunu", "4 boş indirim_orani değeri"],
  ["Bilinçli veri sorunu", "1 yinelenmiş sipariş satırı"],
];
dictionarySheet.showGridLines = false;
dictionarySheet.getRange("A1:C18").format.font = { name: fontFamily, size: 10, color: "#1F2937" };
dictionarySheet.getRange("A1:C1").format = {
  fill: "#1F4E78",
  font: { name: fontFamily, size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
};
dictionarySheet.getRange("A15:B15").format = {
  fill: "#D9EAF7",
  font: { name: fontFamily, size: 10, bold: true, color: "#1F2937" },
};
dictionarySheet.getRange("A16:A18").format.font.bold = true;
dictionarySheet.getRange("A1:C12").format.borders = { preset: "outside", style: "thin", color: "#C7D2E0" };
dictionarySheet.getRange("A15:B18").format.borders = { preset: "outside", style: "thin", color: "#C7D2E0" };
dictionarySheet.getRange("A1:C18").format.verticalAlignment = "center";
dictionarySheet.getRange("A1:C18").format.autofitRows();
dictionarySheet.getRange("A:A").format.columnWidth = 20;
dictionarySheet.getRange("B:B").format.columnWidth = 62;
dictionarySheet.getRange("C:C").format.columnWidth = 16;

const uniqueRows = rows.slice(0, 48);
const completedRows = uniqueRows.filter((row) => row[10] === "Tamamlandı");
const netRevenue = completedRows.reduce((sum, row) => {
  const discount = row[7] ?? 0;
  return sum + row[5] * row[6] * (1 - discount);
}, 0);

workbook.recalculate();

const dataCheck = await workbook.inspect({
  kind: "table",
  sheetId: "Siparisler",
  range: "A1:K12",
  include: "values,formulas",
  tableMaxRows: 12,
  tableMaxCols: 11,
});
console.log(dataCheck.ndjson);

const errorCheck = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errorCheck.ndjson);

await fs.mkdir(previewDir, { recursive: true });
for (const [sheetName, range, fileName] of [
  ["Siparisler", "A1:K18", "siparisler.png"],
  ["Veri_Sozlugu", "A1:C18", "veri_sozlugu.png"],
]) {
  const preview = await workbook.render({ sheetName, range, scale: 1.4, format: "png" });
  await fs.writeFile(`${previewDir}/${fileName}`, new Uint8Array(await preview.arrayBuffer()));
}

await fs.mkdir(outputDir, { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);

console.log(JSON.stringify({
  outputPath,
  rawRows: rows.length,
  uniqueRows: uniqueRows.length,
  completedRows: completedRows.length,
  netRevenue,
  missingDiscounts: uniqueRows.filter((row) => row[7] === null).length,
}, null, 2));
