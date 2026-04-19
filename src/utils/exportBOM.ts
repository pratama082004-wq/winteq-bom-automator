import * as ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { BOMRow } from "../components/ExcelUploader";

export async function exportWinteqBOM(
  sheetData: Record<string, BOMRow[]>,
  projectInfoMap: Record<string, any>,
  tabHeaders: Record<string, string[]>,
  selectedTabs: string[],
  mode: "single" | "separate",
  headerLabelsMap: Record<string, any>
) {
  try {
    const templateUrl = "/templates/master_template_winteq.xlsx";

    const injectData = (sheet: ExcelJS.Worksheet, data: BOMRow[], tabName: string) => {
      // Ambil data spesifik untuk tab yang sedang diproses (atau fallback ke MAIN)
      const pInfo = projectInfoMap[tabName] || projectInfoMap["MAIN"] || {};
      const hLabels = headerLabelsMap[tabName] || headerLabelsMap["MAIN"] || {};

      // 1. SUNTIK LABEL HEADER DINAMIS
      sheet.getCell("F1").value = hLabels.machineName;
      sheet.getCell("F3").value = hLabels.designId;
      sheet.getCell("F4").value = hLabels.customer;
      sheet.getCell("H1").value = hLabels.designed;
      sheet.getCell("I1").value = hLabels.elc;
      sheet.getCell("J1").value = hLabels.checked;
      sheet.getCell("K1").value = hLabels.approved;
      sheet.getCell("C3").value = hLabels.title;
      sheet.getCell("C4").value = hLabels.title;

      // 2. SUNTIK ISI PROJECT INFO
      sheet.getCell("F2").value = pInfo.machineName;
      sheet.getCell("G3").value = pInfo.designId;
      sheet.getCell("G4").value = pInfo.customer;
      sheet.getCell("H2").value = pInfo.date1;
      sheet.getCell("I2").value = pInfo.date2;
      sheet.getCell("J2").value = pInfo.date3;
      sheet.getCell("K2").value = pInfo.date4;
      sheet.getCell("H4").value = pInfo.designed;
      sheet.getCell("I4").value = pInfo.elc;
      sheet.getCell("J4").value = pInfo.checked;
      sheet.getCell("K4").value = pInfo.approved;

      // 3. SUNTIK JUDUL KOLOM TABEL
      const headers = tabHeaders[tabName] || ["PART CODE", "PART NAME", "DETAIL NAME", "SPECIFICATION", "BRAND", "QTY", "UNT", "REMARK"];
      sheet.getCell("C6").value = headers[0];
      sheet.getCell("D6").value = headers[1];
      sheet.getCell("E6").value = headers[2];
      sheet.getCell("F6").value = headers[3];
      sheet.getCell("G6").value = headers[4];
      sheet.getCell("H6").value = headers[5];
      sheet.getCell("I6").value = headers[6];
      sheet.getCell("J6").value = headers[7];

      // 4. SUNTIK DATA BOM
      let startRow = 7;
      data.forEach((row) => {
        sheet.getCell(`A${startRow}`).value = row.NO;
        sheet.getCell(`C${startRow}`).value = row["PART CODE"];
        sheet.getCell(`D${startRow}`).value = row["PART NAME"];
        sheet.getCell(`E${startRow}`).value = row["DETAIL NAME"];
        sheet.getCell(`F${startRow}`).value = row.SPESIFICATION;
        sheet.getCell(`G${startRow}`).value = row.BRAND;
        sheet.getCell(`H${startRow}`).value = Number(row.QTY);
        sheet.getCell(`I${startRow}`).value = row.UNT;
        sheet.getCell(`J${startRow}`).value = row.REMARK;
        startRow++;
      });
    };

    if (mode === "separate") {
      for (const tab of selectedTabs) {
        const data = sheetData[tab] || [];
        if (data.length === 0) continue;
        const response = await fetch(templateUrl);
        const arrayBuffer = await response.arrayBuffer();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(arrayBuffer);
        const sheet = wb.getWorksheet(1);
        if (sheet) {
          sheet.name = tab;
          injectData(sheet, data, tab);
        }
        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const pInfo = projectInfoMap[tab] || projectInfoMap["MAIN"] || { machineName: "BOM" };
        saveAs(blob, `BOM_WINTEQ_${tab}_${(pInfo.machineName || "").replace(/\s+/g, '_')}.xlsx`);
      }
    } else {
      const response = await fetch(templateUrl);
      const arrayBuffer = await response.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(arrayBuffer);
      const masterSheet = wb.getWorksheet(1);
      for (let i = 0; i < selectedTabs.length; i++) {
        const tab = selectedTabs[i];
        const data = sheetData[tab] || [];
        if (data.length === 0) continue;
        let sheet = wb.getWorksheet(tab);
        if (!sheet) {
          if (i === 0 && masterSheet) {
            masterSheet.name = tab;
            sheet = masterSheet;
          } else {
            sheet = wb.addWorksheet(tab);
          }
        }
        if (sheet) injectData(sheet, data, tab);
      }
      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const pInfoMain = projectInfoMap["MAIN"] || { machineName: "BOM" };
      saveAs(blob, `BOM_WINTEQ_FULL_${(pInfoMain.machineName || "").replace(/\s+/g, '_')}.xlsx`);
    }
    return true;
  } catch (error: any) {
    console.error("CRITICAL EXPORT ERROR:", error);
    alert("Gagal download: " + error.message);
    return false;
  }
}