import * as ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { BOMRow } from "../components/ExcelUploader";

export async function exportWinteqBOM(
  sheetData: Record<string, BOMRow[]>,
  projectInfo: any,
  tabHeaders: Record<string, string[]>,
  selectedTabs: string[],
  mode: "single" | "separate"
) {
  try {
    const templateUrl = "/templates/master_template_winteq.xlsx";

    // Fungsi kecil untuk menyuntik data ke sheet manapun
    const injectData = (sheet: ExcelJS.Worksheet, data: BOMRow[], tabName: string) => {
      sheet.getCell("F2").value = projectInfo.machineName;
      sheet.getCell("G3").value = projectInfo.designId;
      sheet.getCell("G4").value = projectInfo.customer;
      
      sheet.getCell("H2").value = projectInfo.date1;
      sheet.getCell("I2").value = projectInfo.date2;
      sheet.getCell("J2").value = projectInfo.date3;
      sheet.getCell("K2").value = projectInfo.date4;

      sheet.getCell("H4").value = projectInfo.designed;
      sheet.getCell("I4").value = projectInfo.elc;
      sheet.getCell("J4").value = projectInfo.checked;
      sheet.getCell("K4").value = projectInfo.approved;

      const headers = tabHeaders[tabName] || ["PART CODE", "PART NAME", "DETAIL NAME", "SPECIFICATION", "BRAND", "QTY", "UNT", "REMARK"];
      sheet.getCell("C6").value = headers[0];
      sheet.getCell("D6").value = headers[1];
      sheet.getCell("E6").value = headers[2];
      sheet.getCell("F6").value = headers[3];
      sheet.getCell("G6").value = headers[4];
      sheet.getCell("H6").value = headers[5];
      sheet.getCell("I6").value = headers[6];
      sheet.getCell("J6").value = headers[7];

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
      // MODE TERPISAH: Download banyak file Excel
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
        saveAs(blob, `BOM_WINTEQ_${tab}_${projectInfo.machineName.replace(/\s+/g, '_')}.xlsx`);
      }
    } else {
      // MODE SATU FILE: Semua sheet digabung
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
            // Sheet kedua dan seterusnya ditambah secara manual
            sheet = wb.addWorksheet(tab);
          }
        }

        if (sheet) injectData(sheet, data, tab);
      }

      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      saveAs(blob, `BOM_WINTEQ_FULL_${projectInfo.machineName.replace(/\s+/g, '_')}.xlsx`);
    }

    return true;
  } catch (error: any) {
    console.error("CRITICAL EXPORT ERROR:", error);
    alert("Gagal download: " + error.message);
    return false;
  }
}