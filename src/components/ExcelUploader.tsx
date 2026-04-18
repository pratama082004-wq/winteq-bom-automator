"use client";
import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import { UploadCloud } from "lucide-react";

export interface BOMRow {
  NO?: number | string;
  "PART CODE": string;
  "PART NAME": string;
  "DETAIL NAME"?: string;
  SPESIFICATION?: string;
  BRAND?: string;
  QTY: number | string;
  UNT?: string;
  REMARK?: string;
}

export default function ExcelUploader({ onDataLoaded }: { onDataLoaded: (data: BOMRow[]) => void }) {
  const [loading, setLoading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setLoading(true);

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const buffer = event.target?.result;
        const workbook = XLSX.read(buffer, { type: "array" });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" }) as any[][];

        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(10, rows.length); i++) {
          if (rows[i].join("").toUpperCase().replace(/\s+/g, "").includes("PARTCODE")) {
            headerRowIndex = i; break;
          }
        }
        if (headerRowIndex === -1) throw new Error("Kolom PART CODE tidak ditemukan.");

        const headers = rows[headerRowIndex].map(h => String(h).toUpperCase().replace(/\s+/g, ""));
        const extractedData: BOMRow[] = [];

        for (let i = headerRowIndex + 1; i < rows.length; i++) {
          const row = rows[i];
          const getVal = (keyStr: string) => row[headers.findIndex(h => h.includes(keyStr))] || "";
          
          const partCode = String(getVal("PARTCODE")).trim();
          const partName = String(getVal("PARTNAME")).trim();

          if (partCode || partName) {
            extractedData.push({
              NO: 0,
              "PART CODE": partCode,
              "PART NAME": partName,
              "DETAIL NAME": getVal("DETAILNAME"),
              SPESIFICATION: getVal("SPESIFICATION") || getVal("SPECIFICATION"),
              BRAND: getVal("BRAND"),
              QTY: Number(getVal("QTY")) || 1,
              UNT: getVal("UNT") || "PCS",
              REMARK: getVal("REMARK")
            });
          }
        }
        onDataLoaded(extractedData);
      } catch (err) {
        alert("Gagal membaca file!");
      } finally {
        setLoading(false);
        if (fileInputRef.current) fileInputRef.current.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div>
      <input type="file" accept=".xlsx" className="hidden" ref={fileInputRef} onChange={handleFileUpload} />
      <button onClick={() => fileInputRef.current?.click()} disabled={loading} className="flex items-center gap-2 bg-white border border-slate-400 text-slate-700 px-3 py-1.5 text-xs font-bold hover:bg-slate-100 transition shadow-sm cursor-pointer">
        <UploadCloud size={14} /> {loading ? "Membaca..." : "Upload Data Visio"}
      </button>
    </div>
  );
}