"use client";

import { useState, useRef, useEffect } from "react";
import { Download, AlertTriangle, Plus, Trash2, X, Grid, Info, ChevronDown, Undo2, Redo2, Eraser } from "lucide-react";
import ExcelUploader, { BOMRow } from "../components/ExcelUploader";
import { exportWinteqBOM } from "../utils/exportBOM";

const DEFAULT_HEADERS = ["PART CODE", "PART NAME", "DETAIL NAME", "SPECIFICATION", "BRAND", "QTY", "UNT", "REMARK"];
const COLUMNS: (keyof BOMRow)[] = ["PART CODE", "PART NAME", "DETAIL NAME", "SPESIFICATION", "BRAND", "QTY", "UNT", "REMARK"];

export default function Home() {
  const [tabs, setTabs] = useState(["MAIN", "PNEUMATIC", "ACCESSORIES"]);
  const [activeTab, setActiveTab] = useState("MAIN");
  const [sheetData, setSheetData] = useState<Record<string, BOMRow[]>>({
    "MAIN": [], "PNEUMATIC": [], "ACCESSORIES": []
  });

  const [tabHeaders, setTabHeaders] = useState<Record<string, string[]>>({
    "MAIN": [...DEFAULT_HEADERS],
    "PNEUMATIC": [...DEFAULT_HEADERS],
    "ACCESSORIES": ["SERIES", "PART NAME", "DETAIL NAME", "SPECIFICATION", "BRAND", "QTY", "UNT", "REMARK"]
  });

  const [projectInfo, setProjectInfo] = useState({
    machineName: "AUTO FAN COMP COOLING ASSY",
    designId: "612469", customer: "ASKI",
    designed: "AGH", elc: "HLM", checked: "DBY", approved: "YSF",
    date1: "Date.", date2: "Date.", date3: "Date.", date4: "Date."
  });

  const [isLoaded, setIsLoaded] = useState(false);
  const [isThickGrid, setIsThickGrid] = useState(false);

  // --- 1. STATE MEMORI & UNDO/REDO ---
  const [history, setHistory] = useState<string[]>([]);
  const [historyIndex, setHistoryIndex] = useState(-1);

  useEffect(() => {
    const savedTabs = localStorage.getItem("winteq_tabs");
    if (savedTabs) setTabs(JSON.parse(savedTabs));
    const savedActiveTab = localStorage.getItem("winteq_activeTab");
    if (savedActiveTab) setActiveTab(savedActiveTab);
    const savedSheetData = localStorage.getItem("winteq_sheetData");
    if (savedSheetData) {
      const parsedData = JSON.parse(savedSheetData);
      setSheetData(parsedData);
      setHistory([JSON.stringify(parsedData)]);
      setHistoryIndex(0);
    } else {
      setHistory([JSON.stringify({ "MAIN": [], "PNEUMATIC": [], "ACCESSORIES": [] })]);
      setHistoryIndex(0);
    }
    const savedHeaders = localStorage.getItem("winteq_headers");
    if (savedHeaders) setTabHeaders(JSON.parse(savedHeaders));
    const savedProject = localStorage.getItem("winteq_projectInfo");
    if (savedProject) setProjectInfo(JSON.parse(savedProject));
    setIsLoaded(true);
  }, []);

  useEffect(() => {
    if (isLoaded) {
      localStorage.setItem("winteq_tabs", JSON.stringify(tabs));
      localStorage.setItem("winteq_activeTab", activeTab);
      localStorage.setItem("winteq_sheetData", JSON.stringify(sheetData));
      localStorage.setItem("winteq_headers", JSON.stringify(tabHeaders));
      localStorage.setItem("winteq_projectInfo", JSON.stringify(projectInfo));
    }
  }, [tabs, activeTab, sheetData, tabHeaders, projectInfo, isLoaded]);

  const commitHistory = (newData: Record<string, BOMRow[]>) => {
    const strData = JSON.stringify(newData);
    if (historyIndex >= 0 && history[historyIndex] === strData) return; 
    const newHistory = history.slice(0, historyIndex + 1);
    newHistory.push(strData);
    if (newHistory.length > 50) newHistory.shift(); 
    setHistory(newHistory);
    setHistoryIndex(newHistory.length - 1);
  };

  const handleUndo = () => {
    if (historyIndex > 0) {
      const prevIdx = historyIndex - 1;
      setHistoryIndex(prevIdx);
      setSheetData(JSON.parse(history[prevIdx]));
    }
  };

  const handleRedo = () => {
    if (historyIndex < history.length - 1) {
      const nextIdx = historyIndex + 1;
      setHistoryIndex(nextIdx);
      setSheetData(JSON.parse(history[nextIdx]));
    }
  };

  // --- 2. STATE DRAG-TO-SELECT & MODAL ---
  const [selection, setSelection] = useState<{ start: { r: number; c: number }; end: { r: number; c: number } } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  
  const [conflictModal, setConflictModal] = useState<{ show: boolean, conflicts: BOMRow[], newCleanData: BOMRow[] }>({
    show: false, conflicts: [], newCleanData: []
  });
  const [isExportMenuOpen, setIsExportMenuOpen] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [exportOptions, setExportOptions] = useState<{ selectedTabs: string[]; mode: "single" | "separate" }>({
    selectedTabs: ["MAIN", "PNEUMATIC", "ACCESSORIES"], mode: "separate"
  });
  const dropdownRef = useRef<HTMLDivElement>(null);

  const [customDialog, setCustomDialog] = useState<{
    show: boolean; type: "alert" | "confirm" | "prompt"; title: string; message: string;
    inputValue?: string; onConfirm?: any;
  }>({ show: false, type: "alert", title: "", message: "" });

  const showAlert = (title: string, message: string) => setCustomDialog({ show: true, type: "alert", title, message });
  const showConfirm = (title: string, message: string, onConfirm: () => void) => setCustomDialog({ show: true, type: "confirm", title, message, onConfirm });
  const showPrompt = (title: string, message: string, defaultVal: string, onConfirm: (val: string) => void) => setCustomDialog({ show: true, type: "prompt", title, message, inputValue: defaultVal, onConfirm });
  const closeDialog = () => setCustomDialog({ ...customDialog, show: false });


  // --- 3. SENSOR GLOBAL: SHORTCUT KEYBOARD (ENTER & ESCAPE) ---
  const shortcutsRef = useRef({ customDialog, conflictModal, isExportMenuOpen, closeDialog, setConflictModal, setIsExportMenuOpen, sheetData, activeTab, updateSheetData: (t: string, d: BOMRow[]) => {} });
  
  useEffect(() => {
    shortcutsRef.current = { customDialog, conflictModal, isExportMenuOpen, closeDialog, setConflictModal, setIsExportMenuOpen, sheetData, activeTab, updateSheetData };
  });

  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      const { customDialog, conflictModal, isExportMenuOpen, closeDialog, setConflictModal, setIsExportMenuOpen, sheetData, activeTab, updateSheetData } = shortcutsRef.current;
      
      if (customDialog.show) {
        if (e.key === "Escape") closeDialog();
        else if (e.key === "Enter") {
          e.preventDefault();
          if (customDialog.onConfirm) customDialog.onConfirm(customDialog.inputValue || "");
          closeDialog();
        }
        return; 
      }

      if (conflictModal.show) {
        if (e.key === "Escape") {
          const currentData = sheetData[activeTab] || [];
          updateSheetData(activeTab, [...currentData, ...conflictModal.newCleanData]);
          setConflictModal({ show: false, conflicts: [], newCleanData: [] });
        } else if (e.key === "Enter") {
          const currentData = sheetData[activeTab] || [];
          const conflictCodes = conflictModal.conflicts.map(c => c["PART CODE"]);
          let updatedData = currentData.filter(item => !conflictCodes.includes(item["PART CODE"]));
          updatedData = [...updatedData, ...conflictModal.newCleanData, ...conflictModal.conflicts];
          updateSheetData(activeTab, updatedData);
          setConflictModal({ show: false, conflicts: [], newCleanData: [] });
        }
        return;
      }

      if (isExportMenuOpen && e.key === "Escape") setIsExportMenuOpen(false);
    };

    window.addEventListener("keydown", handleGlobalKeyDown);
    return () => window.removeEventListener("keydown", handleGlobalKeyDown);
  }, []);

  // --- FUNGSI MOUSE DRAG ---
  useEffect(() => {
    const handleGlobalMouseUp = () => setIsDragging(false);
    const handleClickOutside = (e: MouseEvent) => {
      if (!(e.target as HTMLElement).closest("table")) setSelection(null);
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) setIsExportMenuOpen(false);
    };
    window.addEventListener("mouseup", handleGlobalMouseUp);
    window.addEventListener("mousedown", handleClickOutside);
    return () => {
      window.removeEventListener("mouseup", handleGlobalMouseUp);
      window.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  // --- FUNGSI COPY/PASTE GLOBAL ---
  useEffect(() => {
    const handleCopy = (e: ClipboardEvent) => {
      if (!selection) return;
      const activeEl = document.activeElement as HTMLInputElement;
      if (activeEl?.tagName === "INPUT" && activeEl.selectionStart !== activeEl.selectionEnd) return; 

      e.preventDefault();
      const minR = Math.min(selection.start.r, selection.end.r);
      const maxR = Math.max(selection.start.r, selection.end.r);
      const minC = Math.min(selection.start.c, selection.end.c);
      const maxC = Math.max(selection.start.c, selection.end.c);
      const currentData = sheetData[activeTab] || [];
      
      let tsv = "";
      for (let r = minR; r <= maxR; r++) {
        let rowStr = [];
        for (let c = minC; c <= maxC; c++) {
          rowStr.push(currentData[r] ? currentData[r][COLUMNS[c]] || "" : "");
        }
        tsv += rowStr.join("\t") + "\n";
      }
      e.clipboardData?.setData("text/plain", tsv);
    };

    const handlePaste = (e: ClipboardEvent) => {
      const text = e.clipboardData?.getData("text/plain");
      if (!text || !text.includes("\t")) return;
      if (!selection) return;

      e.preventDefault();
      const rows = text.split(/\r?\n/).filter(r => r.trim() !== "");
      const currentData = [...(sheetData[activeTab] || [])];
      const startRowIdx = Math.min(selection.start.r, selection.end.r);
      const startColIdx = Math.min(selection.start.c, selection.end.c);

      rows.forEach((rowString, rOffset) => {
        const targetRowIdx = startRowIdx + rOffset;
        if (!currentData[targetRowIdx]) {
          currentData.push({ NO: currentData.length + 1, "PART CODE": "", "PART NAME": "", QTY: 1, UNT: "PCS" });
        }
        const cells = rowString.split("\t");
        cells.forEach((cellVal, cOffset) => {
          const targetColIdx = startColIdx + cOffset;
          if (targetColIdx < COLUMNS.length) {
            currentData[targetRowIdx][COLUMNS[targetColIdx]] = cellVal.trim();
          }
        });
      });
      const newData = { ...sheetData, [activeTab]: currentData.map((item, idx) => ({ ...item, NO: idx + 1 })) };
      setSheetData(newData);
      commitHistory(newData);
    };

    document.addEventListener("copy", handleCopy);
    document.addEventListener("paste", handlePaste);
    return () => {
      document.removeEventListener("copy", handleCopy);
      document.removeEventListener("paste", handlePaste);
    };
  }, [selection, sheetData, activeTab]);


  // --- FUNGSI UTAMA ---
  const handleDataLoaded = (newRawData: BOMRow[]) => {
    const currentData = sheetData[activeTab] || [];
    const conflicts: BOMRow[] = [];
    const cleanNewData: BOMRow[] = [];

    newRawData.forEach(row => {
      const exists = currentData.some(existing => existing["PART CODE"] === row["PART CODE"] && row["PART CODE"] !== "");
      if (exists) conflicts.push(row);
      else cleanNewData.push(row);
    });

    if (conflicts.length > 0) {
      setConflictModal({ show: true, conflicts, newCleanData: cleanNewData });
    } else {
      updateSheetData(activeTab, [...currentData, ...cleanNewData]);
    }
  };

  const resolveConflict = (action: "CANCEL" | "OVERWRITE" | "APPEND") => {
    const currentData = sheetData[activeTab] || [];
    let updatedData = [...currentData];
    if (action === "CANCEL") updatedData = [...updatedData, ...conflictModal.newCleanData];
    else if (action === "APPEND") updatedData = [...updatedData, ...conflictModal.newCleanData, ...conflictModal.conflicts];
    else if (action === "OVERWRITE") {
      const conflictCodes = conflictModal.conflicts.map(c => c["PART CODE"]);
      updatedData = updatedData.filter(item => !conflictCodes.includes(item["PART CODE"]));
      updatedData = [...updatedData, ...conflictModal.newCleanData, ...conflictModal.conflicts];
    }
    updateSheetData(activeTab, updatedData);
    setConflictModal({ show: false, conflicts: [], newCleanData: [] });
  };

  const updateSheetData = (tabName: string, data: BOMRow[]) => {
    const renumbered = data.map((item, idx) => ({ ...item, NO: idx + 1 }));
    const newData = { ...sheetData, [tabName]: renumbered };
    setSheetData(newData);
    commitHistory(newData);
  };

  const handleCellChange = (rowIndex: number, field: keyof BOMRow, value: string) => {
    const currentData = [...(sheetData[activeTab] || [])];
    currentData[rowIndex] = { ...currentData[rowIndex], [field]: value };
    setSheetData(prev => ({ ...prev, [activeTab]: currentData }));
  };

  const handleHeaderChange = (colIndex: number, newValue: string) => {
    setTabHeaders(prev => {
      const current = prev[activeTab] || [...DEFAULT_HEADERS];
      const updated = [...current];
      updated[colIndex] = newValue.toUpperCase();
      return { ...prev, [activeTab]: updated };
    });
  };

  const addRow = () => {
    const currentData = [...(sheetData[activeTab] || [])];
    currentData.push({ NO: currentData.length + 1, "PART CODE": "", "PART NAME": "", QTY: 1, UNT: "PCS" });
    updateSheetData(activeTab, currentData);
  };

  const deleteRow = (rowIndex: number) => {
    const currentData = [...(sheetData[activeTab] || [])];
    currentData.splice(rowIndex, 1);
    updateSheetData(activeTab, currentData);
  };

  const clearTable = () => {
    showConfirm("Kosongkan Tabel", `Yakin ingin menghapus SELURUH data di sheet ${activeTab}? (Bisa di-Undo)`, () => {
      updateSheetData(activeTab, []);
    });
  };

  const addNewTab = () => {
    showPrompt("Sheet Baru", "Masukkan nama untuk Sheet baru:", "", (newTabName: string) => {
      if (!newTabName) return;
      const name = newTabName.toUpperCase();
      if (!tabs.includes(name)) {
        setTabs([...tabs, name]);
        setSheetData({ ...sheetData, [name]: [] });
        setTabHeaders({ ...tabHeaders, [name]: [...DEFAULT_HEADERS] });
        setExportOptions(prev => ({ ...prev, selectedTabs: [...prev.selectedTabs, name] })); 
        setActiveTab(name);
      } else showAlert("Peringatan", "Nama sheet sudah ada!");
    });
  };

  const renameTab = (oldName: string) => {
    showPrompt("Ubah Nama Sheet", `Masukkan nama baru untuk ${oldName}:`, oldName, (newName: string) => {
      if (!newName) return;
      const name = newName.toUpperCase();
      if (name !== oldName && !tabs.includes(name)) {
        setTabs(tabs.map(t => t === oldName ? name : t));
        setSheetData(prev => { const newData = { ...prev, [name]: prev[oldName] || [] }; delete newData[oldName]; return newData; });
        setTabHeaders(prev => { const newHeaders = { ...prev, [name]: prev[oldName] || [...DEFAULT_HEADERS] }; delete newHeaders[oldName]; return newHeaders; });
        setExportOptions(prev => ({ ...prev, selectedTabs: prev.selectedTabs.map(t => t === oldName ? name : t) }));
        if (activeTab === oldName) setActiveTab(name);
      }
    });
  };

  const deleteTab = (tabName: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (tabs.length === 1) return showAlert("Dilarang", "Minimal harus ada 1 sheet!");
    showConfirm("Hapus Sheet", `Yakin ingin menghapus sheet ${tabName}? Semua datanya akan musnah.`, () => {
      const newTabs = tabs.filter(t => t !== tabName);
      setTabs(newTabs);
      setSheetData(prev => { const newData = { ...prev }; delete newData[tabName]; return newData; });
      setExportOptions(prev => ({ ...prev, selectedTabs: prev.selectedTabs.filter(t => t !== tabName) }));
      if (activeTab === tabName) setActiveTab(newTabs[0]);
    });
  };

  const executeDownload = async () => {
    if (exportOptions.selectedTabs.length === 0) return showAlert("Peringatan", "Pilih minimal 1 tab untuk di-download.");
    const hasData = exportOptions.selectedTabs.some(tab => (sheetData[tab] || []).length > 0);
    if (!hasData) return showAlert("Peringatan", "Sheet yang Anda pilih masih kosong semua!");
    setIsExporting(true);
    await exportWinteqBOM(sheetData, projectInfo, tabHeaders, exportOptions.selectedTabs, exportOptions.mode);
    setIsExporting(false);
    setIsExportMenuOpen(false); 
  };


  const headerBorder = isThickGrid ? "border-2 border-slate-800" : "border border-slate-300";
  const cellBorder = isThickGrid ? "border-2 border-slate-800" : "border-r border-b border-slate-200";
  const currentHeaders = tabHeaders[activeTab] || DEFAULT_HEADERS;

  // --- KOMPONEN SEL DENGAN ARROW KEY NAVIGASI ---
  const renderCell = (rowIdx: number, colIdx: number, field: keyof BOMRow, isCenter = false, isBold = false) => {
    const value = sheetData[activeTab]?.[rowIdx]?.[field] || "";
    
    let isSelected = false;
    if (selection) {
      const minR = Math.min(selection.start.r, selection.end.r);
      const maxR = Math.max(selection.start.r, selection.end.r);
      const minC = Math.min(selection.start.c, selection.end.c);
      const maxC = Math.max(selection.start.c, selection.end.c);
      isSelected = rowIdx >= minR && rowIdx <= maxR && colIdx >= minC && colIdx <= maxC;
    }

    const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
      const target = e.currentTarget;
      if (e.key === "ArrowUp") {
        e.preventDefault();
        document.getElementById(`cell-${rowIdx - 1}-${colIdx}`)?.focus();
      } else if (e.key === "ArrowDown") {
        e.preventDefault();
        document.getElementById(`cell-${rowIdx + 1}-${colIdx}`)?.focus();
      } else if (e.key === "ArrowLeft" && target.selectionStart === 0) {
        document.getElementById(`cell-${rowIdx}-${colIdx - 1}`)?.focus();
      } else if (e.key === "ArrowRight" && target.selectionEnd === String(value).length) {
        document.getElementById(`cell-${rowIdx}-${colIdx + 1}`)?.focus();
      }
    };

    return (
      <td 
        key={`${rowIdx}-${colIdx}`}
        className={`${cellBorder} p-0 relative transition-colors ${isSelected ? 'bg-blue-100/60 outline outline-blue-400 z-10' : ''}`}
        onMouseDown={() => {
          setIsDragging(true);
          setSelection({ start: { r: rowIdx, c: colIdx }, end: { r: rowIdx, c: colIdx } });
        }}
        onMouseEnter={() => {
          if (isDragging && selection) setSelection({ ...selection, end: { r: rowIdx, c: colIdx } });
        }}
      >
        <input
          id={`cell-${rowIdx}-${colIdx}`}
          value={value}
          onChange={e => handleCellChange(rowIdx, field, e.target.value)}
          onKeyDown={handleKeyDown}
          onBlur={() => commitHistory(sheetData)}
          autoComplete="off" 
          spellCheck="false"
          className={`w-full h-full px-2 py-1.5 outline-none focus:ring-2 focus:ring-blue-600 focus:z-20 relative bg-transparent transition-all ${isCenter ? 'text-center' : ''} ${isBold ? 'font-semibold text-slate-800' : 'text-slate-600'}`}
        />
      </td>
    );
  };

  if (!isLoaded) {
    return <div className="h-screen w-full flex items-center justify-center bg-[#f8f9fa]"><span className="text-slate-500 font-bold animate-pulse tracking-widest">LOADING WORKSPACE...</span></div>;
  }

  return (
    <div className="h-screen w-full bg-[#f8f9fa] flex flex-col font-sans overflow-hidden relative">
      
      {/* --- CUSTOM MODERN DIALOG SYSTEM --- */}
      {customDialog.show && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center z-60 animate-in fade-in">
          <div className="bg-white p-6 rounded-xl shadow-2xl max-w-sm w-full border-t-4 border-blue-600">
            <div className="flex items-center gap-3 text-slate-800 mb-2">
              {customDialog.type === "alert" ? <AlertTriangle className="text-orange-500" /> : <Info className="text-blue-500" />}
              <h3 className="font-bold text-lg">{customDialog.title}</h3>
            </div>
            <p className="text-sm text-slate-600 mb-5">{customDialog.message}</p>
            {customDialog.type === "prompt" && (
              <input type="text" autoFocus value={customDialog.inputValue} onChange={(e) => setCustomDialog({...customDialog, inputValue: e.target.value})} autoComplete="off" spellCheck="false" className="w-full border border-slate-300 rounded px-3 py-2 mb-4 focus:ring-2 focus:ring-blue-500 outline-none" />
            )}
            <div className="flex justify-end gap-2 mt-2">
              {(customDialog.type === "confirm" || customDialog.type === "prompt") && (
                <button onClick={closeDialog} className="px-4 py-2 text-sm font-semibold text-slate-500 hover:bg-slate-100 rounded">Batal (Esc)</button>
              )}
              <button onClick={() => { if (customDialog.onConfirm) customDialog.onConfirm(customDialog.inputValue || ""); closeDialog(); }} className="px-5 py-2 text-sm font-semibold bg-blue-600 text-white rounded hover:bg-blue-700 shadow-sm">
                {customDialog.type === "alert" ? "Mengerti (Enter)" : "Simpan (Enter)"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL KONFLIK DUPLIKAT */}
      {conflictModal.show && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center z-50">
          <div className="bg-white p-6 rounded-xl shadow-2xl max-w-md w-full border-t-4 border-orange-500">
            <div className="flex items-center gap-3 text-orange-600 mb-4"><AlertTriangle size={24} /><h3 className="font-bold text-lg">Data Terdeteksi Sama</h3></div>
            <p className="text-sm text-slate-700 mb-4">Ditemukan <strong>{conflictModal.conflicts.length} komponen</strong> duplikat yang sudah ada di sheet <span className="font-bold">{activeTab}</span>. Apa aksi yang diinginkan?</p>
            <div className="flex flex-col gap-2">
              <button onClick={() => resolveConflict("OVERWRITE")} className="w-full bg-blue-600 text-white font-bold py-2 rounded hover:bg-blue-700 flex justify-between px-4"><span>Timpa Data Lama</span><span className="text-blue-300 font-normal">Enter</span></button>
              <button onClick={() => resolveConflict("APPEND")} className="w-full border border-blue-600 text-blue-600 font-bold py-2 rounded hover:bg-blue-50">Tetap Tambahkan Baris</button>
              <button onClick={() => resolveConflict("CANCEL")} className="w-full text-slate-500 font-bold py-2 hover:bg-slate-100 rounded flex justify-between px-4"><span>Batal Masukkan Duplikat</span><span className="font-normal">Esc</span></button>
            </div>
          </div>
        </div>
      )}

      {/* TOOLBAR ATAS */}
      <div className="w-full bg-white border-b border-slate-300 px-5 py-3 flex justify-between items-center shrink-0 shadow-sm z-40 relative">
        <div className="flex items-center gap-4">
          <div className="w-8 h-8 bg-linear-to-br from-green-500 to-green-700 rounded shadow-sm flex items-center justify-center text-white font-bold text-lg">W</div>
          <div>
            <h1 className="text-lg font-bold text-slate-800 leading-tight">Winteq BOM Automator</h1>
            <p className="text-[10px] font-semibold text-slate-400 uppercase tracking-widest">Electrical Design Engineering</p>
          </div>
        </div>
        
        {/* ACTION BUTTONS */}
        <div className="flex items-center gap-3 relative">
          
          {/* UNDO REDO */}
          <div className="flex items-center border border-slate-300 rounded-lg overflow-hidden bg-white shadow-sm">
            <button onClick={handleUndo} disabled={historyIndex <= 0} className="p-2 text-slate-600 hover:bg-slate-100 disabled:opacity-30 transition" title="Undo (Mundur)"><Undo2 size={16} /></button>
            <div className="w-px h-5 bg-slate-300"></div>
            <button onClick={handleRedo} disabled={historyIndex >= history.length - 1} className="p-2 text-slate-600 hover:bg-slate-100 disabled:opacity-30 transition" title="Redo (Maju)"><Redo2 size={16} /></button>
          </div>

          {/* KOSONGKAN TABEL & GRID */}
          <button onClick={clearTable} className={`flex items-center gap-2 px-3 py-2 text-sm font-bold rounded-lg transition border border-red-200 text-red-600 hover:bg-red-50 bg-white shadow-sm`} title="Hapus semua isi tabel ini">
            <Eraser size={16} /> Kosongkan
          </button>
          
          <button onClick={() => setIsThickGrid(!isThickGrid)} className={`flex items-center gap-2 px-3 py-2 text-sm font-bold rounded-lg transition border ${isThickGrid ? 'bg-blue-100 text-blue-800 border-blue-300 shadow-inner' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50 shadow-sm'}`}>
            <Grid size={16} /> Grid
          </button>
          
          <div className="h-8 w-px bg-slate-300 mx-1"></div>

          <ExcelUploader onDataLoaded={handleDataLoaded} />
          
          {/* TOMBOL DROPDOWN DOWNLOAD */}
          <div className="relative" ref={dropdownRef}>
            <button onClick={() => setIsExportMenuOpen(!isExportMenuOpen)} className={`flex items-center gap-2 px-4 py-2 text-sm font-bold rounded-lg transition shadow-sm bg-green-600 text-white hover:bg-green-700`}>
              <Download size={16} /> Export .xlsx <ChevronDown size={14} />
            </button>

            {isExportMenuOpen && (
              <div className="absolute right-0 top-full mt-2 w-72 bg-white border border-slate-200 shadow-xl rounded-xl p-5 z-50 text-slate-700 animate-in fade-in slide-in-from-top-2">
                <label className="flex items-center gap-3 font-bold mb-3 border-b border-slate-100 pb-3 cursor-pointer text-slate-800">
                  <input type="checkbox" className="w-4 h-4 accent-green-600" checked={exportOptions.selectedTabs.length === tabs.length} onChange={(e) => setExportOptions({...exportOptions, selectedTabs: e.target.checked ? [...tabs] : []})} /> Download All
                </label>
                <div className="text-[10px] font-bold text-slate-400 mb-2 uppercase tracking-wider">Select Tab:</div>
                <div className="flex flex-col gap-2 mb-4 max-h-40 overflow-y-auto pr-2">
                  {tabs.map(t => (
                    <label key={t} className="flex items-center gap-3 text-sm font-medium text-slate-600 hover:text-slate-900 cursor-pointer">
                      <input type="checkbox" className="w-4 h-4 accent-blue-600" checked={exportOptions.selectedTabs.includes(t)} onChange={(e) => setExportOptions({...exportOptions, selectedTabs: e.target.checked ? [...exportOptions.selectedTabs, t] : exportOptions.selectedTabs.filter(x => x !== t)})} /> {t}
                    </label>
                  ))}
                </div>
                <div className="text-[10px] font-bold text-slate-400 mb-2 uppercase tracking-wider border-t border-slate-100 pt-4">Options:</div>
                <div className="flex flex-col gap-3 mb-5">
                  <label className="flex items-center gap-2 text-sm font-medium text-slate-600 cursor-pointer"><input type="radio" name="exportMode" value="single" className="w-4 h-4 accent-blue-600" checked={exportOptions.mode === "single"} onChange={() => setExportOptions({...exportOptions, mode: "single"})} /> Gabung (1 File)</label>
                  <label className="flex items-center gap-2 text-sm font-medium text-slate-600 cursor-pointer"><input type="radio" name="exportMode" value="separate" className="w-4 h-4 accent-blue-600" checked={exportOptions.mode === "separate"} onChange={() => setExportOptions({...exportOptions, mode: "separate"})} /> Terpisah (Per Tab)</label>
                </div>
                <button onClick={executeDownload} disabled={isExporting || exportOptions.selectedTabs.length === 0} className="w-full bg-slate-900 text-white font-bold py-2 text-sm rounded hover:bg-slate-800 disabled:bg-slate-300 transition shadow flex justify-center items-center gap-2">
                  {isExporting ? "Memproses..." : "DOWNLOAD"}
                </button>
              </div>
            )}
          </div>

        </div>
      </div>

      {/* MAIN WORKSPACE AREA */}
      <div className="flex-1 w-full p-2 overflow-hidden flex flex-col bg-[#f8f9fa]">
        <div className="flex-1 bg-white border border-slate-300 flex flex-col overflow-hidden shadow-sm">
          
          {/* SPREADSHEET HEADER AREA */}
          <div className="shrink-0 bg-white p-0">
            <table className="w-full border-collapse text-xs">
              <tbody>
                <tr>
                  <td rowSpan={4} colSpan={3} className={`${headerBorder} p-2 w-[35%] align-middle`}>
                    <div className="flex flex-col items-center justify-center pointer-events-none">
                      <div className="flex items-center gap-2">
                        <div className="w-8 h-8 rounded-full bg-linear-to-br from-blue-900 to-red-600 grid grid-cols-2 grid-rows-2 overflow-hidden opacity-80">
                          <div className="bg-blue-400/30"></div><div className="bg-red-400/30"></div><div className="bg-red-400/30"></div><div className="bg-blue-400/30"></div>
                        </div>
                        <div className="flex flex-col">
                          <span className="text-red-600 font-extrabold text-xl tracking-tighter leading-none italic">ASTRA Otoparts</span>
                          <span className="font-bold text-[#1e3a8a] text-xs italic">DIVISI WINTEQ</span>
                        </div>
                      </div>
                      <span className="font-bold text-lg mt-2 tracking-widest text-slate-800">LIST PART OF ELECTRICAL</span>
                    </div>
                  </td>
                  <td colSpan={2} className={`${headerBorder} bg-slate-100 text-center font-semibold text-slate-600 p-1`}>MACHINE NAME</td>
                  <td className={`${headerBorder} bg-slate-100 text-center font-semibold text-slate-600 p-1 w-24`}>DESIGNED</td>
                  <td className={`${headerBorder} bg-slate-100 text-center font-semibold text-slate-600 p-1 w-24`}>ELC</td>
                  <td className={`${headerBorder} bg-slate-100 text-center font-semibold text-slate-600 p-1 w-24`}>CHECKED</td>
                  <td className={`${headerBorder} bg-slate-100 text-center font-semibold text-slate-600 p-1 w-24`}>APPROVED</td>
                </tr>
                <tr>
                  <td colSpan={2} className={`${headerBorder} p-0 h-8`}>
                    <input value={projectInfo.machineName} onChange={(e) => setProjectInfo({...projectInfo, machineName: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full h-full text-center outline-none font-bold text-slate-700 uppercase focus:bg-blue-50" />
                  </td>
                  <td className={`${headerBorder} p-0 h-6`}><input value={projectInfo.date1} onChange={(e) => setProjectInfo({...projectInfo, date1: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full h-full px-1 text-[9px] text-slate-500 outline-none focus:bg-blue-50 bg-transparent text-center" /></td>
                  <td className={`${headerBorder} p-0 h-6`}><input value={projectInfo.date2} onChange={(e) => setProjectInfo({...projectInfo, date2: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full h-full px-1 text-[9px] text-slate-500 outline-none focus:bg-blue-50 bg-transparent text-center" /></td>
                  <td className={`${headerBorder} p-0 h-6`}><input value={projectInfo.date3} onChange={(e) => setProjectInfo({...projectInfo, date3: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full h-full px-1 text-[9px] text-slate-500 outline-none focus:bg-blue-50 bg-transparent text-center" /></td>
                  <td className={`${headerBorder} p-0 h-6`}><input value={projectInfo.date4} onChange={(e) => setProjectInfo({...projectInfo, date4: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full h-full px-1 text-[9px] text-slate-500 outline-none focus:bg-blue-50 bg-transparent text-center" /></td>
                </tr>
                <tr>
                  <td className={`${headerBorder} p-1 font-medium text-slate-600 bg-slate-50 w-24`}>DESIGN ID</td>
                  <td className={`${headerBorder} p-0`}><input value={projectInfo.designId} onChange={(e) => setProjectInfo({...projectInfo, designId: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full h-full px-2 outline-none focus:bg-blue-50" /></td>
                  <td rowSpan={2} className={`${headerBorder} p-0 text-center align-bottom h-10 hover:bg-slate-50 transition`}><input value={projectInfo.designed} onChange={(e) => setProjectInfo({...projectInfo, designed: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full text-center outline-none uppercase font-bold text-slate-700 pb-1 bg-transparent" /></td>
                  <td rowSpan={2} className={`${headerBorder} p-0 text-center align-bottom h-10 hover:bg-slate-50 transition`}><input value={projectInfo.elc} onChange={(e) => setProjectInfo({...projectInfo, elc: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full text-center outline-none uppercase font-bold text-slate-700 pb-1 bg-transparent" /></td>
                  <td rowSpan={2} className={`${headerBorder} p-0 text-center align-bottom h-10 hover:bg-slate-50 transition`}><input value={projectInfo.checked} onChange={(e) => setProjectInfo({...projectInfo, checked: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full text-center outline-none uppercase font-bold text-slate-700 pb-1 bg-transparent" /></td>
                  <td rowSpan={2} className={`${headerBorder} p-0 text-center align-bottom h-10 hover:bg-slate-50 transition bg-white`}><input value={projectInfo.approved} onChange={(e) => setProjectInfo({...projectInfo, approved: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full text-center outline-none uppercase font-bold text-slate-700 pb-1 bg-transparent" /></td>
                </tr>
                <tr>
                  <td className={`${headerBorder} p-1 font-medium text-slate-600 bg-slate-50`}>CUSTOMER</td>
                  <td className={`${headerBorder} p-0`}><input value={projectInfo.customer} onChange={(e) => setProjectInfo({...projectInfo, customer: e.target.value})} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full h-full px-2 outline-none focus:bg-blue-50" /></td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* SPREADSHEET DATA GRID */}
          <div className="flex-1 overflow-auto relative select-none">
            <table className="w-full border-collapse text-[11px]">
              <thead className="sticky top-0 bg-slate-100 z-20 shadow-sm border-b border-slate-300">
                <tr>
                  <th className={`${headerBorder} font-semibold text-slate-600 py-1 w-10 text-center bg-slate-200`}>#</th>
                  <th className={`${headerBorder} font-semibold text-slate-600 p-0 w-32 bg-slate-200 hover:bg-white transition`}><input value={currentHeaders[0]} onChange={e => handleHeaderChange(0, e.target.value)} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full py-1.5 px-2 bg-transparent outline-none focus:ring-1 focus:ring-blue-500 font-semibold" /></th>
                  <th className={`${headerBorder} font-semibold text-slate-600 p-0 min-w-50 bg-slate-200 hover:bg-white transition`}><input value={currentHeaders[1]} onChange={e => handleHeaderChange(1, e.target.value)} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full py-1.5 px-2 bg-transparent outline-none focus:ring-1 focus:ring-blue-500 font-semibold" /></th>
                  <th className={`${headerBorder} font-semibold text-slate-600 p-0 min-w-37.5 bg-slate-200 hover:bg-white transition`}><input value={currentHeaders[2]} onChange={e => handleHeaderChange(2, e.target.value)} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full py-1.5 px-2 bg-transparent outline-none focus:ring-1 focus:ring-blue-500 font-semibold" /></th>
                  <th className={`${headerBorder} font-semibold text-slate-600 p-0 min-w-50 bg-slate-200 hover:bg-white transition`}><input value={currentHeaders[3]} onChange={e => handleHeaderChange(3, e.target.value)} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full py-1.5 px-2 bg-transparent outline-none focus:ring-1 focus:ring-blue-500 font-semibold" /></th>
                  <th className={`${headerBorder} font-semibold text-slate-600 p-0 w-32 bg-slate-200 hover:bg-white transition`}><input value={currentHeaders[4]} onChange={e => handleHeaderChange(4, e.target.value)} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full py-1.5 px-2 bg-transparent outline-none focus:ring-1 focus:ring-blue-500 font-semibold" /></th>
                  <th className={`${headerBorder} font-semibold text-slate-600 p-0 w-16 bg-slate-200 hover:bg-white transition`}><input value={currentHeaders[5]} onChange={e => handleHeaderChange(5, e.target.value)} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full py-1.5 px-2 text-center bg-transparent outline-none focus:ring-1 focus:ring-blue-500 font-semibold" /></th>
                  <th className={`${headerBorder} font-semibold text-slate-600 p-0 w-16 bg-slate-200 hover:bg-white transition`}><input value={currentHeaders[6]} onChange={e => handleHeaderChange(6, e.target.value)} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full py-1.5 px-2 text-center bg-transparent outline-none focus:ring-1 focus:ring-blue-500 font-semibold" /></th>
                  <th className={`${headerBorder} font-semibold text-slate-600 p-0 min-w-37.5 bg-slate-200 hover:bg-white transition`}><input value={currentHeaders[7]} onChange={e => handleHeaderChange(7, e.target.value)} onBlur={() => commitHistory(sheetData)} autoComplete="off" spellCheck="false" className="w-full py-1.5 px-2 bg-transparent outline-none focus:ring-1 focus:ring-blue-500 font-semibold" /></th>
                  <th className={`${headerBorder} font-semibold text-slate-600 py-1.5 w-10 text-center bg-slate-200`}>Act</th>
                </tr>
              </thead>
              <tbody>
                {(sheetData[activeTab] || []).length === 0 ? (
                  <tr>
                    <td colSpan={10} className="text-center py-10 text-slate-400 italic border-b border-slate-200">Data kosong. Silakan upload atau paste data ke sini.</td>
                  </tr>
                ) : (
                  (sheetData[activeTab] || []).map((row, idx) => (
                    <tr key={idx} className="group hover:bg-blue-50/50 h-7">
                      <td className={`${cellBorder} text-center text-slate-400 bg-slate-50 font-medium group-hover:bg-slate-200`}>{idx + 1}</td>
                      {renderCell(idx, 0, "PART CODE", false, true)}
                      {renderCell(idx, 1, "PART NAME")}
                      {renderCell(idx, 2, "DETAIL NAME")}
                      {renderCell(idx, 3, "SPESIFICATION")}
                      {renderCell(idx, 4, "BRAND")}
                      {renderCell(idx, 5, "QTY", true)}
                      {renderCell(idx, 6, "UNT", true)}
                      {renderCell(idx, 7, "REMARK")}
                      <td className={`${cellBorder} text-center p-0 align-middle bg-slate-50`}>
                        <button onClick={() => deleteRow(idx)} className="text-red-400 hover:text-red-600 opacity-0 group-hover:opacity-100 transition mx-auto flex" title="Hapus Baris">
                          <Trash2 size={14} />
                        </button>
                      </td>
                    </tr>
                  ))
                )}
                {/* TOMBOL TAMBAH BARIS */}
                <tr>
                  <td colSpan={10} className={`${cellBorder} bg-slate-50 hover:bg-slate-100 transition border-t border-slate-300`}>
                    <button onClick={addRow} className="w-full py-2 text-xs font-medium text-slate-500 hover:text-blue-600 flex items-center justify-center gap-1 cursor-pointer">
                      <Plus size={14} /> Tambah Baris Manual
                    </button>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

        </div>

        {/* TABS SHEET BOTTOM AREA */}
        <div className="bg-slate-100 border-t border-slate-300 flex items-center shrink-0">
          <div className="flex bg-slate-200 overflow-x-auto w-full">
            <button onClick={addNewTab} className="px-3 py-2 text-slate-500 hover:text-slate-800 hover:bg-slate-300 flex items-center justify-center border-r border-slate-300">
              <Plus size={16} />
            </button>
            {tabs.map((tab) => (
              <div 
                key={tab} 
                className={`group flex items-center border-r border-slate-300 text-xs font-medium cursor-pointer transition max-w-37.5
                  ${activeTab === tab ? "bg-white text-green-700 shadow-[0_-2px_0_0_#16a34a_inset]" : "bg-slate-100 text-slate-600 hover:bg-slate-50"}`
                }
                onClick={() => setActiveTab(tab)}
              >
                <span className="px-4 py-2 truncate" onDoubleClick={() => renameTab(tab)} title="Double click to rename">{tab}</span>
                <div className="px-1 opacity-0 group-hover:opacity-100 transition">
                  <button onClick={(e) => deleteTab(tab, e)} className="text-slate-400 hover:text-red-500 p-0.5 rounded-full hover:bg-slate-200">
                    <X size={12} />
                  </button>
                </div>
              </div>
            ))}
          </div>
        </div>

      </div>
    </div>
  );
}