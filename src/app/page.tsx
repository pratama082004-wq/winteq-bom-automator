"use client";

import { useState, useRef, useEffect } from "react";
import { Download, AlertTriangle, Plus, Trash2, X, Grid, Info, ChevronDown, Undo2, Redo2, Eraser, Settings2, Upload, Menu, FolderKanban } from "lucide-react";
import ExcelUploader, { BOMRow } from "../components/ExcelUploader";
import { exportWinteqBOM } from "../utils/exportBOM";

const DEFAULT_HEADERS = ["PART CODE", "PART NAME", "DETAIL NAME", "SPECIFICATION", "BRAND", "QTY", "UNT", "REMARK"];
const COLUMNS: (keyof BOMRow)[] = ["PART CODE", "PART NAME", "DETAIL NAME", "SPESIFICATION", "BRAND", "QTY", "UNT", "REMARK"];

const DEFAULT_PROJECT_INFO = {
  machineName: "AUTO FAN COMP COOLING ASSY", designId: "612469", customer: "ASKI",
  designed: "AGH", elc: "HLM", checked: "DBY", approved: "YSF",
  date1: "Date.", date2: "Date.", date3: "Date.", date4: "Date."
};

const DEFAULT_HEADER_LABELS = {
  title: "LIST PART OF ELECTRICAL", machineName: "MACHINE NAME", designed: "DESIGNED",
  elc: "ELC", checked: "CHECKED", approved: "APPROVED", designId: "DESIGN ID", customer: "CUSTOMER"
};

export default function Home() {
  // --- STATE MULTI-PROJECT (SIDEBAR) ---
  const [projects, setProjects] = useState<{ id: string, name: string }[]>([]);
  const [activeProjectId, setActiveProjectId] = useState<string>("");
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  const [draggedProject, setDraggedProject] = useState<string | null>(null);
  
  const [editingProjectId, setEditingProjectId] = useState<string | null>(null);
  const [editingProjectName, setEditingProjectName] = useState("");

  // --- STATE DATA WORKSPACE ---
  const [tabs, setTabs] = useState(["MAIN", "PNEUMATIC", "ACCESSORIES"]);
  const [activeTab, setActiveTab] = useState("MAIN");
  const [sheetData, setSheetData] = useState<Record<string, BOMRow[]>>({ "MAIN": [], "PNEUMATIC": [], "ACCESSORIES": [] });
  const [tabHeaders, setTabHeaders] = useState<Record<string, string[]>>({ "MAIN": [...DEFAULT_HEADERS], "PNEUMATIC": [...DEFAULT_HEADERS], "ACCESSORIES": ["SERIES", "PART NAME", "DETAIL NAME", "SPECIFICATION", "BRAND", "QTY", "UNT", "REMARK"] });
  const [projectInfo, setProjectInfo] = useState<Record<string, typeof DEFAULT_PROJECT_INFO>>({ "MAIN": { ...DEFAULT_PROJECT_INFO }, "PNEUMATIC": { ...DEFAULT_PROJECT_INFO }, "ACCESSORIES": { ...DEFAULT_PROJECT_INFO } });
  const [headerLabels, setHeaderLabels] = useState<Record<string, typeof DEFAULT_HEADER_LABELS>>({ "MAIN": { ...DEFAULT_HEADER_LABELS }, "PNEUMATIC": { ...DEFAULT_HEADER_LABELS }, "ACCESSORIES": { ...DEFAULT_HEADER_LABELS, title: "LIST PART OF ACCESSORIES" } });

  const [isLoaded, setIsLoaded] = useState(false);
  const [isTransitioning, setIsTransitioning] = useState(false);
  const [isThickGrid, setIsThickGrid] = useState(false);
  const [draggedTab, setDraggedTab] = useState<string | null>(null);
  const currentActiveProjectId = useRef("");

  useEffect(() => {
    const savedList = localStorage.getItem("winteq_project_list");
    let initialProjects = [];
    let startProjectId = "p1";

    if (savedList) {
      initialProjects = JSON.parse(savedList);
      startProjectId = localStorage.getItem("winteq_active_project") || initialProjects[0].id;
    } else {
      initialProjects = [{ id: "p1", name: "Project 1" }];
      if (localStorage.getItem("winteq_tabs")) {
        localStorage.setItem("winteq_p1_tabs", localStorage.getItem("winteq_tabs") || "");
        localStorage.setItem("winteq_p1_activeTab", localStorage.getItem("winteq_activeTab") || "MAIN");
        localStorage.setItem("winteq_p1_sheetData", localStorage.getItem("winteq_sheetData") || "{}");
        localStorage.setItem("winteq_p1_headers", localStorage.getItem("winteq_headers") || "{}");
        localStorage.setItem("winteq_p1_projectInfo", localStorage.getItem("winteq_projectInfo") || "{}");
        localStorage.setItem("winteq_p1_headerLabels", localStorage.getItem("winteq_headerLabels") || "{}");
      }
    }

    setProjects(initialProjects);
    setActiveProjectId(startProjectId);
    currentActiveProjectId.current = startProjectId;
    loadProjectData(startProjectId);
  }, []);

  const loadProjectData = (id: string) => {
    const getL = (key: string) => localStorage.getItem(`winteq_${id}_${key}`);
    setTabs(getL("tabs") ? JSON.parse(getL("tabs")!) : ["MAIN", "PNEUMATIC", "ACCESSORIES"]);
    setActiveTab(getL("activeTab") || "MAIN");
    if (getL("sheetData")) {
      const pData = JSON.parse(getL("sheetData")!);
      setSheetData(pData); setHistory([JSON.stringify(pData)]); setHistoryIndex(0);
    } else {
      setSheetData({ "MAIN": [], "PNEUMATIC": [], "ACCESSORIES": [] }); setHistory([JSON.stringify({ "MAIN": [], "PNEUMATIC": [], "ACCESSORIES": [] })]); setHistoryIndex(0);
    }
    setTabHeaders(getL("headers") ? JSON.parse(getL("headers")!) : { "MAIN": [...DEFAULT_HEADERS], "PNEUMATIC": [...DEFAULT_HEADERS], "ACCESSORIES": ["SERIES", "PART NAME", "DETAIL NAME", "SPECIFICATION", "BRAND", "QTY", "UNT", "REMARK"] });
    setProjectInfo(getL("projectInfo") ? JSON.parse(getL("projectInfo")!) : { "MAIN": { ...DEFAULT_PROJECT_INFO }, "PNEUMATIC": { ...DEFAULT_PROJECT_INFO }, "ACCESSORIES": { ...DEFAULT_PROJECT_INFO } });
    setHeaderLabels(getL("headerLabels") ? JSON.parse(getL("headerLabels")!) : { "MAIN": { ...DEFAULT_HEADER_LABELS }, "PNEUMATIC": { ...DEFAULT_HEADER_LABELS }, "ACCESSORIES": { ...DEFAULT_HEADER_LABELS, title: "LIST PART OF ACCESSORIES" } });
    setIsLoaded(true);
  };

  useEffect(() => { currentActiveProjectId.current = activeProjectId; }, [activeProjectId]);

  useEffect(() => {
    if (isLoaded && currentActiveProjectId.current) {
      const id = currentActiveProjectId.current;
      localStorage.setItem(`winteq_${id}_tabs`, JSON.stringify(tabs));
      localStorage.setItem(`winteq_${id}_activeTab`, activeTab);
      localStorage.setItem(`winteq_${id}_sheetData`, JSON.stringify(sheetData));
      localStorage.setItem(`winteq_${id}_headers`, JSON.stringify(tabHeaders));
      localStorage.setItem(`winteq_${id}_projectInfo`, JSON.stringify(projectInfo));
      localStorage.setItem(`winteq_${id}_headerLabels`, JSON.stringify(headerLabels));
    }
  }, [tabs, activeTab, sheetData, tabHeaders, projectInfo, headerLabels, isLoaded]);

  useEffect(() => {
    if (isLoaded) {
      localStorage.setItem("winteq_project_list", JSON.stringify(projects));
      localStorage.setItem("winteq_active_project", activeProjectId);
    }
  }, [projects, activeProjectId, isLoaded]);

  const handleSwitchProject = (id: string) => {
    if (id === activeProjectId) return;
    setIsTransitioning(true);
    setTimeout(() => {
      setActiveProjectId(id);
      loadProjectData(id);
      setTimeout(() => setIsTransitioning(false), 50); 
    }, 200);
  };

  const handleAddNewProject = () => {
    const newId = `p${Date.now()}`;
    const newName = `Project ${projects.length + 1}`;
    setProjects([...projects, { id: newId, name: newName }]);
    handleSwitchProject(newId);
  };

  const startEditingProjectName = (id: string, currentName: string) => { setEditingProjectId(id); setEditingProjectName(currentName); };
  const saveEditingProjectName = () => {
    if (editingProjectId && editingProjectName.trim() !== "") setProjects(prev => prev.map(p => p.id === editingProjectId ? { ...p, name: editingProjectName.trim() } : p));
    setEditingProjectId(null);
  };

  const handleDeleteProject = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (projects.length === 1) return showAlert("Dilarang", "Minimal harus ada 1 project di Workspace!");
    showConfirm("Hapus Project", "Yakin ingin menghapus project ini? Semua data di dalamnya akan musnah permanen.", () => {
      const newProjects = projects.filter(p => p.id !== id);
      setProjects(newProjects);
      ["tabs", "activeTab", "sheetData", "headers", "projectInfo", "headerLabels"].forEach(k => localStorage.removeItem(`winteq_${id}_${k}`));
      if (activeProjectId === id) handleSwitchProject(newProjects[0].id);
    });
  };

  const [syncModal, setSyncModal] = useState<{ show: boolean, fieldType: 'header' | 'project' | null, fieldKey: string, value: string, pendingFieldId: string }>({ show: false, fieldType: null, fieldKey: "", value: "", pendingFieldId: "" });
  const [syncChoice, setSyncChoice] = useState<'single' | 'all' | null>(null);
  const [dontAskAgain, setDontAskAgain] = useState(false);
  const [syncSettings, setSyncSettings] = useState<{ mode: 'single' | 'all' | null, askAgain: boolean }>({ mode: null, askAgain: true });
  const initialValueRef = useRef(""); 
  const [history, setHistory] = useState<string[]>([]);
  const [historyIndex, setHistoryIndex] = useState(-1);

  const updateHeaderLabel = (field: string, value: string, forceMode?: 'single' | 'all') => {
    const mode = forceMode || syncSettings.mode || 'single';
    setHeaderLabels(prev => {
      if (mode === 'all') { const next = { ...prev }; Object.keys(next).forEach(k => next[k] = { ...(next[k] || DEFAULT_HEADER_LABELS), [field]: value }); return next; }
      return { ...prev, [activeTab]: { ...(prev[activeTab] || DEFAULT_HEADER_LABELS), [field]: value } };
    });
  };

  const updateProjectInfo = (field: string, value: string, forceMode?: 'single' | 'all') => {
    const mode = forceMode || syncSettings.mode || 'single';
    setProjectInfo(prev => {
      if (mode === 'all') { const next = { ...prev }; Object.keys(next).forEach(k => next[k] = { ...(next[k] || DEFAULT_PROJECT_INFO), [field]: value }); return next; }
      return { ...prev, [activeTab]: { ...(prev[activeTab] || DEFAULT_PROJECT_INFO), [field]: value } };
    });
  };

  const getTemplateHandlers = (fieldType: 'header' | 'project', fieldKey: string, fieldId: string) => ({
    onFocus: (e: React.FocusEvent<HTMLInputElement>) => { initialValueRef.current = e.target.value; },
    onBlur: (e: React.FocusEvent<HTMLInputElement>) => {
      commitHistory(sheetData);
      if (syncSettings.askAgain && e.target.value !== initialValueRef.current) {
        setSyncChoice(null); setDontAskAgain(false);
        setSyncModal({ show: true, fieldType, fieldKey, value: e.target.value, pendingFieldId: fieldId });
        initialValueRef.current = e.target.value; 
      }
    },
    onKeyDown: (e: React.KeyboardEvent<HTMLInputElement>) => { if (e.key === 'Enter') e.currentTarget.blur(); }
  });

  const confirmSyncModal = () => {
    if (!syncChoice) return;
    setSyncSettings({ mode: syncChoice, askAgain: !dontAskAgain });
    if (syncChoice === 'all' && syncModal.fieldType && syncModal.fieldKey) {
      if (syncModal.fieldType === 'header') updateHeaderLabel(syncModal.fieldKey, syncModal.value, 'all');
      else updateProjectInfo(syncModal.fieldKey, syncModal.value, 'all');
    }
    setSyncModal({ show: false, fieldType: null, fieldKey: "", value: "", pendingFieldId: "" });
  };

  const commitHistory = (newData: Record<string, BOMRow[]>) => {
    const strData = JSON.stringify(newData);
    if (historyIndex >= 0 && history[historyIndex] === strData) return; 
    const newHistory = history.slice(0, historyIndex + 1);
    newHistory.push(strData);
    if (newHistory.length > 50) newHistory.shift(); 
    setHistory(newHistory); setHistoryIndex(newHistory.length - 1);
  };

  const handleUndo = () => { if (historyIndex > 0) { const prevIdx = historyIndex - 1; setHistoryIndex(prevIdx); setSheetData(JSON.parse(history[prevIdx])); } };
  const handleRedo = () => { if (historyIndex < history.length - 1) { const nextIdx = historyIndex + 1; setHistoryIndex(nextIdx); setSheetData(JSON.parse(history[nextIdx])); } };

  const [selection, setSelection] = useState<{ start: { r: number; c: number }; end: { r: number; c: number } } | null>(null);
  const [isDragging, setIsDragging] = useState(false);
  const [conflictModal, setConflictModal] = useState<{ show: boolean, conflicts: BOMRow[], newCleanData: BOMRow[] }>({ show: false, conflicts: [], newCleanData: [] });
  const [isExportMenuOpen, setIsExportMenuOpen] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [exportOptions, setExportOptions] = useState<{ selectedTabs: string[]; mode: "single" | "separate" }>({ selectedTabs: ["MAIN", "PNEUMATIC", "ACCESSORIES"], mode: "separate" });
  const dropdownRef = useRef<HTMLDivElement>(null);
  
  const [customDialog, setCustomDialog] = useState<{ show: boolean; type: "alert" | "confirm" | "prompt"; title: string; message: string; inputValue?: string; onConfirm?: any; }>({ show: false, type: "alert", title: "", message: "" });
  const showAlert = (title: string, message: string) => setCustomDialog({ show: true, type: "alert", title, message });
  const showConfirm = (title: string, message: string, onConfirm: () => void) => setCustomDialog({ show: true, type: "confirm", title, message, onConfirm });
  const showPrompt = (title: string, message: string, defaultVal: string, onConfirm: (val: string) => void) => setCustomDialog({ show: true, type: "prompt", title, message, inputValue: defaultVal, onConfirm });
  const closeDialog = () => setCustomDialog({ ...customDialog, show: false });

  const shortcutsRef = useRef({ customDialog, conflictModal, syncModal, isExportMenuOpen, closeDialog, setConflictModal, setIsExportMenuOpen, sheetData, activeTab, confirmSyncModal, setSyncModal });
  useEffect(() => { shortcutsRef.current = { customDialog, conflictModal, syncModal, isExportMenuOpen, closeDialog, setConflictModal, setIsExportMenuOpen, sheetData, activeTab, confirmSyncModal, setSyncModal }; });

  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      const { customDialog, conflictModal, syncModal, isExportMenuOpen, closeDialog, setConflictModal, setIsExportMenuOpen, sheetData, activeTab, confirmSyncModal, setSyncModal } = shortcutsRef.current;
      
      if (syncModal.show) {
        if (e.key === "Escape") setSyncModal({ show: false, fieldType: null, fieldKey: "", value: "", pendingFieldId: "" });
        else if (e.key === "Enter") { e.preventDefault(); confirmSyncModal(); }
        return;
      }
      if (customDialog.show) {
        if (e.key === "Escape") closeDialog();
        else if (e.key === "Enter") { e.preventDefault(); if (customDialog.onConfirm) customDialog.onConfirm(customDialog.inputValue || ""); closeDialog(); }
        return; 
      }
      if (conflictModal.show) {
        if (e.key === "Escape") setConflictModal({ show: false, conflicts: [], newCleanData: [] });
        return;
      }
      if (isExportMenuOpen && e.key === "Escape") setIsExportMenuOpen(false);
    };
    window.addEventListener("keydown", handleGlobalKeyDown);
    return () => window.removeEventListener("keydown", handleGlobalKeyDown);
  }, []);

  useEffect(() => {
    const handleGlobalMouseUp = () => setIsDragging(false);
    const handleClickOutside = (e: MouseEvent) => {
      if (!(e.target as HTMLElement).closest("table")) setSelection(null);
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) setIsExportMenuOpen(false);
    };
    window.addEventListener("mouseup", handleGlobalMouseUp);
    window.addEventListener("mousedown", handleClickOutside);
    return () => { window.removeEventListener("mouseup", handleGlobalMouseUp); window.removeEventListener("mousedown", handleClickOutside); };
  }, []);

  useEffect(() => {
    const handleCopy = (e: ClipboardEvent) => {
      if (!selection) return;
      const activeEl = document.activeElement as HTMLInputElement;
      if (activeEl?.tagName === "INPUT" && activeEl.selectionStart !== activeEl.selectionEnd) return; 
      e.preventDefault();
      const minR = Math.min(selection.start.r, selection.end.r); const maxR = Math.max(selection.start.r, selection.end.r);
      const minC = Math.min(selection.start.c, selection.end.c); const maxC = Math.max(selection.start.c, selection.end.c);
      const currentData = sheetData[activeTab] || [];
      let tsv = "";
      for (let r = minR; r <= maxR; r++) {
        let rowStr = [];
        for (let c = minC; c <= maxC; c++) rowStr.push(currentData[r] ? currentData[r][COLUMNS[c]] || "" : "");
        tsv += rowStr.join("\t") + "\n";
      }
      e.clipboardData?.setData("text/plain", tsv);
    };

    const handlePaste = (e: ClipboardEvent) => {
      const text = e.clipboardData?.getData("text/plain");
      if (!text || !text.includes("\t") || !selection) return;
      e.preventDefault();
      const rows = text.split(/\r?\n/).filter(r => r.trim() !== "");
      const currentData = [...(sheetData[activeTab] || [])];
      const startRowIdx = Math.min(selection.start.r, selection.end.r); const startColIdx = Math.min(selection.start.c, selection.end.c);
      rows.forEach((rowString, rOffset) => {
        const targetRowIdx = startRowIdx + rOffset;
        if (!currentData[targetRowIdx]) currentData.push({ NO: currentData.length + 1, "PART CODE": "", "PART NAME": "", QTY: 1, UNT: "PCS" });
        const cells = rowString.split("\t");
        cells.forEach((cellVal, cOffset) => {
          const targetColIdx = startColIdx + cOffset;
          if (targetColIdx < COLUMNS.length) currentData[targetRowIdx][COLUMNS[targetColIdx]] = cellVal.trim();
        });
      });
      const newData = { ...sheetData, [activeTab]: currentData.map((item, idx) => ({ ...item, NO: idx + 1 })) };
      setSheetData(newData); commitHistory(newData);
    };

    document.addEventListener("copy", handleCopy); document.addEventListener("paste", handlePaste);
    return () => { document.removeEventListener("copy", handleCopy); document.removeEventListener("paste", handlePaste); };
  }, [selection, sheetData, activeTab]);


  const handleImportClick = () => {
    const hasAnyData = Object.values(sheetData).some(data => data.length > 0);
    if (hasAnyData) {
      showConfirm("Timpa Data Pekerjaan?", "Mengunggah Excel akan menghapus dan menimpa seluruh pekerjaan Anda saat ini di layar. Lanjutkan?", () => {
        document.getElementById("import-excel")?.click();
      });
    } else {
      document.getElementById("import-excel")?.click();
    }
  };

  const handleImportExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      showAlert("Memproses...", "Sedang mengekstrak data dari file Excel...");
      const ExcelJS = (await import("exceljs")).default || await import("exceljs");
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);

      const newTabs: string[] = [];
      const newSheetData: Record<string, BOMRow[]> = {};
      const newProjectInfo: Record<string, typeof DEFAULT_PROJECT_INFO> = {};
      const newHeaderLabels: Record<string, typeof DEFAULT_HEADER_LABELS> = {};
      const newTabHeaders: Record<string, string[]> = {};

      const getCellTextSafe = (sheet: any, cellRef: string) => {
        const cell = sheet.getCell(cellRef);
        if (!cell || cell.value === null || cell.value === undefined) return "";
        if (typeof cell.value === 'object') {
          if (cell.value.richText) return cell.value.richText.map((rt: any) => rt.text).join('');
          if (cell.value.result !== undefined) return String(cell.value.result);
        }
        return String(cell.value);
      };

      workbook.eachSheet((sheet) => {
        const tabName = sheet.name.toUpperCase();
        newTabs.push(tabName);

        newHeaderLabels[tabName] = {
          title: getCellTextSafe(sheet, "C3") || DEFAULT_HEADER_LABELS.title,
          machineName: getCellTextSafe(sheet, "F1") || DEFAULT_HEADER_LABELS.machineName,
          designed: getCellTextSafe(sheet, "H1") || DEFAULT_HEADER_LABELS.designed,
          elc: getCellTextSafe(sheet, "I1") || DEFAULT_HEADER_LABELS.elc,
          checked: getCellTextSafe(sheet, "J1") || DEFAULT_HEADER_LABELS.checked,
          approved: getCellTextSafe(sheet, "K1") || DEFAULT_HEADER_LABELS.approved,
          designId: getCellTextSafe(sheet, "F3") || DEFAULT_HEADER_LABELS.designId,
          customer: getCellTextSafe(sheet, "F4") || DEFAULT_HEADER_LABELS.customer
        };

        newProjectInfo[tabName] = {
          machineName: getCellTextSafe(sheet, "F2"),
          designId: getCellTextSafe(sheet, "G3"),
          customer: getCellTextSafe(sheet, "G4"),
          date1: getCellTextSafe(sheet, "H2"),
          date2: getCellTextSafe(sheet, "I2"),
          date3: getCellTextSafe(sheet, "J2"),
          date4: getCellTextSafe(sheet, "K2"),
          designed: getCellTextSafe(sheet, "H4"),
          elc: getCellTextSafe(sheet, "I4"),
          checked: getCellTextSafe(sheet, "J4"),
          approved: getCellTextSafe(sheet, "K4")
        };

        newTabHeaders[tabName] = [
          getCellTextSafe(sheet, "C6") || "PART CODE",
          getCellTextSafe(sheet, "D6") || "PART NAME",
          getCellTextSafe(sheet, "E6") || "DETAIL NAME",
          getCellTextSafe(sheet, "F6") || "SPECIFICATION",
          getCellTextSafe(sheet, "G6") || "BRAND",
          getCellTextSafe(sheet, "H6") || "QTY",
          getCellTextSafe(sheet, "I6") || "UNT",
          getCellTextSafe(sheet, "J6") || "REMARK"
        ];

        const rows: BOMRow[] = [];
        let r = 7;
        while (true) {
          const no = getCellTextSafe(sheet, `A${r}`);
          const partCode = getCellTextSafe(sheet, `C${r}`);
          const partName = getCellTextSafe(sheet, `D${r}`);
          if (!no && !partCode && !partName && r > 200) break; 
          if (!no && !partCode && !partName) { if (!getCellTextSafe(sheet, `A${r+1}`) && !getCellTextSafe(sheet, `C${r+1}`)) break; }
          if (partCode || partName) {
            rows.push({
              NO: Number(no) || rows.length + 1, "PART CODE": partCode, "PART NAME": partName, "DETAIL NAME": getCellTextSafe(sheet, `E${r}`),
              "SPESIFICATION": getCellTextSafe(sheet, `F${r}`), "BRAND": getCellTextSafe(sheet, `G${r}`), "QTY": Number(sheet.getCell(`H${r}`).value) || 1,
              "UNT": getCellTextSafe(sheet, `I${r}`) || "PCS", "REMARK": getCellTextSafe(sheet, `J${r}`)
            });
          }
          r++;
        }
        newSheetData[tabName] = rows;
      });

      if (newTabs.length > 0) {
        setTabs(newTabs); setSheetData(newSheetData); setProjectInfo(newProjectInfo); setHeaderLabels(newHeaderLabels); setTabHeaders(newTabHeaders);
        setActiveTab(newTabs[0]); setExportOptions(prev => ({ ...prev, selectedTabs: newTabs })); commitHistory(newSheetData);
        showAlert("Berhasil Import!", "Seluruh data, header, dan format Excel berhasil di-load ke dalam sistem.");
      } else {
         showAlert("Error Format", "Tidak menemukan sheet yang valid di file Excel ini.");
      }

    } catch (error: any) {
      showAlert("Error Import", "Gagal membaca file Excel. File mungkin rusak atau formatnya tidak dikenali.");
    }
    e.target.value = ''; 
  };

  const handleDataLoaded = (newRawData: BOMRow[]) => {
    const currentData = sheetData[activeTab] || [];
    const conflicts: BOMRow[] = []; const cleanNewData: BOMRow[] = [];
    newRawData.forEach(row => {
      const exists = currentData.some(existing => existing["PART CODE"] === row["PART CODE"] && row["PART CODE"] !== "");
      if (exists) conflicts.push(row); else cleanNewData.push(row);
    });
    if (conflicts.length > 0) setConflictModal({ show: true, conflicts, newCleanData: cleanNewData });
    else updateSheetData(activeTab, [...currentData, ...cleanNewData]);
  };

  const resolveConflict = (action: "CANCEL" | "OVERWRITE" | "APPEND") => {
    const currentData = sheetData[activeTab] || []; let updatedData = [...currentData];
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
    setSheetData(newData); commitHistory(newData);
  };

  const handleCellChange = (rowIndex: number, field: keyof BOMRow, value: string) => {
    const currentData = [...(sheetData[activeTab] || [])];
    currentData[rowIndex] = { ...currentData[rowIndex], [field]: value };
    setSheetData(prev => ({ ...prev, [activeTab]: currentData }));
  };

  const handleHeaderChange = (colIndex: number, newValue: string) => {
    setTabHeaders(prev => {
      const current = prev[activeTab] || [...DEFAULT_HEADERS];
      const updated = [...current]; updated[colIndex] = newValue.toUpperCase(); return { ...prev, [activeTab]: updated };
    });
  };

  const addRow = () => {
    const currentData = [...(sheetData[activeTab] || [])];
    currentData.push({ NO: currentData.length + 1, "PART CODE": "", "PART NAME": "", QTY: 1, UNT: "PCS" }); updateSheetData(activeTab, currentData);
  };

  const deleteRow = (rowIndex: number) => {
    const currentData = [...(sheetData[activeTab] || [])];
    currentData.splice(rowIndex, 1); updateSheetData(activeTab, currentData);
  };

  const clearTable = () => showConfirm("Kosongkan Tabel", `Yakin ingin menghapus SELURUH data di sheet ${activeTab}? (Bisa di-Undo)`, () => updateSheetData(activeTab, []));

  const addNewTab = () => {
    showPrompt("Sheet Baru", "Masukkan nama untuk Sheet baru:", "", (newTabName: string) => {
      if (!newTabName) return; const name = newTabName.toUpperCase();
      if (!tabs.includes(name)) {
        setTabs([...tabs, name]); setSheetData({ ...sheetData, [name]: [] }); setTabHeaders({ ...tabHeaders, [name]: [...DEFAULT_HEADERS] });
        setProjectInfo({ ...projectInfo, [name]: { ...DEFAULT_PROJECT_INFO } }); setHeaderLabels({ ...headerLabels, [name]: { ...DEFAULT_HEADER_LABELS } });
        setExportOptions(prev => ({ ...prev, selectedTabs: [...prev.selectedTabs, name] })); setActiveTab(name);
      } else showAlert("Peringatan", "Nama sheet sudah ada!");
    });
  };

  const renameTab = (oldName: string) => {
    showPrompt("Ubah Nama Sheet", `Masukkan nama baru untuk ${oldName}:`, oldName, (newName: string) => {
      if (!newName) return; const name = newName.toUpperCase();
      if (name !== oldName && !tabs.includes(name)) {
        setTabs(tabs.map(t => t === oldName ? name : t));
        setSheetData(prev => { const newData = { ...prev, [name]: prev[oldName] || [] }; delete newData[oldName]; return newData; });
        setTabHeaders(prev => { const newHeaders = { ...prev, [name]: prev[oldName] || [...DEFAULT_HEADERS] }; delete newHeaders[oldName]; return newHeaders; });
        setProjectInfo(prev => { const newInfo = { ...prev, [name]: prev[oldName] || { ...DEFAULT_PROJECT_INFO } }; delete newInfo[oldName]; return newInfo; });
        setHeaderLabels(prev => { const newLabels = { ...prev, [name]: prev[oldName] || { ...DEFAULT_HEADER_LABELS } }; delete newLabels[oldName]; return newLabels; });
        setExportOptions(prev => ({ ...prev, selectedTabs: prev.selectedTabs.map(t => t === oldName ? name : t) }));
        if (activeTab === oldName) setActiveTab(name);
      }
    });
  };

  const deleteTab = (tabName: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (tabs.length === 1) return showAlert("Dilarang", "Minimal harus ada 1 sheet!");
    showConfirm("Hapus Sheet", `Yakin ingin menghapus sheet ${tabName}? Semua datanya akan musnah.`, () => {
      const newTabs = tabs.filter(t => t !== tabName); setTabs(newTabs);
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
    await exportWinteqBOM(sheetData, projectInfo, tabHeaders, exportOptions.selectedTabs, exportOptions.mode, headerLabels);
    setIsExporting(false); setIsExportMenuOpen(false); 
  };

  const headerBorder = isThickGrid ? "border-2 border-slate-800" : "border border-slate-300";
  const cellBorder = isThickGrid ? "border-2 border-slate-800" : "border-r border-b border-slate-200";
  
  const currentHeaders = tabHeaders[activeTab] || DEFAULT_HEADERS;
  const currentProjectInfo = projectInfo[activeTab] || DEFAULT_PROJECT_INFO;
  const currentHeaderLabels = headerLabels[activeTab] || DEFAULT_HEADER_LABELS;

  // --- KOMPONEN SEL DENGAN ARROW KEY NAVIGASI ---
  const renderCell = (rowIdx: number, colIdx: number, field: keyof BOMRow, isCenter = false, isBold = false) => {
    const value = sheetData[activeTab]?.[rowIdx]?.[field] || "";
    let isSelected = false;
    if (selection) {
      const minR = Math.min(selection.start.r, selection.end.r); const maxR = Math.max(selection.start.r, selection.end.r);
      const minC = Math.min(selection.start.c, selection.end.c); const maxC = Math.max(selection.start.c, selection.end.c);
      isSelected = rowIdx >= minR && rowIdx <= maxR && colIdx >= minC && colIdx <= maxC;
    }

    const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
      const target = e.currentTarget;
      if (e.key === "ArrowUp") { e.preventDefault(); document.getElementById(`cell-${rowIdx - 1}-${colIdx}`)?.focus(); } 
      else if (e.key === "ArrowDown") { e.preventDefault(); document.getElementById(`cell-${rowIdx + 1}-${colIdx}`)?.focus(); } 
      else if (e.key === "ArrowLeft" && target.selectionStart === 0) document.getElementById(`cell-${rowIdx}-${colIdx - 1}`)?.focus();
      else if (e.key === "ArrowRight" && target.selectionEnd === String(value).length) document.getElementById(`cell-${rowIdx}-${colIdx + 1}`)?.focus();
    };

    return (
      <td 
        key={`${rowIdx}-${colIdx}`}
        className={`${cellBorder} p-0 relative transition-colors ${isSelected ? 'bg-blue-100/60 outline outline-blue-400 z-10' : ''}`}
        onMouseDown={() => { setIsDragging(true); setSelection({ start: { r: rowIdx, c: colIdx }, end: { r: rowIdx, c: colIdx } }); }}
        onMouseEnter={() => { if (isDragging && selection) setSelection({ ...selection, end: { r: rowIdx, c: colIdx } }); }}
      >
        <input
          id={`cell-${rowIdx}-${colIdx}`} value={value} onChange={e => handleCellChange(rowIdx, field, e.target.value)} onKeyDown={handleKeyDown} onBlur={() => commitHistory(sheetData)}
          autoComplete="off" spellCheck="false"
          className={`w-full h-full px-2 py-1.5 outline-none focus:ring-2 focus:ring-blue-600 focus:z-20 relative bg-transparent transition-all ${isCenter ? 'text-center' : ''} ${isBold ? 'font-semibold text-slate-800' : 'text-slate-600'}`}
        />
      </td>
    );
  };

  if (!isLoaded) return <div className="h-screen w-full flex items-center justify-center bg-[#f8f9fa]"><span className="text-slate-500 font-bold animate-pulse tracking-widest">MEMUAT WORKSPACE...</span></div>;

  return (
    <div className="flex h-screen w-full font-sans overflow-hidden bg-[#f8f9fa]">
      
      {/* ================= SIDEBAR PROJECT MANAGER ================= */}
      <div className={`${isSidebarOpen ? 'w-64' : 'w-0'} transition-all duration-300 ease-in-out bg-[#003366] text-white flex flex-col shrink-0 overflow-hidden shadow-2xl z-50`}>
        <div className="p-4 font-bold border-b border-blue-800 flex justify-between items-center whitespace-nowrap bg-[#002244]">
          <div className="flex items-center gap-2 text-blue-300">
             <FolderKanban size={18} />
             <span>WINTEQ OS</span>
          </div>
          <button onClick={() => setIsSidebarOpen(false)} className="hover:text-red-400 transition"><X size={18} /></button>
        </div>
        
        <div className="flex-1 overflow-y-auto p-3 flex flex-col gap-1">
          <div className="text-[10px] font-bold text-blue-300/70 mb-2 uppercase tracking-widest px-1">Daftar Proyek Anda</div>
          {projects.map((p) => (
            <div 
              key={p.id}
              draggable
              onDragStart={() => setDraggedProject(p.id)}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                if (!draggedProject || draggedProject === p.id) return;
                const newProjs = [...projects];
                const dIdx = newProjs.findIndex(x => x.id === draggedProject);
                const tIdx = newProjs.findIndex(x => x.id === p.id);
                const item = newProjs.splice(dIdx, 1)[0];
                newProjs.splice(tIdx, 0, item);
                setProjects(newProjs);
                setDraggedProject(null);
              }}
              onClick={() => { if (editingProjectId !== p.id) handleSwitchProject(p.id); }}
              className={`p-2.5 rounded-lg cursor-pointer border flex justify-between items-center group transition-all duration-200
                ${activeProjectId === p.id ? 'bg-blue-600 border-blue-400 shadow-md' : 'bg-white/5 border-white/5 hover:bg-white/10'}
                ${draggedProject === p.id ? 'opacity-40' : 'opacity-100'}
              `}
            >
              {editingProjectId === p.id ? (
                <input 
                  autoFocus
                  value={editingProjectName}
                  onChange={(e) => setEditingProjectName(e.target.value)}
                  onBlur={saveEditingProjectName}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') saveEditingProjectName();
                    else if (e.key === 'Escape') setEditingProjectId(null);
                  }}
                  className="w-full bg-[#002244] text-white px-2 py-1 rounded outline-none border border-blue-400 text-sm font-semibold"
                />
              ) : (
                <span onDoubleClick={() => startEditingProjectName(p.id, p.name)} className="truncate font-semibold text-sm select-none flex-1" title="Klik 2x untuk ubah nama">{p.name}</span>
              )}
              
              {editingProjectId !== p.id && (
                <button onClick={(e) => handleDeleteProject(p.id, e)} className="opacity-0 group-hover:opacity-100 text-blue-200 hover:text-red-400 transition bg-black/20 rounded-md p-1 ml-2 shrink-0"><Trash2 size={14} /></button>
              )}
            </div>
          ))}
        </div>

        <div className="p-4 border-t border-blue-800 bg-[#002244]">
          <button onClick={handleAddNewProject} className="w-full py-2.5 bg-white/10 hover:bg-blue-600 border border-white/10 hover:border-blue-400 rounded-lg flex items-center justify-center gap-2 font-bold text-sm transition shadow-sm">
            <Plus size={16} /> Proyek Baru
          </button>
        </div>
      </div>

      {/* ================= MAIN WORKSPACE AREA ================= */}
      <div className={`flex-1 flex flex-col min-w-0 h-full relative transition-opacity duration-200 ${isTransitioning ? 'opacity-0' : 'opacity-100'}`}>

        {/* --- SMART SYNC MODAL --- */}
        {syncModal.show && (
          <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center z-70 animate-in fade-in">
            <div className="bg-white p-6 rounded-xl shadow-2xl max-w-md w-full border-t-4 border-blue-500">
              <div className="flex items-center gap-3 text-slate-800 mb-2">
                <Settings2 className="text-blue-500" />
                <h3 className="font-bold text-lg">Pengaturan Template</h3>
              </div>
              <p className="text-sm text-slate-600 mb-4">Anda baru saja mengubah format template standar. Di mana perubahan ini ingin diterapkan?</p>
              <div className="flex flex-col gap-3 mb-5">
                <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-slate-700">
                    <input type="radio" checked={syncChoice === 'single'} onChange={() => setSyncChoice('single')} className="w-4 h-4 accent-blue-600" />
                    Ubah di halaman ini saja
                </label>
                <label className="flex items-center gap-3 cursor-pointer text-sm font-medium text-slate-700">
                    <input type="radio" checked={syncChoice === 'all'} onChange={() => setSyncChoice('all')} className="w-4 h-4 accent-blue-600" />
                    Terapkan di semua halaman
                </label>
              </div>
              <div className="border-t border-slate-100 pt-4 mb-5">
                <label className={`flex items-center gap-3 text-sm font-medium ${!syncChoice ? 'text-slate-400 cursor-not-allowed' : 'text-slate-700 cursor-pointer'}`}>
                    <input type="checkbox" disabled={!syncChoice} checked={dontAskAgain} onChange={e => setDontAskAgain(e.target.checked)} className="w-4 h-4 accent-blue-600 disabled:opacity-50" />
                    Jangan tanyakan lagi
                </label>
              </div>
              <div className="flex justify-end gap-2">
                <button onClick={() => setSyncModal({ show: false, fieldType: null, fieldKey: "", value: "", pendingFieldId: "" })} className="px-4 py-2 text-sm font-semibold text-slate-500 hover:bg-slate-100 rounded">Batal (Esc)</button>
                <button disabled={!syncChoice} onClick={confirmSyncModal} className="px-5 py-2 text-sm font-semibold bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-slate-300 transition shadow-sm">
                  Lanjutkan (Enter)
                </button>
              </div>
            </div>
          </div>
        )}

        {/* --- CUSTOM MODERN DIALOG SYSTEM --- */}
        {customDialog.show && (
          <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center z-60 animate-in fade-in">
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
          <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm flex items-center justify-center z-50">
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
        <div className="w-full bg-white border-b border-slate-300 px-4 py-3 flex justify-between items-center shrink-0 shadow-sm z-40 relative">
          <div className="flex items-center gap-3">
            {!isSidebarOpen && (
              <button onClick={() => setIsSidebarOpen(true)} className="p-2 hover:bg-slate-100 rounded-lg text-slate-600 transition">
                <Menu size={20} />
              </button>
            )}
            <div className="w-8 h-8 rounded-full bg-linear-to-br from-blue-900 to-red-600 shadow-md"></div>
            <div>
              <h1 className="text-lg font-bold text-slate-800 leading-tight">Winteq BOM Automator</h1>
              <p className="text-[10px] font-semibold text-slate-400 uppercase tracking-widest">{projects.find(p => p.id === activeProjectId)?.name || "Workspace"}</p>
            </div>
          </div>
          
          {/* ACTION BUTTONS */}
          <div className="flex items-center gap-3 relative">
            <div className="flex items-center border border-slate-300 rounded-lg overflow-hidden bg-white shadow-sm">
              <button onClick={handleUndo} disabled={historyIndex <= 0} className="p-2 text-slate-600 hover:bg-slate-100 disabled:opacity-30 transition" title="Undo (Mundur)"><Undo2 size={16} /></button>
              <div className="w-px h-5 bg-slate-300"></div>
              <button onClick={handleRedo} disabled={historyIndex >= history.length - 1} className="p-2 text-slate-600 hover:bg-slate-100 disabled:opacity-30 transition" title="Redo (Maju)"><Redo2 size={16} /></button>
            </div>

            <button onClick={clearTable} className={`flex items-center gap-2 px-3 py-2 text-sm font-bold rounded-lg transition border border-red-200 text-red-600 hover:bg-red-50 bg-white shadow-sm`} title="Hapus semua isi tabel ini">
              <Eraser size={16} /> Kosongkan
            </button>
            
            <button onClick={() => setIsThickGrid(!isThickGrid)} className={`flex items-center gap-2 px-3 py-2 text-sm font-bold rounded-lg transition border ${isThickGrid ? 'bg-blue-100 text-blue-800 border-blue-300 shadow-inner' : 'bg-white text-slate-600 border-slate-300 hover:bg-slate-50 shadow-sm'}`}>
              <Grid size={16} /> Grid
            </button>
            
            <div className="h-8 w-px bg-slate-300 mx-1"></div>

            <div>
              <input type="file" id="import-excel" accept=".xlsx" className="hidden" onChange={handleImportExcel} />
              <button onClick={handleImportClick} className={`flex items-center gap-2 px-3 py-2 text-sm font-bold rounded-lg transition border border-blue-200 text-blue-700 hover:bg-blue-50 bg-white shadow-sm`} title="Lanjutkan pekerjaan dari file Excel Winteq">
                <Upload size={16} /> Upload Excel
              </button>
            </div>

            <div className="[&>label]:flex! [&>label]:items-center! [&>label]:gap-2! [&>label]:px-3! [&>label]:py-2! [&>label]:text-sm! [&>label]:font-bold! [&>label]:rounded-lg! [&>label]:transition! [&>label]:border! [&>label]:border-slate-300! [&>label]:text-slate-600! [&>label]:hover:bg-slate-50! [&>label]:bg-white! [&>label]:shadow-sm! [&>label]:cursor-pointer!">
              <ExcelUploader onDataLoaded={handleDataLoaded} />
            </div>
            
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

        {/* WORKSPACE AREA PADA PROYEK AKTIF */}
        <div className="flex-1 w-full p-2 overflow-hidden flex flex-col bg-[#f8f9fa]">
          <div className="flex-1 bg-white border border-slate-300 flex flex-col overflow-hidden shadow-sm">
            
            {/* SPREADSHEET HEADER AREA */}
            <div className="shrink-0 bg-white p-0">
              <table className="w-full border-collapse text-xs">
                <tbody>
                  {/* BARIS 1 */}
                  <tr>
                    <td rowSpan={4} colSpan={3} className={`${headerBorder} p-2 w-[35%] align-middle`}>
                      <div className="flex flex-col items-center justify-center">
                        <div className="flex items-center gap-2 pointer-events-none">
                          <img src="/logo-astra.png" alt="Logo Astra Otoparts" className="w-12 h-12 object-contain" />
                          <div className="flex flex-col">
                            <span className="font-extrabold text-xl tracking-tighter leading-none italic">
                              <span className="text-red-600">ASTRA</span>{" "}
                              <span className="text-black">Otoparts</span>
                            </span>
                            <span className="font-bold text-[#1e3a8a] text-xs italic">DIVISI WINTEQ</span>
                          </div>
                        </div>
                        <input id="lbl-title" value={currentHeaderLabels.title} onChange={e => updateHeaderLabel("title", e.target.value)} {...getTemplateHandlers("header", "title", "lbl-title")} autoComplete="off" spellCheck="false" className="font-bold text-lg mt-2 tracking-widest text-slate-800 text-center w-full bg-transparent outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 uppercase" />
                      </div>
                    </td>
                    <td colSpan={2} className={`${headerBorder} bg-slate-100 p-0`}>
                      <input id="lbl-machine" value={currentHeaderLabels.machineName} onChange={e => updateHeaderLabel("machineName", e.target.value)} {...getTemplateHandlers("header", "machineName", "lbl-machine")} autoComplete="off" spellCheck="false" className="w-full h-full p-1 text-center font-semibold text-slate-600 bg-transparent outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 uppercase" />
                    </td>
                    <td className={`${headerBorder} bg-slate-100 p-0 w-24`}>
                      <input id="lbl-designed" value={currentHeaderLabels.designed} onChange={e => updateHeaderLabel("designed", e.target.value)} {...getTemplateHandlers("header", "designed", "lbl-designed")} autoComplete="off" spellCheck="false" className="w-full h-full p-1 text-center font-semibold text-slate-600 bg-transparent outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 uppercase" />
                    </td>
                    <td className={`${headerBorder} bg-slate-100 p-0 w-24`}>
                      <input id="lbl-elc" value={currentHeaderLabels.elc} onChange={e => updateHeaderLabel("elc", e.target.value)} {...getTemplateHandlers("header", "elc", "lbl-elc")} autoComplete="off" spellCheck="false" className="w-full h-full p-1 text-center font-semibold text-slate-600 bg-transparent outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 uppercase" />
                    </td>
                    <td className={`${headerBorder} bg-slate-100 p-0 w-24`}>
                      <input id="lbl-checked" value={currentHeaderLabels.checked} onChange={e => updateHeaderLabel("checked", e.target.value)} {...getTemplateHandlers("header", "checked", "lbl-checked")} autoComplete="off" spellCheck="false" className="w-full h-full p-1 text-center font-semibold text-slate-600 bg-transparent outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 uppercase" />
                    </td>
                    <td className={`${headerBorder} bg-slate-100 p-0 w-24`}>
                      <input id="lbl-approved" value={currentHeaderLabels.approved} onChange={e => updateHeaderLabel("approved", e.target.value)} {...getTemplateHandlers("header", "approved", "lbl-approved")} autoComplete="off" spellCheck="false" className="w-full h-full p-1 text-center font-semibold text-slate-600 bg-transparent outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 uppercase" />
                    </td>
                  </tr>
                  
                  {/* BARIS 2 (Date Memiliki rowSpan=2) */}
                  <tr>
                    <td colSpan={2} className={`${headerBorder} p-0 h-8`}>
                      <input id="proj-machine" value={currentProjectInfo.machineName} onChange={(e) => updateProjectInfo("machineName", e.target.value)} {...getTemplateHandlers("project", "machineName", "proj-machine")} autoComplete="off" spellCheck="false" className="w-full h-full text-center outline-none font-bold text-slate-700 uppercase focus:bg-blue-50 focus:ring-1 focus:ring-blue-500" />
                    </td>
                    <td rowSpan={2} className={`${headerBorder} p-0 align-top hover:bg-blue-50 transition`}><input id="proj-date1" value={currentProjectInfo.date1} onChange={(e) => updateProjectInfo("date1", e.target.value)} {...getTemplateHandlers("project", "date1", "proj-date1")} autoComplete="off" spellCheck="false" className="w-full h-full px-1 pb-4 text-[9px] text-slate-500 outline-none focus:ring-1 focus:ring-blue-500 bg-transparent text-center" /></td>
                    <td rowSpan={2} className={`${headerBorder} p-0 align-top hover:bg-blue-50 transition`}><input id="proj-date2" value={currentProjectInfo.date2} onChange={(e) => updateProjectInfo("date2", e.target.value)} {...getTemplateHandlers("project", "date2", "proj-date2")} autoComplete="off" spellCheck="false" className="w-full h-full px-1 pb-4 text-[9px] text-slate-500 outline-none focus:ring-1 focus:ring-blue-500 bg-transparent text-center" /></td>
                    <td rowSpan={2} className={`${headerBorder} p-0 align-top hover:bg-blue-50 transition`}><input id="proj-date3" value={currentProjectInfo.date3} onChange={(e) => updateProjectInfo("date3", e.target.value)} {...getTemplateHandlers("project", "date3", "proj-date3")} autoComplete="off" spellCheck="false" className="w-full h-full px-1 pb-4 text-[9px] text-slate-500 outline-none focus:ring-1 focus:ring-blue-500 bg-transparent text-center" /></td>
                    <td rowSpan={2} className={`${headerBorder} p-0 align-top hover:bg-blue-50 transition bg-white`}><input id="proj-date4" value={currentProjectInfo.date4} onChange={(e) => updateProjectInfo("date4", e.target.value)} {...getTemplateHandlers("project", "date4", "proj-date4")} autoComplete="off" spellCheck="false" className="w-full h-full px-1 pb-4 text-[9px] text-slate-500 outline-none focus:ring-1 focus:ring-blue-500 bg-transparent text-center" /></td>
                  </tr>

                  {/* BARIS 3 */}
                  <tr>
                    <td className={`${headerBorder} bg-slate-50 p-0 w-24`}>
                      <input id="lbl-designId" value={currentHeaderLabels.designId} onChange={e => updateHeaderLabel("designId", e.target.value)} {...getTemplateHandlers("header", "designId", "lbl-designId")} autoComplete="off" spellCheck="false" className="w-full h-full p-1 font-medium text-slate-600 bg-transparent outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 uppercase" />
                    </td>
                    <td className={`${headerBorder} p-0`}><input id="proj-designId" value={currentProjectInfo.designId} onChange={(e) => updateProjectInfo("designId", e.target.value)} {...getTemplateHandlers("project", "designId", "proj-designId")} autoComplete="off" spellCheck="false" className="w-full h-full px-2 outline-none focus:bg-blue-50 focus:ring-1 focus:ring-blue-500" /></td>
                  </tr>

                  {/* BARIS 4 (Inisial Tanda Tangan) */}
                  <tr>
                    <td className={`${headerBorder} bg-slate-50 p-0`}>
                      <input id="lbl-customer" value={currentHeaderLabels.customer} onChange={e => updateHeaderLabel("customer", e.target.value)} {...getTemplateHandlers("header", "customer", "lbl-customer")} autoComplete="off" spellCheck="false" className="w-full h-full p-1 font-medium text-slate-600 bg-transparent outline-none focus:bg-white focus:ring-1 focus:ring-blue-500 uppercase" />
                    </td>
                    <td className={`${headerBorder} p-0`}><input id="proj-customer" value={currentProjectInfo.customer} onChange={(e) => updateProjectInfo("customer", e.target.value)} {...getTemplateHandlers("project", "customer", "proj-customer")} autoComplete="off" spellCheck="false" className="w-full h-full px-2 outline-none focus:bg-blue-50 focus:ring-1 focus:ring-blue-500" /></td>
                    <td className={`${headerBorder} p-0 text-center h-8 hover:bg-blue-50 transition`}><input id="proj-designed" value={currentProjectInfo.designed} onChange={(e) => updateProjectInfo("designed", e.target.value)} {...getTemplateHandlers("project", "designed", "proj-designed")} autoComplete="off" spellCheck="false" className="w-full h-full text-center outline-none uppercase font-bold text-slate-700 bg-transparent focus:ring-1 focus:ring-blue-500" /></td>
                    <td className={`${headerBorder} p-0 text-center h-8 hover:bg-blue-50 transition`}><input id="proj-elc" value={currentProjectInfo.elc} onChange={(e) => updateProjectInfo("elc", e.target.value)} {...getTemplateHandlers("project", "elc", "proj-elc")} autoComplete="off" spellCheck="false" className="w-full h-full text-center outline-none uppercase font-bold text-slate-700 bg-transparent focus:ring-1 focus:ring-blue-500" /></td>
                    <td className={`${headerBorder} p-0 text-center h-8 hover:bg-blue-50 transition`}><input id="proj-checked" value={currentProjectInfo.checked} onChange={(e) => updateProjectInfo("checked", e.target.value)} {...getTemplateHandlers("project", "checked", "proj-checked")} autoComplete="off" spellCheck="false" className="w-full h-full text-center outline-none uppercase font-bold text-slate-700 bg-transparent focus:ring-1 focus:ring-blue-500" /></td>
                    <td className={`${headerBorder} p-0 text-center h-8 hover:bg-blue-50 transition bg-white`}><input id="proj-approved" value={currentProjectInfo.approved} onChange={(e) => updateProjectInfo("approved", e.target.value)} {...getTemplateHandlers("project", "approved", "proj-approved")} autoComplete="off" spellCheck="false" className="w-full h-full text-center outline-none uppercase font-bold text-slate-700 bg-transparent focus:ring-1 focus:ring-blue-500" /></td>
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

          {/* TABS SHEET BOTTOM AREA DENGAN DRAG AND DROP */}
          <div className="bg-slate-100 border-t border-slate-300 flex items-center shrink-0">
            <div className="flex bg-slate-200 overflow-x-auto w-full">
              <button onClick={addNewTab} className="px-3 py-2 text-slate-500 hover:text-slate-800 hover:bg-slate-300 flex items-center justify-center border-r border-slate-300">
                <Plus size={16} />
              </button>
              {tabs.map((tab) => (
                <div 
                  key={tab} 
                  draggable
                  onDragStart={(e) => setDraggedTab(tab)}
                  onDragOver={(e) => e.preventDefault()}
                  onDrop={(e) => {
                    e.preventDefault();
                    if (!draggedTab || draggedTab === tab) return;
                    const newTabs = [...tabs];
                    const draggedIdx = newTabs.indexOf(draggedTab);
                    const targetIdx = newTabs.indexOf(tab);
                    newTabs.splice(draggedIdx, 1);
                    newTabs.splice(targetIdx, 0, draggedTab);
                    setTabs(newTabs);
                    setDraggedTab(null);
                  }}
                  className={`group flex items-center border-r border-slate-300 text-xs font-medium cursor-grab active:cursor-grabbing transition max-w-37.5
                    ${activeTab === tab ? "bg-white text-green-700 shadow-[0_-2px_0_0_#16a34a_inset]" : "bg-slate-100 text-slate-600 hover:bg-slate-50"}
                    ${draggedTab === tab ? "opacity-50" : "opacity-100"}
                  `}
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

    </div>
  );
}