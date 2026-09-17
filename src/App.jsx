import React, { useState, useRef, useCallback, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import mammoth from "mammoth";
import {
  Upload, FileText, Image as ImageIcon, Music, Video, Archive, Code2,
  Table as TableIcon, File as FileIcon, X, Search, ZoomIn, ZoomOut,
  RotateCw, Download, Lock, Trash2, Moon, Sun, ChevronRight, ChevronDown,
  AlertTriangle, Copy, Check, FolderOpen, Clock, Eye, EyeOff
} from "lucide-react";

const CATEGORY = {
  pdf: "documents", doc: "documents", docx: "documents", odt: "documents",
  rtf: "documents", txt: "documents", md: "documents",
  csv: "sheets", xls: "sheets", xlsx: "sheets", ods: "sheets",
  ppt: "presentations", pptx: "presentations", odp: "presentations",
  jpg: "images", jpeg: "images", png: "images", webp: "images", gif: "images",
  bmp: "images", svg: "images", tiff: "images", tif: "images", ico: "images", avif: "images",
  mp3: "audio", wav: "audio", ogg: "audio", flac: "audio", m4a: "audio", aac: "audio", opus: "audio",
  mp4: "video", webm: "video", mkv: "video", mov: "video", avi: "video", m4v: "video", "3gp": "video",
  zip: "archives", "7z": "archives", tar: "archives", gz: "archives",
  json: "code", xml: "code", yaml: "code", yml: "code", html: "code", htm: "code",
  css: "code", js: "code", jsx: "code", ts: "code", tsx: "code", py: "code", java: "code",
  c: "code", cpp: "code", h: "code", cs: "code", php: "code", sql: "code", sh: "code",
};

const CATEGORY_META = {
  documents: { label: "Documents", icon: FileText, color: "#4C8DFF" },
  sheets: { label: "Spreadsheets", icon: TableIcon, color: "#3FD0C9" },
  presentations: { label: "Presentations", icon: FileIcon, color: "#F2B84B" },
  images: { label: "Images", icon: ImageIcon, color: "#F2B84B" },
  audio: { label: "Audio", icon: Music, color: "#C77DFF" },
  video: { label: "Video", icon: Video, color: "#FF6B6B" },
  archives: { label: "Archives", icon: Archive, color: "#8B95A5" },
  code: { label: "Code & Data", icon: Code2, color: "#6EE7B7" },
  other: { label: "Other", icon: FileIcon, color: "#8B95A5" },
};

const TEXT_LIKE = new Set(["txt","md","json","xml","yaml","yml","html","htm","css","js","jsx","ts","tsx","py","java","c","cpp","h","cs","php","sql","sh","csv"]);
const NO_PREVIEW = new Set(["odt","rtf","ppt","pptx","odp","zip","7z","tar","gz","doc"]);

function extOf(name) {
  const n = name.toLowerCase();
  if (n.endsWith(".tar.gz")) return "tar.gz";
  const i = n.lastIndexOf(".");
  return i === -1 ? "" : n.slice(i + 1);
}
function categoryOf(ext) { return CATEGORY[ext] || "other"; }
function fmtSize(bytes) {
  if (bytes === 0) return "0 B";
  const u = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  return `${(bytes / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 1)} ${u[i]}`;
}
function fmtDate(date) {
  return new Date(date).toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
}

export default function OmniFile() {
  const [dark, setDark] = useState(true);
  const [files, setFiles] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [collapsed, setCollapsed] = useState({});
  const [showInfo, setShowInfo] = useState(true);
  const inputRef = useRef(null);
  const active = files.find((f) => f.id === activeId) || null;
  const theme = dark ? {
    bg: "#0B0E14", surface: "#11151D", surface2: "#161B25", surface3: "#1D222C",
    border: "#232A38", text: "#E6EDF3", muted: "#7D8A9E", accent: "#4C8DFF",
    success: "#3FBF6B", warning: "#F2B84B", error: "#FF6B6B",
  } : {
    bg: "#F5F6F8", surface: "#FFFFFF", surface2: "#F8F9FA", surface3: "#EDEFF2",
    border: "#E2E5EB", text: "#151B23", muted: "#5B6472", accent: "#2F6FE0",
    success: "#2E9B5A", warning: "#D4A017", error: "#D32F2F",
  };

  const ingest = useCallback((fileList) => {
    const list = Array.from(fileList);
    const entries = list.map((file) => {
      const ext = extOf(file.name);
      const category = categoryOf(ext);
      return {
        id: `${file.name}-${file.size}-${Math.random().toString(36).slice(2, 8)}`,
        file, ext, category, status: "loading", data: null, error: null, url: null, lastModified: file.lastModified,
      };
    });
    setFiles((prev) => [...entries, ...prev]);
    if (!activeId && entries.length) setActiveId(entries[0].id);
    entries.forEach(processFile);
  }, [activeId]);

  const updateFile = (id, patch) => setFiles((prev) => prev.map((f) => (f.id === id ? { ...f, ...patch } : f)));

  async function processFile(entry) {
    const { id, file, ext, category } = entry;
    try {
      if (["images", "audio", "video"].includes(category) || ext === "pdf") {
        const url = URL.createObjectURL(file);
        updateFile(id, { url, status: "ready" });
        return;
      }
      if (ext === "docx") {
        const buf = await file.arrayBuffer();
        const res = await mammoth.convertToHtml({ arrayBuffer: buf });
        updateFile(id, { data: { html: res.value }, status: "ready" });
        return;
      }
      if (ext === "xls" || ext === "xlsx" || ext === "ods") {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const sheets = wb.SheetNames.map((name) => ({
          name, rows: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false }),
        }));
        updateFile(id, { data: { sheets, activeSheet: 0 }, status: "ready" });
        return;
      }
      if (ext === "csv") {
        const text = await file.text();
        const parsed = Papa.parse(text, { skipEmptyLines: true });
        updateFile(id, { data: { sheets: [{ name: file.name, rows: parsed.data }], activeSheet: 0 }, status: "ready" });
        return;
      }
      if (TEXT_LIKE.has(ext) || category === "code") {
        const text = await file.text();
        updateFile(id, { data: { text }, status: "ready" });
        return;
      }
      updateFile(id, { status: "unsupported" });
    } catch (e) {
      updateFile(id, { status: "error", error: String(e.message || e) });
    }
  }

  const onDrop = (e) => { e.preventDefault(); setDragOver(false); if (e.dataTransfer.files?.length) ingest(e.dataTransfer.files); };
  const removeFile = (id) => {
    setFiles((prev) => {
      const f = prev.find((x) => x.id === id);
      if (f?.url) URL.revokeObjectURL(f.url);
      const next = prev.filter((x) => x.id !== id);
      if (activeId === id) setActiveId(next[0]?.id ?? null);
      return next;
    });
  };
  const clearAll = () => { files.forEach((f) => f.url && URL.revokeObjectURL(f.url)); setFiles([]); setActiveId(null); };
  const grouped = useMemo(() => { const g = {}; files.forEach((f) => { g[f.category] = g[f.category] || []; g[f.category].push(f); }); return g; }, [files]);
  const toggleGroup = (cat) => setCollapsed((c) => ({ ...c, [cat]: !c[cat] }));
  const toggleInfo = () => setShowInfo((s) => !s);

  return (
    <div style={{ background: theme.bg, color: theme.text }} className="min-h-screen w-full font-sans transition-colors duration-200">
      <style>{`
        @import url('https://fonts.cdnfonts.com/css/jetbrains-mono-2');
        .mono { font-family: 'JetBrains Mono', ui-monospace, monospace; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: ${theme.border}; border-radius: 4px; }
        ::-webkit-scrollbar-track { background: transparent; }
        * { scrollbar-width: thin; scrollbar-color: ${theme.border} transparent; }
      `}</style>
      <header style={{ borderColor: theme.border, background: theme.surface }} className="border-b sticky top-0 z-20 px-4 py-3 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div style={{ background: theme.accent }} className="w-9 h-9 rounded-lg flex items-center justify-center mono text-[10px] font-bold text-white shrink-0">OF</div>
          <div>
            <div className="font-semibold text-[16px] leading-tight tracking-tight">OmniFile</div>
            <div style={{ color: theme.muted }} className="text-[11px] leading-tight hidden sm:block">Open almost anything in your browser</div>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <div style={{ background: dark ? "#12261C" : "#EAF7EF", color: "#3FBF6B", borderColor: dark ? "#1E3A2A" : "#CFEFDA" }} className="hidden sm:flex items-center gap-1.5 text-[11px] px-2.5 py-1.5 rounded-full border mono">
            <Lock size={11} /> local only
          </div>
          <button onClick={() => setDark((d) => !d)} style={{ borderColor: theme.border, background: theme.surface2 }} className="w-9 h-9 rounded-lg border flex items-center justify-center hover:bg-opacity-80 transition-colors" aria-label="Toggle theme">
            {dark ? <Sun size={16} /> : <Moon size={16} />}
          </button>
        </div>
      </header>
      {files.length === 0 ? (
        <LandingArea theme={theme} dark={dark} dragOver={dragOver} setDragOver={setDragOver} onDrop={onDrop} inputRef={inputRef} ingest={ingest} />
      ) : (
        <div className="flex" style={{ minHeight: "calc(100vh - 65px)" }}>
          <Sidebar theme={theme} grouped={grouped} collapsed={collapsed} toggleGroup={toggleGroup} activeId={activeId} setActiveId={setActiveId} removeFile={removeFile} clearAll={clearAll} inputRef={inputRef} ingest={ingest} count={files.length} />
          <main className="flex-1 min-w-0 flex flex-col">
            {active ? (
              <FileWorkspace key={active.id} entry={active} theme={theme} dark={dark} removeFile={removeFile} showInfo={showInfo} toggleInfo={toggleInfo} />
            ) : (
              <div style={{ color: theme.muted }} className="flex-1 flex items-center justify-center text-sm"><div className="text-center"><FolderOpen size={48} className="mb-4 opacity-50" /><div>Pilih file di sidebar</div></div></div>
            )}
          </main>
        </div>
      )}
      <input ref={inputRef} type="file" multiple className="hidden" onChange={(e) => e.target.files?.length && ingest(e.target.files)} />
    </div>
  );
}

function LandingArea({ theme, dark, dragOver, setDragOver, onDrop, inputRef, ingest }) {
  return (
    <div className="px-4 py-10 sm:py-16 flex flex-col items-center">
      <div onDragOver={(e) => { e.preventDefault(); setDragOver(true); }} onDragLeave={() => setDragOver(false)} onDrop={onDrop} onClick={() => inputRef.current?.click()} style={{ borderColor: dragOver ? theme.accent : theme.border, background: dragOver ? (dark ? "#141C2B" : "#EFF4FF") : theme.surface, boxShadow: dragOver ? (dark ? "0 0 0 3px rgba(76, 141, 255, 0.15)" : "0 0 0 3px rgba(47, 111, 224, 0.15)") : "none" }} className="w-full max-w-2xl border-2 border-dashed rounded-2xl py-16 px-8 flex flex-col items-center gap-4 cursor-pointer transition-all duration-200">
        <div style={{ background: theme.accent + "20" }} className="w-16 h-16 rounded-full flex items-center justify-center"><Upload size={28} color={theme.accent} /></div>
        <div className="text-center"><div className="font-medium text-[16px] sm:text-[18px]">Drop your files here</div><div style={{ color: theme.muted }} className="text-[13px] sm:text-[14px] mt-1">or click to browse — multiple files supported</div></div>
        <button onClick={(e) => { e.stopPropagation(); inputRef.current?.click(); }} style={{ background: theme.accent, boxShadow: "0 2px 8px rgba(76, 141, 255, 0.3)" }} className="mt-2 text-white text-[13px] sm:text-[14px] font-medium px-5 py-2.5 rounded-lg hover:opacity-90 transition-opacity"><Upload size={14} className="inline-block mr-1.5" />Choose Files</button>
      </div>
      <div style={{ color: theme.muted }} className="text-[12px] mt-6 text-center max-w-md">Supported formats: <div className="mt-2 flex flex-wrap justify-center gap-x-4 gap-y-1"><span className="text-[11px] px-2 py-0.5 bg-opacity-50 bg-current rounded">Documents</span><span className="text-[11px] px-2 py-0.5 bg-opacity-50 bg-current rounded">Spreadsheets</span><span className="text-[11px] px-2 py-0.5 bg-opacity-50 bg-current rounded">Images</span><span className="text-[11px] px-2 py-0.5 bg-opacity-50 bg-current rounded">Audio</span><span className="text-[11px] px-2 py-0.5 bg-opacity-50 bg-current rounded">Video</span></div></div>
      <div style={{ borderColor: theme.border, background: theme.surface2 }} className="mt-8 w-full max-w-2xl border rounded-xl p-4 flex items-start gap-3"><Lock size={18} color={theme.accent} className="mt-0.5 shrink-0" /><div className="text-[12.5px]" style={{ color: theme.muted }}><span style={{ color: theme.text }} className="font-medium">Your files stay on your device.</span> Semua pemrosesan berjalan di browser — tidak ada data yang dikirim ke server.</div></div>
    </div>
  );
}

function Sidebar({ theme, grouped, collapsed, toggleGroup, activeId, setActiveId, removeFile, clearAll, inputRef, ingest, count }) {
  return (
    <aside style={{ borderColor: theme.border, background: theme.surface }} className="w-[260px] shrink-0 border-r hidden md:flex flex-col">
      <div className="p-3 flex gap-2">
        <button onClick={() => inputRef.current?.click()} style={{ borderColor: theme.border, background: theme.surface2 }} className="flex-1 border rounded-lg text-[12.5px] py-2 flex items-center justify-center gap-1.5 font-medium hover:bg-opacity-80 transition-colors"><Upload size={14} /> Add file</button>
        <button onClick={clearAll} style={{ borderColor: theme.border, background: theme.surface2, color: theme.error }} className="border rounded-lg px-2.5 flex items-center justify-center hover:bg-opacity-80 transition-colors" title="Clear all opened files"><Trash2 size={14} /></button>
      </div>
      <div style={{ color: theme.muted }} className="px-3 pb-2 text-[11px] mono uppercase tracking-wide flex items-center gap-1.5"><Clock size={11} /> {count} file{count > 1 ? 's' : ''} opened</div>
      <div className="flex-1 overflow-y-auto px-2 pb-3">
        {Object.entries(grouped).map(([cat, list]) => {
          const meta = CATEGORY_META[cat];
          const Icon = meta.icon;
          const isCollapsed = collapsed[cat];
          return (
            <div key={cat} className="mb-1">
              <button onClick={() => toggleGroup(cat)} className="w-full flex items-center gap-2 px-2 py-1.5 rounded-md text-[12px] font-medium hover:bg-opacity-50 transition-colors" style={{ color: theme.muted, background: theme.surface2 }}>
                {isCollapsed ? <ChevronRight size={13} /> : <ChevronDown size={13} />}<Icon size={14} color={meta.color} /><span>{meta.label}</span><span className="ml-auto mono text-[10.5px]">{list.length}</span>
              </button>
              {!isCollapsed && <div className="flex flex-col gap-0.5 mt-0.5">{list.map((f) => (<FileRow key={f.id} f={f} theme={theme} active={f.id === activeId} onClick={() => setActiveId(f.id)} onRemove={() => removeFile(f.id)} />))}</div>}
            </div>
          );
        })}
      </div>
    </aside>
  );
}

function FileRow({ f, theme, active, onClick, onRemove }) {
  const meta = CATEGORY_META[f.category];
  return (
    <div onClick={onClick} style={{ background: active ? theme.accent + "1A" : "transparent", borderColor: active ? theme.accent : theme.border }} className="group flex items-center gap-2.5 px-2.5 py-1.5 rounded-md cursor-pointer border text-left transition-colors">
      <div style={{ color: meta.color }} className="shrink-0"><meta.icon size={14} /></div>
      <div className="min-w-0 flex-1">
        <div className="text-[12.5px] truncate" style={{ color: active ? theme.accent : theme.text }}>{f.file.name}</div>
        <div style={{ color: theme.muted }} className="text-[10.5px] mono">{fmtSize(f.file.size)} · .{f.ext || "?"}</div>
      </div>
      <button onClick={(e) => { e.stopPropagation(); onRemove(); }} className="opacity-0 group-hover:opacity-100 shrink-0 p-1 rounded hover:bg-red-500 hover:text-white transition-colors" style={{ color: theme.muted }}><X size={13} /></button>
    </div>
  );
}

function FileWorkspace({ entry, theme, dark, removeFile, showInfo, toggleInfo }) {
  const meta = CATEGORY_META[entry.category];
  return (
    <div className="flex-1 flex flex-col min-w-0">
      <div style={{ borderColor: theme.border, background: theme.surface }} className="border-b px-4 py-2 flex items-center gap-2">
        <div style={{ background: meta.color + "22", color: meta.color }} className="mono text-[10px] font-bold px-2 py-1 rounded shrink-0 uppercase">.{entry.ext || "?"}</div>
        <div className="min-w-0 flex-1"><div className="text-[13.5px] font-medium truncate">{entry.file.name}</div><div style={{ color: theme.muted }} className="text-[10.5px] mono truncate">{fmtSize(entry.file.size)}</div></div>
        <div className="flex items-center gap-1">
          <button onClick={toggleInfo} style={{ borderColor: theme.border, color: showInfo ? theme.accent : theme.muted }} className="border rounded-md p-1.5 hover:bg-opacity-50 transition-colors" title="Toggle info panel"><Eye size={14} /></button>
          <a href={entry.url || undefined} download={entry.file.name} onClick={(e) => { if (!entry.url) { e.preventDefault(); downloadOriginal(entry.file); } }} style={{ borderColor: theme.border, background: theme.surface2 }} className="border rounded-md p-1.5 hover:bg-opacity-80 transition-colors" title="Download"><Download size={14} /></a>
          <button onClick={() => removeFile(entry.id)} style={{ borderColor: theme.border, color: theme.error }} className="border rounded-md p-1.5 hover:bg-opacity-80 transition-colors" title="Remove file"><X size={14} /></button>
        </div>
      </div>
      <div className="flex-1 flex flex-col md:flex-row min-h-0">
        <div className="flex-1 min-w-0 overflow-auto"><ViewerRouter entry={entry} theme={theme} dark={dark} /></div>
        {showInfo && <InfoPanel entry={entry} theme={theme} />}
      </div>
    </div>
  );
}

function downloadOriginal(file) {
  const url = URL.createObjectURL(file);
  const a = document.createElement("a");
  a.href = url; a.download = file.name; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function InfoPanel({ entry, theme }) {
  const { file, ext, category, data, lastModified } = entry;
  const meta = CATEGORY_META[category];
  const rows = [
    ["Name", file.name],
    ["Type", meta.label],
    ["Extension", `.${ext || "unknown"}`],
    ["MIME type", file.type || "—"],
    ["Size", fmtSize(file.size)],
    ["Last modified", fmtDate(lastModified)],
  ];
  if (data?.sheets) rows.push(["Sheets", data.sheets.length]);
  if (entry.imgMeta) rows.push(["Dimensions", `${entry.imgMeta.w} × ${entry.imgMeta.h}px`]);
  if (data?.text) { const lines = data.text.split('\n').length; const chars = data.text.length; rows.push(["Lines", lines]); rows.push(["Characters", chars]); }
  return (
    <div style={{ borderColor: theme.border, background: theme.surface2 }} className="md:w-[240px] shrink-0 border-t md:border-t-0 md:border-l p-4">
      <div style={{ color: theme.muted }} className="text-[11px] mono uppercase tracking-wide mb-3">File information</div>
      <div className="flex flex-col gap-3">
        {rows.map(([k, v]) => (<div key={k} className="border-b border-current pb-2 last:border-0"><div style={{ color: theme.muted }} className="text-[10.5px]">{k}</div><div className="text-[12.5px] break-words font-medium">{v}</div></div>))}
      </div>
    </div>
  );
}

function ViewerRouter({ entry, theme, dark }) {
  const { ext, category, status, url, data, error } = entry;
  if (status === "loading") return <CenterMsg theme={theme} text="Loading file..." />;
  if (status === "error") return <FallbackViewer entry={entry} theme={theme} reason={`Failed to read file: ${error}`} />;
  if (ext === "pdf" && url) return <PdfViewer url={url} theme={theme} />;
  if (category === "images" && url) return <ImageViewer url={url} file={entry.file} theme={theme} entry={entry} />;
  if (category === "audio" && url) return <MediaViewer url={url} kind="audio" theme={theme} />;
  if (category === "video" && url) return <MediaViewer url={url} kind="video" theme={theme} />;
  if (ext === "docx" && data?.html) return <DocxViewer html={data.html} theme={theme} dark={dark} />;
  if (data?.sheets) return <SheetViewer entry={entry} theme={theme} />;
  if (data?.text !== undefined) return <TextViewer text={data.text} ext={ext} theme={theme} dark={dark} />;
  if (NO_PREVIEW.has(ext) || status === "unsupported") return <FallbackViewer entry={entry} theme={theme} reason="This format cannot be previewed directly in the browser." />;
  return <CenterMsg theme={theme} text="Preparing preview..." />;
}

function CenterMsg({ theme, text }) { return <div style={{ color: theme.muted }} className="h-full flex items-center justify-center text-[13px] p-8 text-center">{text}</div>; }
function FallbackViewer({ entry, theme, reason }) {
  return (
    <div className="h-full flex flex-col items-center justify-center p-8 text-center gap-4">
      <div style={{ background: theme.surface3, borderColor: theme.border }} className="w-16 h-16 rounded-full border flex items-center justify-center"><AlertTriangle size={24} color="#F2B84B" /></div>
      <div><div className="font-medium text-[15px] mono">.{entry.ext || "unknown"}</div><div style={{ color: theme.muted }} className="text-[12.5px] mt-1 max-w-xs">{reason}</div></div>
      <button onClick={() => downloadOriginal(entry.file)} style={{ background: theme.accent }} className="text-white text-[12.5px] font-medium px-5 py-2.5 rounded-lg flex items-center gap-1.5 mt-1 hover:opacity-90 transition-opacity"><Download size={14} /> Download / Open With External App</button>
    </div>
  );
}

function PdfViewer({ url, theme }) { return <div className="h-full flex flex-col"><iframe title="pdf-viewer" src={url} className="w-full flex-1 border-0" style={{ minHeight: "70vh", background: "#fff" }} /></div>; }
function ImageViewer({ url, file, theme, entry }) {
  const [zoom, setZoom] = useState(1);
  const [rotate, setRotate] = useState(0);
  const handleLoad = (e) => { entry.imgMeta = { w: e.target.naturalWidth, h: e.target.naturalHeight }; };
  return (
    <div className="h-full flex flex-col">
      <div style={{ borderColor: theme.border, background: theme.surface2 }} className="border-b px-3 py-2 flex items-center gap-1.5">
        <IconBtn theme={theme} onClick={() => setZoom((z) => Math.max(0.25, z - 0.25))} icon={ZoomOut} tooltip="Zoom out" />
        <div style={{ color: theme.muted }} className="text-[11.5px] mono w-12 text-center">{Math.round(zoom * 100)}%</div>
        <IconBtn theme={theme} onClick={() => setZoom((z) => Math.min(4, z + 0.25))} icon={ZoomIn} tooltip="Zoom in" />
        <IconBtn theme={theme} onClick={() => setRotate((r) => (r + 90) % 360)} icon={RotateCw} tooltip="Rotate" />
        <button onClick={() => { setZoom(1); setRotate(0); }} style={{ color: theme.muted }} className="text-[11.5px] ml-1 hover:text-current">Reset</button>
      </div>
      <div className="flex-1 overflow-auto flex items-center justify-center p-6" style={{ background: "#0000000A" }}>
        <img src={url} alt={file.name} onLoad={handleLoad} style={{ transform: `scale(${zoom}) rotate(${rotate}deg)`, transition: "transform .15s", maxWidth: zoom === 1 ? "100%" : "none" }} className="max-h-[70vh] object-contain" />
      </div>
    </div>
  );
}
function IconBtn({ theme, onClick, icon: Icon, tooltip }) {
  return <button onClick={onClick} style={{ borderColor: theme.border }} className="border rounded-md p-1.5 group relative" title={tooltip}><Icon size={15} />{tooltip && <span className="absolute -top-7 left-1/2 -translate-x-1/2 bg-black text-white text-[10px] px-2 py-0.5 rounded opacity-0 group-hover:opacity-100 transition-opacity whitespace-nowrap">{tooltip}</span>}</button>;
}
function MediaViewer({ url, kind, theme }) {
  return (
    <div className="h-full flex items-center justify-center p-8">
      {kind === "video" ? <video src={url} controls className="max-w-full max-h-[65vh] rounded-lg shadow-lg" /> : <div style={{ background: theme.surface3, borderColor: theme.border }} className="w-full max-w-md border rounded-xl p-6 flex flex-col items-center gap-4"><div style={{ background: theme.accent + "20" }} className="w-16 h-16 rounded-full flex items-center justify-center"><Music size={28} color={theme.accent} /></div><audio src={url} controls className="w-full" /></div>}
    </div>
  );
}
function DocxViewer({ html, theme, dark }) {
  return <div className="p-6 flex justify-center"><div style={{ background: dark ? "#1A1F29" : "#fff", color: dark ? "#E6EDF3" : "#151B23", borderColor: theme.border }} className="w-full max-w-3xl border rounded-lg p-8 text-[14px] leading-relaxed doc-content" dangerouslySetInnerHTML={{ __html: html }} /></div>;
}
function SheetViewer({ entry, theme }) {
  const { data } = entry;
  const [activeSheet, setActiveSheet] = useState(0);
  const [query, setQuery] = useState("");
  const sheet = data.sheets[activeSheet];
  const rows = sheet?.rows || [];
  const header = rows[0] || [];
  const body = rows.slice(1);
  const filtered = query ? body.filter((r) => r.some((c) => String(c ?? "").toLowerCase().includes(query.toLowerCase()))) : body;
  return (
    <div className="h-full flex flex-col">
      {data.sheets.length > 1 && <div style={{ borderColor: theme.border, background: theme.surface2 }} className="border-b flex gap-1 px-2 pt-2 overflow-x-auto">{data.sheets.map((s, i) => <button key={s.name} onClick={() => setActiveSheet(i)} style={{ background: i === activeSheet ? theme.surface : "transparent", color: i === activeSheet ? theme.text : theme.muted, borderColor: theme.border }} className="text-[12px] px-3 py-1.5 rounded-t-md border border-b-0 shrink-0 hover:bg-opacity-50 transition-colors">{s.name}</button>)}</div>}
      <div style={{ borderColor: theme.border, background: theme.surface2 }} className="border-b px-3 py-2 flex items-center gap-2"><Search size={14} color={theme.muted} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search in table..." style={{ color: theme.text }} className="bg-transparent text-[12.5px] outline-none flex-1" /><span style={{ color: theme.muted }} className="text-[11px] mono">{filtered.length} row{filtered.length > 1 ? 's' : ''}</span></div>
      <div className="flex-1 overflow-auto">
        <table className="w-full text-[12px] border-collapse">
          <thead style={{ background: theme.surface3, borderColor: theme.border }} className="sticky top-0 z-10"><tr>{header.map((h, i) => <th key={i} style={{ borderColor: theme.border, color: theme.muted }} className="border-b px-3 py-2 text-left font-medium whitespace-nowrap">{String(h ?? "")}</th>)}</tr></thead>
          <tbody>{filtered.map((r, ri) => <tr key={ri} style={{ borderColor: theme.border, background: ri % 2 ? theme.surface2 : 'transparent' }} className="border-b">{header.map((_, ci) => <td key={ci} className="px-3 py-2 whitespace-nowrap">{String(r[ci] ?? "")}</td>)}</tr>)}</tbody>
        </table>
        {filtered.length === 0 && <div style={{ color: theme.muted }} className="p-4 text-center text-[12px]">No matching rows found</div>}
      </div>
    </div>
  );
}
function TextViewer({ text, ext, theme, dark }) {
  const [query, setQuery] = useState("");
  const [copied, setCopied] = useState(false);
  const isJson = ext === "json";
  const isMd = ext === "md";
  const display = useMemo(() => { if (isJson) { try { const parsed = JSON.parse(text); return JSON.stringify(parsed, null, 2); } catch { return text; } } return text; }, [text, isJson]);
  const lines = display.split("\n");
  const copy = () => { navigator.clipboard?.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 1200); };
  return (
    <div className="h-full flex flex-col">
      <div style={{ borderColor: theme.border, background: theme.surface2 }} className="border-b px-3 py-2 flex items-center gap-2"><Search size={14} color={theme.muted} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search text..." style={{ color: theme.text }} className="bg-transparent text-[12.5px] outline-none flex-1" /><button onClick={copy} style={{ color: theme.muted }} className="text-[11.5px] flex items-center gap-1 hover:text-current">{copied ? <Check size={13} /> : <Copy size={13} />} {copied ? "Copied!" : "Copy"}</button></div>
      {isMd ? <div className="p-6 overflow-auto flex-1"><MarkdownLite text={text} theme={theme} /></div> : <div className="overflow-auto flex-1"><table className="w-full text-[12px] mono border-collapse"><tbody>{lines.map((line, i) => { const match = query && line.toLowerCase().includes(query.toLowerCase()); return <tr key={i} style={{ background: match ? theme.accent + "1A" : "transparent" }}><td style={{ color: theme.muted }} className="select-none text-right pr-3 pl-4 py-0.5 align-top w-10">{i + 1}</td><td className="py-0.5 pr-4 whitespace-pre-wrap break-all">{line || " "}</td></tr>; })}</tbody></table></div>}
    </div>
  );
}
function MarkdownLite({ text, theme }) {
  const html = useMemo(() => {
    return text.split(/\n{2,}/).map((block) => {
      const b = block.trim();
      if (!b) return '<br/>';
      if (/^### /.test(b)) return `<h3 style="font-size:15px;font-weight:600;margin:12px 0 4px;color:${theme.text}">${b.slice(4)}</h3>`;
      if (/^## /.test(b)) return `<h2 style="font-size:17px;font-weight:600;margin:14px 0 6px;color:${theme.text}">${b.slice(3)}</h2>`;
      if (/^# /.test(b)) return `<h1 style="font-size:20px;font-weight:700;margin:16px 0 8px;color:${theme.text}">${b.slice(2)}</h1>`;
      if (/^[-*] /.test(b)) { const items = b.split("\n").map((l) => `<li style="color:${theme.text}">${l.replace(/^[-*] /, "")}</li>`).join(""); return `<ul style="margin:6px 0;padding-left:20px;list-style:disc">${items}</ul>`; }
      if (/^`{3}/.test(b)) { const lang = b.match(/^`{3}(\w*)/)?.[1] || ''; const code = b.slice(b.indexOf('\n') + 1, b.lastIndexOf('`{3}')).trim(); return `<pre style="background:${theme.surface3};padding:12px;border-radius:6px;overflow-x:auto;color:${theme.text}"><code>${code}</code></pre>`; }
      return `<p style="margin:8px 0;line-height:1.6;color:${theme.text}">${b.replace(/\n/g, "<br/>")}</p>`;
    }).join("");
  }, [text, theme.text, theme.surface3]);
  return <div className="text-[13.5px] max-w-3xl mx-auto prose" dangerouslySetInnerHTML={{ __html: html }} />;
}

import React, { useState, useRef, useCallback, useMemo } from "react";