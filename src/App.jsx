import React, { useState, useRef, useCallback, useMemo } from "react";
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
  code: { label: "Code and Data", icon: Code2, color: "#6EE7B7" },
  other: { label: "Other", icon: FileIcon, color: "#8B95A5" },
};

const TEXT_LIKE = new Set(["txt","md","json","xml","yaml","yml","html","htm","css","js","jsx","ts","tsx","py","java","c","cpp","h","cs","php","sql","sh","csv"]);
const NO_PREVIEW = new Set(["odt","rtf","ppt","pptx","odp","zip","7z","tar","gz","doc"]);