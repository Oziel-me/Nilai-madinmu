import React, { useState, useEffect, useMemo, useRef } from 'react';
import ReactECharts from 'echarts-for-react';
import { motion, AnimatePresence } from 'framer-motion';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import * as htmlToImage from 'html-to-image';
import {
  Upload,
  Layers,
  Users,
  Settings,
  Filter,
  CheckCircle,
  AlertTriangle,
  FileSpreadsheet,
  Download,
  Search,
  RefreshCw,
  Sun,
  Moon,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  Check,
  Menu,
  X,
  Info,
  Calendar,
  Grid,
  TrendingUp,
  TrendingDown
} from 'lucide-react';

// Data Awal Bawaan (Mock Data) sesuai dengan struktur MDT Miftahul Ulum
const DEFAULT_RECORDS = [
  // Kelas I A
  { nopes: '07001', nama: 'Ahmad Fauzi', kelas: 'I A', jenjang: 'I', asr: 'A 01', sihhah: 25, makna_mufrodat: 22, makna_murad: 4, jumlah: 236, rata2: 78.7, rank: 1, kkm: 'TUNTAS', ket: '' },
  { nopes: '07002', nama: 'Budi Harjo', kelas: 'I A', jenjang: 'I', asr: 'A 02', sihhah: 20, makna_mufrodat: 18, makna_murad: 3, jumlah: 188, rata2: 62.7, rank: 4, kkm: 'TUNTAS', ket: '' },
  { nopes: '07003', nama: 'M. Khoirul Anam', kelas: 'I A', jenjang: 'I', asr: 'A 01', sihhah: 15, makna_mufrodat: 14, makna_murad: 2, jumlah: 140, rata2: 46.7, rank: 6, kkm: 'TIDAK TUNTAS', ket: '' },
  { nopes: '07004', nama: 'Siti Aminah', kelas: 'I A', jenjang: 'I', asr: 'B 03', sihhah: 32, makna_mufrodat: 32, makna_murad: 5, jumlah: 33, rata2: 11.0, rank: 7, kkm: 'TIDAK HADIR', ket: '' },
  { nopes: '07005', nama: 'Zainal Abidin', kelas: 'I A', jenjang: 'I', asr: 'A 05', sihhah: 24, makna_mufrodat: 25, makna_murad: 4, jumlah: 244, rata2: 81.3, rank: 2, kkm: 'TUNTAS', ket: '' },
  { nopes: '07006', nama: 'Lukman Hakim', kelas: 'I A', jenjang: 'I', asr: 'A 02', sihhah: 22, makna_mufrodat: 21, makna_murad: 3, jumlah: 208, rata2: 69.3, rank: 5, kkm: 'TUNTAS', ket: '' },
  { nopes: '07007', nama: 'M. Rizqi S.', kelas: 'I A', jenjang: 'I', asr: 'A 03', sihhah: 31, makna_mufrodat: 31, makna_murad: 5, jumlah: 39, rata2: 13.0, rank: 8, kkm: 'TIDAK BACA', ket: '' },
  { nopes: '07008', nama: 'Faisal Riza', kelas: 'I A', jenjang: 'I', asr: 'A 04', sihhah: 23, makna_mufrodat: 24, makna_murad: 4, jumlah: 236, rata2: 78.7, rank: 3, kkm: 'TUNTAS', ket: '' },

  // Kelas I B
  { nopes: '07011', nama: 'Ali Akbar', kelas: 'I B', jenjang: 'I', asr: 'B 01', sihhah: 26, makna_mufrodat: 25, makna_murad: 4, jumlah: 252, rata2: 84.0, rank: 1, kkm: 'TUNTAS', ket: '' },
  { nopes: '07012', nama: 'Fatimah Az-Zahra', kelas: 'I B', jenjang: 'I', asr: 'B 02', sihhah: 18, makna_mufrodat: 16, makna_murad: 3, jumlah: 172, rata2: 57.3, rank: 5, kkm: 'TIDAK TUNTAS', ket: '' },
  { nopes: '07013', nama: 'Hassan Basri', kelas: 'I B', jenjang: 'I', asr: 'B 01', sihhah: 23, makna_mufrodat: 22, makna_murad: 3, jumlah: 216, rata2: 72.0, rank: 3, kkm: 'TUNTAS', ket: '' },
  { nopes: '07014', nama: 'Husain R.', kelas: 'I B', jenjang: 'I', asr: 'B 03', sihhah: 32, makna_mufrodat: 32, makna_murad: 5, jumlah: 33, rata2: 11.0, rank: 6, kkm: 'TIDAK HADIR', ket: '' },
  { nopes: '07015', nama: 'Khadijah', kelas: 'I B', jenjang: 'I', asr: 'B 05', sihhah: 25, makna_mufrodat: 24, makna_murad: 4, jumlah: 244, rata2: 81.3, rank: 2, kkm: 'TUNTAS', ket: '' },
  { nopes: '07016', nama: 'Umar bin Khattab', kelas: 'I B', jenjang: 'I', asr: 'B 02', sihhah: 21, makna_mufrodat: 20, makna_murad: 3, jumlah: 198, rata2: 66.0, rank: 4, kkm: 'TUNTAS', ket: '' },

  // Kelas II A
  { nopes: '08001', nama: 'Abdurrahman', kelas: 'II A', jenjang: 'II', asr: 'A 01', sihhah: 28, makna_mufrodat: 27, makna_murad: 5, jumlah: 280, rata2: 93.3, rank: 1, kkm: 'TUNTAS', ket: '' },
  { nopes: '08002', nama: 'Hamzah Fansuri', kelas: 'II A', jenjang: 'II', asr: 'A 02', sihhah: 20, makna_mufrodat: 20, makna_murad: 3, jumlah: 200, rata2: 66.7, rank: 4, kkm: 'TUNTAS', ket: '' },
  { nopes: '08003', nama: 'Ibrahim', kelas: 'II A', jenjang: 'II', asr: 'A 01', sihhah: 12, makna_mufrodat: 11, makna_murad: 2, jumlah: 116, rata2: 38.7, rank: 6, kkm: 'TIDAK TUNTAS', ket: '' },
  { nopes: '08004', nama: 'Ismail', kelas: 'II A', jenjang: 'II', asr: 'A 03', sihhah: 25, makna_mufrodat: 23, makna_murad: 4, jumlah: 240, rata2: 80.0, rank: 2, kkm: 'TUNTAS', ket: '' },
  { nopes: '08005', nama: 'Yusuf', kelas: 'II A', jenjang: 'II', asr: 'A 05', sihhah: 23, makna_mufrodat: 22, makna_murad: 4, jumlah: 228, rata2: 76.0, rank: 3, kkm: 'TUNTAS', ket: '' },
  { nopes: '08006', nama: 'Yahya', kelas: 'II A', jenjang: 'II', asr: 'A 02', sihhah: 31, makna_mufrodat: 31, makna_murad: 5, jumlah: 39, rata2: 13.0, rank: 7, kkm: 'TIDAK BACA', ket: '' },
  { nopes: '08007', nama: 'Zakaria', kelas: 'II A', jenjang: 'II', asr: 'A 04', sihhah: 19, makna_mufrodat: 18, makna_murad: 3, jumlah: 184, rata2: 61.3, rank: 5, kkm: 'TUNTAS', ket: '' },

  // Kelas II B
  { nopes: '08011', nama: 'M. Ali Ridho', kelas: 'II B', jenjang: 'II', asr: 'B 01', sihhah: 22, makna_mufrodat: 24, makna_murad: 4, jumlah: 232, rata2: 77.3, rank: 2, kkm: 'TUNTAS', ket: '' },
  { nopes: '08012', nama: 'M. Zaini', kelas: 'II B', jenjang: 'II', asr: 'B 02', sihhah: 15, makna_mufrodat: 15, makna_murad: 2, jumlah: 144, rata2: 48.0, rank: 5, kkm: 'TIDAK TUNTAS', ket: '' },
  { nopes: '08013', nama: 'Ahmad Nawawi', kelas: 'II B', jenjang: 'II', asr: 'B 04', sihhah: 25, makna_mufrodat: 26, makna_murad: 4, jumlah: 252, rata2: 84.0, rank: 1, kkm: 'TUNTAS', ket: '' },
  { nopes: '08014', nama: 'Ghufron', kelas: 'II B', jenjang: 'II', asr: 'B 03', sihhah: 21, makna_mufrodat: 20, makna_murad: 3, jumlah: 198, rata2: 66.0, rank: 4, kkm: 'TUNTAS', ket: '' },
  { nopes: '08015', nama: 'M. Hasan', kelas: 'II B', jenjang: 'II', asr: 'B 05', sihhah: 32, makna_mufrodat: 32, makna_murad: 5, jumlah: 33, rata2: 11.0, rank: 6, kkm: 'TIDAK HADIR', ket: '' },
  { nopes: '08016', nama: 'Ridwan', kelas: 'II B', jenjang: 'II', asr: 'B 01', sihhah: 22, makna_mufrodat: 22, makna_murad: 3, jumlah: 212, rata2: 70.7, rank: 3, kkm: 'TUNTAS', ket: '' },

  // Kelas III A
  { nopes: '09001', nama: 'Abdur Rauf', kelas: 'III A', jenjang: 'III', asr: 'A 01', sihhah: 28, makna_mufrodat: 29, makna_murad: 5, jumlah: 288, rata2: 96.0, rank: 1, kkm: 'TUNTAS', ket: '' },
  { nopes: '09002', nama: 'Imron Rosyadi', kelas: 'III A', jenjang: 'III', asr: 'A 02', sihhah: 25, makna_mufrodat: 24, makna_murad: 4, jumlah: 244, rata2: 81.3, rank: 2, kkm: 'TUNTAS', ket: '' },
  { nopes: '09003', nama: 'M. Syarif', kelas: 'III A', jenjang: 'III', asr: 'A 03', sihhah: 23, makna_mufrodat: 21, makna_murad: 4, jumlah: 224, rata2: 74.7, rank: 3, kkm: 'TUNTAS', ket: '' },
  { nopes: '09004', nama: 'M. Mustofa', kelas: 'III A', jenjang: 'III', asr: 'A 04', sihhah: 18, makna_mufrodat: 17, makna_murad: 3, jumlah: 176, rata2: 58.7, rank: 5, kkm: 'TIDAK TUNTAS', ket: '' },
  { nopes: '09005', nama: 'Zulkifli', kelas: 'III A', jenjang: 'III', asr: 'A 01', sihhah: 22, makna_mufrodat: 21, makna_murad: 3, jumlah: 208, rata2: 69.3, rank: 4, kkm: 'TUNTAS', ket: '' }
];

// --- CUSTOM SELECT COMPONENT (Cloudflare style) ---
const CustomSelect = ({ value, onChange, options, placeholder, ariaLabel, className = "w-48" }) => {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (containerRef.current && !containerRef.current.contains(e.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleKeyDown = (e) => {
    if (e.key === 'Escape') setIsOpen(false);
    if (e.key === 'ArrowDown' && isOpen) {
      e.preventDefault();
      const currentIndex = options.findIndex(opt => opt.value === value);
      const nextIndex = (currentIndex + 1) % options.length;
      onChange(options[nextIndex].value);
    }
    if (e.key === 'ArrowUp' && isOpen) {
      e.preventDefault();
      const currentIndex = options.findIndex(opt => opt.value === value);
      const prevIndex = (currentIndex - 1 + options.length) % options.length;
      onChange(options[prevIndex].value);
    }
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      setIsOpen(!isOpen);
    }
  };

  const selectedOption = options.find(opt => opt.value === value) || { label: placeholder || '', value: '' };

  return (
    <div className={`relative ${className}`} ref={containerRef} onKeyDown={handleKeyDown}>
      <button
        type="button"
        className="w-full bg-bg-card border border-border-color text-text-main text-xs font-semibold rounded-md px-3.5 py-2 flex items-center justify-between outline-none cursor-pointer hover:border-primary transition-all duration-150"
        onClick={() => setIsOpen(!isOpen)}
        aria-expanded={isOpen}
        aria-haspopup="listbox"
        aria-label={ariaLabel}
      >
        <span>{selectedOption.label}</span>
        <ChevronDown size={14} className={`text-text-muted transition-transform duration-150 ${isOpen ? 'rotate-180' : ''}`} />
      </button>
      {isOpen && (
        <ul
          className="absolute z-50 w-full mt-1 bg-bg-card border border-border-color rounded-md shadow-lg py-1 max-h-60 overflow-y-auto outline-none transition-all duration-150"
          role="listbox"
        >
          {options.map((opt) => (
            <li
              key={opt.value}
              className={`px-3 py-1.5 text-xs cursor-pointer flex items-center justify-between hover:bg-primary-soft hover:text-primary transition-colors ${
                opt.value === value ? 'bg-primary-soft text-primary font-semibold' : 'text-text-main'
              }`}
              role="option"
              aria-selected={opt.value === value}
              onClick={() => {
                onChange(opt.value);
                setIsOpen(false);
              }}
            >
              <span>{opt.label}</span>
              {opt.value === value && <Check size={12} className="text-primary" />}
            </li>
          ))}
        </ul>
      )}
    </div>
  );
};

// --- HELPER TO CONVERT BASE64 DATAURL TO BINARY BLOB ---
const dataUrlToBlob = (dataurl) => {
  try {
    const arr = dataurl.split(',');
    const mime = arr[0].match(/:(.*?);/)[1];
    const bstr = atob(arr[1]);
    let n = bstr.length;
    const u8arr = new Uint8Array(n);
    while (n--) {
      u8arr[n] = bstr.charCodeAt(n);
    }
    return new Blob([u8arr], { type: mime });
  } catch (e) {
    throw new Error("Gagal melakukan decoding base64 gambar.");
  }
};

// --- CHART EXPORT MENU ---
const ChartExportMenu = ({ chartRef, filename }) => {
  const [isOpen, setIsOpen] = useState(false);
  const containerRef = useRef(null);

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (containerRef.current && !containerRef.current.contains(e.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleExport = async (type) => {
    setIsOpen(false);
    if (!chartRef.current) return;

    try {
      const echartsInstance = chartRef.current.getEchartsInstance();
      const dom = echartsInstance.getDom();

      if (type === 'svg') {
        const svgEl = dom.querySelector('svg');
        if (!svgEl) {
          throw new Error("Elemen SVG grafik tidak ditemukan.");
        }
        const serializer = new XMLSerializer();
        let source = serializer.serializeToString(svgEl);
        
        // Pastikan namespace xmlns disertakan
        if (!source.includes('xmlns="http://www.w3.org/2000/svg"')) {
          source = source.replace(/^<svg/, '<svg xmlns="http://www.w3.org/2000/svg"');
        }
        // Pastikan xmlns:xlink disertakan
        if (!source.includes('xmlns:xlink="http://www.w3.org/1999/xlink"')) {
          source = source.replace(/^<svg/, '<svg xmlns:xlink="http://www.w3.org/1999/xlink"');
        }

        const xmlString = `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n${source}`;
        const svgBlob = new Blob([xmlString], { type: 'image/svg+xml;charset=utf-8' });
        
        if (svgBlob.size === 0) {
          throw new Error("File SVG yang dihasilkan kosong (0 byte).");
        }
        saveAs(svgBlob, `${filename}.svg`);
      } else if (type === 'png') {
        const dataUrl = await htmlToImage.toPng(dom, {
          backgroundColor: 'transparent',
          style: {
            background: 'transparent'
          },
          pixelRatio: 2
        });
        if (!dataUrl || dataUrl === 'data:,') {
          throw new Error("Gagal merender gambar PNG.");
        }
        const blob = dataUrlToBlob(dataUrl);
        if (blob.size === 0) {
          throw new Error("File PNG yang dihasilkan kosong (0 byte).");
        }
        saveAs(blob, `${filename}.png`);
      } else if (type === 'jpg') {
        const isDark = document.documentElement.classList.contains('dark');
        const bgColor = isDark ? '#18191b' : '#ffffff';
        
        const dataUrl = await htmlToImage.toJpeg(dom, {
          backgroundColor: bgColor,
          style: {
            background: bgColor
          },
          pixelRatio: 2,
          quality: 0.95
        });
        if (!dataUrl || dataUrl === 'data:,') {
          throw new Error("Gagal merender gambar JPG.");
        }
        const blob = dataUrlToBlob(dataUrl);
        if (blob.size === 0) {
          throw new Error("File JPG yang dihasilkan kosong (0 byte).");
        }
        saveAs(blob, `${filename}.jpg`);
      }
    } catch (error) {
      console.error(error);
      alert(`Gagal mengekspor grafik: ${error.message}`);
    }
  };

  return (
    <div className="relative" ref={containerRef}>
      <button
        onClick={() => setIsOpen(!isOpen)}
        className="p-1.5 rounded-md hover:bg-bg-app border border-border-color text-text-muted hover:text-text-main transition-colors duration-150 cursor-pointer flex items-center justify-center"
        title="Ekspor Grafik"
      >
        <Download size={14} />
      </button>
      {isOpen && (
        <ul className="absolute right-0 mt-1 z-50 w-28 bg-bg-card border border-border-color rounded-md shadow-lg py-1 text-xs select-none">
          <li
            onClick={() => handleExport('png')}
            className="px-3 py-2 text-text-main hover:bg-primary-soft hover:text-primary cursor-pointer transition-colors"
          >
            Ekspor PNG
          </li>
          <li
            onClick={() => handleExport('jpg')}
            className="px-3 py-2 text-text-main hover:bg-primary-soft hover:text-primary cursor-pointer transition-colors"
          >
            Ekspor JPG
          </li>
          <li
            onClick={() => handleExport('svg')}
            className="px-3 py-2 text-text-main hover:bg-primary-soft hover:text-primary cursor-pointer transition-colors"
          >
            Ekspor SVG
          </li>
        </ul>
      )}
    </div>
  );
};

function App() {
  // --- STATE MANAGEMENT ---
  const [records, setRecords] = useState(DEFAULT_RECORDS);
  const [filterMode, setFilterMode] = useState('kelas'); // kelas | jenjang
  const [selectedKelas, setSelectedKelas] = useState('I A');
  const [selectedJenjang, setSelectedJenjang] = useState('I');
  const [pctBasis, setPctBasis] = useState('total'); // total | dinilai
  const [darkMode, setDarkMode] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [studentStatusFilter, setStudentStatusFilter] = useState('Semua');
  const [fileName, setFileName] = useState('');
  const [dragActive, setDragActive] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);
  const [menuOpen, setMenuOpen] = useState(false);
  const fileInputRef = useRef(null);
  const donutChartRef = useRef(null);
  const barChartRef = useRef(null);

  const itemsPerPage = 8;

  // Toggle Dark Mode
  useEffect(() => {
    const root = window.document.documentElement;
    if (darkMode) {
      root.classList.add('dark');
    } else {
      root.classList.remove('dark');
    }
  }, [darkMode]);

  // Extract Lists
  const kelasList = useMemo(() => {
    return [...new Set(records.map(r => r.kelas))].sort();
  }, [records]);

  const jenjangList = useMemo(() => {
    return [...new Set(records.map(r => r.jenjang))].sort();
  }, [records]);

  // Sync selection when records load
  useEffect(() => {
    if (kelasList.length > 0 && !kelasList.includes(selectedKelas)) {
      setSelectedKelas(kelasList[0]);
    }
    if (jenjangList.length > 0 && !jenjangList.includes(selectedJenjang)) {
      setSelectedJenjang(jenjangList[0]);
    }
  }, [records, kelasList, jenjangList]);

  // Filter records based on selected view (Kelas / Jenjang)
  const currentRecords = useMemo(() => {
    if (filterMode === 'kelas') {
      return records.filter(r => r.kelas === selectedKelas);
    } else {
      return records.filter(r => r.jenjang === selectedJenjang);
    }
  }, [records, filterMode, selectedKelas, selectedJenjang]);

  // --- LOGIC CALCULATOR (Core Business Rules) ---
  const stats = useMemo(() => {
    const total = currentRecords.length;

    const dinilai = currentRecords.filter(
      r => r.kkm !== 'TIDAK HADIR' && r.kkm !== 'TIDAK BACA'
    );
    const tidakHadirOrBaca = currentRecords.filter(
      r => r.kkm === 'TIDAK HADIR' || r.kkm === 'TIDAK BACA'
    );

    const tuntas = dinilai.filter(r => r.kkm === 'TUNTAS');
    const tidakTuntas = dinilai.filter(r => r.kkm === 'TIDAK TUNTAS');

    const tuntasAtas70 = tuntas.filter(r => r.rata2 > 70);
    const tuntasBawah70 = tuntas.filter(r => r.rata2 <= 70);

    // Basis percentage selector
    const denom = pctBasis === 'total' ? total : dinilai.length;
    const getPct = (n) => (denom === 0 ? 0 : +(n / denom * 100).toFixed(1));

    return {
      total,
      dinilaiCount: dinilai.length,
      tidakHadirBacaCount: tidakHadirOrBaca.length,
      tuntasCount: tuntas.length,
      tidakTuntasCount: tidakTuntas.length,
      tuntasAtas70Count: tuntasAtas70.length,
      tuntasBawah70Count: tuntasBawah70.length,

      tuntasPct: getPct(tuntas.length),
      tidakTuntasPct: getPct(tidakTuntas.length),
      tuntasAtas70Pct: getPct(tuntasAtas70.length),
      tuntasBawah70Pct: getPct(tuntasBawah70.length),
      tidakHadirBacaPct: getPct(tidakHadirOrBaca.length)
    };
  }, [currentRecords, pctBasis]);

  // --- SHEETJS EXCEL PARSER ---
  const parseExcelFile = (file) => {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const allRecords = [];

        wb.SheetNames.forEach(sheetName => {
          const sheet = wb.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
          if (rows.length === 0) return;

          // Find header row (which contains NOPES and NAMA)
          const headerIdx = rows.findIndex(r => 
            r && r.some(c => typeof c === 'string' && c.toUpperCase().trim() === 'NOPES')
          );
          if (headerIdx === -1) return;

          // Find class name in row or fallback to tab name
          let kelas = sheetName;
          const kelasRow = rows.find(r => 
            r && r.some(c => typeof c === 'string' && c.toUpperCase().includes('KELAS'))
          );
          if (kelasRow) {
            const cell = kelasRow.find(c => typeof c === 'string' && c.toUpperCase().includes('KELAS'));
            if (cell.includes(':')) {
              kelas = cell.split(':')[1].trim();
            } else {
              kelas = cell.replace(/kelas/gi, '').trim();
            }
          }

          const headerRow = rows[headerIdx];
          const colMap = {};
          headerRow.forEach((c, idx) => {
            if (!c) return;
            const cleanHeader = String(c).toUpperCase().trim();
            if (cleanHeader.includes('NOPES')) colMap.nopes = idx;
            else if (cleanHeader.includes('NAMA')) colMap.nama = idx;
            else if (cleanHeader.includes('ASR')) colMap.asr = idx;
            else if (cleanHeader.includes('SIHHAH') || cleanHeader.includes('QIRO')) colMap.sihhah = idx;
            else if (cleanHeader.includes('MUFRODAT') || cleanHeader.includes('MAKNA MUF')) colMap.mufrodat = idx;
            else if (cleanHeader.includes('MURAD') || cleanHeader.includes('MAKNA MUR')) colMap.murad = idx;
            else if (cleanHeader.includes('JML') || cleanHeader.includes('JUMLAH')) colMap.jml = idx;
            else if (cleanHeader.includes('RT') || cleanHeader.includes('RATA2') || cleanHeader.includes('RATA-RATA')) colMap.rata2 = idx;
            else if (cleanHeader.includes('RANK')) colMap.rank = idx;
            else if (cleanHeader.includes('KKM')) colMap.kkm = idx;
            else if (cleanHeader.includes('KET')) colMap.ket = idx;
          });

          // Read data rows
          for (let i = headerIdx + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            
            const nopesVal = String(row[colMap.nopes] || '').trim();
            // Skip rows that don't have valid NOPES numbers or titles
            if (!nopesVal || nopesVal === '' || nopesVal.toUpperCase().includes('NO')) continue;

            const kkmVal = String(row[colMap.kkm] || '').trim().toUpperCase();

            allRecords.push({
              nopes: nopesVal,
              nama: String(row[colMap.nama] || '').trim(),
              kelas: kelas,
              jenjang: kelas.trim().split(' ')[0] || 'Lainnya',
              asr: String(row[colMap.asr] || '').trim(),
              sihhah: Number(row[colMap.sihhah]) || 0,
              makna_mufrodat: Number(row[colMap.mufrodat]) || 0,
              makna_murad: Number(row[colMap.murad]) || 0,
              jumlah: Number(row[colMap.jml]) || 0,
              rata2: Number(row[colMap.rata2]) || 0,
              rank: Number(row[colMap.rank]) || 0,
              kkm: kkmVal === 'TUNTAS' ? 'TUNTAS' :
                   kkmVal === 'TIDAK TUNTAS' ? 'TIDAK TUNTAS' :
                   kkmVal === 'TIDAK HADIR' ? 'TIDAK HADIR' :
                   kkmVal === 'TIDAK BACA' ? 'TIDAK BACA' : kkmVal,
              ket: String(row[colMap.ket] || '').trim()
            });
          }
        });

        if (allRecords.length > 0) {
          setRecords(allRecords);
          setCurrentPage(1);
        } else {
          alert('Format tabel tidak sesuai! Pastikan terdapat kolom NOPES, NAMA, dan KKM.');
        }
      } catch (err) {
        console.error(err);
        alert('Terjadi kesalahan saat membaca file Excel.');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      parseExcelFile(e.dataTransfer.files[0]);
    }
  };

  // --- DOWNLOAD SAMPLE EXCEL ---
  const downloadSampleExcel = () => {
    const link = document.createElement('a');
    link.href = '/TEEMPLATE.xlsx';
    link.download = 'template-penilaian-ujian-baca-kitab.xlsx';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleResetData = () => {
    setRecords(DEFAULT_RECORDS);
    setFileName('');
    setCurrentPage(1);
  };

  // --- DYNAMIC THEMING FOR ECHARTS ---
  const themeColors = useMemo(() => {
    const isDark = darkMode;
    return {
      text: isDark ? '#f4f4f5' : '#18181b',
      textMuted: isDark ? '#a1a1aa' : '#71717a',
      border: isDark ? '#27292c' : '#e5e7eb',
      bgCard: isDark ? '#18191b' : '#ffffff',
      bgTooltip: isDark ? '#18191b' : '#ffffff',
      primary: isDark ? '#22c55e' : '#16a34a',
      danger: isDark ? '#f87171' : '#dc2626',
      neutral: isDark ? '#52525b' : '#a1a1aa',
      // Emerald and slate for aspect charts
      emerald: isDark ? '#34d399' : '#059669',
      slate: isDark ? '#94a3b8' : '#475569'
    };
  }, [darkMode]);

  // Donut Chart - proporsi kelulusan siswa
  const donutChartOptions = useMemo(() => {
    return {
      title: {
        text: stats.total,
        subtext: 'Total Santri',
        left: 'center',
        top: '38%',
        textStyle: {
          fontSize: 24,
          fontWeight: 'semibold',
          fontFamily: 'Inter, sans-serif',
          color: themeColors.text
        },
        subtextStyle: {
          fontSize: 10,
          fontFamily: 'Inter, sans-serif',
          color: themeColors.textMuted
        }
      },
      tooltip: {
        trigger: 'item',
        backgroundColor: themeColors.bgTooltip,
        borderColor: themeColors.border,
        borderWidth: 1,
        borderRadius: 8,
        shadowColor: 'rgba(0, 0, 0, 0.05)',
        shadowBlur: 10,
        textStyle: {
          color: themeColors.text,
          fontFamily: 'Inter, sans-serif',
          fontSize: 11
        }
      },
      legend: {
        show: false
      },
      series: [
        {
          name: 'Status Ujian',
          type: 'pie',
          radius: ['60%', '75%'],
          center: ['50%', '50%'],
          avoidLabelOverlap: false,
          roundCap: true,
          itemStyle: {
            borderColor: themeColors.bgCard,
            borderWidth: 2
          },
          label: { show: false },
          labelLine: { show: false },
          data: [
            { value: stats.tuntasCount, name: 'Tuntas', itemStyle: { color: themeColors.primary } },
            { value: stats.tidakTuntasCount, name: 'Tidak Tuntas', itemStyle: { color: themeColors.danger } },
            { value: stats.tidakHadirBacaCount, name: 'Tidak Hadir/Baca', itemStyle: { color: themeColors.neutral } }
          ],
          animationDuration: 350,
          animationEasing: 'cubicOut'
        }
      ]
    };
  }, [stats, themeColors, darkMode]);

  // Bar Chart - Perbandingan kelas dalam satu jenjang / rata-rata nilai aspek kelas
  const barChartOptions = useMemo(() => {
    if (filterMode === 'jenjang') {
      // Dapatkan daftar kelas pada jenjang terpilih
      const classesInJenjang = [...new Set(records.filter(r => r.jenjang === selectedJenjang).map(r => r.kelas))].sort();
      
      const seriesData = classesInJenjang.map(k => {
        const clsRecords = records.filter(r => r.kelas === k);
        const total = clsRecords.length;
        const dinilai = clsRecords.filter(r => r.kkm !== 'TIDAK HADIR' && r.kkm !== 'TIDAK BACA');
        const tuntas = dinilai.filter(r => r.kkm === 'TUNTAS');
        const denom = pctBasis === 'total' ? total : dinilai.length;
        const pct = denom === 0 ? 0 : +(tuntas.length / denom * 100).toFixed(1);
        return { name: k, value: pct };
      }).sort((a, b) => a.value - b.value); 

      return {
        tooltip: {
          trigger: 'axis',
          axisPointer: { type: 'shadow' },
          backgroundColor: themeColors.bgTooltip,
          borderColor: themeColors.border,
          borderWidth: 1,
          borderRadius: 8,
          shadowColor: 'rgba(0, 0, 0, 0.05)',
          shadowBlur: 10,
          textStyle: {
            color: themeColors.text,
            fontFamily: 'Inter, sans-serif',
            fontSize: 11
          }
        },
        grid: { left: '3%', right: '8%', bottom: '5%', top: '5%', containLabel: true },
        xAxis: {
          type: 'value',
          name: '% Tuntas',
          nameTextStyle: { color: themeColors.textMuted },
          axisLabel: { color: themeColors.textMuted },
          splitLine: { show: true, lineStyle: { type: 'dashed', color: themeColors.border } },
          max: 100
        },
        yAxis: {
          type: 'category',
          data: seriesData.map(d => `Kelas ${d.name}`),
          axisLabel: { color: themeColors.text, fontSize: 11 },
          axisLine: { lineStyle: { color: themeColors.border } }
        },
        series: [
          {
            name: 'Persentase Tuntas',
            type: 'bar',
            data: seriesData.map(d => d.value),
            itemStyle: {
              color: themeColors.primary,
              borderRadius: [0, 4, 4, 0]
            },
            animationDuration: 350,
            animationEasing: 'cubicOut'
          }
        ]
      };
    } else {
      // Mode kelas: Hitung rata-rata nilai sub-aspek (Sihhah Al-Qiro'ah [max 30], Mufrodat [max 30], Murad [max 5])
      const dinilai = currentRecords.filter(r => r.kkm !== 'TIDAK HADIR' && r.kkm !== 'TIDAK BACA');
      const avgSihhah = dinilai.length === 0 ? 0 : +(dinilai.reduce((s, r) => s + r.sihhah, 0) / dinilai.length).toFixed(1);
      const avgMufrodat = dinilai.length === 0 ? 0 : +(dinilai.reduce((s, r) => s + r.makna_mufrodat, 0) / dinilai.length).toFixed(1);
      const avgMurad = dinilai.length === 0 ? 0 : +(dinilai.reduce((s, r) => s + r.makna_murad, 0) / dinilai.length).toFixed(1);

      // Konversi persentase capaian dari maksimal, diurutkan ascending untuk horizontal bar
      const pctSihhah = +(avgSihhah / 30 * 100).toFixed(1);
      const pctMufrodat = +(avgMufrodat / 30 * 100).toFixed(1);
      const pctMurad = +(avgMurad / 5 * 100).toFixed(1);

      const items = [
        { name: 'Makna Murad (%)', value: pctMurad },
        { name: 'Makna Mufrodat (%)', value: pctMufrodat },
        { name: 'Sihhah Al-Qiro\'ah (%)', value: pctSihhah }
      ];

      return {
        tooltip: {
          trigger: 'axis',
          axisPointer: { type: 'shadow' },
          backgroundColor: themeColors.bgTooltip,
          borderColor: themeColors.border,
          borderWidth: 1,
          borderRadius: 8,
          shadowColor: 'rgba(0, 0, 0, 0.05)',
          shadowBlur: 10,
          textStyle: {
            color: themeColors.text,
            fontFamily: 'Inter, sans-serif',
            fontSize: 11
          }
        },
        grid: { left: '3%', right: '8%', bottom: '5%', top: '5%', containLabel: true },
        xAxis: {
          type: 'value',
          axisLabel: { color: themeColors.textMuted },
          splitLine: { show: true, lineStyle: { type: 'dashed', color: themeColors.border } },
          max: 100
        },
        yAxis: {
          type: 'category',
          data: items.map(i => i.name),
          axisLabel: { color: themeColors.text, fontSize: 11 },
          axisLine: { lineStyle: { color: themeColors.border } }
        },
        series: [
          {
            name: 'Rata-rata Aspek',
            type: 'bar',
            data: items.map(i => i.value),
            itemStyle: {
              color: function(params) {
                // Return matching design token colors
                const colorsMap = [
                  themeColors.slate,   // Murad
                  themeColors.emerald, // Mufrodat
                  themeColors.primary  // Sihhah
                ];
                return colorsMap[params.dataIndex];
              },
              borderRadius: [0, 4, 4, 0]
            },
            animationDuration: 350,
            animationEasing: 'cubicOut'
          }
        ]
      };
    }
  }, [filterMode, records, selectedJenjang, pctBasis, currentRecords, themeColors, darkMode]);

  // --- REKAP TABEL KELAS / JENJANG (Comparison Table) ---
  const comparisonData = useMemo(() => {
    const targetList = filterMode === 'kelas' ? kelasList : jenjangList;
    return targetList.map(item => {
      const filtered = records.filter(r => (filterMode === 'kelas' ? r.kelas === item : r.jenjang === item));
      const total = filtered.length;
      const dinilai = filtered.filter(r => r.kkm !== 'TIDAK HADIR' && r.kkm !== 'TIDAK BACA');
      const tuntas = dinilai.filter(r => r.kkm === 'TUNTAS');
      const tidakTuntas = dinilai.filter(r => r.kkm === 'TIDAK TUNTAS');
      const tidakHadirOrBaca = filtered.filter(r => r.kkm === 'TIDAK HADIR' || r.kkm === 'TIDAK BACA');

      const denom = pctBasis === 'total' ? total : dinilai.length;
      const tuntasPct = denom === 0 ? 0 : +(tuntas.length / denom * 100).toFixed(1);
      const tidakTuntasPct = denom === 0 ? 0 : +(tidakTuntas.length / denom * 100).toFixed(1);

      return {
        name: item,
        total,
        dinilai: dinilai.length,
        tidakHadir: tidakHadirOrBaca.length,
        tuntas: tuntas.length,
        tidakTuntas: tidakTuntas.length,
        tuntasPct,
        tidakTuntasPct
      };
    });
  }, [records, filterMode, kelasList, jenjangList, pctBasis]);

  // --- DETAILED STUDENTS TABLE FILTERS & PAGINATION ---
  const filteredStudents = useMemo(() => {
    return currentRecords.filter(student => {
      const matchesSearch = 
        student.nama.toLowerCase().includes(searchQuery.toLowerCase()) ||
        student.nopes.toLowerCase().includes(searchQuery.toLowerCase()) ||
        student.asr.toLowerCase().includes(searchQuery.toLowerCase());
      
      const matchesStatus = studentStatusFilter === 'Semua' || student.kkm === studentStatusFilter;

      return matchesSearch && matchesStatus;
    });
  }, [currentRecords, searchQuery, studentStatusFilter]);

  const totalPages = Math.ceil(filteredStudents.length / itemsPerPage);
  
  const paginatedStudents = useMemo(() => {
    const start = (currentPage - 1) * itemsPerPage;
    return filteredStudents.slice(start, start + itemsPerPage);
  }, [filteredStudents, currentPage]);

  useEffect(() => {
    setCurrentPage(1);
  }, [searchQuery, studentStatusFilter, selectedKelas, selectedJenjang, filterMode]);

  return (
    <div className="min-h-screen transition-colors duration-300">
      
      {/* HEADER BANNER */}
      <header className="bg-bg-card border-b border-border-color text-text-main sticky top-0 z-50 transition-colors duration-150 h-14">
        <div className="w-full px-4 sm:px-6 h-full flex items-center justify-between">
          
          {/* Kiri: Logo + Nama Aplikasi */}
          <div className="flex items-center gap-3">
            <div className="w-6 h-6 rounded-[6px] bg-primary flex items-center justify-center font-bold text-[10px] text-white select-none transition-colors duration-150">
              MU
            </div>
            <div className="flex items-baseline gap-2">
              <span className="hidden sm:inline text-[15px] font-semibold text-text-main">
                Infografis Penilaian Ujian Baca Kitab
              </span>
              <span className="inline sm:hidden text-[15px] font-semibold text-text-main">
                Ujian Baca Kitab
              </span>
            </div>
            
            {/* Tengah: Breadcrumb/Konteks Kelas Aktif (Desktop Only) */}
            <div className="hidden lg:flex items-center gap-2 text-xs text-text-muted select-none">
              <span className="text-text-muted/40 font-light">/</span>
              <span className="font-medium bg-bg-app px-2 py-0.5 rounded border border-border-color">
                {filterMode === 'kelas' ? `Kelas ${selectedKelas}` : `Jenjang ${selectedJenjang}`}
              </span>
            </div>
          </div>

          {/* Kanan: Kontrol rapat */}
          <div className="flex items-center gap-3">
            
            {/* Kontrol Desktop & Tablet */}
            <div className="hidden sm:flex items-center gap-2.5">

              {/* Reset Data Button */}
              <button
                onClick={handleResetData}
                className="p-1.5 rounded-md border border-border-color bg-bg-card text-text-muted hover:text-text-main hover:bg-bg-app transition-all duration-150 cursor-pointer"
                title="Reset Data ke Awal"
              >
                <RefreshCw size={16} />
              </button>

              {/* Dark Mode switcher */}
              <button 
                onClick={() => setDarkMode(!darkMode)}
                className="p-1.5 rounded-md border border-border-color bg-bg-card text-text-muted hover:text-text-main hover:bg-bg-app transition-all duration-150 cursor-pointer"
                title="Ganti Tema Visual"
              >
                {darkMode ? <Sun size={16} /> : <Moon size={16} />}
              </button>

              {/* Avatar / Inisial */}
              <div 
                className="w-6 h-6 rounded-full bg-primary/10 text-primary border border-primary/20 flex items-center justify-center font-semibold text-[10px] select-none" 
                title="Administrator MDTMU"
              >
                AD
              </div>

            </div>

            {/* Kontrol Mobile (Hamburger Menu Icon) */}
            <div className="flex sm:hidden">
              <button 
                onClick={() => setMenuOpen(!menuOpen)}
                className="p-1.5 rounded-md hover:bg-bg-app text-text-muted hover:text-text-main border border-border-color transition-colors duration-150 cursor-pointer"
              >
                {menuOpen ? <X size={16} /> : <Menu size={16} />}
              </button>
            </div>

          </div>

        </div>
      </header>

      {/* Mobile Dropdown Panel */}
      {menuOpen && (
        <div className="sm:hidden bg-bg-card border-b border-border-color px-4 py-4 flex flex-col gap-4 transition-colors duration-150">
          <div className="flex items-center justify-between">
            <span className="text-xs text-text-muted font-medium">Konteks Aktif:</span>
            <span className="text-xs font-semibold bg-bg-app border border-border-color px-2.5 py-1 rounded">
              {filterMode === 'kelas' ? `Kelas ${selectedKelas}` : `Jenjang ${selectedJenjang}`}
            </span>
          </div>
          <div className="flex items-center justify-between">
            <span className="text-xs text-text-muted font-medium">Tema Visual:</span>
            <button 
              onClick={() => {
                setDarkMode(!darkMode);
                setMenuOpen(false);
              }}
              className="flex items-center gap-2 px-3 py-1.5 rounded-md border border-border-color bg-bg-app text-text-muted hover:text-text-main text-xs font-medium cursor-pointer"
            >
              {darkMode ? (
                <>
                  <Sun size={14} />
                  <span>Mode Terang</span>
                </>
              ) : (
                <>
                  <Moon size={14} />
                  <span>Mode Gelap</span>
                </>
              )}
            </button>
          </div>
          <div className="flex items-center justify-between">
            <span className="text-xs text-text-muted font-medium">Data Aplikasi:</span>
            <button 
              onClick={() => {
                handleResetData();
                setMenuOpen(false);
              }}
              className="flex items-center gap-2 px-3 py-1.5 rounded-md border border-border-color bg-bg-app text-text-muted hover:text-text-main text-xs font-medium cursor-pointer"
            >
              <RefreshCw size={14} />
              <span>Reset Data</span>
            </button>
          </div>
          <div className="flex items-center justify-between pt-2 border-t border-border-color">
            <span className="text-xs text-text-muted font-medium">Pengguna:</span>
            <div className="flex items-center gap-2">
              <div className="w-6 h-6 rounded-full bg-primary/10 text-primary border border-primary/20 flex items-center justify-center font-bold text-[10px] select-none">
                AD
              </div>
              <span className="text-xs font-medium text-text-main">Administrator</span>
            </div>
          </div>
        </div>
      )}

      {/* DASHBOARD CONTAINER */}
      <main className="w-full px-4 sm:px-6 py-8 flex flex-col gap-8">
        
        {/* ROW 1: UPLOAD EXCEL SECTION */}
        <section className="bg-bg-card border border-border-color rounded-2xl p-6 shadow-sm flex flex-col md:flex-row gap-6 items-center transition-all duration-200">
          
          <div className="flex-1">
            <h2 className="text-base font-bold text-text-main flex items-center gap-2">
              <Upload className="text-primary" size={20} />
              <span>Unggah Data Nilai Ujian</span>
            </h2>
            <p className="text-xs text-text-muted mt-1 max-w-xl">
              Unggah file Excel hasil penilaian ujian santri. Aplikasi akan mem-parsing otomatis seluruh kelas yang terdapat pada setiap tab sheet secara langsung.
            </p>
            <div className="flex items-center gap-3 mt-4 flex-wrap">
              <button
                className="px-4.5 py-2.5 bg-primary hover:bg-primary-hover text-white rounded-xl text-xs font-semibold shadow-md shadow-primary/20 transition-all duration-200 flex items-center gap-2 cursor-pointer"
                onClick={downloadSampleExcel}
              >
                <Download size={14} />
                <span>Unduh Template Excel</span>
              </button>
              {fileName && (
                <div className="px-3.5 py-2 bg-primary/10 text-primary border border-primary/20 rounded-lg text-xs font-medium flex items-center gap-1.5 animate-pulse">
                  <CheckCircle size={14} />
                  <span>Aktif: {fileName}</span>
                </div>
              )}
            </div>
          </div>

          <div 
            className={`w-full md:w-80 h-32 border-2 border-dashed rounded-xl flex flex-col items-center justify-center p-4 transition-all duration-200 text-center cursor-pointer ${
              dragActive 
                ? 'border-primary bg-primary/10' 
                : 'border-border-color hover:border-primary bg-bg-app'
            }`}
            onDragEnter={handleDrag}
            onDragLeave={handleDrag}
            onDragOver={handleDrag}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
          >
            <input 
              type="file" 
              ref={fileInputRef} 
              className="hidden" 
              accept=".xlsx, .xls"
              onChange={(e) => e.target.files?.[0] && parseExcelFile(e.target.files[0])}
            />
            <FileSpreadsheet className="text-text-muted mb-2" size={32} />
            <span className="text-xs font-semibold text-text-main">Tarik berkas atau klik di sini</span>
            <span className="text-[10px] text-text-muted mt-1">Hanya mendukung format file .xlsx / .xls</span>
          </div>

        </section>

        {/* ROW 2: FILTERS & CONTROL PANEL */}
        <section className="bg-bg-card border border-border-color rounded-2xl p-6 shadow-sm flex flex-col gap-6 transition-all duration-200">
          
          <div className="flex flex-col lg:flex-row justify-between lg:items-center gap-6">
            
            {/* Filter Mode Selector */}
            <div className="flex flex-col gap-2">
              <span className="text-xs font-bold text-text-muted uppercase tracking-wider">Mode Analisis</span>
              <div className="inline-flex p-1 bg-bg-app rounded-xl gap-1 w-fit border border-border-color">
                <button
                  className={`px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                    filterMode === 'kelas'
                      ? 'bg-bg-card text-primary shadow-sm'
                      : 'text-text-muted hover:text-text-main'
                  }`}
                  onClick={() => setFilterMode('kelas')}
                >
                  Per Kelas
                </button>
                <button
                  className={`px-4 py-2 rounded-lg text-xs font-bold transition-all cursor-pointer ${
                    filterMode === 'jenjang'
                      ? 'bg-bg-card text-primary shadow-sm'
                      : 'text-text-muted hover:text-text-main'
                  }`}
                  onClick={() => setFilterMode('jenjang')}
                >
                  Per Jenjang (Agregasi)
                </button>
              </div>
            </div>

            {/* Dropdown Selector */}
            <div className="flex flex-col gap-2">
              <span className="text-xs font-bold text-text-muted uppercase tracking-wider">Pilih Kategori</span>
              <div className="flex items-center gap-2">
                <Filter className="text-text-muted" size={16} />
                {filterMode === 'kelas' ? (
                  <CustomSelect
                    value={selectedKelas}
                    onChange={setSelectedKelas}
                    options={kelasList.map(k => ({ label: `Kelas ${k}`, value: k }))}
                    ariaLabel="Pilih Kategori Kelas"
                  />
                ) : (
                  <CustomSelect
                    value={selectedJenjang}
                    onChange={setSelectedJenjang}
                    options={jenjangList.map(j => ({ label: `Jenjang ${j}`, value: j }))}
                    ariaLabel="Pilih Kategori Jenjang"
                  />
                )}
              </div>
            </div>

            {/* Percentage Basis Toggle */}
            <div className="flex flex-col gap-2">
              <span className="text-xs font-bold text-text-muted uppercase tracking-wider">Basis Perhitungan Persentase</span>
              <div className="flex items-center gap-4 bg-bg-app border border-border-color rounded-xl px-4 py-2.5 w-fit">
                <label className="flex items-center gap-2 text-xs font-bold text-text-main cursor-pointer">
                  <input
                    type="radio"
                    name="pctBasis"
                    value="total"
                    checked={pctBasis === 'total'}
                    onChange={() => setPctBasis('total')}
                    className="text-primary focus:ring-primary h-4 w-4 border-border-color"
                  />
                  <span>Semua Siswa</span>
                </label>
                <label className="flex items-center gap-2 text-xs font-bold text-text-main cursor-pointer">
                  <input
                    type="radio"
                    name="pctBasis"
                    value="dinilai"
                    checked={pctBasis === 'dinilai'}
                    onChange={() => setPctBasis('dinilai')}
                    className="text-primary focus:ring-primary h-4 w-4 border-border-color"
                  />
                  <span>Hanya Yang Dinilai</span>
                </label>
              </div>
            </div>

          </div>

          {/* Business rules notes banner */}
          <div className="bg-bg-app border border-border-color rounded-xl p-4 flex gap-3 text-xs text-text-muted transition-all duration-200">
            <Info className="text-primary shrink-0 mt-0.5" size={16} />
            <div className="flex-1 leading-relaxed">
              <strong>Aturan Bisnis Penilaian:</strong> Santri dengan KKM status <code className="px-1.5 py-0.5 rounded bg-bg-card border border-border-color text-primary font-mono font-bold">TIDAK HADIR</code> (Rata-rata 11.0) atau <code className="px-1.5 py-0.5 rounded bg-bg-card border border-border-color text-primary font-mono font-bold">TIDAK BACA</code> (Rata-rata 13.0) <strong>otomatis dikecualikan</strong> dari penghitungan Tuntas / Tidak Tuntas. Persentase di atas dihitung dengan basis pembagi aktif yang dipilih di sebelah kanan.
            </div>
          </div>

        </section>

        {/* ROW 3: STAT CARDS GRID (Optimized for TVs) */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-5">
          
          {/* Card 1: Keseluruhan */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 hover:border-primary transition-all duration-150">
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-brand-neutral"></span>
                <span className="text-[11px] font-medium text-text-muted uppercase tracking-wide">Keseluruhan</span>
              </div>
              <div className="p-2 bg-bg-app text-text-muted rounded-md border border-border-color">
                <Users size={16} />
              </div>
            </div>
            <div className="flex items-baseline gap-2">
              <span className="text-[32px] font-semibold text-text-main leading-none">{stats.total}</span>
              <span className="text-xs font-semibold text-text-muted">Santri</span>
            </div>
            <p className="text-[10px] text-text-muted mt-2">Total terdaftar di {filterMode === 'kelas' ? `Kelas ${selectedKelas}` : `Jenjang ${selectedJenjang}`}</p>
          </div>

          {/* Card 2: Tuntas */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 hover:border-primary transition-all duration-150 relative overflow-hidden group">
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-primary"></span>
                <span className="text-[11px] font-medium text-text-muted uppercase tracking-wide">Tuntas (Lulus)</span>
              </div>
              <div className="p-2 bg-primary/10 text-primary rounded-md">
                <CheckCircle size={16} />
              </div>
            </div>
            <div className="flex items-baseline justify-between">
              <div className="flex items-baseline gap-1.5">
                <span className="text-[32px] font-semibold text-text-main leading-none">{stats.tuntasCount}</span>
                <span className="text-xs font-semibold text-text-muted">Santri</span>
              </div>
              <span className="text-lg font-bold text-primary">{stats.tuntasPct}%</span>
            </div>
            <p className="text-[10px] text-text-muted mt-2">Mendapatkan predikat kelulusan TUNTAS</p>
          </div>

          {/* Card 3: Tidak Tuntas */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 hover:border-primary transition-all duration-150 relative overflow-hidden group">
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-brand-danger"></span>
                <span className="text-[11px] font-medium text-text-muted uppercase tracking-wide">Tidak Tuntas</span>
              </div>
              <div className="p-2 bg-brand-danger/10 text-brand-danger rounded-md">
                <AlertTriangle size={16} />
              </div>
            </div>
            <div className="flex items-baseline justify-between">
              <div className="flex items-baseline gap-1.5">
                <span className="text-[32px] font-semibold text-text-main leading-none">{stats.tidakTuntasCount}</span>
                <span className="text-xs font-semibold text-text-muted">Santri</span>
              </div>
              <span className="text-lg font-bold text-brand-danger">{stats.tidakTuntasPct}%</span>
            </div>
            <p className="text-[10px] text-text-muted mt-2">Mendapatkan predikat TIDAK TUNTAS</p>
          </div>

          {/* Card 4: Tuntas > 70 */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 hover:border-primary transition-all duration-150 relative overflow-hidden group">
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-emerald-500"></span>
                <span className="text-[11px] font-medium text-text-muted uppercase tracking-wide">Tuntas &gt; 70</span>
              </div>
              <div className="p-2 bg-emerald-500/10 text-emerald-500 rounded-md">
                <TrendingUp size={16} />
              </div>
            </div>
            <div className="flex items-baseline justify-between">
              <div className="flex items-baseline gap-1.5">
                <span className="text-[32px] font-semibold text-text-main leading-none">{stats.tuntasAtas70Count}</span>
                <span className="text-xs font-semibold text-text-muted">Santri</span>
              </div>
              <span className="text-lg font-bold text-emerald-500">{stats.tuntasAtas70Pct}%</span>
            </div>
            <p className="text-[10px] text-text-muted mt-2">Tuntas dengan rata-rata nilai tinggi</p>
          </div>

          {/* Card 5: Tuntas <= 70 */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 hover:border-primary transition-all duration-150 relative overflow-hidden group">
            <div className="flex items-center justify-between mb-3">
              <div className="flex items-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                <span className="text-[11px] font-medium text-text-muted uppercase tracking-wide">Tuntas &le; 70</span>
              </div>
              <div className="p-2 bg-amber-500/10 text-amber-500 rounded-md">
                <TrendingDown size={16} />
              </div>
            </div>
            <div className="flex items-baseline justify-between">
              <div className="flex items-baseline gap-1.5">
                <span className="text-[32px] font-semibold text-text-main leading-none">{stats.tuntasBawah70Count}</span>
                <span className="text-xs font-semibold text-text-muted">Santri</span>
              </div>
              <span className="text-lg font-bold text-amber-500">{stats.tuntasBawah70Pct}%</span>
            </div>
            <p className="text-[10px] text-text-muted mt-2">Tuntas dengan nilai pas-pasan/ambang</p>
          </div>

        </section>        {/* ROW 4: CHARTS GRID (Donut + Bar Chart) */}
        <section className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          
          {/* Donut Chart: Proporsi Kelulusan */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 flex flex-col gap-4 transition-all duration-150">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-text-main">Proporsi Kelulusan Siswa</h3>
                <p className="text-[11px] text-text-muted mt-0.5">Proporsi predikat KKM santri saat ini</p>
              </div>
              <ChartExportMenu chartRef={donutChartRef} filename={`proporsi-kelulusan-${filterMode === 'kelas' ? 'kelas-' + selectedKelas.toLowerCase().replace(' ', '-') : 'jenjang-' + selectedJenjang.toLowerCase()}`} />
            </div>
            <div className="h-48 relative">
              {stats.total === 0 ? (
                <div className="absolute inset-0 flex items-center justify-center text-xs text-text-muted">Tidak ada data.</div>
              ) : (
                <ReactECharts ref={donutChartRef} option={donutChartOptions} opts={{ renderer: 'svg' }} style={{ height: '100%', width: '100%' }} notMerge={true} />
              )}
            </div>
            {stats.total > 0 && (
              <div className="flex flex-col gap-2 mt-2 pt-2 border-t border-border-color/60 select-none">
                <div className="flex items-center justify-between text-xs">
                  <div className="flex items-center gap-2">
                    <span className="w-2 h-2 rounded-full bg-primary" />
                    <span className="text-text-muted font-medium">Tuntas</span>
                  </div>
                  <span className="text-text-main font-semibold">{stats.tuntasCount} ({stats.tuntasPct}%)</span>
                </div>
                <div className="flex items-center justify-between text-xs">
                  <div className="flex items-center gap-2">
                    <span className="w-2 h-2 rounded-full bg-brand-danger" />
                    <span className="text-text-muted font-medium">Tidak Tuntas</span>
                  </div>
                  <span className="text-text-main font-semibold">{stats.tidakTuntasCount} ({stats.tidakTuntasPct}%)</span>
                </div>
                <div className="flex items-center justify-between text-xs">
                  <div className="flex items-center gap-2">
                    <span className="w-2 h-2 rounded-full bg-brand-neutral" />
                    <span className="text-text-muted font-medium">Tidak Hadir/Baca</span>
                  </div>
                  <span className="text-text-main font-semibold">{stats.tidakHadirBacaCount} ({stats.tidakHadirBacaPct}%)</span>
                </div>
              </div>
            )}
          </div>

          {/* Bar Chart: Perbandingan Antar Kelas / Nilai Aspek */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 flex flex-col gap-4 lg:col-span-2 transition-all duration-150">
            <div className="flex items-center justify-between">
              <div>
                <h3 className="text-sm font-bold text-text-main">
                  {filterMode === 'jenjang' 
                    ? `Perbandingan Persentase Tuntas Jenjang ${selectedJenjang}` 
                    : `Rata-rata Nilai Aspek Kelas ${selectedKelas}`}
                </h3>
                <p className="text-[11px] text-text-muted mt-0.5">
                  {filterMode === 'jenjang'
                    ? 'Urutan performa tingkat tuntas tertinggi per kelas'
                    : 'Rata-rata persentase capaian per aspek ujian'}
                </p>
              </div>
              <ChartExportMenu chartRef={barChartRef} filename={`rata-rata-aspek-${filterMode === 'kelas' ? 'kelas-' + selectedKelas.toLowerCase().replace(' ', '-') : 'jenjang-' + selectedJenjang.toLowerCase()}`} />
            </div>
            <div className="h-64 relative">
              {stats.total === 0 ? (
                <div className="absolute inset-0 flex items-center justify-center text-xs text-text-muted">Tidak ada data.</div>
              ) : (
                <ReactECharts ref={barChartRef} option={barChartOptions} opts={{ renderer: 'svg' }} style={{ height: '100%', width: '100%' }} notMerge={true} />
              )}
            </div>
          </div>

        </section>

        {/* ROW 5: TABLES SECTION (Rekap Kelas + Detail Siswa) */}
        <section className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          
          {/* Left Column: Tabel Rekapitulasi Ringkas Kelas/Jenjang */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 flex flex-col gap-4 transition-all duration-150">
            <div>
              <h3 className="text-sm font-bold text-text-main">
                Tabel Rekap {filterMode === 'kelas' ? 'Seluruh Kelas' : 'Seluruh Jenjang'}
              </h3>
              <p className="text-[11px] text-text-muted mt-0.5">Perbandingan performa data secara ringkas</p>
            </div>

            <div className="overflow-x-auto border border-border-color rounded-md">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-bg-app border-b border-border-color text-text-muted font-semibold">
                    <th className="p-3">{filterMode === 'kelas' ? 'Kelas' : 'Jenjang'}</th>
                    <th className="p-3 text-center">Total</th>
                    <th className="p-3 text-center">Tuntas</th>
                    <th className="p-3 text-right">% Tuntas</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border-color text-text-main">
                  {comparisonData.map(c => (
                    <tr 
                      key={c.name} 
                      className={`hover:bg-primary/5 transition-all cursor-pointer ${
                        (filterMode === 'kelas' && selectedKelas === c.name) || (filterMode === 'jenjang' && selectedJenjang === c.name)
                          ? 'bg-primary/5 font-semibold'
                          : ''
                      }`}
                      onClick={() => {
                        if (filterMode === 'kelas') setSelectedKelas(c.name);
                        else setSelectedJenjang(c.name);
                      }}
                    >
                      <td className="p-3">{c.name}</td>
                      <td className="p-3 text-center">{c.total}</td>
                      <td className="p-3 text-center">{c.tuntas}</td>
                      <td className="p-3 text-right text-primary font-bold">{c.tuntasPct}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Right Column: Detailed Students Records Table */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 flex flex-col gap-4 lg:col-span-2 transition-all duration-150">
            
            <div className="flex flex-col sm:flex-row justify-between sm:items-center gap-3">
              <div>
                <h3 className="text-sm font-bold text-text-main">
                  Rincian Nilai Santri ({filteredStudents.length} siswa)
                </h3>
                <p className="text-[11px] text-text-muted mt-0.5">Daftar nilai riil ujian baca kitab santri</p>
              </div>

              {/* Table search controls */}
              <div className="flex items-center gap-2">
                <div className="relative flex items-center">
                  <Search className="absolute left-2.5 text-text-muted" size={14} />
                  <input
                    type="text"
                    placeholder="Cari santri..."
                    className="pl-8 pr-3 py-1.5 bg-bg-app border border-border-color rounded-md text-xs outline-none focus:border-primary text-text-main transition-all duration-150"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                  />
                </div>
                <CustomSelect
                  value={studentStatusFilter}
                  onChange={setStudentStatusFilter}
                  options={[
                    { label: 'Semua KKM', value: 'Semua' },
                    { label: 'TUNTAS', value: 'TUNTAS' },
                    { label: 'TIDAK TUNTAS', value: 'TIDAK TUNTAS' },
                    { label: 'TIDAK HADIR', value: 'TIDAK HADIR' },
                    { label: 'TIDAK BACA', value: 'TIDAK BACA' }
                  ]}
                  ariaLabel="Saring Status KKM"
                  className="w-36"
                />
              </div>
            </div>

            <div className="overflow-x-auto border border-border-color rounded-md">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-bg-app border-b border-border-color text-text-muted font-semibold">
                    <th className="p-3">Nopes</th>
                    <th className="p-3">Nama Santri</th>
                    <th className="p-3 text-center">Kelas</th>
                    <th className="p-3 text-center">Sihhah (30)</th>
                    <th className="p-3 text-center">Mufrodat (30)</th>
                    <th className="p-3 text-center">Murad (5)</th>
                    <th className="p-3 text-center">Rata-rata</th>
                    <th className="p-3 text-right">Status KKM</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-border-color text-text-main">
                  {paginatedStudents.length > 0 ? (
                    paginatedStudents.map(student => (
                      <tr key={student.nopes} className="hover:bg-primary/5 transition-all">
                        <td className="p-3 font-semibold">{student.nopes}</td>
                        <td className="p-3 truncate max-w-[120px] font-medium" title={student.nama}>{student.nama}</td>
                        <td className="p-3 text-center">{student.kelas}</td>
                        <td className="p-3 text-center">{student.sihhah}</td>
                        <td className="p-3 text-center">{student.makna_mufrodat}</td>
                        <td className="p-3 text-center">{student.makna_murad}</td>
                        <td className="p-3 text-center font-semibold text-text-main">{student.rata2}</td>
                        <td className="p-3 text-right">
                          <span className={`px-2.5 py-0.5 rounded-full text-[10px] font-semibold border ${
                            student.kkm === 'TUNTAS' ? 'bg-primary-soft text-primary border-primary/20' :
                            student.kkm === 'TIDAK TUNTAS' ? 'bg-red-500/10 text-brand-danger border-brand-danger/20' :
                            'bg-brand-neutral/10 text-brand-neutral border-brand-neutral/20'
                          }`}>
                            {student.kkm}
                          </span>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="8" className="p-6 text-center text-text-muted">
                        Tidak ada data santri yang cocok.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls */}
            {totalPages > 1 && (
              <div className="flex items-center justify-between mt-2">
                <span className="text-[11px] text-text-muted">
                  Menampilkan {Math.min(filteredStudents.length, (currentPage - 1) * itemsPerPage + 1)}-
                  {Math.min(filteredStudents.length, currentPage * itemsPerPage)} dari {filteredStudents.length} santri
                </span>
                <div className="flex items-center gap-1">
                  <button
                    className="p-1.5 rounded-md border border-border-color text-text-muted hover:text-text-main hover:bg-bg-app disabled:opacity-50 transition-all duration-150 cursor-pointer"
                    disabled={currentPage === 1}
                    onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                  >
                    <ChevronLeft size={16} />
                  </button>
                  <span className="text-xs px-2 text-text-main font-semibold">
                    {currentPage} / {totalPages}
                  </span>
                  <button
                    className="p-1.5 rounded-md border border-border-color text-text-muted hover:text-text-main hover:bg-bg-app disabled:opacity-50 transition-all duration-150 cursor-pointer"
                    disabled={currentPage === totalPages}
                    onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                  >
                    <ChevronRight size={16} />
                  </button>
                </div>
              </div>
            )}

          </div>

        </section>

      </main>

      {/* DASHBOARD FOOTER */}
      <footer className="bg-bg-card border-t border-border-color text-text-muted py-6 mt-16 text-center text-xs transition-colors duration-150">
        <div className="w-full px-4 sm:px-6">
          <p>© 2026 MDT Miftahul Ulum — Pondok Pesantren Miftahul Ulum Banyuputih Kidul. Hak Cipta Dilindungi.</p>
        </div>
      </footer>

    </div>
  );
}

export default App;
