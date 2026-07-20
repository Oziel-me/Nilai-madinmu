import React, { useState, useEffect, useMemo, useRef } from 'react';
import ReactECharts from 'echarts-for-react';
import { motion, AnimatePresence } from 'framer-motion';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import * as htmlToImage from 'html-to-image';
import { jsPDF } from 'jspdf';
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
  TrendingDown,
  Eye,
  EyeOff,
  Printer,
  BarChart3,
  FileText
} from 'lucide-react';



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

  const handleExport = (type) => {
    setIsOpen(false);
    if (!chartRef.current) return;

    try {
      const echartsInstance = chartRef.current.getEchartsInstance();
      const dom = echartsInstance.getDom();
      const svgEl = dom.querySelector('svg');
      if (!svgEl) {
        throw new Error("Elemen SVG grafik tidak ditemukan.");
      }

      const serializer = new XMLSerializer();
      let source = serializer.serializeToString(svgEl);
      
      // Ensure namespaces are correct
      if (!source.includes('xmlns="http://www.w3.org/2000/svg"')) {
        source = source.replace(/^<svg/, '<svg xmlns="http://www.w3.org/2000/svg"');
      }
      if (!source.includes('xmlns:xlink="http://www.w3.org/1999/xlink"')) {
        source = source.replace(/^<svg/, '<svg xmlns:xlink="http://www.w3.org/1999/xlink"');
      }

      if (type === 'svg') {
        const xmlString = `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n${source}`;
        const svgBlob = new Blob([xmlString], { type: 'image/svg+xml;charset=utf-8' });
        
        if (svgBlob.size === 0) {
          throw new Error("File SVG yang dihasilkan kosong (0 byte).");
        }
        
        const url = window.URL.createObjectURL(svgBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${filename}.svg`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
      } else {
        // Render raster PNG/JPG using native image + canvas rendering of SVG source
        const svgUrl = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(`<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n${source}`);
        const img = new Image();
        img.onload = () => {
          try {
            const canvas = document.createElement('canvas');
            const bbox = svgEl.getBoundingClientRect();
            const width = bbox.width || 400;
            const height = bbox.height || 300;
            
            // Set 2x canvas size for high-DPI rendering
            canvas.width = width * 2;
            canvas.height = height * 2;
            
            const ctx = canvas.getContext('2d');
            ctx.scale(2, 2);
            
            if (type === 'jpg') {
              const isDark = document.documentElement.classList.contains('dark');
              ctx.fillStyle = isDark ? '#18191b' : '#ffffff';
              ctx.fillRect(0, 0, width, height);
            } else {
              ctx.clearRect(0, 0, width, height);
            }
            
            ctx.drawImage(img, 0, 0, width, height);
            
            const outputMime = type === 'png' ? 'image/png' : 'image/jpeg';
            canvas.toBlob((blob) => {
              if (!blob || blob.size === 0) {
                alert("Gagal merender gambar (0 byte).");
                return;
              }
              const url = window.URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = url;
              a.download = `${filename}.${type}`;
              document.body.appendChild(a);
              a.click();
              document.body.removeChild(a);
              window.URL.revokeObjectURL(url);
            }, outputMime, type === 'jpg' ? 0.95 : undefined);
          } catch (e) {
            console.error(e);
            alert(`Gagal memproses gambar canvas: ${e.message}`);
          }
        };
        img.onerror = (err) => {
          console.error(err);
          alert("Gagal memuat gambar SVG untuk ekspor.");
        };
        img.src = svgUrl;
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

// --- PDF EXPORT MENU (Dropdown ekspor PDF kustom) ---
const PDFExportMenu = ({ onExport, isExporting }) => {
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

  return (
    <div className="relative" ref={containerRef}>
      <button
        onClick={() => setIsOpen(!isOpen)}
        disabled={isExporting}
        className="p-1.5 rounded-md border border-border-color bg-bg-card text-text-muted hover:text-text-main hover:bg-bg-app disabled:opacity-50 transition-all duration-150 cursor-pointer flex items-center justify-center gap-1.5 font-semibold text-xs h-8"
        title="Ekspor Laporan PDF"
      >
        {isExporting ? (
          <RefreshCw size={14} className="animate-spin text-primary" />
        ) : (
          <Download size={14} />
        )}
        <span className="hidden md:inline">Ekspor PDF</span>
      </button>
      {isOpen && (
        <ul className="absolute right-0 mt-1 z-50 w-56 bg-bg-card border border-border-color rounded-md shadow-lg py-1 text-xs select-none">
          <li
            onClick={() => { onExport('all'); setIsOpen(false); }}
            className="px-4 py-2.5 text-text-main hover:bg-primary-soft hover:text-primary cursor-pointer transition-colors flex items-center gap-2.5"
          >
            <Printer size={14} className="text-blue-600 dark:text-blue-400 shrink-0" />
            <span>Cetak Laporan Lengkap (A4)</span>
          </li>
          <li
            onClick={() => { onExport('charts'); setIsOpen(false); }}
            className="px-4 py-2.5 text-text-main hover:bg-primary-soft hover:text-primary cursor-pointer transition-colors flex items-center gap-2.5"
          >
            <BarChart3 size={14} className="text-primary shrink-0" />
            <span>Unduh Grafik Saja (PDF 1 Halaman)</span>
          </li>
          <li
            onClick={() => { onExport('data'); setIsOpen(false); }}
            className="px-4 py-2.5 text-text-main hover:bg-primary-soft hover:text-primary cursor-pointer transition-colors flex items-center gap-2.5"
          >
            <FileText size={14} className="text-amber-600 dark:text-amber-400 shrink-0" />
            <span>Cetak Tabel Data Saja (A4)</span>
          </li>
        </ul>
      )}
    </div>
  );
};

function App() {
  // --- STATE MANAGEMENT ---
  const [records, setRecords] = useState([]);
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
  const [showRincianCategory, setShowRincianCategory] = useState(true);
  const [printMode, setPrintMode] = useState(null); // null | 'all' | 'charts' | 'data'
  const [isExporting, setIsExporting] = useState(false);
  const fileInputRef = useRef(null);
  const donutChartRef = useRef(null);
  const barChartRef = useRef(null);
  const keseluruhanChartRef = useRef(null);
  const tuntasChartRef = useRef(null);
  const tidakTuntasChartRef = useRef(null);
  const tuntasAtas70ChartRef = useRef(null);
  const tuntasBawah70ChartRef = useRef(null);

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

  const DEFAULT_RECORDS = [];

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

          const kelas = sheetName.trim();

          const headerRow = rows[headerIdx];
          const subHeaderRow = rows[headerIdx + 1] || [];
          const colMap = {};

          headerRow.forEach((c, idx) => {
            if (!c) return;
            const cleanHeader = String(c).toUpperCase().trim();
            if (cleanHeader.includes('NOPES')) colMap.nopes = idx;
            else if (cleanHeader.includes('NAMA')) colMap.nama = idx;
            else if (cleanHeader.includes('ASR')) colMap.asr = idx;
            else if (cleanHeader.includes('JML') || cleanHeader.includes('JUMLAH')) colMap.jml = idx;
            else if (cleanHeader.includes('RT') || cleanHeader.includes('RATA2') || cleanHeader.includes('RATA-RATA')) colMap.rata2 = idx;
            else if (cleanHeader.includes('RANK')) colMap.rank = idx;
            else if (cleanHeader.includes('KKM')) colMap.kkm = idx;
            else if (cleanHeader.includes('KET')) colMap.ket = idx;
            else if (cleanHeader.includes('SIHHAH') || cleanHeader.includes('QIRO')) colMap.sihhah = idx;
            else if (cleanHeader.includes('MUFRODAT') || cleanHeader.includes('MAKNA MUF')) colMap.mufrodat = idx;
            else if (cleanHeader.includes('MURAD') || cleanHeader.includes('MAKNA MUR')) colMap.murad = idx;
          });

          subHeaderRow.forEach((c, idx) => {
            if (!c) return;
            const cleanHeader = String(c).toUpperCase().trim();
            if (cleanHeader.includes('SIHHAH') || cleanHeader.includes('QIRO')) colMap.sihhah = idx;
            else if (cleanHeader.includes('MUFRODAT') || cleanHeader.includes('MAKNA MUF')) colMap.mufrodat = idx;
            else if (cleanHeader.includes('MURAD') || cleanHeader.includes('MAKNA MUR')) colMap.murad = idx;
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
              jenjang: (() => {
                const clean = kelas.trim();
                const match = clean.match(/^([0-9]+|[ivxIVX]+)/);
                return match ? match[0].toUpperCase() : (clean.split(' ')[0] || 'Lainnya');
              })(),
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
    try {
      const wb = XLSX.utils.book_new();
      
      const wsData = [
        ['PENILAIAN UJIAN BACA KITAB'],
        ['MADRASAH DINIYAH TAKMILIYAH MIFTAHUL ULUM'],
        ['Banyuputih Kidul Jatiroto Lumajang'],
        ['ASUHAN : UST. FIRRUL AZIZ', '', '', '', '', '', '', '', 'KELAS: I A'],
        [],
        ['NO', 'NOPES', 'NAMA', 'ASR', 'ASPEK PENILAIAN', '', '', 'JML', 'RT', 'RANK', 'KKM', 'KET'],
        ['', '', '', '', "Sihhah Al-Qiro'ah", 'Makna Mufrodat', 'Makna Murad', '', '', '', '', ''],
        ['', '', '', '', 'MAX 30', 'MAX 30', 'MAX 5', '', '', '', '60', ''],
        [1, '07001', 'M Khoironi Hasan', 'A 01', 26, 25, 3, 222, 74.0, 1, 'TUNTAS', ''],
        [2, '07002', 'Achmad Naufal', 'A 06', 11, 11, 11, 33, 11.0, 2, 'TIDAK HADIR', 'BOYONG'],
        [3, '07003', 'Rama Afriyanto', 'A 10', 17, 15, 6, 114, 38.0, 3, 'TIDAK TUNTAS', 'TM']
      ];

      const ws = XLSX.utils.aoa_to_sheet(wsData);

      ws['!merges'] = [
        { s: { r: 5, c: 4 }, e: { r: 5, c: 6 } }, // ASPEK PENILAIAN
        { s: { r: 5, c: 0 }, e: { r: 7, c: 0 } }, // NO
        { s: { r: 5, c: 1 }, e: { r: 7, c: 1 } }, // NOPES
        { s: { r: 5, c: 2 }, e: { r: 7, c: 2 } }, // NAMA
        { s: { r: 5, c: 3 }, e: { r: 7, c: 3 } }, // ASR
        { s: { r: 5, c: 7 }, e: { r: 7, c: 7 } }, // JML
        { s: { r: 5, c: 8 }, e: { r: 7, c: 8 } }, // RT
        { s: { r: 5, c: 9 }, e: { r: 7, c: 9 } }, // RANK
        { s: { r: 5, c: 10 }, e: { r: 6, c: 10 } }, // KKM
        { s: { r: 5, c: 11 }, e: { r: 7, c: 11 } }  // KET
      ];

      XLSX.utils.book_append_sheet(wb, ws, 'I A');

      // Generate binary string for download
      const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
      const s2ab = (s) => {
        const buf = new ArrayBuffer(s.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
      };
      
      const fileBlob = new Blob([s2ab(wbout)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      if (fileBlob.size === 0) {
        throw new Error("File Excel yang dihasilkan kosong (0 byte).");
      }
      saveAs(fileBlob, 'template-penilaian-ujian-baca-kitab.xlsx');
    } catch (error) {
      console.error(error);
      alert(`Gagal mengunduh template Excel: ${error.message}`);
    }
  };

  const handleResetData = () => {
    setRecords(DEFAULT_RECORDS);
    setFileName('');
    setCurrentPage(1);
  };

  const handleExportPDF = async (type) => {
    if (records.length === 0) {
      alert("Tidak ada data untuk diekspor! Silakan unggah berkas Excel terlebih dahulu.");
      return;
    }

    if (type === 'charts') {
      setIsExporting(true);
      // Beri jeda kecil agar React merender logo & judul PDF formal
      setTimeout(async () => {
        const element = document.getElementById('visual-charts-section');
        if (!element) {
          alert("Elemen grafik tidak ditemukan!");
          setIsExporting(false);
          return;
        }
        try {
          const dataUrl = await htmlToImage.toPng(element, {
            backgroundColor: darkMode ? '#18191b' : '#ffffff',
            style: {
              transform: 'scale(1)',
              transformOrigin: 'top left',
              width: element.offsetWidth + 'px',
              height: element.offsetHeight + 'px'
            },
            pixelRatio: 3, // Perbesar resolusi gambar 3x agar sangat tajam saat dicetak/zoom
            filter: (node) => {
              if (node.tagName === 'BUTTON') return false;
              if (node.classList && (node.classList.contains('print:hidden') || node.classList.contains('no-print'))) {
                return false;
              }
              return true;
            },
            quality: 1.0
          });

          const pdf = new jsPDF({
            orientation: 'landscape',
            unit: 'px',
            format: [element.offsetWidth, element.offsetHeight]
          });

          pdf.addImage(dataUrl, 'PNG', 0, 0, element.offsetWidth, element.offsetHeight);
          const fileBase = fileName ? fileName.replace(/\.[^/.]+$/, "") : "laporan-nilai";
          const catName = filterMode === 'kelas' ? `kelas-${selectedKelas}` : `jenjang-${selectedJenjang}`;
          pdf.save(`grafik-${fileBase}-${catName}.pdf`);
        } catch (error) {
          console.error("Gagal mengekspor grafik ke PDF:", error);
          alert(`Gagal mengekspor grafik ke PDF: ${error.message}`);
        } finally {
          setIsExporting(false);
        }
      }, 200);
    } else {
      setPrintMode(type);
      setTimeout(() => {
        try {
          window.print();
        } catch (error) {
          console.error("Gagal melakukan pencetakan:", error);
        } finally {
          setPrintMode(null);
        }
      }, 500);
    }
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

  // Donut Chart - proporsi kelulusan siswa (diubah ke Column Chart minimalis)
  const donutChartOptions = useMemo(() => {
    const data = [
      { name: 'Tuntas', value: stats.tuntasCount, pct: stats.tuntasPct, color: themeColors.primary },
      { name: 'Tidak Tuntas', value: stats.tidakTuntasCount, pct: stats.tidakTuntasPct, color: themeColors.danger },
      { name: 'Tidak Hadir/Baca', value: stats.tidakHadirBacaCount, pct: stats.tidakHadirBacaPct, color: themeColors.neutral }
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
      grid: { left: '8%', right: '8%', bottom: '10%', top: '15%', containLabel: true },
      xAxis: {
        type: 'category',
        data: data.map(d => d.name),
        axisLabel: { color: themeColors.text, fontSize: 10 },
        axisLine: { lineStyle: { color: themeColors.border } }
      },
      yAxis: {
        type: 'value',
        axisLabel: { color: themeColors.textMuted, fontSize: 10 },
        splitLine: { show: true, lineStyle: { type: 'dashed', color: themeColors.border } },
        splitNumber: 4
      },
      series: [
        {
          name: 'Jumlah Santri',
          type: 'bar',
          barWidth: 20,
          data: data.map(d => ({
            value: d.value,
            itemStyle: { color: d.color, borderRadius: [4, 4, 0, 0] }
          })),
          label: {
            show: true,
            position: 'top',
            formatter: (params) => {
              const idx = params.dataIndex;
              return `${params.value} (${data[idx].pct}%)`;
            },
            color: themeColors.text,
            fontFamily: 'Inter',
            fontSize: 9,
            fontWeight: 'medium'
          },
          animationDuration: 350,
          animationEasing: 'cubicOut'
        }
      ]
    };
  }, [stats, themeColors]);

  // Bar Chart - Perbandingan kelas dalam satu jenjang / rata-rata nilai aspek kelas (diubah ke Column Chart minimalis)
  const barChartOptions = useMemo(() => {
    const dinilai = currentRecords.filter(r => r.kkm !== 'TIDAK HADIR' && r.kkm !== 'TIDAK BACA');
    const avgSihhahMistakes = dinilai.length === 0 ? 0 : dinilai.reduce((s, r) => s + r.sihhah, 0) / dinilai.length;
    const avgMufrodatMistakes = dinilai.length === 0 ? 0 : dinilai.reduce((s, r) => s + r.makna_mufrodat, 0) / dinilai.length;
    const avgMuradMistakes = dinilai.length === 0 ? 0 : dinilai.reduce((s, r) => s + r.makna_murad, 0) / dinilai.length;

    const pctSihhah = Math.max(0, +(100 - avgSihhahMistakes * 3).toFixed(1));
    const pctMufrodat = Math.max(0, +(100 - avgMufrodatMistakes * 3).toFixed(1));
    const pctMurad = Math.max(0, +(100 - avgMuradMistakes * 15).toFixed(1));

    const items = [
      { name: 'Sihhah Al-Qiro\'ah', value: pctSihhah, color: themeColors.primary },
      { name: 'Makna Mufrodat', value: pctMufrodat, color: themeColors.emerald },
      { name: 'Makna Murad', value: pctMurad, color: themeColors.slate }
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
      grid: { left: '5%', right: '5%', bottom: '10%', top: '15%', containLabel: true },
      xAxis: {
        type: 'category',
        data: items.map(i => i.name),
        axisLabel: { color: themeColors.text, fontSize: 10 },
        axisLine: { lineStyle: { color: themeColors.border } }
      },
      yAxis: {
        type: 'value',
        axisLabel: { color: themeColors.textMuted, fontSize: 10 },
        splitLine: { show: true, lineStyle: { type: 'dashed', color: themeColors.border } },
        max: 100,
        splitNumber: 4
      },
      series: [
        {
          name: 'Capaian Rata-rata',
          type: 'bar',
          barWidth: 20,
          data: items.map(i => ({
            value: i.value,
            itemStyle: { color: i.color, borderRadius: [4, 4, 0, 0] }
          })),
          label: {
            show: true,
            position: 'top',
            formatter: '{c}%',
            color: themeColors.text,
            fontFamily: 'Inter',
            fontSize: 9,
            fontWeight: 'medium'
          },
          animationDuration: 350,
          animationEasing: 'cubicOut'
        }
      ]
    };
  }, [filterMode, records, selectedJenjang, pctBasis, currentRecords, themeColors, darkMode]);


  // --- INFOGRAPHIC CATEGORY OPTIONS (Semua grafik menggunakan tipe Area Chart secara seragam) ---
  const categoryChartOptions = useMemo(() => {
    if (currentRecords.length === 0) return null;

    const classesInJenjang = [...new Set(records.filter(r => r.jenjang === selectedJenjang).map(r => r.kelas))].sort();

    // 1. KESELURUHAN -> Area Chart
    const total = currentRecords.length;
    const dinilai = currentRecords.filter(r => r.kkm !== 'TIDAK HADIR' && r.kkm !== 'TIDAK BACA');
    const tidakHadirOrBaca = currentRecords.filter(r => r.kkm === 'TIDAK HADIR' || r.kkm === 'TIDAK BACA');
    const tuntas = dinilai.filter(r => r.kkm === 'TUNTAS');
    const tidakTuntas = dinilai.filter(r => r.kkm === 'TIDAK TUNTAS');
    
    const denom = pctBasis === 'total' ? total : dinilai.length;
    const getPct = (n) => (denom === 0 ? 0 : +(n / denom * 100).toFixed(1));

    const keseluruhanData = [
      { name: 'Tuntas', value: tuntas.length, pct: getPct(tuntas.length) },
      { name: 'Tidak Tuntas', value: tidakTuntas.length, pct: getPct(tidakTuntas.length) },
      { name: 'Tidak Hadir/Baca', value: tidakHadirOrBaca.length, pct: getPct(tidakHadirOrBaca.length) }
    ];

    const keseluruhanChart = {
      tooltip: {
        trigger: 'axis',
        backgroundColor: themeColors.bgTooltip,
        borderColor: themeColors.border,
        borderWidth: 1,
        borderRadius: 8,
        textStyle: { color: themeColors.text, fontFamily: 'Inter', fontSize: 11 }
      },
      grid: { left: '5%', right: '5%', bottom: '10%', top: '15%', containLabel: true },
      xAxis: {
        type: 'category',
        data: keseluruhanData.map(d => d.name),
        boundaryGap: false,
        axisLabel: { color: themeColors.text, fontSize: 10 },
        axisLine: { lineStyle: { color: themeColors.border } }
      },
      yAxis: {
        type: 'value',
        axisLabel: { color: themeColors.textMuted, fontSize: 10 },
        splitLine: { show: true, lineStyle: { type: 'dashed', color: themeColors.border } },
        splitNumber: 4
      },
      series: [
        {
          name: 'Jumlah Santri',
          type: 'line',
          smooth: true,
          symbol: 'circle',
          symbolSize: 6,
          lineStyle: { width: 2, color: themeColors.primary },
          itemStyle: { color: themeColors.primary },
          areaStyle: {
            color: darkMode ? 'rgba(34, 197, 94, 0.15)' : 'rgba(22, 163, 74, 0.1)'
          },
          data: keseluruhanData.map(d => d.value),
          label: {
            show: true,
            position: 'top',
            formatter: (params) => {
              const idx = params.dataIndex;
              return `${params.value} (${keseluruhanData[idx].pct}%)`;
            },
            color: themeColors.text,
            fontFamily: 'Inter',
            fontSize: 9,
            fontWeight: 'medium'
          },
          animationDuration: 300,
          animationEasing: 'cubicOut'
        }
      ]
    };

    // 2. TUNTAS -> Area Chart
    let tuntasX = [];
    let tuntasY = [];
    if (filterMode === 'jenjang') {
      tuntasX = classesInJenjang.map(k => `Kelas ${k}`);
      tuntasY = classesInJenjang.map(k => records.filter(r => r.kelas === k && r.kkm === 'TUNTAS').length);
    } else {
      tuntasX = ['50-60', '61-70', '71-80', '81-90', '91-100'];
      const tuntasRecs = currentRecords.filter(r => r.kkm === 'TUNTAS');
      const binCounts = [0, 0, 0, 0, 0];
      tuntasRecs.forEach(r => {
        const v = r.rata2;
        if (v >= 50 && v <= 60) binCounts[0]++;
        else if (v > 60 && v <= 70) binCounts[1]++;
        else if (v > 70 && v <= 80) binCounts[2]++;
        else if (v > 80 && v <= 90) binCounts[3]++;
        else if (v > 90 && v <= 100) binCounts[4]++;
      });
      tuntasY = binCounts;
    }

    const tuntasChart = {
      tooltip: {
        trigger: 'axis',
        backgroundColor: themeColors.bgTooltip,
        borderColor: themeColors.border,
        borderWidth: 1,
        borderRadius: 8,
        textStyle: { color: themeColors.text, fontFamily: 'Inter', fontSize: 11 }
      },
      grid: { left: '5%', right: '5%', bottom: '10%', top: '15%', containLabel: true },
      xAxis: {
        type: 'category',
        data: tuntasX,
        boundaryGap: false,
        axisLabel: { color: themeColors.text, fontSize: 10 },
        axisLine: { lineStyle: { color: themeColors.border } }
      },
      yAxis: {
        type: 'value',
        axisLabel: { color: themeColors.textMuted, fontSize: 10 },
        splitLine: { show: true, lineStyle: { type: 'dashed', color: themeColors.border } },
        splitNumber: 4
      },
      series: [
        {
          name: 'Jumlah Tuntas',
          type: 'line',
          smooth: true,
          symbol: 'circle',
          symbolSize: 6,
          lineStyle: { width: 2, color: themeColors.primary },
          itemStyle: { color: themeColors.primary },
          areaStyle: {
            color: darkMode ? 'rgba(34, 197, 94, 0.15)' : 'rgba(22, 163, 74, 0.1)'
          },
          data: tuntasY,
          label: {
            show: true,
            position: 'top',
            color: themeColors.text,
            fontFamily: 'Inter',
            fontSize: 9,
            fontWeight: 'medium'
          },
          animationDuration: 300,
          animationEasing: 'cubicOut'
        }
      ]
    };

    // 3. TIDAK TUNTAS -> Area Chart
    let tidakTuntasX = [];
    let tidakTuntasY = [];
    if (filterMode === 'jenjang') {
      tidakTuntasX = classesInJenjang.map(k => `Kelas ${k}`);
      tidakTuntasY = classesInJenjang.map(k => records.filter(r => r.kelas === k && r.kkm === 'TIDAK TUNTAS').length);
    } else {
      tidakTuntasX = ['0-20', '21-40', '41-50', '51-60', '61-70'];
      const tidakTuntasRecs = currentRecords.filter(r => r.kkm === 'TIDAK TUNTAS');
      const binCounts = [0, 0, 0, 0, 0];
      tidakTuntasRecs.forEach(r => {
        const v = r.rata2;
        if (v >= 0 && v <= 20) binCounts[0]++;
        else if (v > 20 && v <= 40) binCounts[1]++;
        else if (v > 40 && v <= 50) binCounts[2]++;
        else if (v > 50 && v <= 60) binCounts[3]++;
        else if (v > 60 && v <= 70) binCounts[4]++;
      });
      tidakTuntasY = binCounts;
    }

    const tidakTuntasChart = {
      tooltip: {
        trigger: 'axis',
        backgroundColor: themeColors.bgTooltip,
        borderColor: themeColors.border,
        borderWidth: 1,
        borderRadius: 8,
        textStyle: { color: themeColors.text, fontFamily: 'Inter', fontSize: 11 }
      },
      grid: { left: '5%', right: '5%', bottom: '10%', top: '15%', containLabel: true },
      xAxis: {
        type: 'category',
        data: tidakTuntasX,
        boundaryGap: false,
        axisLabel: { color: themeColors.text, fontSize: 10 },
        axisLine: { lineStyle: { color: themeColors.border } }
      },
      yAxis: {
        type: 'value',
        axisLabel: { color: themeColors.textMuted, fontSize: 10 },
        splitLine: { show: true, lineStyle: { type: 'dashed', color: themeColors.border } },
        splitNumber: 4
      },
      series: [
        {
          name: 'Jumlah Tidak Tuntas',
          type: 'line',
          smooth: true,
          symbol: 'circle',
          symbolSize: 6,
          lineStyle: { width: 2, color: themeColors.danger },
          itemStyle: { color: themeColors.danger },
          areaStyle: {
            color: darkMode ? 'rgba(239, 68, 68, 0.15)' : 'rgba(220, 38, 38, 0.1)'
          },
          data: tidakTuntasY,
          label: {
            show: true,
            position: 'top',
            color: themeColors.text,
            fontFamily: 'Inter',
            fontSize: 9,
            fontWeight: 'medium'
          },
          animationDuration: 300,
          animationEasing: 'cubicOut'
        }
      ]
    };

    // 4. TUNTAS > 70 -> Area Chart
    let tuntasAtas70X = [];
    let tuntasAtas70Y = [];
    if (filterMode === 'jenjang') {
      tuntasAtas70X = classesInJenjang.map(k => `Kelas ${k}`);
      tuntasAtas70Y = classesInJenjang.map(k => records.filter(r => r.kelas === k && r.kkm === 'TUNTAS' && r.rata2 > 70).length);
    } else {
      tuntasAtas70X = ['71-80', '81-90', '91-100'];
      const tuntasRecs = currentRecords.filter(r => r.kkm === 'TUNTAS' && r.rata2 > 70);
      const binCounts = [0, 0, 0];
      tuntasRecs.forEach(r => {
        const v = r.rata2;
        if (v >= 71 && v <= 80) binCounts[0]++;
        else if (v > 80 && v <= 90) binCounts[1]++;
        else if (v > 90 && v <= 100) binCounts[2]++;
      });
      tuntasAtas70Y = binCounts;
    }

    const tuntasAtas70Chart = {
      tooltip: {
        trigger: 'axis',
        backgroundColor: themeColors.bgTooltip,
        borderColor: themeColors.border,
        borderWidth: 1,
        borderRadius: 8,
        textStyle: { color: themeColors.text, fontFamily: 'Inter', fontSize: 11 }
      },
      grid: { left: '5%', right: '5%', bottom: '10%', top: '15%', containLabel: true },
      xAxis: {
        type: 'category',
        data: tuntasAtas70X,
        boundaryGap: false,
        axisLabel: { color: themeColors.text, fontSize: 10 },
        axisLine: { lineStyle: { color: themeColors.border } }
      },
      yAxis: {
        type: 'value',
        axisLabel: { color: themeColors.textMuted, fontSize: 10 },
        splitLine: { show: true, lineStyle: { type: 'dashed', color: themeColors.border } },
        splitNumber: 4
      },
      series: [
        {
          name: 'Jumlah Tuntas > 70',
          type: 'line',
          smooth: true,
          symbol: 'circle',
          symbolSize: 6,
          lineStyle: { width: 2, color: themeColors.emerald },
          itemStyle: { color: themeColors.emerald },
          areaStyle: {
            color: darkMode ? 'rgba(16, 185, 129, 0.15)' : 'rgba(5, 150, 105, 0.1)'
          },
          data: tuntasAtas70Y,
          label: {
            show: true,
            position: 'top',
            color: themeColors.text,
            fontFamily: 'Inter',
            fontSize: 9,
            fontWeight: 'medium'
          },
          animationDuration: 300,
          animationEasing: 'cubicOut'
        }
      ]
    };

    // 5. TUNTAS <= 70 -> Area Chart
    let tuntasBawah70X = [];
    let tuntasBawah70Y = [];
    if (filterMode === 'jenjang') {
      tuntasBawah70X = classesInJenjang.map(k => `Kelas ${k}`);
      tuntasBawah70Y = classesInJenjang.map(k => records.filter(r => r.kelas === k && r.kkm === 'TUNTAS' && r.rata2 <= 70).length);
    } else {
      tuntasBawah70X = ['50-60', '61-70'];
      const tuntasRecs = currentRecords.filter(r => r.kkm === 'TUNTAS' && r.rata2 <= 70);
      const binCounts = [0, 0];
      tuntasRecs.forEach(r => {
        const v = r.rata2;
        if (v >= 50 && v <= 60) binCounts[0]++;
        else if (v > 60 && v <= 70) binCounts[1]++;
      });
      tuntasBawah70Y = binCounts;
    }

    const tuntasBawah70Chart = {
      tooltip: {
        trigger: 'axis',
        backgroundColor: themeColors.bgTooltip,
        borderColor: themeColors.border,
        borderWidth: 1,
        borderRadius: 8,
        textStyle: { color: themeColors.text, fontFamily: 'Inter', fontSize: 11 }
      },
      grid: { left: '5%', right: '5%', bottom: '10%', top: '15%', containLabel: true },
      xAxis: {
        type: 'category',
        data: tuntasBawah70X,
        boundaryGap: false,
        axisLabel: { color: themeColors.text, fontSize: 10 },
        axisLine: { lineStyle: { color: themeColors.border } }
      },
      yAxis: {
        type: 'value',
        axisLabel: { color: themeColors.textMuted, fontSize: 10 },
        splitLine: { show: true, lineStyle: { type: 'dashed', color: themeColors.border } },
        splitNumber: 4
      },
      series: [
        {
          name: 'Jumlah Tuntas ≤ 70',
          type: 'line',
          smooth: true,
          symbol: 'circle',
          symbolSize: 6,
          lineStyle: { width: 2, color: themeColors.slate },
          itemStyle: { color: themeColors.slate },
          areaStyle: {
            color: darkMode ? 'rgba(148, 163, 184, 0.15)' : 'rgba(71, 85, 105, 0.1)'
          },
          data: tuntasBawah70Y,
          label: {
            show: true,
            position: 'top',
            color: themeColors.text,
            fontFamily: 'Inter',
            fontSize: 9,
            fontWeight: 'medium'
          },
          animationDuration: 300,
          animationEasing: 'cubicOut'
        }
      ]
    };

    return {
      keseluruhan: keseluruhanChart,
      tuntasChart,
      tidakTuntasChart,
      tuntasAtas70Chart,
      tuntasBawah70Chart
    };
  }, [currentRecords, filterMode, records, selectedJenjang, selectedKelas, themeColors, pctBasis, darkMode]);

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

  const studentsToRender = useMemo(() => {
    return printMode ? filteredStudents : paginatedStudents;
  }, [printMode, filteredStudents, paginatedStudents]);

  useEffect(() => {
    setCurrentPage(1);
  }, [searchQuery, studentStatusFilter, selectedKelas, selectedJenjang, filterMode]);


  const renderSparkline = (dataPoints, strokeColor = "#16a34a", fillColor = "rgba(22, 163, 74, 0.1)") => {
    let pointsArray = dataPoints;
    if (!pointsArray || !Array.isArray(pointsArray) || pointsArray.length === 0) {
      pointsArray = [0, 0, 0, 0, 0];
    }
    pointsArray = pointsArray.map(item => typeof item === 'object' && item !== null ? item.value : Number(item) || 0);

    if (pointsArray.length === 1) {
      pointsArray = [pointsArray[0], pointsArray[0]];
    }
    const maxVal = Math.max(...pointsArray, 1);
    const minVal = Math.min(...pointsArray, 0);
    const range = maxVal - minVal || 1;
    
    const width = 120;
    const height = 30;
    
    const points = pointsArray.map((val, idx) => {
      const x = (idx / (pointsArray.length - 1)) * width;
      const y = height - ((val - minVal) / range) * (height - 6) - 3;
      return { x, y };
    });
    
    const pathData = points.map((p, idx) => `${idx === 0 ? 'M' : 'L'} ${p.x.toFixed(1)} ${p.y.toFixed(1)}`).join(' ');
    const areaData = `${pathData} L ${width} ${height} L 0 ${height} Z`;
    
    return (
      <svg viewBox={`0 0 ${width} ${height}`} className="w-full h-10 mt-1 overflow-visible select-none pointer-events-none opacity-85" aria-hidden="true">
        <path d={areaData} fill={fillColor} stroke="none" />
        <path d={pathData} fill="none" stroke={strokeColor} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
        {points.length > 0 && (
          <circle cx={points[points.length - 1].x} cy={points[points.length - 1].y} r="2" fill={strokeColor} />
        )}
      </svg>
    );
  };

  return (
    <div className={`min-h-screen transition-colors duration-300 ${printMode ? `print-mode-${printMode}` : ''}`}>
      
      {/* HEADER BANNER */}
      <header className="bg-bg-card border-b border-border-color text-text-main sticky top-0 z-50 transition-colors duration-150 h-14">
        <div className="w-full px-4 sm:px-6 h-full flex items-center justify-between">
          
          {/* Kiri: Logo + Nama Aplikasi */}
          <div className="flex items-center gap-3">
            <img src="/favicon.svg" alt="Logo Madin MU" className="w-7 h-7 select-none" />
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

              {/* PDF Export Dropdown */}
              <PDFExportMenu onExport={handleExportPDF} isExporting={isExporting} />

              {/* Reset Data Button */}
              <button
                onClick={handleResetData}
                className="p-1.5 rounded-md border border-border-color bg-bg-card text-text-muted hover:text-text-main hover:bg-bg-app transition-all duration-150 cursor-pointer h-8 w-8 flex items-center justify-center"
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
          <div className="flex items-center justify-between">
            <span className="text-xs text-text-muted font-medium">Ekspor Laporan:</span>
            <div className="flex gap-1.5 flex-wrap">
              <button 
                onClick={() => { handleExportPDF('all'); setMenuOpen(false); }}
                className="px-2.5 py-1.5 rounded-md border border-border-color bg-bg-app text-text-main text-[10px] font-semibold cursor-pointer flex items-center gap-1"
              >
                <Printer size={11} className="text-blue-600 dark:text-blue-400 shrink-0" />
                <span>Lengkap</span>
              </button>
              <button 
                onClick={() => { handleExportPDF('charts'); setMenuOpen(false); }}
                className="px-2.5 py-1.5 rounded-md border border-border-color bg-bg-app text-text-main text-[10px] font-semibold cursor-pointer flex items-center gap-1"
              >
                <BarChart3 size={11} className="text-primary shrink-0" />
                <span>Grafik</span>
              </button>
              <button 
                onClick={() => { handleExportPDF('data'); setMenuOpen(false); }}
                className="px-2.5 py-1.5 rounded-md border border-border-color bg-bg-app text-text-main text-[10px] font-semibold cursor-pointer flex items-center gap-1"
              >
                <FileText size={11} className="text-amber-600 dark:text-amber-400 shrink-0" />
                <span>Tabel</span>
              </button>
            </div>
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
        
        {/* PRINT-ONLY HEADER */}
        <div className="hidden print:block mb-6 border-b-2 border-border-color pb-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <img src="/favicon.svg" alt="Logo" className="w-9 h-9" />
              <div>
                <h1 className="text-base font-bold text-black dark:text-white uppercase tracking-wide">
                  Laporan Hasil Penilaian Ujian Baca Kitab
                </h1>
                <p className="text-[10px] text-text-muted font-medium">
                  Madrasah Diniyah Takmiliyah Miftahul Ulum Banyuputih Kidul Jatiroto Lumajang
                </p>
              </div>
            </div>
            <div className="text-right text-[10px] text-text-muted">
              <p className="font-semibold text-text-main">
                {filterMode === 'kelas' ? `Kelas: ${selectedKelas}` : `Jenjang: ${selectedJenjang}`}
              </p>
              <p>Tanggal Cetak: {new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}</p>
            </div>
          </div>
        </div>

        {/* ROW 1: UPLOAD EXCEL SECTION */}
        <section className="bg-bg-card border border-border-color rounded-lg p-5 flex flex-col md:flex-row gap-6 items-center transition-all duration-150 print:hidden">
          
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
                className="px-4.5 py-2.5 bg-primary hover:bg-primary-hover text-white rounded-lg text-xs font-semibold transition-all duration-150 flex items-center gap-2 cursor-pointer"
                onClick={downloadSampleExcel}
              >
                <Download size={14} />
                <span>Unduh Template Excel</span>
              </button>
              {fileName && (
                <div className="px-3.5 py-2 bg-primary/10 text-primary border border-primary/20 rounded-md text-xs font-medium flex items-center gap-1.5 animate-pulse">
                  <CheckCircle size={14} />
                  <span>Aktif: {fileName}</span>
                </div>
              )}
            </div>
          </div>

          <div 
            className={`w-full md:w-80 h-28 border-2 border-dashed rounded-lg flex flex-col items-center justify-center p-4 transition-all duration-150 text-center cursor-pointer ${
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
            <FileSpreadsheet className="text-text-muted mb-2" size={28} />
            <span className="text-xs font-semibold text-text-main">Tarik berkas atau klik di sini</span>
            <span className="text-[10px] text-text-muted mt-1">Hanya mendukung format file .xlsx / .xls</span>
          </div>

        </section>

        {/* ROW 2: FILTERS & CONTROL PANEL */}
        <section className="bg-bg-card border border-border-color rounded-lg p-5 flex flex-col gap-5 transition-all duration-150 print:hidden">
          
          <div className="flex flex-col lg:flex-row justify-between lg:items-center gap-6">
            
            {/* Filter Mode Selector */}
            <div className="flex flex-col gap-2">
              <span className="text-xs font-bold text-text-muted uppercase tracking-wider">Mode Analisis</span>
              <div className="inline-flex p-1 bg-bg-app rounded-lg gap-1 w-fit border border-border-color">
                <button
                  className={`px-4 py-2 rounded-md text-xs font-bold transition-all cursor-pointer ${
                    filterMode === 'kelas'
                      ? 'bg-bg-card text-primary shadow-sm'
                      : 'text-text-muted hover:text-text-main'
                  }`}
                  onClick={() => setFilterMode('kelas')}
                >
                  Per Kelas
                </button>
                <button
                  className={`px-4 py-2 rounded-md text-xs font-bold transition-all cursor-pointer ${
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
              <div className="flex items-center gap-4 bg-bg-app border border-border-color rounded-lg px-4 py-2.5 w-fit">
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
          <div className="bg-bg-app border border-border-color rounded-lg p-4 flex gap-3 text-xs text-text-muted transition-all duration-150">
            <Info className="text-primary shrink-0 mt-0.5" size={16} />
            <div className="flex-1 leading-relaxed">
              <strong>Aturan Bisnis Penilaian:</strong> Santri dengan KKM status <code className="px-1.5 py-0.5 rounded bg-bg-card border border-border-color text-primary font-mono font-bold">TIDAK HADIR</code> (Rata-rata 11.0) atau <code className="px-1.5 py-0.5 rounded bg-bg-card border border-border-color text-primary font-mono font-bold">TIDAK BACA</code> (Rata-rata 13.0) <strong>otomatis dikecualikan</strong> dari penghitungan Tuntas / Tidak Tuntas. Persentase di atas dihitung dengan basis pembagi aktif yang dipilih di sebelah kanan.
            </div>
          </div>

        </section>

        {/* ROW 3: STAT CARDS GRID (Optimized for TVs) */}
        <div id="visual-charts-section" className={`flex flex-col gap-8 p-4 rounded-xl ${isExporting ? 'bg-bg-card' : ''} ${printMode === 'data' ? 'print:hidden' : ''}`}>
          
          {/* PDF EXPORT FORMAL HEADER (Visible only during PDF export) */}
          {isExporting && (
            <div className="border-b-2 border-border-color pb-4 mb-2 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <img src="/favicon.svg" alt="Logo" className="w-10 h-10" />
                <div>
                  <h1 className="text-sm sm:text-base font-bold text-text-main uppercase tracking-wide">
                    Laporan Grafik Penilaian Ujian Baca Kitab
                  </h1>
                  <p className="text-[10px] text-text-muted font-medium">
                    Madrasah Diniyah Takmiliyah Miftahul Ulum Banyuputih Kidul Jatiroto Lumajang
                  </p>
                </div>
              </div>
              <div className="text-right text-[10px] text-text-muted">
                <p className="font-semibold text-text-main">
                  {filterMode === 'kelas' ? `Kelas: ${selectedKelas}` : `Jenjang: ${selectedJenjang}`}
                </p>
                <p>Tanggal: {new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}</p>
              </div>
            </div>
          )}

          <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 print:grid-cols-5 gap-5">
          
          {/* Card 1: Keseluruhan */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 relative overflow-hidden group flex flex-col justify-between">
            <div>
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
            </div>
            <div>
              <div className="-mx-5 mt-2">
                {renderSparkline(
                  categoryChartOptions?.keseluruhan?.series[0]?.data,
                  darkMode ? '#94a3b8' : '#64748b',
                  darkMode ? 'rgba(148, 163, 184, 0.1)' : 'rgba(100, 116, 139, 0.08)'
                )}
              </div>
              <p className="text-[10px] text-text-muted mt-2">Total terdaftar di {filterMode === 'kelas' ? `Kelas ${selectedKelas}` : `Jenjang ${selectedJenjang}`}</p>
            </div>
          </div>

          {/* Card 2: Tuntas */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 relative overflow-hidden group flex flex-col justify-between">
            <div>
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
            </div>
            <div>
              <div className="-mx-5 mt-2">
                {renderSparkline(
                  categoryChartOptions?.tuntasChart?.series[0]?.data,
                  darkMode ? '#22c55e' : '#16a34a',
                  darkMode ? 'rgba(34, 197, 94, 0.1)' : 'rgba(22, 163, 74, 0.08)'
                )}
              </div>
              <p className="text-[10px] text-text-muted mt-2">Mendapatkan predikat kelulusan TUNTAS</p>
            </div>
          </div>

          {/* Card 3: Tidak Tuntas */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 relative overflow-hidden group flex flex-col justify-between">
            <div>
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
            </div>
            <div>
              <div className="-mx-5 mt-2">
                {renderSparkline(
                  categoryChartOptions?.tidakTuntasChart?.series[0]?.data,
                  darkMode ? '#f87171' : '#dc2626',
                  darkMode ? 'rgba(248, 113, 113, 0.1)' : 'rgba(220, 38, 38, 0.08)'
                )}
              </div>
              <p className="text-[10px] text-text-muted mt-2">Mendapatkan predikat TIDAK TUNTAS</p>
            </div>
          </div>

          {/* Card 4: Tuntas > 70 */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 relative overflow-hidden group flex flex-col justify-between">
            <div>
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
            </div>
            <div>
              <div className="-mx-5 mt-2">
                {renderSparkline(
                  categoryChartOptions?.tuntasAtas70Chart?.series[0]?.data,
                  darkMode ? '#10b981' : '#059669',
                  darkMode ? 'rgba(16, 185, 129, 0.1)' : 'rgba(5, 150, 105, 0.08)'
                )}
              </div>
              <p className="text-[10px] text-text-muted mt-2">Tuntas dengan rata-rata nilai tinggi</p>
            </div>
          </div>

          {/* Card 5: Tuntas <= 70 */}
          <div className="bg-bg-card border border-border-color rounded-lg p-5 relative overflow-hidden group flex flex-col justify-between">
            <div>
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
            </div>
            <div>
              <div className="-mx-5 mt-2">
                {renderSparkline(
                  categoryChartOptions?.tuntasBawah70Chart?.series[0]?.data,
                  darkMode ? '#facc15' : '#d97706',
                  darkMode ? 'rgba(250, 204, 21, 0.1)' : 'rgba(217, 119, 6, 0.08)'
                )}
              </div>
              <p className="text-[10px] text-text-muted mt-2">Tuntas dengan nilai pas-pasan/ambang</p>
            </div>
          </div>

        </section>

        {/* DETAILED CATEGORY INFOGRAPHICS SECTION */}
        <section className="flex flex-col gap-4">
          <div className="flex items-center justify-between flex-wrap gap-2">
            <div>
              <h2 className="text-sm font-bold text-text-main flex items-center gap-2">
                <Layers className="text-primary" size={16} />
                <span>Rincian Analisis Kategori Ujian</span>
              </h2>
              <p className="text-[11px] text-text-muted mt-0.5">
                Visualisasi sebaran, performa, dan pencapaian target berdasarkan status dan kategori ketuntasan.
              </p>
            </div>
            {/* Toggle View/Hide Button */}
            <button
              onClick={() => setShowRincianCategory(!showRincianCategory)}
              className="p-2 bg-bg-card border border-border-color hover:border-primary hover:text-primary rounded-lg transition-all duration-150 cursor-pointer flex items-center justify-center"
              title={showRincianCategory ? "Sembunyikan Rincian" : "Tampilkan Rincian"}
              aria-label={showRincianCategory ? "Sembunyikan Rincian" : "Tampilkan Rincian"}
            >
              {showRincianCategory ? <EyeOff size={15} /> : <Eye size={15} />}
            </button>
          </div>

          <AnimatePresence initial={false}>
            {showRincianCategory && (
              <motion.div
                initial={{ opacity: 0, height: 0 }}
                animate={{ opacity: 1, height: 'auto' }}
                exit={{ opacity: 0, height: 0 }}
                transition={{ duration: 0.2, ease: 'easeInOut' }}
                className="overflow-hidden"
              >
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 print:grid-cols-5 gap-4 pt-1">
                  
                  {/* Infografis 1: Keseluruhan */}
                  <div className="bg-bg-card border border-border-color rounded-lg p-4 flex flex-col gap-3 transition-all duration-150">
                    <div className="flex items-center justify-between">
                      <div>
                        <h4 className="text-[10px] font-bold text-text-main uppercase tracking-wide">1. Status Kelulusan</h4>
                        <p className="text-[9px] text-text-muted mt-0.5">Proporsi predikat KKM</p>
                      </div>
                      <ChartExportMenu chartRef={keseluruhanChartRef} filename={`infografis-status-${filterMode}-${filterMode === 'kelas' ? selectedKelas : selectedJenjang}`} />
                    </div>
                    <div className="h-40 relative">
                      {stats.total === 0 ? (
                        <div className="absolute inset-0 flex items-center justify-center text-[10px] text-text-muted">Tidak ada data.</div>
                      ) : (
                        <ReactECharts ref={keseluruhanChartRef} option={categoryChartOptions?.keseluruhan} opts={{ renderer: 'svg' }} style={{ height: '100%', width: '100%' }} notMerge={true} />
                      )}
                    </div>
                  </div>

                  {/* Infografis 2: Tuntas */}
                  <div className="bg-bg-card border border-border-color rounded-lg p-4 flex flex-col gap-3 transition-all duration-150">
                    <div className="flex items-center justify-between">
                      <div>
                        <h4 className="text-[10px] font-bold text-text-main uppercase tracking-wide">2. Santri Tuntas</h4>
                        <p className="text-[9px] text-text-muted mt-0.5">Tren distribusi tuntas</p>
                      </div>
                      <ChartExportMenu chartRef={tuntasChartRef} filename={`infografis-tuntas-${filterMode}-${filterMode === 'kelas' ? selectedKelas : selectedJenjang}`} />
                    </div>
                    <div className="h-40 relative">
                      {stats.total === 0 ? (
                        <div className="absolute inset-0 flex items-center justify-center text-[10px] text-text-muted">Tidak ada data.</div>
                      ) : (
                        <ReactECharts ref={tuntasChartRef} option={categoryChartOptions?.tuntasChart} opts={{ renderer: 'svg' }} style={{ height: '100%', width: '100%' }} notMerge={true} />
                      )}
                    </div>
                  </div>

                  {/* Infografis 3: Tidak Tuntas */}
                  <div className="bg-bg-card border border-border-color rounded-lg p-4 flex flex-col gap-3 transition-all duration-150">
                    <div className="flex items-center justify-between">
                      <div>
                        <h4 className="text-[10px] font-bold text-text-main uppercase tracking-wide">3. Tidak Tuntas</h4>
                        <p className="text-[9px] text-text-muted mt-0.5">Tren tidak tuntas</p>
                      </div>
                      <ChartExportMenu chartRef={tidakTuntasChartRef} filename={`infografis-tidak-tuntas-${filterMode}-${filterMode === 'kelas' ? selectedKelas : selectedJenjang}`} />
                    </div>
                    <div className="h-40 relative">
                      {stats.total === 0 ? (
                        <div className="absolute inset-0 flex items-center justify-center text-[10px] text-text-muted">Tidak ada data.</div>
                      ) : (
                        <ReactECharts ref={tidakTuntasChartRef} option={categoryChartOptions?.tidakTuntasChart} opts={{ renderer: 'svg' }} style={{ height: '100%', width: '100%' }} notMerge={true} />
                      )}
                    </div>
                  </div>

                  {/* Infografis 4: Tuntas > 70 */}
                  <div className="bg-bg-card border border-border-color rounded-lg p-4 flex flex-col gap-3 transition-all duration-150">
                    <div className="flex items-center justify-between">
                      <div>
                        <h4 className="text-[10px] font-bold text-text-main uppercase tracking-wide">4. Tuntas &gt; 70</h4>
                        <p className="text-[9px] text-text-muted mt-0.5">Nilai rata-rata &gt; 70</p>
                      </div>
                      <ChartExportMenu chartRef={tuntasAtas70ChartRef} filename={`infografis-tuntas-atas-70-${filterMode}-${filterMode === 'kelas' ? selectedKelas : selectedJenjang}`} />
                    </div>
                    <div className="h-40 relative">
                      {stats.total === 0 ? (
                        <div className="absolute inset-0 flex items-center justify-center text-[10px] text-text-muted">Tidak ada data.</div>
                      ) : (
                        <ReactECharts ref={tuntasAtas70ChartRef} option={categoryChartOptions?.tuntasAtas70Chart} opts={{ renderer: 'svg' }} style={{ height: '100%', width: '100%' }} notMerge={true} />
                      )}
                    </div>
                  </div>

                  {/* Infografis 5: Tuntas <= 70 */}
                  <div className="bg-bg-card border border-border-color rounded-lg p-4 flex flex-col gap-3 transition-all duration-150">
                    <div className="flex items-center justify-between">
                      <div>
                        <h4 className="text-[10px] font-bold text-text-main uppercase tracking-wide">5. Tuntas &le; 70</h4>
                        <p className="text-[9px] text-text-muted mt-0.5">Nilai rata-rata &le; 70</p>
                      </div>
                      <ChartExportMenu chartRef={tuntasBawah70ChartRef} filename={`infografis-tuntas-bawah-70-${filterMode}-${filterMode === 'kelas' ? selectedKelas : selectedJenjang}`} />
                    </div>
                    <div className="h-40 relative">
                      {stats.total === 0 ? (
                        <div className="absolute inset-0 flex items-center justify-center text-[10px] text-text-muted">Tidak ada data.</div>
                      ) : (
                        <ReactECharts ref={tuntasBawah70ChartRef} option={categoryChartOptions?.tuntasBawah70Chart} opts={{ renderer: 'svg' }} style={{ height: '100%', width: '100%' }} notMerge={true} />
                      )}
                    </div>
                  </div>

                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </section>

        {/* ROW 4: CHARTS GRID */}
        <section className="grid grid-cols-1 lg:grid-cols-3 print:grid-cols-3 gap-6">
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
                    ? `Rata-rata Nilai Aspek Jenjang ${selectedJenjang}` 
                    : `Rata-rata Nilai Aspek Kelas ${selectedKelas}`}
                </h3>
                <p className="text-[11px] text-text-muted mt-0.5">
                  {filterMode === 'jenjang'
                    ? 'Rata-rata persentase capaian per aspek ujian jenjang'
                    : 'Rata-rata persentase capaian per aspek ujian kelas'}
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
      </div>
      {/* ROW 5: TABLES SECTION (Rekap Kelas + Detail Siswa) */}
      <section className={`grid grid-cols-1 lg:grid-cols-3 gap-6 ${printMode === 'charts' ? 'print:hidden' : ''}`}>
          
          {/* Left Column: Tabel Rekapitulasi Ringkas Kelas/Jenjang */}
          <div className={`bg-bg-card border border-border-color rounded-lg p-5 flex flex-col gap-4 transition-all duration-150 ${printMode === 'data' ? 'print:hidden' : ''}`}>
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
          <div className={`bg-bg-card border border-border-color rounded-lg p-5 flex flex-col gap-4 transition-all duration-150 ${printMode === 'data' ? 'lg:col-span-3' : 'lg:col-span-2'}`}>
            
            <div className="flex flex-col sm:flex-row justify-between sm:items-center gap-3">
              <div>
                <h3 className="text-sm font-bold text-text-main">
                  Rincian Nilai Santri ({filteredStudents.length} siswa)
                </h3>
                <p className="text-[11px] text-text-muted mt-0.5">Daftar nilai riil ujian baca kitab santri</p>
              </div>

              {/* Table search controls */}
              <div className="flex items-center gap-2 print:hidden">
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
                  {studentsToRender.length > 0 ? (
                    studentsToRender.map(student => (
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
              <div className="flex items-center justify-between mt-2 print:hidden">
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
