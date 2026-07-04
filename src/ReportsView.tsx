import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx-js-style';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { 
  FileText,
  Filter,
  CheckCircle2,
  FileSpreadsheet,
  FileWarning,
  Loader2,
  FileDown
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { format, parseISO } from 'date-fns';
import { id } from 'date-fns/locale';
import { kopSuratBase64 } from './kop-surat-b64';

type Status = 'Hadir' | 'Sakit' | 'Izin' | 'Alpa' | 'Dispen' | '';

interface Student {
  id: string;
  nisn: string;
  name: string;
  class: string;
  userId?: string;
}

interface AttendanceSession {
  id: string;
  date: string;
  className: string;
  meetingNumber: number;
  records: Record<string, Status>;
  userId?: string;
}

interface ReportsViewProps {
  classList: string[];
  students: Student[];
  attendanceSessions: AttendanceSession[];
  profileData: {
    namaGuruMapel: string;
    namaKepalaSekolah: string;
    nipGuruMapel: string;
    nipKepalaSekolah: string;
    semester: string;
    tahunPelajaran: string;
    mataPelajaran: string;
  };
}

export default function ReportsView({ classList, students, attendanceSessions, profileData }: ReportsViewProps) {
  const [selectedClass, setSelectedClass] = useState<string>('');
  const [reportType, setReportType] = useState<'daily' | 'monthly' | 'summary'>('summary');
  const [selectedMonth, setSelectedMonth] = useState<string>(format(new Date(), 'yyyy-MM'));
  const [selectedDailyDate, setSelectedDailyDate] = useState<string>(format(new Date(), 'yyyy-MM-dd'));
  const [showExportSuccess, setShowExportSuccess] = useState<'excel' | 'pdf' | 'none' | 'no_data'>('none');
  const [isExporting, setIsExporting] = useState(false);

  // Logic for computing report based on selectedClass, ReportType, and selectedMonth
  const classStudents = useMemo(() => {
    if (!selectedClass) return [];
    if (selectedClass === 'all') {
      return [...students].sort((a, b) => a.class.localeCompare(b.class, undefined, { numeric: true, sensitivity: 'base' }) || a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
    }
    return students.filter(s => s.class === selectedClass).sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
  }, [students, selectedClass]);

  const classSessions = useMemo(() => {
    if (!selectedClass) return [];
    if (selectedClass === 'all') {
      return [...attendanceSessions].sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    }
    return attendanceSessions.filter(s => s.className === selectedClass).sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
  }, [attendanceSessions, selectedClass]);

  const monthlySessions = useMemo(() => {
    return classSessions.filter(session => session.date.startsWith(selectedMonth));
  }, [classSessions, selectedMonth]);

  const availableDailyDates = useMemo(() => {
    return Array.from(new Set(classSessions.map(s => s.date))).sort().reverse();
  }, [classSessions]);

  React.useEffect(() => {
    if (reportType === 'daily' && availableDailyDates.length > 0 && !availableDailyDates.includes(selectedDailyDate)) {
      setSelectedDailyDate(availableDailyDates[0]);
    }
  }, [reportType, availableDailyDates, selectedDailyDate]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const autoFitColumns = (data: any[]) => {
    if (data.length === 0) return [];
    const keys = Object.keys(data[0]);
    return keys.map(key => {
      let max = key.length;
      data.forEach(row => {
        const val = row[key];
        if (val !== undefined && val !== null) {
          max = Math.max(max, val.toString().length);
        }
      });
      return { wch: max + 2 };
    });
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const applyStyles = (worksheet: any) => {
    if (!worksheet['!ref']) return;
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell_address = {c: C, r: R};
        const cell_ref = XLSX.utils.encode_cell(cell_address);
        if (!worksheet[cell_ref]) continue;

        const cell = worksheet[cell_ref];
        if (!cell.s) cell.s = {};
        
        const headerCell = worksheet[XLSX.utils.encode_cell({c: C, r: 0})];
        const headerText = headerCell ? headerCell.v : '';

        // Determine alignment
        let hAlign = 'center'; 
        if (headerText === 'Nama Lengkap Siswa' || headerText === 'Nama Siswa') {
          hAlign = 'left';
        }

        cell.s.alignment = { horizontal: hAlign, vertical: 'center' };
        
        // Header row styling
        if (R === 0) {
          cell.s.font = { bold: true };
          cell.s.alignment = { horizontal: 'center', vertical: 'center' };
        }
      }
    }
  };

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const appendSignatures = (worksheet: any, data: any[]) => {
    if (data.length === 0) return;
    const keys = Object.keys(data[0]);
    const numCols = keys.length;
    
    const leftCol = 1; 
    let rightCol = numCols > 3 ? numCols - 2 : numCols - 1;
    
    if (rightCol <= leftCol) rightCol = leftCol + 1; // Ensure they don't overlap if table is too small
    
    const createRow = () => new Array(numCols).fill('');
    
    const row1 = createRow();
    row1[leftCol] = 'Mengetahui,';
    row1[rightCol] = `Cililin, ${format(new Date(), 'dd MMMM yyyy', {locale: id})}`;
    
    const row2 = createRow();
    row2[leftCol] = 'Kepala Sekolah';
    row2[rightCol] = (profileData?.mataPelajaran && profileData.mataPelajaran.trim()) 
      ? `Guru Mata Pelajaran ${profileData.mataPelajaran.trim()}` 
      : 'Guru Mata Pelajaran';
    
    const row3 = createRow();
    const row4 = createRow();
    const row5 = createRow();
    
    const row6 = createRow();
    row6[leftCol] = profileData.namaKepalaSekolah || '(________________________)';
    row6[rightCol] = profileData.namaGuruMapel || '(________________________)';

    const row7 = createRow();
    row7[leftCol] = profileData.nipKepalaSekolah ? `NIP. ${profileData.nipKepalaSekolah}` : '';
    row7[rightCol] = profileData.nipGuruMapel ? `NIP. ${profileData.nipGuruMapel}` : '';

    XLSX.utils.sheet_add_aoa(worksheet, [[], [], row1, row2, row3, row4, row5, row6, row7], { origin: -1 });

    // Center alignment for signatures
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    for (let r = range.e.r - 8; r <= range.e.r; r++) {
      for (let c = 0; c <= range.e.c; c++) {
        const cell_ref = XLSX.utils.encode_cell({c, r});
        if (worksheet[cell_ref]) {
           if (!worksheet[cell_ref].s) worksheet[cell_ref].s = {};
           if (c === leftCol || c === rightCol) {
             worksheet[cell_ref].s.alignment = { horizontal: 'center', vertical: 'center' };
           }
        }
      }
    }
  };

  const getExportData = () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let data: any[] = [];
    let fileName = '';
    let sheetName = '';

    if (reportType === 'summary') {
      data = classStudents.map((student, i) => {
        const studentSessions = classSessions.filter(s => s.className === student.class);
        const totalSessions = studentSessions.length;
        let hadir = 0, sakit = 0, izin = 0, alpa = 0, dispen = 0;

        studentSessions.forEach(session => {
          const status = session.records[student.id];
          if (status === 'Hadir') hadir++;
          if (status === 'Sakit') sakit++;
          if (status === 'Izin') izin++;
          if (status === 'Alpa') alpa++;
          if (status === 'Dispen') dispen++;
        });

        const presentPercentage = totalSessions > 0 ? Math.round((hadir / totalSessions) * 100) : 0;

        const row: Record<string, string | number | boolean> = {
          'No': i + 1,
          'NISN': student.nisn,
          'Nama Lengkap Siswa': student.name
        };
        
        if (selectedClass === 'all') {
          row['Kelas'] = student.class;
        }

        row['Hadir'] = hadir;
        row['Sakit'] = sakit;
        row['Izin'] = izin;
        row['Alpa'] = alpa;
        row['Dispen'] = dispen;
        row['Presentase %'] = presentPercentage;

        return row;
      });
      fileName = `Rekap_Total_${selectedClass === 'all' ? 'Semua_Kelas' : selectedClass}`;
      sheetName = 'Rekap Total';
    } else if (reportType === 'monthly') {
      data = classStudents.map((student, i) => {
        const row: Record<string, string | number | boolean> = {
          'No': i + 1,
          'NISN': student.nisn,
          'Nama Lengkap Siswa': student.name
        };
        
        if (selectedClass === 'all') {
          row['Kelas'] = student.class;
        }

        let hadir = 0, sakit = 0, izin = 0, alpa = 0, dispen = 0;
        const studentMonthlySessions = monthlySessions.filter(s => s.className === student.class);

        studentMonthlySessions.forEach(session => {
          const status = session.records[student.id];
          if (status === 'Hadir') hadir++;
          if (status === 'Sakit') sakit++;
          if (status === 'Izin') izin++;
          if (status === 'Alpa') alpa++;
          if (status === 'Dispen') dispen++;
        });

        row['Hadir'] = hadir;
        row['Sakit'] = sakit;
        row['Izin'] = izin;
        row['Alpa'] = alpa;
        row['Dispen'] = dispen;
        
        const totalMonthlySessions = studentMonthlySessions.length;
        const presentPercentage = totalMonthlySessions > 0 ? Math.round((hadir / totalMonthlySessions) * 100) : 0;
        row['Presentase %'] = presentPercentage;

        return row;
      });
      fileName = `Rekap_Bulan_${selectedMonth}_${selectedClass === 'all' ? 'Semua_Kelas' : selectedClass}`;
      sheetName = `Bulan_${selectedMonth}`;
    } else {
      const dailySessionList = classSessions.filter(s => s.date === selectedDailyDate);
      const targetSession = dailySessionList[dailySessionList.length - 1];
      if (targetSession) {
        data = classStudents.map((student, i) => {
          const totalSessions = classSessions.length;
          let hadir = 0;
          classSessions.forEach(session => {
            if (session.records[student.id] === 'Hadir') hadir++;
          });
          const presentPercentage = totalSessions > 0 ? Math.round((hadir / totalSessions) * 100) : 0;

          return {
            'No': i + 1,
            'NISN': student.nisn,
            'Nama Lengkap Siswa': student.name,
            'Status': targetSession.records[student.id] || '-',
            'Presentase %': presentPercentage
          };
        });
        fileName = `Rekap_Harian_${selectedDailyDate}_${selectedClass}`;
        sheetName = `Harian_${selectedDailyDate}`;
      }
    }
    
    return { data, fileName, sheetName };
  };

  const exportToExcel = () => {
    if (!selectedClass) return;
    const { data, fileName, sheetName } = getExportData();
    if (!data || data.length === 0) {
      setShowExportSuccess('no_data');
      return;
    }

    setIsExporting(true);
    setTimeout(() => {
      try {
        const worksheet = XLSX.utils.json_to_sheet(data);
        applyStyles(worksheet);
        worksheet['!cols'] = autoFitColumns(data);
        appendSignatures(worksheet, data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        XLSX.writeFile(workbook, `${fileName}.xlsx`);
        setShowExportSuccess('excel');
      } catch (err) {
        console.error(err);
      } finally {
        setIsExporting(false);
      }
    }, 800);
  };

  const exportToPdf = () => {
    if (!selectedClass) return;
    const { data, fileName } = getExportData();
    if (!data || data.length === 0) {
      setShowExportSuccess('no_data');
      return;
    }

    setIsExporting(true);
    setTimeout(() => {
      try {
        const isLandscape = data[0] && Object.keys(data[0]).length > 10;
        const doc = new jsPDF({ orientation: isLandscape ? 'landscape' : 'portrait' });
        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        
        // PDF Header - Kop Surat Formal with Custom Image
        const imgWidth = pageWidth - 28;
        const imgHeight = imgWidth * (341 / 1450);
        
        try {
          doc.addImage(kopSuratBase64, 'PNG', 14, 10, imgWidth, imgHeight);
        } catch (e) {
          console.error("Failed to add custom header image", e);
        }
        
        const startY = 10 + imgHeight + 10; // margin top + image + spacing

        doc.setFontSize(14);
        doc.setFont("helvetica", "bold");
        doc.text('LAPORAN REKAPITULASI KEHADIRAN SISWA', pageWidth / 2, startY, { align: 'center' });
        
        doc.setFontSize(11);
        doc.setFont("helvetica", "normal");
        // Left Column Data
        doc.text(`Tahun Pelajaran`, 14, startY + 10);
        doc.text(`: ${profileData.tahunPelajaran || '-'}`, 50, startY + 10);
        doc.text(`Mata Pelajaran`, 14, startY + 16);
        doc.text(`: ${profileData.mataPelajaran || '-'}`, 50, startY + 16);

        // Right Column Data (Wider spacing to prevent overlapping)
        const labelX = pageWidth - 85;
        const valueX = pageWidth - 50;

        doc.text(`Kelas`, labelX, startY + 10);
        doc.text(`: ${selectedClass === 'all' ? 'Semua Kelas' : selectedClass}`, valueX, startY + 10);
        doc.text(`Semester`, labelX, startY + 16);
        doc.text(`: ${profileData.semester || '-'}`, valueX, startY + 16);
        
        // Periode in Left Column (Y: 22)
        if (reportType === 'monthly') {
          doc.text(`Periode Bulan`, 14, startY + 22);
          let displayMonth = selectedMonth;
          try {
            if (selectedMonth) {
              const [year, month] = selectedMonth.split('-');
              const dateObj = new Date(parseInt(year), parseInt(month) - 1, 1);
              displayMonth = format(dateObj, 'MMMM yyyy', { locale: id });
            }
          } catch (e) {
            console.error("Error formatting month:", e);
          }
          doc.text(`: ${displayMonth}`, 50, startY + 22);
        } else if (reportType === 'daily') {
          doc.text(`Periode Tanggal`, 14, startY + 22);
          let displayDate = selectedDailyDate || '-';
          try {
            if (selectedDailyDate) {
              displayDate = format(parseISO(selectedDailyDate), 'dd MMMM yyyy', { locale: id });
            }
          } catch (e) {
            console.error("Error formatting date:", e);
          }
          doc.text(`: ${displayDate}`, 50, startY + 22);
        } else if (reportType === 'summary') {
          doc.text(`Periode`, 14, startY + 22);
          let summaryDateText = 'Semua Sesi (Total)';
          if (classSessions.length > 0) {
            try {
              const startDate = format(parseISO(classSessions[0].date), 'dd MMMM yyyy', {locale: id});
              const endDate = format(parseISO(classSessions[classSessions.length - 1].date), 'dd MMMM yyyy', {locale: id});
              summaryDateText = `${startDate} - ${endDate}`;
            } catch (e) {
              console.error("Error formatting summary dates:", e);
            }
          }
          doc.text(`: ${summaryDateText}`, 50, startY + 22);
        }
        
        const headers = Object.keys(data[0]);
        const body = data.map(row => Object.values(row).map(val => val !== undefined && val !== null ? val.toString() : ''));

        const autoTableOptions = {
          startY: startY + 30,
          head: [headers],
          body: body,
          theme: 'grid' as const,
          headStyles: { fillColor: [230, 230, 230], textColor: 0, fontStyle: 'bold', halign: 'center', lineWidth: 0.2, lineColor: [0, 0, 0] },
          styles: { fontSize: 8, cellPadding: 2, lineWidth: 0.2, lineColor: [0, 0, 0], textColor: 0 },
          columnStyles: {
            /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
            2: { halign: 'left' } as any // Assuming index 2 is 'Nama Lengkap Siswa' always
          },
          /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
          didParseCell: function(cellData: any) {
            if (cellData.section === 'body' && cellData.column.index !== 2) {
              cellData.cell.styles.halign = 'center';
            }
          }
        };

        // Ensure we call autoTable safely across all bundler/Vite environments
        if (typeof autoTable === 'function') {
          /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
          autoTable(doc, autoTableOptions as any);
        /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
        } else if (typeof (doc as any).autoTable === 'function') {
          /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
          (doc as any).autoTable(autoTableOptions as any);
        } else {
          throw new Error("jsPDF AutoTable plugin is not loaded correctly. Please check integration.");
        }

        /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
        const finalY = (doc as any).lastAutoTable?.finalY || 40;
        
        let sigY = finalY + 20;
        if (sigY + 40 > pageHeight) {
           doc.addPage();
           sigY = 20;
        }

        doc.setFontSize(10);
        
        // Signatures
        let dateStr = '';
        try {
          dateStr = `Cililin, ${format(new Date(), 'dd MMMM yyyy', { locale: id })}`;
        } catch {
          // Safe manual fallback formatting if format or locale throws
          const d = new Date();
          const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
          dateStr = `Cililin, ${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
        }

        const headmasterName = profileData?.namaKepalaSekolah || '(________________________)';
        const headmasterNIP = profileData?.nipKepalaSekolah ? `NIP. ${profileData.nipKepalaSekolah}` : '';
        
        const teacherName = profileData?.namaGuruMapel || '(________________________)';
        const teacherNIP = profileData?.nipGuruMapel ? `NIP. ${profileData.nipGuruMapel}` : '';

        const teacherLabel = (profileData?.mataPelajaran && profileData.mataPelajaran.trim())
          ? `Guru Mata Pelajaran ${profileData.mataPelajaran.trim()}`
          : 'Guru Mata Pelajaran';

        doc.text('Mengetahui,', 20, sigY);
        doc.text('Kepala Sekolah', 20, sigY + 5);
        doc.text(dateStr, pageWidth - 80, sigY);
        doc.text(teacherLabel, pageWidth - 80, sigY + 5);

        doc.text(headmasterName, 20, sigY + 25);
        if (headmasterNIP) doc.text(headmasterNIP, 20, sigY + 30);
        
        doc.text(teacherName, pageWidth - 80, sigY + 25);
        if (teacherNIP) doc.text(teacherNIP, pageWidth - 80, sigY + 30);

        doc.save(`${fileName}.pdf`);
        setShowExportSuccess('pdf');
      } catch (error) {
        console.error('Error generating PDF:', error);
      } finally {
        setIsExporting(false);
      }
    }, 800);
  };

  return (
    <div className="p-4 sm:p-6 lg:p-8 space-y-6">
      <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
        <div>
          <h2 className="text-2xl sm:text-3xl font-black text-slate-800 tracking-tight">Laporan & Rekapitulasi</h2>
          <p className="text-slate-600 mt-1">Unduh hasil absensi siswa untuk pelaporan</p>
        </div>
        
        {selectedClass && (
          <div className="flex w-full sm:w-auto gap-3">
            <button 
              onClick={exportToExcel}
              disabled={isExporting}
              className="flex-1 sm:flex-none bg-emerald-600 hover:bg-emerald-700 text-white px-6 py-3 rounded-2xl font-bold flex items-center justify-center gap-2.5 shadow-md hover:shadow-lg transition-all active:scale-95 disabled:opacity-50 disabled:active:scale-100"
            >
              {isExporting ? <Loader2 className="w-5 h-5 animate-spin" /> : <FileSpreadsheet className="w-5 h-5" />} 
              Export Excel
            </button>
            <button 
              onClick={exportToPdf}
              disabled={isExporting}
              className="flex-1 sm:flex-none bg-rose-600 hover:bg-rose-700 text-white px-6 py-3 rounded-2xl font-bold flex items-center justify-center gap-2.5 shadow-md hover:shadow-lg transition-all active:scale-95 disabled:opacity-50 disabled:active:scale-100"
            >
              {isExporting ? <Loader2 className="w-5 h-5 animate-spin" /> : <FileDown className="w-5 h-5" />} 
              Export PDF
            </button>
          </div>
        )}
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="md:col-span-1 space-y-6">
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
            <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
              <Filter className="w-5 h-5 text-emerald-700" /> Filter Laporan
            </h3>
            
            <div className="space-y-4">
              <div>
                <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">Pilih Kelas</label>
                <select 
                  className="w-full p-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors"
                  value={selectedClass}
                  onChange={(e) => setSelectedClass(e.target.value)}
                >
                  <option value="">-- Pilih Kelas --</option>
                  <option value="all">Semua Kelas</option>
                  {classList.slice().sort((a,b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' })).map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>

              {selectedClass && (
                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">Jenis Laporan</label>
                  <select 
                    className="w-full p-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors"
                    value={reportType}
                    onChange={(e) => setReportType(e.target.value as 'daily' | 'monthly' | 'summary')}
                  >
                    <option value="summary">Rekap Total Keseluruhan</option>
                    <option value="monthly">Rekap Bulanan</option>
                    <option value="daily">Harian (Berdasarkan Tanggal)</option>
                  </select>
                </div>
              )}

              {selectedClass && reportType === 'daily' && (
                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">Pilih Tanggal</label>
                  {availableDailyDates.length > 0 ? (
                    <select 
                      className="w-full p-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors"
                      value={selectedDailyDate}
                      onChange={(e) => setSelectedDailyDate(e.target.value)}
                    >
                      {availableDailyDates.map(d => (
                        <option key={d} value={d}>{format(parseISO(d), 'dd MMMM yyyy', {locale: id})}</option>
                      ))}
                    </select>
                  ) : (
                    <div className="w-full p-3 bg-slate-100 text-slate-600 border-2 border-slate-300 rounded-xl text-sm italic">
                      Belum ada sesi absensi
                    </div>
                  )}
                </div>
              )}

              {selectedClass && reportType === 'monthly' && (
                <div>
                  <label className="block text-xs font-bold text-slate-600 uppercase tracking-wider mb-2">Pilih Bulan</label>
                  <input 
                    type="month" 
                    className="w-full p-3 bg-slate-50 border-2 border-slate-300 rounded-xl focus:ring-2 focus:ring-emerald-100 focus:border-emerald-600 transition-colors"
                    value={selectedMonth}
                    onChange={(e) => setSelectedMonth(e.target.value)}
                  />
                </div>
              )}
            </div>
          </div>
        </div>

        <div className="md:col-span-2">
          {!selectedClass ? (
            <div className="bg-slate-50/50 rounded-2xl border border-slate-100 h-full min-h-[300px] flex flex-col items-center justify-center p-8 text-center space-y-4">
              <div className="p-4 bg-white rounded-full shadow-sm">
                <FileText className="w-8 h-8 text-slate-300" />
              </div>
              <div>
                <h3 className="text-lg font-bold text-slate-700">Pilih Kelas Terlebih Dahulu</h3>
                <p className="text-slate-600 mt-1">Silakan pilih kelas melalui filter di samping untuk melihat preview laporan.</p>
              </div>
            </div>
          ) : (
            <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
              <div className="p-6 border-b border-slate-100 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                <div>
                  <h3 className="text-lg font-bold text-slate-800">
                    Preview Laporan {reportType === 'summary' ? 'Total' : reportType === 'monthly' ? 'Bulanan' : 'Harian'}
                  </h3>
                  <p className="text-slate-600 text-sm mt-1">
                    Kelas: <span className="font-bold text-slate-700">{selectedClass}</span> 
                    {reportType === 'monthly' && ` | Bulan: ${selectedMonth}`}
                    {reportType === 'daily' && selectedDailyDate && availableDailyDates.length > 0 && ` | Tanggal: ${format(parseISO(selectedDailyDate), 'dd MMMM yyyy', {locale: id})}`}
                    {reportType === 'summary' && classSessions.length > 0 && ` | Tanggal: ${format(parseISO(classSessions[0].date), 'dd MMM yyyy', {locale: id})} - ${format(parseISO(classSessions[classSessions.length - 1].date), 'dd MMM yyyy', {locale: id})}`}
                  </p>
                </div>
                <div className="flex bg-slate-50 rounded-lg p-1 border border-slate-100">
                  <span className="px-3 py-1 text-xs font-bold text-slate-600">Total: {classStudents.length} Siswa</span>
                </div>
              </div>

              <div className="overflow-x-auto max-h-[500px] scrollbar-thin">
                <table className="w-full min-w-[600px] text-sm text-left">
                  <thead className="bg-slate-50 sticky top-0 z-10 shadow-[0_1px_0_0_rgba(226,232,240,1)]">
                    <tr>
                      <th className="p-4 font-bold text-slate-600 border-b border-slate-200">No</th>
                      <th className="p-4 font-bold text-slate-600 border-b border-slate-200">Nama Lengkap Siswa</th>
                      {selectedClass === 'all' && (
                        <th className="p-4 font-bold text-slate-600 border-b border-slate-200">Kelas</th>
                      )}
                      {(reportType === 'summary' || (reportType === 'monthly' && selectedClass === 'all')) && (
                        <>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200"><span className="text-emerald-700">Hadir</span></th>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Sakit</th>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Izin</th>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200"><span className="text-rose-600">Alpa</span></th>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Dispen</th>
                          <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Presentase %</th>
                        </>
                      )}
                      {(reportType === 'monthly' && selectedClass !== 'all') && (
                        <>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Hadir</th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Sakit</th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Izin</th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200"><span className="text-rose-600">Alpa</span></th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Dispen</th>
                           <th className="p-4 font-bold text-slate-600 text-center border-b border-slate-200">Presentase %</th>
                        </>
                      )}
                      {reportType === 'daily' && (
                         <th className="p-4 font-bold text-slate-600 border-b border-slate-200">Status Terbaru</th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {classStudents.map((student, i) => {
                      let hadir = 0, sakit = 0, izin = 0, alpa = 0, dispen = 0;
                      let recentStatus = '-';

                      if (reportType === 'summary') {
                        const studentSessions = classSessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
                        studentSessions.forEach(session => {
                          const status = session.records[student.id];
                          if (status === 'Hadir') hadir++;
                          if (status === 'Sakit') sakit++;
                          if (status === 'Izin') izin++;
                          if (status === 'Alpa') alpa++;
                          if (status === 'Dispen') dispen++;
                        });
                      } else if (reportType === 'monthly') {
                        const studentMonthlySessions = monthlySessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
                        studentMonthlySessions.forEach(session => {
                          const status = session.records[student.id];
                          if (status === 'Hadir') hadir++;
                          if (status === 'Sakit') sakit++;
                          if (status === 'Izin') izin++;
                          if (status === 'Alpa') alpa++;
                          if (status === 'Dispen') dispen++;
                        });
                      } else {
                        const studentSessions = classSessions.filter(s => selectedClass === 'all' ? s.className === student.class : true);
                        const lastSession = studentSessions[studentSessions.length - 1];
                        if (lastSession) {
                          recentStatus = lastSession.records[student.id] || '-';
                        }
                      }

                      let totalSessions = 0;
                      if (reportType === 'summary') {
                        totalSessions = classSessions.filter(s => selectedClass === 'all' ? s.className === student.class : true).length;
                      } else if (reportType === 'monthly') {
                        totalSessions = monthlySessions.filter(s => selectedClass === 'all' ? s.className === student.class : true).length;
                      }
                      
                      const persentase = totalSessions > 0 ? Math.round((hadir / totalSessions) * 100) : 0;

                      return (
                        <tr key={student.id} className="border-b last:border-b-0 border-slate-100 hover:bg-slate-50/50 transition-colors">
                          <td className="p-4 text-slate-600 font-medium">{i + 1}</td>
                          <td className="p-4">
                            <div className="font-bold text-slate-800">{student.name}</div>
                            <div className="text-[10px] text-slate-500 font-mono mt-0.5">{student.nisn}</div>
                          </td>
                          {selectedClass === 'all' && (
                            <td className="p-4 text-slate-600 font-bold">{student.class}</td>
                          )}
                          {(reportType === 'summary' || (reportType === 'monthly' && selectedClass === 'all')) && (
                            <>
                              <td className="p-4 text-center font-bold text-emerald-700">{hadir}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{sakit}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{izin}</td>
                              <td className="p-4 text-center font-bold text-rose-500">{alpa}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{dispen}</td>
                              <td className="p-4 text-center">
                                <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${persentase >= 80 ? 'bg-emerald-100 text-lime-700' : persentase >= 50 ? 'bg-amber-100 text-amber-700' : 'bg-rose-100 text-rose-700'}`}>
                                  {persentase}%
                                </span>
                              </td>
                            </>
                          )}
                          {(reportType === 'monthly' && selectedClass !== 'all') && (
                             <>
                              <td className="p-4 text-center font-bold text-emerald-700">{hadir}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{sakit}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{izin}</td>
                              <td className="p-4 text-center font-bold text-rose-500">{alpa}</td>
                              <td className="p-4 text-center font-medium text-slate-600">{dispen}</td>
                              <td className="p-4 text-center">
                                <span className={`px-2.5 py-1 rounded-full text-xs font-bold ${persentase >= 80 ? 'bg-emerald-100 text-lime-700' : persentase >= 50 ? 'bg-amber-100 text-amber-700' : 'bg-rose-100 text-rose-700'}`}>
                                  {persentase}%
                                </span>
                              </td>
                             </>
                          )}
                          {reportType === 'daily' && (
                             <td className="p-4">
                               <span className={`px-2.5 py-1 rounded-full text-xs font-bold inline-block
                                 ${recentStatus === 'Hadir' ? 'bg-emerald-100 text-emerald-600' : 
                                   recentStatus === 'Sakit' || recentStatus === 'Izin' || recentStatus === 'Dispen' ? 'bg-amber-100 text-amber-600' : 
                                   recentStatus === 'Alpa' ? 'bg-rose-100 text-rose-600' : 'bg-slate-100 text-slate-500'}
                               `}>
                                 {recentStatus || '-'}
                               </span>
                             </td>
                          )}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      </div>

      <AnimatePresence>
        {isExporting && (
           <motion.div 
             initial={{ opacity: 0 }}
             animate={{ opacity: 1 }}
             exit={{ opacity: 0 }}
             className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 backdrop-blur-md p-4"
           >
             <motion.div
               initial={{ scale: 0.9, opacity: 0 }}
               animate={{ scale: 1, opacity: 1 }}
               exit={{ scale: 0.9, opacity: 0 }}
               className="bg-white rounded-[2.5rem] p-10 max-w-sm w-full shadow-[0_25px_60px_-15px_rgba(0,0,0,0.2)] text-center border border-slate-100"
             >
                <div className="relative mx-auto w-24 h-24 mb-6">
                  <motion.div 
                    animate={{ rotate: 360 }}
                    transition={{ duration: 2, repeat: Infinity, ease: "linear" }}
                    className="absolute inset-0 rounded-full border-4 border-emerald-100 border-t-emerald-500"
                  />
                  <div className="absolute inset-0 flex items-center justify-center">
                    <FileDown className="w-10 h-10 text-emerald-500" />
                  </div>
                </div>
                <h3 className="text-xl font-black text-slate-800 mb-2">Menyiapkan Laporan</h3>
                <p className="text-slate-600 text-sm leading-relaxed">
                  Mohon tunggu sebentar, kami sedang memproses data absensi menjadi format yang elegan untuk Anda.
                </p>
             </motion.div>
           </motion.div>
        )}

        {showExportSuccess !== 'none' && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/60 backdrop-blur-xl p-4"
          >
            <motion.div
              initial={{ y: 50, scale: 0.9, opacity: 0 }}
              animate={{ y: 0, scale: 1, opacity: 1 }}
              exit={{ y: 50, scale: 0.9, opacity: 0 }}
              transition={{ type: "spring", damping: 25, stiffness: 300 }}
              className="bg-white rounded-[3rem] p-10 max-w-sm w-full shadow-[0_40px_100px_-20px_rgba(0,0,0,0.3)] text-center relative overflow-hidden group"
            >
               {/* Decorative background elements */}
               <div className={`absolute top-0 left-0 w-full h-2 bg-gradient-to-r ${showExportSuccess === 'no_data' ? 'from-amber-400 to-rose-400' : 'from-emerald-400 to-sky-400'}`}></div>
               <div className={`absolute -top-12 -right-12 w-32 h-32 ${showExportSuccess === 'no_data' ? 'bg-amber-50' : 'bg-emerald-50'} rounded-full opacity-50 group-hover:scale-110 transition-transform duration-700`}></div>

               <div className="relative z-10">
                  <div className={`mx-auto w-24 h-24 ${showExportSuccess === 'no_data' ? 'bg-amber-100' : 'bg-emerald-100'} rounded-[2rem] flex items-center justify-center mb-8 rotate-3 group-hover:rotate-6 transition-transform`}>
                     {showExportSuccess === 'no_data' ? (
                       <FileWarning className="w-12 h-12 text-amber-600" />
                     ) : (
                       <CheckCircle2 className="w-12 h-12 text-emerald-600" />
                     )}
                  </div>
                  
                  <h3 className="text-2xl font-black text-slate-900 mb-3 tracking-tight">
                    {showExportSuccess === 'no_data' ? 'Data Kosong' : 'Ekspor Berhasil!'}
                  </h3>
                  <p className="text-slate-600 text-sm mb-10 leading-relaxed font-medium">
                    {showExportSuccess === 'no_data' 
                      ? 'Maaf, kami tidak menemukan data absensi untuk kriteria yang Anda pilih. Silakan pastikan kelas dan rentang waktu sudah benar.' 
                      : `Dokumen ${showExportSuccess === 'excel' ? 'Excel (.xlsx)' : 'PDF (.pdf)'} telah berhasil dibuat dan siap untuk Anda bagikan atau simpan.`}
                  </p>

                  <button 
                    onClick={() => setShowExportSuccess('none')}
                    className={`w-full ${showExportSuccess === 'no_data' ? 'bg-amber-600 hover:bg-amber-700' : 'bg-slate-900 hover:bg-slate-800'} text-white py-4 rounded-2xl font-black text-sm hover:scale-[1.02] active:scale-[0.98] transition-all shadow-xl ${showExportSuccess === 'no_data' ? 'shadow-amber-200' : 'shadow-slate-200'}`}
                  >
                    {showExportSuccess === 'no_data' ? 'Coba Lagi' : 'Selesai'}
                  </button>
               </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
