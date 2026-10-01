import React, { useState, useEffect, useMemo, useRef, memo } from 'react';
import { 
  ShieldCheck, 
  ArrowRight, 
  Eye, 
  EyeOff, 
  Receipt, 
  Calendar as CalendarIcon, 
  Users, 
  FileSpreadsheet, 
  Download, 
  LogOut, 
  AlertTriangle, 
  Filter, 
  DollarSign, 
  Clock, 
  CheckCircle2, 
  AlertCircle, 
  FileUp, 
  Upload, 
  Table as TableIcon, 
  Search, 
  MoveHorizontal, 
  ChevronLeft, 
  ChevronRight, 
  Trash2, 
  X, 
  Plus, 
  RefreshCw, 
  Loader2, 
  UploadCloud, 
  Check,
  TrendingUp,
  Crosshair,
  ListFilter,
  CalendarRange,
  Building2,
  Bell,
  AlertOctagon
} from 'lucide-react';
import { initializeApp, getApps, getApp } from 'firebase/app';
import { 
  getFirestore, 
  initializeFirestore,
  type Firestore,
  collection, 
  doc, 
  getDocs,
  setDoc, 
  updateDoc, 
  deleteDoc, 
  onSnapshot, 
  writeBatch 
} from 'firebase/firestore';
import { getAuth, signInAnonymously } from 'firebase/auth';
import * as XLSX from 'xlsx';
import ChartJS from 'chart.js/auto';

// Firebase Config
const firebaseConfig = {
  projectId: "compelling-analyst-sh7nb",
  appId: "1:889993914042:web:53c7ca50eff15cc3d08483",
  apiKey: "AIzaSyA03QSzv4cxwouIQgyrO0oZK8if5VB81_E",
  authDomain: "compelling-analyst-sh7nb.firebaseapp.com",
  firestoreDatabaseId: "ai-studio-extractordefactu-2ad37653-b923-4cfa-831d-cfa5998501de"
};

const app = getApps().length === 0 ? initializeApp(firebaseConfig) : getApp();

let db: Firestore;
try {
  db = initializeFirestore(app, {
    experimentalForceLongPolling: true,
  }, firebaseConfig.firestoreDatabaseId);
} catch {
  db = getFirestore(app, firebaseConfig.firestoreDatabaseId);
}
const auth = getAuth(app);

// Utilidad de fecha local inmune a desfases de huso horario UTC
function toLocalDateString(d: Date): string {
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// Validación estricta de fecha ISO (YYYY-MM-DD con año de 4 dígitos entre 1900 y 2099)
export function isValidDate(dateStr: any): boolean {
  if (!dateStr || typeof dateStr !== 'string') return false;
  const trimmed = dateStr.trim();
  const regex = /^(19|20)\d{2}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$/;
  return regex.test(trimmed);
}

export function sanitizeIsoDate(val: any): string {
  if (!val || typeof val !== 'string') return '';
  const trimmed = val.trim();
  if (isValidDate(trimmed)) return trimmed;
  // Si viene en formato DD/MM/AAAA o DD-MM-AAAA convertirlo
  const match = trimmed.match(/^(\d{1,2})[/-](\d{1,2})[/-]((?:19|20)\d{2})$/);
  if (match) {
    const reconstructed = `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
    if (isValidDate(reconstructed)) return reconstructed;
  }
  return '';
}

export function parseDateParts(dateStr: any): { year: string; month: string; day: string } {
  if (!dateStr || typeof dateStr !== 'string') return { year: '', month: '', day: '' };
  const trimmed = dateStr.trim();
  if (!isValidDate(trimmed)) return { year: '', month: '', day: '' };
  const parts = trimmed.split('-');
  return {
    year: parts[0] || '',
    month: parts[1] || '',
    day: parts[2] || ''
  };
}

export const DATE_DAYS = Array.from({ length: 31 }, (_, i) => String(i + 1).padStart(2, '0'));
export const DATE_MONTHS = [
  { val: '01', label: '01 - Ene' },
  { val: '02', label: '02 - Feb' },
  { val: '03', label: '03 - Mar' },
  { val: '04', label: '04 - Abr' },
  { val: '05', label: '05 - May' },
  { val: '06', label: '06 - Jun' },
  { val: '07', label: '07 - Jul' },
  { val: '08', label: '08 - Ago' },
  { val: '09', label: '09 - Sep' },
  { val: '10', label: '10 - Oct' },
  { val: '11', label: '11 - Nov' },
  { val: '12', label: '12 - Dic' }
];
export const DATE_YEARS = ['2025', '2026', '2027', '2028', '2029', '2030'];

interface CompoundDateSelectorProps {
  id: string;
  value: string;
  onSave: (dateStr: string) => void;
}

const CompoundDateSelector: React.FC<CompoundDateSelectorProps> = ({ id, value, onSave }) => {
  const parts = parseDateParts(value);
  const [selectedDay, setSelectedDay] = useState(parts.day);
  const [selectedMonth, setSelectedMonth] = useState(parts.month);
  const [selectedYear, setSelectedYear] = useState(parts.year);

  useEffect(() => {
    const p = parseDateParts(value);
    setSelectedDay(p.day);
    setSelectedMonth(p.month);
    setSelectedYear(p.year);
  }, [value]);

  const handleDayChange = (newDay: string) => {
    setSelectedDay(newDay);
    if (newDay && selectedMonth && selectedYear) {
      onSave(`${selectedYear}-${selectedMonth.padStart(2, '0')}-${newDay.padStart(2, '0')}`);
    }
  };

  const handleMonthChange = (newMonth: string) => {
    setSelectedMonth(newMonth);
    if (selectedDay && newMonth && selectedYear) {
      onSave(`${selectedYear}-${newMonth.padStart(2, '0')}-${selectedDay.padStart(2, '0')}`);
    }
  };

  const handleYearChange = (newYear: string) => {
    setSelectedYear(newYear);
    if (selectedDay && selectedMonth && newYear) {
      onSave(`${newYear}-${selectedMonth.padStart(2, '0')}-${selectedDay.padStart(2, '0')}`);
    }
  };

  const handleClear = () => {
    setSelectedDay('');
    setSelectedMonth('');
    setSelectedYear('');
    onSave('');
  };

  const hasFullDate = Boolean(selectedDay && selectedMonth && selectedYear);

  return (
    <div className="flex items-center gap-1">
      <select
        id={`date-day-${id}`}
        value={selectedDay}
        onChange={(e) => handleDayChange(e.target.value)}
        className="bg-slate-950 border border-slate-700/80 rounded px-1.5 py-1 text-xs text-slate-200 focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer font-mono hover:border-slate-500 transition-colors"
        title="Día de pago"
      >
        <option value="">Día</option>
        {DATE_DAYS.map((d) => (
          <option key={d} value={d}>{d}</option>
        ))}
      </select>

      <select
        id={`date-month-${id}`}
        value={selectedMonth}
        onChange={(e) => handleMonthChange(e.target.value)}
        className="bg-slate-950 border border-slate-700/80 rounded px-1.5 py-1 text-xs text-slate-200 focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer font-mono hover:border-slate-500 transition-colors"
        title="Mes de pago"
      >
        <option value="">Mes</option>
        {DATE_MONTHS.map((m) => (
          <option key={m.val} value={m.val}>{m.label}</option>
        ))}
      </select>

      <select
        id={`date-year-${id}`}
        value={selectedYear}
        onChange={(e) => handleYearChange(e.target.value)}
        className="bg-slate-950 border border-slate-700/80 rounded px-1.5 py-1 text-xs text-slate-200 focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer font-mono hover:border-slate-500 transition-colors"
        title="Año de pago"
      >
        <option value="">Año</option>
        {DATE_YEARS.map((y) => (
          <option key={y} value={y}>{y}</option>
        ))}
      </select>

      {hasFullDate && (
        <button
          type="button"
          onClick={handleClear}
          className="p-1 text-slate-500 hover:text-rose-400 rounded transition-colors cursor-pointer"
          title="Limpiar fecha"
        >
          <X className="w-3.5 h-3.5" />
        </button>
      )}
    </div>
  );
};

// Convertidor robusto para fechas de Excel / CSV
function parseExcelDate(val: any): string {
  if (!val) return '';
  if (typeof val === 'number') {
    // Número serial de fecha de Excel
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    return sanitizeIsoDate(toLocalDateString(date));
  }
  const str = String(val).trim();
  if (/^\d{5}$/.test(str)) {
    const num = parseInt(str, 10);
    const date = new Date(Math.round((num - 25569) * 86400 * 1000));
    return sanitizeIsoDate(toLocalDateString(date));
  }
  if (/^\d{1,2}[/-]\d{1,2}[/-]\d{4}$/.test(str)) {
    const parts = str.split(/[/-]/);
    return sanitizeIsoDate(`${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`);
  }
  return sanitizeIsoDate(str.split('T')[0]);
}

export const INVOICE_STATUS_OPTIONS = [
  'Generado',
  'Procedimiento parcial',
  'Procedimiento terminado',
  'Cancelada',
  'Pagada sin complemento',
  'Finalizado',
  'Complemento',
  'Problema'
] as const;

export type InvoiceStatus = typeof INVOICE_STATUS_OPTIONS[number];

export const DIA_PAGO_FIJO_OPTIONS = [
  'Mismo día exacto',
  'Lunes',
  'Martes',
  'Miércoles',
  'Jueves',
  'Viernes'
] as const;

export type DiaPagoFijo = typeof DIA_PAGO_FIJO_OPTIONS[number];

export interface InvoiceItem {
  id: string;
  numero_factura: string;
  empresa: string;
  rfc_cliente?: string;
  orden_de_compra: string;
  concepto: string;
  precio_unitario: number;
  monto_total: number;
  fecha_emision: string;
  fecha_probable_pago: string;
  estatus: InvoiceStatus;
  complemento: string; // Columna NOTAS
  documento_relacionado?: string;
  uuid?: string;
  folio_fiscal?: string;
  es_hueco_pendiente?: boolean;
  estatus_modificado_manualmente?: boolean;
  alerta_complemento_descartada?: boolean;
  createdAt?: string;
  updatedAt?: string;
}

export interface ClientItem {
  id: string;
  nombre: string;
  rfc: string;
  dias_credito: number;
  dia_pago_fijo: DiaPagoFijo;
  notas?: string;
}

const AUTH_KEY = 'cobranza_standalone_auth';

// Subcomponente memoizado para el campo de NOTAS
// Evita desincronizaciones y saltos de cursor al tipear
const NotesCell = memo(({ 
  initialValue, 
  onSave 
}: { 
  initialValue: string; 
  onSave: (val: string) => void;
}) => {
  const [val, setVal] = useState(initialValue);

  useEffect(() => {
    setVal(initialValue);
  }, [initialValue]);

  return (
    <input
      type="text"
      value={val}
      onChange={(e) => setVal(e.target.value)}
      onBlur={() => {
        if (val !== initialValue) {
          onSave(val);
        }
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') {
          e.currentTarget.blur();
        }
      }}
      placeholder="Escribir nota..."
      className="w-full bg-slate-950/70 border border-slate-800 focus:border-emerald-500 text-xs text-slate-200 rounded-lg px-2.5 py-1.5 focus:outline-none focus:ring-1 focus:ring-emerald-500 transition-colors placeholder:text-slate-600"
    />
  );
});

export default function App() {
  // Authentication state
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(() => {
    return sessionStorage.getItem(AUTH_KEY) === 'true' || localStorage.getItem(AUTH_KEY) === 'true';
  });
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loginError, setLoginError] = useState(false);

  // Data state
  const [invoices, setInvoices] = useState<InvoiceItem[]>([]);
  const [clients, setClients] = useState<ClientItem[]>([]);
  const [isFirebaseConnected, setIsFirebaseConnected] = useState<boolean>(false);

  // Filters & search
  const [periodFilter, setPeriodFilter] = useState<'todo' | 'ultimo_mes' | 'ultimos_dos_meses' | 'ano_actual' | 'personalizado'>('todo');
  const [customStartDate, setCustomStartDate] = useState('');
  const [customEndDate, setCustomEndDate] = useState('');
  const [searchQuery, setSearchQuery] = useState('');

  // Table direct advanced filters
  const [tableStatusFilter, setTableStatusFilter] = useState<string>('todos');
  const [tableDateStart, setTableDateStart] = useState<string>('');
  const [tableDateEnd, setTableDateEnd] = useState<string>('');
  const [tableEmpresaFilter, setTableEmpresaFilter] = useState<string>('todas');
  const [showEmpresaFilterModal, setShowEmpresaFilterModal] = useState<boolean>(false);
  const [empresaSearchQuery, setEmpresaSearchQuery] = useState<string>('');

  // Modals state
  const [showCalendarModal, setShowCalendarModal] = useState(false);
  const [showClientsModal, setShowClientsModal] = useState(false);
  const [showCsvModal, setShowCsvModal] = useState(false);
  const [showChartModal, setShowChartModal] = useState(false);
  const [showUrgentModal, setShowUrgentModal] = useState(false);
  const [showProblemaModal, setShowProblemaModal] = useState(false);

  // Mountain Chart state
  const [chartCenterDate, setChartCenterDate] = useState<Date>(new Date());
  const [chartDaysSpan, setChartDaysSpan] = useState<number>(30);
  const chartCanvasRef = useRef<HTMLCanvasElement | null>(null);
  const chartInstanceRef = useRef<ChartJS | null>(null);

  // Calendar month state
  const [calDate, setCalDate] = useState(new Date());

  // Dual scroll references
  const tableContainerRef = useRef<HTMLDivElement>(null);
  const topScrollRef = useRef<HTMLDivElement>(null);
  const [tableScrollWidth, setTableScrollWidth] = useState<number>(1750);
  const isSyncingRef = useRef(false);

  // New Client Form in Modal
  const [newClientName, setNewClientName] = useState('');
  const [newClientRfc, setNewClientRfc] = useState('');
  const [newClientDays, setNewClientDays] = useState(30);
  const [newClientDayFixed, setNewClientDayFixed] = useState<DiaPagoFijo>('Mismo día exacto');

  // CSV progress indicator
  const [csvLoading, setCsvLoading] = useState(false);
  const [csvStatusText, setCsvStatusText] = useState('');
  const csvInputRef = useRef<HTMLInputElement>(null);
  const xmlInputRef = useRef<HTMLInputElement>(null);

  // 1. Firebase Anonymous Auth
  useEffect(() => {
    signInAnonymously(auth)
      .then(() => setIsFirebaseConnected(true))
      .catch((err) => {
        console.warn('Anonymous auth note:', err?.message || err);
        setIsFirebaseConnected(true);
      });
  }, []);

  // 2. Firestore Listeners
  useEffect(() => {
    if (!isAuthenticated) return;

    const unsubscribeInvoices = onSnapshot(
      collection(db, 'invoices'),
      (snapshot) => {
        const items: InvoiceItem[] = [];
        snapshot.forEach((docSnap) => {
          const d = docSnap.data();
          const monto = typeof d.monto_total === 'number' ? d.monto_total : (parseFloat(d.monto_total) || 0);

          let estatus = d.estatus as InvoiceStatus;
          if (!INVOICE_STATUS_OPTIONS.includes(estatus)) {
            if (estatus === ('Completo' as any) || estatus === ('Pagada con complemento' as any)) {
              estatus = monto === 0 ? 'Complemento' : 'Finalizado';
            } else {
              estatus = 'Generado';
            }
          }

          items.push({
            id: docSnap.id,
            numero_factura: String(d.numero_factura || ''),
            empresa: String(d.empresa || ''),
            rfc_cliente: String(d.rfc_cliente || ''),
            orden_de_compra: String(d.orden_de_compra || ''),
            concepto: String(d.concepto || ''),
            precio_unitario: typeof d.precio_unitario === 'number' ? d.precio_unitario : (parseFloat(d.precio_unitario) || 0),
            monto_total: monto,
            fecha_emision: String(d.fecha_emision || ''),
            fecha_probable_pago: String(d.fecha_probable_pago || ''),
            estatus: estatus,
            complemento: String(d.complemento || ''),
            documento_relacionado: String(d.documento_relacionado || ''),
            uuid: String(d.uuid || d.folio_fiscal || ''),
            es_hueco_pendiente: Boolean(d.es_hueco_pendiente),
            estatus_modificado_manualmente: Boolean(d.estatus_modificado_manualmente),
            alerta_complemento_descartada: Boolean(d.alerta_complemento_descartada)
          });
        });

        // Transición Automática por Vencimiento de Fecha Probable de Pago
        const todayYMD = toLocalDateString(new Date());
        const overdueToUpdate = items.filter(inv => {
          return (
            inv.fecha_probable_pago &&
            isValidDate(inv.fecha_probable_pago) &&
            inv.fecha_probable_pago <= todayYMD &&
            inv.monto_total > 0 &&
            (inv.estatus === 'Generado' || inv.estatus === 'Procedimiento terminado') &&
            !inv.estatus_modificado_manualmente
          );
        });

        if (overdueToUpdate.length > 0) {
          const batch = writeBatch(db);
          overdueToUpdate.forEach(inv => {
            batch.update(doc(db, 'invoices', inv.id), {
              estatus: 'Pagada sin complemento',
              updatedAt: new Date().toISOString()
            });
            inv.estatus = 'Pagada sin complemento';
          });
          batch.commit().catch(err => console.error('Error auto-transición vencimiento:', err));
        }

        items.sort((a, b) => {
          const numA = parseInt(a.numero_factura.replace(/\D/g, ''), 10);
          const numB = parseInt(b.numero_factura.replace(/\D/g, ''), 10);
          if (!isNaN(numA) && !isNaN(numB) && numA !== numB) {
            return numB - numA;
          }
          return b.numero_factura.localeCompare(a.numero_factura);
        });

        setInvoices(items);
      },
      (error) => console.warn('Firestore Invoices listener note:', error?.message || error)
    );

    const unsubscribeClients = onSnapshot(
      collection(db, 'clientes'),
      (snapshot) => {
        const items: ClientItem[] = [];
        snapshot.forEach((docSnap) => {
          const d = docSnap.data();
          items.push({
            id: docSnap.id,
            nombre: String(d.nombre || ''),
            rfc: String(d.rfc || ''),
            dias_credito: typeof d.dias_credito === 'number' ? d.dias_credito : (parseInt(d.dias_credito, 10) || 30),
            dia_pago_fijo: (d.dia_pago_fijo as DiaPagoFijo) || 'Mismo día exacto',
            notas: String(d.notas || '')
          });
        });
        items.sort((a, b) => a.nombre.localeCompare(b.nombre));
        setClients(items);
      },
      (error) => console.warn('Firestore Clientes listener note:', error?.message || error)
    );

    return () => {
      unsubscribeInvoices();
      unsubscribeClients();
    };
  }, [isAuthenticated]);

  // Dual scroll sync and dynamic scrollWidth measurement
  useEffect(() => {
    const tableEl = tableContainerRef.current;
    const topScrollEl = topScrollRef.current;
    if (!tableEl || !topScrollEl) return;

    if (tableEl.scrollWidth > 0) {
      setTableScrollWidth(tableEl.scrollWidth);
    }

    const onTableScroll = () => {
      if (isSyncingRef.current) return;
      isSyncingRef.current = true;
      topScrollEl.scrollLeft = tableEl.scrollLeft;
      requestAnimationFrame(() => { isSyncingRef.current = false; });
    };

    const onTopScroll = () => {
      if (isSyncingRef.current) return;
      isSyncingRef.current = true;
      tableEl.scrollLeft = topScrollEl.scrollLeft;
      requestAnimationFrame(() => { isSyncingRef.current = false; });
    };

    tableEl.addEventListener('scroll', onTableScroll);
    topScrollEl.addEventListener('scroll', onTopScroll);

    return () => {
      tableEl.removeEventListener('scroll', onTableScroll);
      topScrollEl.removeEventListener('scroll', onTopScroll);
    };
  }, [invoices.length]);

  // Currency Formatter
  const formatCurrency = (val: number) => {
    return new Intl.NumberFormat('es-MX', {
      style: 'currency',
      currency: 'MXN'
    }).format(val);
  };

  // Login handler
  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (password === 'Altair16') {
      sessionStorage.setItem(AUTH_KEY, 'true');
      localStorage.setItem(AUTH_KEY, 'true');
      setIsAuthenticated(true);
      setLoginError(false);
    } else {
      setLoginError(true);
    }
  };

  const handleLogout = () => {
    sessionStorage.removeItem(AUTH_KEY);
    localStorage.removeItem(AUTH_KEY);
    setIsAuthenticated(false);
    setPassword('');
  };

  // Filter invoices by global metrics period (con fecha local)
  const periodFilteredInvoices = useMemo(() => {
    if (periodFilter === 'todo') return invoices;

    const now = new Date();
    const todayYMD = toLocalDateString(now);

    if (periodFilter === 'ultimo_mes') {
      const d = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
      const startYMD = toLocalDateString(d);
      return invoices.filter((i) => (i.fecha_emision || todayYMD) >= startYMD);
    }
    if (periodFilter === 'ultimos_dos_meses') {
      const d = new Date(now.getFullYear(), now.getMonth() - 2, now.getDate());
      const startYMD = toLocalDateString(d);
      return invoices.filter((i) => (i.fecha_emision || todayYMD) >= startYMD);
    }
    if (periodFilter === 'ano_actual') {
      const startYMD = `${now.getFullYear()}-01-01`;
      return invoices.filter((i) => (i.fecha_emision || todayYMD) >= startYMD);
    }
    if (periodFilter === 'personalizado') {
      return invoices.filter((i) => {
        const date = i.fecha_emision || todayYMD;
        if (customStartDate && date < customStartDate) return false;
        if (customEndDate && date > customEndDate) return false;
        return true;
      });
    }
    return invoices;
  }, [invoices, periodFilter, customStartDate, customEndDate]);

  // Catálogo y estadísticas de empresas receptoras
  const empresaStats = useMemo(() => {
    const map = new Map<string, {
      nombre: string;
      rfc: string;
      totalFacturas: number;
      montoTotal: number;
      pendiente: number;
      cobrado: number;
    }>();

    invoices.forEach((inv) => {
      const rawName = (inv.empresa || '').trim();
      if (!rawName) return;
      const key = rawName.toLowerCase();
      const existing = map.get(key);
      const monto = inv.monto_total || 0;
      const isCobrado = inv.estatus === 'Finalizado' || inv.estatus === 'Pagada sin complemento';
      const isExcluded = inv.estatus === 'Cancelada' || inv.estatus === 'Complemento' || monto <= 0;

      if (!existing) {
        const cli = clients.find(c => c.nombre.trim().toLowerCase() === key);
        map.set(key, {
          nombre: rawName,
          rfc: inv.rfc_cliente || cli?.rfc || '',
          totalFacturas: 1,
          montoTotal: isExcluded ? 0 : monto,
          pendiente: isExcluded || isCobrado ? 0 : monto,
          cobrado: isCobrado ? monto : 0
        });
      } else {
        existing.totalFacturas += 1;
        if (!isExcluded) {
          existing.montoTotal += monto;
          if (isCobrado) {
            existing.cobrado += monto;
          } else {
            existing.pendiente += monto;
          }
        }
        if (!existing.rfc && inv.rfc_cliente) {
          existing.rfc = inv.rfc_cliente;
        }
      }
    });

    return Array.from(map.values()).sort((a, b) => a.nombre.localeCompare(b.nombre));
  }, [invoices, clients]);

  const filteredEmpresaStats = useMemo(() => {
    const q = empresaSearchQuery.trim().toLowerCase();
    if (!q) return empresaStats;
    return empresaStats.filter(e => 
      e.nombre.toLowerCase().includes(q) || 
      (e.rfc && e.rfc.toLowerCase().includes(q))
    );
  }, [empresaStats, empresaSearchQuery]);

  // Direct table display filtering (Buscador + Estatus + Rango de fecha pago + Empresa)
  const displayInvoices = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    const targetEmpresa = tableEmpresaFilter.trim().toLowerCase();

    return periodFilteredInvoices.filter((inv) => {
      // 1. Text search
      if (q) {
        const matches = (
          inv.numero_factura.toLowerCase().includes(q) ||
          inv.empresa.toLowerCase().includes(q) ||
          inv.orden_de_compra.toLowerCase().includes(q) ||
          inv.concepto.toLowerCase().includes(q) ||
          inv.complemento.toLowerCase().includes(q)
        );
        if (!matches) return false;
      }

      // 2. Status filter
      if (tableStatusFilter !== 'todos') {
        if (tableStatusFilter === 'Complemento') {
          if (inv.estatus !== 'Complemento' && inv.monto_total > 0) return false;
        } else {
          if (inv.estatus !== tableStatusFilter) return false;
        }
      }

      // 3. Date range for payment
      if (tableDateStart && inv.fecha_probable_pago && inv.fecha_probable_pago < tableDateStart) return false;
      if (tableDateEnd && inv.fecha_probable_pago && inv.fecha_probable_pago > tableDateEnd) return false;

      // 4. Empresa filter
      if (targetEmpresa !== 'todas') {
        if ((inv.empresa || '').trim().toLowerCase() !== targetEmpresa) {
          return false;
        }
      }

      return true;
    });
  }, [periodFilteredInvoices, searchQuery, tableStatusFilter, tableDateStart, tableDateEnd, tableEmpresaFilter]);

  // Metrics calculation
  const metrics = useMemo(() => {
    let totalFacturado = 0;
    let pendienteCobro = 0;
    let totalCobrado = 0;
    let problemaMonto = 0;
    let countFacturas = 0;
    let countPendiente = 0;
    let countCobrado = 0;
    let countProblema = 0;

    periodFilteredInvoices.forEach((inv) => {
      if (inv.estatus === 'Cancelada' || inv.estatus === 'Complemento' || inv.monto_total <= 0) {
        return;
      }

      totalFacturado += inv.monto_total;
      countFacturas++;

      if (inv.estatus === 'Finalizado' || inv.estatus === 'Pagada sin complemento') {
        totalCobrado += inv.monto_total;
        countCobrado++;
      } else if (inv.estatus === 'Procedimiento parcial' || inv.estatus === 'Problema') {
        problemaMonto += inv.monto_total;
        countProblema++;
      } else {
        pendienteCobro += inv.monto_total;
        countPendiente++;
      }
    });

    return {
      totalFacturado,
      pendienteCobro,
      totalCobrado,
      problemaMonto,
      countFacturas,
      countPendiente,
      countCobrado,
      countProblema
    };
  }, [periodFilteredInvoices]);

  // Facturas de Atención Urgente (Complementos Pendientes)
  const urgentInvoices = useMemo(() => {
    const todayYMD = toLocalDateString(new Date());
    return invoices.filter(inv => {
      if (inv.estatus === 'Complemento' || inv.estatus === 'Cancelada' || inv.monto_total <= 0) {
        return false;
      }
      if (inv.alerta_complemento_descartada) {
        return false;
      }
      if (inv.estatus === 'Finalizado') {
        return false;
      }
      const isPagadaSinComp = inv.estatus === 'Pagada sin complemento';
      const isVencidaSinComp = (
        inv.fecha_probable_pago &&
        isValidDate(inv.fecha_probable_pago) &&
        inv.fecha_probable_pago <= todayYMD
      );
      return isPagadaSinComp || isVencidaSinComp;
    });
  }, [invoices]);

  // Facturas con Problema
  const problemaInvoices = useMemo(() => {
    return invoices.filter(inv => {
      if (inv.estatus === 'Complemento' || inv.estatus === 'Cancelada' || inv.monto_total <= 0) {
        return false;
      }
      return inv.estatus === 'Problema';
    });
  }, [invoices]);

  const dismissUrgentAlert = async (id: string) => {
    try {
      const docRef = doc(db, 'invoices', id);
      await updateDoc(docRef, {
        alerta_complemento_descartada: true,
        updatedAt: new Date().toISOString()
      });
    } catch (err) {
      console.error('Error al descartar alerta de complemento:', err);
    }
  };

  // Sequence Gap Detection
  const missingFolios = useMemo(() => {
    const numericFolios = invoices
      .map((i) => parseInt(i.numero_factura.replace(/\D/g, ''), 10))
      .filter((n) => !isNaN(n))
      .sort((a, b) => a - b);

    if (numericFolios.length < 2) return [];

    const missing: number[] = [];
    for (let i = 0; i < numericFolios.length - 1; i++) {
      const curr = numericFolios[i];
      const next = numericFolios[i + 1];
      if (next - curr > 1 && next - curr <= 15) {
        for (let m = curr + 1; m < next; m++) {
          missing.push(m);
        }
      }
    }
    return missing;
  }, [invoices]);

  // Mountain Chart Render Logic (Chart.js)
  const mountainChartData = useMemo(() => {
    const half = Math.floor(chartDaysSpan / 2);
    const labels: string[] = [];
    const keysYMD: string[] = [];
    const cobradoData: number[] = [];
    const pendienteData: number[] = [];

    let sumCobrado = 0;
    let sumPendiente = 0;

    const todayYMD = toLocalDateString(new Date());

    for (let i = -half; i <= half; i++) {
      const d = new Date(chartCenterDate);
      d.setDate(d.getDate() + i);
      const ymd = toLocalDateString(d);
      keysYMD.push(ymd);

      const dayName = d.toLocaleDateString('es-MX', { weekday: 'short', day: 'numeric', month: 'short' });
      labels.push(ymd === todayYMD ? `★ HOY (${d.getDate()} ${d.toLocaleDateString('es-MX', { month: 'short' })})` : dayName);

      let dayCobrado = 0;
      let dayPendiente = 0;

      invoices.forEach((inv) => {
        if (!inv.fecha_probable_pago || inv.estatus === 'Cancelada' || inv.estatus === 'Complemento' || inv.monto_total <= 0) {
          return;
        }
        if (inv.fecha_probable_pago === ymd) {
          if (inv.estatus === 'Finalizado' || inv.estatus === 'Pagada sin complemento') {
            dayCobrado += inv.monto_total;
          } else {
            dayPendiente += inv.monto_total;
          }
        }
      });

      cobradoData.push(dayCobrado);
      pendienteData.push(dayPendiente);
      sumCobrado += dayCobrado;
      sumPendiente += dayPendiente;
    }

    return {
      labels,
      cobradoData,
      pendienteData,
      sumCobrado,
      sumPendiente,
      sumTotal: sumCobrado + sumPendiente
    };
  }, [invoices, chartCenterDate, chartDaysSpan]);

  // Effect to instantiate/update Chart.js in the modal
  useEffect(() => {
    if (!showChartModal) return;

    const timer = setTimeout(() => {
      if (!chartCanvasRef.current) return;

      const existingChart = ChartJS.getChart(chartCanvasRef.current);
      if (existingChart) {
        existingChart.destroy();
      }

      if (chartInstanceRef.current) {
        chartInstanceRef.current.destroy();
        chartInstanceRef.current = null;
      }

      const ctx = chartCanvasRef.current.getContext('2d');
      if (!ctx) return;

      const gradEmerald = ctx.createLinearGradient(0, 0, 0, 360);
      gradEmerald.addColorStop(0, 'rgba(16, 185, 129, 0.45)');
      gradEmerald.addColorStop(1, 'rgba(16, 185, 129, 0.02)');

      const gradAmber = ctx.createLinearGradient(0, 0, 0, 360);
      gradAmber.addColorStop(0, 'rgba(251, 191, 36, 0.45)');
      gradAmber.addColorStop(1, 'rgba(251, 191, 36, 0.02)');

      chartInstanceRef.current = new ChartJS(ctx, {
        type: 'line',
        data: {
          labels: mountainChartData.labels,
          datasets: [
            {
              label: 'Total Cobrado (Finalizado)',
              data: mountainChartData.cobradoData,
              borderColor: '#10b981',
              backgroundColor: gradEmerald,
              borderWidth: 2.5,
              tension: 0.35,
              fill: true,
              pointBackgroundColor: '#10b981',
              pointBorderColor: '#064e3b',
              pointRadius: 3.5,
              pointHoverRadius: 6
            },
            {
              label: 'Pendiente de Cobro Proyectado',
              data: mountainChartData.pendienteData,
              borderColor: '#f59e0b',
              backgroundColor: gradAmber,
              borderWidth: 2.5,
              tension: 0.35,
              fill: true,
              pointBackgroundColor: '#f59e0b',
              pointBorderColor: '#78350f',
              pointRadius: 3.5,
              pointHoverRadius: 6
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: {
            mode: 'index',
            intersect: false
          },
          plugins: {
            legend: {
              labels: {
                color: '#cbd5e1',
                font: { size: 11, weight: 'bold', family: 'Inter' }
              }
            },
            tooltip: {
              backgroundColor: 'rgba(15, 23, 42, 0.95)',
              titleColor: '#f8fafc',
              bodyColor: '#cbd5e1',
              borderColor: '#334155',
              borderWidth: 1,
              padding: 10,
              callbacks: {
                label: function(context: any) {
                  return ` ${context.dataset.label}: ${formatCurrency(context.parsed.y)}`;
                }
              }
            }
          },
          scales: {
            x: {
              grid: { color: 'rgba(51, 65, 85, 0.3)' },
              ticks: {
                color: (val: any) => {
                  const label = mountainChartData.labels[val.index] || '';
                  return label.includes('★ HOY') ? '#10b981' : '#94a3b8';
                },
                font: (val: any) => {
                  const label = mountainChartData.labels[val.index] || '';
                  return {
                    weight: label.includes('★ HOY') ? 'bold' : 'normal',
                    size: 10
                  };
                }
              }
            },
            y: {
              grid: { color: 'rgba(51, 65, 85, 0.3)' },
              ticks: {
                color: '#94a3b8',
                callback: (value: any) => formatCurrency(Number(value))
              }
            }
          }
        }
      });
    }, 50);

    return () => {
      clearTimeout(timer);
      if (chartCanvasRef.current) {
        const existing = ChartJS.getChart(chartCanvasRef.current);
        if (existing) {
          existing.destroy();
        }
      }
      if (chartInstanceRef.current) {
        chartInstanceRef.current.destroy();
        chartInstanceRef.current = null;
      }
    };
  }, [showChartModal, mountainChartData]);

  // Update field in Firestore
  const updateInvoiceField = async (id: string, field: keyof InvoiceItem, value: any) => {
    try {
      const docRef = doc(db, 'invoices', id);
      let finalValue = value;
      if (field === 'fecha_probable_pago') {
        finalValue = sanitizeIsoDate(value);
      }
      const updates: any = {
        [field]: finalValue,
        updatedAt: new Date().toISOString()
      };
      if (field === 'estatus') {
        updates.estatus_modificado_manualmente = true;
      }
      await updateDoc(docRef, updates);
    } catch (err) {
      console.error('Error updating invoice:', err);
    }
  };

  // Delete invoice in Firestore
  const deleteInvoice = async (id: string, folio: string) => {
    const ok = window.confirm(`¿Estás seguro de que deseas eliminar permanentemente la factura ${folio || 'seleccionada'} de Cloud Firestore?`);
    if (!ok) return;

    try {
      await deleteDoc(doc(db, 'invoices', id));
    } catch (err) {
      console.error('Error deleting invoice:', err);
      alert('No se pudo eliminar la factura de Cloud Firestore.');
    }
  };

  // Status style helper
  const getStatusBadgeStyle = (status: InvoiceStatus) => {
    switch (status) {
      case 'Generado':
        return 'bg-slate-700/80 text-slate-200 border-slate-600 hover:bg-slate-700';
      case 'Procedimiento parcial':
        return 'bg-amber-500/90 text-slate-950 font-bold border-amber-400 hover:bg-amber-400';
      case 'Problema':
        return 'bg-red-600/95 text-white font-extrabold border-2 border-orange-400 hover:bg-red-500 shadow-sm';
      case 'Procedimiento terminado':
        return 'bg-blue-600/90 text-white font-bold border-blue-400 hover:bg-blue-500';
      case 'Cancelada':
        return 'bg-rose-700/90 text-white font-bold border-rose-500 hover:bg-rose-600';
      case 'Pagada sin complemento':
        return 'bg-emerald-400/90 text-slate-950 font-bold border-emerald-300 hover:bg-emerald-300';
      case 'Finalizado':
        return 'bg-emerald-600 text-white font-bold border-emerald-400 hover:bg-emerald-500';
      case 'Complemento':
        return 'bg-purple-600 text-white font-bold border-purple-400 hover:bg-purple-500';
      default:
        return 'bg-slate-700 text-slate-200 border-slate-600';
    }
  };

  // XML CFDI Upload Handler con detección de folios existentes y actualización de huecos
  const handleXmlUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const xmlText = event.target?.result as string;
        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(xmlText, 'text/xml');

        const comprobante = xmlDoc.getElementsByTagName('cfdi:Comprobante')[0] || 
                            xmlDoc.getElementsByTagName('Comprobante')[0] ||
                            xmlDoc.documentElement;
        if (!comprobante) {
          alert('El archivo no parece ser un comprobante CFDI válido del SAT.');
          return;
        }

        const allElements = Array.from(xmlDoc.getElementsByTagName('*'));

        const rawFolio = (comprobante.getAttribute('Folio') || comprobante.getAttribute('folio') || file.name.replace(/\D/g, '') || 'S/N').trim();
        const rawSerie = (comprobante.getAttribute('Serie') || comprobante.getAttribute('serie') || '').trim();
        const rawFecha = comprobante.getAttribute('Fecha') || comprobante.getAttribute('fecha') || '';
        const fecha = rawFecha ? rawFecha.split('T')[0] : toLocalDateString(new Date());
        const totalStr = comprobante.getAttribute('Total') || comprobante.getAttribute('total') || '0';
        const montoTotal = parseFloat(totalStr) || 0;
        const tipoComprobante = (comprobante.getAttribute('TipoDeComprobante') || comprobante.getAttribute('tipoDeComprobante') || 'I').toUpperCase();

        // Extraer UUID fiscal del propio comprobante (TimbreFiscalDigital)
        const timbreNode = allElements.find(el => (el.localName || el.nodeName || '').toLowerCase().includes('timbrefiscaldigital'));
        const currentUuid = timbreNode ? (timbreNode.getAttribute('UUID') || timbreNode.getAttribute('uuid') || '').trim() : '';

        const receptor = xmlDoc.getElementsByTagName('cfdi:Receptor')[0] || 
                         xmlDoc.getElementsByTagName('Receptor')[0] ||
                         allElements.find(el => (el.localName || el.nodeName || '').toLowerCase().endsWith('receptor'));
        const nombreReceptor = receptor ? (receptor.getAttribute('Nombre') || receptor.getAttribute('nombre') || 'Cliente SAT') : 'Cliente SAT';
        const rfcReceptor = receptor ? (receptor.getAttribute('Rfc') || receptor.getAttribute('rfc') || '') : '';

        const conceptoElem = xmlDoc.getElementsByTagName('cfdi:Concepto')[0] || 
                             xmlDoc.getElementsByTagName('Concepto')[0] ||
                             allElements.find(el => (el.localName || el.nodeName || '').toLowerCase().endsWith('concepto'));
        const descripcion = conceptoElem ? (conceptoElem.getAttribute('Descripcion') || conceptoElem.getAttribute('descripcion') || 'Servicios') : 'Servicios';
        const valorUnitario = conceptoElem ? (parseFloat(conceptoElem.getAttribute('ValorUnitario') || '0') || montoTotal) : montoTotal;

        // 1. Detección exhaustiva de Documentos Relacionados (DoctoRelacionado en pago20, pago10, cfdi, etc.)
        const doctosRelElements = allElements.filter(el => {
          const name = (el.localName || el.nodeName || '').toLowerCase();
          return name.endsWith('doctorelacionado') || name === 'doctorelacionado';
        });

        const relatedDocs = doctosRelElements.map(el => ({
          folio: (el.getAttribute('Folio') || el.getAttribute('folio') || '').trim(),
          serie: (el.getAttribute('Serie') || el.getAttribute('serie') || '').trim(),
          uuid: (el.getAttribute('IdDocumento') || el.getAttribute('idDocumento') || '').trim()
        })).filter(r => r.folio || r.uuid);

        // También buscar en cfdi:CfdiRelacionado si existiera
        const cfdiRelElements = allElements.filter(el => {
          const name = (el.localName || el.nodeName || '').toLowerCase();
          return name.endsWith('cfdirelacionado') || name === 'cfdirelacionado';
        });
        cfdiRelElements.forEach(cr => {
          const relUuid = (cr.getAttribute('UUID') || cr.getAttribute('uuid') || '').trim();
          if (relUuid && !relatedDocs.some(r => r.uuid.toLowerCase() === relUuid.toLowerCase())) {
            relatedDocs.push({ folio: '', serie: '', uuid: relUuid });
          }
        });

        // Determinar si es un Complemento de Pago (REP)
        const hasPagosTag = allElements.some(el => (el.localName || el.nodeName || '').toLowerCase().includes('pagos'));
        const isComplement = montoTotal === 0 || tipoComprobante === 'P' || relatedDocs.length > 0 || hasPagosTag;
        const estatus: InvoiceStatus = isComplement ? 'Complemento' : 'Generado';

        let docRelacionadoSummary = '';
        if (relatedDocs.length > 0) {
          const firstRel = relatedDocs[0];
          docRelacionadoSummary = firstRel.folio || firstRel.uuid || (firstRel.serie ? `${firstRel.serie}-${firstRel.folio}` : '');
        }

        // Check if client exists in directory to calculate probable payment date
        let probableDate = '';
        const clientMatch = clients.find(c => (
          c.nombre.trim().toLowerCase() === nombreReceptor.trim().toLowerCase() || 
          (rfcReceptor && c.rfc.trim().toUpperCase() === rfcReceptor.trim().toUpperCase())
        ));

        if (clientMatch && !isComplement) {
          const emDate = new Date(fecha + 'T00:00:00');
          emDate.setDate(emDate.getDate() + (clientMatch.dias_credito || 30));
          if (clientMatch.dia_pago_fijo && clientMatch.dia_pago_fijo !== 'Mismo día exacto') {
            const targetDay = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'].indexOf(clientMatch.dia_pago_fijo);
            if (targetDay !== -1) {
              const diff = (targetDay + 7 - emDate.getDay()) % 7;
              emDate.setDate(emDate.getDate() + (diff === 0 ? 7 : diff));
            }
          }
          probableDate = toLocalDateString(emDate);
        }

        // PREVENCIÓN DE DUPLICADOS Y HUECOS: Buscar coincidencia exacta o numérica
        const existingInv = invoices.find(i => {
          const normA = (i.numero_factura || '').trim().toLowerCase();
          const normB = rawFolio.toLowerCase();
          return normA === normB || (normA.replace(/\D/g, '') === normB.replace(/\D/g, '') && normB.replace(/\D/g, '').length > 0);
        });

        const targetDocId = existingInv ? existingInv.id : `inv_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;

        const payload: any = {
          numero_factura: rawFolio,
          empresa: nombreReceptor,
          rfc_cliente: rfcReceptor,
          orden_de_compra: existingInv?.orden_de_compra || '',
          concepto: descripcion,
          precio_unitario: valorUnitario,
          monto_total: montoTotal,
          fecha_emision: fecha,
          fecha_probable_pago: existingInv?.fecha_probable_pago || probableDate,
          estatus: estatus,
          complemento: isComplement && docRelacionadoSummary ? `Relacionado con factura ${docRelacionadoSummary}` : (existingInv?.complemento || ''),
          documento_relacionado: docRelacionadoSummary,
          uuid: currentUuid || existingInv?.uuid || '',
          es_hueco_pendiente: false,
          updatedAt: new Date().toISOString()
        };

        if (!existingInv) {
          payload.createdAt = new Date().toISOString();
        }

        await setDoc(doc(db, 'invoices', targetDocId), payload, { merge: true });

        // 2. BÚSQUEDA Y ACTUALIZACIÓN INMEDIATA DE LA FACTURA ORIGINAL RELACIONADA
        let vinculadaCount = 0;
        if (isComplement && relatedDocs.length > 0) {
          let poolInvoices = [...invoices];
          try {
            const snap = await getDocs(collection(db, 'invoices'));
            const firestoreItems: any[] = [];
            snap.forEach(dSnap => firestoreItems.push({ id: dSnap.id, ...dSnap.data() }));
            if (firestoreItems.length > 0) poolInvoices = firestoreItems;
          } catch (errSnap) {
            console.warn('Consulta directa a Firestore de respaldo:', errSnap);
          }

          for (const rel of relatedDocs) {
            const relFolio = (rel.folio || '').trim().toLowerCase();
            const relSerie = (rel.serie || '').trim().toLowerCase();
            const relUuid = (rel.uuid || '').trim().toLowerCase();
            const fullRel = (rel.serie && rel.folio) ? `${relSerie}${relFolio}` : '';
            const hyphenRel = (rel.serie && rel.folio) ? `${relSerie}-${relFolio}` : '';
            const relDigits = relFolio.replace(/\D/g, '');

            const originalInvoice = poolInvoices.find(inv => {
              const m = typeof inv.monto_total === 'number' ? inv.monto_total : (parseFloat(inv.monto_total) || 0);
              if (m <= 0) return false;

              const invFolio = (inv.numero_factura || '').trim().toLowerCase();
              const invDigits = invFolio.replace(/\D/g, '');
              const invUuid = (inv.uuid || inv.folio_fiscal || '').trim().toLowerCase();
              const invDocRel = (inv.documento_relacionado || '').trim().toLowerCase();

              // 1. Coincidencia por UUID fiscal
              if (relUuid && invUuid && invUuid === relUuid) return true;
              if (relUuid && invFolio === relUuid) return true;

              // 2. Coincidencia por Folio exacto
              if (relFolio && invFolio === relFolio) return true;
              if (fullRel && invFolio === fullRel) return true;
              if (hyphenRel && invFolio === hyphenRel) return true;

              // 3. Coincidencia numérica
              if (relDigits && invDigits && relDigits === invDigits && relDigits.length >= 1) {
                return true;
              }

              // 4. Coincidencia con documento_relacionado
              if (relFolio && invDocRel === relFolio) return true;
              if (relUuid && invDocRel === relUuid) return true;

              return false;
            });

            if (originalInvoice) {
              const folioCompRef = rawFolio || 'S/N';
              const notaActualizada = `Complemento ${folioCompRef}`;

              // PRIORIDAD ABSOLUTA: se actualiza sin importar estatus_modificado_manualmente
              await updateDoc(doc(db, 'invoices', originalInvoice.id), {
                estatus: 'Finalizado',
                complemento: notaActualizada,
                documento_relacionado: folioCompRef,
                alerta_complemento_descartada: true,
                estatus_modificado_manualmente: false,
                updatedAt: new Date().toISOString()
              });

              originalInvoice.estatus = 'Finalizado';
              originalInvoice.complemento = notaActualizada;
              originalInvoice.documento_relacionado = folioCompRef;
              originalInvoice.alerta_complemento_descartada = true;
              vinculadaCount++;
            }
          }
        }

        let msg = `¡Factura ${rawFolio} registrada exitosamente!`;
        if (isComplement) {
          if (vinculadaCount > 0) {
            msg = `¡Complemento ${rawFolio} vinculado exitosamente con ${vinculadaCount} factura(s) de ingreso! Su estatus cambió a "Finalizado".`;
          } else if (relatedDocs.length > 0) {
            msg = `¡Complemento ${rawFolio} registrado! Nota: Se buscó la factura con folio/UUID "${docRelacionadoSummary}", pero aún no se encuentra registrada en el sistema.`;
          }
        } else if (existingInv) {
          msg += ' (Hueco/registro anterior actualizado sin duplicados)';
        }

        alert(msg);
      } catch (err) {
        console.error('Error parsing XML:', err);
        alert('Hubo un error al procesar el archivo XML.');
      } finally {
        if (xmlInputRef.current) xmlInputRef.current.value = '';
      }
    };
    reader.readAsText(file);
  };

  // CSV Import handler con conversor de fechas Excel
  const handleCsvImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setCsvLoading(true);
    setCsvStatusText(`Leyendo ${file.name}...`);

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = evt.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows: any[] = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

        if (rows.length === 0) {
          alert('El archivo CSV o Excel está vacío.');
          setCsvLoading(false);
          return;
        }

        setCsvStatusText(`Guardando ${rows.length} registros sin duplicar en Firestore...`);

        let batch = writeBatch(db);
        let count = 0;
        let committed = 0;

        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          const folio = String(
            row.Folio || row.folio || row.Factura || row.factura || 
            row.Numero || row.numero || row['No.'] || row['No'] || ''
          ).trim();
          if (!folio) continue;

          const empresa = String(
            row.Receptor || row.receptor || row.Empresa || row.empresa || 
            row.Cliente || row.cliente || row['Razon Social'] || 'Cliente'
          ).trim();
          const ordenCompra = String(
            row.Orden_de_Compra || row.OC || row.oc || row['Orden de Compra'] || row.Orden || ''
          ).trim();
          const concepto = String(
            row.Concepto || row.concepto || row.Descripcion || row.descripcion || 'Servicios'
          ).trim();
          const montoStr = String(
            row.Total || row.total || row.Monto || row.monto || row.Importe || row.importe || '0'
          ).replace(/[^0-9.-]+/g, '');
          const monto = parseFloat(montoStr) || 0;
          
          const rawFecha = row.Fecha || row.fecha || row['Fecha Emision'] || row['Fecha de Emision'] || '';
          const fecha = parseExcelDate(rawFecha) || toLocalDateString(new Date());

          const isComplement = monto === 0 || String(row.Estatus || row.estatus || '').toLowerCase().includes('complemento');
          const estatus: InvoiceStatus = isComplement ? 'Complemento' : 'Generado';

          // Detectar existencia de folio para mergear sin duplicar
          const existing = invoices.find(inv => inv.numero_factura.toLowerCase() === folio.toLowerCase());
          const id = existing ? existing.id : `inv_${Date.now()}_${Math.random().toString(36).substr(2, 6)}`;

          batch.set(doc(db, 'invoices', id), {
            numero_factura: folio,
            empresa: empresa,
            rfc_cliente: '',
            orden_de_compra: ordenCompra,
            concepto: concepto,
            precio_unitario: monto,
            monto_total: monto,
            fecha_emision: fecha,
            fecha_probable_pago: existing?.fecha_probable_pago || '',
            estatus: estatus,
            complemento: isComplement ? 'Complemento de pago' : (existing?.complemento || ''),
            documento_relacionado: '',
            es_hueco_pendiente: false,
            updatedAt: new Date().toISOString()
          }, { merge: true });

          count++;
          if (count === 400) {
            await batch.commit();
            committed += count;
            batch = writeBatch(db);
            count = 0;
          }
        }

        if (count > 0) {
          await batch.commit();
          committed += count;
        }

        alert(`¡Se procesaron ${committed} facturas exitosamente en Cloud Firestore!`);
        setShowCsvModal(false);
      } catch (err) {
        console.error('Error importing CSV:', err);
        alert('Hubo un error al procesar el archivo CSV.');
      } finally {
        setCsvLoading(false);
        if (csvInputRef.current) csvInputRef.current.value = '';
      }
    };
    reader.readAsBinaryString(file);
  };

  // Export to Excel
  const handleExportExcel = () => {
    if (invoices.length === 0) {
      alert('No hay facturas registradas para exportar.');
      return;
    }

    const rows = invoices.map((i) => ({
      '1. Folio': i.numero_factura,
      '2. Empresa / Receptor': i.empresa,
      '3. Orden de Compra': i.orden_de_compra,
      '4. Concepto': i.concepto,
      '5. Precio Unitario': i.precio_unitario,
      '6. Monto Total': i.monto_total,
      '7. Fecha Probable de Pago': i.fecha_probable_pago,
      '8. Estatus': i.estatus,
      '9. NOTAS': i.complemento
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Facturas');
    XLSX.writeFile(wb, `Control_Cobranza_CFDI_${toLocalDateString(new Date())}.xlsx`);
  };

  // Add client in modal
  const handleAddClient = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newClientName.trim()) return;

    try {
      const newId = `cli_${Date.now()}`;
      await setDoc(doc(db, 'clientes', newId), {
        nombre: newClientName.trim(),
        rfc: newClientRfc.trim().toUpperCase(),
        dias_credito: Number(newClientDays) || 30,
        dia_pago_fijo: newClientDayFixed,
        notas: '',
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString()
      });

      setNewClientName('');
      setNewClientRfc('');
      setNewClientDays(30);
      setNewClientDayFixed('Mismo día exacto');
    } catch (err) {
      console.error('Error adding client:', err);
    }
  };

  // Recalculate client dates
  const handleRecalculateClient = async (cli: ClientItem) => {
    const normName = cli.nombre.trim().toLowerCase();
    const normRfc = cli.rfc.trim().toUpperCase();

    const batch = writeBatch(db);
    let count = 0;

    invoices.forEach((inv) => {
      const invEmpresa = inv.empresa.trim().toLowerCase();
      if (invEmpresa === normName || (normRfc && inv.rfc_cliente === normRfc)) {
        if (inv.estatus_modificado_manualmente || inv.estatus === 'Complemento' || inv.monto_total === 0 || inv.estatus === 'Cancelada') {
          return;
        }

        const emission = inv.fecha_emision || toLocalDateString(new Date());
        const emDate = new Date(emission + 'T00:00:00');
        emDate.setDate(emDate.getDate() + (cli.dias_credito || 30));

        if (cli.dia_pago_fijo && cli.dia_pago_fijo !== 'Mismo día exacto') {
          const targetDay = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'].indexOf(cli.dia_pago_fijo);
          if (targetDay !== -1) {
            const diff = (targetDay + 7 - emDate.getDay()) % 7;
            emDate.setDate(emDate.getDate() + (diff === 0 ? 7 : diff));
          }
        }

        const newProbable = toLocalDateString(emDate);
        batch.update(doc(db, 'invoices', inv.id), {
          fecha_probable_pago: newProbable,
          updatedAt: new Date().toISOString()
        });
        count++;
      }
    });

    if (count > 0) {
      await batch.commit();
      alert(`Se actualizaron las fechas de pago para ${count} facturas de ${cli.nombre}.`);
    } else {
      alert(`No se encontraron facturas pendientes para recalcular.`);
    }
  };

  // Calendar calculations
  const calYear = calDate.getFullYear();
  const calMonth = calDate.getMonth();
  const monthNames = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const firstDayOfWeek = new Date(calYear, calMonth, 1).getDay();
  const daysInCalMonth = new Date(calYear, calMonth + 1, 0).getDate();

  const calendarDayInvoices = useMemo(() => {
    const map: Record<string, InvoiceItem[]> = {};
    invoices.forEach((inv) => {
      if (!inv.fecha_probable_pago || inv.estatus === 'Cancelada' || inv.estatus === 'Complemento' || inv.monto_total <= 0) {
        return;
      }
      if (!map[inv.fecha_probable_pago]) {
        map[inv.fecha_probable_pago] = [];
      }
      map[inv.fecha_probable_pago].push(inv);
    });
    return map;
  }, [invoices]);

  // LOGIN GATE SCREEN
  if (!isAuthenticated) {
    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/95 backdrop-blur-md px-4">
        <div className="bg-slate-900 border border-slate-800 rounded-3xl p-8 max-w-md w-full shadow-2xl space-y-6 relative overflow-hidden">
          <div className="absolute -right-12 -top-12 w-40 h-40 bg-emerald-500/10 rounded-full blur-2xl pointer-events-none" />

          <div className="text-center space-y-2">
            <div className="w-16 h-16 bg-emerald-950/80 border border-emerald-500/40 rounded-2xl mx-auto flex items-center justify-center text-emerald-400 shadow-lg shadow-emerald-950">
              <ShieldCheck className="w-8 h-8" />
            </div>
            <h2 className="text-2xl font-bold text-white tracking-tight">Control de Cobranza CFDI</h2>
            <p className="text-xs text-slate-400">Ingresa la clave de acceso para acceder al panel financiero</p>
          </div>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-slate-300 mb-1.5 uppercase tracking-wider">
                Contraseña de Acceso
              </label>
              <div className="relative">
                <input
                  type={showPassword ? 'text' : 'password'}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                  required
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-4 py-3 text-sm text-white placeholder-slate-500 focus:outline-none focus:ring-2 focus:ring-emerald-500 focus:border-transparent transition-all"
                />
                <button
                  type="button"
                  onClick={() => setShowPassword(!showPassword)}
                  className="absolute right-3 top-3 text-slate-400 hover:text-slate-200 transition-colors"
                >
                  {showPassword ? <EyeOff className="w-5 h-5" /> : <Eye className="w-5 h-5" />}
                </button>
              </div>
            </div>

            {loginError && (
              <div className="p-3 bg-rose-950/50 border border-rose-800/60 rounded-xl text-xs text-rose-300 flex items-center gap-2">
                <AlertTriangle className="w-4 h-4 shrink-0 text-rose-400" />
                <span>Contraseña incorrecta. Inténtalo nuevamente.</span>
              </div>
            )}

            <button
              type="submit"
              className="w-full bg-emerald-600 hover:bg-emerald-500 text-white font-semibold py-3 rounded-xl shadow-lg shadow-emerald-950 transition-all flex items-center justify-center gap-2 text-sm cursor-pointer"
            >
              <span>Ingresar al Sistema</span>
              <ArrowRight className="w-4 h-4" />
            </button>
          </form>

          <div className="pt-2 text-center text-[11px] text-slate-500">
            Plataforma protegida con autenticación y sincronización segura en Cloud Firestore
          </div>
        </div>
      </div>
    );
  }

  // MAIN APPLICATION SCREEN: FULL WIDTH DESIGN
  return (
    <div className="min-h-screen flex flex-col w-full bg-slate-950 text-slate-100 selection:bg-emerald-500 selection:text-white">
      {/* Hidden file inputs */}
      <input
        ref={xmlInputRef}
        type="file"
        accept=".xml"
        onChange={handleXmlUpload}
        className="hidden"
      />
      <input
        ref={csvInputRef}
        type="file"
        accept=".csv,.xlsx,.xls"
        onChange={handleCsvImport}
        className="hidden"
      />

      {/* HEADER: FULL WIDTH */}
      <header className="bg-slate-900/90 border-b border-slate-800 sticky top-0 z-40 backdrop-blur-md w-full">
        <div className="w-full px-3 sm:px-6 lg:px-8 min-h-[4rem] py-2 sm:py-0 flex flex-col sm:flex-row items-center justify-between gap-2.5 sm:gap-4">
          <div className="flex items-center justify-between w-full sm:w-auto gap-3 shrink-0">
            <div className="flex items-center gap-2.5 sm:gap-3">
              <div className="w-9 h-9 sm:w-10 sm:h-10 bg-emerald-950 border border-emerald-500/40 rounded-xl flex items-center justify-center text-emerald-400 shadow-md shadow-emerald-950 shrink-0">
                <Receipt className="w-4 h-4 sm:w-5 sm:h-5" />
              </div>
              <div>
                <h1 className="text-sm sm:text-base font-bold text-white leading-tight">Control de Cobranza CFDI</h1>
                <p className="text-[10px] sm:text-[11px] text-emerald-400 font-medium flex items-center gap-1.5">
                  <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
                  <span>{isFirebaseConnected ? 'Cloud Firestore Conectado' : 'Conectando a base de datos...'}</span>
                </p>
              </div>
            </div>

            {/* Logout button visible on mobile header */}
            <button
              onClick={handleLogout}
              title="Cerrar sesión"
              className="sm:hidden p-1.5 text-slate-400 hover:text-rose-400 rounded-lg hover:bg-slate-800"
            >
              <LogOut className="w-4 h-4" />
            </button>
          </div>

          {/* Action buttons with proper wrap, gap-2 and full labels */}
          <div className="flex flex-wrap items-center gap-2 w-full md:w-auto justify-start md:justify-end">
            
            {/* BOTÓN: GRÁFICA DE COBRANZA */}
            <button
              id="openChartModalBtn"
              onClick={() => {
                setChartCenterDate(new Date());
                setShowChartModal(true);
              }}
              className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white rounded-xl text-xs font-semibold border border-slate-700 hover:border-emerald-500/40 transition-all flex items-center gap-1.5 shadow-sm cursor-pointer whitespace-nowrap shrink-0"
              title="Abrir flujo financiero y proyección en gráfica tipo montaña"
            >
              <TrendingUp className="w-4 h-4 text-emerald-400" />
              <span>Gráfica de Cobranza</span>
            </button>

            {/* Calendario de Cobranza */}
            <button
              id="openCalendarBtn"
              onClick={() => setShowCalendarModal(true)}
              className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white rounded-xl text-xs font-semibold border border-slate-700 hover:border-emerald-500/40 transition-all flex items-center gap-1.5 shadow-sm cursor-pointer whitespace-nowrap shrink-0"
            >
              <CalendarIcon className="w-4 h-4 text-emerald-400" />
              <span>Calendario de Cobranza</span>
            </button>

            {/* Directorio de Clientes */}
            <button
              id="openClientsBtn"
              onClick={() => setShowClientsModal(true)}
              className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white rounded-xl text-xs font-semibold border border-slate-700 hover:border-blue-500/40 transition-all flex items-center gap-1.5 shadow-sm cursor-pointer whitespace-nowrap shrink-0"
            >
              <Users className="w-4 h-4 text-blue-400" />
              <span>Directorio de Clientes</span>
            </button>

            {/* Importar CSV Anual */}
            <button
              id="openCsvImportBtn"
              onClick={() => setShowCsvModal(true)}
              className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white rounded-xl text-xs font-semibold border border-slate-700 hover:border-amber-500/40 transition-all flex items-center gap-1.5 shadow-sm cursor-pointer whitespace-nowrap shrink-0"
            >
              <FileSpreadsheet className="w-4 h-4 text-amber-400" />
              <span>Importar CSV Anual</span>
            </button>

            {/* Excel */}
            <button
              id="exportExcelBtn"
              onClick={handleExportExcel}
              className="px-3 py-2 bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-300 border border-emerald-500/40 rounded-xl text-xs font-semibold transition-all flex items-center gap-1.5 shadow-sm cursor-pointer whitespace-nowrap shrink-0"
            >
              <Download className="w-4 h-4 text-emerald-400" />
              <span>Excel</span>
            </button>

            {/* Cerrar Sesión (Desktop) */}
            <button
              id="logoutBtn"
              onClick={handleLogout}
              title="Cerrar sesión"
              className="hidden md:inline-flex p-2 text-slate-400 hover:text-rose-400 hover:bg-rose-950/30 rounded-xl border border-transparent hover:border-rose-900 transition-all cursor-pointer shrink-0"
            >
              <LogOut className="w-4 h-4" />
            </button>
          </div>
        </div>
      </header>

      {/* MAIN CONTAINER: FULL WIDTH */}
      <main className="flex-1 w-full px-3 sm:px-5 lg:px-6 py-5 space-y-5">

        {/* SEQUENCE GAP ALERT BANNER */}
        {missingFolios.length > 0 && (
          <div className="bg-amber-950/50 border border-amber-600/50 rounded-2xl p-4 shadow-lg flex items-start gap-3.5 text-amber-200 text-xs">
            <AlertTriangle className="w-5 h-5 text-amber-400 shrink-0 mt-0.5" />
            <div className="flex-1 space-y-1">
              <div className="font-bold text-amber-300 text-sm">Discontinuidad detectada en la secuencia de facturas</div>
              <div>
                Se detectaron {missingFolios.length} folios faltantes en la secuencia:{' '}
                <span className="font-mono font-bold text-amber-100">
                  [{missingFolios.slice(0, 10).join(', ')}{missingFolios.length > 10 ? '...' : ''}]
                </span>
                . Al subir el XML correspondiente posteriormente, se actualizará automáticamente sin duplicar registros.
              </div>
            </div>
          </div>
        )}

        {/* METRICS & GLOBAL PERIOD FILTER */}
        <div className="space-y-3">
          <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-900/60 border border-slate-800 rounded-2xl p-3">
            <div className="flex items-center gap-2 text-xs font-semibold text-slate-300">
              <Filter className="w-4 h-4 text-emerald-400" />
              <span>Filtrar período global de métricas:</span>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <select
                value={periodFilter}
                onChange={(e) => setPeriodFilter(e.target.value as any)}
                className="bg-slate-950 border border-slate-700 text-xs text-white rounded-xl px-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
              >
                <option value="todo">Todo el historial</option>
                <option value="ultimo_mes">Último mes</option>
                <option value="ultimos_dos_meses">Últimos dos meses</option>
                <option value="ano_actual">Lo que va del año</option>
                <option value="personalizado">Personalizado</option>
              </select>

              {periodFilter === 'personalizado' && (
                <div className="flex items-center gap-2">
                  <input
                    type="date"
                    value={customStartDate}
                    onChange={(e) => setCustomStartDate(e.target.value)}
                    className="bg-slate-950 border border-slate-700 text-xs text-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                  <span className="text-xs text-slate-400">al</span>
                  <input
                    type="date"
                    value={customEndDate}
                    onChange={(e) => setCustomEndDate(e.target.value)}
                    className="bg-slate-950 border border-slate-700 text-xs text-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                </div>
              )}
            </div>
          </div>

          {/* 4 Cards Grid (2x2 en móvil, 4 cols en desktop) */}
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-2.5 sm:gap-3.5">
            <div className="bg-slate-900/90 border border-slate-800 rounded-2xl p-3 sm:p-4 shadow-md relative overflow-hidden">
              <div className="flex items-center justify-between text-slate-400 text-xs font-semibold mb-1">
                <span className="truncate text-[11px] sm:text-xs">Monto Total Facturado</span>
                <div className="w-6 h-6 sm:w-7 sm:h-7 bg-blue-950/60 rounded-lg flex items-center justify-center text-blue-400 shrink-0">
                  <DollarSign className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
                </div>
              </div>
              <div className="text-base sm:text-xl lg:text-2xl font-black text-white truncate">{formatCurrency(metrics.totalFacturado)}</div>
              <div className="text-[10px] sm:text-[11px] text-slate-400 mt-1 truncate">{metrics.countFacturas} comprobantes activos</div>
            </div>

            <div className="bg-slate-900/90 border border-slate-800 rounded-2xl p-3 sm:p-4 shadow-md relative overflow-hidden">
              <div className="flex items-center justify-between text-slate-400 text-xs font-semibold mb-1">
                <span className="truncate text-[11px] sm:text-xs">Pendiente de Cobro</span>
                <div className="w-6 h-6 sm:w-7 sm:h-7 bg-amber-950/60 rounded-lg flex items-center justify-center text-amber-400 shrink-0">
                  <Clock className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
                </div>
              </div>
              <div className="text-base sm:text-xl lg:text-2xl font-black text-amber-400 truncate">{formatCurrency(metrics.pendienteCobro)}</div>
              <div className="text-[10px] sm:text-[11px] text-amber-300/80 mt-1 truncate">{metrics.countPendiente} facturas por liquidar</div>
            </div>

            <div className="bg-slate-900/90 border border-slate-800 rounded-2xl p-3 sm:p-4 shadow-md relative overflow-hidden">
              <div className="flex items-center justify-between text-slate-400 text-xs font-semibold mb-1">
                <span className="truncate text-[11px] sm:text-xs">Total Cobrado</span>
                <div className="w-6 h-6 sm:w-7 sm:h-7 bg-emerald-950/60 rounded-lg flex items-center justify-center text-emerald-400 shrink-0">
                  <CheckCircle2 className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
                </div>
              </div>
              <div className="text-base sm:text-xl lg:text-2xl font-black text-emerald-400 truncate">{formatCurrency(metrics.totalCobrado)}</div>
              <div className="text-[10px] sm:text-[11px] text-emerald-300/80 mt-1 truncate">{metrics.countCobrado} facturas concluidas</div>
            </div>

            <div className="bg-slate-900/90 border border-slate-800 rounded-2xl p-3 sm:p-4 shadow-md relative overflow-hidden">
              <div className="flex items-center justify-between text-slate-400 text-xs font-semibold mb-1">
                <span className="truncate text-[11px] sm:text-xs">Procedimiento Parcial</span>
                <div className="w-6 h-6 sm:w-7 sm:h-7 bg-rose-950/60 rounded-lg flex items-center justify-center text-rose-400 shrink-0">
                  <AlertCircle className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
                </div>
              </div>
              <div className="text-base sm:text-xl lg:text-2xl font-black text-rose-400 truncate">{formatCurrency(metrics.problemaMonto)}</div>
              <div className="text-[10px] sm:text-[11px] text-rose-300/80 mt-1 truncate">{metrics.countProblema} facturas en revisión</div>
            </div>
          </div>
        </div>

        {/* XML CFDI UPLOADER (GESTIÓN DE HUECOS / CERO DUPLICIDAD) */}
        <div className="bg-slate-900/80 border border-slate-800 rounded-2xl p-4 shadow-lg flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-emerald-950/80 border border-emerald-500/40 rounded-xl flex items-center justify-center text-emerald-400 shrink-0">
              <FileUp className="w-5 h-5" />
            </div>
            <div>
              <div className="text-sm font-bold text-white">Subir Factura o Complemento CFDI (XML)</div>
              <div className="text-xs text-slate-400">
                Lectura y cruce inteligente. Si existía un hueco o registro con el mismo folio, se actualiza automáticamente sin crear duplicados.
              </div>
            </div>
          </div>

          <div className="flex items-center gap-2 w-full md:w-auto">
            <button
              onClick={() => xmlInputRef.current?.click()}
              className="w-full md:w-auto px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold shadow-md shadow-emerald-950 transition-all flex items-center justify-center gap-2 cursor-pointer"
            >
              <Upload className="w-4 h-4" />
              <span>Cargar Archivo XML</span>
            </button>
          </div>
        </div>

        {/* TABLA DE CONTROL DE COBRANZA (ANCHO COMPLETO + FILTRADO AVANZADO) */}
        <div className="bg-slate-900/90 border border-slate-800 rounded-2xl shadow-xl overflow-hidden flex flex-col w-full">
          
          {/* Barra de herramientas superior con FILTRADO AVANZADO */}
          <div className="p-4 border-b border-slate-800 flex flex-col gap-3">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
              <div className="flex items-center gap-2">
                <TableIcon className="w-4 h-4 text-emerald-400" />
                <h2 className="text-base font-bold text-white">Tabla de Facturas y Complementos</h2>
                <span className="px-2.5 py-0.5 bg-slate-800 text-slate-300 rounded-full text-xs font-semibold">
                  {displayInvoices.length} de {invoices.length} registros
                </span>
              </div>

              {/* Grupo Buscador, Botón Atención Urgente y Botón Facturas con Problema */}
              <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5 w-full md:w-auto">
                <button
                  type="button"
                  onClick={() => setShowUrgentModal(true)}
                  className={`w-full sm:w-auto justify-center px-3.5 py-2.5 rounded-xl text-xs font-bold border flex items-center gap-2 transition-all shadow-sm cursor-pointer whitespace-nowrap ${
                    urgentInvoices.length > 0
                      ? 'bg-amber-950/40 hover:bg-amber-900/60 border-amber-500/50 text-amber-200 shadow-amber-950/50 ring-1 ring-amber-500/30'
                      : 'bg-slate-800/90 hover:bg-slate-800 border-slate-700/80 text-slate-400'
                  }`}
                  title="Ver facturas que requieren complemento de pago"
                >
                  {urgentInvoices.length > 0 ? (
                    <AlertTriangle className="w-4 h-4 text-amber-400 animate-pulse" />
                  ) : (
                    <Bell className="w-4 h-4 text-slate-400" />
                  )}
                  <span>Atención Urgente</span>
                  <span
                    className={`px-2 py-0.5 rounded-full text-[11px] font-black ${
                      urgentInvoices.length > 0
                        ? 'bg-amber-500 text-slate-950'
                        : 'bg-slate-700 text-slate-400'
                    }`}
                  >
                    {urgentInvoices.length}
                  </span>
                </button>

                {/* Botón Facturas con Problema */}
                <button
                  type="button"
                  onClick={() => setShowProblemaModal(true)}
                  className={`w-full sm:w-auto justify-center px-3.5 py-2.5 rounded-xl text-xs font-bold border flex items-center gap-2 transition-all shadow-sm cursor-pointer whitespace-nowrap ${
                    problemaInvoices.length > 0
                      ? 'bg-rose-950/40 hover:bg-rose-900/60 border-rose-500/50 text-rose-200 shadow-rose-950/50 ring-1 ring-rose-500/30'
                      : 'bg-slate-800/90 hover:bg-slate-800 border-slate-700/80 text-slate-400'
                  }`}
                  title="Ver facturas reportadas con problema"
                >
                  <AlertOctagon
                    className={`w-4 h-4 ${
                      problemaInvoices.length > 0 ? 'text-rose-400 animate-pulse' : 'text-slate-400'
                    }`}
                  />
                  <span>Facturas con Problema</span>
                  <span
                    className={`px-2 py-0.5 rounded-full text-[11px] font-black ${
                      problemaInvoices.length > 0
                        ? 'bg-rose-600 text-white'
                        : 'bg-slate-700 text-slate-400'
                    }`}
                  >
                    {problemaInvoices.length}
                  </span>
                </button>

                {/* Buscador global */}
                <div className="relative w-full md:w-80">
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Buscar folio, empresa, OC, concepto..."
                    className="w-full bg-slate-950 border border-slate-700 text-xs text-white rounded-xl pl-9 pr-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-emerald-500 placeholder:text-slate-500"
                  />
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-3" />
                </div>
              </div>
            </div>

            {/* Controles de Filtrado Rápido Adicional sobre la tabla */}
            <div className="pt-2 border-t border-slate-800/80 flex flex-wrap items-center justify-between gap-2.5 text-xs">
              
              <div className="flex items-center gap-3 flex-wrap">
                {/* BOTÓN Y FILTRO DIRECTO POR EMPRESA */}
                <div className="flex items-center gap-1.5 flex-wrap">
                  <button
                    type="button"
                    onClick={() => setShowEmpresaFilterModal(true)}
                    className={`px-3 py-1.5 rounded-xl text-xs font-semibold border flex items-center gap-1.5 transition-all shadow-sm cursor-pointer ${
                      tableEmpresaFilter !== 'todas'
                        ? 'bg-cyan-950/80 border-cyan-500/60 text-cyan-300 ring-2 ring-cyan-500/20 shadow-cyan-950/50'
                        : 'bg-slate-800 hover:bg-slate-700 border-slate-700 text-slate-200 hover:text-white'
                    }`}
                    title="Abrir catálogo y filtrar por empresa"
                  >
                    <Building2 className="w-3.5 h-3.5 text-cyan-400" />
                    <span>Filtrar por Empresa</span>
                    {tableEmpresaFilter !== 'todas' ? (
                      <span className="px-1.5 py-0.2 bg-cyan-500/20 text-cyan-300 rounded-md text-[10px] font-bold border border-cyan-500/30">
                        Activo
                      </span>
                    ) : (
                      <span className="text-slate-400 text-[10px]">({empresaStats.length})</span>
                    )}
                  </button>

                  <select
                    value={tableEmpresaFilter}
                    onChange={(e) => setTableEmpresaFilter(e.target.value)}
                    className="bg-slate-950 border border-slate-700 text-xs text-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:ring-2 focus:ring-cyan-500 cursor-pointer max-w-[210px] truncate"
                    title="Seleccionar empresa receptora"
                  >
                    <option value="todas">Todas las empresas ({empresaStats.length})</option>
                    {empresaStats.map((item) => (
                      <option key={item.nombre} value={item.nombre}>
                        {item.nombre} ({item.totalFacturas})
                      </option>
                    ))}
                  </select>

                  {tableEmpresaFilter !== 'todas' && (
                    <div className="flex items-center gap-1 px-2.5 py-1 bg-cyan-950/70 border border-cyan-700/80 text-cyan-300 rounded-xl text-xs font-medium animate-fadeIn">
                      <span className="max-w-[130px] truncate font-semibold" title={tableEmpresaFilter}>
                        {tableEmpresaFilter}
                      </span>
                      <button
                        type="button"
                        onClick={() => setTableEmpresaFilter('todas')}
                        className="hover:text-white text-cyan-400 hover:bg-cyan-900/60 rounded p-0.5 transition-colors cursor-pointer"
                        title="Quitar filtro de empresa"
                      >
                        <X className="w-3 h-3" />
                      </button>
                    </div>
                  )}
                </div>

                {/* Filtro rápido por Estatus */}
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-slate-400 font-semibold flex items-center gap-1">
                    <ListFilter className="w-3.5 h-3.5 text-emerald-400" />
                    <span>Estatus:</span>
                  </span>
                  <select
                    value={tableStatusFilter}
                    onChange={(e) => setTableStatusFilter(e.target.value)}
                    className="bg-slate-950 border border-slate-700 text-xs text-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer"
                  >
                    <option value="todos">Todos los estatus</option>
                    <option value="Generado">Generado</option>
                    <option value="Procedimiento parcial">Procedimiento parcial</option>
                    <option value="Problema">Problema</option>
                    <option value="Procedimiento terminado">Procedimiento terminado</option>
                    <option value="Cancelada">Cancelada</option>
                    <option value="Pagada sin complemento">Pagada sin complemento</option>
                    <option value="Finalizado">Finalizado</option>
                    <option value="Complemento">Complemento ($0.00)</option>
                  </select>
                </div>
              </div>

              {/* Filtro por Rango Directo de Fecha de Pago Probable */}
              <div className="flex items-center gap-2 flex-wrap">
                <span className="text-slate-400 font-semibold flex items-center gap-1">
                  <CalendarRange className="w-3.5 h-3.5 text-blue-400" />
                  <span>F. Probable Pago:</span>
                </span>
                <input
                  type="date"
                  value={tableDateStart}
                  onChange={(e) => setTableDateStart(e.target.value)}
                  className="bg-slate-950 border border-slate-700 text-xs text-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                />
                <span className="text-slate-500 text-xs">-</span>
                <input
                  type="date"
                  value={tableDateEnd}
                  onChange={(e) => setTableDateEnd(e.target.value)}
                  className="bg-slate-950 border border-slate-700 text-xs text-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:ring-2 focus:ring-emerald-500"
                />
                <button
                  onClick={() => {
                    setSearchQuery('');
                    setTableEmpresaFilter('todas');
                    setTableStatusFilter('todos');
                    setTableDateStart('');
                    setTableDateEnd('');
                  }}
                  className="px-2.5 py-1.5 text-[11px] bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl transition-all border border-slate-700 cursor-pointer"
                  title="Restablecer todos los filtros de la tabla"
                >
                  Limpiar
                </button>
              </div>

            </div>
          </div>

          {/* Barra de desplazamiento superior sincronizada interactiva */}
          <div className="px-4 py-2.5 bg-slate-950/80 border-b border-slate-800 flex flex-wrap items-center justify-between gap-3 text-xs">
            <div className="flex items-center gap-2 text-slate-300">
              <span className="font-semibold text-emerald-400 flex items-center gap-1.5">
                <MoveHorizontal className="w-4 h-4" />
                <span>Navegación Horizontal:</span>
              </span>
              <span className="text-[11px] text-slate-400 hidden sm:inline">
                Desplaza a lo largo de las 10 columnas en pantallas de cualquier tamaño
              </span>
            </div>

            <div className="flex items-center gap-2 flex-1 max-w-lg">
              <button
                type="button"
                onClick={() => tableContainerRef.current?.scrollBy({ left: -400, behavior: 'smooth' })}
                className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white rounded-lg border border-slate-700 font-semibold text-[11px] flex items-center gap-1 transition-all shadow-sm cursor-pointer"
              >
                <ChevronLeft className="w-3.5 h-3.5 text-emerald-400" />
                <span className="hidden md:inline">Izquierda</span>
              </button>

              <div
                ref={topScrollRef}
                className="custom-scroll-top flex-1 h-3.5 bg-slate-900 rounded-full border border-slate-700 overflow-x-auto overflow-y-hidden shadow-inner cursor-pointer"
              >
                <div style={{ width: `${tableScrollWidth}px`, height: '14px' }} />
              </div>

              <button
                type="button"
                onClick={() => tableContainerRef.current?.scrollBy({ left: 400, behavior: 'smooth' })}
                className="px-2.5 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white rounded-lg border border-slate-700 font-semibold text-[11px] flex items-center gap-1 transition-all shadow-sm cursor-pointer"
              >
                <span className="hidden md:inline">Derecha</span>
                <ChevronRight className="w-3.5 h-3.5 text-emerald-400" />
              </button>
            </div>
          </div>

          {/* Tabla de 10 Columnas Completa (con aceleración táctil para móviles) */}
          <div
            ref={tableContainerRef}
            className="w-full overflow-x-auto overflow-y-auto max-h-[70vh] shadow-md rounded-xl p-2 min-h-[380px] focus:outline-none touch-scroll"
            style={{ WebkitOverflowScrolling: 'touch' }}
          >
            <table className="w-full min-w-[1700px] text-left text-xs border-collapse">
              <thead>
                <tr className="bg-slate-950/90 border-b border-slate-800 text-slate-400 font-semibold uppercase tracking-wider text-[11px] select-none sticky top-0 z-20 backdrop-blur-md">
                  <th className="py-3 px-3 text-center min-w-[110px] w-28">1. Folio</th>
                  <th className="py-3 px-4 min-w-[250px]">
                    <div className="flex items-center justify-between gap-1">
                      <span>2. Empresa / Receptor</span>
                      <button
                        type="button"
                        onClick={() => setShowEmpresaFilterModal(true)}
                        title="Filtrar por empresa"
                        className="p-1 hover:bg-slate-800 text-slate-400 hover:text-cyan-400 rounded transition-colors cursor-pointer"
                      >
                        <Building2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  </th>
                  <th className="py-3 px-3 min-w-[140px]">3. Orden de Compra</th>
                  <th className="py-3 px-4 min-w-[280px]">4. Concepto</th>
                  <th className="py-3 px-3 text-right min-w-[120px]">5. P. Unitario</th>
                  <th className="py-3 px-3 text-right min-w-[130px]">6. Monto Total</th>
                  <th className="py-3 px-3 min-w-[160px]">7. Fecha Probable Pago</th>
                  <th className="py-3 px-3 min-w-[195px]">8. Estatus</th>
                  <th className="py-3 px-4 min-w-[240px]">9. NOTAS</th>
                  <th className="py-3 px-3 text-center min-w-[110px]">10. ACCIÓN</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/60 font-medium">
                {displayInvoices.length === 0 ? (
                  <tr>
                    <td colSpan={10} className="py-12 text-center text-slate-500">
                      No se encontraron facturas con los filtros aplicados.
                    </td>
                  </tr>
                ) : (
                  displayInvoices.map((inv) => {
                    const isGap = Boolean(inv.es_hueco_pendiente);
                    let rowBg = 'hover:bg-slate-800/40 transition-colors';
                    if (isGap) {
                      rowBg = 'bg-amber-950/20 border-l-4 border-l-amber-500 hover:bg-amber-950/30';
                    } else if (inv.estatus === 'Cancelada') {
                      rowBg = 'bg-rose-950/20 opacity-75 hover:bg-rose-950/30';
                    } else if (inv.estatus === 'Finalizado') {
                      rowBg = 'bg-emerald-950/10 hover:bg-emerald-950/20';
                    } else if (inv.estatus === 'Complemento') {
                      rowBg = 'bg-purple-950/10 hover:bg-purple-950/20';
                    }

                    return (
                      <tr key={inv.id} className={rowBg}>
                        {/* 1. Folio */}
                        <td className="py-2.5 px-3 text-center whitespace-nowrap">
                          <span className="font-bold text-white px-2 py-1 bg-slate-800 rounded-lg border border-slate-700/80">
                            {inv.numero_factura || 'S/N'}
                          </span>
                        </td>

                        {/* 2. Empresa */}
                        <td className="py-2.5 px-4 font-semibold text-slate-200">
                          <div className="flex items-center justify-between gap-1.5 group">
                            <span className="truncate max-w-[240px]" title={inv.empresa}>
                              {inv.empresa || '-'}
                            </span>
                            {inv.empresa && (
                              <button
                                type="button"
                                onClick={() => setTableEmpresaFilter(inv.empresa)}
                                title={`Filtrar tabla por "${inv.empresa}"`}
                                className="opacity-0 group-hover:opacity-100 hover:text-cyan-300 text-slate-500 hover:bg-slate-800 transition-all p-1 rounded cursor-pointer shrink-0"
                              >
                                <Filter className="w-3 h-3" />
                              </button>
                            )}
                          </div>
                        </td>

                        {/* 3. Orden de Compra */}
                        <td className="py-2.5 px-3 text-slate-300 font-mono text-[11px] whitespace-nowrap">
                          {inv.orden_de_compra ? (
                            <span className="bg-slate-800 px-2 py-0.5 rounded border border-slate-700">
                              {inv.orden_de_compra}
                            </span>
                          ) : (
                            <span className="text-slate-600">-</span>
                          )}
                        </td>

                        {/* 4. Concepto */}
                        <td className="py-2.5 px-4 text-slate-300">
                          <div className="truncate max-w-[290px]" title={inv.concepto}>
                            {inv.concepto || 'Sin descripción'}
                          </div>
                        </td>

                        {/* 5. Precio Unitario */}
                        <td className="py-2.5 px-3 text-right font-mono text-slate-300 whitespace-nowrap">
                          {formatCurrency(inv.precio_unitario)}
                        </td>

                        {/* 6. Monto Total */}
                        <td className={`py-2.5 px-3 text-right font-mono font-bold whitespace-nowrap ${inv.monto_total === 0 ? 'text-purple-400' : 'text-emerald-400'}`}>
                          {formatCurrency(inv.monto_total)}
                        </td>

                        {/* 7. Fecha Probable de Pago (3 Selectores: Día, Mes, Año) */}
                        <td className="py-2.5 px-3 whitespace-nowrap">
                          <CompoundDateSelector
                            id={inv.id}
                            value={inv.fecha_probable_pago}
                            onSave={(dateStr) => updateInvoiceField(inv.id, 'fecha_probable_pago', dateStr)}
                          />
                        </td>

                        {/* 8. Estatus */}
                        <td className="py-2.5 px-3 whitespace-nowrap">
                          <select
                            value={inv.estatus}
                            onChange={(e) => updateInvoiceField(inv.id, 'estatus', e.target.value as InvoiceStatus)}
                            className={`w-full px-2.5 py-1.5 rounded-lg text-xs font-semibold border focus:outline-none focus:ring-2 focus:ring-emerald-500 cursor-pointer ${getStatusBadgeStyle(inv.estatus)}`}
                          >
                            {INVOICE_STATUS_OPTIONS.map((opt) => (
                              <option key={opt} value={opt}>
                                {opt}
                              </option>
                            ))}
                          </select>
                        </td>

                        {/* 9. NOTAS (Componente seguro) */}
                        <td className="py-2.5 px-4">
                          <NotesCell
                            initialValue={inv.complemento || ''}
                            onSave={(newVal) => updateInvoiceField(inv.id, 'complemento', newVal)}
                          />
                        </td>

                        {/* 10. ACCIÓN */}
                        <td className="py-2.5 px-3 text-center whitespace-nowrap">
                          <button
                            type="button"
                            onClick={() => deleteInvoice(inv.id, inv.numero_factura)}
                            className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-semibold text-rose-300 hover:text-white bg-rose-950/40 hover:bg-rose-600 border border-rose-800/60 hover:border-rose-500 transition-all cursor-pointer shadow-sm"
                          >
                            <Trash2 className="w-3.5 h-3.5 text-rose-400" />
                            <span>Eliminar</span>
                          </button>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      </main>

      {/* FOOTER */}
      <footer className="border-t border-slate-800 py-4 px-4 text-center text-xs text-slate-500 w-full">
        Control de Cobranza CFDI &bull; Plataforma fluida y de alta precisión &bull; Cloud Firestore
      </footer>

      {/* ==================== NUEVO MODAL: GRÁFICA TIPO MONTAÑA DE COBRANZA ==================== */}
      {showChartModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/85 backdrop-blur-md p-2 sm:p-5 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl w-full max-w-6xl p-4 sm:p-7 shadow-2xl space-y-4 sm:space-y-5 my-auto max-h-[94vh] sm:max-h-[90vh] overflow-y-auto touch-scroll">
            
            <div className="flex items-center justify-between pb-3 border-b border-slate-800">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-emerald-950 border border-emerald-500/40 rounded-xl flex items-center justify-center text-emerald-400 shrink-0">
                  <TrendingUp className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white">Gráfica de Cobranza (Flujo Tipo Montaña)</h3>
                  <p className="text-xs text-slate-400">Visualización de montos cobrados frente a cobros proyectados con la fecha de hoy centrada</p>
                </div>
              </div>
              <button
                onClick={() => setShowChartModal(false)}
                className="text-slate-400 hover:text-white p-2 rounded-xl hover:bg-slate-800 transition-colors cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Controles de Navegación y Centrado */}
            <div className="flex flex-wrap items-center justify-between gap-3 bg-slate-950/70 p-3 rounded-2xl border border-slate-800">
              <div className="flex items-center gap-2 text-xs">
                <span className="text-slate-400 font-medium">Ventana de tiempo:</span>
                <button
                  onClick={() => setChartDaysSpan(14)}
                  className={`px-2.5 py-1 rounded-lg text-xs font-semibold border cursor-pointer ${chartDaysSpan === 14 ? 'bg-emerald-600/30 text-emerald-300 border-emerald-500/40' : 'bg-slate-800 text-slate-300 border-slate-700'}`}
                >
                  14 Días
                </button>
                <button
                  onClick={() => setChartDaysSpan(30)}
                  className={`px-2.5 py-1 rounded-lg text-xs font-semibold border cursor-pointer ${chartDaysSpan === 30 ? 'bg-emerald-600/30 text-emerald-300 border-emerald-500/40' : 'bg-slate-800 text-slate-300 border-slate-700'}`}
                >
                  30 Días (±15d)
                </button>
                <button
                  onClick={() => setChartDaysSpan(60)}
                  className={`px-2.5 py-1 rounded-lg text-xs font-semibold border cursor-pointer ${chartDaysSpan === 60 ? 'bg-emerald-600/30 text-emerald-300 border-emerald-500/40' : 'bg-slate-800 text-slate-300 border-slate-700'}`}
                >
                  60 Días (±30d)
                </button>
              </div>

              <div className="flex items-center gap-2">
                <button
                  onClick={() => {
                    const shift = Math.max(7, Math.floor(chartDaysSpan / 2));
                    const d = new Date(chartCenterDate);
                    d.setDate(d.getDate() - shift);
                    setChartCenterDate(d);
                  }}
                  className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-semibold flex items-center gap-1 border border-slate-700 cursor-pointer"
                >
                  <ChevronLeft className="w-3.5 h-3.5" />
                  <span>Anterior</span>
                </button>

                <button
                  onClick={() => setChartCenterDate(new Date())}
                  className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md shadow-emerald-950 cursor-pointer"
                >
                  <Crosshair className="w-3.5 h-3.5" />
                  <span>Hoy (Centrar)</span>
                </button>

                <button
                  onClick={() => {
                    const shift = Math.max(7, Math.floor(chartDaysSpan / 2));
                    const d = new Date(chartCenterDate);
                    d.setDate(d.getDate() + shift);
                    setChartCenterDate(d);
                  }}
                  className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-xs font-semibold flex items-center gap-1 border border-slate-700 cursor-pointer"
                >
                  <span>Siguiente</span>
                  <ChevronRight className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>

            {/* Resumen numérico */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs">
              <div className="p-3 rounded-xl bg-slate-950/60 border border-slate-800 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <span className="w-3 h-3 rounded-full bg-emerald-500 shrink-0" />
                  <span className="text-slate-300">Cobrado (Finalizado):</span>
                </div>
                <span className="font-bold text-emerald-400">{formatCurrency(mountainChartData.sumCobrado)}</span>
              </div>

              <div className="p-3 rounded-xl bg-slate-950/60 border border-slate-800 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <span className="w-3 h-3 rounded-full bg-amber-400 shrink-0" />
                  <span className="text-slate-300">Pendiente Proyectado:</span>
                </div>
                <span className="font-bold text-amber-400">{formatCurrency(mountainChartData.sumPendiente)}</span>
              </div>

              <div className="p-3 rounded-xl bg-slate-950/60 border border-slate-800 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <span className="w-3 h-3 rounded-full bg-blue-500 shrink-0" />
                  <span className="text-slate-300">Total en Ventana:</span>
                </div>
                <span className="font-bold text-blue-400">{formatCurrency(mountainChartData.sumTotal)}</span>
              </div>
            </div>

            {/* Canvas de la Gráfica */}
            <div className="relative w-full h-[380px] bg-slate-950/80 rounded-2xl p-4 border border-slate-800/80">
              <canvas ref={chartCanvasRef} />
            </div>

            <div className="text-[11px] text-slate-500 text-center">
              * La gráfica agrupa los montos según la Fecha Probable de Pago de cada comprobante. Las facturas con estatus "Finalizado" o "Pagada sin complemento" se muestran en verde esmeralda y lo pendiente en ámbar cálido.
            </div>
          </div>
        </div>
      )}

      {/* ==================== MODAL: CALENDARIO DE COBRANZA ==================== */}
      {showCalendarModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/85 backdrop-blur-md p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-4xl w-full p-6 shadow-2xl space-y-4 my-8">
            <div className="flex items-center justify-between pb-3 border-b border-slate-800">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 bg-emerald-950 border border-emerald-500/40 rounded-xl flex items-center justify-center text-emerald-400">
                  <CalendarIcon className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white">Calendario de Cobranza Mensual</h3>
                  <p className="text-xs text-slate-400">Proyección de cobros por fecha probable de pago</p>
                </div>
              </div>
              <button
                onClick={() => setShowCalendarModal(false)}
                className="text-slate-400 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Navegación mes */}
            <div className="flex items-center justify-between">
              <div className="text-base font-bold text-white capitalize">
                {monthNames[calMonth]} {calYear}
              </div>
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setCalDate(new Date(calYear, calMonth - 1, 1))}
                  className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl cursor-pointer"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>
                <button
                  onClick={() => setCalDate(new Date())}
                  className="px-3 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-semibold cursor-pointer"
                >
                  Hoy
                </button>
                <button
                  onClick={() => setCalDate(new Date(calYear, calMonth + 1, 1))}
                  className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl cursor-pointer"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>

            {/* Cuadrícula */}
            <div className="border border-slate-800 rounded-2xl overflow-hidden">
              <div className="grid grid-cols-7 bg-slate-950 text-slate-400 text-center py-2 text-xs font-semibold border-b border-slate-800">
                <span>Dom</span><span>Lun</span><span>Mar</span><span>Mié</span><span>Jue</span><span>Vie</span><span>Sáb</span>
              </div>
              <div className="grid grid-cols-7 gap-px bg-slate-800">
                {Array.from({ length: firstDayOfWeek }).map((_, i) => (
                  <div key={`empty-${i}`} className="bg-slate-950/40 min-h-[90px] p-1.5 opacity-30" />
                ))}

                {Array.from({ length: daysInCalMonth }).map((_, idx) => {
                  const day = idx + 1;
                  const ymd = `${calYear}-${String(calMonth + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                  const dayInvoices = calendarDayInvoices[ymd] || [];
                  const dayTotal = dayInvoices.reduce((acc, curr) => acc + curr.monto_total, 0);
                  const isToday = ymd === toLocalDateString(new Date());

                  return (
                    <div
                      key={`day-${day}`}
                      className={`bg-slate-900 min-h-[90px] p-2 flex flex-col justify-between hover:bg-slate-850 transition-colors border border-slate-800/40 ${isToday ? 'ring-1 ring-emerald-500 bg-slate-900/90' : ''}`}
                    >
                      <div className="flex items-center justify-between">
                        <span className={`text-xs font-bold ${isToday ? 'bg-emerald-500 text-slate-950 px-1.5 py-0.5 rounded-md' : 'text-slate-400'} border border-transparent`}>
                          {day}
                        </span>
                        {dayTotal > 0 && (
                          <span className="text-[10px] font-black text-emerald-400">
                            {formatCurrency(dayTotal)}
                          </span>
                        )}
                      </div>

                      <div className="space-y-1 mt-1">
                        {dayInvoices.slice(0, 2).map((inv) => (
                          <div
                            key={inv.id}
                            className="text-[10px] truncate px-1.5 py-0.5 rounded bg-slate-800 border border-slate-700 text-slate-300 font-medium"
                            title={`${inv.numero_factura} - ${inv.empresa} (${formatCurrency(inv.monto_total)})`}
                          >
                            <span className="font-bold text-white">{inv.numero_factura}</span>: {inv.empresa}
                          </div>
                        ))}
                        {dayInvoices.length > 2 && (
                          <div className="text-[9px] text-slate-400 text-right">
                            +{dayInvoices.length - 2} más
                          </div>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MODAL: DIRECTORIO DE CLIENTES ==================== */}
      {showClientsModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/85 backdrop-blur-md p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-4xl w-full p-6 shadow-2xl space-y-4 my-8">
            <div className="flex items-center justify-between pb-3 border-b border-slate-800">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 bg-blue-950 border border-blue-500/40 rounded-xl flex items-center justify-center text-blue-400">
                  <Users className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white">Directorio de Clientes y Condiciones de Crédito</h3>
                  <p className="text-xs text-slate-400">Configuración de días de crédito y reglas de día fijo de pago</p>
                </div>
              </div>
              <button
                onClick={() => setShowClientsModal(false)}
                className="text-slate-400 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Formulario nuevo cliente */}
            <form onSubmit={handleAddClient} className="grid grid-cols-1 sm:grid-cols-5 gap-2 bg-slate-950 p-3 rounded-2xl border border-slate-800">
              <input
                type="text"
                placeholder="Nombre o Razón Social"
                value={newClientName}
                onChange={(e) => setNewClientName(e.target.value)}
                required
                className="sm:col-span-2 bg-slate-900 border border-slate-700 text-xs text-white rounded-xl px-3 py-2"
              />
              <input
                type="text"
                placeholder="RFC (Opcional)"
                value={newClientRfc}
                onChange={(e) => setNewClientRfc(e.target.value)}
                className="bg-slate-900 border border-slate-700 text-xs text-white rounded-xl px-3 py-2"
              />
              <input
                type="number"
                placeholder="Días Crédito"
                value={newClientDays}
                onChange={(e) => setNewClientDays(Number(e.target.value))}
                min="0"
                className="bg-slate-900 border border-slate-700 text-xs text-white rounded-xl px-3 py-2"
              />
              <button
                type="submit"
                className="bg-blue-600 hover:bg-blue-500 text-white text-xs font-bold rounded-xl py-2 flex items-center justify-center gap-1.5 cursor-pointer shadow-md shadow-blue-950"
              >
                <Plus className="w-4 h-4" />
                <span>Agregar</span>
              </button>
            </form>

            {/* Tabla de clientes */}
            <div className="border border-slate-800 rounded-2xl overflow-hidden max-h-96 overflow-y-auto">
              <table className="w-full text-left text-xs">
                <thead className="bg-slate-950 text-slate-400 font-semibold sticky top-0 border-b border-slate-800">
                  <tr>
                    <th className="py-2.5 px-3">Cliente / Razón Social</th>
                    <th className="py-2.5 px-3">RFC</th>
                    <th className="py-2.5 px-3 w-28">Días de Crédito</th>
                    <th className="py-2.5 px-3 min-w-[160px]">Día de Pago Fijo</th>
                    <th className="py-2.5 px-3">Acciones</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800">
                  {clients.length === 0 ? (
                    <tr>
                      <td colSpan={5} className="py-6 text-center text-slate-500">
                        No hay clientes registrados en el directorio.
                      </td>
                    </tr>
                  ) : (
                    clients.map((cli) => (
                      <tr key={cli.id} className="hover:bg-slate-850">
                        <td className="py-2.5 px-3 font-semibold text-slate-200">{cli.nombre}</td>
                        <td className="py-2.5 px-3 font-mono text-slate-400">{cli.rfc || '-'}</td>
                        <td className="py-2.5 px-3">
                          <input
                            type="number"
                            defaultValue={cli.dias_credito}
                            min="0"
                            onBlur={async (e) => {
                              await updateDoc(doc(db, 'clientes', cli.id), {
                                dias_credito: parseInt(e.target.value, 10) || 30,
                                updatedAt: new Date().toISOString()
                              });
                            }}
                            className="w-20 bg-slate-950 border border-slate-700 rounded-lg px-2 py-1 text-xs text-white text-center focus:outline-none focus:ring-1 focus:ring-blue-500"
                          />
                        </td>
                        <td className="py-2.5 px-3">
                          <select
                            defaultValue={cli.dia_pago_fijo}
                            onChange={async (e) => {
                              await updateDoc(doc(db, 'clientes', cli.id), {
                                dia_pago_fijo: e.target.value as DiaPagoFijo,
                                updatedAt: new Date().toISOString()
                              });
                            }}
                            className="bg-slate-950 border border-slate-700 rounded-lg px-2 py-1 text-xs text-slate-200 focus:outline-none focus:ring-1 focus:ring-blue-500 cursor-pointer"
                          >
                            {DIA_PAGO_FIJO_OPTIONS.map((opt) => (
                              <option key={opt} value={opt}>
                                {opt}
                              </option>
                            ))}
                          </select>
                        </td>
                        <td className="py-2.5 px-3 text-slate-400 text-xs">
                          <button
                            onClick={() => handleRecalculateClient(cli)}
                            className="px-2 py-1 bg-blue-950 hover:bg-blue-900 border border-blue-800 text-blue-300 rounded text-[11px] font-medium cursor-pointer"
                          >
                            Recalcular Fechas
                          </button>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* ==================== MODAL: IMPORTADOR CSV ANUAL ==================== */}
      {showCsvModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/85 backdrop-blur-md p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-xl w-full p-6 shadow-2xl space-y-4 my-8">
            <div className="flex items-center justify-between pb-3 border-b border-slate-800">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 bg-amber-950 border border-amber-500/40 rounded-xl flex items-center justify-center text-amber-400">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white">Importar Catálogo Anual de Facturas (CSV)</h3>
                  <p className="text-xs text-slate-400">Carga masiva sin consumir cuotas de API de IA</p>
                </div>
              </div>
              <button
                onClick={() => setShowCsvModal(false)}
                className="text-slate-400 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-4">
              <div className="p-4 bg-slate-950 border border-slate-800 rounded-2xl space-y-2 text-xs text-slate-300">
                <div className="font-bold text-emerald-400 flex items-center gap-1.5">
                  <Check className="w-4 h-4" />
                  Reglas de Transformación Automática:
                </div>
                <ul className="list-disc list-inside space-y-1 text-slate-400 text-[11px]">
                  <li>Columnas admitidas: <strong>Folio, Fecha, Receptor, Concepto, Total, Estatus</strong>.</li>
                  <li>Conversión inteligente de fechas numéricas de Excel y formatos DD/MM/AAAA.</li>
                  <li>Si el <strong>Total es $0.00</strong>, se asigna automáticamente como <strong>Complemento</strong> y se excluye de cobros pendientes.</li>
                  <li>Si el folio ya existe en Firestore, se actualiza automáticamente previniendo duplicados.</li>
                </ul>
              </div>

              <div
                onClick={() => csvInputRef.current?.click()}
                className="border-2 border-dashed border-slate-700 hover:border-amber-400 rounded-2xl p-8 text-center cursor-pointer bg-slate-950/50 hover:bg-slate-950 transition-all flex flex-col items-center justify-center gap-3"
              >
                <div className="w-12 h-12 bg-amber-950/60 rounded-xl flex items-center justify-center text-amber-400">
                  <UploadCloud className="w-6 h-6" />
                </div>
                <div className="text-sm font-semibold text-white">Haz clic aquí o arrastra tu archivo CSV o Excel</div>
                <div className="text-[11px] text-slate-400">Archivos separados por comas (.csv) o libros (.xlsx, .xls) exportados del SAT o ERP</div>
              </div>

              {csvLoading && (
                <div className="p-3 bg-slate-950 border border-emerald-500/40 rounded-xl flex items-center gap-3 text-xs text-emerald-400">
                  <Loader2 className="w-4 h-4 animate-spin text-emerald-400" />
                  <span>{csvStatusText}</span>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ==================== MODAL: FILTRAR POR EMPRESA ==================== */}
      {showEmpresaFilterModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/85 backdrop-blur-md p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-2xl w-full p-6 shadow-2xl space-y-4 my-8 max-h-[88vh] flex flex-col">
            <div className="flex items-center justify-between pb-3 border-b border-slate-800 shrink-0">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 bg-cyan-950 border border-cyan-500/40 rounded-xl flex items-center justify-center text-cyan-400">
                  <Building2 className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white">Filtrar por Empresa</h3>
                  <p className="text-xs text-slate-400">
                    Selecciona una empresa receptora para filtrar la tabla de facturas y complementos
                  </p>
                </div>
              </div>
              <button
                onClick={() => {
                  setShowEmpresaFilterModal(false);
                  setEmpresaSearchQuery('');
                }}
                className="text-slate-400 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Buscador de empresas */}
            <div className="flex items-center gap-2.5 shrink-0">
              <div className="relative flex-1">
                <input
                  type="text"
                  value={empresaSearchQuery}
                  onChange={(e) => setEmpresaSearchQuery(e.target.value)}
                  placeholder="Buscar empresa por nombre o RFC..."
                  className="w-full bg-slate-950 border border-slate-700 text-xs text-white rounded-xl pl-9 pr-3 py-2.5 focus:outline-none focus:ring-2 focus:ring-cyan-500 placeholder:text-slate-500"
                  autoFocus
                />
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-3" />
              </div>

              {tableEmpresaFilter !== 'todas' && (
                <button
                  type="button"
                  onClick={() => {
                    setTableEmpresaFilter('todas');
                    setShowEmpresaFilterModal(false);
                    setEmpresaSearchQuery('');
                  }}
                  className="px-3 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-xl text-xs font-semibold border border-slate-700 transition-all cursor-pointer whitespace-nowrap"
                >
                  Mostrar todas
                </button>
              )}
            </div>

            {/* Listado de empresas */}
            <div className="overflow-y-auto flex-1 space-y-2 pr-1 min-h-[220px]">
              {filteredEmpresaStats.length === 0 ? (
                <div className="py-12 text-center text-slate-500 text-xs">
                  No se encontraron empresas con el término buscado.
                </div>
              ) : (
                filteredEmpresaStats.map((emp) => {
                  const isSelected = tableEmpresaFilter.toLowerCase() === emp.nombre.toLowerCase();
                  return (
                    <div
                      key={emp.nombre}
                      onClick={() => {
                        setTableEmpresaFilter(emp.nombre);
                        setShowEmpresaFilterModal(false);
                        setEmpresaSearchQuery('');
                      }}
                      className={`p-3.5 rounded-2xl border transition-all cursor-pointer flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${
                        isSelected
                          ? 'bg-cyan-950/60 border-cyan-500 text-white shadow-md shadow-cyan-950/40 ring-1 ring-cyan-500/50'
                          : 'bg-slate-950/60 border-slate-800 hover:border-slate-700 hover:bg-slate-800/40 text-slate-200'
                      }`}
                    >
                      <div className="space-y-1">
                        <div className="flex items-center gap-2">
                          <span className="font-bold text-sm text-white">{emp.nombre}</span>
                          {isSelected && (
                            <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-cyan-500/20 text-cyan-300 border border-cyan-500/30">
                              Activa
                            </span>
                          )}
                        </div>
                        {emp.rfc && (
                          <div className="text-xs font-mono text-slate-400">RFC: {emp.rfc}</div>
                        )}
                      </div>

                      <div className="flex items-center gap-4 text-xs">
                        <div className="text-right">
                          <div className="text-slate-400 text-[11px]">{emp.totalFacturas} comprobantes</div>
                          <div className="font-bold text-emerald-400">{formatCurrency(emp.montoTotal)}</div>
                        </div>

                        <button
                          type="button"
                          onClick={(e) => {
                            e.stopPropagation();
                            setTableEmpresaFilter(emp.nombre);
                            setShowEmpresaFilterModal(false);
                            setEmpresaSearchQuery('');
                          }}
                          className={`px-3 py-1.5 rounded-xl font-semibold text-xs transition-all cursor-pointer ${
                            isSelected
                              ? 'bg-cyan-500 text-slate-950 font-bold'
                              : 'bg-slate-800 hover:bg-cyan-600 hover:text-white text-slate-300'
                          }`}
                        >
                          {isSelected ? 'Seleccionada' : 'Filtrar'}
                        </button>
                      </div>
                    </div>
                  );
                })
              )}
            </div>

            <div className="pt-2 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400 shrink-0">
              <span>Total de empresas: <strong>{empresaStats.length}</strong></span>
              <button
                type="button"
                onClick={() => {
                  setShowEmpresaFilterModal(false);
                  setEmpresaSearchQuery('');
                }}
                className="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl cursor-pointer"
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: ATENCIÓN URGENTE (COMPLEMENTOS PENDIENTES) */}
      {showUrgentModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/85 backdrop-blur-md p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-3xl w-full p-6 shadow-2xl space-y-4 my-8 max-h-[88vh] flex flex-col">
            <div className="flex items-center justify-between pb-3 border-b border-slate-800 shrink-0">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 bg-amber-950 border border-amber-500/40 rounded-xl flex items-center justify-center text-amber-400">
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white flex items-center gap-2">
                    <span>Atención Urgente: Complementos Pendientes</span>
                    <span className="px-2 py-0.5 bg-amber-500 text-slate-950 rounded-full text-xs font-black">
                      {urgentInvoices.length}
                    </span>
                  </h3>
                  <p className="text-xs text-slate-400">
                    Facturas pagadas o con fecha vencida que requieren emisión de complemento de pago CFDI
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowUrgentModal(false)}
                className="text-slate-400 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Lista de facturas urgentes */}
            <div className="overflow-y-auto flex-1 space-y-2 pr-1 min-h-[220px]">
              {urgentInvoices.length === 0 ? (
                <div className="py-12 text-center text-slate-500 flex flex-col items-center justify-center gap-2">
                  <CheckCircle2 className="w-10 h-10 text-emerald-500/60 mb-1" />
                  <p className="text-sm font-semibold text-slate-300">¡Al día! No hay complementos pendientes</p>
                  <p className="text-xs text-slate-500">
                    Todas las facturas cobradas cuentan con su complemento CFDI o fueron descartadas.
                  </p>
                </div>
              ) : (
                urgentInvoices.map((inv) => (
                  <div
                    key={inv.id}
                    className="flex items-center justify-between p-3 rounded-2xl bg-slate-950/70 border border-slate-800 hover:border-amber-500/40 transition-all gap-3"
                  >
                    <div className="flex items-center gap-3 min-w-0">
                      <span className="font-bold text-white px-2.5 py-1 bg-slate-800 rounded-lg border border-slate-700/80 text-xs font-mono shrink-0">
                        {inv.numero_factura || 'S/N'}
                      </span>
                      <div className="min-w-0">
                        <div className="text-xs font-bold text-slate-200 truncate" title={inv.empresa}>
                          {inv.empresa || 'Cliente sin nombre'}
                        </div>
                        <div className="text-[11px] text-slate-400 flex items-center gap-2 mt-0.5">
                          <span>
                            Pago: <strong className="text-amber-300 font-mono">{inv.fecha_probable_pago || 'Vencida'}</strong>
                          </span>
                          <span>•</span>
                          <span className="text-slate-500 font-mono">
                            {inv.orden_de_compra ? `OC: ${inv.orden_de_compra}` : 'Sin OC'}
                          </span>
                        </div>
                      </div>
                    </div>

                    <div className="flex items-center gap-3 shrink-0">
                      <div className="text-right">
                        <div className="text-xs font-black font-mono text-emerald-400">
                          {formatCurrency(inv.monto_total)}
                        </div>
                        <span className="inline-block px-1.5 py-0.5 rounded text-[10px] font-semibold bg-amber-950/80 text-amber-300 border border-amber-800/60 mt-0.5">
                          {inv.estatus}
                        </span>
                      </div>

                      <button
                        type="button"
                        onClick={() => dismissUrgentAlert(inv.id)}
                        className="p-1.5 text-slate-500 hover:text-rose-400 hover:bg-slate-800 rounded-lg transition-colors cursor-pointer"
                        title="Descartar aviso para esta factura"
                      >
                        <X className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>

            <div className="pt-3 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400 shrink-0">
              <span className="font-medium text-slate-300">
                Monto total pendiente:{' '}
                {formatCurrency(urgentInvoices.reduce((acc, curr) => acc + curr.monto_total, 0))}
              </span>
              <button
                type="button"
                onClick={() => setShowUrgentModal(false)}
                className="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl cursor-pointer"
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: FACTURAS CON PROBLEMA */}
      {showProblemaModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-950/85 backdrop-blur-md p-4 overflow-y-auto">
          <div className="bg-slate-900 border border-slate-800 rounded-3xl max-w-4xl w-full p-6 shadow-2xl space-y-4 my-8 max-h-[88vh] flex flex-col">
            <div className="flex items-center justify-between pb-3 border-b border-slate-800 shrink-0">
              <div className="flex items-center gap-2.5">
                <div className="w-9 h-9 bg-rose-950 border border-rose-500/40 rounded-xl flex items-center justify-center text-rose-400">
                  <AlertOctagon className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white flex items-center gap-2">
                    <span>Revisión de Facturas con Problema</span>
                    <span className="px-2 py-0.5 bg-rose-600 text-white rounded-full text-xs font-black">
                      {problemaInvoices.length}
                    </span>
                  </h3>
                  <p className="text-xs text-slate-400">
                    Comprobantes retenidos, aclaraciones pendientes o incidencias operativas
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setShowProblemaModal(false)}
                className="text-slate-400 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Lista de facturas con problema */}
            <div className="overflow-y-auto flex-1 space-y-2 pr-1 min-h-[220px]">
              {problemaInvoices.length === 0 ? (
                <div className="py-12 text-center text-slate-500 flex flex-col items-center justify-center gap-2">
                  <CheckCircle2 className="w-10 h-10 text-emerald-500/60 mb-1" />
                  <p className="text-sm font-semibold text-slate-300">¡Sin incidencias activas!</p>
                  <p className="text-xs text-slate-500">
                    No hay facturas registradas con el estatus "Problema".
                  </p>
                </div>
              ) : (
                problemaInvoices.map((inv) => {
                  const notaOrConcepto = inv.concepto || inv.complemento || 'Sin notas registradas';
                  return (
                    <div
                      key={inv.id}
                      className="flex flex-col sm:flex-row sm:items-center justify-between p-3.5 rounded-2xl bg-slate-950/70 border border-slate-800 hover:border-rose-500/40 transition-all gap-3"
                    >
                      <div className="flex items-start gap-3 min-w-0 flex-1">
                        <span className="font-bold text-white px-2.5 py-1 bg-slate-800 rounded-lg border border-slate-700/80 text-xs font-mono shrink-0">
                          {inv.numero_factura || 'S/N'}
                        </span>
                        <div className="min-w-0 flex-1">
                          <div className="text-xs font-bold text-slate-200 truncate" title={inv.empresa}>
                            {inv.empresa || 'Cliente sin nombre'}
                          </div>
                          <div className="text-[11px] text-slate-400 flex flex-wrap items-center gap-2 mt-0.5">
                            <span>
                              Pago: <strong className="text-amber-300 font-mono">{inv.fecha_probable_pago || 'Sin fecha'}</strong>
                            </span>
                            <span>•</span>
                            <span className="text-slate-500 font-mono">
                              {inv.orden_de_compra ? `OC: ${inv.orden_de_compra}` : 'Sin OC'}
                            </span>
                          </div>
                          <div className="mt-1 text-[11px] text-rose-300/90 bg-rose-950/40 border border-rose-900/40 rounded-lg px-2 py-1 line-clamp-2" title={notaOrConcepto}>
                            <span className="font-semibold text-rose-400">Nota/Concepto:</span> {notaOrConcepto}
                          </div>
                        </div>
                      </div>

                      <div className="flex items-center justify-between sm:justify-end gap-3 shrink-0 pt-2 sm:pt-0 border-t sm:border-t-0 border-slate-800/80">
                        <div className="text-left sm:text-right">
                          <div className="text-xs font-black font-mono text-emerald-400">
                            {formatCurrency(inv.monto_total)}
                          </div>
                          <select
                            value={inv.estatus}
                            onChange={(e) => updateInvoiceField(inv.id, 'estatus', e.target.value)}
                            className="mt-1 bg-slate-900 border border-rose-500/40 text-[10px] text-rose-200 font-semibold rounded-lg px-2 py-0.5 cursor-pointer focus:outline-none focus:ring-1 focus:ring-rose-500"
                            title="Cambiar estatus si se resolvió el problema"
                          >
                            {INVOICE_STATUS_OPTIONS.map((st) => (
                              <option key={st} value={st}>
                                {st}
                              </option>
                            ))}
                          </select>
                        </div>

                        <button
                          type="button"
                          onClick={() => {
                            setShowProblemaModal(false);
                            setSearchQuery(inv.numero_factura || '');
                          }}
                          className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white rounded-xl text-xs font-semibold flex items-center gap-1.5 transition-colors cursor-pointer shrink-0"
                          title="Ver y filtrar en la tabla"
                        >
                          <Search className="w-3.5 h-3.5 text-cyan-400" />
                          <span className="hidden sm:inline">Ver en tabla</span>
                        </button>
                      </div>
                    </div>
                  );
                })
              )}
            </div>

            <div className="pt-3 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400 shrink-0">
              <span className="font-medium text-slate-300">
                Monto total con problema:{' '}
                {formatCurrency(problemaInvoices.reduce((acc, curr) => acc + curr.monto_total, 0))}
              </span>
              <button
                type="button"
                onClick={() => setShowProblemaModal(false)}
                className="px-4 py-1.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl cursor-pointer"
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
