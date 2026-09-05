import { useState, useMemo, useRef, useEffect } from 'react';
import { useInventario } from '../context/InventarioContext';
import { useSolicitudes } from '../context/SolicitudesContext';
import { useAuth } from '../context/AuthContext';
import { useAlert } from '../context/AlertContext';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { Download, Search, Package, UserCircle, MonitorSmartphone, Printer, Eye, Upload, Pencil, CheckCircle, UploadCloud, AlertCircle, FileWarning, AlertTriangle, PlusCircle, UserPlus, QrCode, FileText, Info } from 'lucide-react';
import { QRCodeSVG } from 'qrcode.react';
import { useSearchParams } from 'react-router-dom';
import { saveDocument, getDocument } from '../utils/db';
import EditarEquipoModal from '../components/EditarEquipoModal';
import NuevoEquipoModal from '../components/NuevoEquipoModal';
import { isSameUser } from '../utils/userUtils';
import AutocompleteInput from '../components/AutocompleteInput';
import { supabase } from '../lib/supabaseClient';
import { generateActaDocx } from '../utils/docxUtils';
import { getActaFirmadaUrl } from '../utils/storageUtils';
import { encodeQRData } from '../utils/cryptoUtils';
import Badge from '../components/Badge';

const COLUMNS = [
  'Descripción del Bien', 'Marca', 'Modelo', 'Nº de serie',
  'ID Publicación',
  'Orden de Compra', 'Factura', 'Proveedor', 'SubDirección', 'Usuario'
];

function norm(s) {
  return (s == null ? '' : String(s)).trim().toLowerCase().replace(/\s+/g, ' ');
}

const HEADER_ALIASES = {
  'Descripción del Bien': ['descripción del bien', 'descripcion del bien', 'descripcion bien', 'descripción bien'],
  'Marca': ['marca'],
  'Modelo': ['modelo'],
  'Nº de serie': ['nº de serie', 'n° de serie', 'no de serie', 'numero de serie', 'número de serie', 'serie'],
  'ID Publicación': ['id publicación', 'id publicacion', 'codigo compra agil / licitación / codigo convenio marco', 'codigo compra agil / licitacion / codigo convenio marco', 'código compra ágil', 'codigo compra', 'licitacion', 'licitación', 'convenio marco', 'codigo'],
  'Tipo Publicación': ['tipo publicación', 'tipo publicacion', 'tipo'],
  'Orden de Compra': ['orden de compra', 'oc', 'orden compra'],
  'Factura': ['factura', 'n° factura', 'nº factura'],
  'Proveedor': ['proveedor', 'empresa', 'vendedor', 'distribuidor'],
  'SubDirección': ['subcdirección', 'subcdireccion', 'subdirección', 'subdireccion', 'departamento', 'area', 'área'],
  'Usuario': ['usuario', 'funcionario', 'asignado a', 'responsable']
};

function normalizeRow(rawRow) {
  const out = {};
  const lowerMap = {};
  Object.keys(rawRow).forEach(k => lowerMap[norm(k)] = rawRow[k]);
  COLUMNS.forEach(canonical => {
    const aliases = HEADER_ALIASES[canonical] || [norm(canonical)];
    let found = '';
    for (const a of aliases) {
      if (lowerMap[a] !== undefined && lowerMap[a] !== null && String(lowerMap[a]).trim() !== '') {
        found = lowerMap[a]; break;
      }
    }
    if (found === '' && lowerMap[norm(canonical)] !== undefined) found = lowerMap[norm(canonical)];
    out[canonical] = found == null ? '' : String(found).trim();
  });
  if (lowerMap['_id_interno_'] !== undefined) out['_id_interno_'] = String(lowerMap['_id_interno_']).trim();
  
  const tipoAliases = HEADER_ALIASES['Tipo Publicación'];
  let foundTipo = '';
  for (const a of tipoAliases) {
    if (lowerMap[a] !== undefined && lowerMap[a] !== null && String(lowerMap[a]).trim() !== '') {
      foundTipo = lowerMap[a]; break;
    }
  }
  if (foundTipo === '' && lowerMap[norm('Tipo Publicación')] !== undefined) foundTipo = lowerMap[norm('Tipo Publicación')];
  out['Tipo Publicación'] = foundTipo == null ? '' : String(foundTipo).trim();

  return out;
}

function isAvailable(usuario) {
  const v = norm(usuario);
  return v === '' || v === 'disponible';
}

function getEstadoFinal(row) {
  if (!row) return 'DISPONIBLE';
  const isDisp = isAvailable(row['Usuario']);
  const dbEstado = (row.estado || '').trim().toUpperCase();

  if (!isDisp) {
    if (dbEstado === 'EN PRESTAMO' || dbEstado === 'EN PRÉSTAMO') {
      return 'EN PRESTAMO';
    }
    if (dbEstado === 'BAJA' || dbEstado === 'DE BAJA') {
      return 'DE BAJA';
    }
    return 'ASIGNADO';
  } else {
    if (dbEstado === 'ASIGNADO') {
      return 'DISPONIBLE';
    }
    if (dbEstado === 'EN PRESTAMO' || dbEstado === 'EN PRÉSTAMO') {
      return 'EN PRESTAMO';
    }
    if (dbEstado === 'BAJA' || dbEstado === 'DE BAJA') {
      return 'DE BAJA';
    }
    if (dbEstado === 'PARA PRESTAMO' || dbEstado === 'PARA PRÉSTAMO') {
      return 'PARA PRESTAMO';
    }
    return dbEstado || 'DISPONIBLE';
  }
}

function safe(v) {
  return (v == null || String(v).trim() === '') ? '—' : String(v).trim();
}

function getInitials(name) {
  if (!name || name === '—') return '??';
  const words = String(name).trim().split(/\s+/);
  if (words.length === 1) return words[0].substring(0, 2).toUpperCase();
  return (words[0][0] + words[words.length - 1][0]).toUpperCase();
}

const isQRSupported = (tipo) => {
  const t = (tipo || '').toLowerCase();
  return (
    t.includes('notebook') ||
    t.includes('aio') ||
    t.includes('tablet') ||
    t.includes('all in one') ||
    t.includes('todo en uno') ||
    t.includes('impresora') ||
    t.includes('switch') ||
    t.includes('router') ||
    t.includes('monitor') ||
    t.includes('proyector') ||
    t.includes('dron') ||
    t.includes('drone') ||
    t.includes('dock') ||
    t.includes('camara') ||
    t.includes('cámara') ||
    t.includes('tv')
  );
};

export default function DashboardPage() {
  const { equipos, loading, setFileStatus, addMasivo, updateEquipo } = useInventario();
  const { solicitudes } = useSolicitudes();
  const { user, session, perfil, canEdit } = useAuth();
  const { showAlertConfirm, showAlertPrompt } = useAlert();
  const [searchParams, setSearchParams] = useSearchParams();
  const activeTab = searchParams.get('tab') || 'disp';
  const setActiveTab = (tab) => {
    setSearchParams({ tab });
  };
  const [globalSearch, setGlobalSearch] = useState(() => searchParams.get('q') || searchParams.get('search') || '');
  const [sortConfig, setSortConfig] = useState({ col: 'Descripción del Bien', dir: 1 });

  const [isMasivaModalOpen, setIsMasivaModalOpen] = useState(false);
  const [isNuevoEquipoModalOpen, setIsNuevoEquipoModalOpen] = useState(false);
  const [editingEquipo, setEditingEquipo] = useState(null);
  const [status, setStatus] = useState({ type: 'idle', message: '' });
  const [localToast, setLocalToast] = useState(null);
  const toastTimerRef = useRef(null);
  const [toastPos, setToastPos] = useState({ x: 0, y: 0 });
  const isDragging = useRef(false);
  const dragStart = useRef({ x: 0, y: 0 });

  const [assignModalData, setAssignModalData] = useState(null);
  const [assignUserName, setAssignUserName] = useState('');
  const [assignDate, setAssignDate] = useState('');
  const [assignObservation, setAssignObservation] = useState('');
  const [qrModalData, setQrModalData] = useState(null);
  const [perfilesOptions, setPerfilesOptions] = useState([]);
  const [perfilesData, setPerfilesData] = useState([]);

  useEffect(() => {
    async function fetchPerfiles() {
      const { data } = await supabase.from('perfiles').select('*');
      if (data) {
        setPerfilesData(data);
        const ops = new Set(data.map(p => p.nombre || p.email).filter(Boolean));
        setPerfilesOptions(Array.from(ops).sort());
      }
    }
    fetchPerfiles();
  }, []);

  useEffect(() => {
    const handleGlobalMouseMove = (e) => {
      if (!isDragging.current) return;
      setToastPos({
        x: e.clientX - dragStart.current.x,
        y: e.clientY - dragStart.current.y
      });
    };

    const handleGlobalMouseUp = () => {
      if (isDragging.current) {
        isDragging.current = false;
        document.body.style.userSelect = '';
      }
    };

    if (localToast) {
      window.addEventListener('mousemove', handleGlobalMouseMove);
      window.addEventListener('mouseup', handleGlobalMouseUp);
    }
    return () => {
      window.removeEventListener('mousemove', handleGlobalMouseMove);
      window.removeEventListener('mouseup', handleGlobalMouseUp);
    };
  }, [localToast]);

  const startToastTimer = (title, duration) => {
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    toastTimerRef.current = setTimeout(() => {
      setLocalToast(prev => (prev && prev.title === title ? null : prev));
    }, duration);
  };

  const showLocalToast = (title, message, type = 'success', details = null) => {
    const duration = type === 'warning' ? 15000 : 8000;
    setLocalToast({ title, message, type, ...details });
    setToastPos({ x: 0, y: 0 });
    startToastTimer(title, duration);
  };

  const handleToastMouseEnter = () => {
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
  };

  const handleToastMouseLeave = () => {
    if (localToast) {
      startToastTimer(localToast.title, 4000);
    }
  };

  const handleVerActa = async (path) => {
    if (!path) return;
    const url = await getActaFirmadaUrl(path);
    if (url) {
      window.open(url, '_blank');
    } else {
      showLocalToast('Error', 'No se pudo abrir el acta', 'error');
    }
  };

  const handleGenerateActaAsignacion = async (row) => {
    try {
      let adminName = perfil?.nombre || 'Administrador TI';
      let adminRut = perfil?.rut || '—';
      let adminSub = perfil?.subdireccion || 'Tecnologías de la Información';

      const dbEstado = (row.estado || '').trim().toUpperCase();
      const isEnPrestamo = dbEstado === 'EN PRESTAMO' || dbEstado === 'EN PRÉSTAMO';

      let activeLoan = null;
      let userName = row['Usuario'] || '—';
      let userRut = '—';
      let userSub = '—';

      if (isEnPrestamo) {
        // Buscar el préstamo activo en solicitudes
        activeLoan = solicitudes?.find(s => (s.equipo_id === row.id || s.equipo_id === row['Nº de serie']) && s.estado === 'aprobado' && s.tipo === 'prestamo');
        
        if (activeLoan) {
          // Extraer RUT y subdirección del aprobador si es posible
          const adminMatch = activeLoan.observaciones_admin ? activeLoan.observaciones_admin.match(/\[Aprobado por:\s*(.*?)\]/) : null;
          const approvedByName = adminMatch ? adminMatch[1].trim() : null;
          if (approvedByName) {
             const approver = perfilesData.find(p => p.nombre === approvedByName || p.email === approvedByName);
             if (approver) {
               adminName = approver.nombre || approvedByName;
               adminRut = approver.rut || '—';
               adminSub = approver.subdireccion || 'Tecnologías de la Información';
             } else {
               adminName = approvedByName;
             }
          }
          
          userName = activeLoan.perfil?.nombre || userName;
          userRut = activeLoan.perfil?.rut || '—';
          userSub = activeLoan.perfil?.subdireccion || '—';
        }
      } else {
        const userProfile = perfilesData.find(p => p.nombre === userName || p.email === userName) || {};
        userRut = userProfile.rut || '—';
        userSub = userProfile.subdireccion || '—';
      }

      const templateName = isEnPrestamo ? 'acta_prestamo.docx' : 'acta_asigna.docx';

      let fechaEntrega = new Date().toLocaleDateString();
      let dia = '';
      let mes = '';
      let ano = '';

      if (!isEnPrestamo) {
        const dateToUse = row.fecha_asignacion ? row.fecha_asignacion : new Date().toISOString().split('T')[0];
        const dateParts = dateToUse.split('-');
        if (dateParts.length === 3) {
          const [y, m, d] = dateParts;
          const dateObj = new Date(y, m - 1, d);
          fechaEntrega = dateObj.toLocaleDateString();
          dia = dateObj.getDate().toString().padStart(2, '0');
          const monthNames = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
          mes = monthNames[dateObj.getMonth()];
          ano = dateObj.getFullYear().toString();
        }
      }

      const data = {
        ti_nombre: adminName,
        ti_rut: adminRut,
        ti_subdireccion: adminSub,
        solicitante_nombre: userName,
        solicitante_rut: userRut,
        solicitante_subdireccion: userSub,
        fecha_entrega: fechaEntrega,
        dia: dia,
        día: dia,
        DIA: dia,
        DÍA: dia,
        Día: dia,
        mes: mes,
        año: ano,
        fecha_inicio: activeLoan?.fecha_inicio || '',
        fecha_fin: activeLoan?.fecha_fin || '',
        hora_inicio: activeLoan?.hora_inicio || '',
        hora_fin: activeLoan?.hora_fin || '',
        equipos: [
          {
            tipo: row['Descripción del Bien'] || 'Equipo',
            marca_modelo: `${row.Marca || ''} ${row.Modelo || ''}`.trim(),
            serie: row['Nº de serie'] || '—'
          }
        ]
      };

      const result = await generateActaDocx(data, templateName);
      if (!result.success) {
         showLocalToast('Error', result.error || 'No se pudo generar el acta', 'error');
      }
    } catch(err) {
       console.error(err);
       showLocalToast('Error', 'Hubo un error al crear el acta.', 'error');
    }
  };

  const handleGenerateActaMasivaFuncionario = async () => {
    if (!selectedFunc) {
      alert('Debe seleccionar un funcionario primero.');
      return;
    }
    if (baseData.length === 0) {
      alert('El funcionario no tiene equipos asignados.');
      return;
    }
    
    try {
      const firstEq = baseData[0];
      const userName = firstEq['Usuario'] || selectedFunc;
      const userSub = firstEq['SubDirección'] || '—';

      // Obtener RUT del usuario si existe
      const userProfile = perfilesData.find(p => p.nombre === userName || p.email === userName) || {};
      const userRut = userProfile.rut || '—';

      let adminName = perfil?.nombre || 'Administrador TI';
      let adminRut = perfil?.rut || '—';
      let adminSub = perfil?.subdireccion || 'Tecnologías de la Información';

      const d = new Date();
      const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
      const dia = d.getDate().toString().padStart(2, '0');
      const mes = meses[d.getMonth()];
      const ano = d.getFullYear().toString();

      const allPrestamo = baseData.every(eq => eq.estado === 'EN PRESTAMO' || eq.estado === 'EN PRÉSTAMO');
      const templateName = allPrestamo ? 'acta_prestamo.docx' : 'acta_asigna.docx';

      const data = {
        ti_nombre: adminName,
        ti_rut: adminRut,
        ti_subdireccion: adminSub,
        solicitante_nombre: userName,
        solicitante_rut: userRut,
        solicitante_subdireccion: userSub,
        fecha_entrega: `${dia} de ${mes} de ${ano}`,
        dia: dia,
        día: dia,
        DIA: dia,
        DÍA: dia,
        Día: dia,
        mes: mes,
        año: ano,
        fecha_inicio: '',
        fecha_fin: '',
        hora_inicio: '',
        hora_fin: '',
        equipos: baseData.map(eq => ({
          tipo: eq['Descripción del Bien'] || 'Equipo',
          marca_modelo: `${eq.Marca || ''} ${eq.Modelo || ''}`.trim(),
          serie: eq['Nº de serie'] || '—',
          codigo_interno: eq.id || eq['ID Publicación'] || '',
          estado: eq.estado || '—',
          fecha_asignacion: eq.fecha_asignacion || ''
        }))
      };

      const result = await generateActaDocx(data, templateName);
      if (!result.success) {
         showLocalToast('Error', result.error || 'No se pudo generar el acta', 'error');
      }
    } catch(err) {
       console.error(err);
       showLocalToast('Error', 'Hubo un error al crear el acta.', 'error');
    }
  };

  const handleDragStart = (e) => {
    if (e.target.tagName.toLowerCase() === 'button' || e.target.closest('button')) return;
    isDragging.current = true;
    dragStart.current = {
      x: e.clientX - toastPos.x,
      y: e.clientY - toastPos.y
    };
    document.body.style.userSelect = 'none';
  };

  const processData = (rows) => {
    if (!rows || rows.length === 0) {
      setStatus({ type: 'error', message: 'El archivo está vacío' });
      showLocalToast('Error', 'El archivo cargado está vacío.', 'error');
      return;
    }

    const normalized = rows.map(r => normalizeRow(r)).filter(r => {
      return COLUMNS.some(col => r[col] && r[col].toString().trim() !== '');
    });

    if (normalized.length === 0) {
      setStatus({ type: 'error', message: 'No se encontraron filas válidas.' });
      showLocalToast('Error', 'No se encontraron filas válidas en el archivo.', 'error');
      return;
    }

    const isFirstUpload = equipos.length === 0;
    const noSerialNew = [];
    const importableItems = [];
    const updatedItems = [];
    const newItems = [];

    normalized.forEach((row, index) => {
      let serial = row['Nº de serie'] ? String(row['Nº de serie']).trim().toLowerCase() : '';
      if (serial === '—') {
        serial = '';
        row['Nº de serie'] = '';
      }
      
      const internalId = row['_id_interno_'] ? String(row['_id_interno_']).trim() : null;

      // Check if item exists
      let exists = false;
      if (internalId && equipos.some(e => String(e.id) === internalId)) {
        exists = true;
      } else if (serial && equipos.some(e => (e['Nº de serie'] ? String(e['Nº de serie']).trim().toLowerCase() : '') === serial)) {
        exists = true;
      }

      if (exists) {
        updatedItems.push(row);
        importableItems.push(row);
      } else {
        if (!serial) {
          row['Nº de serie'] = `S/N-${Date.now()}-${Math.floor(Math.random() * 10000) + index}`;
          noSerialNew.push({ rowNumber: index + 2, desc: row['Descripción del Bien'] || 'Sin descripción' });
        }
        newItems.push(row);
        importableItems.push(row);
      }
    });

    if (importableItems.length > 0) {
      addMasivo(importableItems);
      const partialMsg = noSerialNew.length > 0 ? `\n• Se auto-generaron ${noSerialNew.length} Nº de Serie temporales.` : '';
      const message = `Se agregaron ${newItems.length} equipos y se actualizaron ${updatedItems.length}.${partialMsg}`;
      
      setStatus({
        type: 'success',
        message: message
      });
      showLocalToast(
        noSerialNew.length > 0 ? 'Carga con Auto-generación' : 'Carga Exitosa',
        message,
        'success',
        {
          duplicateSerials: [],
          addedCount: newItems.length + updatedItems.length,
          noSerialCount: noSerialNew.length,
          isFirstUpload: isFirstUpload
        }
      );
      setTimeout(() => setIsMasivaModalOpen(false), 2000);
    } else {
      setStatus({
        type: 'error',
        message: 'No se encontraron registros válidos para agregar o actualizar.'
      });
      showLocalToast(
        'Error de Carga',
        'El archivo no contenía registros válidos.',
        'error',
        {
          duplicateSerials: [],
          addedCount: 0,
          noSerialCount: 0,
          isFirstUpload: false
        }
      );
    }
  };

  const masivaFileInputRef = useRef(null);

  const handleMasivaFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setStatus({ type: 'processing', message: 'Procesando archivo...' });
    setLocalToast(null);

    const ext = file.name.split('.').pop().toLowerCase();

    if (ext === 'csv') {
      Papa.parse(file, {
        header: true, skipEmptyLines: true, dynamicTyping: false,
        complete: res => { setTimeout(() => processData(res.data), 50); },
        error: err => setStatus({ type: 'error', message: err.message })
      });
    } else if (ext === 'xlsx' || ext === 'xls') {
      const reader = new FileReader();
      reader.onload = ev => {
        setTimeout(() => {
          try {
            const wb = XLSX.read(ev.target.result, { type: 'array' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
            processData(json);
          } catch (err) {
            setStatus({ type: 'error', message: 'Error al leer el archivo Excel' });
          }
        }, 50);
      };
      reader.onerror = () => setStatus({ type: 'error', message: 'Error de lectura de archivo' });
      reader.readAsArrayBuffer(file);
    } else {
      setStatus({ type: 'error', message: 'Formato no soportado. Use .csv, .xls o .xlsx' });
    }

    if (masivaFileInputRef.current) {
      masivaFileInputRef.current.value = '';
    }
  };

  // Tab: func state
  const [selectedFunc, setSelectedFunc] = useState('');
  const [funcSearch, setFuncSearch] = useState('');
  const [showFuncSug, setShowFuncSug] = useState(false);
  const [focusedIndex, setFocusedIndex] = useState(-1);

  // Tab: equip state
  const [selectedDesc, setSelectedDesc] = useState('');
  const [selectedMod, setSelectedMod] = useState('');

  const [uploadTarget, setUploadTarget] = useState(null);
  const fileInputRef = useRef(null);

  const handlePreview = async (id, type) => {
    // Abrir la pestaña inmediatamente para evitar el bloqueo de popups por el navegador
    const newTab = window.open('', '_blank');
    if (newTab) newTab.document.write('Cargando documento...');

    try {
      const eq = equipos.find(e => e.id === id);
      const rawCode = eq ? (type === 'factura' ? eq['Factura'] : eq['Orden de Compra']) : '';
      const code = rawCode != null ? String(rawCode) : '';
      const defaultStorageKey = (code && code.trim() !== '—' && code.trim() !== '')
        ? `${type}_${code.trim().toLowerCase()}`
        : `${type}_${id}`;

      let doc = await getDocument(defaultStorageKey, type);
      
      // Fallback para documentos antiguos que sufrieron el bug de colisión y se guardaron sin prefijo
      if (!doc && (!code || code.trim() === '—' || code.trim() === '')) {
        doc = await getDocument(id, type);
      }

      if (!doc) {
        if (newTab) newTab.close();
        alert('No se encontró el documento asociado.');
        return;
      }
      const fileURL = URL.createObjectURL(doc.blob);
      if (newTab) {
        newTab.location.href = fileURL;
      } else {
        window.open(fileURL, '_blank');
      }
    } catch (err) {
      if (newTab) newTab.close();
      console.error('Error al abrir el documento:', err);
      alert('Error al abrir el documento.');
    }
  };

  const triggerUpload = (id, type) => {
    setUploadTarget({ id, type });
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };

  const handleDirectUpload = async (e) => {
    const file = e.target.files[0];
    if (!file || !uploadTarget) return;

    try {
      const eq = equipos.find(item => item.id === uploadTarget.id);
      const originalCode = eq ? (uploadTarget.type === 'factura' ? eq['Factura'] : eq['Orden de Compra']) : '';
      let code = originalCode;

      // Extract from filename if code is empty
      if (!code || code.trim() === '—' || code.trim() === '') {
        const cleanName = file.name.replace(/\.[^/.]+$/, "");
        let extractedCode = null;
        
        if (uploadTarget.type === 'factura') {
          const matchFactura = cleanName.match(/(?:n[°º]|nro\.?|numero|número)\s*(\d+)/i);
          if (matchFactura && matchFactura[1]) {
            extractedCode = matchFactura[1];
          } else {
            const matchOld = cleanName.match(/(?:factura|fact|f)[\s_.-]*([a-z0-9-]+)/i);
            if (matchOld && matchOld[1]) extractedCode = matchOld[1].toUpperCase();
          }
        } else {
          const matchOC = cleanName.match(/(1456839-\d+-[a-z]{2}\d{2})/i);
          if (matchOC && matchOC[1]) {
            extractedCode = matchOC[1].toUpperCase();
          } else {
            const matchOld = cleanName.match(/(?:oc|orden|compra)[\s_.-]*([a-z0-9-]+)/i);
            if (matchOld && matchOld[1]) extractedCode = matchOld[1].toUpperCase();
          }
        }

        if (!extractedCode) {
           const numMatch = cleanName.match(/\d{4,}/);
           if (numMatch) extractedCode = numMatch[0];
        }

        if (extractedCode && !['PDF', 'JPG', 'PNG', 'DOC', 'DOCK'].includes(extractedCode)) {
           code = extractedCode;
           // Update DB record
           const fieldName = uploadTarget.type === 'factura' ? 'Factura' : 'Orden de Compra';
           const index = equipos.findIndex(item => item.id === uploadTarget.id);
           if (index !== -1) {
             const updated = { ...eq, [fieldName]: code };
             await updateEquipo(index, updated);
           }
        }
      }

      const storageKey = (code && code.trim() !== '—' && code.trim() !== '')
        ? `${uploadTarget.type}_${code.trim().toLowerCase()}`
        : `${uploadTarget.type}_${uploadTarget.id}`;

      await saveDocument(storageKey, uploadTarget.type, file);
      setFileStatus(uploadTarget.id, uploadTarget.type, true);
    } catch (err) {
      console.error('Error uploading file directly:', err);
      alert('Error al guardar el archivo.');
    } finally {
      setUploadTarget(null);
      e.target.value = '';
    }
  };

  const uniqueUsers = useMemo(() => {
    const set = new Set();
    equipos.forEach(r => {
      if (!isAvailable(r['Usuario'])) set.add(r['Usuario'].trim());
    });
    return [...set].sort((a, b) => a.localeCompare(b, 'es'));
  }, [equipos]);

  const funcSuggestions = useMemo(() => {
    if (!funcSearch) return uniqueUsers;
    const v = norm(funcSearch);
    return uniqueUsers.filter(u => norm(u).includes(v));
  }, [uniqueUsers, funcSearch]);

  const { descList, modList } = useMemo(() => {
    const dSet = new Set();
    const mSet = new Set();
    equipos.forEach(r => {
      if (r['Descripción del Bien']) dSet.add(r['Descripción del Bien'].trim());
      if (r['Modelo']) mSet.add(r['Modelo'].trim());
    });
    return {
      descList: [...dSet].sort((a, b) => a.localeCompare(b, 'es')),
      modList: [...mSet].sort((a, b) => a.localeCompare(b, 'es'))
    };
  }, [equipos]);

  // Derived visible data based on active tab
  let baseData = [];
  let kpisEquip = null;

  const isSearchingQR = globalSearch && (searchParams.get('q') || searchParams.get('search')) === globalSearch;

  if (isSearchingQR) {
    baseData = equipos;
  } else if (activeTab === 'disp') {
    baseData = equipos.filter(r => isAvailable(r['Usuario']));
  } else if (activeTab === 'func') {
    if (selectedFunc) {
      baseData = equipos.filter(r => norm(r['Usuario']) === norm(selectedFunc));
    }
  } else if (activeTab === 'equip') {
    let pool = equipos;
    if (selectedDesc && selectedDesc !== 'ALL' && selectedDesc !== '') {
      pool = pool.filter(r => (r['Descripción del Bien'] || '').trim() === selectedDesc.trim());
    }
    if (selectedMod && selectedMod !== 'ALL' && selectedMod !== '') {
      pool = pool.filter(r => (r['Modelo'] || '').trim() === selectedMod.trim());
    }

    if (selectedDesc !== '' || selectedMod !== '') {
      const dispCount = pool.filter(r => isAvailable(r['Usuario'])).length;
      kpisEquip = { total: pool.length, disp: dispCount, asig: pool.length - dispCount };
      baseData = pool;
    }
  }

  // Apply Global Search
  if (globalSearch) {
    const q = norm(globalSearch);
    baseData = baseData.filter(r => {
      const matchCols = COLUMNS.some(c => norm(r[c]).includes(q));
      const estadoFinal = getEstadoFinal(r);
      const matchEstado = norm(estadoFinal).includes(q);
      return matchCols || matchEstado;
    });
  }

  // Define active columns dynamically
  const activeCols = useMemo(() => {
    const transformCols = (cols) => {
      let newCols = [];
      cols.forEach(c => {
        if (c === 'Orden de Compra') newCols.push('Respaldo');
        else if (c === 'Factura') {} // Skip
        else if (c === 'SubDirección') {} // Skip, we will merge it with Usuario
        else newCols.push(c);
      });
      return newCols;
    };

    if (isSearchingQR) {
      const cols = ['Imagen'];
      COLUMNS.forEach(c => {
        if (c === 'Usuario') {
          cols.push('Estado');
        }
        cols.push(c);
      });
      return transformCols(cols);
    }
    if (activeTab === 'disp') {
      const cols = ['Imagen', ...COLUMNS.filter(c => c !== 'Usuario' && c !== 'SubDirección')];
      cols.push('Estado');
      return transformCols(cols);
    }
    const cols = ['Imagen'];
    COLUMNS.forEach(c => {
      if (c === 'Usuario') {
        cols.push('Estado');
      }
      cols.push(c);
    });
    return transformCols(cols);
  }, [activeTab, isSearchingQR]);

  // Apply Sort
  if (sortConfig.col) {
    baseData.sort((a, b) => {
      let va, vb;
      if (sortConfig.col === 'Estado') {
        va = norm(getEstadoFinal(a));
        vb = norm(getEstadoFinal(b));
      } else {
        va = norm(a[sortConfig.col]);
        vb = norm(b[sortConfig.col]);
      }
      if (va < vb) return -1 * sortConfig.dir;
      if (va > vb) return 1 * sortConfig.dir;
      return 0;
    });
  }

  const handleSort = (col) => {
    if (sortConfig.col === col) {
      setSortConfig({ col, dir: sortConfig.dir * -1 });
    } else {
      setSortConfig({ col, dir: 1 });
    }
  };

  const exportData = async (format) => {
    if (baseData.length === 0) {
      alert('No hay datos para exportar.');
      return;
    }

    const activeCols = (function() {
      const cols = COLUMNS.filter(c => activeTab === 'disp' ? (c !== 'Usuario' && c !== 'SubDirección') : true);
      cols.push('Estado');
      let newCols = [];
      cols.forEach(c => {
        if (c === 'Orden de Compra') newCols.push('Respaldo');
        else if (c === 'Factura') {}
        else newCols.push(c);
      });
      return newCols;
    })();

    const stamp = new Date().toISOString().slice(0, 16).replace(/[:T]/g, '-');
    const baseName = `inventario_${activeTab}_${stamp}`;

    if (format === 'xlsx') {
      const { Workbook } = await import('exceljs');
      const { saveAs } = await import('file-saver');
      const wb = new Workbook();
      const ws = wb.addWorksheet('Inventario');

      const exportCols = COLUMNS.filter(c => c !== 'Imagen');
      const idx = exportCols.indexOf('ID Publicación');
      const allExportCols = [
        ...exportCols.slice(0, idx + 1),
        'Tipo Publicación',
        ...exportCols.slice(idx + 1),
        '_id_interno_'
      ];

      ws.columns = allExportCols.map(c => ({
        header: c.toUpperCase(),
        key: c,
        width: Math.max(c.length + 5, 20)
      }));

      const headerRow = ws.getRow(1);
      headerRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF006BB9' } };
        cell.font = { color: { argb: 'FFFFFFFF' }, bold: true, size: 11 };
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
      });
      headerRow.height = 25;

      baseData.forEach((r, i) => {
        const rowData = {};
        allExportCols.forEach(c => {
          if (c === 'Estado') {
            rowData[c] = getEstadoFinal(r) || '';
          } else if (c === '_id_interno_') {
            rowData[c] = r.id || '';
          } else if (c === 'Tipo Publicación') {
            rowData[c] = r['Tipo Publicación'] || '';
          } else {
            rowData[c] = r[c] == null ? '' : String(r[c]);
          }
        });
        const row = ws.addRow(rowData);

        if (i % 2 !== 0) {
          row.eachCell(cell => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } }; });
        }

        row.eachCell((cell, colNumber) => {
          cell.font = { size: 11, color: { argb: 'FF374151' } };
          cell.alignment = { vertical: 'middle', horizontal: 'left' };
          const colName = activeCols[colNumber - 1];
          if (colName === 'Estado') {
            const val = cell.value;
            let color = 'FF19214D'; // ASIGNADO
            if (val === 'DISPONIBLE') color = 'FF4A7A1B';
            else if (val === 'PARA PRESTAMO') color = 'FF6B21A8';
            else if (val === 'EN PRESTAMO') color = 'FFC2410C';
            else if (val === 'DE BAJA') color = 'FF991B1B';

            cell.font = { bold: true, size: 10, color: { argb: color } };
          }
          if (colName === 'Nº de serie') {
            cell.font = { ...cell.font, name: 'Courier New' };
          }
        });
        row.height = 20;
      });

      const buffer = await wb.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), baseName + '.xlsx');

    } else if (format === 'pdf') {
      const { default: jsPDF } = await import('jspdf');
      const { default: autoTable } = await import('jspdf-autotable');

      const doc = new jsPDF('landscape', 'pt', 'a4');
      const tableRows = baseData.map(r => activeCols.map(c => {
        const estadoFinal = getEstadoFinal(r);
        if (c === 'Estado') return estadoFinal;
        if (c === 'Respaldo') {
          const fac = safe(r['Factura']);
          const oc = safe(r['Orden de Compra']);
          return [fac ? `Factura: ${fac}` : '', oc ? `OC: ${oc}` : ''].filter(Boolean).join(' / ');
        }
        return safe(r[c]);
      }));

      doc.setFontSize(16);
      doc.setTextColor(37, 48, 107);
      doc.text(`Reporte de Inventario - ${activeTab === 'disp' ? 'Equipos Disponibles' : activeTab === 'func' ? 'Por Funcionario' : 'Por Equipamiento'}`, 40, 40);

      autoTable(doc, {
        head: [activeCols.map(c => c.toUpperCase())],
        body: tableRows,
        startY: 60,
        styles: { fontSize: 8, font: 'helvetica', cellPadding: 6, textColor: [55, 65, 81] },
        headStyles: { fillColor: [0, 107, 185], textColor: 255, fontStyle: 'bold' },
        alternateRowStyles: { fillColor: [248, 250, 252] },
        columnStyles: {
          [activeCols.indexOf('Estado')]: { fontStyle: 'bold' },
          [activeCols.indexOf('Nº de serie')]: { font: 'courier' }
        },
        didParseCell: function (data) {
          if (data.section === 'body' && data.column.index === activeCols.indexOf('Estado')) {
            const val = data.cell.raw;
            if (val === 'DISPONIBLE') data.cell.styles.textColor = [100, 160, 40];
            else if (val === 'PARA PRESTAMO') data.cell.styles.textColor = [107, 33, 168];
            else if (val === 'EN PRESTAMO') data.cell.styles.textColor = [194, 65, 12];
            else if (val === 'DE BAJA') data.cell.styles.textColor = [153, 27, 27];
            else data.cell.styles.textColor = [37, 48, 107];
          }
        }
      });
      doc.save(baseName + '.pdf');

    } else {
      const exportCols = COLUMNS.filter(c => c !== 'Imagen');
      const idx = exportCols.indexOf('ID Publicación');
      const allExportCols = [
        ...exportCols.slice(0, idx + 1),
        'Tipo Publicación',
        ...exportCols.slice(idx + 1),
        '_id_interno_'
      ];

      const exportRows = baseData.map(r => {
        const o = {};
        allExportCols.forEach(c => {
          if (c === 'Estado') {
            o[c] = getEstadoFinal(r) || '';
          } else if (c === '_id_interno_') {
            o[c] = r.id || '';
          } else if (c === 'Tipo Publicación') {
            o[c] = r['Tipo Publicación'] || '';
          } else {
            o[c] = r[c] == null ? '' : String(r[c]);
          }
        });
        return o;
      });
      const csv = Papa.unparse(exportRows);
      const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = baseName + '.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }
  };

  const generarCodigoInventario = (eq) => {
    if (!eq) return 'SLEPLC-XXX-0000';
    const d = (eq['Descripción del Bien'] || '').toLowerCase();
    let prefix = 'XXX';
    if (d.includes('notebook') || d.includes('laptop')) prefix = 'NOT';
    else if (d.includes('impresora')) prefix = 'IMP';
    else if (d.includes('all in one') || d.includes('aio')) prefix = 'AIO';
    else if (d.includes('proyector')) prefix = 'PRY';
    else if (d.includes('monitor')) prefix = 'MON';
    else if (d.includes('router')) prefix = 'ROT';
    else if (d.includes('switch')) prefix = 'SWT';
    else if (d.includes('dron') || d.includes('drone')) prefix = 'DRO';
    else if (d.includes('dock')) prefix = 'DOC';
    else if (d.includes('camara') || d.includes('cámara') || d.includes('fotografica') || d.includes('fotográfica')) prefix = 'CAM';
    else if (d.includes('tv') || d.includes('smart tv') || d.includes('televisor')) prefix = 'STV';
    else if (d.includes('tablet')) prefix = 'TAB';
    else {
      prefix = d.substring(0, 3).toUpperCase().replace(/[^A-Z]/g, 'X').padEnd(3, 'X');
    }
    
    const equiposConMismoPrefix = baseData.filter(e => {
       const desc = (e['Descripción del Bien'] || '').toLowerCase();
       let p = 'XXX';
       if (desc.includes('notebook') || desc.includes('laptop')) p = 'NOT';
       else if (desc.includes('impresora')) p = 'IMP';
       else if (desc.includes('all in one') || desc.includes('aio')) p = 'AIO';
       else if (desc.includes('proyector')) p = 'PRY';
       else if (desc.includes('monitor')) p = 'MON';
       else if (desc.includes('router')) p = 'ROT';
       else if (desc.includes('switch')) p = 'SWT';
       else if (desc.includes('dron') || desc.includes('drone')) p = 'DRO';
       else if (desc.includes('dock')) p = 'DOC';
       else if (desc.includes('camara') || desc.includes('cámara') || desc.includes('fotografica') || desc.includes('fotográfica')) p = 'CAM';
       else if (desc.includes('tv') || desc.includes('smart tv') || desc.includes('televisor')) p = 'STV';
       else if (desc.includes('tablet')) p = 'TAB';
       else {
         p = desc.substring(0, 3).toUpperCase().replace(/[^A-Z]/g, 'X').padEnd(3, 'X');
       }
       return p === prefix;
    }).sort((a, b) => a.id - b.id);

    const index = equiposConMismoPrefix.findIndex(e => e.id === eq.id);
    const sequential = index !== -1 ? index + 1 : equiposConMismoPrefix.length + 1;
    const sequentialStr = String(sequential).padStart(4, '0');
    return `SLEPLC-${prefix}-${sequentialStr}`;
  };

  const generateStickerCanvas = (svgElement, eq) => {
    return new Promise((resolve) => {
      const svgString = new XMLSerializer().serializeToString(svgElement);
      const svgBlob = new Blob([svgString], { type: "image/svg+xml;charset=utf-8" });
      const URL = window.URL || window.webkitURL || window;
      const blobURL = URL.createObjectURL(svgBlob);
      const image = new Image();
      image.onload = () => {
        const canvas = document.createElement("canvas");
        const scale = 3;
        const baseWidth = 208;
        const baseHeight = 302;
        canvas.width = baseWidth * scale;
        canvas.height = baseHeight * scale;
        const ctx = canvas.getContext("2d");
        ctx.scale(scale, scale);
        
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(0, 0, baseWidth, baseHeight);
        
        ctx.fillStyle = "#1e293b";
        ctx.fillRect(10, 10, baseWidth - 20, 24);
        ctx.fillStyle = "#FFFFFF";
        ctx.font = "bold 11px 'Segoe UI', sans-serif";
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText("SLEP LOS COPIHUES", baseWidth / 2, 22);
        
        const qrSize = 130;
        const qrX = (baseWidth - qrSize) / 2;
        ctx.drawImage(image, qrX, 45, qrSize, qrSize);
        
        const boxY = 185;
        const boxH = 26;
        ctx.fillStyle = "#f1f5f9";
        ctx.strokeStyle = "#cbd5e1";
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.roundRect(15, boxY, baseWidth - 30, boxH, 4);
        ctx.fill();
        ctx.stroke();
        
        const codigoInventario = generarCodigoInventario(eq);
        ctx.fillStyle = "#0f172a";
        ctx.font = "bold 13px 'Segoe UI', sans-serif";
        ctx.fillText(codigoInventario, baseWidth / 2, boxY + (boxH / 2));
        
        const descripcionCompleta = `${eq['Descripción del Bien'] || ''} ${eq['Marca'] || ''} ${eq['Modelo'] || ''}`.trim();
        ctx.font = "bold 10px 'Segoe UI', sans-serif";
        ctx.fillText(descripcionCompleta, baseWidth / 2, 225);
        
        const serialText = eq['Nº de serie'] ? `S/N: ${eq['Nº de serie']}` : '';
        ctx.fillStyle = "#475569";
        ctx.font = "9px 'Segoe UI', sans-serif";
        ctx.fillText(serialText, baseWidth / 2, 240);
        
        URL.revokeObjectURL(blobURL);
        resolve(canvas.toDataURL("image/png"));
      };
      image.src = blobURL;
    });
  };

  const handleDownloadQR = async (eq) => {
    const svgElement = document.getElementById("qr-code-svg");
    if (!svgElement) return;
    const pngDataUrl = await generateStickerCanvas(svgElement, eq);
    const downloadLink = document.createElement("a");
    downloadLink.href = pngDataUrl;
    downloadLink.download = `QR_${eq['Nº de serie'] || 'equipo'}.png`;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
  };

  const handlePrintQR = async (eq) => {
    const svgElement = document.getElementById("qr-code-svg");
    if (!svgElement) return;
    const pngDataUrl = await generateStickerCanvas(svgElement, eq);
    const printWindow = window.open('', '_blank', 'width=400,height=600');
    printWindow.document.write(`
      <html>
        <head>
          <title>Imprimir Etiqueta QR</title>
          <style>
            body {
              margin: 0;
              display: flex;
              justify-content: center;
              align-items: center;
              height: 100vh;
              background: #fff;
            }
            img {
              width: 5.5cm;
              height: auto;
              max-width: 100%;
            }
            @media print {
              body { align-items: flex-start; justify-content: flex-start; }
            }
          </style>
        </head>
        <body>
          <img src="${pngDataUrl}" onload="window.print(); setTimeout(() => window.close(), 500);" />
        </body>
      </html>
    `);
    printWindow.document.close();
  };

  const total = equipos.length;
  const totalDisp = equipos.filter(r => isAvailable(r['Usuario'])).length;
  const totalAsig = total - totalDisp;

  if (loading) return <div className="p-8 text-center">Cargando...</div>;

  return (
    <div className="p-6 w-full max-w-[1920px] mx-auto space-y-6">

      {/* Header */}
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 border-b border-gray-200 pb-5">
        <div>
          <h1 className="text-2xl font-bold text-gray-800 flex items-center gap-2">
            <MonitorSmartphone size={26} className="text-[#006BB9]" /> Equipos
          </h1>
          <p className="text-sm text-gray-500 mt-1">Gestión general, asignaciones y métricas de equipos informáticos.</p>
        </div>
        <div className="flex gap-2 flex-wrap">
          {canEdit && (
            <>
              <button
                onClick={() => { setStatus({ type: 'idle', message: '' }); setIsMasivaModalOpen(true); }}
                className="flex items-center gap-2 bg-blue-100 text-[#006BB9] px-4 py-2 rounded-lg hover:bg-blue-200 transition-colors text-sm font-medium border border-blue-200"
              >
                <UploadCloud size={16} /> Carga Masiva
              </button>
              <button
                onClick={() => setIsNuevoEquipoModalOpen(true)}
                className="flex items-center gap-2 bg-[#112A46] text-white px-4 py-2 rounded-lg hover:bg-[#1A3A5F] transition-colors text-sm font-medium shadow-sm cursor-pointer"
              >
                <PlusCircle size={16} /> Nuevo Equipo
              </button>
            </>
          )}
        </div>
      </div>



      {/* Tabs & Controls */}
      <section className="bg-white rounded-xl shadow-sm">
        <div className="flex flex-col xl:flex-row justify-between items-stretch xl:items-center border-b border-gray-200">
          <nav className="flex flex-wrap w-full xl:w-auto">
            <button onClick={() => setActiveTab('disp')} className={`px-4 py-3 text-sm font-medium border-b-2 flex items-center gap-2 transition-colors ${activeTab === 'disp' ? 'border-[#25306B] bg-[#25306B] text-white' : 'border-transparent text-gray-600 hover:bg-gray-50'}`}>
              <Package size={16} /> Disponibles
            </button>
            <button onClick={() => setActiveTab('func')} className={`px-4 py-3 text-sm font-medium border-b-2 flex items-center gap-2 transition-colors ${activeTab === 'func' ? 'border-[#25306B] bg-[#25306B] text-white' : 'border-transparent text-gray-600 hover:bg-gray-50'}`}>
              <UserCircle size={16} /> Por Funcionario
            </button>
            <button onClick={() => setActiveTab('equip')} className={`px-4 py-3 text-sm font-medium border-b-2 flex items-center gap-2 transition-colors ${activeTab === 'equip' ? 'border-[#25306B] bg-[#25306B] text-white' : 'border-transparent text-gray-600 hover:bg-gray-50'}`}>
              <MonitorSmartphone size={16} /> Por Equipamiento
            </button>
          </nav>
          <div className="flex flex-wrap gap-2 p-3 xl:p-0 xl:pr-4 bg-gray-50 xl:bg-transparent no-print-interactive">
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
              <input type="text" value={globalSearch} onChange={e => setGlobalSearch(e.target.value)} placeholder="Filtrar en tabla activa..." className="pl-9 pr-3 py-1.5 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-[#006BB9] focus:outline-none min-w-[200px]" />
            </div>
            <button onClick={() => exportData('xlsx')} className="px-3 py-1.5 text-green-800 rounded-lg text-sm font-medium shadow-sm transition-colors bg-green-200 hover:bg-green-300 flex items-center gap-2">
              <Download size={14} /> Excel
            </button>
            <button onClick={() => exportData('csv')} className="px-3 py-1.5 text-sky-800 rounded-lg text-sm font-medium shadow-sm transition-colors bg-sky-200 hover:bg-sky-300 flex items-center gap-2">
              <Download size={14} /> CSV
            </button>
            <button onClick={() => exportData('pdf')} className="px-3 py-1.5 text-rose-800 rounded-lg text-sm font-medium shadow-sm transition-colors bg-rose-200 hover:bg-rose-300 flex items-center gap-2">
              <Printer size={14} /> PDF
            </button>
          </div>
        </div>

        <div className="p-5">
          {/* Funcionario Tab Controls */}
          {activeTab === 'func' && (
            <div className="flex flex-col sm:flex-row gap-4 mb-4 items-stretch sm:items-center justify-between no-print-interactive">
              <div className="relative flex-1 sm:max-w-md">
                <div className="relative flex items-center">
                  <Search className="absolute left-3 w-4 h-4 text-gray-400" />
                  <input
                    type="text"
                    value={funcSearch}
                    onChange={e => {
                      setFuncSearch(e.target.value);
                      setShowFuncSug(true);
                      setFocusedIndex(-1);
                      if (!e.target.value) setSelectedFunc('');
                    }}
                    onKeyDown={e => {
                      if (!showFuncSug) return;
                      if (e.key === 'ArrowDown') {
                        e.preventDefault();
                        setFocusedIndex(prev => (prev < funcSuggestions.length - 1 ? prev + 1 : prev));
                      } else if (e.key === 'ArrowUp') {
                        e.preventDefault();
                        setFocusedIndex(prev => (prev > 0 ? prev - 1 : 0));
                      } else if (e.key === 'Enter') {
                        e.preventDefault();
                        if (focusedIndex >= 0 && focusedIndex < funcSuggestions.length) {
                          const u = funcSuggestions[focusedIndex];
                          setSelectedFunc(u);
                          setFuncSearch(u);
                          setShowFuncSug(false);
                          setFocusedIndex(-1);
                        }
                      } else if (e.key === 'Escape') {
                        setShowFuncSug(false);
                        setFocusedIndex(-1);
                      }
                    }}
                    onFocus={() => setShowFuncSug(true)}
                    onClick={() => {
                      if (selectedFunc) {
                        setFuncSearch('');
                        setSelectedFunc('');
                        setShowFuncSug(true);
                      }
                    }}
                    onBlur={() => setTimeout(() => { setShowFuncSug(false); setFocusedIndex(-1); }, 200)}
                    placeholder="Buscar funcionario..."
                    className="w-full pl-9 pr-4 py-2.5 bg-white border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-[#006BB9] focus:outline-none shadow-sm transition-all"
                  />
                </div>
                {showFuncSug && (
                  <div className="absolute z-50 w-full bg-white border border-gray-200 rounded-lg mt-1 max-h-60 overflow-y-auto shadow-xl">
                    <div className="py-1">
                      {funcSuggestions.length > 0 ? funcSuggestions.map((u, idx) => (
                        <div
                          key={u}
                          onMouseDown={() => { setSelectedFunc(u); setFuncSearch(u); setShowFuncSug(false); setFocusedIndex(-1); }}
                          className={`flex items-center gap-3 px-3 py-2 cursor-pointer transition-colors ${focusedIndex === idx ? 'bg-blue-100' : 'hover:bg-slate-50'}`}
                        >
                          <Badge variant="user" categoria="nombres" estado="Funcionario" text={u} />
                        </div>
                      )) : <div className="px-4 py-3 text-slate-500 italic text-center text-sm">Sin coincidencias...</div>}
                    </div>
                  </div>
                )}
              </div>
              <div className="flex items-center gap-4">
                <div className="text-sm font-medium text-[#25306B] flex items-center gap-2">
                  {selectedFunc ? <><CheckCircle size={16} className="text-green-600" /> Mostrando activos de {selectedFunc}</> : 'Seleccione un funcionario'}
                </div>
                {selectedFunc && baseData.length > 0 && (
                  <button 
                    onClick={handleGenerateActaMasivaFuncionario}
                    className="px-3 py-1.5 text-blue-800 rounded-lg text-sm font-bold shadow-sm transition-colors bg-blue-100 hover:bg-blue-200 flex items-center gap-2 border border-blue-200"
                  >
                    <Download size={14} strokeWidth={2.5} /> Acta Consolidada
                  </button>
                )}
              </div>
            </div>
          )}

          {/* Equipamiento Tab Controls */}
          {activeTab === 'equip' && (
            <div className="mb-4 no-print-interactive">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-[#EDF0F5] p-4 rounded-lg border border-gray-200 mb-3">
                <div>
                  <label className="block text-xs font-semibold text-[#25306B] mb-1 uppercase tracking-wide">Por Descripción del Bien</label>
                  <select value={selectedDesc} onChange={e => setSelectedDesc(e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-[#006BB9] focus:outline-none shadow-sm bg-white">
                    <option value="">— Seleccionar opción —</option>
                    <option value="ALL">— Mostrar Todo —</option>
                    {descList.map(d => <option key={d} value={d}>{d}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-semibold text-[#25306B] mb-1 uppercase tracking-wide">Por Modelo</label>
                  <select value={selectedMod} onChange={e => setSelectedMod(e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-[#006BB9] focus:outline-none shadow-sm bg-white">
                    <option value="">— Seleccionar opción —</option>
                    <option value="ALL">— Mostrar Todos los Modelos —</option>
                    {modList.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                </div>
              </div>

              {kpisEquip && (
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-4">
                  <div className="bg-white rounded-lg p-3 border-l-4 shadow-sm" style={{ borderColor: 'var(--slep-primary)' }}>
                    <div className="text-xs text-gray-500 uppercase font-semibold">Total del tipo</div>
                    <div className="text-2xl font-bold text-[#25306B]">{kpisEquip.total}</div>
                  </div>
                  <div className="bg-white rounded-lg p-3 border-l-4 shadow-sm" style={{ borderColor: 'var(--slep-green)' }}>
                    <div className="text-xs text-gray-500 uppercase font-semibold">Disponibles en bodega</div>
                    <div className="text-2xl font-bold text-[#90d039]">{kpisEquip.disp}</div>
                  </div>
                  <div className="bg-white rounded-lg p-3 border-l-4 shadow-sm" style={{ borderColor: 'var(--slep-secondary)' }}>
                    <div className="text-xs text-gray-500 uppercase font-semibold">Asignados</div>
                    <div className="text-2xl font-bold text-[#006BB9]">{kpisEquip.asig}</div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Table */}
          <div className="table-scroll rounded-lg border border-gray-200">
            {baseData.length === 0 ? (
              <div className="p-8 text-center text-gray-500 text-sm">
                {activeTab === 'func' && !selectedFunc ? 'Seleccione un funcionario para ver sus equipos asignados.' :
                  activeTab === 'equip' && (selectedDesc === '' && selectedMod === '') ? 'Seleccione una descripción o modelo para comenzar el análisis.' :
                    'No hay registros que mostrar.'}
              </div>
            ) : (
              <table>
                <thead>
                  <tr>
                    {activeCols.map(c => {
                      let headerClass = c === 'Estado' ? "sortable text-left align-top w-24" : "sortable text-left align-top";
                      if (c === 'ID Publicación') headerClass += " max-w-[150px]";
                      if (c === 'Orden de Compra') headerClass += " max-w-[80px]";
                      if (c === 'Respaldo') headerClass += " min-w-[200px] w-[200px]";

                      const onClickHandler = () => handleSort(c);

                      return (
                        <th key={c} onClick={onClickHandler} className={headerClass}>
                          <div className="flex items-start gap-1 justify-between">
                            <span className="whitespace-normal leading-snug">{c === 'Imagen' ? '' : c}</span>
                            {sortConfig.col === c ? <span className="text-[11px] mt-0.5 shrink-0">{sortConfig.dir === 1 ? '▲' : '▼'}</span> : null}
                          </div>
                        </th>
                      );
                    })}
                    <th className="text-center w-24 no-print-interactive align-top">
                      <div className="flex justify-center items-start gap-1">Acciones</div>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {baseData.map((row, i) => {
                    const disp = isAvailable(row['Usuario']);
                    const estadoFinal = getEstadoFinal(row);
                    return (
                      <tr key={i} className="hover:bg-blue-50 even:bg-slate-50 transition-colors">
                        {activeCols.map(c => {
                          const value = c === 'Estado' ? '' : safe(row[c]);
                          const itemId = row.id || row['Nº de serie'] || `temp_${i}`;

                          if (c === 'Imagen') {
                            return (
                              <td key={c} className="px-3 py-2 w-[52px]">
                                <div className="w-[52px] h-[52px] rounded-[6px] bg-white border border-gray-200 overflow-hidden flex items-center justify-center shrink-0 shadow-sm">
                                  {row.imagen_url ? (
                                    <img src={row.imagen_url} alt={row['Descripción del Bien'] || 'Equipo'} className="w-full h-full object-contain" />
                                  ) : (
                                    <span className="text-[8px] text-gray-400 font-bold uppercase text-center leading-tight">Sin<br/>Img</span>
                                  )}
                                </div>
                              </td>
                            );
                          }

                          if (c === 'Estado') {
                            return (
                              <td key={c} className="px-4 py-3 align-middle">
                                <Badge categoria="equipos" estado={estadoFinal} />
                              </td>
                            );
                          }

                          if (c === 'Respaldo') {
                            const fac = row['Factura'];
                            const oc = row['Orden de Compra'];
                            const hasFacFile = row.hasFacturaFile;
                            const hasOcFile = row.hasOcFile;
                            return (
                              <td key={c} className="px-3 py-2 text-[12px] text-gray-700 min-w-[220px] w-[220px]">
                                <div className="flex flex-col gap-1.5">
                                  <div className="flex items-center gap-2">
                                    {fac ? (
                                      <button
                                        onClick={() => hasFacFile ? handlePreview(itemId, 'factura') : null}
                                        className={`flex items-center gap-1.5 transition-colors whitespace-nowrap ${hasFacFile ? 'text-[#006BB9] hover:text-blue-800 text-[11px] font-bold cursor-pointer' : 'text-gray-500 text-[11px] font-bold cursor-default'}`}
                                        title={hasFacFile ? `Ver Factura ${fac}` : `Factura: ${fac} (Sin archivo)`}
                                      >
                                        <FileText size={12} className="shrink-0" /> FACTURA N° {fac}
                                      </button>
                                    ) : (
                                      <span className="flex items-center gap-1 text-amber-700 text-[10px] font-bold uppercase whitespace-nowrap" title="Falta Factura">
                                        <AlertCircle size={10} className="shrink-0" /> Sin Factura
                                      </span>
                                    )}
                                  </div>
                                  <div className="flex items-center gap-2">
                                    {oc ? (
                                      <button
                                        onClick={() => hasOcFile ? handlePreview(itemId, 'oc') : null}
                                        className={`flex items-center gap-1.5 transition-colors whitespace-nowrap ${hasOcFile ? 'text-emerald-700 hover:text-emerald-900 text-[11px] font-bold cursor-pointer' : 'text-gray-500 text-[11px] font-bold cursor-default'}`}
                                        title={hasOcFile ? `Ver OC ${oc}` : `OC: ${oc} (Sin archivo)`}
                                      >
                                        <FileText size={12} className="shrink-0" /> OC N° {oc}
                                      </button>
                                    ) : (
                                      <span className="flex items-center gap-1 text-amber-700 text-[10px] font-bold uppercase whitespace-nowrap" title="Falta Orden de Compra">
                                        <AlertCircle size={10} className="shrink-0" /> Sin OC
                                      </span>
                                    )}
                                  </div>
                                </div>
                              </td>
                            );
                          }

                          if (c === 'ID Publicación') {
                            const tipo = row['Tipo Publicación'];
                            let badgeText = '';
                            if (tipo === 'Convenio Marco') badgeText = 'CM';
                            else if (tipo === 'Compra Ágil') badgeText = 'CA';
                            else if (tipo === 'Licitación') badgeText = 'LI';
                            else badgeText = tipo;

                            return (
                              <td key={c} className="px-3 py-2 text-[12px] text-gray-700 whitespace-nowrap">
                                <div className="flex items-center gap-1.5">
                                  <span>{value}</span>
                                  {tipo && (
                                    <span
                                      title={tipo}
                                      className="inline-flex items-center justify-center px-1.5 py-0.5 bg-blue-50 border border-blue-200 text-[#006BB9] rounded text-[9px] uppercase font-bold cursor-help"
                                    >
                                      {badgeText}
                                    </span>
                                  )}
                                </div>
                              </td>
                            );
                          }

                          if (c === 'Usuario') {
                            const valSubdir = row['SubDirección'];
                            const obsAsig = row.observacion_asignacion || row['observacion_asignacion'];
                            return (
                              <td key={c} className="px-3 py-2 whitespace-nowrap">
                                <div className="flex items-center gap-2">
                                  <div>
                                    <div className="font-[800] text-[#111827] text-[14px]">
                                      {isAvailable(value) ? '—' : value}
                                    </div>
                                    <div className="font-[500] text-[12px] text-[#6b7280]">
                                      {isAvailable(value) ? '' : (valSubdir || '—')}
                                    </div>
                                  </div>
                                  {obsAsig && !isAvailable(value) && (
                                    <div className="group relative flex items-center justify-center cursor-help">
                                      <Info size={16} className="text-red-500 hover:text-red-600 transition-colors" />
                                      <div className="pointer-events-none absolute bottom-full left-1/2 -translate-x-1/2 mb-2 w-max max-w-xs opacity-0 transition-opacity group-hover:opacity-100 z-50">
                                        <div className="bg-gray-800 text-white text-xs rounded-lg py-2 px-3 shadow-xl whitespace-normal break-words text-center border border-gray-700">
                                          {obsAsig}
                                        </div>
                                        <div className="w-0 h-0 border-l-[6px] border-l-transparent border-t-[6px] border-t-gray-800 border-r-[6px] border-r-transparent absolute left-1/2 -translate-x-1/2 top-full" />
                                      </div>
                                    </div>
                                  )}
                                </div>
                              </td>
                            );
                          }

                          if (c === 'Departamento') {
                            return (
                              <td key={c} className="px-3 py-2 text-[12px] font-[500] text-[#6b7280] whitespace-nowrap">
                                {value}
                              </td>
                            );
                          }

                          let cellClass = "px-3 py-2 max-w-[200px] break-words";
                          if (c === 'Descripción del Bien' || c === 'Marca' || c === 'Modelo') {
                            cellClass += " text-[14px] font-[800] text-[#334155]";
                          } else if (c === 'Nº de serie') {
                            cellClass += " text-[14px] font-[500] text-[#334155]";
                          } else {
                            cellClass += " text-[12px] text-gray-700";
                          }

                          return <td key={c} className={cellClass}>{value}</td>;
                        })}
                        <td className="text-center no-print-interactive">
                          <div className="flex justify-center items-center gap-2">
                            {/* Acta Button */}
                            {!isAvailable(row['Usuario']) && (
                              row.acta_firmada_url ? (
                                <button
                                  onClick={() => handleVerActa(row.acta_firmada_url)}
                                  className="p-2 text-emerald-600 hover:text-white hover:bg-emerald-600 border border-emerald-200 hover:border-emerald-600 rounded-lg transition-all shadow-xs flex items-center justify-center cursor-pointer"
                                  title="Ver Acta Firmada"
                                >
                                  <FileText size={14} className="stroke-[2.5]" />
                                </button>
                              ) : (
                                <button
                                  onClick={() => handleGenerateActaAsignacion(row)}
                                  className="p-2 text-indigo-600 hover:text-white hover:bg-indigo-600 border border-indigo-200 hover:border-indigo-600 rounded-lg transition-all shadow-xs flex items-center justify-center cursor-pointer"
                                  title="Generar Acta de Asignación"
                                >
                                  <Download size={14} className="stroke-[2.5]" />
                                </button>
                              )
                            )}
                            <button
                              disabled={!isQRSupported(row['Descripción del Bien'])}
                              onClick={() => setQrModalData(row)}
                              className={`p-2 rounded-lg transition-all shadow-xs flex items-center justify-center cursor-pointer border ${
                                isQRSupported(row['Descripción del Bien'])
                                  ? 'text-[#006BB9] hover:text-white hover:bg-[#006BB9] border-blue-200 hover:border-[#006BB9]'
                                  : 'text-gray-300 border-gray-100 bg-gray-50 cursor-not-allowed'
                              }`}
                              title={isQRSupported(row['Descripción del Bien']) ? "Generar Código QR" : "Código QR no soportado para este tipo de bien"}
                            >
                              <QrCode size={14} className="stroke-[2.5]" />
                            </button>
                            {canEdit && (
                              <>
                                <button
                                  onClick={() => { 
                                    setAssignModalData(row); 
                                    setAssignUserName(row['Usuario'] && row['Usuario'] !== '—' ? row['Usuario'] : '');
                                    setAssignDate(row.fecha_asignacion ? row.fecha_asignacion.split('T')[0] : new Date().toISOString().split('T')[0]);
                                    setAssignObservation(row.observacion_asignacion || '');
                                  }}
                                  className="p-2 text-emerald-600 hover:text-white hover:bg-emerald-600 border border-emerald-200 hover:border-emerald-600 rounded-lg transition-all shadow-xs flex items-center justify-center cursor-pointer"
                                  title="Asignar Equipo a un Funcionario"
                                >
                                  <UserPlus size={14} className="stroke-[2.5]" />
                                </button>
                                <button
                                  onClick={() => setEditingEquipo(row)}
                                  className="p-2 text-[#006BB9] hover:text-white hover:bg-[#006BB9] border border-blue-200 hover:border-[#006BB9] rounded-lg transition-all shadow-xs flex items-center justify-center cursor-pointer"
                                  title="Editar Ficha del Equipo"
                                >
                                  <Pencil size={14} className="stroke-[2.5]" />
                                </button>
                              </>
                            )}
                          </div>
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            )}
          </div>
        </div>
      </section>
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleDirectUpload}
        accept="application/pdf,image/*"
        className="hidden"
      />

      {/* Modales */}
      {isNuevoEquipoModalOpen && (
        <NuevoEquipoModal 
          isOpen={isNuevoEquipoModalOpen}
          onClose={() => setIsNuevoEquipoModalOpen(false)}
        />
      )}

      {/* Modal Asignar Equipo */}
      {assignModalData && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white p-6 rounded-xl shadow-2xl w-full max-w-sm animate-fade-in relative flex flex-col">
            <button
              onClick={() => setAssignModalData(null)}
              className="absolute top-4 right-4 text-gray-400 hover:text-gray-600 bg-gray-100 hover:bg-gray-200 rounded-full w-8 h-8 flex items-center justify-center text-xl transition-colors cursor-pointer"
            >
              &times;
            </button>
            <h3 className="text-xl font-bold text-[#25306B] mb-2 flex items-center gap-2">
              <UserPlus className="text-[#006BB9]" /> Asignar Equipo
            </h3>
            <p className="text-sm text-gray-600 mb-4 border-b pb-4">
              <span className="font-semibold block text-base">{assignModalData['Descripción del Bien'] || 'Equipo'}</span>
              <span className="block mt-1"><strong>Marca:</strong> {assignModalData['Marca'] || 'N/A'}</span>
              <span className="block"><strong>Modelo:</strong> {assignModalData['Modelo'] || 'N/A'}</span>
              <span className="block"><strong>S/N:</strong> {assignModalData['Nº de serie'] || 'N/A'}</span>
              <span className="block mt-2"><strong>Estado Actual:</strong> <span className="uppercase text-xs font-bold px-2 py-1 bg-gray-100 rounded text-gray-600">{assignModalData.estado || 'DISPONIBLE'}</span></span>
            </p>
            <div className="mb-6 relative">
              <label className="block text-sm font-semibold text-[#25306B] mb-2">Nombre del Usuario / Funcionario</label>
              <AutocompleteInput
                value={assignUserName}
                onChange={(e) => setAssignUserName(e.target.value)}
                options={perfilesOptions}
                placeholder="Ej. Juan Pérez (Dejar vacío para Disponible)"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-[#006BB9] focus:outline-none bg-white mb-4"
              />
              <label className="block text-sm font-semibold text-[#25306B] mb-2 mt-4">Fecha de Asignación</label>
              <input
                type="date"
                value={assignDate}
                onChange={(e) => setAssignDate(e.target.value)}
                onClick={(e) => e.target.showPicker && e.target.showPicker()}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-[#006BB9] focus:outline-none bg-white cursor-pointer"
              />
              <label className="block text-sm font-semibold text-[#25306B] mb-2 mt-4">Observación Asignación</label>
              <textarea
                value={assignObservation}
                onChange={(e) => setAssignObservation(e.target.value)}
                placeholder="Ej: Entrega sin cargador, pantalla con rayón, etc."
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-[#006BB9] focus:outline-none bg-white min-h-[60px] resize-y mb-4"
              />
            </div>
            <div className="flex justify-end gap-3 mt-auto relative z-0">
              <button
                onClick={() => setAssignModalData(null)}
                className="px-4 py-2 text-sm font-semibold text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors"
              >
                Cancelar
              </button>
              <button
                onClick={async () => {
                  const currentDesc = assignModalData['Descripción del Bien'];
                  const isDisp = isAvailable(assignUserName);
                  let obsData = '';
                  
                  if (currentDesc && !isDisp) {
                    const descLower = currentDesc.toLowerCase();
                    const isPC = descLower.includes('notebook') || descLower.includes('all in one') || descLower.includes('pc') || descLower.includes('computador') || descLower.includes('desktop');

                    let hasDuplicate = false;
                    let conflictDesc = '';

                    if (isPC) {
                      const existingPC = equipos.find(eq =>
                        eq.id !== assignModalData.id &&
                        (eq['Usuario'] && isSameUser(eq['Usuario'], assignUserName)) &&
                        (eq['Descripción del Bien'] && (
                           eq['Descripción del Bien'].toLowerCase().includes('notebook') ||
                           eq['Descripción del Bien'].toLowerCase().includes('all in one') ||
                           eq['Descripción del Bien'].toLowerCase().includes('pc') ||
                           eq['Descripción del Bien'].toLowerCase().includes('computador') ||
                           eq['Descripción del Bien'].toLowerCase().includes('desktop')
                        ))
                      );
                      if (existingPC) {
                        hasDuplicate = true;
                        conflictDesc = existingPC['Descripción del Bien'];
                      }
                    } else {
                      const existingSame = equipos.find(eq =>
                        eq.id !== assignModalData.id &&
                        eq['Descripción del Bien'] === currentDesc &&
                        (eq['Usuario'] && isSameUser(eq['Usuario'], assignUserName))
                      );
                      if (existingSame) {
                        hasDuplicate = true;
                        conflictDesc = currentDesc;
                      }
                    }

                    if (hasDuplicate) {
                      const confirmed = await showAlertConfirm(
                        'Advertencia de Duplicidad',
                        `El usuario ya tiene asignado un equipo de este tipo (<strong>${conflictDesc}</strong>).<br/><br/>¿Desea continuar y asignarlo de igual manera?`
                      );
                      if (!confirmed) {
                        return;
                      }
                      
                      const reason = await showAlertPrompt(
                        'Justificación de Asignación',
                        'Ingrese el motivo por el cual se asigna un equipo adicional a este usuario (OBLIGATORIO):'
                      );
                      if (!reason || !reason.trim()) {
                        showLocalToast('Cancelado', 'La asignación de equipo adicional requiere una observación obligatoria.', 'error');
                        return;
                      }
                      obsData = reason.trim();
                    }
                  }

                  const updated = { ...assignModalData, 'Usuario': assignUserName, fecha_asignacion: assignDate, observacion_asignacion: assignObservation };
                  updated.estado = isDisp ? 'DISPONIBLE' : 'ASIGNADO';
                  if (obsData) {
                    updated.Observaciones = updated.Observaciones 
                      ? `${updated.Observaciones} | Asignación adicional: ${obsData}` 
                      : `Asignación adicional: ${obsData}`;
                  }
                  
                  updateEquipo(equipos.indexOf(assignModalData), updated);
                  setAssignModalData(null);
                  showLocalToast('Equipo Asignado', `El equipo se ha asignado correctamente a ${isDisp ? 'DISPONIBLE' : assignUserName}.`, 'success');
                }}
                className="px-4 py-2 text-sm font-semibold text-white bg-[#006BB9] hover:bg-[#25306B] rounded-lg transition-colors flex items-center gap-2 cursor-pointer"
              >
                <CheckCircle size={16} /> Guardar
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Carga Masiva */}
      {isMasivaModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-50 p-4">
          <div className="bg-white p-6 rounded-xl shadow-2xl w-full max-w-lg animate-fade-in relative max-h-[90vh] flex flex-col">
            <button
              onClick={() => { setIsMasivaModalOpen(false); setStatus({ type: 'idle', message: '' }); setLocalToast(null); }}
              className="absolute top-4 right-4 text-gray-400 hover:text-gray-600 bg-gray-100 hover:bg-gray-200 rounded-full p-1 transition-colors cursor-pointer"
            >
              &times;
            </button>
            <h2 className="text-xl font-bold mb-2 text-[#25306B] flex items-center gap-2">
              <UploadCloud className="text-[#006BB9]" /> Carga Masiva de Equipos
            </h2>

            <div className="overflow-y-auto flex-1 pr-2 mt-2 custom-scrollbar relative">

              {/* Toast Flotante Interno */}
              {localToast && (
                <div
                  onMouseDown={handleDragStart}
                  style={{
                    position: 'absolute',
                    left: '50%',
                    bottom: `calc(100% - 150px - ${toastPos.y}px)`,
                    transform: `translateX(calc(-50% + ${toastPos.x}px))`,
                    cursor: isDragging.current ? 'grabbing' : 'grab',
                    zIndex: 9999
                  }}
                >
                  <div
                    onMouseEnter={handleToastMouseEnter}
                    onMouseLeave={handleToastMouseLeave}
                    className={`flex items-start gap-3 px-4 py-3 rounded-xl shadow-xl border w-[340px] text-sm text-left ${localToast.type === 'success' ? 'bg-emerald-50 border-emerald-200 text-emerald-800' :
                      localToast.type === 'warning' ? 'bg-amber-50 border-amber-200 text-amber-800' :
                        'bg-red-50 border-red-200 text-red-800'
                      }`}
                  >
                    {localToast.type === 'success' && <CheckCircle className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />}
                    {localToast.type === 'warning' && <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />}
                    {localToast.type === 'error' && <AlertCircle className="w-5 h-5 text-red-600 shrink-0 mt-0.5" />}

                    <div className="flex-1">
                      <p className="font-bold text-[11px] uppercase tracking-wider">{localToast.title}</p>

                      {localToast.addedCount !== undefined || (localToast.duplicateSerials && localToast.duplicateSerials.length > 0) ? (
                        <div className="text-[11px] opacity-90 mt-1.5 space-y-1.5">
                          {localToast.addedCount !== undefined && (
                            <p>
                              {localToast.addedCount > 0
                                ? `Se agregaron ${localToast.addedCount} equipos en total.`
                                : 'No se agregaron nuevos equipos.'}
                            </p>
                          )}
                          {localToast.noSerialCount > 0 && (
                            <div>
                              <p className={localToast.isFirstUpload ? "text-amber-900 font-bold" : ""}>
                                • {localToast.noSerialCount} equipos {localToast.isFirstUpload ? "ingresados SIN N° de Serie." : "omitidos por falta de N° de Serie."}
                              </p>
                              {localToast.isFirstUpload && (
                                <p className="mt-1 text-[9px] leading-tight text-amber-800 bg-amber-100 p-1.5 rounded border border-amber-200">
                                  ⚠️ <strong>¡CRUCIAL!</strong> Es necesario que edites estos equipos a la brevedad y les asignes un número de serie o identificador único por seguridad.
                                </p>
                              )}
                            </div>
                          )}
                          {localToast.duplicateSerials && localToast.duplicateSerials.length > 0 && (
                            <div className="group relative cursor-help inline-block mt-1">
                              <span className="border-b border-dashed border-amber-500 font-semibold text-amber-900">
                                • {localToast.duplicateSerials.length} equipos omitidos por duplicidad (Ver detalle)
                              </span>

                              <div className="invisible group-hover:visible absolute bottom-full left-0 pb-2 z-50">
                                <div className="bg-slate-800 text-white p-2.5 rounded-lg shadow-xl w-60 max-h-48 overflow-y-auto leading-relaxed font-mono whitespace-normal normal-case border border-slate-700">
                                  <strong className="text-slate-300 block border-b border-slate-700 pb-1 mb-1">Series duplicadas omitidas:</strong>
                                  <div className="flex flex-wrap gap-1">
                                    {localToast.duplicateSerials.map((s, idx) => (
                                      <span key={idx} className="bg-slate-700 px-1 py-0.5 rounded text-[9px]">{s}</span>
                                    ))}
                                  </div>
                                </div>
                              </div>
                            </div>
                          )}
                        </div>
                      ) : (
                        <p className="text-[11px] opacity-90 mt-1 whitespace-pre-line">{localToast.message}</p>
                      )}
                    </div>

                    <button onClick={() => setLocalToast(null)} className="text-gray-400 hover:text-gray-600 font-bold ml-1 shrink-0 text-lg leading-none focus:outline-none cursor-pointer">
                      &times;
                    </button>
                  </div>
                </div>
              )}

              <p className="text-[13px] text-gray-600 mb-4 bg-blue-50 p-3 rounded-lg border border-blue-100">
                Sube un archivo Excel o CSV para importar o actualizar equipos de forma masiva.
              </p>

              <div className="bg-blue-50 rounded-lg p-4 border border-blue-100 text-xs text-blue-800 space-y-2 mb-5">
                <h3 className="font-bold flex items-center gap-1.5">
                  <AlertCircle className="w-4 h-4" /> Columnas Requeridas en el Archivo
                </h3>
                <ul className="list-disc pl-5 mt-1 space-y-0.5 text-blue-700">
                  {COLUMNS.map(c => <li key={c}>{c}</li>)}
                </ul>
                <p className="mt-2 text-[10px] opacity-80 leading-normal border-t border-blue-200/50 pt-1.5">
                  * Todos los equipos deben tener obligatoriamente un <strong>Nº de serie</strong>. Si se detecta un número de serie existente, se omitirá su carga para evitar alterar los datos.
                </p>
              </div>

              <div className="border-2 border-dashed border-blue-200 rounded-xl p-8 text-center bg-gray-50 hover:bg-blue-50 hover:border-blue-400 transition-colors relative group">
                <label className="cursor-pointer flex flex-col items-center justify-center">
                  <div className="w-14 h-14 bg-white rounded-full shadow-sm flex items-center justify-center mb-3 text-blue-600 group-hover:scale-110 group-hover:bg-blue-600 group-hover:text-white transition-all">
                    <UploadCloud size={28} />
                  </div>
                  <span className="font-bold text-[#006BB9] text-sm group-hover:underline">Haz clic para buscar el archivo</span>
                  <span className="text-[11px] text-gray-500 mt-1">o arrastra el archivo aquí</span>
                  <input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleMasivaFile} ref={masivaFileInputRef} />
                </label>
              </div>

              {status.type !== 'idle' && (
                <div className={`mt-4 p-3 rounded-lg flex items-start gap-2 text-xs ${status.type === 'processing' ? 'bg-blue-50 text-blue-700 border border-blue-100' :
                  status.type === 'success' ? 'bg-emerald-50 text-emerald-700 border border-emerald-100' :
                    'bg-red-50 text-red-700 border border-red-100'
                  }`}>
                  {status.type === 'processing' && <AlertCircle className="w-4 h-4 animate-pulse shrink-0 mt-0.5" />}
                  {status.type === 'success' && <CheckCircle className="w-4 h-4 shrink-0 mt-0.5" />}
                  {status.type === 'error' && <FileWarning className="w-4 h-4 shrink-0 mt-0.5" />}
                  <span className="font-semibold leading-normal">{status.message}</span>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Modal Editar Equipo */}
      {editingEquipo && (
        <EditarEquipoModal
          equipo={editingEquipo}
          onClose={() => setEditingEquipo(null)}
        />
      )}

      {/* Modal Código QR */}
      {qrModalData && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-2xl border border-gray-200 shadow-2xl w-full max-w-sm flex flex-col overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            {/* Header */}
            <div className="flex justify-between items-center px-4 py-3 bg-[#25306B] text-white">
              <h2 className="text-sm font-bold flex items-center gap-1.5">
                <QrCode size={16} /> Etiqueta QR del Equipo
              </h2>
              <button
                onClick={() => setQrModalData(null)}
                className="text-white/80 hover:text-white hover:bg-white/10 px-2 py-0.5 rounded-lg transition-colors cursor-pointer text-lg font-bold"
              >
                &times;
              </button>
            </div>

            {/* Body */}
            <div className="p-5 flex flex-col items-center gap-4 text-xs">
              <div className="bg-slate-50 border border-slate-200 p-3 rounded-xl shadow-inner flex flex-col items-center justify-center gap-3">
                <QRCodeSVG
                  id="qr-code-svg"
                  value={`${window.location.origin}/qr-info?q=${encodeURIComponent(encodeQRData('E', qrModalData.id))}`}
                  size={180}
                  level="H"
                  includeMargin={true}
                />
                <a 
                  href={`${window.location.origin}/qr-info?q=${encodeURIComponent(encodeQRData('E', qrModalData.id))}`}
                  target="_blank" 
                  rel="noopener noreferrer"
                  className="text-[#006BB9] hover:underline font-bold text-xs flex items-center justify-center gap-1.5 bg-blue-50 px-3 py-1.5 rounded-lg w-full border border-blue-100"
                  title="Abrir información en nueva pestaña"
                >
                  <QrCode size={14} /> Abrir Link del Código QR
                </a>
              </div>

              {/* Detalles */}
              <div className="w-full bg-slate-50 border border-gray-150 p-3 rounded-xl space-y-1.5">
                <div>
                  <span className="font-bold text-[#25306B]">Descripción:</span>{' '}
                  <span className="text-gray-705 font-medium">{qrModalData['Descripción del Bien'] || '—'}</span>
                </div>
                <div>
                  <span className="font-bold text-[#25306B]">Marca:</span>{' '}
                  <span className="text-gray-705 font-medium">{qrModalData['Marca'] || '—'}</span>
                </div>
                <div>
                  <span className="font-bold text-[#25306B]">Modelo:</span>{' '}
                  <span className="text-gray-705 font-medium">{qrModalData['Modelo'] || '—'}</span>
                </div>
                <div>
                  <span className="font-bold text-[#25306B]">N° Serie:</span>{' '}
                  <span className="text-gray-755 font-mono font-bold">{qrModalData['Nº de serie'] || '—'}</span>
                </div>
                <div>
                  <span className="font-bold text-[#25306B]">Usuario:</span>{' '}
                  <span className="text-gray-705 font-medium">
                    {isAvailable(qrModalData['Usuario']) ? 'Disponible' : qrModalData['Usuario']}
                  </span>
                </div>
              </div>

              {/* Botones de acción */}
              <div className="flex gap-2 w-full mt-1">
                <button
                  onClick={() => handleDownloadQR(qrModalData)}
                  className="flex-1 py-2 bg-slate-100 hover:bg-slate-200 text-gray-700 font-bold rounded-xl border border-gray-300 flex items-center justify-center gap-1.5 transition-colors cursor-pointer"
                >
                  <Download size={14} /> Descargar PNG
                </button>
                <button
                  onClick={() => handlePrintQR(qrModalData)}
                  className="flex-1 py-2 bg-[#006BB9] hover:bg-[#25306B] text-white font-bold rounded-xl flex items-center justify-center gap-1.5 transition-colors shadow-sm cursor-pointer"
                >
                  <Printer size={14} /> Imprimir
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
