import { useEffect, useMemo, useRef, useState } from 'react';
import { SuperDoc } from 'superdoc';
import 'superdoc/style.css';
import { renderAsync } from 'docx-preview';
import JSZip from 'jszip';
import { api } from './lib/api';
import { ChevronLeft, Download, FileText, Loader2, Save, Eye, Trash2 } from 'lucide-react';

type Template = { id: string; name: string; filename: string; content_type: string; };
type FieldKey =
  | 'nombre' | 'tipoDocumento' | 'numeroDocumento' | 'programa' | 'periodoAcademico'
  | 'referenciaPago' | 'anio' | 'periodo' | 'totalLiquidado' | 'descuentos'
  | 'saldoFavor' | 'totalPagar' | 'estado' | 'fechaExpedicion' | 'analista' | 'proyecto'
  | 'nombreResponsableFirma' | 'cargoResponsableFirma';

type Field = { key: FieldKey; label: string; value: string; group: 'student' | 'liquidation' | 'signatures'; sample?: string };

const SIGNATURE_FIELDS: Field[] = [
  { key: 'fechaExpedicion', label: 'Fecha de expedición', value: '11/08/2026', group: 'signatures' },
  { key: 'analista', label: 'Nombre del analista', value: 'Daniel Samuel Moreno', group: 'signatures' },
  { key: 'proyecto', label: 'Proyecto / Revisó', value: 'Daniel Samuel Moreno / Analista de Apoyo Financiero.', group: 'signatures' },
  { key: 'nombreResponsableFirma', label: 'Nombre responsable de firma', value: 'Javier Alexander Mendez Tovar', group: 'signatures' },
  { key: 'cargoResponsableFirma', label: 'Cargo responsable de firma', value: 'Líder de Planeación Financiera.', group: 'signatures' },
];

const NORMAL_FIELDS: Field[] = [
  { key: 'nombre', label: 'Nombre completo', value: 'GIL RUBIO DIANA KATALINA', group: 'student' },
  { key: 'tipoDocumento', label: 'Tipo de documento', value: 'Cédula de ciudadanía', group: 'student' },
  { key: 'numeroDocumento', label: 'Número', value: '1014738616', group: 'student' },
  { key: 'programa', label: 'Programa', value: 'DISEÑO VISUAL', group: 'student' },
  { key: 'periodoAcademico', label: 'Periodo académico', value: '2026 - 1', group: 'student' },
  { key: 'referenciaPago', label: 'Referencia de pago', value: '946740 - 7', group: 'liquidation' },
  { key: 'anio', label: 'Año', value: '2026', group: 'liquidation' },
  { key: 'periodo', label: 'Periodo', value: '1', group: 'liquidation' },
  { key: 'totalLiquidado', label: 'Total liquidado', value: '4.515.200,00', group: 'liquidation' },
  { key: 'descuentos', label: 'Total descuentos', value: '903.000,00', group: 'liquidation' },
  { key: 'saldoFavor', label: 'Valor saldos a favor', value: '-', group: 'liquidation' },
  { key: 'totalPagar', label: 'Valor total por pagar', value: '3.612.200,00', group: 'liquidation' },
  { key: 'estado', label: 'Estado', value: 'PAGADO', group: 'liquidation' },
  ...SIGNATURE_FIELDS,
];

const FINANCING_FIELDS: Field[] = [
  { ...NORMAL_FIELDS[0], value: 'MORA CAMPOS VALERY' },
  { ...NORMAL_FIELDS[2], value: '1126805074' },
  { ...NORMAL_FIELDS[3], value: 'PROFESIONAL EN LENGUAS' },
  { ...NORMAL_FIELDS[4], value: '2026 - 2' },
  { ...NORMAL_FIELDS[5], value: '9847138' },
  { ...NORMAL_FIELDS[6], value: '2026' },
  { ...NORMAL_FIELDS[7], value: '2' },
  { ...NORMAL_FIELDS[8], value: '4.771.764,00' },
  { ...NORMAL_FIELDS[9], value: '-' },
  { ...NORMAL_FIELDS[11], value: '4.771.764,00' },
  { ...NORMAL_FIELDS[12], value: 'PENDIENTE' },
  { ...SIGNATURE_FIELDS[0], value: '04/08/2026' },
  ...SIGNATURE_FIELDS.slice(1),
];

const MULTI_FIELDS: Field[] = [
  { ...NORMAL_FIELDS[0], value: 'CASTRO VILLANUEVA GIONNY LEANDRO' },
  { ...NORMAL_FIELDS[2], value: '1000937521' },
  { ...NORMAL_FIELDS[3], value: 'TÉCNICA PROFESIONAL EN RECONOCIMIENTO ADUANERO' },
  { ...NORMAL_FIELDS[4], value: '2022 - 1' },
  { ...NORMAL_FIELDS[5], value: '691015 - 2' },
  { ...NORMAL_FIELDS[6], value: '2022' },
  { ...NORMAL_FIELDS[7], value: '1' },
  { ...NORMAL_FIELDS[8], value: '3.120.000,00' },
  { ...NORMAL_FIELDS[9], value: '468.000,00' },
  { ...NORMAL_FIELDS[11], value: '2.652.000,00' },
  { ...NORMAL_FIELDS[12], value: 'PAGADO' },
  { ...SIGNATURE_FIELDS[0], value: '18/08/2026' },
  ...SIGNATURE_FIELDS.slice(1),
];

function decodeBase64(value: string) {
  const clean = value.replace(/^data:[^,]+,/, '');
  const binary = atob(clean);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function fieldsForTemplate(name: string): Field[] {
  if (/financiaci[oó]n/i.test(name)) return FINANCING_FIELDS;
  if (/varios/i.test(name)) return MULTI_FIELDS;
  return NORMAL_FIELDS;
}

function groupLabel(group: Field['group']) {
  if (group === 'student') return 'Datos del estudiante';
  if (group === 'liquidation') return 'Datos de la liquidación';
  return 'Fecha y firmas';
}

function isMultiTemplate(name: string) { return /varios/i.test(name); }

export default function SuperDocCertificateEditor({ template, adminPassword, onBack, onChanged }: {
  template: Template; adminPassword: string; onBack: () => void; onChanged: () => Promise<void>;
}) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<HTMLDivElement | null>(null);
  const instanceRef = useRef<any>(null);
  const originalBytesRef = useRef<Uint8Array | null>(null);
  const fallbackBytesRef = useRef<Uint8Array | null>(null);
  const fallbackAttemptedRef = useRef(false);
  const [loading, setLoading] = useState(true);
  const [ready, setReady] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [fields, setFields] = useState<Field[]>(() => fieldsForTemplate(template.name));
  const [zoom, setZoom] = useState(100);
  const [fallbackMode, setFallbackMode] = useState(false);

  const groups = useMemo(() => ['student', 'liquidation', 'signatures'] as const, []);

  useEffect(() => { setFields(fieldsForTemplate(template.name)); }, [template.name]);

  const replaceInDocument = async (oldValue: string, newValue: string) => {
    const doc = instanceRef.current?.activeEditor?.doc;
    if (!doc || !oldValue || oldValue === newValue) return;
    try {
      const match = await doc.query.match({ select: { type: 'text', pattern: oldValue }, require: 'first' });
      const item = match.items?.[0];
      if (!item || item.matchKind !== 'text') return;
      await doc.replace({ target: item.target, text: newValue }, { expectedRevision: match.evaluatedRevision });
    } catch (e) { console.warn('No fue posible actualizar el campo en DOCX', oldValue, e); }
  };

  const replaceInDocxPackage = async (bytes: Uint8Array, replacements: Array<[string, string]>) => {
    const zip = await JSZip.loadAsync(bytes);
    const names = Object.keys(zip.files).filter(name => name.endsWith('.xml') && (name.startsWith('word/') || name.startsWith('customXml/')));
    for (const name of names) {
      const entry = zip.file(name);
      if (!entry) continue;
      let xml = await entry.async('string');
      for (const [from, to] of replacements) if (from && from !== to) xml = xml.split(from).join(to);
      zip.file(name, xml);
    }
    return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
  };

  const buildFallbackDocx = async (nextFields: Field[]) => {
    if (!originalBytesRef.current) return null;
    const values = Object.fromEntries(nextFields.map(field => [field.key, field.value])) as Record<string, string>;
    const replacements: Array<[string, string]> = [];
    const defaults = fieldsForTemplate(template.name);
    for (const field of defaults) {
      const next = values[field.key];
      if (field.value && next !== undefined && field.value !== next) replacements.push([field.value, next]);
    }
    replacements.push(['Javier Alexander Mendez Tovar', values.nombreResponsableFirma || 'Javier Alexander Mendez Tovar']);
    replacements.push(['Líder de Planeación Financiera.', values.cargoResponsableFirma || 'Líder de Planeación Financiera.']);
    return replaceInDocxPackage(originalBytesRef.current, replacements);
  };

  const renderFallback = async (bytes: Uint8Array) => {
    if (!hostRef.current) return;
    hostRef.current.innerHTML = '';
    await renderAsync(bytes, hostRef.current, undefined, { className: 'docx-preview-container', inWrapper: false, breakPages: true, ignoreWidth: false, ignoreHeight: false });
    fallbackBytesRef.current = bytes;
  };

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setLoading(true); setReady(false); setError(''); setMessage(''); setFallbackMode(false); fallbackAttemptedRef.current = false;
      try {
        const result = await api.get('/api/certificados/plantillas/' + encodeURIComponent(template.id) + '/file');
        const encoded = result?.data?.data;
        if (!encoded) throw new Error('La plantilla no contiene datos.');
        const bytes = decodeBase64(encoded);
        originalBytesRef.current = bytes;
        if (!hostRef.current || !toolbarRef.current) throw new Error('No se encontró el contenedor del editor.');

        const instance = new SuperDoc({
          selector: hostRef.current,
          document: new Blob([bytes.buffer as ArrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }),
          documentMode: 'editing',
          pagination: true,
          toolbar: toolbarRef.current,
          ui: { search: true },
          title: template.name,
          onReady: ({ superdoc }: any) => {
            if (cancelled) return;
            instanceRef.current = superdoc;
            setReady(true); setLoading(false);
          },
          onContentError: async ({ error: detail }: any) => {
            console.error('SuperDoc content error', detail);
            if (!cancelled && isMultiTemplate(template.name) && !fallbackAttemptedRef.current) {
              fallbackAttemptedRef.current = true;
              try {
                try { instanceRef.current?.destroy?.(); } catch {}
                instanceRef.current = null;
                setFallbackMode(true); setError('');
                await renderFallback(bytes);
                setReady(true); setLoading(false);
                setMessage('Vista compatible activada para esta plantilla DOCX.');
              } catch (fallbackError) {
                console.error('DOCX fallback error', fallbackError);
                setError('No fue posible abrir la plantilla DOCX ni activar la vista compatible.'); setLoading(false);
              }
            } else if (!cancelled) { setError('SuperDoc no pudo interpretar esta plantilla DOCX.'); setLoading(false); }
          },
          onException: async ({ error: detail }: any) => {
            console.error('SuperDoc exception', detail);
            if (!cancelled && isMultiTemplate(template.name) && !fallbackAttemptedRef.current) {
              fallbackAttemptedRef.current = true;
              try {
                try { instanceRef.current?.destroy?.(); } catch {}
                instanceRef.current = null;
                setFallbackMode(true); setError('');
                await renderFallback(bytes);
                setReady(true); setLoading(false);
                setMessage('Vista compatible activada para esta plantilla DOCX.');
              } catch (fallbackError) {
                console.error('DOCX fallback error', fallbackError);
                setError('No fue posible abrir la plantilla DOCX ni activar la vista compatible.'); setLoading(false);
              }
            } else if (!cancelled) { setError('Error técnico al abrir la plantilla DOCX.'); setLoading(false); }
          },
        });
        instanceRef.current = instance;
      } catch (e) {
        console.error('Certificate editor load error', e);
        if (!cancelled) { setError(e instanceof Error ? e.message : 'No fue posible cargar la plantilla.'); setLoading(false); }
      }
    };
    void load();
    return () => {
      cancelled = true;
      try { instanceRef.current?.destroy?.(); } catch {}
      instanceRef.current = null;
      originalBytesRef.current = null;
      fallbackBytesRef.current = null;
    };
  }, [template.id, template.name]);

  const updateField = async (key: FieldKey, value: string) => {
    const current = fields.find(field => field.key === key);
    const nextFields = fields.map(field => field.key === key ? { ...field, value } : field);
    setFields(nextFields);
    if (fallbackMode) {
      try {
        const nextBytes = await buildFallbackDocx(nextFields);
        if (nextBytes) await renderFallback(nextBytes);
      } catch (e) { console.error('Fallback DOCX update error', e); setMessage('No se pudo actualizar la vista del DOCX.'); }
      return;
    }
    if (current) await replaceInDocument(current.value, value);
  };

  const getCurrentDocxBlob = async () => {
    if (fallbackMode) {
      const bytes = fallbackBytesRef.current || await buildFallbackDocx(fields);
      if (!bytes) throw new Error('No hay DOCX disponible.');
      return new Blob([bytes.buffer as ArrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    }
    const instance = instanceRef.current;
    if (!instance || !ready) throw new Error('El documento todavía no está listo.');
    const blob = await instance.export({ exportType: ['docx'], triggerDownload: false });
    if (!(blob instanceof Blob)) throw new Error('export-invalid');
    return blob;
  };

  const exportDocx = async () => {
    if (!ready) return;
    setBusy(true); setMessage('');
    try {
      const blob = await getCurrentDocxBlob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = template.name.replace(/[^a-zA-Z0-9_-]+/g, '_') + '_editado.docx';
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
      setMessage('Word descargado correctamente.');
    } catch (e) { console.error('SuperDoc DOCX export error', e); setMessage('No fue posible exportar el Word.'); }
    finally { setBusy(false); }
  };

  const saveBackup = async () => {
    if (!ready || !adminPassword) return false;
    setBusy(true);
    try {
      let html = '';
      if (!fallbackMode) { const doc = instanceRef.current?.activeEditor?.doc; html = doc ? await doc.getHtml({}) : ''; }
      await api.post('/api/certificados/respaldos', {
        password: adminPassword, template_id: template.id, template_name: template.name,
        filename: template.name + '_respaldo', fields: Object.fromEntries(fields.map(field => [field.key, field.value])), html,
      });
      await onChanged(); setMessage('Respaldo guardado correctamente.'); return true;
    } catch (e) { console.error('Certificate backup error', e); setMessage('No se pudo guardar el respaldo.'); return false; }
    finally { setBusy(false); }
  };

  const preview = async () => {
    if (!ready || busy) return;
    setBusy(true); setMessage('Actualizando vista previa…');
    try {
      if (fallbackMode) {
        const bytes = await buildFallbackDocx(fields);
        if (bytes) await renderFallback(bytes);
      }
      setMessage('Vista previa actualizada.');
    } catch (e) { console.error('Preview error', e); setMessage('No se pudo actualizar la vista previa.'); }
    finally { setBusy(false); }
  };

  const generatePdf = async () => {
    if (!ready) return;
    setBusy(true); setMessage('Preparando PDF…');
    try {
      const ok = await saveBackup();
      if (!ok) return;
      window.print();
      setMessage('Se abrió la impresión. Selecciona “Guardar como PDF”.');
    } finally { setBusy(false); }
  };

  return (
    <div className='superdoc-certificate-shell'>
      <div className='superdoc-certificate-titlebar'>
        <button type='button' className='certificate-back-button' onClick={onBack}><ChevronLeft size={17}/> Certificados</button>
        <div className='superdoc-title-copy'><h2>{template.name}</h2><p>Edita los campos del certificado y genera el PDF.</p></div>
        <button type='button' className='certificate-settings-button' title='Configuración'><span>⚙</span></button>
      </div>
      <div className='superdoc-certificate-layout'>
        <aside className='superdoc-fields-panel'>
          <div className='superdoc-fixed-notice'><span>ⓘ</span><div>El encabezado y el pie de página se mantienen automáticamente en todas las plantillas.</div></div>
          {groups.map(group => {
            const groupFields = fields.filter(field => field.group === group);
            if (!groupFields.length) return null;
            return <section className='superdoc-field-group' key={group}><h3>{groupLabel(group)}</h3><div className='superdoc-field-grid'>
              {groupFields.map(field => <label key={field.key} className={['numeroDocumento','tipoDocumento','referenciaPago','anio','periodo','descuentos','saldoFavor','totalPagar','estado','fechaExpedicion','analista'].includes(field.key) ? 'compact-field' : ''}>
                <span>{field.label} <b>*</b></span>
                {field.key === 'tipoDocumento' ? <select value={field.value} onChange={e => void updateField(field.key, e.target.value)}><option>Cédula de ciudadanía</option><option>Tarjeta de identidad</option><option>Cédula de extranjería</option><option>Pasaporte</option></select>
                : field.key === 'estado' ? <select value={field.value} onChange={e => void updateField(field.key, e.target.value)}><option>PAGADO</option><option>PENDIENTE</option><option>ANULADO</option><option>VENCIDO</option></select>
                : <input value={field.value} onChange={e => void updateField(field.key, e.target.value)} />}
              </label>)}
            </div></section>;
          })}
          <div className='superdoc-left-actions'>
            <button type='button' className='secondary' disabled={!ready || busy} onClick={() => void preview()}><Eye size={15}/> Vista previa</button>
            <button type='button' className='primary' disabled={!ready || busy} onClick={() => void generatePdf()}><FileText size={15}/> Generar PDF</button>
            <button type='button' className='secondary' disabled={busy} onClick={() => setFields(fieldsForTemplate(template.name))}><Trash2 size={15}/> Limpiar campos</button>
          </div>
        </aside>
        <section className='superdoc-document-panel'>
          {!fallbackMode && <div ref={toolbarRef} className='superdoc-toolbar' />}
          <div className='superdoc-stage' style={{ ['--certificate-zoom' as string]: zoom / 100 }}>
            {loading && <div className='superdoc-loading'><Loader2 className='spin' size={25}/> Cargando documento original…</div>}
            {error && <div className='superdoc-error'><b>No se pudo abrir la plantilla.</b><span>{error}</span></div>}
            <div ref={hostRef} className='superdoc-document-host' />
          </div>
          <div className='superdoc-document-footer'>
            <button type='button' className='page-nav' aria-label='Página anterior'>‹</button><span>Página 1 de 1</span><button type='button' className='page-nav' aria-label='Página siguiente'>›</button>
            <div className='zoom-controls'><button type='button' onClick={() => setZoom(value => Math.max(70, value - 10))}>−</button><span>{zoom}%</span><button type='button' onClick={() => setZoom(value => Math.min(120, value + 10))}>＋</button></div>
            <button type='button' className='secondary superdoc-save-button' disabled={!ready || busy} onClick={() => void saveBackup()}><Save size={14}/> Guardar respaldo</button>
            <button type='button' className='secondary superdoc-word-button' disabled={!ready || busy} onClick={() => void exportDocx()}><Download size={14}/> Descargar Word</button>
          </div>
          {message && <div className='superdoc-message'>{message}</div>}
        </section>
      </div>
    </div>
  );
}
