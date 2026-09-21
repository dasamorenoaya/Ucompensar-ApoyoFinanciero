import { useEffect, useMemo, useRef, useState } from 'react';
import { Archive, ChevronLeft, Download, FileText, Loader2, RotateCcw, Save, Trash2, X, ExternalLink } from 'lucide-react';
import { api } from './lib/api';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderAsync } from 'docx-preview';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';

type CertificateTemplate = { id: string; name: string; description: string; filename: string; path: string; content_type: string; size?: number; created_at: number; updated_at?: number; };
type CertificateBackup = { id: string; template_id: string; template_name: string; filename?: string; fields: Record<string, string>; html: string; created_at: number; };
type PdfField = { key: string; x: number; top: number; width: number; height: number; fontSize: number; bold?: boolean; align?: 'left' | 'center' | 'right'; };
type Props = { template: CertificateTemplate; templates: CertificateTemplate[]; backups: CertificateBackup[]; adminPassword: string; onBack: () => void; onChanged: () => Promise<void>; sourceUrl?: string; };

const PAGE_WIDTH = 612;
const PAGE_HEIGHT = 792;
export const NORMAL_REFERENCE_PDF = '/resources/certificacion-financiera-normal-referencia.pdf';

const NORMAL_FIELDS: PdfField[] = [
  { key: 'NOMBRE', x: 85.46, top: 265.9, width: 126.71, height: 12, fontSize: 10, bold: true },
  { key: 'DOCUMENTO', x: 425.35, top: 265.9, width: 82.84, height: 12, fontSize: 10, bold: true },
  { key: 'PERIODO_ADMISION', x: 206.81, top: 277.3, width: 46.59, height: 12, fontSize: 10, bold: true },
  { key: 'SEMESTRE', x: 320.83, top: 277.3, width: 46.7, height: 12, fontSize: 10, bold: true },
  { key: 'VALOR_MATRICULA', x: 85.46, top: 288.82, width: 113.29, height: 12, fontSize: 10, bold: true },
  { key: 'VALOR_MATRICULA_NUMERO', x: 235.37, top: 288.82, width: 65.6, height: 12, fontSize: 10 },
  { key: 'PROGRAMA', x: 275.57, top: 332.61, width: 65.92, height: 14, fontSize: 11, bold: true, align: 'center' },
  { key: 'REFERENCIA_PAGO', x: 102.5, top: 391.86, width: 32, height: 11, fontSize: 8, align: 'center' },
  { key: 'ANIO', x: 153.38, top: 391.86, width: 21.32, height: 11, fontSize: 8, align: 'center' },
  { key: 'PERIODO', x: 189.62, top: 391.86, width: 19.33, height: 11, fontSize: 8, align: 'center' },
  { key: 'TOTAL_LIQUIDADO', x: 235.01, top: 391.86, width: 37.28, height: 11, fontSize: 8, align: 'center' },
  { key: 'TOTAL_DESCUENTOS', x: 304.13, top: 391.86, width: 32, height: 11, fontSize: 8, align: 'center' },
  { key: 'SALDO_FAVOR', x: 370.03, top: 391.86, width: 26.6, height: 11, fontSize: 8, align: 'center' },
  { key: 'TOTAL_POR_PAGAR', x: 419.35, top: 391.86, width: 32, height: 11, fontSize: 8, align: 'center' },
  { key: 'ESTADO', x: 474.82, top: 391.86, width: 35.26, height: 11, fontSize: 8, bold: true, align: 'center' },
  { key: 'DIA_NUMERO', x: 271.13, top: 443.4, width: 27.97, height: 12, fontSize: 10 },
  { key: 'DIA_LETRAS', x: 303.41, top: 443.4, width: 33.22, height: 12, fontSize: 10 },
  { key: 'MES', x: 411.55, top: 443.4, width: 39.99, height: 12, fontSize: 10 },
  { key: 'ANIO_EXPEDICION', x: 467.74, top: 443.4, width: 40, height: 12, fontSize: 10 },
  { key: 'NOMBRE_PROYECTO', x: 86.9, top: 540.96, width: 173.37, height: 12, fontSize: 10, bold: true },
  { key: 'CARGO_PROYECTO', x: 86.9, top: 554.28, width: 53.31, height: 12, fontSize: 10, bold: true },
  { key: 'ANALISTA', x: 146.66, top: 569.12, width: 44.6, height: 10, fontSize: 8 },
];

const FIELD_LABELS: Record<string, string> = {
  NOMBRE: 'Nombre del estudiante', DOCUMENTO: 'Documento de identidad', PERIODO_ADMISION: 'Periodo de admisión', SEMESTRE: 'Semestre', VALOR_MATRICULA: 'Valor de matrícula en letras', VALOR_MATRICULA_NUMERO: 'Valor de matrícula en números', PROGRAMA: 'Programa académico', REFERENCIA_PAGO: 'Referencia de pago', ANIO: 'Año', PERIODO: 'Periodo', TOTAL_LIQUIDADO: 'Total liquidado', TOTAL_DESCUENTOS: 'Total descuentos', SALDO_FAVOR: 'Saldo a favor', TOTAL_POR_PAGAR: 'Total por pagar', ESTADO: 'Estado', DIA_NUMERO: 'Día en número', DIA_LETRAS: 'Día en letras', MES: 'Mes', ANIO_EXPEDICION: 'Año de expedición', NOMBRE_PROYECTO: 'Nombre / firma', CARGO_PROYECTO: 'Cargo / dependencia', ANALISTA: 'Analista',
};

function formatDate(value: unknown) { const numeric = typeof value === 'number' ? value : Number(value); const date = Number.isFinite(numeric) ? new Date(numeric) : null; if (!date || Number.isNaN(date.getTime())) return 'Fecha no disponible'; return new Intl.DateTimeFormat('es-CO', { dateStyle: 'short', timeStyle: 'short' }).format(date); }
function safeTemplate(template: CertificateTemplate): CertificateTemplate { return { id: String(template?.id || ''), name: String(template?.name || 'Certificado financiero'), description: String(template?.description || ''), filename: String(template?.filename || 'certificado.pdf'), path: String(template?.path || ''), content_type: String(template?.content_type || 'application/pdf'), size: Number(template?.size) || 0, created_at: Number(template?.created_at) || 0, updated_at: Number(template?.updated_at) || undefined }; }
function safeBackups(backups: CertificateBackup[]) { return (Array.isArray(backups) ? backups : []).filter(item => item && item.template_id); }

const DOCX_RENDER_OPTIONS = { className: 'docx', inWrapper: true, breakPages: true, ignoreLastRenderedPageBreak: false, renderHeaders: true, renderFooters: true, renderFootnotes: true, renderEndnotes: true, useBase64URL: true, experimental: false };

async function renderWordPreview(bytes: ArrayBuffer, host: HTMLElement) {
  host.innerHTML = '';
  const styleHost = document.createElement('div');
  styleHost.style.position = 'absolute'; styleHost.style.width = '0'; styleHost.style.height = '0'; styleHost.style.overflow = 'hidden';
  document.body.appendChild(styleHost);
  await renderAsync(bytes, host, styleHost, DOCX_RENDER_OPTIONS);
  styleHost.querySelectorAll('style').forEach(source => { const style = document.createElement('style'); style.setAttribute('data-crm-certificate-word-style','true'); style.textContent = source.textContent || ''; document.head.appendChild(style); });
  styleHost.remove();
  host.querySelectorAll<HTMLElement>('.docx').forEach(page => { page.style.setProperty('display','block','important'); page.style.setProperty('visibility','visible','important'); page.style.setProperty('opacity','1','important'); page.style.setProperty('background','#fff','important'); page.style.setProperty('background-color','#fff','important'); });
}

async function fetchArrayBuffer(url: string) { if (url.startsWith('/api/')) { const result = await api.get(url); const encoded = result?.data?.data; if (!encoded) throw new Error('pdf-missing-data'); const binary = atob(encoded); return Uint8Array.from(binary, char => char.charCodeAt(0)).buffer; } const response = await fetch(url, { cache: 'no-store' }); if (!response.ok) throw new Error(`pdf-fetch-${response.status}`); return await response.arrayBuffer(); }

function PdfPage({ canvasRef, fields, values, onChange, onFocus }: {
  canvasRef: (node: HTMLCanvasElement | null) => void;
  pageNumber: number;
  fields: PdfField[];
  values: Record<string, string>;
  onChange: (key: string, value: string) => void;
  onFocus: (node: HTMLDivElement | null) => void;
}) {
  return <div className='pdf-edit-page' style={{ position: 'relative', width: '612px', aspectRatio: '612 / 792', background: '#fff', overflow: 'hidden' }}>
    <canvas ref={canvasRef} style={{ position: 'absolute', inset: 0, width: '100%', height: '100%', display: 'block' }} />
    {fields.map(field => <div key={field.key} style={{ position: 'absolute', left: `${(field.x / PAGE_WIDTH) * 100}%`, top: `${(field.top / PAGE_HEIGHT) * 100}%`, width: `${(field.width / PAGE_WIDTH) * 100}%`, minHeight: `${(field.height / PAGE_HEIGHT) * 100}%`, display: 'flex', alignItems: 'center' }}>
      <input
        value={values[field.key] || ''}
        onChange={event => onChange(field.key, event.target.value)}
        onFocus={event => onFocus(event.currentTarget.parentElement)}
        aria-label={FIELD_LABELS[field.key] || field.key}
        style={{ width: '100%', height: '100%', border: '1px dashed rgba(255,104,1,.65)', background: 'rgba(255,249,242,.82)', color: '#171717', fontSize: `${Math.max(8, field.fontSize)}px`, fontWeight: field.bold ? 700 : 400, textAlign: field.align || 'left', padding: '1px 3px', outline: 'none', boxSizing: 'border-box' }}
      />
    </div>)}
  </div>;
}

export default function PdfCertificateEditor({ template: rawTemplate, templates: rawTemplates, backups: rawBackups, adminPassword, onBack, onChanged, sourceUrl }: Props) {
  const template = useMemo(() => safeTemplate(rawTemplate), [rawTemplate]);
  const templates = useMemo(() => (Array.isArray(rawTemplates) ? rawTemplates.map(safeTemplate).filter(item => item.id) : []), [rawTemplates]);
  const backups = useMemo(() => safeBackups(rawBackups), [rawBackups]);
  const [fields, setFields] = useState<Record<string, string>>({});
  const [loading, setLoading] = useState(true);
  const [errorText, setErrorText] = useState('');
  const [message, setMessage] = useState('');
  const [savingBackup, setSavingBackup] = useState(false);
  const [showBackups, setShowBackups] = useState(true);
  const [zoom, setZoom] = useState(82);
  const [activeEditor, setActiveEditor] = useState<HTMLDivElement | null>(null);
  const pageRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const wordRef = useRef<HTMLDivElement | null>(null);
  const sourceBytesRef = useRef<ArrayBuffer | null>(null);
  const [isWordTemplate, setIsWordTemplate] = useState(false);
  const isNormal = template.name.toLowerCase().includes('normal');
  const fieldDefs = isNormal ? NORMAL_FIELDS : [];
  const fieldEntries = Object.entries(fields);
  const hasData = fieldEntries.some(([, value]) => value.trim().length > 0);

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setLoading(true); setErrorText(''); setMessage(''); setFields({});
      try {
        const url = sourceUrl || (isNormal ? NORMAL_REFERENCE_PDF : `/api/certificados/plantillas/${encodeURIComponent(template.id)}/file`);
        const bytes = await fetchArrayBuffer(url);
        sourceBytesRef.current = bytes;
        if (cancelled) return;
        const signature = new Uint8Array(bytes.slice(0, 4));
        const isZipPackage = signature[0] === 0x50 && signature[1] === 0x4b;
        if (isZipPackage) {
          if (!wordRef.current) throw new Error('word-host-missing');
          await renderWordPreview(bytes, wordRef.current);
          wordRef.current.contentEditable = 'true';
          wordRef.current.spellcheck = false;
          setIsWordTemplate(true);
          return;
        }
        setIsWordTemplate(false);
        const pdf = await getDocument({ data: new Uint8Array(bytes), disableWorker: true }).promise;
        if (cancelled || pdf.numPages < 1) return;
        const page = await pdf.getPage(1);
        const viewport = page.getViewport({ scale: 2 });
        const canvas = canvasRef.current;
        if (!canvas) throw new Error('canvas-missing');
        const context = canvas.getContext('2d');
        if (!context) throw new Error('canvas-context-missing');
        canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
        await page.render({ canvasContext: context, viewport }).promise;
      } catch (error) { console.error('PDF certificate template load error', error); if (!cancelled) setErrorText('No fue posible cargar la plantilla PDF.'); }
      finally { if (!cancelled) setLoading(false); }
    };
    void load();
    return () => { cancelled = true; };
  }, [template.id, sourceUrl, isNormal]);

  const runCommand = (command: string, value?: string) => { activeEditor?.focus(); document.execCommand(command, false, value); };
  const updateField = (key: string, value: string) => setFields(current => ({ ...current, [key]: value }));
  const setFieldFromPanel = (key: string, value: string) => { updateField(key, value); const node = pageRef.current?.querySelector<HTMLElement>(`[data-field-key="${CSS.escape(key)}"]`); if (node && node.textContent !== value) node.textContent = value; };
  const clearCertificate = async () => {
    if (hasData && !window.confirm('¿Deseas limpiar los datos ingresados?')) return;
    setFields({});
    if (isWordTemplate && wordRef.current && sourceBytesRef.current) await renderWordPreview(sourceBytesRef.current, wordRef.current);
    pageRef.current?.querySelectorAll<HTMLElement>('[data-field-key]').forEach(node => { node.textContent = ''; });
    setMessage('Certificado limpio. La plantilla original permanece intacta.');
  };

  const saveBackup = async () => {
    if (!adminPassword || !hasData) { setMessage('Diligencia al menos un dato antes de guardar un respaldo.'); return false; }
    setSavingBackup(true);
    try {
      await api.post('/api/certificados/respaldos', { password: adminPassword, template_id: template.id, template_name: template.name, filename: template.name + '_certificado', fields, html: isWordTemplate ? (wordRef.current?.innerHTML || '') : (pageRef.current?.outerHTML || '') });
      await onChanged(); return true;
    } catch (error) { console.error('PDF certificate backup error', error); setMessage('No se pudo guardar el respaldo.'); return false; }
    finally { setSavingBackup(false); }
  };

  const generatePdf = async () => {
    const target = isWordTemplate ? wordRef.current : pageRef.current;
    if (!target || !hasData) { setMessage('Diligencia los campos antes de generar el PDF.'); return; }
    setMessage('Generando PDF…'); if (!await saveBackup()) return;
    try {
      const pdf = new jsPDF('p', 'mm', 'a4');
      if (isWordTemplate) {
        const pages = Array.from(target.querySelectorAll<HTMLElement>('.docx'));
        if (!pages.length) throw new Error('no-word-pages');
        for (let index = 0; index < pages.length; index += 1) {
          const canvas = await html2canvas(pages[index], { scale: 2, useCORS: true, backgroundColor: '#ffffff', logging: false });
          const width = 190;
          const height = canvas.height * width / canvas.width;
          if (index > 0) pdf.addPage();
          pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 10, 10, width, Math.min(height, 277));
        }
      } else {
        const canvas = await html2canvas(target, { scale: 2, useCORS: true, backgroundColor: '#ffffff', logging: false });
        pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, 210, 297);
      }
      pdf.save(template.name.replace(/[^a-zA-Z0-9_-]+/g, '_') + '_certificado.pdf');
      await clearCertificate();
      await onChanged();
      setMessage('PDF generado. El certificado quedó limpio para el siguiente estudiante.');
    } catch (error) { console.error('PDF certificate generation error', error); setMessage('No fue posible generar el PDF. El respaldo quedó guardado.'); }
  };

  const downloadWord = async () => {
    const target = isWordTemplate ? wordRef.current : pageRef.current;
    if (!target || !hasData) { setMessage('Diligencia los datos antes de descargar el Word.'); return; }
    if (!await saveBackup()) return;
    const imageCanvas = await html2canvas(target, { scale: 2, useCORS: true, backgroundColor: '#ffffff', logging: false });
    const image = imageCanvas.toDataURL('image/png');
    const html = "<!DOCTYPE html><html><head><meta charset='utf-8'><style>@page{size:A4;margin:0}html,body{margin:0;padding:0}img{width:210mm;display:block}</style></head><body><img src='" + image + "' alt='Certificado financiero'/></body></html>";
    const blob = new Blob([html], { type: 'application/msword' }); const url = URL.createObjectURL(blob); const anchor = document.createElement('a'); anchor.href = url; anchor.download = template.name.replace(/[^a-zA-Z0-9_-]+/g, '_') + '_respaldo.doc'; document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url); setMessage('Respaldo Word descargado y guardado.');
  };

  const restoreBackup = (backup: CertificateBackup) => { const next = backup.fields && typeof backup.fields === 'object' ? backup.fields : {}; setFields(Object.fromEntries(Object.entries(next).map(([key, value]) => [key, String(value ?? '')]))); window.requestAnimationFrame(() => pageRef.current?.querySelectorAll<HTMLElement>('[data-field-key]').forEach(node => { const key = node.dataset.fieldKey || ''; node.textContent = String(next[key] ?? ''); })); setMessage(`Respaldo del ${formatDate(backup.created_at)} cargado nuevamente.`); };
  const deleteBackup = async (backup: CertificateBackup) => { if (!adminPassword || !window.confirm(`¿Eliminar el respaldo de ${formatDate(backup.created_at)}?`)) return; try { await api.post(`/api/certificados/respaldos/${encodeURIComponent(backup.id)}/delete`, { password: adminPassword }); await onChanged(); setMessage('Respaldo eliminado.'); } catch (error) { console.error('PDF certificate backup delete error', error); setMessage('No se pudo eliminar el respaldo.'); } };
  const templateBackups = backups.filter(item => String(item.template_id) === template.id).sort((a, b) => Number(b.created_at || 0) - Number(a.created_at || 0));

  if (loading) return <div className='certificate-safe-state'><Loader2 className='spin' size={28} /><h2>Preparando plantilla PDF…</h2><p>Se está renderizando la página original para conservar su diseño.</p><button className='ghost' onClick={onBack}><ChevronLeft size={16} /> Volver a certificados</button></div>;
  if (errorText) return <div className='certificate-safe-state error'><FileText size={30} /><h2>No fue posible cargar esta plantilla</h2><p>{errorText} La plantilla maestra no se modifica.</p><button className='primary' onClick={onBack}><ChevronLeft size={16} /> Volver a certificados</button></div>;

  return <div className='certificate-editor pdf-certificate-editor'>
    <div className='certificate-toolbar'><button className='ghost' onClick={() => { if (!hasData || window.confirm('¿Deseas salir? Los datos ingresados no se guardarán.')) onBack(); }}><ChevronLeft size={16} /> Certificados</button><div className='certificate-title'><span className='certificate-breadcrumb'>CERTIFICADOS FINANCIEROS</span><h1>{template.name}</h1></div><div className='certificate-actions'><button className='secondary' onClick={clearCertificate}><RotateCcw size={15} /> Limpiar</button><button className='secondary' disabled={savingBackup} onClick={() => void downloadWord()}><Save size={15} /> Descargar respaldo Word</button><button className='primary' disabled={savingBackup} onClick={() => void generatePdf()}><Download size={15} /> Generar certificado PDF</button></div></div>
    <div className='certificate-tabs' aria-label='Certificados disponibles'>{templates.map(item => <button key={item.id} className={`certificate-tab ${item.id === template.id ? 'active' : ''}`} onClick={() => { if (!hasData || window.confirm('¿Deseas cambiar de certificado? Los datos ingresados no se guardarán.')) window.dispatchEvent(new CustomEvent('crm-select-certificate', { detail: item.id })); }}><FileText size={14} /><span>{item.name}</span></button>)}</div>
    {message && <div className='certificate-message'>{message}</div>}
    <div className={`certificate-workspace certificate-workspace-${showBackups ? 'with-backups' : 'without-backups'}`}>
      <aside className='certificate-fields'><div className='certificate-panel-title'><div><b>Datos del certificado</b><span>{fieldDefs.length} campos editables</span></div><Archive size={17} /></div>
        <div className='wysiwyg-tools pdf-wysiwyg-tools' aria-label='Herramientas de edición'><button type='button' title='Deshacer' onMouseDown={e => e.preventDefault()} onClick={() => runCommand('undo')}>↶</button><button type='button' title='Rehacer' onMouseDown={e => e.preventDefault()} onClick={() => runCommand('redo')}>↷</button><span className='wysiwyg-divider' /><button type='button' title='Negrilla' onMouseDown={e => e.preventDefault()} onClick={() => runCommand('bold')}><b>B</b></button><button type='button' title='Cursiva' onMouseDown={e => e.preventDefault()} onClick={() => runCommand('italic')}><i>I</i></button><button type='button' title='Subrayado' onMouseDown={e => e.preventDefault()} onClick={() => runCommand('underline')}><u>U</u></button><span className='wysiwyg-divider' /><select title='Tipo de fuente' defaultValue='Arial' onChange={e => runCommand('fontName', e.target.value)}><option>Arial</option><option>Calibri</option><option>Times New Roman</option><option>Verdana</option></select><select title='Tamaño de letra' defaultValue='3' onChange={e => runCommand('fontSize', e.target.value)}><option value='1'>10</option><option value='2'>11</option><option value='3'>12</option><option value='4'>14</option><option value='5'>18</option><option value='6'>24</option></select><button type='button' title='Alinear izquierda' onMouseDown={e => e.preventDefault()} onClick={() => runCommand('justifyLeft')}>≡</button><button type='button' title='Centrar' onMouseDown={e => e.preventDefault()} onClick={() => runCommand('justifyCenter')}>≡</button><button type='button' title='Alinear derecha' onMouseDown={e => e.preventDefault()} onClick={() => runCommand('justifyRight')}>≡</button><span className='wysiwyg-divider' /><span className='wysiwyg-zoom-label'>Zoom</span><select title='Zoom' value={zoom} onChange={e => setZoom(Number(e.target.value))}><option value={70}>70%</option><option value={82}>82%</option><option value={92}>92%</option><option value={100}>100%</option></select></div>
        {fieldDefs.map(field => <label key={field.key}><span>{FIELD_LABELS[field.key] || field.key}</span><input value={fields[field.key] || ''} onChange={event => setFieldFromPanel(field.key, event.target.value)} placeholder='Editar campo' /></label>)}
      </aside>
      <section className='certificate-preview-area'><div className='preview-label'><span>PLANTILLA ORIGINAL · EDICIÓN DIRECTA</span><small>Diseño institucional intacto · copia temporal</small></div><div className='document-stage pdf-document-stage'>{isWordTemplate ? <div ref={wordRef} className='certificate-word-web-editor' /> : <div ref={pageRef} className='pdf-page-stack' style={{ width: `${zoom}%` }}><PdfPage canvasRef={node => { canvasRef.current = node; }} pageNumber={1} fields={fieldDefs} values={fields} onChange={updateField} onFocus={setActiveEditor} /></div>}</div></section>
      {showBackups && <aside className='certificate-backups'><div className='certificate-backups-head'><div><b>Respaldos de certificados</b><span>{templateBackups.length} guardados</span></div><button className='icon-btn-small' title='Ocultar respaldos' onClick={() => setShowBackups(false)}><X size={14} /></button></div><p className='backup-helper'>Cada PDF o Word generado guarda una copia para poder reutilizarla.</p>{templateBackups.length ? <div className='backup-list'>{templateBackups.map((backup, index) => { const safe = backup.fields && typeof backup.fields === 'object' ? backup.fields : {}; const backupKey = String(backup.id || `${template.id}-${backup.created_at || index}`); return <article className='backup-card' key={backupKey}><div className='backup-card-top'><FileText size={16} /><span>{formatDate(backup.created_at)}</span></div><b>{Object.values(safe).filter(value => typeof value === 'string' && value.trim()).slice(0, 2).join(' · ') || 'Certificado guardado'}</b><small>{Object.keys(safe).length} campos · {String(backup.filename || 'respaldo')}</small><div className='backup-actions'><button onClick={() => restoreBackup(backup)}><ExternalLink size={13} /> Utilizar nuevamente</button>{adminPassword && <button title='Eliminar respaldo' onClick={() => void deleteBackup(backup)}><Trash2 size={13} /></button>}</div></article>; })}</div> : <div className='backup-empty'><Archive size={22} /><b>Aún no hay respaldos</b><span>Cuando generes un PDF o Word aparecerán aquí.</span></div>}</aside>}
    </div>
    {!showBackups && <button className='show-backups' onClick={() => setShowBackups(true)}><Archive size={14} /> Mostrar respaldos</button>}
  </div>;
}
