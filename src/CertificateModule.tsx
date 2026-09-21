import { Component, type ErrorInfo, type ReactNode, useEffect, useMemo, useRef, useState } from 'react';
import { api } from './lib/api';
import { FileText, Upload, RefreshCw, Download, Settings2, Trash2, Pencil, Plus, X, Check, ChevronLeft, Save, RotateCcw, Archive, ExternalLink, AlertTriangle, Loader2 } from 'lucide-react';
import { renderAsync } from 'docx-preview';
import PdfCertificateEditor from './PdfCertificateEditor';
import mammoth from 'mammoth';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';

type CertificateTemplate = { id: string; name: string; description: string; filename: string; path: string; content_type: string; size?: number; created_at: number; updated_at?: number; };
type CertificateBackup = { id: string; template_id: string; template_name: string; filename?: string; fields: Record<string, string>; html: string; created_at: number; };
type Props = { templates: CertificateTemplate[]; backups: CertificateBackup[]; admin: boolean; adminPassword: string; onChanged: () => Promise<void>; };
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

function escapeHtml(value: string) { return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\"/g, '&quot;'); }
function formatBytes(value?: number) { if (!value) return '—'; if (value < 1024 * 1024) return `${Math.round(value / 1024)} KB`; return `${(value / (1024 * 1024)).toFixed(1)} MB`; }
function formatDate(value: unknown) { const numeric = typeof value === 'number' ? value : Number(value); const date = Number.isFinite(numeric) ? new Date(numeric) : null; if (!date || Number.isNaN(date.getTime())) return 'Fecha no disponible'; return new Intl.DateTimeFormat('es-CO', { dateStyle: 'short', timeStyle: 'short' }).format(date); }
function normalizeTemplate(template: CertificateTemplate): CertificateTemplate { return { id: String(template?.id || ''), name: String(template?.name || 'Certificado financiero'), description: String(template?.description || ''), filename: String(template?.filename || 'plantilla.docx'), path: String(template?.path || ''), content_type: String(template?.content_type || DOCX), size: typeof template?.size === 'number' ? template.size : Number(template?.size) || 0, created_at: typeof template?.created_at === 'number' ? template.created_at : Number(template?.created_at) || 0, updated_at: typeof template?.updated_at === 'number' ? template.updated_at : Number(template?.updated_at) || undefined }; }

function lockFixedCertificateAreas(root: HTMLElement) {
  const fixedNodes = Array.from(root.querySelectorAll<HTMLElement>('header, footer'));
  fixedNodes.forEach(node => {
    node.contentEditable = 'false';
    node.setAttribute('data-certificate-fixed', 'true');
    node.setAttribute('aria-label', 'Área institucional fija');
    node.classList.add('certificate-fixed-area');
    node.querySelectorAll<HTMLElement>('*').forEach(child => { child.contentEditable = 'false'; });
  });
}

function decorateTemplateFields(root: HTMLElement) {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes: Text[] = [];
  let current: Node | null;
  while ((current = walker.nextNode())) nodes.push(current as Text);
  nodes.forEach(node => {
    const text = node.nodeValue || '';
    let start = text.indexOf('{{');
    if (start < 0) return;
    const fragment = document.createDocumentFragment();
    let cursor = 0;
    while (start >= 0) {
      const end = text.indexOf('}}', start + 2);
      if (end < 0) break;
      if (start > cursor) fragment.appendChild(document.createTextNode(text.slice(cursor, start)));
      const key = text.slice(start + 2, end).trim();
      const span = document.createElement('span');
      span.className = 'certificate-variable';
      span.dataset.fieldKey = key;
      span.textContent = `{{${key}}}`;
      fragment.appendChild(span);
      cursor = end + 2;
      start = text.indexOf('{{', cursor);
    }
    if (cursor === 0) return;
    if (cursor < text.length) fragment.appendChild(document.createTextNode(text.slice(cursor)));
    node.parentNode?.replaceChild(fragment, node);
  });
}

function readTemplateFields(root: HTMLElement) {
  const fields: Record<string, string> = {};
  root.querySelectorAll<HTMLElement>('[data-field-key]').forEach(node => {
    const key = node.dataset.fieldKey;
    if (!key) return;
    const value = node.textContent || '';
    fields[key] = value === `{{${key}}}` ? '' : value;
  });
  return fields;
}

function writeTemplateField(root: HTMLElement, key: string, value: string) {
  root.querySelectorAll<HTMLElement>('[data-field-key]').forEach(node => {
    if (node.dataset.fieldKey === key) node.textContent = value || `{{${key}}}`;
  });
}

function downloadCertificateBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const DOCX_RENDER_OPTIONS = {
  className: 'docx',
  inWrapper: true,
  breakPages: true,
  ignoreLastRenderedPageBreak: false,
  renderHeaders: true,
  renderFooters: true,
  renderFootnotes: true,
  renderEndnotes: true,
  useBase64URL: true,
  experimental: false,
};

function clearCertificateDocxStyles() {
  document.head.querySelectorAll('style[data-crm-certificate-docx]').forEach(node => node.remove());
}

async function renderCertificateDocument(bytes: ArrayBuffer, host: HTMLElement) {
  host.innerHTML = '';
  host.style.display = 'block';
  clearCertificateDocxStyles();
  if (bytes.byteLength < 4) throw new Error('empty-template-file');
  const signature = new Uint8Array(bytes.slice(0, 4));
  const isZipPackage = signature[0] === 0x50 && signature[1] === 0x4b;
  if (!isZipPackage) throw new Error('invalid-docx-package');

  let rendered = false;
  try {
    // docx-preview debe recibir el HEAD real como contenedor de estilos.
    // Un DIV dentro de <head> puede ser reubicado por el navegador y dejar
    // las reglas generadas fuera del contexto correcto, provocando una hoja
    // completamente vacía aunque el DOCX sí se haya descargado.
    await renderAsync(bytes, host, document.head, DOCX_RENDER_OPTIONS);
    document.head.querySelectorAll('style').forEach(style => {
      if (style.textContent?.includes('docx')) {
        style.setAttribute('data-crm-certificate-docx', 'true');
      }
    });
    const pages = Array.from(host.querySelectorAll<HTMLElement>('.docx'));
    // La hoja del certificado es papel blanco real. Forzamos el fondo en línea
    // porque docx-preview puede insertar después sus propios estilos de página.
    pages.forEach(page => {
      page.style.setProperty('background', '#FFFFFF', 'important');
      page.style.setProperty('background-color', '#FFFFFF', 'important');
      page.style.setProperty('background-image', 'none', 'important');
      page.style.setProperty('opacity', '1', 'important');
      page.style.setProperty('filter', 'none', 'important');
      page.style.setProperty('mix-blend-mode', 'normal', 'important');
      // Algunos DOCX traen un sombreado aplicado a un contenedor que ocupa
      // prácticamente toda la hoja. No lo queremos: el papel debe ser blanco.
      const pageArea = Math.max(1, page.getBoundingClientRect().width * page.getBoundingClientRect().height);
      page.querySelectorAll<HTMLElement>('*').forEach(node => {
        const box = node.getBoundingClientRect();
        const area = box.width * box.height;
        const bg = getComputedStyle(node).backgroundColor;
        const image = getComputedStyle(node).backgroundImage;
        const hasSolidBackground = bg && bg !== 'transparent' && !bg.includes('rgba(0, 0, 0, 0)');
        if (area > pageArea * 0.45 && (hasSolidBackground || image !== 'none')) {
          node.style.setProperty('background', 'transparent', 'important');
          node.style.setProperty('background-color', 'transparent', 'important');
          node.style.setProperty('background-image', 'none', 'important');
        }
      });
    });
    await new Promise<void>(resolve => requestAnimationFrame(() => resolve()));
    rendered = pages.length > 0 && pages.some(page => {
      const box = page.getBoundingClientRect();
      const text = page.textContent?.replace(/\\s+/g, ' ').trim() || '';
      return box.width > 20 && box.height > 20 && Boolean(text || page.querySelector('img,table,svg'));
    });
    if (rendered) return;
  } catch (error) {
    console.warn('docx-preview failed, using compatibility renderer', error);
  }

  host.innerHTML = '';
  clearCertificateDocxStyles();
  document.head.querySelector('[data-crm-certificate-docx-style-host]')?.remove();

  const fallback = await mammoth.convertToHtml(
    { arrayBuffer: bytes },
    {
      includeDefaultStyleMap: true,
      convertImage: mammoth.images.inline(async image => ({
        src: `data:${image.contentType};base64,${await image.read('base64')}`,
      })),
    },
  );
  if (!fallback.value.trim()) throw new Error('docx-render-empty');
  host.innerHTML = fallback.value;
}

function CertificateEditor({ template: rawTemplate, templates: rawTemplates, backups: rawBackups, adminPassword, onBack, onChanged }: { template: CertificateTemplate; templates: CertificateTemplate[]; backups: CertificateBackup[]; adminPassword: string; onBack: () => void; onChanged: () => Promise<void> }) {
  const template = useMemo(() => normalizeTemplate(rawTemplate), [rawTemplate]);
  const templates = useMemo(() => (Array.isArray(rawTemplates) ? rawTemplates : []).map(normalizeTemplate).filter(item => item.id), [rawTemplates]);
  const backups = useMemo(() => (Array.isArray(rawBackups) ? rawBackups : []).filter(item => item && item.template_id), [rawBackups]);
  const [fields, setFields] = useState<Record<string, string>>({});
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(false);
  const [message, setMessage] = useState('');
  const [savingBackup, setSavingBackup] = useState(false);
  const [showBackups, setShowBackups] = useState(true);
  const [zoom, setZoom] = useState(82);
  const previewRef = useRef<HTMLDivElement | null>(null);
  const editorRef = useRef<HTMLDivElement | null>(null);
  const documentHostRef = useRef<HTMLDivElement | null>(null);
  const sourceBytesRef = useRef<ArrayBuffer | null>(null);

  useEffect(() => {
    let cancelled = false;
    const loadTemplate = async () => {
      setLoading(true); setLoadError(false); setMessage(''); setFields({});
      try {
        if (!template.id) throw new Error('missing-template-id');
        const result = await api.get(`/api/certificados/plantillas/${encodeURIComponent(template.id)}/file`);
        const encoded = result?.data?.data;
        if (typeof encoded !== 'string' || !encoded) throw new Error('missing-file-data');
        const binary = atob(encoded);
        const bytes = Uint8Array.from(binary, char => char.charCodeAt(0));
        sourceBytesRef.current = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
        if (cancelled || !documentHostRef.current) return;
        await renderCertificateDocument(sourceBytesRef.current, documentHostRef.current);
        decorateTemplateFields(documentHostRef.current);
        documentHostRef.current.contentEditable = 'true';
        documentHostRef.current.spellcheck = false;
        lockFixedCertificateAreas(documentHostRef.current);
        if (cancelled) return;
        setFields(readTemplateFields(documentHostRef.current));
      } catch (error) {
        console.error('Certificate template load error', error);
        if (!cancelled) { setLoadError(true); setMessage('No fue posible cargar esta plantilla.'); }
      } finally { if (!cancelled) setLoading(false); }
    };
    void loadTemplate();
    return () => { cancelled = true; };
  }, [template.id]);

  const syncEditor = () => { if (documentHostRef.current) setFields(readTemplateFields(documentHostRef.current)); };
  const replaceFieldInEditor = (key: string, value: string) => {
    const editor = editorRef.current;
    if (!editor) return;
    writeTemplateField(editor, key, value);
    syncEditor();
  };
  const runEditorCommand = (command: string, value?: string) => {
    editorRef.current?.focus();
    document.execCommand(command, false, value);
    syncEditor();
  };
  const restoreWorkingCopy = async () => {
    if (!documentHostRef.current || !sourceBytesRef.current) return;
    await renderCertificateDocument(sourceBytesRef.current, documentHostRef.current);
    decorateTemplateFields(documentHostRef.current);
    documentHostRef.current.contentEditable = 'true';
    documentHostRef.current.spellcheck = false;
    lockFixedCertificateAreas(documentHostRef.current);
    setFields(readTemplateFields(documentHostRef.current));
  };
  const hasData = Object.values(fields).some(value => value.trim()) || Boolean(documentHostRef.current?.classList.contains('certificate-user-edit'));
  const clearFields = async () => {
    if (!hasData || window.confirm('¿Deseas limpiar los datos ingresados?')) {
      await restoreWorkingCopy();
      setMessage('Datos limpiados. La plantilla maestra permanece intacta.');
    }
  };
  const confirmExit = () => { if (!hasData || window.confirm('¿Deseas salir? Los datos ingresados no se guardarán.')) onBack(); };
  const saveBackup = async () => {
    if (!adminPassword || !hasData) { setMessage('Diligencia al menos un dato antes de guardar un respaldo.'); return false; }
    setSavingBackup(true);
    try {
      await api.post('/api/certificados/respaldos', { password: adminPassword, template_id: template.id, template_name: template.name, filename: `${template.name}_certificado`, fields: documentHostRef.current ? readTemplateFields(documentHostRef.current) : fields, html: documentHostRef.current?.innerHTML || '' });
      await onChanged(); return true;
    } catch (error) { console.error('Certificate backup error', error); setMessage('No se pudo guardar el respaldo.'); return false; }
    finally { setSavingBackup(false); }
  };
  const downloadWord = async () => {
    if (!hasData) { setMessage('Diligencia los datos antes de descargar el Word.'); return; }
    if (!await saveBackup()) return;
    const workingHtml = documentHostRef.current?.innerHTML || editorRef.current?.innerHTML || '';
    const documentStyles = Array.from(document.head.querySelectorAll<HTMLStyleElement>('style[data-crm-certificate-docx]')).map(style => style.textContent || '').join('\n');
    const wordHtml = `<!DOCTYPE html><html><head><meta charset='utf-8'><style>${documentStyles}.docx-wrapper{background:#fff!important;padding:0!important}.docx{margin:0 auto!important;box-shadow:none!important}</style></head><body>${workingHtml}</body></html>`;
    const blob = new Blob([wordHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob); const anchor = document.createElement('a'); anchor.href = url; anchor.download = `${template.name.replace(/[^a-zA-Z0-9_-]+/g, '_')}_respaldo.doc`; document.body.appendChild(anchor); anchor.click(); anchor.remove(); URL.revokeObjectURL(url);
    setMessage('Respaldo Word descargado y guardado en Respaldos de certificados.');
  };
  const generatePdf = async () => {
    if (!previewRef.current || !hasData) { setMessage('Diligencia los campos obligatorios antes de generar el PDF.'); return; }
    setMessage('Generando PDF…');
    if (!await saveBackup()) return;
    try {
      const pages = Array.from(documentHostRef.current?.querySelectorAll<HTMLElement>('.docx') || []);
      if (!pages.length) throw new Error('no-rendered-pages');
      const pdf = new jsPDF('p', 'mm', 'a4');
      for (let index = 0; index < pages.length; index += 1) {
        const canvas = await html2canvas(pages[index], { scale: 2, useCORS: true, backgroundColor: '#ffffff' });
        const imageData = canvas.toDataURL('image/jpeg', 0.96);
        const pageWidth = 190;
        const pageHeight = canvas.height * pageWidth / canvas.width;
        if (index > 0) pdf.addPage();
        pdf.addImage(imageData, 'JPEG', 10, 10, pageWidth, pageHeight);
      }
      pdf.save(`${template.name.replace(/[^a-zA-Z0-9_-]+/g, '_')}_certificado.pdf`);
      await restoreWorkingCopy();
      setMessage('PDF generado y descargado. La plantilla quedó limpia para el siguiente certificado; la plantilla maestra permanece intacta.');
      await onChanged();
    } catch (error) { console.error('Certificate PDF error', error); setMessage('No fue posible generar el PDF. El respaldo quedó guardado para recuperar los datos.'); }
  };
  const restoreBackup = async (backup: CertificateBackup) => {
    const nextFields = backup.fields && typeof backup.fields === 'object' ? backup.fields : {};
    await restoreWorkingCopy();
    if (documentHostRef.current) Object.entries(nextFields).forEach(([key, value]) => writeTemplateField(documentHostRef.current!, key, String(value || '')));
    setMessage(`Respaldo del ${formatDate(backup.created_at)} cargado nuevamente.`);
  };
  const deleteBackup = async (backup: CertificateBackup) => {
    if (!adminPassword || !window.confirm(`¿Eliminar el respaldo de ${formatDate(backup.created_at)}?`)) return;
    try { await api.post(`/api/certificados/respaldos/${encodeURIComponent(backup.id)}/delete`, { password: adminPassword }); await onChanged(); setMessage('Respaldo eliminado.'); }
    catch (error) { console.error('Certificate backup delete error', error); setMessage('No se pudo eliminar el respaldo.'); }
  };
  const templateBackups = backups.filter(item => String(item.template_id) === template.id).sort((a, b) => Number(b.created_at || 0) - Number(a.created_at || 0));
  const fieldEntries = Object.entries(fields || {});
  if (loading) return <div className='certificate-safe-state'><Loader2 className='spin' size={28} /><h2>Cargando plantilla…</h2><p>Estamos preparando la plantilla institucional para edición.</p><button className='ghost' onClick={onBack}><ChevronLeft size={16} /> Volver a certificados</button></div>;
  if (loadError) return <div className='certificate-safe-state error'><AlertTriangle size={30} /><h2>No fue posible cargar esta plantilla</h2><p>La plantilla no está disponible o no pudo convertirse para la vista previa. Ningún dato de la plantilla maestra fue modificado.</p><button className='primary' onClick={onBack}><ChevronLeft size={16} /> Volver a certificados</button></div>;
  return <div className='certificate-editor'>
    <div className='certificate-toolbar'><button className='ghost' onClick={confirmExit}><ChevronLeft size={16} /> Certificados</button><div className='certificate-title'><span className='certificate-breadcrumb'>CERTIFICADOS FINANCIEROS</span><h1>{template.name}</h1></div><div className='certificate-actions'><button className='secondary' onClick={clearFields}><RotateCcw size={15} /> Limpiar</button><button className='secondary' disabled={savingBackup} onClick={() => void downloadWord()}><Save size={15} /> Descargar respaldo Word</button><button className='primary' disabled={savingBackup} onClick={() => void generatePdf()}><Download size={15} /> Generar certificado PDF</button></div></div>
    <div className='certificate-tabs' aria-label='Certificados disponibles'>{templates.map(item => <button key={item.id} className={`certificate-tab ${item.id === template.id ? 'active' : ''}`} onClick={() => { if (!hasData || window.confirm('¿Deseas cambiar de certificado? Los datos ingresados no se guardarán.')) window.dispatchEvent(new CustomEvent('crm-select-certificate', { detail: item.id })); }}><FileText size={14} /><span>{item.name}</span></button>)}</div>
    {message && <div className='certificate-message'>{message}</div>}
    <div className={`certificate-workspace certificate-workspace-${showBackups ? 'with-backups' : 'without-backups'}`}>
      <aside className='certificate-fields'>
        <div className='certificate-panel-title'><div><b>Datos del certificado</b><span>{fieldEntries.length ? `${fieldEntries.length} campos editables` : 'Edición directa'}</span></div><Archive size={17} /></div>
        <div className='certificate-fixed-notice'><b>🔒 Encabezado y pie protegidos</b><span>Estas áreas institucionales no se pueden editar y se conservan en todas las certificaciones.</span></div>
        <div className='wysiwyg-tools' aria-label='Herramientas de edición'>
          <button type='button' title='Deshacer' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('undo')}>↶</button>
          <button type='button' title='Rehacer' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('redo')}>↷</button>
          <span className='wysiwyg-divider' />
          <button type='button' title='Negrilla' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('bold')}><b>B</b></button>
          <button type='button' title='Cursiva' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('italic')}><i>I</i></button>
          <button type='button' title='Subrayado' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('underline')}><u>U</u></button>
          <span className='wysiwyg-divider' />
          <select title='Tipo de fuente' defaultValue='Arial' onChange={e => runEditorCommand('fontName', e.target.value)}><option>Arial</option><option>Calibri</option><option>Times New Roman</option><option>Georgia</option><option>Verdana</option></select>
          <select title='Tamaño' defaultValue='3' onChange={e => runEditorCommand('fontSize', e.target.value)}><option value='1'>10</option><option value='2'>11</option><option value='3'>12</option><option value='4'>14</option><option value='5'>18</option><option value='6'>24</option><option value='7'>32</option></select>
          <button type='button' title='Alinear a la izquierda' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('justifyLeft')}>≡</button>
          <button type='button' title='Centrar' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('justifyCenter')}>≡</button>
          <button type='button' title='Alinear a la derecha' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('justifyRight')}>≡</button>
          <button type='button' title='Lista con viñetas' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('insertUnorderedList')}>•</button>
          <button type='button' title='Lista numerada' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('insertOrderedList')}>1.</button>
          <button type='button' title='Salto de línea' onMouseDown={e => e.preventDefault()} onClick={() => runEditorCommand('insertHTML', '<br>')}>↵</button>
          <button type='button' title='Agregar fila debajo' onMouseDown={e => e.preventDefault()} onClick={() => {
            const cell = document.getSelection()?.anchorNode?.parentElement?.closest('td,th');
            const row = cell?.closest('tr');
            if (!row || !row.parentElement) return;
            const clone = row.cloneNode(true) as HTMLTableRowElement;
            clone.querySelectorAll('[data-field-key]').forEach(node => node.textContent = `{{${node.getAttribute('data-field-key') || ''}}}`);
            row.parentElement.insertBefore(clone, row.nextSibling);
            syncEditor();
          }}>＋fila</button>
          <button type='button' title='Eliminar fila actual' onMouseDown={e => e.preventDefault()} onClick={() => {
            const cell = document.getSelection()?.anchorNode?.parentElement?.closest('td,th');
            const row = cell?.closest('tr');
            if (row && row.parentElement && row.parentElement.children.length > 1) row.remove();
            syncEditor();
          }}>−fila</button>
          <span className='wysiwyg-divider' />
          <span className='wysiwyg-zoom-label'>Zoom</span>
          <select title='Zoom del documento' value={zoom} onChange={e => setZoom(Number(e.target.value))}><option value={70}>70%</option><option value={82}>82%</option><option value={92}>92%</option><option value={100}>100%</option></select>
        </div>
        {fieldEntries.length ? fieldEntries.map(([key, value]) => <label key={key}><span>{key.replace(/[_-]+/g, ' ')}</span><input value={value} onChange={e => { const next = e.target.value; setFields(current => ({ ...current, [key]: next })); replaceFieldInEditor(key, next); }} placeholder={`Ingresa ${key.toLowerCase()}`} /></label>) : <div className='certificate-empty-fields'><FileText size={22} /><b>Plantilla cargada</b><p>Edita directamente el documento para conservar y ajustar su formato visual.</p></div>}
        {fieldEntries.length > 0 && <button className='reset-link' onClick={clearFields}><RefreshCw size={14} /> Limpiar datos</button>}
      </aside>
      <section className='certificate-preview-area'><div className='preview-label'><span>DOCUMENTO ORIGINAL · VISTA EN TIEMPO REAL</span><small>🔒 Encabezado y pie fijos · contenido editable</small></div><div className='document-stage'><div ref={previewRef} className='document-page'><div ref={node => { editorRef.current = node; documentHostRef.current = node; }} className='document-content wysiwyg-editor' contentEditable suppressContentEditableWarning onInput={e => { e.currentTarget.classList.add('certificate-user-edit'); syncEditor(); }} style={{ zoom: `${zoom}%` }} /></div></div></section>
      {showBackups && <aside className='certificate-backups'><div className='certificate-backups-head'><div><b>Respaldos de certificados</b><span>{templateBackups.length} guardados</span></div><button className='icon-btn-small' title='Ocultar respaldos' onClick={() => setShowBackups(false)}><X size={14} /></button></div><p className='backup-helper'>Cada PDF o Word generado guarda una copia de trabajo para poder reutilizarla.</p>{templateBackups.length ? <div className='backup-list'>{templateBackups.map((backup, index) => { const safeFields = backup.fields && typeof backup.fields === 'object' ? backup.fields : {}; const backupKey = String(backup.id || `${template.id}-${backup.created_at || index}`); return <article className='backup-card' key={backupKey}><div className='backup-card-top'><FileText size={16} /><span>{formatDate(backup.created_at)}</span></div><b>{Object.values(safeFields).filter(value => typeof value === 'string' && value.trim()).slice(0, 2).join(' · ') || 'Certificado guardado'}</b><small>{Object.keys(safeFields).length} campos · {String(backup.filename || 'respaldo')}</small><div className='backup-actions'><button onClick={() => restoreBackup(backup)}><ExternalLink size={13} /> Utilizar nuevamente</button>{adminPassword && <button title='Eliminar respaldo' onClick={() => void deleteBackup(backup)}><Trash2 size={13} /></button>}</div></article>; })}</div> : <div className='backup-empty'><Archive size={22} /><b>Aún no hay respaldos</b><span>Cuando generes un PDF o Word aparecerán aquí.</span></div>}</aside>}
    </div>
    {!showBackups && <button className='show-backups' onClick={() => setShowBackups(true)}><Archive size={14} /> Mostrar respaldos</button>}
  </div>;
}

export function CertificateSidebar({ templates, admin }: { templates: CertificateTemplate[]; admin: boolean }) {
  const safeTemplates = Array.isArray(templates) ? templates.map(normalizeTemplate).filter(item => item.id) : [];
  const [selectedId, setSelectedId] = useState('');
  useEffect(() => { const selectHandler = (event: Event) => setSelectedId(String((event as CustomEvent<string>).detail || '')); const resetHandler = () => setSelectedId(''); window.addEventListener('crm-certificate-selected', selectHandler); window.addEventListener('crm-certificate-sidebar-reset', resetHandler); return () => { window.removeEventListener('crm-certificate-selected', selectHandler); window.removeEventListener('crm-certificate-sidebar-reset', resetHandler); }; }, []);
  return <><div className='side-title certificate-side-title'>CERTIFICADOS <span>{safeTemplates.length}</span></div><button className={`category ${selectedId === '' ? 'active' : ''}`} onClick={() => { setSelectedId(''); window.dispatchEvent(new CustomEvent('crm-select-certificate', { detail: '' })); }}><FileText size={15} /> Todos los certificados</button>{safeTemplates.map(template => <button key={template.id} className={`category certificate-side-item ${selectedId === template.id ? 'active' : ''}`} onClick={() => { setSelectedId(template.id); window.dispatchEvent(new CustomEvent('crm-select-certificate', { detail: template.id })); }}><FileText size={15} /><span>{template.name}</span></button>)}{admin && <><div className='certificate-side-divider' /><div className='certificate-side-label'>GESTIÓN</div><button className='add-side' onClick={() => window.dispatchEvent(new CustomEvent('crm-open-certificate-upload'))}><Plus size={16} /> Nueva plantilla</button><button className='certificate-side-settings' onClick={() => window.dispatchEvent(new CustomEvent('crm-open-certificate-manager'))}><Settings2 size={15} /> Configuración</button></>}</>;
}

class CertificateErrorBoundary extends Component<{ children: ReactNode }, { hasError: boolean }> {
  state = { hasError: false };
  static getDerivedStateFromError() { return { hasError: true }; }
  componentDidCatch(error: Error, info: ErrorInfo) { console.error('Certificate module render error', error, info); }
  render() { if (this.state.hasError) return <div className='certificate-safe-state error'><AlertTriangle size={30} /><h2>No fue posible mostrar Certificaciones Financieras</h2><p>Ocurrió un error al mostrar este módulo. El Banco de Respuestas no se ve afectado.</p><button className='primary' onClick={() => this.setState({ hasError: false })}>Reintentar</button></div>; return this.props.children; }
}

function CertificateModuleContent({ templates, backups, admin, adminPassword, onChanged }: Props) {
  const safeTemplates = Array.isArray(templates) ? templates.map(normalizeTemplate).filter(item => item.id) : [];
  const safeBackups = Array.isArray(backups) ? backups : [];
  const [selected, setSelected] = useState<CertificateTemplate | null>(null); const [managing, setManaging] = useState(false); const [showUpload, setShowUpload] = useState(false); const [editing, setEditing] = useState<CertificateTemplate | null>(null); const [name, setName] = useState(''); const [description, setDescription] = useState(''); const [file, setFile] = useState<File | null>(null); const [busy, setBusy] = useState(false); const [errorText, setErrorText] = useState('');
  useEffect(() => { if (selected && !safeTemplates.some(item => item.id === selected.id)) setSelected(null); }, [safeTemplates, selected]);
  useEffect(() => { const selectHandler = (event: Event) => { const id = String((event as CustomEvent<string>).detail || ''); if (!id) { setSelected(null); window.dispatchEvent(new CustomEvent('crm-certificate-sidebar-reset')); return; } const next = safeTemplates.find(item => item.id === id); if (next) setSelected(next); }; const managerHandler = () => { setManaging(true); setSelected(null); }; const uploadHandler = () => { setManaging(true); setSelected(null); setEditing(null); setName(''); setDescription(''); setFile(null); setErrorText(''); setShowUpload(true); }; window.addEventListener('crm-select-certificate', selectHandler); window.addEventListener('crm-open-certificate-manager', managerHandler); window.addEventListener('crm-open-certificate-upload', uploadHandler); return () => { window.removeEventListener('crm-select-certificate', selectHandler); window.removeEventListener('crm-open-certificate-manager', managerHandler); window.removeEventListener('crm-open-certificate-upload', uploadHandler); }; }, [safeTemplates]);
  useEffect(() => { window.dispatchEvent(new CustomEvent('crm-certificate-selected', { detail: selected?.id || '' })); }, [selected]);
  const openCreate = () => { setEditing(null); setName(''); setDescription(''); setFile(null); setErrorText(''); setShowUpload(true); };
  const openEdit = (template: CertificateTemplate) => { setEditing(template); setName(template.name); setDescription(template.description); setFile(null); setErrorText(''); setShowUpload(true); };
  const readBase64 = (selectedFile: File) => new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result || '').split(',')[1] || ''); reader.onerror = () => reject(new Error('read')); reader.readAsDataURL(selectedFile); });
  const saveTemplate = async () => { if (!admin || !adminPassword) return; if (!name.trim()) return setErrorText('Escribe un nombre para la plantilla.'); if (!editing && !file) return setErrorText('Selecciona una plantilla Word .docx o PDF.'); if (file && ((!file.name.toLowerCase().endsWith('.docx') && !file.name.toLowerCase().endsWith('.pdf')) || file.size > 10 * 1024 * 1024)) return setErrorText('La plantilla debe ser .docx o .pdf y pesar máximo 10 MB.'); setBusy(true); setErrorText(''); try { if (editing) { await api.put(`/api/certificados/plantillas/${encodeURIComponent(editing.id)}`, { password: adminPassword, name, description }); if (file) await api.post(`/api/certificados/plantillas/${encodeURIComponent(editing.id)}/replace`, { password: adminPassword, filename: file.name, contentType: file.type || DOCX, data: await readBase64(file) }); } else await api.post('/api/certificados/plantillas/upload', { password: adminPassword, name, description, filename: file!.name, contentType: file!.type || DOCX, data: await readBase64(file!) }); setShowUpload(false); setFile(null); await onChanged(); } catch (error) { console.error('Certificate template save error', error); setErrorText('No se pudo guardar la plantilla.'); } finally { setBusy(false); } };
  const deleteTemplate = async (template: CertificateTemplate) => { if (!admin || !adminPassword || !window.confirm(`¿Eliminar la plantilla “${template.name}”? Esta acción elimina solo esa plantilla y no afecta al Banco de Respuestas.`)) return; try { await api.post(`/api/certificados/plantillas/${encodeURIComponent(template.id)}/delete`, { password: adminPassword }); if (selected?.id === template.id) setSelected(null); await onChanged(); } catch (error) { console.error('Certificate template delete error', error); setErrorText('No se pudo eliminar la plantilla.'); } };
  if (selected) {
    const isPdfTemplate = selected.content_type === 'application/pdf' || /\.pdf$/i.test(selected.filename);
    if (isPdfTemplate) return <PdfCertificateEditor template={selected} templates={safeTemplates} backups={safeBackups} adminPassword={adminPassword} onBack={() => setSelected(null)} onChanged={onChanged} />;
    return <CertificateEditor template={selected} templates={safeTemplates} backups={safeBackups} adminPassword={adminPassword} onBack={() => setSelected(null)} onChanged={onChanged} />;
  }
  return <div className='certificates-module'><div className='certificates-hero'><div><span className='certificate-breadcrumb'>HERRAMIENTA INTERNA</span><h1>Certificados financieros</h1><p>Selecciona una plantilla y trabaja sobre una copia editable. El encabezado institucional y el pie de página se mantienen protegidos.</p></div><FileText size={38} /></div><div className='certificate-module-head'><div><h2>Certificados disponibles</h2><p>Selecciona una plantilla para abrir el editor o gestiona las plantillas desde la barra lateral.</p></div>{admin && <button className='primary' onClick={() => setManaging(current => !current)}><Settings2 size={16} /> {managing ? 'Cerrar gestión' : 'Gestionar plantillas'}</button>}</div>{safeTemplates.length > 0 && <div className='certificate-selector-bar'>{safeTemplates.map(item => <button key={item.id} className='certificate-selector-button' onClick={() => setSelected(item)}><FileText size={15} /><span>{item.name}</span></button>)}</div>}{managing && admin && <div className='template-manager'><div className='template-manager-head'><div><b>Plantillas Word</b><span>Administra las plantillas oficiales sin alterar el Banco de Respuestas.</span></div><button className='secondary' onClick={openCreate}><Plus size={16} /> Nueva plantilla</button></div>{safeTemplates.length ? <div className='template-list'>{safeTemplates.map(template => <article className='template-admin-card' key={template.id}><div className='template-file-icon'><FileText size={22} /></div><div className='template-admin-copy'><b>{template.name}</b><span>{template.filename} · {formatBytes(template.size)}</span><small>{template.description}</small></div><div className='template-admin-actions'><button onClick={() => openEdit(template)} title='Editar o reemplazar'><Pencil size={15} /></button><button onClick={() => void deleteTemplate(template)} title='Eliminar'><Trash2 size={15} /></button></div></article>)}</div> : <div className='template-empty'>Aún no hay plantillas. Sube los Word oficiales para comenzar.</div>}</div>}{!safeTemplates.length && !managing && <div className='certificate-no-templates'><FileText size={30} /><h3>Aún no hay certificados configurados</h3><p>Cuando tengas las plantillas Word, un administrador puede subirlas aquí.</p>{admin && <button className='primary' onClick={openCreate}><Upload size={16} /> Subir primera plantilla</button>}</div>}{showUpload && <div className='overlay'><div className='modal small certificate-upload-modal'><button className='close' onClick={() => setShowUpload(false)}><X /></button><div className='modal-icon'><FileText /></div><h2>{editing ? 'Editar plantilla' : 'Nueva plantilla Word'}</h2><p>La plantilla maestra existente no se modifica al diligenciar certificados. Solo se reemplaza si tú lo confirmas aquí.</p><label>Nombre<input autoFocus value={name} onChange={e => setName(e.target.value)} placeholder='Nombre real del certificado' /></label><label>Descripción<input value={description} onChange={e => setDescription(e.target.value)} placeholder='Descripción breve' /></label><label className='file-picker'><span>Plantilla Word o PDF {editing ? '(opcional para reemplazar)' : ''}</span><input type='file' accept='.docx,.pdf' onChange={e => setFile(e.target.files?.[0] || null)} />{file ? <b>{file.name}</b> : <small>{editing ? 'Conserva la plantilla actual si no seleccionas otro archivo.' : 'Acepta .docx o .pdf · máximo 10 MB'}</small>}</label>{errorText && <div className='form-error'>{errorText}</div>}<div className='modal-footer'><button className='ghost' onClick={() => setShowUpload(false)}>Cancelar</button><button className='primary' disabled={busy} onClick={() => void saveTemplate()}>{busy ? 'Guardando…' : <><Check size={16} /> Guardar plantilla</>}</button></div></div></div>}</div>;
}
export default function CertificateModule(props: Props) { return <CertificateErrorBoundary><CertificateModuleContent {...props} /></CertificateErrorBoundary>; }
