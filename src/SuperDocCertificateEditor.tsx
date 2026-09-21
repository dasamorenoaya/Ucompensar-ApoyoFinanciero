import { useEffect, useRef, useState } from 'react';
import { SuperDoc } from 'superdoc';
import 'superdoc/style.css';
import { api } from './lib/api';
import { Download, FileText, Loader2, Save, Trash2, ChevronLeft } from 'lucide-react';

type Template = { id: string; name: string; filename: string; content_type: string; };
type Backup = { id: string; template_id: string; template_name: string; fields: Record<string,string>; html: string; created_at: number; };

function decodeBase64(value: string) {
  const clean = value.replace(/^data:[^,]+,/, '');
  const binary = atob(clean);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

export default function SuperDocCertificateEditor({ template, adminPassword, onBack, onChanged }: {
  template: Template; adminPassword: string; onBack: () => void; onChanged: () => Promise<void>;
}) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<HTMLDivElement | null>(null);
  const instanceRef = useRef<any>(null);
  const sourceRef = useRef<Blob | null>(null);
  const [loading, setLoading] = useState(true);
  const [ready, setReady] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setLoading(true); setError(''); setMessage('');
      try {
        const result = await api.get('/api/certificados/plantillas/' + encodeURIComponent(template.id) + '/file');
        const encoded = result?.data?.data;
        if (!encoded) throw new Error('template-empty');
        const bytes = decodeBase64(encoded);
        const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        sourceRef.current = blob;
        if (!hostRef.current || !toolbarRef.current) throw new Error('editor-host-missing');

        const instance = new SuperDoc({
          selector: hostRef.current,
          document: blob,
          documentMode: 'editing',
          pagination: true,
          toolbar: toolbarRef.current,
          title: template.name,
          onReady: () => {
            if (cancelled) return;
            setReady(true);
            setLoading(false);
          },
          onContentError: ({ error: detail }: any) => {
            console.error('SuperDoc content error', detail);
            if (!cancelled) { setError('SuperDoc no pudo interpretar esta plantilla DOCX.'); setLoading(false); }
          },
          onException: ({ error: detail }: any) => {
            console.error('SuperDoc exception', detail);
            if (!cancelled) { setError('Error técnico al abrir la plantilla DOCX.'); setLoading(false); }
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
    };
  }, [template.id, template.name]);

  const exportDocx = async () => {
    const instance = instanceRef.current;
    if (!instance || !ready) return;
    setBusy(true); setMessage('');
    try {
      const blob = await instance.export({ exportType: ['docx'], triggerDownload: false });
      if (!(blob instanceof Blob)) throw new Error('export-invalid');
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = template.name.replace(/[^a-zA-Z0-9_-]+/g, '_') + '_editado.docx';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setMessage('Word descargado correctamente.');
    } catch (e) {
      console.error('SuperDoc DOCX export error', e);
      setMessage('No fue posible exportar el Word.');
    } finally { setBusy(false); }
  };

  const saveBackup = async () => {
    const instance = instanceRef.current;
    if (!instance || !adminPassword) return false;
    setBusy(true);
    try {
      const doc = instance.activeEditor?.doc;
      const html = doc ? await doc.getHtml({}) : '';
      await api.post('/api/certificados/respaldos', {
        password: adminPassword,
        template_id: template.id,
        template_name: template.name,
        filename: template.name + '_respaldo',
        fields: {},
        html,
      });
      await onChanged();
      return true;
    } catch (e) {
      console.error('Certificate backup error', e);
      setMessage('No se pudo guardar el respaldo.');
      return false;
    } finally { setBusy(false); }
  };

  const generatePdf = async () => {
    const host = hostRef.current;
    if (!host || !ready) return;
    setBusy(true); setMessage('Generando PDF…');
    try {
      const ok = await saveBackup();
      if (!ok) return;
      window.print();
      setMessage('Se abrió la impresión del certificado. Selecciona “Guardar como PDF”.');
    } finally { setBusy(false); }
  };

  return <div className='superdoc-certificate-shell'>
    <div className='superdoc-certificate-head'>
      <button type='button' className='secondary' onClick={onBack}><ChevronLeft size={16}/> Certificados</button>
      <div><h2>{template.name}</h2><p>Documento DOCX real · edición directa tipo Word</p></div>
      <div className='superdoc-certificate-actions'>
        <button type='button' className='secondary' onClick={() => void saveBackup()} disabled={!ready || busy}><Save size={16}/> Guardar respaldo</button>
        <button type='button' className='secondary' onClick={() => void exportDocx()} disabled={!ready || busy}><Download size={16}/> Descargar Word</button>
        <button type='button' className='primary' onClick={() => void generatePdf()} disabled={!ready || busy}><FileText size={16}/> Generar PDF</button>
      </div>
    </div>
    <div className='superdoc-certificate-body'>
      <div className='superdoc-editor-column'>
        <div ref={toolbarRef} className='superdoc-toolbar' />
        <div className='superdoc-stage'>
          {loading && <div className='superdoc-loading'><Loader2 className='spin' size={26}/> Cargando documento original…</div>}
          {error && <div className='superdoc-error'><b>No se pudo abrir la plantilla.</b><span>{error}</span></div>}
          <div ref={hostRef} className='superdoc-document-host' />
        </div>
      </div>
    </div>
  </div>;
}
