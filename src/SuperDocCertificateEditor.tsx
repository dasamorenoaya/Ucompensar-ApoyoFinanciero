import { useEffect, useMemo, useRef, useState } from 'react';
import { SuperDoc } from 'superdoc';
import 'superdoc/style.css';
import { api } from './lib/api';
import { ChevronLeft, Download, FileText, Loader2, Save, Eye, Trash2 } from 'lucide-react';

type Template = { id: string; name: string; filename: string; content_type: string; };
type FieldKey =
  | 'nombre' | 'tipoDocumento' | 'numeroDocumento' | 'programa' | 'periodoAcademico'
  | 'referenciaPago' | 'anio' | 'periodo' | 'totalLiquidado' | 'descuentos'
  | 'saldoFavor' | 'totalPagar' | 'estado' | 'fechaExpedicion' | 'analista' | 'proyecto';

type Field = { key: FieldKey; label: string; value: string; group: 'student' | 'liquidation' | 'signatures'; sample?: string };

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
  { key: 'fechaExpedicion', label: 'Fecha de expedición', value: '11/08/2026', group: 'signatures' },
  { key: 'analista', label: 'Nombre del analista', value: 'Daniel Samuel Moreno', group: 'signatures' },
  { key: 'proyecto', label: 'Proyecto / Revisó', value: 'Daniel Samuel Moreno / Analista de Apoyo Financiero.', group: 'signatures' },
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
  { ...NORMAL_FIELDS[13], value: '04/08/2026' },
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
  { ...NORMAL_FIELDS[13], value: '18/08/2026' },
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

export default function SuperDocCertificateEditor({ template, adminPassword, onBack, onChanged }: {
  template: Template; adminPassword: string; onBack: () => void; onChanged: () => Promise<void>;
}) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<HTMLDivElement | null>(null);
  const instanceRef = useRef<any>(null);
  const [loading, setLoading] = useState(true);
  const [ready, setReady] = useState(false);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [fields, setFields] = useState<Field[]>(() => fieldsForTemplate(template.name));
  const [zoom, setZoom] = useState(100);

  const groups = useMemo(() => ['student', 'liquidation', 'signatures'] as const, []);

  useEffect(() => {
    setFields(fieldsForTemplate(template.name));
  }, [template.name]);

  const replaceInDocument = async (oldValue: string, newValue: string) => {
    const doc = instanceRef.current?.activeEditor?.doc;
    if (!doc || !oldValue || oldValue === newValue) return;
    try {
      const match = await doc.query.match({
        select: { type: 'text', pattern: oldValue },
        require: 'first',
      });
      const item = match.items?.[0];
      if (!item || item.matchKind !== 'text') return;
      await doc.replace({ target: item.target, text: newValue }, { expectedRevision: match.evaluatedRevision });
    } catch (e) {
      console.warn('No fue posible actualizar el campo en DOCX', oldValue, e);
    }
  };

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      setLoading(true); setReady(false); setError(''); setMessage('');
      try {
        const result = await api.get('/api/certificados/plantillas/' + encodeURIComponent(template.id) + '/file');
        const encoded = result?.data?.data;
        if (!encoded) throw new Error('La plantilla no contiene datos.');
        const bytes = decodeBase64(encoded);
        const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        if (!hostRef.current || !toolbarRef.current) throw new Error('No se encontró el contenedor del editor.');

        const instance = new SuperDoc({
          selector: hostRef.current,
          document: blob,
          documentMode: 'editing',
          pagination: true,
          toolbar: toolbarRef.current,
          ui: { search: true },
          title: template.name,
          onReady: ({ superdoc }: any) => {
            if (cancelled) return;
            instanceRef.current = superdoc;
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

  const updateField = async (key: FieldKey, value: string) => {
    const current = fields.find(field => field.key === key);
    setFields(items => items.map(field => field.key === key ? { ...field, value } : field));
    if (current) await replaceInDocument(current.value, value);
  };

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
        fields: Object.fromEntries(fields.map(field => [field.key, field.value])),
        html,
      });
      await onChanged();
      setMessage('Respaldo guardado correctamente.');
      return true;
    } catch (e) {
      console.error('Certificate backup error', e);
      setMessage('No se pudo guardar el respaldo.');
      return false;
    } finally { setBusy(false); }
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
        <div className='superdoc-title-copy'>
          <h2>{template.name}</h2>
          <p>Edita los campos del certificado y genera el PDF.</p>
        </div>
        <button type='button' className='certificate-settings-button' title='Configuración'><span>⚙</span></button>
      </div>

      <div className='superdoc-certificate-layout'>
        <aside className='superdoc-fields-panel'>
          <div className='superdoc-fixed-notice'><span>ⓘ</span><div>El encabezado y el pie de página se mantienen automáticamente en todas las plantillas.</div></div>
          {groups.map(group => {
            const groupFields = fields.filter(field => field.group === group);
            if (!groupFields.length) return null;
            return (
              <section className='superdoc-field-group' key={group}>
                <h3>{groupLabel(group)}</h3>
                <div className='superdoc-field-grid'>
                  {groupFields.map(field => (
                    <label key={field.key} className={field.key === 'numeroDocumento' || field.key === 'tipoDocumento' || field.key === 'referenciaPago' || field.key === 'anio' || field.key === 'periodo' || field.key === 'descuentos' || field.key === 'saldoFavor' || field.key === 'totalPagar' || field.key === 'estado' || field.key === 'fechaExpedicion' || field.key === 'analista' ? 'compact-field' : ''}>
                      <span>{field.label} <b>*</b></span>
                      {field.key === 'tipoDocumento' ? (
                        <select value={field.value} onChange={e => void updateField(field.key, e.target.value)}>
                          <option>Cédula de ciudadanía</option><option>Tarjeta de identidad</option><option>Cédula de extranjería</option><option>Pasaporte</option>
                        </select>
                      ) : field.key === 'estado' ? (
                        <select value={field.value} onChange={e => void updateField(field.key, e.target.value)}>
                          <option>PAGADO</option><option>PENDIENTE</option><option>ANULADO</option><option>VENCIDO</option>
                        </select>
                      ) : (
                        <input value={field.value} onChange={e => void updateField(field.key, e.target.value)} />
                      )}
                    </label>
                  ))}
                </div>
              </section>
            );
          })}
          <div className='superdoc-left-actions'>
            <button type='button' className='secondary' onClick={() => setMessage('Vista previa: el documento que ves a la derecha es el DOCX real.')}><Eye size={15}/> Vista previa</button>
            <button type='button' className='primary' disabled={!ready || busy} onClick={() => void generatePdf()}><FileText size={15}/> Generar PDF</button>
            <button type='button' className='secondary' disabled={busy} onClick={() => setFields(fieldsForTemplate(template.name))}><Trash2 size={15}/> Limpiar campos</button>
          </div>
        </aside>

        <section className='superdoc-document-panel'>
          <div ref={toolbarRef} className='superdoc-toolbar' />
          <div className='superdoc-stage' style={{ ['--certificate-zoom' as string]: zoom / 100 }}>
            {loading && <div className='superdoc-loading'><Loader2 className='spin' size={25}/> Cargando documento original…</div>}
            {error && <div className='superdoc-error'><b>No se pudo abrir la plantilla.</b><span>{error}</span></div>}
            <div ref={hostRef} className='superdoc-document-host' />
          </div>
          <div className='superdoc-document-footer'>
            <button type='button' className='page-nav' aria-label='Página anterior'>‹</button>
            <span>Página 1 de 1</span>
            <button type='button' className='page-nav' aria-label='Página siguiente'>›</button>
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
