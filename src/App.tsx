import { useEffect, useMemo, useRef, useState, type CSSProperties } from 'react';
import { api } from './lib/api';
import CertificateModule, { CertificateSidebar } from './CertificateModule';
import {
  Search,
  Copy,
  ShieldCheck,
  Bold,
  Italic,
  Underline,
  Strikethrough,
  List,
  ListOrdered,
  AlignLeft,
  AlignCenter,
  AlignRight,
  RemoveFormatting,
  Settings,
  Plus,
  Pencil,
  Trash2,
  X,
  Check,
  Lock,
  LogOut,
  Download,
  Upload,
} from 'lucide-react';

type Category = { id: string; name: string };
type ResponseItem = {
  id: string;
  category_id: string;
  title: string;
  keywords: string;
  text: string;
  instruction_id?: string;
  updated_at: number;
};

const allowedRichTags = new Set(['B', 'STRONG', 'I', 'EM', 'U', 'S', 'BR', 'P', 'DIV', 'UL', 'OL', 'LI']);

function sanitizeRichText(value: string) {
  if (!value) return '';
  const parser = new DOMParser();
  const doc = parser.parseFromString(value, 'text/html');
  const clean = doc.createElement('div');
  const walk = (node: Node, parent: HTMLElement) => {
    node.childNodes.forEach(child => {
      if (child.nodeType === Node.TEXT_NODE) {
        parent.appendChild(doc.createTextNode(child.textContent || ''));
        return;
      }
      if (child.nodeType !== Node.ELEMENT_NODE) return;
      const element = child as HTMLElement;
      const tag = element.tagName.toUpperCase();
      if (!allowedRichTags.has(tag)) {
        walk(element, parent);
        return;
      }
      const safe = doc.createElement(tag.toLowerCase());
      walk(element, safe);
      parent.appendChild(safe);
    });
  };
  walk(doc.body, clean);
  return clean.innerHTML;
}

function richTextToPlain(value: string) {
  if (!value) return '';
  const parser = new DOMParser();
  const doc = parser.parseFromString(sanitizeRichText(value), 'text/html');
  const walk = (node: Node): string => { if (node.nodeType === Node.TEXT_NODE) return (node.textContent || '').replace(/\u00a0/g, ' '); if (node.nodeType !== Node.ELEMENT_NODE) return ''; const element = node as HTMLElement; const tag = element.tagName.toUpperCase(); if (tag === 'BR') return '\n'; const content = Array.from(element.childNodes).map(walk).join(''); if (tag === 'DIV' || tag === 'P') return `${content}\n`; if (tag === 'LI') return `• ${content}\n`; return content; }; return walk(doc.body).replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

function normalizeRichText(value: string) {
  const trimmed = value.trim();
  if (!trimmed) return '';
  const looksFormatted = /<\/?(strong|b|em|i|u|s|br|p|div|ul|ol|li)\b/i.test(trimmed);
  if (looksFormatted) return sanitizeRichText(trimmed);
  return trimmed.split(/\r?\n/).map(line => `<div>${line.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') || '<br>'}</div>`).join('');
}
type Instruction = { id: string; name: string; path: string; url: string; content_type: string; size?: number; created_at: number };
type CertificateTemplate = { id: string; name: string; description: string; filename: string; path: string; content_type: string; size?: number; created_at: number; updated_at?: number };
type CertificateBackup = { id: string; template_id: string; template_name: string; filename?: string; fields: Record<string, string>; html: string; created_at: number }; 
type SettingsData = { institution_name: string; logo_url: string; logo_path?: string; background_image_url: string; background_path?: string; background_type?: 'image' | 'video'; background_video_url?: string; background_video_path?: string; background_transparency?: number; primary_color: string; accent_color: string; background_color: string; topbar_mode?: 'solid' | 'gradient'; topbar_color_1?: string; topbar_color_2?: string; topbar_color_3?: string; sidebar_mode?: 'solid' | 'gradient'; sidebar_color_1?: string; sidebar_color_2?: string; sidebar_color_3?: string; topbar_font_family?: string; sidebar_font_family?: string; hero_mode?: 'solid' | 'gradient'; hero_color_1?: string; hero_color_2?: string; hero_color_3?: string; hero_font_family?: string };  

const defaultCategories = [
  'Ampliación de fechas',
  'Anulación de recibos',
  'Aplicación de descuentos',
  'Aplicación de descuentos educación continuada',
  'Becas',
  'Cargue de pago',
  'Certificación financiera',
  'Deudas pendientes',
  'Devolución',
  'Financiación directa',
  'Fraccionamiento',
  'Generación recibo de pago',
  'ICETEX',
  'Saldos a favor',
  'Pagos de empresas',
  'Traslado de pagos por cambios de programas',
];
  const defaultSettings: SettingsData = { institution_name: 'Fundación Universitaria Compensar', logo_url: '/resources/logo-ucompensar.png', logo_path: '', background_image_url: '/resources/fondo-ucompensar.png', background_path: '', background_type: 'image', background_video_url: '', background_video_path: '', primary_color: '#FF6801', accent_color: '#4E2A54', background_color: '#F6F7F9', topbar_mode: 'solid', topbar_color_1: '#FFFFFF', topbar_color_2: '#FFFFFF', topbar_color_3: '#FFFFFF', sidebar_mode: 'gradient', sidebar_color_1: '#4E2A54', sidebar_color_2: '#6E5E9D', sidebar_color_3: '#1896A8', topbar_font_family: 'ui-rounded', sidebar_font_family: 'ui-rounded', hero_mode: 'gradient', hero_color_1: '#4E2A54', hero_color_2: '#6E5E9D', hero_color_3: '#FF6801', hero_font_family: 'ui-rounded' }; 

function App() {
  const [categories, setCategories] = useState<Category[]>([]);
  const [responses, setResponses] = useState<ResponseItem[]>([]);
  const [instructions, setInstructions] = useState<Instruction[]>([]);
  const [certificateTemplates, setCertificateTemplates] = useState<CertificateTemplate[]>([]);
  const [certificateBackups, setCertificateBackups] = useState<CertificateBackup[]>([]);
  const [activeModule, setActiveModule] = useState<'crm' | 'certificates'>('crm');
  const [settings, setSettings] = useState<SettingsData>(defaultSettings);
  const [selected, setSelected] = useState('all');
  const [query, setQuery] = useState('');
  const [admin, setAdmin] = useState(false);
  const [showLogin, setShowLogin] = useState(false);
  const [password, setPassword] = useState('');
  const [adminPassword, setAdminPassword] = useState('');
  const [modal, setModal] = useState<
    'response' | 'category' | 'settings' | 'password' | null
  >(null);
  const [editing, setEditing] = useState<ResponseItem | null>(null);
  const [editingCategory, setEditingCategory] = useState<Category | null>(null);
  const [toast, setToast] = useState('');
  const [loading, setLoading] = useState(true);
  const modalRef = useRef<'response' | 'category' | 'settings' | 'password' | null>(null);
  modalRef.current = modal;

  const load = async () => {
    setLoading(true);
    try {
      const result = await api.get('/api/data');
      setCategories(result.data.categories || []);
      setResponses(result.data.responses || []);
      setInstructions(result.data.instructions || []);
      setCertificateTemplates(result.data.certificate_templates || []);
      setCertificateBackups(result.data.certificate_backups || []);
      setSettings(result.data.settings || defaultSettings);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void load();
  }, []);
  useEffect(() => { const unsubscribe = api.subscribe(() => { if (modalRef.current === null) void load(); }); return unsubscribe; }, []);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return responses.filter(
      r =>
        (selected === 'all' || r.category_id === selected) &&
        (!q || `${r.title} ${r.keywords} ${r.text}`.toLowerCase().includes(q))
    );
  }, [responses, selected, query]);

  const categoryName = (id: string) =>
    categories.find(c => c.id === id)?.name || 'Sin tipología';
  const notify = (message: string) => {
    setToast(message);
    window.setTimeout(() => setToast(''), 2200);
  };
  const copy = async (text: string) => {
    const html = sanitizeRichText(text);
    const plain = richTextToPlain(text);
    try {
      const blobHtml = new Blob([html], { type: 'text/html' });
      const blobText = new Blob([plain], { type: 'text/plain' });
      const ClipboardItemCtor = window.ClipboardItem;
      if (ClipboardItemCtor) {
        await navigator.clipboard.write([new ClipboardItemCtor({ 'text/html': blobHtml, 'text/plain': blobText })]);
      } else {
        await navigator.clipboard.writeText(plain);
      }
    } catch {
      await navigator.clipboard.writeText(plain);
    }
    notify('Respuesta copiada al portapapeles');
  }; 

  const login = async () => {
    try {
      const r = await api.post('/api/admin/login', { password });
      if (r.data.ok) {
        setAdmin(true);
        setAdminPassword(password);
        setShowLogin(false);
        setPassword('');
        notify('Modo administrador activado');
      } else notify('Contraseña incorrecta');
    } catch {
      notify('No fue posible validar el acceso');
    }
  };
  const logout = () => {
      setAdmin(false);
    setAdminPassword('');
    notify('Modo consulta activado');
  };

  const saveResponse = async (
    data: Omit<ResponseItem, 'id' | 'updated_at'>,
    id?: string
  ) => {
    if (!admin) return;
    if (!richTextToPlain(data.title) || !richTextToPlain(data.text) || !data.category_id)
      return notify('Completa tipología, título y respuesta');
    if (data.instruction_id === '__choose__')
      return notify('Selecciona el instructivo que debe adjuntarse o desactiva la opción');
    try {
      const payload = { ...data, title: sanitizeRichText(data.title), text: normalizeRichText(data.text), instruction_id: data.instruction_id || '', password: adminPassword };
      if (id) await api.put(`/api/responses/${id}`, payload);
      else await api.post('/api/responses', payload);
      setModal(null);
      setEditing(null);
      await load();
      notify('Respuesta guardada');
    } catch {
      notify('No se pudo guardar la respuesta');
    }
  };
  const deleteResponse = async (r: ResponseItem) => {
    if (!admin || !window.confirm(`¿Eliminar la respuesta “${r.title}”?`))
      return;
    try {
      await api.post(`/api/responses/${r.id}/delete`, { password: adminPassword });
      await load();
      notify('Respuesta eliminada');
    } catch {
      notify('No se pudo eliminar');
    }
  };
  const saveCategory = async (name: string, id?: string) => {
    if (!admin || !name.trim()) return;
    try {
      if (id)
        await api.put(`/api/categories/${id}`, { name: name.trim(), password: adminPassword });
      else await api.post('/api/categories', { name: name.trim(), password: adminPassword });
      setModal(null);
      setEditingCategory(null);
      await load();
      notify('Tipología guardada');
    } catch {
      notify('No se pudo guardar la tipología');
    }
  };
  const deleteCategory = async (c: Category) => {
    if (!admin) return;
    const count = responses.filter(r => r.category_id === c.id).length;
    const warning = count
      ? `La tipología tiene ${count} respuesta(s). Al eliminarla también se eliminarán esas respuestas.`
      : 'Esta acción no se puede deshacer.';
    if (!window.confirm(`¿Eliminar “${c.name}”?\n\n${warning}`)) return;
    try {
      await api.delete(`/api/categories/${c.id}`, { password: adminPassword });
      if (selected === c.id) setSelected('all');
      await load();
      notify('Tipología eliminada');
    } catch {
      notify('No se pudo eliminar');
    }
  };
  const saveSettings = async (s: SettingsData) => {
    if (!admin) return;
      try { await api.put('/api/settings', { ...s, password: adminPassword }); await load(); setModal(null); notify('Configuración actualizada'); } catch { notify('No se pudo actualizar la configuración'); }
  };
  const uploadBrandAsset = async (file: File, kind: 'logo' | 'background') => {
    if (!admin || !adminPassword) return null;
    const allowed = kind === 'logo' ? ['image/png', 'image/jpeg'] : ['image/png', 'image/jpeg', 'video/mp4', 'video/webm'];
    if (!allowed.includes(file.type)) { notify(kind === 'logo' ? 'Solo se permiten JPG o PNG' : 'El fondo debe ser JPG, PNG, MP4 o WebM'); return null; }
    const maxSize = kind === 'logo' ? 700 * 1024 : 10 * 1024 * 1024;
    if (file.size > maxSize) { notify(kind === 'logo' ? 'La imagen debe pesar máximo 700 KB' : 'El video debe pesar máximo 4 MB'); return null; }
    const dataUrl = await new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result || '')); reader.onerror = () => reject(new Error('read')); reader.readAsDataURL(file); });
    const base64 = dataUrl.split(',')[1] || '';
    try { const r = await api.post('/api/settings/upload', { password: adminPassword, kind, filename: file.name, contentType: file.type, data: base64 }); notify(kind === 'logo' ? 'Logo cargado correctamente' : file.type.startsWith('video/') ? 'Video de fondo cargado correctamente' : 'Imagen de fondo cargada correctamente'); return { ...(r.data as { path: string; url: string }), contentType: file.type }; } catch { notify(kind === 'logo' ? 'No se pudo cargar el logo.' : 'No se pudo cargar el fondo. Verifica el formato y tamaño.'); return null; }
  }; 
  const changePassword = async (newPassword: string) => {
    if (!admin || newPassword.length < 6)
      return notify('La contraseña debe tener mínimo 6 caracteres');
    try {
      await api.put('/api/admin/password', {
        password: adminPassword,
        new_password: newPassword,
      });
      setAdminPassword(newPassword);
      setPassword('');
      setModal(null);
      notify('Contraseña actualizada');
    } catch {
      notify('No se pudo cambiar la contraseña');
    }
  };
  const exportBackup = async () => {
    const r = await api.get('/api/data');
    const blob = new Blob([JSON.stringify(r.data, null, 2)], {
      type: 'application/json',
    });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'backup_banco_respuestas_crm.json';
    a.click();
    URL.revokeObjectURL(a.href);
  };
  const uploadInstruction = async (file: File) => {
    if (!admin || !adminPassword) return;
    const allowed = ['application/pdf', 'image/png', 'image/jpeg', 'image/webp'];
    if (!allowed.includes(file.type)) return notify('El instructivo debe ser PDF, JPG, PNG o WebP');
    if (file.size > 8 * 1024 * 1024) return notify('El instructivo debe pesar máximo 8 MB');
    const dataUrl = await new Promise<string>((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result || '')); reader.onerror = () => reject(new Error('read')); reader.readAsDataURL(file); });
    try { const r = await api.post('/api/instructivos/upload', { password: adminPassword, filename: file.name, contentType: file.type, data: dataUrl.split(',')[1] || '' }); setInstructions(current => [...current, r.data.instruction]); notify('Instructivo agregado'); } catch { notify('No se pudo cargar el instructivo'); }
  };
  const deleteInstruction = async (item: Instruction) => {
    if (!admin || !window.confirm(`¿Eliminar el instructivo “${item.name}”?`)) return;
    try { await api.post(`/api/instructivos/${item.id}/delete`, { password: adminPassword }); await load(); notify('Instructivo eliminado'); } catch { notify('No se pudo eliminar el instructivo'); }
  };
  const importBackup = async (file: File) => {
    try {
      const data = JSON.parse(await file.text());
      await api.post('/api/import', { ...data, password: adminPassword });
      await load();
      notify('Respaldo importado');
    } catch {
      notify('El respaldo no es válido o no pudo importarse');
    }
  };

  const readableColor = (hex: string) => { const value = (hex || '').replace('#', ''); if (![3, 6].includes(value.length)) return '#FFFFFF'; const normalized = value.length === 3 ? value.split('').map(c => c + c).join('') : value; const r = parseInt(normalized.slice(0, 2), 16); const g = parseInt(normalized.slice(2, 4), 16); const b = parseInt(normalized.slice(4, 6), 16); const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255; return luminance > 0.62 ? '#1F2937' : '#FFFFFF'; };
  const backgroundTransparency = Math.min(100, Math.max(0, Number(settings.background_transparency ?? 88)));
  const themeStyle = { '--brand-primary': settings.primary_color, '--brand-accent': settings.accent_color, '--page-bg': settings.background_color, '--app-bg-image': settings.background_type === 'video' ? 'none' : (settings.background_image_url ? `url("${settings.background_image_url}")` : 'none'), '--bg-transparency': String(backgroundTransparency / 100), '--bg-video-opacity': String(1 - backgroundTransparency / 100), '--topbar-bg': settings.topbar_mode === 'gradient' ? `linear-gradient(105deg, ${settings.topbar_color_1 || settings.primary_color}, ${settings.topbar_color_2 || settings.accent_color}, ${settings.topbar_color_3 || settings.accent_color})` : (settings.topbar_color_1 || settings.primary_color), '--sidebar-bg': settings.sidebar_mode === 'gradient' ? `linear-gradient(165deg, ${settings.sidebar_color_1 || settings.accent_color}, ${settings.sidebar_color_2 || settings.primary_color}, ${settings.sidebar_color_3 || settings.accent_color})` : (settings.sidebar_color_1 || settings.accent_color), '--hero-bg': settings.hero_mode === 'gradient' ? `linear-gradient(118deg, ${settings.hero_color_1 || settings.primary_color}, ${settings.hero_color_2 || settings.accent_color}, ${settings.hero_color_3 || settings.accent_color})` : (settings.hero_color_1 || settings.primary_color), '--topbar-text': readableColor(settings.topbar_color_1 || settings.primary_color), '--sidebar-text': readableColor(settings.sidebar_color_1 || settings.accent_color), '--hero-text': '#FFFFFF', '--topbar-font': settings.topbar_font_family || 'ui-rounded', '--sidebar-font': settings.sidebar_font_family || 'ui-rounded', '--hero-font': settings.hero_font_family || 'ui-rounded' } as CSSProperties;
  return (
    <div className="app-shell" style={themeStyle}>
      {settings.background_type === 'video' && (settings.background_video_url || settings.background_image_url) && <video className="app-bg-video" src={settings.background_video_url || settings.background_image_url} autoPlay muted loop playsInline aria-hidden="true" style={{ opacity: 'var(--bg-video-opacity)' }} onError={e => { e.currentTarget.style.display = 'none'; }} />}
      <header className="topbar" style={{ background: 'var(--topbar-bg)' }}>
        <div className="brand">
          {settings.logo_url ? (
            <img className="brand-logo" src={settings.logo_url} alt={settings.institution_name} />
          ) : (
            <div className="brand-wordmark"><span className="brand-flower">✦</span><div><strong>compensar</strong><small>fundación universitaria</small></div></div>
          )}
          <div className="brand-copy">
            <strong>{settings.institution_name}</strong>
            <span>Área de Apoyo Financiero · Banco de respuestas CRM</span>
          </div>
        </div>
        <div className="top-actions">
          {admin ? (
            <>
              <span className="admin-pill">
                <ShieldCheck size={15} /> Administrador
              </span>
              <button className="ghost" onClick={logout}>
                <LogOut size={16} /> Salir
              </button>
            </>
          ) : (
            <button className="ghost" onClick={() => setShowLogin(true)}>
              <Lock size={16} /> Acceso administrador
            </button>
          )}
        </div>
      </header>
      <nav className="module-tabs" aria-label="Módulos de la aplicación">
        <button className={`module-tab ${activeModule === 'crm' ? 'active' : ''}`} onClick={() => setActiveModule('crm')}>
          <Copy size={16} />
          <span>Banco de respuestas</span>
        </button>
        <button className={`module-tab ${activeModule === 'certificates' ? 'active' : ''}`} onClick={() => setActiveModule('certificates')}>
          <ShieldCheck size={16} />
          <span>Certificados financieros</span>
        </button>
      </nav>
      <main className="layout">
        <aside className="sidebar" style={{ background: 'var(--sidebar-bg)' }}>
          {activeModule === 'certificates' ? <CertificateSidebar templates={certificateTemplates} admin={admin} /> : <>
            <div className="side-title">TIPOLOGÍAS <span>{categories.length}</span></div>
            <button className={`category ${selected === 'all' ? 'active' : ''}`} onClick={() => setSelected('all')}>
              Todas las respuestas <b>{responses.length}</b>
            </button>
            {categories.map(c => (
              <div className="cat-row" key={c.id}>
                <button className={`category ${selected === c.id ? 'active' : ''}`} onClick={() => setSelected(c.id)}>
                  {c.name}
                  <b>{responses.filter(r => r.category_id === c.id).length}</b>
                </button>
                {admin && <div className="mini-actions">
                  <button onClick={() => { setEditingCategory(c); setModal('category'); }}><Pencil size={13} /></button>
                  <button onClick={() => void deleteCategory(c)}><Trash2 size={13} /></button>
                </div>}
              </div>
            ))}
            {admin && <button className="add-side" onClick={() => { setEditingCategory(null); setModal('category'); }}>
              <Plus size={16} /> Nueva tipología
            </button>}
          </>}
        </aside>
        <section className={`content ${activeModule === 'certificates' ? 'certificates-content' : ''}`}>
          {activeModule === 'certificates' ? <CertificateModule templates={certificateTemplates} backups={certificateBackups} admin={admin} adminPassword={adminPassword} onChanged={load} /> : <div className="crm-main-content">
          <div className="hero" style={{ background: 'var(--hero-bg)', color: 'var(--hero-text)', fontFamily: 'var(--hero-font)' }}> 
            <div>
              <div className="eyebrow">CONSULTA RÁPIDA</div>
              <h1>¿Qué respuesta necesitas hoy?</h1>
              <p>
                Busca, consulta y copia respuestas listas para llevar al CRM.
              </p>
            </div>
            <div className="hero-icon">
              <Search size={30} />
            </div>
          </div>
          <div className="toolbar">
            <div className="search">
              <Search size={19} />
              <input
                value={query}
                onChange={e => setQuery(e.target.value)}
                placeholder="Buscar por título, palabras clave o contenido..."
              />
              <kbd>⌘ K</kbd>
            </div>
            {admin && (
              <>
                <button
                  className="primary"
                  onClick={() => {
                    setEditing(null);
                    setModal('response');
                  }}
                >
                  <Plus size={17} /> Nueva respuesta
                </button>
                <button
                  className="icon-btn"
                  title="Configuración"
                  onClick={() => setModal('settings')}
                >
                  <Settings size={18} />
                </button>
              </>
            )}
          </div>
          <div className="result-head">
            <span>
              {loading
                ? 'Cargando…'
                : `${filtered.length} respuesta${filtered.length === 1 ? '' : 's'}`}
            </span>
            <span>
              {selected === 'all'
                ? 'Todas las tipologías'
                : categoryName(selected)}
            </span>
          </div>
          <div className="cards">
            {filtered.map(r => (
              <article className="card" key={r.id}>
                <div className="card-top">
                  <span className="tag">{categoryName(r.category_id)}</span>
                  {admin && (
                    <div className="card-actions">
                      <button
                        onClick={() => {
                          setEditing(r);
                          setModal('response');
                        }}
                      >
                        <Pencil size={15} />
                      </button>
                      <button onClick={() => void deleteResponse(r)}>
                        <Trash2 size={15} />
                      </button>
                    </div>
                  )}
                </div>
                <h3 dangerouslySetInnerHTML={{ __html: sanitizeRichText(r.title) }} />
                {r.keywords && (
                  <div className="keywords">
                    {r.keywords.split(',').map(k => (
                      <span key={k}>{k.trim()}</span>
                    ))}
                  </div>
                )}
                <div className="response-content" dangerouslySetInnerHTML={{ __html: sanitizeRichText(r.text) }} />
                {r.instruction_id && (() => { const instruction = instructions.find(item => item.id === r.instruction_id); return <div className="response-instruction">
                  <div className="response-instruction-warning">⚠️ ESTA RESPUESTA REQUIERE INSTRUCTIVO</div>
                  {instruction ? <a className="response-instruction-link" href={instruction.url} target="_blank" rel="noreferrer" download>
                    <Download size={15} /> Descargar instructivo: <strong>{instruction.name}</strong>
                  </a> : <div className="response-instruction-missing">El instructivo asociado ya no está disponible. Contacta al administrador.</div>}
                </div>; })()}
                <button className="copy" onClick={() => void copy(r.text)}>
                  <Copy size={16} /> Copiar respuesta
                </button>
              </article>
            ))}
            {!loading && filtered.length === 0 && (
              <div className="empty">
                <Search size={30} />
                <h3>No encontramos respuestas</h3>
                <p>Prueba con otra palabra, título o tipología.</p>
              </div>
            )}
          </div>
          </div>
          }
        </section>
      </main>
      {toast && (
        <div className="toast">
          <Check size={17} />
          {toast}
        </div>
      )}
      {showLogin && (
        <div className="overlay">
          <div className="modal small">
            <button className="close" onClick={() => setShowLogin(false)}>
              <X />
            </button>
            <div className="modal-icon">
              <Lock />
            </div>
            <h2>Acceso administrador</h2>
            <p>Ingresa la contraseña para administrar el banco compartido.</p>
            <input
              autoFocus
              type="password"
              value={password}
              onChange={e => setPassword(e.target.value)}
              onKeyDown={e => {
                if (e.key === 'Enter') void login();
              }}
              placeholder="Contraseña"
            />
            <button className="primary wide" onClick={() => void login()}>
              Entrar
            </button>
          </div>
        </div>
      )}
      {modal === 'response' && (
        <ResponseModal
          item={editing}
          categories={categories}
          instructions={instructions}
          onClose={() => {
            setModal(null);
            setEditing(null);
          }}
          onSave={saveResponse}
        />
      )}
      {modal === 'category' && (
        <CategoryModal
          item={editingCategory}
          onClose={() => {
            setModal(null);
            setEditingCategory(null);
          }}
          onSave={saveCategory}
        />
      )}
      {modal === 'settings' && (
        <SettingsModal
          settings={settings}
          onClose={() => setModal(null)}
          onSave={saveSettings}
          onPassword={() => setModal('password')}
          onExport={exportBackup}
          onImport={importBackup}
          categories={categories}
          instructions={instructions}
          onUploadInstruction={uploadInstruction}
          onDeleteInstruction={deleteInstruction}
          onNewCategory={() => {
            setEditingCategory(null);
            setModal('category');
          }}
          onEditCategory={category => {
            setEditingCategory(category);
            setModal('category');
          }}
          onDeleteCategory={deleteCategory}
          onUpload={uploadBrandAsset}
        />
      )}
      {modal === 'password' && (
        <PasswordModal
          onClose={() => setModal('settings')}
          onSave={changePassword}
        />
      )}
    </div>
  );
}

function ResponseModal({
  item,
  categories,
  instructions,
  onClose,
  onSave,
}: {
  item: ResponseItem | null;
  categories: Category[];
  instructions: Instruction[];
  onClose: () => void;
  onSave: (
    d: Omit<ResponseItem, 'id' | 'updated_at'>,
    id?: string
  ) => Promise<void>;
}) {
  const [d, setD] = useState({
    category_id: item?.category_id || categories[0]?.id || '',
    title: item?.title || '',
    keywords: item?.keywords || '',
    text: item?.text || '',
    instruction_id: item?.instruction_id || '',
  });
  const titleEditorRef = useRef<HTMLDivElement | null>(null);
  const editorRef = useRef<HTMLDivElement | null>(null);
  useEffect(() => {
    if (titleEditorRef.current) titleEditorRef.current.innerHTML = sanitizeRichText(item?.title || '');
    if (editorRef.current) editorRef.current.innerHTML = sanitizeRichText(item?.text || '');
  }, [item?.id]);
  const runFormat = (ref: React.RefObject<HTMLDivElement | null>, field: 'title' | 'text', command: string, value?: string) => {
    const editor = ref.current;
    if (!editor) return;
    editor.focus();
    try { document.execCommand(command, false, value); } catch { return; }
  };
  const applyTitleFormat = (command: string, value?: string) => runFormat(titleEditorRef, 'title', command, value);
  const applyTextFormat = (command: string, value?: string) => runFormat(editorRef, 'text', command, value);
  const handlePaste = (event: React.ClipboardEvent<HTMLDivElement>, ref: React.RefObject<HTMLDivElement | null>) => {
    event.preventDefault();
    const editor = ref.current;
    if (!editor) return;
    const html = event.clipboardData.getData('text/html');
    const plain = event.clipboardData.getData('text/plain').replace(/\\r\\n?/g, '\\n');
    const safeHtml = html ? sanitizeRichText(html) : normalizeRichText(plain);
    try {
      editor.focus();
      if (safeHtml) document.execCommand('insertHTML', false, safeHtml);
      else if (plain) document.execCommand('insertText', false, plain);
    } catch {
      try { document.execCommand('insertText', false, plain); } catch { return; }
    }
  };
  return (
    <div className="overlay">
      <div className="modal">
        <button className="close" onClick={onClose}>
          <X />
        </button>
        <h2>{item ? 'Editar respuesta' : 'Nueva respuesta'}</h2>
        <p>La información queda disponible para todo el equipo.</p>
        <label>
          Tipología
          <select
            value={d.category_id}
            onChange={e => setD({ ...d, category_id: e.target.value })}
          >
            {categories.map(c => (
              <option key={c.id} value={c.id}>
                {c.name}
              </option>
            ))}
          </select>
        </label>
        <label>
          Título
          <div className="rich-editor title-editor">
            <div className="rich-toolbar" aria-label="Formato del título">
              <button type="button" title="Negrilla" onMouseDown={e => e.preventDefault()} onClick={() => applyTitleFormat('bold')}><Bold size={15} /></button>
              <button type="button" title="Cursiva" onMouseDown={e => e.preventDefault()} onClick={() => applyTitleFormat('italic')}><Italic size={15} /></button>
              <button type="button" title="Subrayado" onMouseDown={e => e.preventDefault()} onClick={() => applyTitleFormat('underline')}><Underline size={15} /></button>
              <button type="button" title="Tachado" onMouseDown={e => e.preventDefault()} onClick={() => applyTitleFormat('strikeThrough')}><Strikethrough size={15} /></button>
              <span className="toolbar-divider" />
              <button type="button" title="Alinear a la izquierda" onMouseDown={e => e.preventDefault()} onClick={() => applyTitleFormat('justifyLeft')}><AlignLeft size={15} /></button>
              <button type="button" title="Centrar" onMouseDown={e => e.preventDefault()} onClick={() => applyTitleFormat('justifyCenter')}><AlignCenter size={15} /></button>
              <button type="button" title="Alinear a la derecha" onMouseDown={e => e.preventDefault()} onClick={() => applyTitleFormat('justifyRight')}><AlignRight size={15} /></button>
              <button type="button" title="Quitar formato" onMouseDown={e => e.preventDefault()} onClick={() => applyTitleFormat('removeFormat')}><RemoveFormatting size={15} /></button>
            </div>
            <div ref={titleEditorRef} className="rich-input title-rich-input" contentEditable suppressContentEditableWarning onPaste={e => handlePaste(e, titleEditorRef)} data-placeholder="Escribe el título y aplica negrilla, cursiva, subrayado..." />
          </div>
          <small className="editor-helper">Selecciona el texto y usa la barra para darle formato.</small>
        </label>
        <label>
          Palabras clave
          <input
            value={d.keywords}
            onChange={e => setD({ ...d, keywords: e.target.value })}
            placeholder="Ej.: recibo, fecha, pago"
          />
        </label>
        <label className="instruction-toggle-field">
          <span>¿Esta respuesta requiere instructivo?</span>
          <div className="instruction-toggle-row">
            <button type="button" className={`switch ${d.instruction_id ? 'on' : ''}`} role="switch" aria-checked={Boolean(d.instruction_id)} onClick={() => setD({ ...d, instruction_id: d.instruction_id ? '' : '__choose__' })}>
              <span />
            </button>
            <b>{d.instruction_id ? 'Sí, requiere instructivo' : 'No requiere instructivo'}</b>
          </div>
          {d.instruction_id && (
            <select value={d.instruction_id === '__choose__' ? '' : d.instruction_id} onChange={e => setD({ ...d, instruction_id: e.target.value || '__choose__' })}>
              <option value="">Selecciona el instructivo que debe adjuntarse</option>
              {instructions.map(instruction => <option key={instruction.id} value={instruction.id}>{instruction.name}</option>)}
            </select>
          )}
          <small className="editor-helper">El analista verá el aviso y el acceso directo al instructivo desde esta respuesta.</small>
        </label>
        <label>
          Respuesta
          <div className="rich-editor">
            <div className="rich-toolbar" aria-label="Formato de respuesta">
              <button type="button" title="Negrilla" onMouseDown={e => e.preventDefault()} onClick={() => applyTextFormat('bold')}><Bold size={15} /></button>
              <button type="button" title="Cursiva" onMouseDown={e => e.preventDefault()} onClick={() => applyTextFormat('italic')}><Italic size={15} /></button>
              <button type="button" title="Subrayado" onMouseDown={e => e.preventDefault()} onClick={() => applyTextFormat('underline')}><Underline size={15} /></button>
              <button type="button" title="Tachado" onMouseDown={e => e.preventDefault()} onClick={() => applyTextFormat('strikeThrough')}><Strikethrough size={15} /></button>
              <span className="toolbar-divider" />
              <button type="button" title="Lista con viñetas" onMouseDown={e => e.preventDefault()} onClick={() => applyTextFormat('insertUnorderedList')}><List size={15} /></button>
              <button type="button" title="Lista numerada" onMouseDown={e => e.preventDefault()} onClick={() => applyTextFormat('insertOrderedList')}><ListOrdered size={15} /></button>
              <span className="toolbar-divider" />
              <button type="button" title="Alinear a la izquierda" onMouseDown={e => e.preventDefault()} onClick={() => applyTextFormat('justifyLeft')}><AlignLeft size={15} /></button>
              <button type="button" title="Centrar" onMouseDown={e => e.preventDefault()} onClick={() => applyTextFormat('justifyCenter')}><AlignCenter size={15} /></button>
              <button type="button" title="Alinear a la derecha" onMouseDown={e => e.preventDefault()} onClick={() => applyTextFormat('justifyRight')}><AlignRight size={15} /></button>
              <button type="button" title="Quitar formato" onMouseDown={e => e.preventDefault()} onClick={() => applyTextFormat('removeFormat')}><RemoveFormatting size={15} /></button>
            </div>
            <div ref={editorRef} className="rich-input" contentEditable suppressContentEditableWarning onPaste={e => handlePaste(e, editorRef)} data-placeholder="Escribe la respuesta y aplica negrilla, cursiva, listas y alineación..." />
          </div>
          <small className="editor-helper">Selecciona el texto y usa la barra para darle formato.</small>
        </label>
        <div className="modal-footer">
          <button className="ghost" onClick={onClose}>
            Cancelar
          </button>
          <button className="primary" onClick={() => { const currentData = { ...d, title: titleEditorRef.current?.innerHTML || '', text: editorRef.current?.innerHTML || '' }; void onSave(currentData, item?.id); }}>
            <Check size={16} /> Guardar
          </button>
        </div>
      </div>
    </div>
  );
}
function CategoryModal({
  item,
  onClose,
  onSave,
}: {
  item: Category | null;
  onClose: () => void;
  onSave: (name: string, id?: string) => Promise<void>;
}) {
  const [name, setName] = useState(item?.name || '');
  return (
    <div className="overlay">
      <div className="modal small">
        <button className="close" onClick={onClose}>
          <X />
        </button>
        <h2>{item ? 'Editar tipología' : 'Nueva tipología'}</h2>
        <label>
          Nombre
          <input
            autoFocus
            value={name}
            onChange={e => setName(e.target.value)}
            onKeyDown={e => {
              if (e.key === 'Enter') void onSave(name, item?.id);
            }}
          />
        </label>
        <div className="modal-footer">
          <button className="ghost" onClick={onClose}>
            Cancelar
          </button>
          <button
            className="primary"
            onClick={() => void onSave(name, item?.id)}
          >
            Guardar
          </button>
        </div>
      </div>
    </div>
  );
}
function SettingsModal({
  settings, onClose, onSave, onPassword, onExport, onImport, categories, instructions, onUploadInstruction, onDeleteInstruction, onNewCategory, onEditCategory, onDeleteCategory, onUpload,
}: {
  settings: SettingsData; onClose: () => void; onSave: (s: SettingsData) => Promise<void>; onPassword: () => void; onExport: () => Promise<void>; onImport: (f: File) => Promise<void>; categories: Category[]; instructions: Instruction[]; onUploadInstruction: (file: File) => Promise<void>; onDeleteInstruction: (item: Instruction) => Promise<void>; onNewCategory: () => void; onEditCategory: (category: Category) => void; onDeleteCategory: (category: Category) => Promise<void>; onUpload: (file: File, kind: 'logo' | 'background') => Promise<{ path: string; url: string; contentType: string } | null>;
}) {
  const [s, setS] = useState(settings);
  const [errorText, setErrorText] = useState('');
  const chooseAsset = async (file: File, kind: 'logo' | 'background') => {
    setErrorText('');
    const result = await onUpload(file, kind);
    if (result) setS(current => kind === 'logo' ? { ...current, logo_url: result.url, logo_path: result.path } : { ...current, background_image_url: result.url, background_path: result.path, background_type: result.contentType.startsWith('video/') ? 'video' : 'image', background_video_url: result.contentType.startsWith('video/') ? result.url : '', background_video_path: result.contentType.startsWith('video/') ? result.path : '' });
  };
  return (
    <div className="overlay"><div className="modal settings-modal">
      <button className="close" onClick={onClose}><X /></button>
      <div className="settings-heading"><div className="settings-icon"><Settings size={20} /></div><div><h2>Configuración del banco</h2><p>Personaliza la imagen del banco para todo el equipo.</p></div></div>
      <section className="settings-section">
        <div className="section-title"><span>Identidad corporativa</span><small>Se sincroniza para todos</small></div>
        <label>Nombre de la institución<input value={s.institution_name} onChange={e => setS({ ...s, institution_name: e.target.value })} /></label>
        <div className="branding-grid">
          <div className="branding-card"><div className="branding-preview light">{s.logo_url ? <img className="logo-preview" src={s.logo_url} alt="Logo" /> : <b>Sin logo</b>}</div><div className="branding-copy"><b>Logo de la empresa</b><span>JPG o PNG · máximo 700 KB. Se sube directamente desde tu computador y queda guardado para todo el equipo.</span><label className="upload-logo-btn"><Upload size={16} /> Seleccionar logo<input type="file" accept="image/jpeg,image/png" onChange={e => { const file = e.target.files?.[0]; if (file) void chooseAsset(file, 'logo'); e.currentTarget.value = ''; }} /></label>{s.logo_path && <button className="link-danger" onClick={() => setS({ ...s, logo_url: '', logo_path: '' })}>Quitar logo</button>}</div></div>
          <div className="branding-card"><div className="branding-preview background">{s.background_type === 'video' && (s.background_video_url || s.background_image_url) ? <video src={s.background_video_url || s.background_image_url} autoPlay muted loop playsInline controls /> : s.background_image_url ? <img src={s.background_image_url} alt="Vista previa del fondo" /> : <b>Sin fondo</b>}</div><div className="branding-copy"><b>Fondo del banco</b><span>Imagen JPG/PNG o video corto MP4/WebM. El video se reproduce automáticamente, en silencio y en bucle. Imagen: máximo 700 KB · Video: máximo 10 MB.</span><label className="upload-logo-btn"><Upload size={16} /> Seleccionar fondo<input type="file" accept="image/jpeg,image/png,video/mp4,video/webm" onChange={e => { const file = e.target.files?.[0]; if (file) void chooseAsset(file, 'background'); e.currentTarget.value = ''; }} /></label>{s.background_image_url && <button className="link-danger" onClick={() => setS({ ...s, background_image_url: '', background_path: '', background_type: 'image', background_video_url: '', background_video_path: '' })}>Quitar fondo</button>}</div></div>
        </div>
      </section>
      <section className="settings-section">
        <div className="section-title"><span>Transparencia del fondo</span><small>{Math.round(Number(s.background_transparency ?? 88))}% transparente</small></div><div className="transparency-control"><input type="range" min="0" max="100" step="1" value={s.background_transparency ?? 88} onChange={e => setS({ ...s, background_transparency: Number(e.target.value) })} /><output>{Math.round(Number(s.background_transparency ?? 88))}%</output></div><small className="helper-text">0% muestra el fondo completo · 100% lo oculta. Funciona igual para imagen y video.</small></section><section className="settings-section"><div className="section-title"><span>Colores generales</span><small>Personaliza el estilo visual</small></div>
        <div className="color-grid"><label>Color principal<div className="color-control"><input type="color" value={s.primary_color} onChange={e => setS({ ...s, primary_color: e.target.value })} /><input value={s.primary_color} onChange={e => setS({ ...s, primary_color: e.target.value })} /></div></label><label>Color secundario<div className="color-control"><input type="color" value={s.accent_color} onChange={e => setS({ ...s, accent_color: e.target.value })} /><input value={s.accent_color} onChange={e => setS({ ...s, accent_color: e.target.value })} /></div></label><label>Fondo general<div className="color-control"><input type="color" value={s.background_color} onChange={e => setS({ ...s, background_color: e.target.value })} /><input value={s.background_color} onChange={e => setS({ ...s, background_color: e.target.value })} /></div></label></div>
        <div className="bar-theme-grid">
          <div className="bar-theme-card">
            <div className="bar-theme-heading"><b>Barra superior</b><span>Color fijo o degradado</span></div>
            <div className="mode-switch"><button className={s.topbar_mode === 'solid' ? 'selected' : ''} onClick={() => setS({ ...s, topbar_mode: 'solid' })}>Color fijo</button><button className={s.topbar_mode === 'gradient' ? 'selected' : ''} onClick={() => setS({ ...s, topbar_mode: 'gradient' })}>Degradado</button></div>
            <div className="bar-preview" style={{ background: s.topbar_mode === 'gradient' ? `linear-gradient(105deg, ${s.topbar_color_1 || s.primary_color}, ${s.topbar_color_2 || s.accent_color}, ${s.topbar_color_3 || s.accent_color})` : (s.topbar_color_1 || s.primary_color) }}>Barra superior</div>
            <div className="bar-color-row"><label>Color 1<div className="color-control"><input type="color" value={s.topbar_color_1 || s.primary_color} onChange={e => setS({ ...s, topbar_color_1: e.target.value })} /><input value={s.topbar_color_1 || s.primary_color} onChange={e => setS({ ...s, topbar_color_1: e.target.value })} /></div></label>{s.topbar_mode === 'gradient' && <><label>Color 2<div className="color-control"><input type="color" value={s.topbar_color_2 || s.accent_color} onChange={e => setS({ ...s, topbar_color_2: e.target.value })} /><input value={s.topbar_color_2 || s.accent_color} onChange={e => setS({ ...s, topbar_color_2: e.target.value })} /></div></label><label>Color 3<div className="color-control"><input type="color" value={s.topbar_color_3 || s.accent_color} onChange={e => setS({ ...s, topbar_color_3: e.target.value })} /><input value={s.topbar_color_3 || s.accent_color} onChange={e => setS({ ...s, topbar_color_3: e.target.value })} /></div></label></>}</div>
          </div>
          <div className="bar-theme-card">
            <div className="bar-theme-heading"><b>Barra de tipologías</b><span>Panel lateral de categorías</span></div>
            <div className="mode-switch"><button className={s.sidebar_mode === 'solid' ? 'selected' : ''} onClick={() => setS({ ...s, sidebar_mode: 'solid' })}>Color fijo</button><button className={s.sidebar_mode === 'gradient' ? 'selected' : ''} onClick={() => setS({ ...s, sidebar_mode: 'gradient' })}>Degradado</button></div>
            <div className="bar-preview" style={{ background: s.sidebar_mode === 'gradient' ? `linear-gradient(165deg, ${s.sidebar_color_1 || s.accent_color}, ${s.sidebar_color_2 || s.primary_color}, ${s.sidebar_color_3 || s.accent_color})` : (s.sidebar_color_1 || s.accent_color) }}>TIPOLOGÍAS</div>
            <div className="bar-color-row"><label>Color 1<div className="color-control"><input type="color" value={s.sidebar_color_1 || s.accent_color} onChange={e => setS({ ...s, sidebar_color_1: e.target.value })} /><input value={s.sidebar_color_1 || s.accent_color} onChange={e => setS({ ...s, sidebar_color_1: e.target.value })} /></div></label>{s.sidebar_mode === 'gradient' && <><label>Color 2<div className="color-control"><input type="color" value={s.sidebar_color_2 || s.primary_color} onChange={e => setS({ ...s, sidebar_color_2: e.target.value })} /><input value={s.sidebar_color_2 || s.primary_color} onChange={e => setS({ ...s, sidebar_color_2: e.target.value })} /></div></label><label>Color 3<div className="color-control"><input type="color" value={s.sidebar_color_3 || s.accent_color} onChange={e => setS({ ...s, sidebar_color_3: e.target.value })} /><input value={s.sidebar_color_3 || s.accent_color} onChange={e => setS({ ...s, sidebar_color_3: e.target.value })} /></div></label></>}</div>
          </div>
        </div>
        <div className="bar-advanced-grid">
          <div className="bar-theme-card">
            <div className="bar-theme-heading"><b>Opciones de la barra superior</b><span>Tercer color del degradado y tipografía</span></div>
            <div className="bar-color-row">
              {s.topbar_mode === 'gradient' && <label>Color 3<div className="color-control"><input type="color" value={s.topbar_color_3 || s.accent_color} onChange={e => setS({ ...s, topbar_color_3: e.target.value })} /><input value={s.topbar_color_3 || s.accent_color} onChange={e => setS({ ...s, topbar_color_3: e.target.value })} /></div></label>}
              <label>Tipografía<select value={s.topbar_font_family || 'ui-rounded'} onChange={e => setS({ ...s, topbar_font_family: e.target.value })}><option value="ui-rounded">Redondeada del sistema</option><option value="'Avenir Next'">Avenir Next</option><option value="'Trebuchet MS'">Trebuchet MS</option><option value="'Segoe UI'">Segoe UI</option><option value="Arial">Arial</option><option value="Georgia">Georgia</option></select></label>
            </div>
          </div>
          <div className="bar-theme-card">
            <div className="bar-theme-heading"><b>Opciones de la barra de tipologías</b><span>Tercer color del degradado y tipografía</span></div>
            <div className="bar-color-row">
              {s.sidebar_mode === 'gradient' && <label>Color 3<div className="color-control"><input type="color" value={s.sidebar_color_3 || s.accent_color} onChange={e => setS({ ...s, sidebar_color_3: e.target.value })} /><input value={s.sidebar_color_3 || s.accent_color} onChange={e => setS({ ...s, sidebar_color_3: e.target.value })} /></div></label>}
              <label>Tipografía<select value={s.sidebar_font_family || 'ui-rounded'} onChange={e => setS({ ...s, sidebar_font_family: e.target.value })}><option value="ui-rounded">Redondeada del sistema</option><option value="'Avenir Next'">Avenir Next</option><option value="'Trebuchet MS'">Trebuchet MS</option><option value="'Segoe UI'">Segoe UI</option><option value="Arial">Arial</option><option value="Georgia">Georgia</option></select></label>
            </div>
          </div>
        </div>
      </section>
      <section className="settings-section">
        <div className="section-title"><span>Consulta rápida</span><small>Bloque principal de bienvenida</small></div>
        <div className="hero-theme-card">
          <div className="hero-theme-heading"><b>Colores de “¿Qué respuesta necesitas hoy?”</b><span>Color fijo o degradado independiente de las demás barras</span></div>
          <div className="mode-switch"><button className={s.hero_mode === 'solid' ? 'selected' : ''} onClick={() => setS({ ...s, hero_mode: 'solid' })}>Color fijo</button><button className={s.hero_mode === 'gradient' ? 'selected' : ''} onClick={() => setS({ ...s, hero_mode: 'gradient' })}>Degradado</button></div>
          <div className="hero-preview" style={{ background: s.hero_mode === 'gradient' ? `linear-gradient(118deg, ${s.hero_color_1 || s.primary_color}, ${s.hero_color_2 || s.accent_color}, ${s.hero_color_3 || s.primary_color})` : (s.hero_color_1 || s.primary_color) }}>¿Qué respuesta necesitas hoy?</div>
          <div className="bar-color-row"><label>Color 1<div className="color-control"><input type="color" value={s.hero_color_1 || s.primary_color} onChange={e => setS({ ...s, hero_color_1: e.target.value })} /><input value={s.hero_color_1 || s.primary_color} onChange={e => setS({ ...s, hero_color_1: e.target.value })} /></div></label>{s.hero_mode === 'gradient' && <><label>Color 2<div className="color-control"><input type="color" value={s.hero_color_2 || s.accent_color} onChange={e => setS({ ...s, hero_color_2: e.target.value })} /><input value={s.hero_color_2 || s.accent_color} onChange={e => setS({ ...s, hero_color_2: e.target.value })} /></div></label><label>Color 3<div className="color-control"><input type="color" value={s.hero_color_3 || s.primary_color} onChange={e => setS({ ...s, hero_color_3: e.target.value })} /><input value={s.hero_color_3 || s.primary_color} onChange={e => setS({ ...s, hero_color_3: e.target.value })} /></div></label></>}</div>
          <label>Tipografía<select value={s.hero_font_family || 'ui-rounded'} onChange={e => setS({ ...s, hero_font_family: e.target.value })}><option value="ui-rounded">Redondeada del sistema</option><option value="'Avenir Next'">Avenir Next</option><option value="Montserrat, 'Segoe UI', sans-serif">Montserrat</option><option value="'Trebuchet MS'">Trebuchet MS</option><option value="'Segoe UI'">Segoe UI</option><option value="Arial">Arial</option><option value="Georgia">Georgia</option><option value="Verdana">Verdana</option></select></label>
        </div>
      </section>
      <section className="settings-section"><div className="section-title"><span>Tipologías</span><small>{categories.length} activas</small></div><div className="category-manager">{categories.map(category => <div className="manager-row" key={category.id}><span>{category.name}</span><div><button onClick={() => onEditCategory(category)} title="Editar"><Pencil size={15} /></button><button onClick={() => void onDeleteCategory(category)} title="Eliminar"><Trash2 size={15} /></button></div></div>)}</div><button className="secondary add-category-btn" onClick={onNewCategory}><Plus size={16} /> Agregar nueva tipología</button><small className="helper-text">Puedes agregar, renombrar o eliminar tipologías cuando cambien los procesos.</small></section>
      <section className="settings-section">
        <div className="section-title"><span>Instructivos a la mano</span><small>Administración de archivos complementarios</small></div>
        <div className="instructions-panel settings-instructions-panel">
          <div className="instructions-head"><div><b>Biblioteca de instructivos</b><span>Gestiona los PDF e imágenes que pueden asociarse directamente a las respuestas.</span></div><label className="instruction-upload"><Upload size={15} /> Agregar instructivo<input type="file" accept="application/pdf,image/png,image/jpeg,image/webp" onChange={e => { const file = e.target.files?.[0]; if (file) void onUploadInstruction(file); e.currentTarget.value = ''; }} /></label></div>
          {instructions.length ? <div className="instruction-list">{instructions.map(item => <article className="instruction-card" key={item.id}><div className="instruction-icon">{item.content_type === 'application/pdf' ? 'PDF' : 'IMG'}</div><div className="instruction-copy"><b title={item.name}>{item.name}</b><span>{item.content_type === 'application/pdf' ? 'Documento PDF' : 'Imagen'}</span></div><a className="instruction-open" href={item.url} target="_blank" rel="noreferrer">Abrir</a><button className="instruction-delete" onClick={() => void onDeleteInstruction(item)} title="Eliminar"><Trash2 size={14} /></button></article>)}</div> : <div className="instructions-empty">Todavía no hay instructivos cargados.</div>}
        </div>
        <small className="helper-text">Los instructivos asociados a una respuesta seguirán apareciendo directamente dentro de esa respuesta para el analista.</small>
      </section>
      <section className="settings-section"><div className="section-title"><span>Seguridad y respaldos</span><small>Solo administrador</small></div><div className="settings-grid"><button className="secondary" onClick={onPassword}><Lock size={16} /> Cambiar contraseña</button><button className="secondary" onClick={() => void onExport()}><Download size={16} /> Exportar respaldo</button><label className="secondary file-btn"><Upload size={16} /> Importar respaldo<input type="file" accept="application/json" onChange={e => { const f = e.target.files?.[0]; if (f) void onImport(f); }} /></label></div></section>
      {errorText && <div className="form-error">{errorText}</div>}
      <div className="modal-footer"><button className="ghost" onClick={onClose}>Cerrar</button><button className="primary" onClick={() => void onSave(s)}><Check size={16} /> Guardar cambios</button></div>
    </div></div>
  );
}
function PasswordModal({
  onClose,
  onSave,
}: {
  onClose: () => void;
  onSave: (p: string) => Promise<void>;
}) {
  const [p, setP] = useState('');
  const [c, setC] = useState('');
  return (
    <div className="overlay">
      <div className="modal small">
        <button className="close" onClick={onClose}>
          <X />
        </button>
        <h2>Cambiar contraseña</h2>
        <label>
          Nueva contraseña
          <input
            type="password"
            value={p}
            onChange={e => setP(e.target.value)}
          />
        </label>
        <label>
          Confirmar contraseña
          <input
            type="password"
            value={c}
            onChange={e => setC(e.target.value)}
          />
        </label>
        <div className="modal-footer">
          <button className="ghost" onClick={onClose}>
            Cancelar
          </button>
          <button
            className="primary"
            disabled={!p || p !== c}
            onClick={() => void onSave(p)}
          >
            Actualizar
          </button>
        </div>
      </div>
    </div>
  );
}
export default App;
