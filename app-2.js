/* =====================================================================
   SISTEMA COMERCIAL · app.js
   Ventas, compras, inventario por bodegas con costo promedio, kardex,
   cartera, usuarios con roles e informes en Excel. Backend: Supabase.
   ===================================================================== */
'use strict';

const CFG = window.APP_CONFIG || {};
const LOC = CFG.LOCALE || 'es-CO';
let sb = null;
const S = { perfil: null, empresa: {}, bodegas: [], productos: [], clientes: [], proveedores: [], loginMsg: '' };
const isAdmin = () => S.perfil?.rol === 'admin';
const MEDIOS = ['EFECTIVO', 'TRANSFERENCIA', 'TARJETA DÉBITO', 'TARJETA CRÉDITO', 'NEQUI', 'DAVIPLATA', 'CHEQUE', 'OTRO'];

/* ---------------------------------------------------------------------
   Utilidades
   --------------------------------------------------------------------- */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = v => String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const MONEY = new Intl.NumberFormat(LOC, { style: 'currency', currency: CFG.MONEDA || 'COP', maximumFractionDigits: 0 });
const fm = n => MONEY.format(Number(n) || 0);
const fq = n => (Number(n) || 0).toLocaleString(LOC, { maximumFractionDigits: 3 });
const num = v => { const n = parseFloat(v); return isNaN(n) ? 0 : n; };
const r2 = n => Math.round((Number(n) || 0) * 100) / 100;
const r3 = n => Math.round((Number(n) || 0) * 1000) / 1000;
const r4 = n => Math.round((Number(n) || 0) * 10000) / 10000;
function hoy(off = 0) { const d = new Date(); d.setDate(d.getDate() + off); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); }
function sumarDias(fecha, n) { const d = new Date(fecha + 'T12:00:00'); d.setDate(d.getDate() + (n || 0)); return d.toISOString().slice(0, 10); }
const inicioMes = () => hoy().slice(0, 8) + '01';
const fd = s => s ? String(s).slice(0, 10) : '';
const fdt = s => s ? new Date(s).toLocaleString(LOC, { dateStyle: 'short', timeStyle: 'short' }) : '';
const dIni = d => new Date(d + 'T00:00:00').toISOString();
const dFin = d => new Date(d + 'T23:59:59.999').toISOString();
const diasEntre = (a, b) => Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 86400000);
const go = r => { location.hash = r; };
const prodLabel = p => `${p.codigo} | ${p.nombre}`;
const terLabel = t => `${t.documento} - ${t.nombre}`;
const findBy = (items, fn, val) => items.find(i => fn(i) === val) || null;
const opts = (items, sel, fnV = i => i.id, fnL = i => i.nombre) => items.map(i => `<option value="${esc(fnV(i))}" ${String(fnV(i)) === String(sel) ? 'selected' : ''}>${esc(fnL(i))}</option>`).join('');

function toast(msg, type = 'ok') {
  let box = $('#toasts');
  if (!box) { box = document.createElement('div'); box.id = 'toasts'; document.body.appendChild(box); }
  const t = document.createElement('div'); t.className = 'toast ' + type; t.textContent = msg; box.appendChild(t);
  setTimeout(() => t.classList.add('out'), type === 'err' ? 6000 : 3500);
  setTimeout(() => t.remove(), type === 'err' ? 6500 : 4000);
}
function traducirError(m = '') {
  if (m.includes('Invalid login credentials')) return 'Correo o contraseña incorrectos';
  if (m.includes('duplicate key')) return 'Ya existe un registro con ese código o número de documento';
  if (m.includes('row-level security')) return 'No tienes permisos para esta acción';
  if (m.includes('Password should be')) return 'La contraseña debe tener al menos 6 caracteres';
  if (m.includes('User already registered')) return 'Ese correo ya está registrado';
  if (m.includes('Email not confirmed')) return 'El correo aún no ha sido confirmado';
  if (m.includes('Failed to fetch')) return 'No hay conexión con el servidor. Revisa tu internet o la configuración.';
  return m;
}
function chk(res) { if (res.error) throw new Error(traducirError(res.error.message)); return res.data; }
async function run(fn, btn) {
  if (btn) btn.disabled = true;
  try { return await fn(); }
  catch (e) { console.error(e); toast(e.message || String(e), 'err'); }
  finally { if (btn) btn.disabled = false; }
}
async function fetchAll(build) {
  const out = [];
  for (let from = 0; ; from += 1000) {
    const data = chk(await build().range(from, from + 999));
    out.push(...data);
    if (data.length < 1000) break;
  }
  return out;
}

function modal(title, body, { wide = false, footer = '' } = {}) {
  const m = document.createElement('div');
  m.className = 'modal-bg';
  m.innerHTML = `<div class="modal ${wide ? 'wide' : ''}" role="dialog" aria-modal="true">
    <div class="modal-h"><h3>${esc(title)}</h3><button class="btn icon ghost" data-close aria-label="Cerrar">✕</button></div>
    <div class="modal-b">${body}</div>${footer ? `<div class="modal-f">${footer}</div>` : ''}</div>`;
  m.addEventListener('mousedown', e => { if (e.target === m) m.remove(); });
  m.addEventListener('keydown', e => { if (e.key === 'Escape') m.remove(); });
  $$('[data-close]', m).forEach(b => b.addEventListener('click', () => m.remove()));
  document.body.appendChild(m);
  const f = $('.modal-b input:not([disabled]), .modal-b select, .modal-b textarea', m);
  setTimeout(() => (f || $('.modal', m)).focus?.(), 50);
  return m;
}
function confirmar(msg, okText = 'Aceptar') {
  return new Promise(res => {
    const m = modal('Confirmar', `<p>${esc(msg)}</p>`, { footer: `<button class="btn" data-close>Cancelar</button><button class="btn primary" id="cf-ok">${esc(okText)}</button>` });
    $$('[data-close]', m).forEach(b => b.addEventListener('click', () => res(false)));
    $('#cf-ok', m).onclick = () => { m.remove(); res(true); };
  });
}
function pedirTexto(title, label, okText = 'Aceptar') {
  return new Promise(res => {
    const m = modal(title, `<label>${esc(label)}<input id="pt-v"></label>`, { footer: `<button class="btn" data-close>Cancelar</button><button class="btn primary" id="pt-ok">${esc(okText)}</button>` });
    $$('[data-close]', m).forEach(b => b.addEventListener('click', () => res(null)));
    const ok = () => { const v = $('#pt-v', m).value.trim(); m.remove(); res(v); };
    $('#pt-ok', m).onclick = ok;
    $('#pt-v', m).addEventListener('keydown', e => { if (e.key === 'Enter') ok(); });
  });
}
function table(cols, rows, empty = 'Sin registros para mostrar') {
  if (!rows.length) return `<div class="empty">${esc(empty)}</div>`;
  return `<div class="tbl-wrap"><table class="tbl"><thead><tr>${cols.map(c => `<th class="${c.c || ''}">${c.t}</th>`).join('')}</tr></thead>
    <tbody>${rows.map(r => `<tr class="${r._cls || ''}">${cols.map(c => `<td class="${c.c || ''}">${c.f ? c.f(r) : esc(r[c.k])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}
const badge = e => `<span class="badge b-${esc(String(e || '').toLowerCase())}">${esc(e)}</span>`;
const ENTREGAS = ['PENDIENTE', 'EN PREPARACIÓN', 'DESPACHADO', 'ENTREGADO'];
const ENT_CLS = { 'PENDIENTE': 'b-pendiente', 'EN PREPARACIÓN': 'b-parcial', 'DESPACHADO': 'b-cerrada', 'ENTREGADO': 'b-pagada' };
const entTxt = e => e ? e[0] + e.slice(1).toLowerCase() : '';
const badgeEnt = e => e ? `<span class="badge ${ENT_CLS[e] || ''}">${esc(entTxt(e))}</span>` : '';
const pageHead = (t, sub = '', actions = '') => `<div class="page-h"><div><h1>${esc(t)}</h1>${sub ? `<div class="sub">${sub}</div>` : ''}</div><div class="actions">${actions}</div></div>`;
const btnA = (act, id, txt, cls = '') => `<button class="btn sm ${cls}" data-act="${act}" data-id="${id}">${txt}</button>`;

function excel(nombre, hojas) {
  if (!window.XLSX) return toast('No se pudo cargar la librería de Excel', 'err');
  const wb = XLSX.utils.book_new();
  hojas.forEach(h => {
    const rows = h.rows.length ? h.rows : [{ Mensaje: 'Sin datos para el periodo seleccionado' }];
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = Object.keys(rows[0]).map(k => ({ wch: Math.min(48, Math.max(10, k.length + 2, ...rows.slice(0, 300).map(r => String(r[k] ?? '').length + 1))) }));
    XLSX.utils.book_append_sheet(wb, ws, h.name.replace(/[\\/?*[\]:]/g, '').slice(0, 31));
  });
  XLSX.writeFile(wb, `${nombre.replace(/[^\w-]+/g, '_')}_${hoy()}.xlsx`);
  toast('Informe descargado');
}

/* ---------------------------------------------------------------------
   Tipos de documento
   --------------------------------------------------------------------- */
const DT = {
  COT: { n: 'Cotización', pl: 'Cotizaciones', ter: 'cliente', ruta: 'cotizaciones', nuevo: 'Nueva cotización', estados: ['ABIERTA', 'APROBADA', 'RECHAZADA', 'CONVERTIDA', 'ANULADA'] },
  REM: { n: 'Remisión', pl: 'Remisiones', ter: 'cliente', ruta: 'remisiones', nuevo: 'Nueva remisión', bod: true, salida: true, estados: ['PENDIENTE', 'FACTURADA', 'ANULADA'] },
  FAC: { n: 'Factura de venta', pl: 'Facturas de venta', ter: 'cliente', ruta: 'facturas', nuevo: 'Nueva factura', bod: true, salida: true, estados: ['EMITIDA', 'PAGADA', 'ANULADA'] },
  OC: { n: 'Orden de compra', pl: 'Órdenes de compra', ter: 'proveedor', ruta: 'ordenes', nuevo: 'Nueva orden de compra', compra: true, estados: ['PENDIENTE', 'PARCIAL', 'RECIBIDA', 'CERRADA', 'ANULADA'] },
  COM: { n: 'Compra', pl: 'Compras', ter: 'proveedor', ruta: 'compras', nuevo: 'Registrar compra', bod: true, compra: true, estados: ['RECIBIDA', 'ANULADA'] },
  AJU: { n: 'Ajuste de inventario', pl: 'Ajustes', ruta: 'ajustes' },
  TRA: { n: 'Traslado entre bodegas', pl: 'Traslados', ruta: 'ajustes' },
};

/* ---------------------------------------------------------------------
   Rutas
   --------------------------------------------------------------------- */
const ROUTES = {
  inicio: { t: 'Inicio', g: '', f: vInicio },
  cotizaciones: { t: 'Cotizaciones', g: 'Ventas', f: () => vDocs('COT') },
  remisiones: { t: 'Remisiones', g: 'Ventas', f: () => vDocs('REM') },
  facturas: { t: 'Facturas', g: 'Ventas', f: () => vDocs('FAC') },
  entregas: { t: 'Seguimiento de entregas', g: 'Ventas', f: vEntregas },
  clientes: { t: 'Clientes', g: 'Ventas', f: () => vTerceros('clientes') },
  cartera: { t: 'Cartera', g: 'Ventas', f: vCartera },
  ordenes: { t: 'Órdenes de compra', g: 'Compras', f: () => vDocs('OC') },
  compras: { t: 'Compras', g: 'Compras', f: () => vDocs('COM') },
  proveedores: { t: 'Proveedores', g: 'Compras', f: () => vTerceros('proveedores') },
  existencias: { t: 'Existencias', g: 'Inventario', f: vExistencias },
  kardex: { t: 'Kardex', g: 'Inventario', f: id => vKardex(id) },
  informes: { t: 'Informes Excel', g: 'Inventario', f: vInformes },
  productos: { t: 'Productos', g: 'Administración', admin: true, f: vProductos },
  bodegas: { t: 'Bodegas', g: 'Administración', admin: true, f: vBodegas },
  ajustes: { t: 'Ajustes y traslados', g: 'Administración', admin: true, f: vAjustes },
  usuarios: { t: 'Usuarios y roles', g: 'Administración', admin: true, f: vUsuarios },
  empresa: { t: 'Empresa y consecutivos', g: 'Administración', admin: true, f: vEmpresa },
  nuevo: { hidden: true, f: (t, o, m) => vEditor(t, o, m) },
};

async function route() {
  if (!S.perfil) return;
  const raw = decodeURIComponent(location.hash.slice(1) || 'inicio');
  const [key, ...args] = raw.split('/');
  const R = ROUTES[key] || ROUTES.inicio;
  const activa = key === 'nuevo' ? DT[args[0]]?.ruta : (ROUTES[key] ? key : 'inicio');
  $$('#nav a').forEach(a => a.classList.toggle('active', a.dataset.r === activa));
  document.body.classList.remove('nav-open');
  const main = $('#main');
  main.onclick = null;
  if (R.admin && !isAdmin()) { main.innerHTML = `<div class="empty">Esta sección es solo para administradores.</div>`; return; }
  main.innerHTML = '<div class="loading">Cargando…</div>';
  window.scrollTo(0, 0);
  try { await R.f(...args); }
  catch (e) { console.error(e); main.innerHTML = `<div class="card empty err">${esc(e.message)}</div>`; }
}

/* ---------------------------------------------------------------------
   Arranque, sesión y login
   --------------------------------------------------------------------- */
document.addEventListener('DOMContentLoaded', init);

async function init() {
  if (!window.supabase || !CFG.SUPABASE_URL || /TU-PROYECTO/.test(CFG.SUPABASE_URL)) {
    $('#app').innerHTML = `<div class="login-wrap"><div class="card login"><div class="logo big">SC</div><h1>Falta la configuración</h1>
      <p class="muted">Abre el archivo <b>config.js</b> y escribe la URL y la clave <b>anon public</b> de tu proyecto Supabase (Project Settings › API). Luego recarga esta página.</p></div></div>`;
    return;
  }
  sb = window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_ANON_KEY);
  sb.auth.onAuthStateChange(ev => {
    setTimeout(() => {
      if (ev === 'PASSWORD_RECOVERY') recoveryModal();
      if (ev === 'SIGNED_OUT') { S.perfil = null; showLogin(S.loginMsg); S.loginMsg = ''; }
    }, 0);
  });
  window.addEventListener('hashchange', route);
  const { data: { session } } = await sb.auth.getSession();
  if (session) await startApp(session.user); else showLogin();
}

async function showLogin(msg = '') {
  let hay = true;
  try { const { data, error } = await sb.rpc('hay_usuarios'); if (!error) hay = data !== false; } catch (e) { /* sin conexión */ }
  $('#app').innerHTML = `<div class="login-wrap"><div class="card login">
    <div class="logo big">SC</div>
    <h1>${hay ? 'Iniciar sesión' : 'Configuración inicial'}</h1>
    <p class="muted">${hay ? 'Ingresa con el usuario que te asignó el administrador.' : 'Crea la cuenta del administrador principal. Tendrá control total del sistema y podrá crear a los demás usuarios.'}</p>
    ${msg ? `<div class="alert">${esc(msg)}</div>` : ''}
    <form id="lf">
      ${hay ? '' : '<label>Nombre completo<input id="l-nom" required></label>'}
      <label>Correo electrónico<input type="email" id="l-em" required autocomplete="username"></label>
      <label>Contraseña<input type="password" id="l-pw" required minlength="6" autocomplete="${hay ? 'current-password' : 'new-password'}"></label>
      <button class="btn primary lg block" id="l-btn">${hay ? 'Entrar' : 'Crear administrador'}</button>
    </form>
    ${hay ? '<button class="link" id="l-forgot" type="button">Olvidé mi contraseña</button>' : ''}
  </div></div>`;
  $('#lf').onsubmit = e => {
    e.preventDefault();
    run(async () => {
      const email = $('#l-em').value.trim(), password = $('#l-pw').value;
      if (hay) {
        const data = chk(await sb.auth.signInWithPassword({ email, password }));
        await startApp(data.user);
      } else {
        const data = chk(await sb.auth.signUp({ email, password, options: { data: { nombre: $('#l-nom').value.trim() } } }));
        if (data.session) await startApp(data.user);
        else showLogin('Cuenta creada. Revisa tu correo para confirmarla y luego inicia sesión.');
      }
    }, $('#l-btn'));
  };
  const fg = $('#l-forgot');
  if (fg) fg.onclick = () => run(async () => {
    const email = $('#l-em').value.trim();
    if (!email) throw new Error('Escribe tu correo y vuelve a pulsar "Olvidé mi contraseña"');
    chk(await sb.auth.resetPasswordForEmail(email, { redirectTo: location.href.split('#')[0] }));
    toast('Te enviamos un correo para restablecer la contraseña');
  });
}

async function startApp(user) {
  const p = chk(await sb.from('perfiles').select('*').eq('id', user.id).maybeSingle());
  if (!p || !p.activo) {
    S.loginMsg = 'Tu cuenta está pendiente de activación. Pide a un administrador que la active.';
    await sb.auth.signOut();
    return;
  }
  S.perfil = p;
  await loadCatalogs();
  renderShell();
  route();
}

async function loadCatalogs() {
  const [bod, prod, cli, prov, emp] = await Promise.all([
    fetchAll(() => sb.from('bodegas').select('*').order('nombre')),
    fetchAll(() => sb.from('productos').select('*').order('nombre')),
    fetchAll(() => sb.from('clientes').select('*').order('nombre')),
    fetchAll(() => sb.from('proveedores').select('*').order('nombre')),
    sb.from('empresa').select('*').eq('id', 1).maybeSingle(),
  ]);
  Object.assign(S, { bodegas: bod, productos: prod, clientes: cli, proveedores: prov, empresa: emp.data || {} });
}

function renderShell() {
  const groups = {};
  Object.entries(ROUTES).forEach(([k, r]) => {
    if (r.hidden || (r.admin && !isAdmin())) return;
    (groups[r.g] = groups[r.g] || []).push([k, r]);
  });
  const ini = (S.empresa.nombre || 'Sistema Comercial').split(/\s+/).map(w => w[0]).join('').slice(0, 2).toUpperCase();
  $('#app').innerHTML = `<div class="shell">
    <aside class="side">
      <div class="brand"><div class="logo">${esc(ini)}</div><div><b>${esc(S.empresa.nombre || 'Sistema Comercial')}</b><small>Gestión comercial</small></div></div>
      <nav id="nav">${Object.entries(groups).map(([g, items]) => `${g ? `<div class="nav-g">${esc(g)}</div>` : ''}${items.map(([k, r]) => `<a href="#${k}" data-r="${k}">${esc(r.t)}</a>`).join('')}`).join('')}</nav>
    </aside>
    <div class="side-bg" id="side-bg"></div>
    <div class="content">
      <header class="top">
        <button class="btn icon ghost menu" id="menu-btn" aria-label="Menú">☰</button>
        <div class="grow"></div>
        <div class="user"><b>${esc(S.perfil.nombre || S.perfil.email)}</b>${badge(S.perfil.rol === 'admin' ? 'Admin' : 'Trabajador')}</div>
        <button class="btn sm ghost" id="btn-pass">Cambiar contraseña</button>
        <button class="btn sm" id="btn-out">Salir</button>
      </header>
      <main id="main"></main>
    </div></div>`;
  $('#menu-btn').onclick = () => document.body.classList.toggle('nav-open');
  $('#side-bg').onclick = () => document.body.classList.remove('nav-open');
  $('#btn-out').onclick = () => sb.auth.signOut();
  $('#btn-pass').onclick = () => recoveryModal(true);
}

function recoveryModal(propio = false) {
  const m = modal(propio ? 'Cambiar mi contraseña' : 'Nueva contraseña', `<div class="form-grid">
    <label class="span-all">Nueva contraseña<input type="password" id="np-1" minlength="6" autocomplete="new-password"></label>
    <label class="span-all">Repite la contraseña<input type="password" id="np-2" minlength="6" autocomplete="new-password"></label></div>`,
    { footer: `<button class="btn" data-close>Cancelar</button><button class="btn primary" id="np-ok">Guardar contraseña</button>` });
  $('#np-ok', m).onclick = e => run(async () => {
    const a = $('#np-1', m).value, b = $('#np-2', m).value;
    if (a.length < 6) throw new Error('La contraseña debe tener al menos 6 caracteres');
    if (a !== b) throw new Error('Las contraseñas no coinciden');
    chk(await sb.auth.updateUser({ password: a }));
    toast('Contraseña actualizada'); m.remove();
    if (!S.perfil) { const { data: { session } } = await sb.auth.getSession(); if (session) startApp(session.user); }
  }, e.currentTarget);
}

/* ---------------------------------------------------------------------
   INICIO (tablero)
   --------------------------------------------------------------------- */
async function vInicio() {
  const mes = inicioMes(), h = hoy();
  const [fac, cart, inv, ocp, remp, entp] = await Promise.all([
    fetchAll(() => sb.from('documentos').select('id,numero,fecha,total,saldo,estado,clientes(nombre)').eq('tipo', 'FAC').neq('estado', 'ANULADA').gte('fecha', mes).lte('fecha', h).order('id', { ascending: false })),
    fetchAll(() => sb.from('documentos').select('saldo,fecha,fecha_vencimiento').eq('tipo', 'FAC').gt('saldo', 0).neq('estado', 'ANULADA')),
    fetchAll(() => sb.from('inventario').select('producto_id,cantidad,costo_promedio')),
    sb.from('documentos').select('id', { count: 'exact', head: true }).eq('tipo', 'OC').in('estado', ['PENDIENTE', 'PARCIAL']),
    sb.from('documentos').select('id', { count: 'exact', head: true }).eq('tipo', 'REM').eq('estado', 'PENDIENTE'),
    sb.from('documentos').select('id', { count: 'exact', head: true }).in('tipo', ['REM', 'FAC']).neq('estado', 'ANULADA').neq('estado', 'FACTURADA').neq('estado_entrega', 'ENTREGADO'),
  ]);
  const ventasMes = fac.reduce((a, r) => a + +r.total, 0);
  const ventasHoy = fac.filter(r => fd(r.fecha) === h).reduce((a, r) => a + +r.total, 0);
  const cartera = cart.reduce((a, r) => a + +r.saldo, 0);
  const vencida = cart.filter(r => (r.fecha_vencimiento || r.fecha) < h).reduce((a, r) => a + +r.saldo, 0);
  const valorInv = inv.reduce((a, r) => a + r.cantidad * r.costo_promedio, 0);
  const stock = {}; inv.forEach(r => stock[r.producto_id] = (stock[r.producto_id] || 0) + +r.cantidad);
  const bajos = S.productos.filter(p => p.activo && +p.stock_minimo > 0 && (stock[p.id] || 0) < +p.stock_minimo)
    .map(p => ({ ...p, stock: stock[p.id] || 0 }));

  $('#main').innerHTML = pageHead(`Hola, ${(S.perfil.nombre || '').split(' ')[0] || 'bienvenido'}`, `Resumen al ${h}`,
    `<div class="quick"><button class="btn" onclick="go('nuevo/COT')">Cotizar</button><button class="btn" onclick="go('nuevo/COM')">Registrar compra</button><button class="btn primary" onclick="go('nuevo/FAC')">Facturar</button></div>`) +
    `<div class="kpis">
      <div class="kpi accent"><span>Ventas del mes</span><b>${fm(ventasMes)}</b><small>${fac.length} factura(s) · hoy ${fm(ventasHoy)}</small></div>
      <div class="kpi"><span>Cartera por cobrar</span><b>${fm(cartera)}</b><small>${cart.length} factura(s) con saldo</small></div>
      <div class="kpi ${vencida > 0 ? 'bad' : ''}"><span>Cartera vencida</span><b>${fm(vencida)}</b><small>Facturas con fecha de vencimiento pasada</small></div>
      <div class="kpi"><span>Inventario valorizado</span><b>${fm(valorInv)}</b><small>A costo promedio, todas las bodegas</small></div>
      <div class="kpi ${bajos.length ? 'warn' : ''}"><span>Bajo stock mínimo</span><b>${bajos.length}</b><small>Productos por reponer</small></div>
      <a class="kpi ${entp.count ? 'warn' : ''}" href="#entregas" style="text-decoration:none;color:inherit"><span>Pedidos por entregar</span><b>${entp.count || 0}</b><small>Remisiones y facturas sin entregar · ver seguimiento</small></a>
      <div class="kpi"><span>Pendientes</span><b>${ocp.count || 0} / ${remp.count || 0}</b><small>Órdenes por recibir / remisiones por facturar</small></div>
    </div>
    <div class="grid-2">
      <div class="card"><div class="card-h"><h3>Últimas facturas del mes</h3><a href="#facturas">Ver todas</a></div>
        ${table([{ t: 'Número', f: r => `<a class="clickable" data-ver="${r.id}">${esc(r.numero)}</a>` }, { t: 'Cliente', f: r => esc(r.clientes?.nombre) }, { t: 'Total', c: 'num', f: r => fm(r.total) }, { t: 'Estado', f: r => badge(r.estado) }], fac.slice(0, 8), 'Aún no hay facturas este mes')}</div>
      <div class="card"><div class="card-h"><h3>Productos por reponer</h3><a href="#existencias">Ver existencias</a></div>
        ${table([{ t: 'Producto', f: r => esc(prodLabel(r)) }, { t: 'Stock', c: 'num', f: r => `<span class="txt-err">${fq(r.stock)}</span>` }, { t: 'Mínimo', c: 'num', f: r => fq(r.stock_minimo) }], bajos.slice(0, 10), 'Todo el inventario está sobre el mínimo')}</div>
    </div>`;
  $('#main').onclick = e => { const a = e.target.closest('[data-ver]'); if (a) run(() => verDoc(+a.dataset.ver)); };
}

/* ---------------------------------------------------------------------
   LISTADO DE DOCUMENTOS (COT, REM, FAC, OC, COM)
   --------------------------------------------------------------------- */
function docActions(r) {
  const T = r.tipo, E = r.estado, a = [btnA('ver', r.id, 'Ver')];
  if (T === 'COT' && ['ABIERTA', 'APROBADA'].includes(E)) {
    if (E === 'ABIERTA') a.push(btnA('aprobar', r.id, 'Aprobar'));
    a.push(btnA('rechazar', r.id, 'Rechazar'), btnA('rem', r.id, 'Remisionar'), btnA('fac', r.id, 'Facturar', 'primary'));
  }
  if (T === 'COT' || T === 'OC') a.push(btnA('copiar', r.id, 'Duplicar'));
  if (T === 'REM' && E === 'PENDIENTE') a.push(btnA('fac', r.id, 'Facturar', 'primary'));
  if (['REM', 'FAC'].includes(T) && E !== 'ANULADA') a.push(btnA('entrega', r.id, 'Entrega'));
  if (T === 'FAC' && +r.saldo > 0 && E !== 'ANULADA') a.push(btnA('abonar', r.id, 'Abonar', 'primary'));
  if (T === 'OC' && ['PENDIENTE', 'PARCIAL'].includes(E)) a.push(btnA('cerrar', r.id, 'Cerrar'), btnA('recibir', r.id, 'Recibir mercancía', 'primary'));
  if (isAdmin() && E !== 'ANULADA' && !(T === 'REM' && E === 'FACTURADA') && E !== 'CONVERTIDA') a.push(btnA('anular', r.id, 'Anular', 'danger'));
  return a.join('');
}

async function docAction(act, r, reload) {
  switch (act) {
    case 'ver': return run(() => verDoc(r.id));
    case 'rem': return go(`nuevo/REM/${r.id}`);
    case 'fac': return go(`nuevo/FAC/${r.id}`);
    case 'recibir': return go(`nuevo/COM/${r.id}`);
    case 'copiar': return go(`nuevo/${r.tipo}/${r.id}/copia`);
    case 'aprobar': case 'rechazar': case 'cerrar': {
      const est = { aprobar: 'APROBADA', rechazar: 'RECHAZADA', cerrar: 'CERRADA' }[act];
      if (act === 'cerrar' && !await confirmar(`¿Cerrar la orden ${r.numero}? Las cantidades pendientes ya no se recibirán.`, 'Cerrar orden')) return;
      return run(async () => { chk(await sb.rpc('cambiar_estado', { p_id: r.id, p_estado: est })); toast(`${r.numero}: ${est.toLowerCase()}`); reload?.(); });
    }
    case 'abonar': return pagoModal(r, reload);
    case 'entrega': return entregaModal(r, reload);
    case 'anular': {
      const motivo = await pedirTexto(`Anular ${r.numero}`, 'Motivo de la anulación (el inventario se reversará en el kardex)', 'Anular documento');
      if (motivo === null) return;
      return run(async () => { chk(await sb.rpc('anular_documento', { p_id: r.id, p_motivo: motivo || 'Sin motivo' })); toast(`${r.numero} anulado`); reload?.(); });
    }
  }
}

async function vDocs(tipo) {
  const d = DT[tipo], esCli = d.ter === 'cliente';
  $('#main').innerHTML = pageHead(d.pl, '', `<button class="btn" id="btn-excel">Exportar a Excel</button><button class="btn primary" id="btn-nuevo">${d.nuevo}</button>`) +
    `<div class="card filters">
      <label>Desde<input type="date" id="f-desde" value="${inicioMes()}"></label>
      <label>Hasta<input type="date" id="f-hasta" value="${hoy()}"></label>
      <label>Estado<select id="f-estado"><option value="">Todos</option>${d.estados.map(e => `<option>${e}</option>`).join('')}</select></label>
      <label class="grow">Buscar<input id="f-q" placeholder="Número, ${esCli ? 'cliente' : 'proveedor'} o documento"></label>
    </div><div class="card" id="lista"></div>`;
  let rows = [];
  const load = async () => {
    $('#lista').innerHTML = '<div class="loading">Cargando…</div>';
    rows = await fetchAll(() => {
      let q = sb.from('documentos').select('*, clientes(nombre,documento), proveedores(nombre,documento), bodega:bodegas!documentos_bodega_id_fkey(nombre)')
        .eq('tipo', tipo).gte('fecha', $('#f-desde').value).lte('fecha', $('#f-hasta').value).order('id', { ascending: false });
      const e = $('#f-estado').value; if (e) q = q.eq('estado', e);
      return q;
    });
    paint();
  };
  const filtered = () => {
    const q = $('#f-q').value.trim().toLowerCase();
    return !q ? rows : rows.filter(r => [r.numero, r.clientes?.nombre, r.clientes?.documento, r.proveedores?.nombre, r.proveedores?.documento, r.referencia].join(' ').toLowerCase().includes(q));
  };
  const paint = () => {
    const fr = filtered(), tot = fr.filter(r => r.estado !== 'ANULADA').reduce((a, r) => a + +r.total, 0);
    $('#lista').innerHTML = table([
      { t: 'Número', f: r => `<b>${esc(r.numero)}</b>` },
      { t: 'Fecha', f: r => fd(r.fecha) },
      { t: esCli ? 'Cliente' : 'Proveedor', f: r => esc((r.clientes || r.proveedores)?.nombre) },
      ...(d.bod ? [{ t: 'Bodega', f: r => esc(r.bodega?.nombre || '') }] : []),
      ...(tipo === 'COM' ? [{ t: 'Factura proveedor', k: 'referencia' }] : []),
      ...(tipo === 'FAC' ? [{ t: 'Pago', f: r => esc(r.forma_pago === 'CREDITO' ? `Crédito · vence ${fd(r.fecha_vencimiento)}` : 'Contado') }] : []),
      { t: 'Total', c: 'num', f: r => fm(r.total) },
      ...(tipo === 'FAC' ? [{ t: 'Saldo', c: 'num', f: r => fm(r.saldo) }] : []),
      ...(['REM', 'FAC'].includes(tipo) ? [{ t: 'Entrega', f: r => r.estado === 'ANULADA' ? '' : badgeEnt(r.estado_entrega) }] : []),
      { t: 'Estado', f: r => badge(r.estado) },
      { t: '', c: 'acts', f: docActions },
    ], fr, `No hay ${d.pl.toLowerCase()} en este periodo. Usa "${d.nuevo}" para crear una.`) +
      (fr.length ? `<div class="tbl-foot">${fr.length} documento(s) · Total sin anulados: <b>${fm(tot)}</b></div>` : '');
  };
  ['f-desde', 'f-hasta', 'f-estado'].forEach(id => $('#' + id).onchange = () => run(load));
  $('#f-q').oninput = paint;
  $('#btn-nuevo').onclick = () => go(`nuevo/${tipo}`);
  $('#btn-excel').onclick = () => excel(d.pl, [{
    name: d.pl, rows: filtered().map(r => ({
      Número: r.numero, Fecha: fd(r.fecha), [esCli ? 'Cliente' : 'Proveedor']: (r.clientes || r.proveedores)?.nombre, Documento: (r.clientes || r.proveedores)?.documento,
      Bodega: r.bodega?.nombre || '', Subtotal: +r.subtotal, Descuento: +r.descuento, IVA: +r.iva, Total: +r.total, Saldo: +r.saldo,
      'Forma de pago': r.forma_pago || '', Vence: fd(r.fecha_vencimiento), Referencia: r.referencia || '', Estado: r.estado, ...(['REM', 'FAC'].includes(tipo) ? { Entrega: r.estado_entrega || '', 'Fecha entrega': fdt(r.fecha_entrega), 'Recibió': r.recibido_por || '' } : {}), Observaciones: r.observaciones || '',
    }))
  }]);
  $('#lista').onclick = e => { const b = e.target.closest('[data-act]'); if (!b) return; docAction(b.dataset.act, rows.find(x => x.id == b.dataset.id), () => run(load)); };
  await load();
}

/* ---------------------------------------------------------------------
   EDITOR DE DOCUMENTOS
   --------------------------------------------------------------------- */
async function vEditor(tipo, origenId, modo) {
  const d = DT[tipo];
  if (!d || ['AJU', 'TRA'].includes(tipo)) return go('inicio');
  const copia = modo === 'copia';
  let origen = null, oItems = [];
  if (origenId) {
    origen = chk(await sb.from('documentos').select('*').eq('id', origenId).single());
    oItems = chk(await sb.from('documento_items').select('*').eq('documento_id', origenId).order('id'));
  }
  const desdeRem = tipo === 'FAC' && origen?.tipo === 'REM' && !copia;
  const desdeOC = tipo === 'COM' && origen?.tipo === 'OC' && !copia;
  const esCli = d.ter === 'cliente', terTabla = esCli ? 'clientes' : 'proveedores';
  const terceros = () => (esCli ? S.clientes : S.proveedores).filter(t => t.activo);
  const prods = S.productos.filter(p => p.activo);
  const prodById = new Map(S.productos.map(p => [p.id, p]));
  const usaBodega = d.bod && !desdeRem;
  const showStock = d.salida && !desdeRem;
  const esCompra = !!d.compra;
  let inv = new Map();
  const nl = () => ({ producto_id: null, cantidad: 1, precio: 0, descuento_pct: 0, iva_pct: 0, valor: '' });
  let lines = oItems.map(it => {
    const cantidad = desdeOC ? r3(it.cantidad - it.cantidad_recibida) : +it.cantidad;
    const l = { producto_id: it.producto_id, cantidad, precio: +it.precio, descuento_pct: +it.descuento_pct, iva_pct: +it.iva_pct, valor: '' };
    if (esCompra) { l.precio = r4(l.precio * (1 - l.descuento_pct / 100)); l.descuento_pct = 0; l.valor = r2(l.precio * cantidad); }
    return l;
  }).filter(l => l.cantidad > 0);
  if (!lines.length) lines = [nl()];
  const terIni = origen ? (esCli ? S.clientes : S.proveedores).find(t => t.id === (esCli ? origen.cliente_id : origen.proveedor_id)) : null;
  const bodIni = origen?.bodega_id || S.bodegas.find(b => b.activo)?.id || '';
  const sub = origen ? (copia ? `Copia de ${esc(origen.numero)}` : `Basada en ${DT[origen.tipo].n.toLowerCase()} <b>${esc(origen.numero)}</b>`) : '';
  const vencLabel = { COT: 'Válida hasta', OC: 'Fecha esperada de entrega', FAC: 'Fecha de vencimiento' }[tipo];
  const vencIni = { COT: hoy(15), OC: hoy(8), FAC: hoy() }[tipo];

  $('#main').innerHTML = pageHead(d.nuevo, sub, `<button class="btn" id="btn-cancel">Cancelar</button>`) +
    `<div class="card form-grid">
      <label class="span2">${esCli ? 'Cliente' : 'Proveedor'}
        <div class="row"><input id="e-ter" list="dl-ter" value="${terIni ? esc(terLabel(terIni)) : ''}" placeholder="Escribe documento o nombre y elige de la lista" autocomplete="off" ${desdeRem ? 'disabled' : ''}>
        ${desdeRem ? '' : `<button class="btn" id="btn-ter" type="button">Crear ${esCli ? 'cliente' : 'proveedor'}</button>`}</div>
        <small id="ter-info" class="muted"></small>
      </label>
      <label>Fecha<input type="date" id="e-fecha" value="${hoy()}"></label>
      ${usaBodega ? `<label>${esCompra ? 'Bodega que recibe' : 'Bodega que despacha'}<select id="e-bod">${opts(S.bodegas.filter(b => b.activo), bodIni)}</select></label>` : ''}
      ${desdeRem ? `<label>Bodega<input disabled value="${esc(S.bodegas.find(b => b.id === origen.bodega_id)?.nombre || '')}"></label>` : ''}
      ${tipo === 'FAC' ? `<label>Forma de pago<select id="e-fp"><option value="CONTADO">Contado</option><option value="CREDITO">Crédito</option></select></label>
        <label id="lb-mp">Medio de pago<select id="e-mp">${MEDIOS.map(m => `<option>${m}</option>`).join('')}</select></label>` : ''}
      ${vencLabel ? `<label id="lb-venc">${vencLabel}<input type="date" id="e-venc" value="${vencIni}"></label>` : ''}
      ${tipo === 'COM' ? `<label>N° factura del proveedor<input id="e-ref" placeholder="Ej: FE-1234"></label>` : ''}
      ${esCompra ? `<label class="chk span2"><input type="checkbox" id="e-coniva"> Los valores digitados ya incluyen IVA</label>` : ''}
      <label class="span-all">Observaciones<textarea id="e-obs" rows="2" placeholder="Condiciones, notas de entrega, garantía…"></textarea></label>
    </div>
    <div class="card">
      <div class="card-h"><h3>Productos</h3>${desdeRem ? '<small class="muted">Las cantidades vienen de la remisión y ya salieron del inventario</small>' : '<button class="btn" id="btn-add" type="button">Agregar producto</button>'}</div>
      <div class="tbl-wrap"><table class="tbl lines"><thead><tr><th>Producto</th>${showStock ? '<th class="num">Disponible</th>' : ''}<th class="num">Cantidad</th>${esCompra ? '<th class="num">Valor total del producto</th><th class="num">Costo unitario (calculado)</th>' : '<th class="num">Precio unitario</th><th class="num">Desc. %</th>'}<th class="num">IVA %</th><th class="num">Total con IVA</th><th></th></tr></thead>
      <tbody id="e-lines"></tbody></table></div>
      <datalist id="dl-prod">${prods.map(p => `<option value="${esc(prodLabel(p))}"></option>`).join('')}</datalist>
      <datalist id="dl-ter"></datalist>
      ${esCompra ? '<p class="muted"><small>Digita la cantidad y el valor total de cada producto como aparece en la factura (ya con descuentos). El sistema divide y calcula el costo unitario. Si lo prefieres, también puedes escribir el costo unitario y el total se calcula solo. Al inventario entra el valor sin IVA y el costo promedio de la bodega se recalcula automáticamente.</small></p>' : ''}
      <div class="totals" id="e-tot"></div>
      <div class="form-actions"><button class="btn primary lg" id="btn-save">Guardar ${d.n.toLowerCase()}</button></div>
    </div>`;

  const fillTer = () => { $('#dl-ter').innerHTML = terceros().map(t => `<option value="${esc(terLabel(t))}"></option>`).join(''); };
  const getTer = () => desdeRem ? terIni : findBy(terceros(), terLabel, $('#e-ter').value.trim());
  const onTer = () => {
    const t = getTer();
    $('#ter-info').textContent = t ? [t.telefono, t.ciudad, esCli ? (t.dias_credito ? `Crédito a ${t.dias_credito} días` : 'Contado') : ''].filter(Boolean).join(' · ') : '';
    if (tipo === 'FAC' && t) { $('#e-fp').value = t.dias_credito > 0 ? 'CREDITO' : 'CONTADO'; onFp(); }
  };
  const onFp = () => {
    if (tipo !== 'FAC') return;
    const cred = $('#e-fp').value === 'CREDITO';
    $('#lb-mp').style.display = cred ? 'none' : '';
    $('#lb-venc').style.display = cred ? '' : 'none';
    if (cred) $('#e-venc').value = sumarDias($('#e-fecha').value || hoy(), getTer()?.dias_credito || 30);
  };
  fillTer();
  $('#e-ter').addEventListener('change', onTer);
  if (tipo === 'FAC') { $('#e-fp').onchange = onFp; $('#e-fecha').addEventListener('change', onFp); }
  onTer(); onFp();
  if ($('#btn-ter')) $('#btn-ter').onclick = () => terceroModal(terTabla, null, t => { fillTer(); $('#e-ter').value = terLabel(t); onTer(); });

  async function loadInv() {
    inv = new Map();
    if (usaBodega) {
      (await fetchAll(() => sb.from('inventario').select('producto_id,cantidad,costo_promedio').eq('bodega_id', $('#e-bod').value))).forEach(r => inv.set(r.producto_id, r));
    } else if (esCompra) {
      const acc = {};
      (await fetchAll(() => sb.from('inventario').select('producto_id,cantidad,costo_promedio'))).forEach(r => {
        const a = acc[r.producto_id] || (acc[r.producto_id] = { cantidad: 0, valor: 0, max: 0 });
        a.cantidad += +r.cantidad; a.valor += r.cantidad * r.costo_promedio; a.max = Math.max(a.max, +r.costo_promedio);
      });
      Object.entries(acc).forEach(([k, a]) => inv.set(+k, { cantidad: a.cantidad, costo_promedio: a.cantidad > 0 ? a.valor / a.cantidad : a.max }));
    }
    renderLines();
  }
  if (usaBodega) $('#e-bod').onchange = () => run(loadInv);

  // Compras: el valor total digitado manda; el costo unitario = valor neto / cantidad
  const conIva = () => esCompra && !!$('#e-coniva')?.checked;
  const netoLinea = l => conIva() ? num(l.valor) / (1 + num(l.iva_pct) / 100) : num(l.valor);
  const syncPrecio = l => { l.precio = num(l.cantidad) > 0 ? r4(netoLinea(l) / num(l.cantidad)) : 0; };
  const syncValor = l => { const n = num(l.cantidad) * num(l.precio); l.valor = r2(conIva() ? n * (1 + num(l.iva_pct) / 100) : n); };
  const calc = l => {
    if (esCompra) { const s = netoLinea(l), iva = s * num(l.iva_pct) / 100; return { bruto: s, sub: s, iva, total: s + iva }; }
    const bruto = num(l.cantidad) * num(l.precio); const s = bruto * (1 - num(l.descuento_pct) / 100); const iva = s * num(l.iva_pct) / 100; return { bruto, sub: s, iva, total: s + iva };
  };
  const tb = $('#e-lines');
  function renderLines() {
    tb.innerHTML = lines.map((l, i) => {
      const p = prodById.get(l.producto_id), disp = +(inv.get(l.producto_id)?.cantidad || 0);
      const falta = showStock && p && num(l.cantidad) > disp;
      return `<tr data-i="${i}" class="${falta ? 'warn' : ''}">
        <td class="w-prod"><input class="in-prod" list="dl-prod" value="${p ? esc(prodLabel(p)) : ''}" placeholder="Código o nombre del producto" ${desdeRem ? 'disabled' : ''}></td>
        ${showStock ? `<td class="num ${falta ? 'txt-err' : 'muted'}">${p ? fq(disp) + ' ' + esc(p.unidad) : ''}</td>` : ''}
        <td><input type="number" class="in-n" data-f="cantidad" step="any" min="0" value="${l.cantidad}" ${desdeRem ? 'disabled' : ''}></td>
        ${esCompra
          ? `<td><input type="number" class="in-n" data-f="valor" step="any" min="0" value="${l.valor}" placeholder="Valor de la factura"></td>
             <td><input type="number" class="in-n" data-f="precio" step="any" min="0" value="${l.precio}"></td>`
          : `<td><input type="number" class="in-n" data-f="precio" step="any" min="0" value="${l.precio}"></td>
             <td><input type="number" class="in-n xs" data-f="descuento_pct" step="any" min="0" max="100" value="${l.descuento_pct}"></td>`}
        <td><input type="number" class="in-n xs" data-f="iva_pct" step="any" min="0" value="${l.iva_pct}"></td>
        <td class="num lt">${fm(calc(l).total)}</td>
        <td>${desdeRem ? '' : '<button class="btn icon ghost" data-del type="button" aria-label="Quitar línea">✕</button>'}</td></tr>`;
    }).join('');
    totals();
  }
  function totals() {
    let b = 0, s = 0, iv = 0;
    lines.forEach(l => { const c = calc(l); b += c.bruto; s += c.sub; iv += c.iva; });
    $('#e-tot').innerHTML = `<div><span>Subtotal</span><b>${fm(b)}</b></div><div><span>Descuentos</span><b>−${fm(b - s)}</b></div><div><span>IVA</span><b>${fm(iv)}</b></div><div class="grand"><span>Total</span><b>${fm(s + iv)}</b></div>`;
  }
  tb.addEventListener('change', e => {
    const tr = e.target.closest('tr'); if (!tr) return;
    const i = +tr.dataset.i;
    if (e.target.classList.contains('in-prod')) {
      const p = findBy(prods, prodLabel, e.target.value.trim());
      if (!p) { if (e.target.value) toast('Elige un producto de la lista desplegable', 'err'); e.target.value = ''; lines[i].producto_id = null; return; }
      Object.assign(lines[i], { producto_id: p.id, iva_pct: +p.iva, precio: esCompra ? r2(inv.get(p.id)?.costo_promedio || 0) : +p.precio_venta });
      if (esCompra) syncValor(lines[i]);
      renderLines();
      tb.querySelector(`tr[data-i="${i}"] [data-f="cantidad"]`)?.select();
    } else if (e.target.dataset.f === 'cantidad' && showStock) renderLines();
  });
  tb.addEventListener('input', e => {
    const f = e.target.dataset.f; if (!f) return;
    const tr = e.target.closest('tr'), l = lines[+tr.dataset.i];
    l[f] = e.target.value;
    if (esCompra) {
      if (f === 'precio') { syncValor(l); $('[data-f="valor"]', tr).value = l.valor; }
      else { syncPrecio(l); $('[data-f="precio"]', tr).value = l.precio; }
    }
    $('.lt', tr).textContent = fm(calc(l).total); totals();
  });
  if (esCompra) $('#e-coniva').onchange = () => { lines.forEach(syncPrecio); renderLines(); };
  tb.addEventListener('click', e => {
    if (!e.target.closest('[data-del]')) return;
    lines.splice(+e.target.closest('tr').dataset.i, 1);
    if (!lines.length) lines.push(nl());
    renderLines();
  });
  if ($('#btn-add')) $('#btn-add').onclick = () => { lines.push(nl()); renderLines(); const ins = $$('.in-prod', tb); ins[ins.length - 1].focus(); };
  $('#btn-cancel').onclick = () => history.back();

  $('#btn-save').onclick = e => run(async () => {
    const t = getTer();
    if (!t) throw new Error(`Elige un ${d.ter} de la lista o créalo con el botón "Crear ${d.ter}"`);
    const items = lines.filter(l => l.producto_id).map(l => {
      const it = { producto_id: l.producto_id, cantidad: num(l.cantidad), precio: num(l.precio), descuento_pct: num(l.descuento_pct), iva_pct: num(l.iva_pct) };
      if (esCompra) { it.valor_total = r2(netoLinea(l)); it.descuento_pct = 0; }
      return it;
    });
    if (esCompra && items.some(i => i.valor_total <= 0) && !await confirmar('Hay productos con valor total en cero. ¿Guardar de todas formas?', 'Guardar')) return;
    if (!items.length) throw new Error('Agrega al menos un producto');
    if (items.some(i => i.cantidad <= 0)) throw new Error('Hay productos con cantidad en cero');
    const p = { tipo, fecha: $('#e-fecha').value, observaciones: $('#e-obs').value.trim() || null, items };
    p[esCli ? 'cliente_id' : 'proveedor_id'] = t.id;
    if (usaBodega) p.bodega_id = +$('#e-bod').value;
    if (origen && !copia) p.origen_id = origen.id;
    if (tipo === 'FAC') { p.forma_pago = $('#e-fp').value; p.medio_pago = $('#e-mp').value; if (p.forma_pago === 'CREDITO') p.fecha_vencimiento = $('#e-venc').value; }
    if (['COT', 'OC'].includes(tipo)) p.fecha_vencimiento = $('#e-venc').value;
    if (tipo === 'COM') p.referencia = $('#e-ref').value.trim() || null;
    const res = chk(await sb.rpc('crear_documento', { p }));
    toast(`${d.n} ${res.numero} guardada`);
    go(d.ruta);
    setTimeout(() => run(() => verDoc(res.id)), 350);
  }, e.currentTarget);

  await loadInv();
}

/* ---------------------------------------------------------------------
   VER / IMPRIMIR DOCUMENTO
   --------------------------------------------------------------------- */
async function verDoc(id) {
  const doc = chk(await sb.from('documentos').select('*, clientes(*), proveedores(*), bodega:bodegas!documentos_bodega_id_fkey(nombre), destino:bodegas!documentos_bodega_destino_id_fkey(nombre), perfiles(nombre)').eq('id', id).single());
  const items = chk(await sb.from('documento_items').select('*, productos(codigo,nombre,unidad)').eq('documento_id', id).order('id'));
  const pagos = doc.tipo === 'FAC' ? chk(await sb.from('pagos').select('*').eq('factura_id', id).order('fecha')) : [];
  const origen = doc.origen_id ? chk(await sb.from('documentos').select('numero,tipo').eq('id', doc.origen_id).single()) : null;
  const E = S.empresa || {}, ter = doc.clientes || doc.proveedores, esInv = ['AJU', 'TRA'].includes(doc.tipo);
  const vLabel = { COT: 'Válida hasta', OC: 'Entrega esperada', FAC: 'Vence' }[doc.tipo];

  const head = `<div class="dp-head">
    <div><h2>${esc(E.nombre || '')}</h2>${E.nit ? `<div>NIT ${esc(E.nit)}</div>` : ''}<div>${esc([E.direccion, E.ciudad].filter(Boolean).join(', '))}</div><div>${esc([E.telefono, E.email].filter(Boolean).join(' · '))}</div></div>
    <div class="dp-num"><div class="dp-tipo">${esc(DT[doc.tipo].n)}</div><div class="dp-n">${esc(doc.numero)}</div>
      <div>Fecha: ${fd(doc.fecha)}</div>${vLabel && doc.fecha_vencimiento ? `<div>${vLabel}: ${fd(doc.fecha_vencimiento)}</div>` : ''}<div style="margin-top:6px">${badge(doc.estado)}</div></div></div>`;
  const ter_ = ter ? `<div class="dp-ter">
      <div><span>${doc.clientes ? 'Cliente' : 'Proveedor'}</span><b>${esc(ter.nombre)}</b></div>
      <div><span>${esc(ter.tipo_doc || 'Documento')}</span>${esc(ter.documento)}</div>
      <div><span>Dirección</span>${esc([ter.direccion, ter.ciudad].filter(Boolean).join(', ') || '—')}</div>
      <div><span>Teléfono</span>${esc(ter.telefono || '—')}</div>
      ${doc.bodega ? `<div><span>Bodega</span>${esc(doc.bodega.nombre)}</div>` : ''}
      ${doc.forma_pago ? `<div><span>Forma de pago</span>${doc.forma_pago === 'CREDITO' ? 'Crédito' : 'Contado' + (doc.medio_pago ? ' · ' + esc(doc.medio_pago) : '')}</div>` : ''}
      ${doc.referencia ? `<div><span>Factura del proveedor</span>${esc(doc.referencia)}</div>` : ''}
      ${origen ? `<div><span>Documento origen</span>${esc(origen.numero)}</div>` : ''}
      ${doc.estado_entrega && doc.estado !== 'ANULADA' ? `<div><span>Entrega</span>${badgeEnt(doc.estado_entrega)}${doc.fecha_entrega ? ' ' + fdt(doc.fecha_entrega) : ''}${doc.recibido_por ? '<br>Recibió: ' + esc(doc.recibido_por) : ''}</div>` : ''}
    </div>` : `<div class="dp-ter">
      <div><span>${doc.tipo === 'TRA' ? 'Bodega origen' : 'Bodega'}</span><b>${esc(doc.bodega?.nombre || '')}</b></div>
      ${doc.destino ? `<div><span>Bodega destino</span><b>${esc(doc.destino.nombre)}</b></div>` : ''}
      <div><span>Elaborado por</span>${esc(doc.perfiles?.nombre || '')}</div></div>`;

  const tbl = esInv
    ? table([{ t: 'Código', f: i => esc(i.productos?.codigo) }, { t: 'Producto', f: i => esc(i.productos?.nombre) }, { t: 'Tipo', k: 'descripcion' },
    { t: 'Cantidad', c: 'num', f: i => fq(i.cantidad) }, { t: 'Costo unit.', c: 'num', f: i => fm(i.precio) }, { t: 'Valor', c: 'num', f: i => fm(i.total) }], items)
    : table([{ t: 'Código', f: i => esc(i.productos?.codigo) }, { t: 'Descripción', f: i => esc(i.descripcion || i.productos?.nombre) },
    { t: 'Cant.', c: 'num', f: i => fq(i.cantidad) }, ...(doc.tipo === 'OC' ? [{ t: 'Recibido', c: 'num', f: i => fq(i.cantidad_recibida) }] : []),
    { t: 'Und', f: i => esc(i.productos?.unidad) }, { t: 'Vr. unitario', c: 'num', f: i => fm(i.precio) }, { t: 'Desc.', c: 'num', f: i => +i.descuento_pct ? i.descuento_pct + '%' : '' },
    { t: 'IVA', c: 'num', f: i => i.iva_pct + '%' }, { t: 'Total', c: 'num', f: i => fm(i.total) }], items);

  const tot = esInv ? `<div class="totals"><div class="grand"><span>Valor total</span><b>${fm(doc.total)}</b></div></div>`
    : `<div class="totals"><div><span>Subtotal</span><b>${fm(doc.subtotal)}</b></div>${+doc.descuento ? `<div><span>Descuentos</span><b>−${fm(doc.descuento)}</b></div>` : ''}<div><span>IVA</span><b>${fm(doc.iva)}</b></div><div class="grand"><span>Total</span><b>${fm(doc.total)}</b></div>${doc.tipo === 'FAC' && doc.forma_pago === 'CREDITO' ? `<div><span>Saldo pendiente</span><b>${fm(doc.saldo)}</b></div>` : ''}</div>`;
  const pg = pagos.length && doc.forma_pago === 'CREDITO' ? `<h3 style="margin-top:18px">Abonos recibidos</h3>${table([{ t: 'Fecha', f: x => fd(x.fecha) }, { t: 'Medio', k: 'medio' }, { t: 'Referencia', k: 'referencia' }, { t: 'Valor', c: 'num', f: x => fm(x.valor) }, ...(isAdmin() ? [{ t: '', c: 'acts', f: x => `<button class="btn sm danger" data-delpago="${x.id}">Eliminar</button>` }] : [])], pagos)}` : '';
  const pie = [doc.tipo === 'FAC' ? E.resolucion : '', E.pie_documento].filter(Boolean).join('\n');
  const firmas = ['REM', 'OC', 'COM', 'AJU', 'TRA'].includes(doc.tipo) ? `<div class="dp-firmas"><div>Elaboró</div><div>${doc.tipo === 'REM' ? 'Recibí conforme' : 'Aprobó'}</div></div>` : '';

  const m = modal(`${DT[doc.tipo].n} ${doc.numero}`, `<div id="print-area" class="doc-print">${head}${ter_}${tbl}${tot}${doc.observaciones ? `<div class="dp-obs">${esc(doc.observaciones)}</div>` : ''}${pg}${pie ? `<div class="dp-pie">${esc(pie)}</div>` : ''}${firmas}</div>`,
    { wide: true, footer: `<button class="btn" id="dp-xls">Exportar a Excel</button><button class="btn primary" id="dp-print">Imprimir o guardar PDF</button>` });
  $('#dp-print', m).onclick = () => window.print();
  $('#dp-xls', m).onclick = () => excel(doc.numero, [{
    name: doc.numero, rows: items.map(i => ({
      Código: i.productos?.codigo, Producto: i.productos?.nombre, Cantidad: +i.cantidad, 'Valor unitario': +i.precio,
      'Desc %': +i.descuento_pct, 'IVA %': +i.iva_pct, Subtotal: +i.subtotal, IVA: +i.iva_valor, Total: +i.total,
    }))
  }]);
  m.addEventListener('click', e => {
    const b = e.target.closest('[data-delpago]'); if (!b) return;
    confirmar('¿Eliminar este abono? El saldo de la factura aumentará.', 'Eliminar abono').then(ok => ok && run(async () => {
      chk(await sb.rpc('eliminar_pago', { p_id: +b.dataset.delpago })); toast('Abono eliminado'); m.remove(); verDoc(id);
    }));
  });
}

function pagoModal(doc, reload) {
  const m = modal(`Registrar abono a ${doc.numero}`, `<div class="form-grid">
    <div class="span-all info">Saldo pendiente: <b>${fm(doc.saldo)}</b></div>
    <label>Fecha<input type="date" id="p-fecha" value="${hoy()}"></label>
    <label>Valor<input type="number" id="p-valor" step="any" min="0" value="${doc.saldo}"></label>
    <label>Medio de pago<select id="p-medio">${MEDIOS.map(x => `<option>${x}</option>`).join('')}</select></label>
    <label>Referencia<input id="p-ref" placeholder="N° de transferencia, recibo…"></label></div>`,
    { footer: `<button class="btn" data-close>Cancelar</button><button class="btn primary" id="p-ok">Registrar abono</button>` });
  $('#p-ok', m).onclick = e => run(async () => {
    chk(await sb.rpc('registrar_pago', { p: { factura_id: doc.id, fecha: $('#p-fecha', m).value, valor: num($('#p-valor', m).value), medio: $('#p-medio', m).value, referencia: $('#p-ref', m).value.trim() || null } }));
    toast('Abono registrado'); m.remove(); reload?.();
  }, e.currentTarget);
}

/* ---------------------------------------------------------------------
   SEGUIMIENTO DE ENTREGAS (remisiones y facturas)
   --------------------------------------------------------------------- */
async function entregaModal(doc, reload) {
  const log = chk(await sb.from('entregas_log').select('*, perfiles(nombre)').eq('documento_id', doc.id).order('id', { ascending: false }));
  const ter = doc.clientes?.nombre || S.clientes.find(c => c.id === doc.cliente_id)?.nombre || '';
  const m = modal(`Entrega · ${doc.numero}`, `<div class="form-grid">
      <div class="span-all info">${esc(ter)} · Total ${fm(doc.total)} · Estado actual: ${badgeEnt(doc.estado_entrega)}${doc.recibido_por ? ' · Recibió: ' + esc(doc.recibido_por) : ''}</div>
      <label>Nuevo estado<select id="en-est">${ENTREGAS.map(e => `<option value="${e}" ${e === doc.estado_entrega ? 'selected' : ''}>${entTxt(e)}</option>`).join('')}</select></label>
      <label id="en-lb-rec">Recibido por<input id="en-rec" value="${esc(doc.recibido_por || '')}" placeholder="Nombre de quien recibe"></label>
      <label class="span-all">Observación<input id="en-obs" placeholder="Ej: Transportadora, número de guía, novedad en la entrega…"></label>
    </div>
    <h3 style="margin:18px 0 8px">Historial</h3>
    ${table([{ t: 'Fecha', f: x => fdt(x.fecha) }, { t: 'Estado', f: x => badgeEnt(x.estado) }, { t: 'Recibió', k: 'recibido_por' }, { t: 'Observación', k: 'observacion' }, { t: 'Usuario', f: x => esc(x.perfiles?.nombre || '') }], log, 'Sin cambios registrados: el pedido está pendiente desde su creación')}`,
    { wide: true, footer: `<button class="btn" data-close>Cancelar</button><button class="btn primary" id="en-ok">Guardar estado</button>` });
  const sync = () => { $('#en-lb-rec', m).style.display = $('#en-est', m).value === 'ENTREGADO' ? '' : 'none'; };
  $('#en-est', m).onchange = sync; sync();
  $('#en-ok', m).onclick = e => run(async () => {
    const estado = $('#en-est', m).value, recibido_por = $('#en-rec', m).value.trim();
    if (estado === 'ENTREGADO' && !recibido_por) throw new Error('Escribe el nombre de quien recibió el pedido');
    chk(await sb.rpc('cambiar_entrega', { p: { documento_id: doc.id, estado, recibido_por, observacion: $('#en-obs', m).value.trim() } }));
    toast(`${doc.numero}: ${entTxt(estado)}`); m.remove(); reload?.();
  }, e.currentTarget);
}

async function cargarPedidos(desde, hasta, bod) {
  return fetchAll(() => {
    let q = sb.from('documentos').select('*, clientes(nombre,documento,telefono,direccion,ciudad), bodega:bodegas!documentos_bodega_id_fkey(nombre), perfiles(nombre)')
      .in('tipo', ['REM', 'FAC']).neq('estado', 'ANULADA').neq('estado', 'FACTURADA')
      .gte('fecha', desde).lte('fecha', hasta).order('fecha', { ascending: false }).order('id', { ascending: false });
    if (bod) q = q.eq('bodega_id', bod);
    return q;
  });
}
const pedidoExcel = r => ({
  Documento: r.numero, Tipo: r.tipo === 'REM' ? 'Remisión' : 'Factura', Fecha: fd(r.fecha), 'Días desde el pedido': r.estado_entrega === 'ENTREGADO' ? '' : diasEntre(fd(r.fecha), hoy()),
  Cliente: r.clientes?.nombre, Teléfono: r.clientes?.telefono || '', Dirección: [r.clientes?.direccion, r.clientes?.ciudad].filter(Boolean).join(', '),
  Bodega: r.bodega?.nombre || '', Total: +r.total, Entrega: r.estado_entrega || '', 'Fecha entrega': fdt(r.fecha_entrega), Recibió: r.recibido_por || '', Elaboró: r.perfiles?.nombre || '',
});

async function vEntregas() {
  $('#main').innerHTML = pageHead('Seguimiento de entregas', 'Estado de despacho de remisiones y facturas. Una factura hecha desde una remisión aparece una sola vez.', `<button class="btn" id="en-xls">Exportar a Excel</button>`) +
    `<div class="kpis" id="en-kpis"></div>
    <div class="card filters">
      <label>Desde<input type="date" id="en-desde" value="${hoy(-60)}"></label>
      <label>Hasta<input type="date" id="en-hasta" value="${hoy()}"></label>
      <label>Entrega<select id="en-f"><option value="POR">Por entregar</option><option value="">Todas</option>${ENTREGAS.map(e => `<option value="${e}">${entTxt(e)}</option>`).join('')}</select></label>
      <label>Bodega<select id="en-bod"><option value="">Todas</option>${opts(S.bodegas, '')}</select></label>
      <label class="grow">Buscar<input id="en-q" placeholder="Número, cliente o quien recibió"></label>
    </div><div class="card" id="en-list"></div>`;
  let rows = [];
  const load = async () => { $('#en-list').innerHTML = '<div class="loading">Cargando…</div>'; rows = await cargarPedidos($('#en-desde').value, $('#en-hasta').value, $('#en-bod').value || null); paint(); };
  const filtered = () => {
    const f = $('#en-f').value, q = $('#en-q').value.trim().toLowerCase();
    return rows.filter(r => (f === '' || (f === 'POR' ? r.estado_entrega !== 'ENTREGADO' : r.estado_entrega === f))
      && (!q || [r.numero, r.clientes?.nombre, r.clientes?.documento, r.recibido_por].join(' ').toLowerCase().includes(q)));
  };
  const paint = () => {
    const cnt = Object.fromEntries(ENTREGAS.map(e => [e, 0])); rows.forEach(r => cnt[r.estado_entrega] = (cnt[r.estado_entrega] || 0) + 1);
    $('#en-kpis').innerHTML = ENTREGAS.map(e => `<div class="kpi ${e === 'PENDIENTE' && cnt[e] ? 'warn' : ''} ${e === 'ENTREGADO' ? 'accent' : ''}"><span>${entTxt(e)}</span><b>${cnt[e]}</b></div>`).join('');
    const fr = filtered();
    fr.forEach(r => { const dias = diasEntre(fd(r.fecha), hoy()); r._dias = dias; r._cls = r.estado_entrega !== 'ENTREGADO' && dias > 3 ? 'warn' : ''; });
    $('#en-list').innerHTML = table([
      { t: 'Documento', f: r => `<a class="clickable" data-act="ver" data-id="${r.id}">${esc(r.numero)}</a><br><small class="muted">${r.tipo === 'REM' ? 'Remisión' : 'Factura'}</small>` },
      { t: 'Fecha', f: r => fd(r.fecha) + (r.estado_entrega !== 'ENTREGADO' ? `<br><small class="${r._dias > 3 ? 'txt-err' : 'muted'}">hace ${r._dias} día(s)</small>` : '') },
      { t: 'Cliente', f: r => `<b>${esc(r.clientes?.nombre)}</b><br><small class="muted">${esc([r.clientes?.telefono, r.clientes?.direccion, r.clientes?.ciudad].filter(Boolean).join(' · '))}</small>` },
      { t: 'Bodega', f: r => esc(r.bodega?.nombre || '') },
      { t: 'Total', c: 'num', f: r => fm(r.total) },
      { t: 'Entrega', f: r => badgeEnt(r.estado_entrega) + (r.estado_entrega === 'ENTREGADO' ? `<br><small class="muted">${fdt(r.fecha_entrega)} · ${esc(r.recibido_por || '')}</small>` : '') },
      { t: '', c: 'acts', f: r => btnA('entrega', r.id, r.estado_entrega === 'ENTREGADO' ? 'Ver / cambiar' : 'Actualizar entrega', r.estado_entrega === 'ENTREGADO' ? '' : 'primary') },
    ], fr, $('#en-f').value === 'POR' ? 'No hay pedidos pendientes por entregar en este periodo' : 'No hay pedidos con estos filtros');
  };
  ['en-desde', 'en-hasta', 'en-bod'].forEach(id => $('#' + id).onchange = () => run(load));
  $('#en-f').onchange = paint; $('#en-q').oninput = paint;
  $('#en-list').onclick = e => { const b = e.target.closest('[data-act]'); if (b) docAction(b.dataset.act, rows.find(x => x.id == b.dataset.id), () => run(load)); };
  $('#en-xls').onclick = () => excel('Seguimiento_entregas', [{ name: 'Entregas', rows: filtered().map(pedidoExcel) }]);
  await load();
}

/* ---------------------------------------------------------------------
   CLIENTES Y PROVEEDORES
   --------------------------------------------------------------------- */
function terceroModal(tabla, row, onSaved) {
  const esCli = tabla === 'clientes';
  const r = row || { tipo_doc: esCli ? 'CC' : 'NIT', activo: true, dias_credito: 0, cupo_credito: 0 };
  const m = modal(`${row ? 'Editar' : 'Crear'} ${esCli ? 'cliente' : 'proveedor'}`, `<div class="form-grid">
    <label>Tipo de documento<select id="c-td">${['CC', 'NIT', 'CE', 'PAS', 'TI', 'OTRO'].map(t => `<option ${t === r.tipo_doc ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
    <label>Número de documento<input id="c-doc" value="${esc(r.documento || '')}"></label>
    <label class="span-all">Nombre o razón social<input id="c-nom" value="${esc(r.nombre || '')}"></label>
    <label>Teléfono<input id="c-tel" value="${esc(r.telefono || '')}"></label>
    <label>Correo<input type="email" id="c-em" value="${esc(r.email || '')}"></label>
    <label class="span2">Dirección<input id="c-dir" value="${esc(r.direccion || '')}"></label>
    <label>Ciudad<input id="c-ciu" value="${esc(r.ciudad || '')}"></label>
    ${esCli ? `<label>Días de crédito<input type="number" id="c-dias" min="0" value="${r.dias_credito}"><small class="muted">0 = solo contado</small></label>
      <label>Cupo de crédito<input type="number" id="c-cupo" min="0" step="any" value="${r.cupo_credito}"><small class="muted">0 = sin límite</small></label>` : ''}
    ${isAdmin() && row ? `<label class="chk span-all"><input type="checkbox" id="c-act" ${r.activo ? 'checked' : ''}> Activo (los inactivos no aparecen en documentos nuevos)</label>` : ''}
  </div>`, { footer: `<button class="btn" data-close>Cancelar</button><button class="btn primary" id="c-ok">Guardar</button>` });
  $('#c-ok', m).onclick = e => run(async () => {
    const data = {
      tipo_doc: $('#c-td', m).value, documento: $('#c-doc', m).value.trim(), nombre: $('#c-nom', m).value.trim(),
      telefono: $('#c-tel', m).value.trim() || null, email: $('#c-em', m).value.trim() || null,
      direccion: $('#c-dir', m).value.trim() || null, ciudad: $('#c-ciu', m).value.trim() || null,
    };
    if (!data.documento || !data.nombre) throw new Error('El documento y el nombre son obligatorios');
    if (esCli) { data.dias_credito = parseInt($('#c-dias', m).value) || 0; data.cupo_credito = num($('#c-cupo', m).value); }
    if (isAdmin() && row) data.activo = $('#c-act', m).checked;
    const saved = chk(row ? await sb.from(tabla).update(data).eq('id', row.id).select().single() : await sb.from(tabla).insert(data).select().single());
    const arr = S[tabla], i = arr.findIndex(x => x.id === saved.id);
    if (i >= 0) arr[i] = saved; else arr.push(saved);
    arr.sort((a, b) => a.nombre.localeCompare(b.nombre));
    toast(`${esCli ? 'Cliente' : 'Proveedor'} guardado`); m.remove(); onSaved?.(saved);
  }, e.currentTarget);
}

async function vTerceros(tabla) {
  const esCli = tabla === 'clientes', titulo = esCli ? 'Clientes' : 'Proveedores';
  $('#main').innerHTML = pageHead(titulo, '', `<button class="btn" id="t-xls">Exportar a Excel</button><button class="btn primary" id="t-new">Crear ${esCli ? 'cliente' : 'proveedor'}</button>`) +
    `<div class="card filters"><label class="grow">Buscar<input id="t-q" placeholder="Nombre, documento, ciudad o teléfono"></label><label class="chk"><input type="checkbox" id="t-inact"> Mostrar inactivos</label></div><div class="card" id="t-list"></div>`;
  const rows = await fetchAll(() => sb.from(tabla).select('*').order('nombre'));
  S[tabla] = rows;
  const saldos = {};
  if (esCli) (await fetchAll(() => sb.from('documentos').select('cliente_id,saldo').eq('tipo', 'FAC').gt('saldo', 0).neq('estado', 'ANULADA'))).forEach(x => saldos[x.cliente_id] = (saldos[x.cliente_id] || 0) + +x.saldo);
  const filtered = () => {
    const q = $('#t-q').value.trim().toLowerCase(), inact = $('#t-inact').checked;
    return S[tabla].filter(r => (inact || r.activo) && (!q || [r.nombre, r.documento, r.ciudad, r.telefono, r.email].join(' ').toLowerCase().includes(q)));
  };
  const paint = () => {
    $('#t-list').innerHTML = table([
      { t: 'Documento', f: r => `${esc(r.tipo_doc || '')} ${esc(r.documento)}` },
      { t: 'Nombre', f: r => `<b>${esc(r.nombre)}</b> ${r.activo ? '' : badge('INACTIVO')}` },
      { t: 'Teléfono', k: 'telefono' }, { t: 'Correo', k: 'email' }, { t: 'Ciudad', k: 'ciudad' },
      ...(esCli ? [{ t: 'Condición', f: r => r.dias_credito ? `Crédito ${r.dias_credito} días` : 'Contado' }, { t: 'Saldo en cartera', c: 'num', f: r => fm(saldos[r.id] || 0) }] : []),
      { t: '', c: 'acts', f: r => btnA('edit', r.id, 'Editar') + (esCli ? btnA('cuenta', r.id, 'Estado de cuenta') : '') },
    ], filtered(), `No hay ${titulo.toLowerCase()} registrados. Usa "Crear ${esCli ? 'cliente' : 'proveedor'}".`);
  };
  $('#t-q').oninput = paint; $('#t-inact').onchange = paint;
  $('#t-new').onclick = () => terceroModal(tabla, null, paint);
  $('#t-list').onclick = e => {
    const b = e.target.closest('[data-act]'); if (!b) return;
    const r = S[tabla].find(x => x.id == b.dataset.id);
    if (b.dataset.act === 'edit') terceroModal(tabla, r, paint);
    if (b.dataset.act === 'cuenta') run(() => estadoCuenta(r));
  };
  $('#t-xls').onclick = () => excel(titulo, [{
    name: titulo, rows: filtered().map(r => ({
      'Tipo doc': r.tipo_doc, Documento: r.documento, Nombre: r.nombre, Teléfono: r.telefono || '', Correo: r.email || '', Dirección: r.direccion || '', Ciudad: r.ciudad || '',
      ...(esCli ? { 'Días crédito': r.dias_credito, 'Cupo crédito': +r.cupo_credito, 'Saldo cartera': saldos[r.id] || 0 } : {}), Activo: r.activo ? 'Sí' : 'No',
    }))
  }]);
  paint();
}

async function estadoCuenta(c) {
  const rows = await fetchAll(() => sb.from('documentos').select('id,numero,fecha,fecha_vencimiento,forma_pago,total,saldo,estado').eq('tipo', 'FAC').eq('cliente_id', c.id).neq('estado', 'ANULADA').order('fecha', { ascending: false }));
  const pend = rows.filter(r => +r.saldo > 0), h = hoy();
  const total = pend.reduce((a, r) => a + +r.saldo, 0), vencido = pend.filter(r => (r.fecha_vencimiento || r.fecha) < h).reduce((a, r) => a + +r.saldo, 0);
  const m = modal(`Estado de cuenta · ${c.nombre}`, `<div class="kpis"><div class="kpi"><span>Saldo total</span><b>${fm(total)}</b></div><div class="kpi ${vencido ? 'bad' : ''}"><span>Vencido</span><b>${fm(vencido)}</b></div><div class="kpi"><span>Cupo</span><b>${+c.cupo_credito ? fm(c.cupo_credito) : 'Sin límite'}</b></div><div class="kpi"><span>Facturas en el historial</span><b>${rows.length}</b></div></div>
    <h3 style="margin-bottom:8px">Facturas con saldo</h3>
    ${table([{ t: 'Factura', f: r => `<a class="clickable" data-ver="${r.id}">${esc(r.numero)}</a>` }, { t: 'Fecha', f: r => fd(r.fecha) }, { t: 'Vence', f: r => fd(r.fecha_vencimiento) },
    { t: 'Días vencida', c: 'num', f: r => { const dv = diasEntre(r.fecha_vencimiento || r.fecha, h); return dv > 0 ? `<span class="txt-err">${dv}</span>` : '—'; } },
    { t: 'Total', c: 'num', f: r => fm(r.total) }, { t: 'Saldo', c: 'num', f: r => `<b>${fm(r.saldo)}</b>` }], pend, 'El cliente está al día')}`,
    { wide: true, footer: `<button class="btn" id="ec-xls">Exportar a Excel</button><button class="btn primary" data-close>Cerrar</button>` });
  m.addEventListener('click', e => { const a = e.target.closest('[data-ver]'); if (a) run(() => verDoc(+a.dataset.ver)); });
  $('#ec-xls', m).onclick = () => excel(`Estado_cuenta_${c.documento}`, [{ name: 'Estado de cuenta', rows: pend.map(r => ({ Cliente: c.nombre, Documento: c.documento, Factura: r.numero, Fecha: fd(r.fecha), Vence: fd(r.fecha_vencimiento), 'Días vencida': Math.max(0, diasEntre(r.fecha_vencimiento || r.fecha, h)), Total: +r.total, Saldo: +r.saldo })) }]);
}

/* ---------------------------------------------------------------------
   CARTERA
   --------------------------------------------------------------------- */
const EDADES = ['Al día', '1 a 30 días', '31 a 60 días', '61 a 90 días', 'Más de 90 días'];
const edadDe = dv => dv <= 0 ? EDADES[0] : dv <= 30 ? EDADES[1] : dv <= 60 ? EDADES[2] : dv <= 90 ? EDADES[3] : EDADES[4];
async function cargarCartera() {
  const h = hoy();
  return (await fetchAll(() => sb.from('documentos').select('id,tipo,numero,fecha,fecha_vencimiento,total,saldo,estado,cliente_id,clientes(nombre,documento,telefono)').eq('tipo', 'FAC').gt('saldo', 0).neq('estado', 'ANULADA').order('fecha')))
    .map(r => { const dv = diasEntre(r.fecha_vencimiento || r.fecha, h); return { ...r, dv, edad: edadDe(dv) }; });
}

async function vCartera() {
  $('#main').innerHTML = pageHead('Cartera', 'Facturas a crédito con saldo pendiente, por antigüedad de vencimiento', `<button class="btn" id="c-xls">Exportar a Excel</button>`) +
    `<div class="kpis" id="c-kpis"></div><div class="card filters"><label class="grow">Buscar cliente o factura<input id="c-q"></label><label>Antigüedad<select id="c-edad"><option value="">Todas</option>${EDADES.map(e => `<option>${e}</option>`).join('')}</select></label></div><div class="card" id="c-list"></div>`;
  let rows = [];
  const load = async () => { rows = await cargarCartera(); paint(); };
  const filtered = () => {
    const q = $('#c-q').value.trim().toLowerCase(), ed = $('#c-edad').value;
    return rows.filter(r => (!ed || r.edad === ed) && (!q || [r.numero, r.clientes?.nombre, r.clientes?.documento].join(' ').toLowerCase().includes(q)));
  };
  const paint = () => {
    const by = Object.fromEntries(EDADES.map(e => [e, 0])); rows.forEach(r => by[r.edad] += +r.saldo);
    const total = rows.reduce((a, r) => a + +r.saldo, 0);
    $('#c-kpis').innerHTML = `<div class="kpi accent"><span>Total cartera</span><b>${fm(total)}</b><small>${rows.length} factura(s)</small></div>` +
      EDADES.map((e, i) => `<div class="kpi ${i >= 3 && by[e] > 0 ? 'bad' : i >= 1 && by[e] > 0 ? 'warn' : ''}"><span>${e}</span><b>${fm(by[e])}</b></div>`).join('');
    const fr = filtered();
    $('#c-list').innerHTML = table([
      { t: 'Cliente', f: r => `<b>${esc(r.clientes?.nombre)}</b><br><small class="muted">${esc(r.clientes?.documento)} ${r.clientes?.telefono ? '· ' + esc(r.clientes.telefono) : ''}</small>` },
      { t: 'Factura', f: r => `<a class="clickable" data-act="ver" data-id="${r.id}">${esc(r.numero)}</a>` },
      { t: 'Fecha', f: r => fd(r.fecha) }, { t: 'Vence', f: r => fd(r.fecha_vencimiento) },
      { t: 'Antigüedad', f: r => r.dv > 0 ? `<span class="${r.dv > 60 ? 'txt-err' : ''}">${r.dv} días vencida</span>` : 'Al día' },
      { t: 'Total', c: 'num', f: r => fm(r.total) }, { t: 'Saldo', c: 'num', f: r => `<b>${fm(r.saldo)}</b>` },
      { t: '', c: 'acts', f: r => btnA('abonar', r.id, 'Registrar abono', 'primary') },
    ], fr, 'No hay cartera pendiente');
  };
  $('#c-q').oninput = paint; $('#c-edad').onchange = paint;
  $('#c-list').onclick = e => { const b = e.target.closest('[data-act]'); if (b) docAction(b.dataset.act, rows.find(x => x.id == b.dataset.id), () => run(load)); };
  $('#c-xls').onclick = () => excel('Cartera', carteraHojas(filtered()));
  await load();
}
function carteraHojas(rows) {
  const porCli = {};
  rows.forEach(r => {
    const k = r.cliente_id, c = porCli[k] || (porCli[k] = { Cliente: r.clientes?.nombre, Documento: r.clientes?.documento, Teléfono: r.clientes?.telefono || '', ...Object.fromEntries(EDADES.map(e => [e, 0])), Total: 0 });
    c[r.edad] += +r.saldo; c.Total += +r.saldo;
  });
  return [
    { name: 'Resumen por cliente', rows: Object.values(porCli).sort((a, b) => b.Total - a.Total) },
    { name: 'Detalle', rows: rows.map(r => ({ Cliente: r.clientes?.nombre, Documento: r.clientes?.documento, Factura: r.numero, Fecha: fd(r.fecha), Vence: fd(r.fecha_vencimiento), 'Días vencida': Math.max(0, r.dv), Antigüedad: r.edad, Total: +r.total, Saldo: +r.saldo })) },
  ];
}

/* ---------------------------------------------------------------------
   EXISTENCIAS
   --------------------------------------------------------------------- */
async function cargarExistencias() {
  return fetchAll(() => sb.from('inventario').select('producto_id,bodega_id,cantidad,costo_promedio,productos(codigo,nombre,categoria,unidad,stock_minimo,activo),bodegas(nombre)'));
}
function consolidar(rows) {
  const acc = {};
  rows.forEach(r => {
    const a = acc[r.producto_id] || (acc[r.producto_id] = { ...r, bodega_id: null, bodegas: { nombre: 'Todas' }, cantidad: 0, valor: 0 });
    a.cantidad += +r.cantidad; a.valor += r.cantidad * r.costo_promedio;
  });
  return Object.values(acc).map(a => ({ ...a, costo_promedio: a.cantidad > 0 ? a.valor / a.cantidad : 0 }));
}
async function vExistencias() {
  $('#main').innerHTML = pageHead('Existencias', 'Inventario valorizado a costo promedio ponderado', `<button class="btn" id="x-xls">Exportar a Excel</button>`) +
    `<div class="kpis" id="x-kpis"></div>
    <div class="card filters">
      <label>Bodega<select id="x-bod"><option value="">Todas</option>${opts(S.bodegas, '')}</select></label>
      <label>Vista<select id="x-vista"><option value="bod">Por bodega</option><option value="cons">Consolidado por producto</option></select></label>
      <label class="grow">Buscar<input id="x-q" placeholder="Código, nombre o categoría"></label>
      <label class="chk"><input type="checkbox" id="x-stock" checked> Solo con existencias</label>
      <label class="chk"><input type="checkbox" id="x-min"> Solo bajo mínimo</label>
    </div><div class="card" id="x-list"></div>`;
  const rows = await cargarExistencias();
  const data = () => {
    const b = $('#x-bod').value, q = $('#x-q').value.trim().toLowerCase();
    let r = rows.filter(x => !b || x.bodega_id == b);
    if ($('#x-vista').value === 'cons') r = consolidar(r);
    return r.map(x => ({ ...x, valor: x.cantidad * x.costo_promedio, bajo: +x.productos.stock_minimo > 0 && +x.cantidad < +x.productos.stock_minimo }))
      .filter(x => (!$('#x-stock').checked || +x.cantidad !== 0) && (!$('#x-min').checked || x.bajo) && (!q || [x.productos.codigo, x.productos.nombre, x.productos.categoria].join(' ').toLowerCase().includes(q)))
      .sort((a, z) => a.productos.nombre.localeCompare(z.productos.nombre));
  };
  const paint = () => {
    const d = data(), total = d.reduce((a, x) => a + x.valor, 0);
    const porBod = {}; rows.forEach(x => porBod[x.bodegas.nombre] = (porBod[x.bodegas.nombre] || 0) + x.cantidad * x.costo_promedio);
    $('#x-kpis').innerHTML = `<div class="kpi accent"><span>Valor en la vista actual</span><b>${fm(total)}</b><small>${d.length} línea(s)</small></div>` +
      Object.entries(porBod).map(([n, v]) => `<div class="kpi"><span>${esc(n)}</span><b>${fm(v)}</b></div>`).join('');
    d.forEach(x => x._cls = x.bajo ? 'warn' : '');
    $('#x-list').innerHTML = table([
      { t: 'Código', f: x => esc(x.productos.codigo) },
      { t: 'Producto', f: x => `<a class="clickable" href="#kardex/${x.producto_id}">${esc(x.productos.nombre)}</a>` },
      { t: 'Categoría', f: x => esc(x.productos.categoria || '') }, { t: 'Bodega', f: x => esc(x.bodegas.nombre) },
      { t: 'Cantidad', c: 'num', f: x => `${fq(x.cantidad)} ${esc(x.productos.unidad)}` },
      { t: 'Mínimo', c: 'num', f: x => +x.productos.stock_minimo ? fq(x.productos.stock_minimo) : '—' },
      { t: 'Costo promedio', c: 'num', f: x => fm(x.costo_promedio) }, { t: 'Valor total', c: 'num', f: x => `<b>${fm(x.valor)}</b>` },
    ], d, 'No hay existencias con estos filtros. El inventario inicial se carga en Administración › Ajustes y traslados.');
  };
  ['x-bod', 'x-vista', 'x-stock', 'x-min'].forEach(id => $('#' + id).onchange = paint);
  $('#x-q').oninput = paint;
  $('#x-xls').onclick = () => excel('Existencias', [{ name: 'Existencias', rows: data().map(x => ({ Código: x.productos.codigo, Producto: x.productos.nombre, Categoría: x.productos.categoria || '', Unidad: x.productos.unidad, Bodega: x.bodegas.nombre, Cantidad: +x.cantidad, 'Stock mínimo': +x.productos.stock_minimo, 'Costo promedio': r2(x.costo_promedio), 'Valor total': r2(x.valor), 'Bajo mínimo': x.bajo ? 'Sí' : 'No' })) }]);
  paint();
}

/* ---------------------------------------------------------------------
   KARDEX
   --------------------------------------------------------------------- */
async function vKardex(prodId) {
  const p0 = prodId ? S.productos.find(p => p.id == prodId) : null;
  $('#main').innerHTML = pageHead('Kardex', 'Movimientos valorizados con saldo y costo promedio después de cada movimiento', `<button class="btn" id="k-xls">Exportar a Excel</button>`) +
    `<div class="card filters">
      <label class="grow">Producto<input id="k-prod" list="k-dl" value="${p0 ? esc(prodLabel(p0)) : ''}" placeholder="Todos los productos"><datalist id="k-dl">${S.productos.map(p => `<option value="${esc(prodLabel(p))}"></option>`).join('')}</datalist></label>
      <label>Bodega<select id="k-bod"><option value="">Todas</option>${opts(S.bodegas, '')}</select></label>
      <label>Desde<input type="date" id="k-desde" value="${p0 ? hoy(-365) : inicioMes()}"></label>
      <label>Hasta<input type="date" id="k-hasta" value="${hoy()}"></label>
      <button class="btn primary" id="k-go">Consultar</button>
    </div><div class="card" id="k-list"></div>`;
  let rows = [];
  const load = async () => {
    const pv = $('#k-prod').value.trim(), p = pv ? findBy(S.productos, prodLabel, pv) : null;
    if (pv && !p) throw new Error('Elige un producto de la lista o deja el campo vacío para ver todos');
    $('#k-list').innerHTML = '<div class="loading">Cargando…</div>';
    rows = await fetchAll(() => {
      let q = sb.from('kardex').select('*, productos(codigo,nombre,unidad), bodegas(nombre), perfiles(nombre)')
        .gte('fecha', dIni($('#k-desde').value)).lte('fecha', dFin($('#k-hasta').value)).order('id');
      if (p) q = q.eq('producto_id', p.id);
      if ($('#k-bod').value) q = q.eq('bodega_id', $('#k-bod').value);
      return q;
    });
    const te = rows.reduce((a, r) => a + +r.entrada, 0), ts = rows.reduce((a, r) => a + +r.salida, 0);
    const ve = rows.filter(r => +r.entrada > 0).reduce((a, r) => a + +r.valor, 0), vs = rows.filter(r => +r.salida > 0).reduce((a, r) => a + +r.valor, 0);
    $('#k-list').innerHTML = table([
      { t: 'Fecha', f: r => fdt(r.fecha) },
      { t: 'Documento', f: r => r.doc_id ? `<a class="clickable" data-ver="${r.doc_id}">${esc(r.doc_numero)}</a>` : esc(r.doc_numero) },
      { t: 'Movimiento', f: r => esc(r.tipo) },
      ...(p ? [] : [{ t: 'Producto', f: r => esc(r.productos?.codigo + ' · ' + r.productos?.nombre) }]),
      { t: 'Bodega', f: r => esc(r.bodegas?.nombre) }, { t: 'Detalle', f: r => `<small>${esc(r.detalle || '')}</small>` },
      { t: 'Entrada', c: 'num', f: r => +r.entrada ? fq(r.entrada) : '' }, { t: 'Salida', c: 'num', f: r => +r.salida ? fq(r.salida) : '' },
      { t: 'Costo unit.', c: 'num', f: r => fm(r.costo_unitario) }, { t: 'Valor mov.', c: 'num', f: r => fm(r.valor) },
      { t: 'Saldo cant.', c: 'num', f: r => `<b>${fq(r.saldo_cantidad)}</b>` }, { t: 'Costo prom.', c: 'num', f: r => fm(r.saldo_costo_promedio) },
      { t: 'Saldo valor', c: 'num', f: r => fm(r.saldo_valor) }, { t: 'Usuario', f: r => `<small>${esc(r.perfiles?.nombre || '')}</small>` },
    ], rows, 'No hay movimientos en este periodo') +
      (rows.length ? `<div class="tbl-foot">${rows.length} movimiento(s) · Entradas ${fq(te)} (${fm(ve)}) · Salidas ${fq(ts)} (${fm(vs)})${p ? '' : ' · Los saldos son por producto y bodega'}</div>` : '');
  };
  $('#k-go').onclick = e => run(load, e.currentTarget);
  $('#k-list').onclick = e => { const a = e.target.closest('[data-ver]'); if (a) run(() => verDoc(+a.dataset.ver)); };
  $('#k-xls').onclick = () => excel('Kardex', [{ name: 'Kardex', rows: kardexRows(rows) }]);
  await run(load);
}
const kardexRows = rows => rows.map(r => ({
  Fecha: fdt(r.fecha), Documento: r.doc_numero, Movimiento: r.tipo, Código: r.productos?.codigo, Producto: r.productos?.nombre, Bodega: r.bodegas?.nombre, Detalle: r.detalle || '',
  Entrada: +r.entrada, Salida: +r.salida, 'Costo unitario': +r.costo_unitario, 'Valor movimiento': +r.valor,
  'Saldo cantidad': +r.saldo_cantidad, 'Costo promedio': +r.saldo_costo_promedio, 'Saldo valor': +r.saldo_valor, Usuario: r.perfiles?.nombre || '',
}));

/* ---------------------------------------------------------------------
   ADMINISTRACIÓN · PRODUCTOS
   --------------------------------------------------------------------- */
function productoModal(row, onSaved) {
  const r = row || { unidad: 'UND', iva: 19, precio_venta: 0, stock_minimo: 0, activo: true };
  const cats = [...new Set(S.productos.map(p => p.categoria).filter(Boolean))].sort();
  const m = modal(row ? 'Editar producto' : 'Crear producto', `<div class="form-grid">
    <label>Código<input id="pr-cod" value="${esc(r.codigo || '')}"></label>
    <label class="span2">Nombre<input id="pr-nom" value="${esc(r.nombre || '')}"></label>
    <label>Categoría<input id="pr-cat" list="pr-cats" value="${esc(r.categoria || '')}"><datalist id="pr-cats">${cats.map(c => `<option value="${esc(c)}">`).join('')}</datalist></label>
    <label>Unidad<input id="pr-und" value="${esc(r.unidad)}" placeholder="UND, KG, MT, CAJA…"></label>
    <label>Precio de venta<input type="number" id="pr-pv" min="0" step="any" value="${r.precio_venta}"></label>
    <label>IVA %<select id="pr-iva">${[19, 5, 0].map(v => `<option value="${v}" ${+r.iva === v ? 'selected' : ''}>${v}%</option>`).join('')}</select></label>
    <label>Stock mínimo<input type="number" id="pr-min" min="0" step="any" value="${r.stock_minimo}"></label>
    ${row ? `<label class="chk span-all"><input type="checkbox" id="pr-act" ${r.activo ? 'checked' : ''}> Activo</label>` : ''}
  </div>`, { footer: `<button class="btn" data-close>Cancelar</button><button class="btn primary" id="pr-ok">Guardar producto</button>` });
  $('#pr-ok', m).onclick = e => run(async () => {
    const data = { codigo: $('#pr-cod', m).value.trim(), nombre: $('#pr-nom', m).value.trim(), categoria: $('#pr-cat', m).value.trim() || null, unidad: $('#pr-und', m).value.trim().toUpperCase() || 'UND', precio_venta: num($('#pr-pv', m).value), iva: num($('#pr-iva', m).value), stock_minimo: num($('#pr-min', m).value) };
    if (!data.codigo || !data.nombre) throw new Error('El código y el nombre son obligatorios');
    if (row) data.activo = $('#pr-act', m).checked;
    const saved = chk(row ? await sb.from('productos').update(data).eq('id', row.id).select().single() : await sb.from('productos').insert(data).select().single());
    const i = S.productos.findIndex(x => x.id === saved.id); if (i >= 0) S.productos[i] = saved; else S.productos.push(saved);
    S.productos.sort((a, b) => a.nombre.localeCompare(b.nombre));
    toast('Producto guardado'); m.remove(); onSaved?.();
  }, e.currentTarget);
}

async function vProductos() {
  $('#main').innerHTML = pageHead('Productos', 'Catálogo de productos. El inventario inicial se carga con un ajuste de entrada.',
    `<button class="btn" id="pr-plant">Descargar plantilla</button><label class="btn" for="pr-file">Importar desde Excel</label><input type="file" id="pr-file" accept=".xlsx,.xls,.csv" hidden>
     <button class="btn" id="pr-xls">Exportar a Excel</button><button class="btn primary" id="pr-new">Crear producto</button>`) +
    `<div class="card filters"><label class="grow">Buscar<input id="pr-q" placeholder="Código, nombre o categoría"></label><label class="chk"><input type="checkbox" id="pr-inact"> Mostrar inactivos</label></div><div class="card" id="pr-list"></div>`;
  S.productos = await fetchAll(() => sb.from('productos').select('*').order('nombre'));
  const filtered = () => { const q = $('#pr-q').value.trim().toLowerCase(), ia = $('#pr-inact').checked; return S.productos.filter(p => (ia || p.activo) && (!q || [p.codigo, p.nombre, p.categoria].join(' ').toLowerCase().includes(q))); };
  const paint = () => {
    $('#pr-list').innerHTML = table([
      { t: 'Código', f: p => `<b>${esc(p.codigo)}</b>` }, { t: 'Nombre', f: p => `${esc(p.nombre)} ${p.activo ? '' : badge('INACTIVO')}` },
      { t: 'Categoría', k: 'categoria' }, { t: 'Unidad', k: 'unidad' }, { t: 'Precio venta', c: 'num', f: p => fm(p.precio_venta) },
      { t: 'IVA', c: 'num', f: p => p.iva + '%' }, { t: 'Stock mín.', c: 'num', f: p => fq(p.stock_minimo) },
      { t: '', c: 'acts', f: p => btnA('edit', p.id, 'Editar') + `<a class="btn sm" href="#kardex/${p.id}">Kardex</a>` },
    ], filtered(), 'Aún no hay productos. Créalos uno a uno o impórtalos desde Excel.');
  };
  $('#pr-q').oninput = paint; $('#pr-inact').onchange = paint;
  $('#pr-new').onclick = () => productoModal(null, paint);
  $('#pr-list').onclick = e => { const b = e.target.closest('[data-act="edit"]'); if (b) productoModal(S.productos.find(p => p.id == b.dataset.id), paint); };
  $('#pr-xls').onclick = () => excel('Productos', [{ name: 'Productos', rows: filtered().map(p => ({ codigo: p.codigo, nombre: p.nombre, categoria: p.categoria || '', unidad: p.unidad, precio_venta: +p.precio_venta, iva: +p.iva, stock_minimo: +p.stock_minimo, activo: p.activo ? 'SI' : 'NO' })) }]);
  $('#pr-plant').onclick = () => excel('Plantilla_productos', [{ name: 'Productos', rows: [{ codigo: 'P001', nombre: 'Producto de ejemplo', categoria: 'General', unidad: 'UND', precio_venta: 10000, iva: 19, stock_minimo: 5 }] }]);
  $('#pr-file').onchange = e => run(async () => {
    const f = e.target.files[0]; if (!f) return;
    const wb = XLSX.read(await f.arrayBuffer());
    const raw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    const norm = o => Object.fromEntries(Object.entries(o).map(([k, v]) => [k.toString().trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, '_'), v]));
    const rows = raw.map(norm).filter(r => String(r.codigo).trim() && String(r.nombre).trim()).map(r => ({
      codigo: String(r.codigo).trim(), nombre: String(r.nombre).trim(), categoria: String(r.categoria || '').trim() || null,
      unidad: String(r.unidad || 'UND').trim().toUpperCase() || 'UND', precio_venta: num(r.precio_venta), iva: r.iva === '' ? 19 : num(r.iva), stock_minimo: num(r.stock_minimo),
      activo: !/^(no|n|0|false)$/i.test(String(r.activo || 'si')),
    }));
    if (!rows.length) throw new Error('El archivo no tiene filas con columnas "codigo" y "nombre"');
    if (!await confirmar(`Se importarán ${rows.length} producto(s). Los códigos existentes se actualizarán.`, 'Importar')) return;
    for (let i = 0; i < rows.length; i += 500) chk(await sb.from('productos').upsert(rows.slice(i, i + 500), { onConflict: 'codigo' }));
    S.productos = await fetchAll(() => sb.from('productos').select('*').order('nombre'));
    toast(`${rows.length} producto(s) importados`); paint();
  }).finally(() => { e.target.value = ''; });
  paint();
}

/* ---------------------------------------------------------------------
   ADMINISTRACIÓN · BODEGAS
   --------------------------------------------------------------------- */
async function vBodegas() {
  $('#main').innerHTML = pageHead('Bodegas', 'Cada bodega lleva sus propias existencias y su costo promedio', `<button class="btn primary" id="b-new">Crear bodega</button>`) + `<div class="card" id="b-list"></div>`;
  const load = async () => {
    S.bodegas = await fetchAll(() => sb.from('bodegas').select('*').order('nombre'));
    const inv = await fetchAll(() => sb.from('inventario').select('bodega_id,cantidad,costo_promedio'));
    const val = {}; inv.forEach(r => val[r.bodega_id] = (val[r.bodega_id] || 0) + r.cantidad * r.costo_promedio);
    $('#b-list').innerHTML = table([{ t: 'Código', f: b => `<b>${esc(b.codigo)}</b>` }, { t: 'Nombre', f: b => `${esc(b.nombre)} ${b.activo ? '' : badge('INACTIVO')}` }, { t: 'Dirección', k: 'direccion' },
    { t: 'Valor inventario', c: 'num', f: b => fm(val[b.id] || 0) }, { t: '', c: 'acts', f: b => btnA('edit', b.id, 'Editar') }], S.bodegas);
  };
  const edit = row => {
    const r = row || { activo: true };
    const m = modal(row ? 'Editar bodega' : 'Crear bodega', `<div class="form-grid"><label>Código<input id="bo-cod" value="${esc(r.codigo || '')}"></label><label class="span2">Nombre<input id="bo-nom" value="${esc(r.nombre || '')}"></label><label class="span-all">Dirección<input id="bo-dir" value="${esc(r.direccion || '')}"></label>${row ? `<label class="chk"><input type="checkbox" id="bo-act" ${r.activo ? 'checked' : ''}> Activa</label>` : ''}</div>`,
      { footer: `<button class="btn" data-close>Cancelar</button><button class="btn primary" id="bo-ok">Guardar bodega</button>` });
    $('#bo-ok', m).onclick = e => run(async () => {
      const data = { codigo: $('#bo-cod', m).value.trim().toUpperCase(), nombre: $('#bo-nom', m).value.trim(), direccion: $('#bo-dir', m).value.trim() || null };
      if (!data.codigo || !data.nombre) throw new Error('El código y el nombre son obligatorios');
      if (row) data.activo = $('#bo-act', m).checked;
      chk(row ? await sb.from('bodegas').update(data).eq('id', row.id) : await sb.from('bodegas').insert(data));
      toast('Bodega guardada'); m.remove(); load();
    }, e.currentTarget);
  };
  $('#b-new').onclick = () => edit(null);
  $('#b-list').onclick = e => { const b = e.target.closest('[data-act="edit"]'); if (b) edit(S.bodegas.find(x => x.id == b.dataset.id)); };
  await load();
}

/* ---------------------------------------------------------------------
   ADMINISTRACIÓN · AJUSTES Y TRASLADOS
   --------------------------------------------------------------------- */
async function vAjustes() {
  $('#main').innerHTML = pageHead('Ajustes y traslados', 'Control administrativo del inventario. Cada movimiento queda registrado en el kardex con su motivo.') +
    `<div class="tabs"><button class="tab active" data-t="aju">Ajuste de inventario</button><button class="tab" data-t="tra">Traslado entre bodegas</button><button class="tab" data-t="his">Historial</button></div><div id="aj-body"></div>`;
  const tabs = { aju: () => formInventario('AJU'), tra: () => formInventario('TRA'), his: histInventario };
  $('.tabs').onclick = e => {
    const t = e.target.closest('.tab'); if (!t) return;
    $$('.tab').forEach(x => x.classList.toggle('active', x === t));
    run(tabs[t.dataset.t]);
  };
  await formInventario('AJU');
}

async function formInventario(tipo) {
  const aju = tipo === 'AJU', bods = S.bodegas.filter(b => b.activo), prods = S.productos.filter(p => p.activo);
  const MODOS = { ENTRADA: 'Entrada (suma)', SALIDA: 'Salida (resta)', CONTEO: 'Conteo físico (fija la cantidad)', COSTO: 'Corregir costo promedio' };
  const body = $('#aj-body');
  body.innerHTML = `<div class="card form-grid">
      <label>${aju ? 'Bodega' : 'Bodega de origen'}<select id="aj-bod">${opts(bods, bods[0]?.id)}</select></label>
      ${aju ? '' : `<label>Bodega de destino<select id="aj-dest">${opts(bods, bods[1]?.id)}</select></label>`}
      <label>Fecha<input type="date" id="aj-fecha" value="${hoy()}"></label>
      <label class="span-all">${aju ? 'Motivo del ajuste (obligatorio)' : 'Observaciones'}<input id="aj-obs" placeholder="${aju ? 'Ej: Inventario inicial, conteo físico de octubre, mercancía averiada…' : 'Ej: Reposición de punto de venta'}"></label>
    </div>
    <div class="card"><div class="card-h"><h3>Productos</h3><button class="btn" id="aj-add">Agregar producto</button></div>
      <div class="tbl-wrap"><table class="tbl lines"><thead><tr><th>Producto</th><th class="num">Existencia</th><th class="num">Costo actual</th>${aju ? '<th>Tipo de ajuste</th>' : ''}<th class="num">Cantidad</th>${aju ? '<th class="num">Costo unitario</th><th class="num">Efecto</th>' : ''}<th></th></tr></thead><tbody id="aj-lines"></tbody></table></div>
      <datalist id="aj-dl">${prods.map(p => `<option value="${esc(prodLabel(p))}">`).join('')}</datalist>
      ${aju ? '<p class="muted"><small>Para cargar el inventario inicial usa "Entrada" con el costo unitario de cada producto. "Conteo físico" calcula la diferencia contra el sistema y la ajusta.</small></p>' : ''}
      <div class="form-actions"><button class="btn primary lg" id="aj-save">${aju ? 'Aplicar ajuste' : 'Aplicar traslado'}</button></div></div>`;
  let inv = new Map(), lines = [];
  const nl = () => ({ producto_id: null, modo: 'ENTRADA', cantidad: '', costo: '' });
  lines.push(nl());
  const loadInv = async () => { inv = new Map(); (await fetchAll(() => sb.from('inventario').select('producto_id,cantidad,costo_promedio').eq('bodega_id', $('#aj-bod').value))).forEach(r => inv.set(r.producto_id, r)); render(); };
  const efecto = l => {
    const st = +(inv.get(l.producto_id)?.cantidad || 0), q = num(l.cantidad);
    if (!l.producto_id) return '';
    if (l.modo === 'COSTO') return 'Cambia el costo';
    const dlt = l.modo === 'ENTRADA' ? q : l.modo === 'SALIDA' ? -q : q - st;
    return `<span class="${dlt < 0 ? 'txt-err' : ''}">${dlt > 0 ? '+' : ''}${fq(dlt)} → ${fq(st + dlt)}</span>`;
  };
  const tb = $('#aj-lines');
  function render() {
    tb.innerHTML = lines.map((l, i) => {
      const p = S.productos.find(x => x.id === l.producto_id), st = inv.get(l.producto_id);
      return `<tr data-i="${i}">
        <td class="w-prod"><input class="in-prod" list="aj-dl" value="${p ? esc(prodLabel(p)) : ''}" placeholder="Código o nombre"></td>
        <td class="num muted">${p ? fq(st?.cantidad || 0) : ''}</td><td class="num muted">${p ? fm(st?.costo_promedio || 0) : ''}</td>
        ${aju ? `<td><select data-f="modo">${Object.entries(MODOS).map(([k, v]) => `<option value="${k}" ${k === l.modo ? 'selected' : ''}>${v}</option>`).join('')}</select></td>` : ''}
        <td><input type="number" data-f="cantidad" step="any" min="0" value="${l.cantidad}" ${l.modo === 'COSTO' && aju ? 'disabled' : ''} placeholder="${l.modo === 'CONTEO' ? 'Cantidad contada' : ''}"></td>
        ${aju ? `<td><input type="number" data-f="costo" step="any" min="0" value="${l.costo}" ${['SALIDA', 'CONTEO'].includes(l.modo) ? 'disabled' : ''} placeholder="${l.modo === 'ENTRADA' ? 'Costo' : l.modo === 'COSTO' ? 'Nuevo costo' : 'Promedio'}"></td><td class="num ef">${efecto(l)}</td>` : ''}
        <td><button class="btn icon ghost" data-del aria-label="Quitar">✕</button></td></tr>`;
    }).join('');
  }
  tb.addEventListener('change', e => {
    const i = +e.target.closest('tr').dataset.i, l = lines[i];
    if (e.target.classList.contains('in-prod')) {
      const p = findBy(prods, prodLabel, e.target.value.trim());
      if (!p) { if (e.target.value) toast('Elige un producto de la lista', 'err'); l.producto_id = null; e.target.value = ''; return; }
      l.producto_id = p.id; if (aju && !l.costo) l.costo = r2(inv.get(p.id)?.costo_promedio || 0) || ''; render();
    } else if (e.target.dataset.f === 'modo') { l.modo = e.target.value; render(); }
  });
  tb.addEventListener('input', e => { const f = e.target.dataset.f; if (!f || f === 'modo') return; const tr = e.target.closest('tr'), l = lines[+tr.dataset.i]; l[f] = e.target.value; if (aju) $('.ef', tr).innerHTML = efecto(l); });
  tb.addEventListener('click', e => { if (!e.target.closest('[data-del]')) return; lines.splice(+e.target.closest('tr').dataset.i, 1); if (!lines.length) lines.push(nl()); render(); });
  $('#aj-add').onclick = () => { lines.push(nl()); render(); $$('.in-prod', tb).pop().focus(); };
  $('#aj-bod').onchange = () => run(loadInv);
  $('#aj-save').onclick = e => run(async () => {
    const items = lines.filter(l => l.producto_id);
    if (!items.length) throw new Error('Agrega al menos un producto');
    const p = { bodega_id: +$('#aj-bod').value, fecha: $('#aj-fecha').value, observaciones: $('#aj-obs').value.trim() };
    let res;
    if (aju) {
      if (!p.observaciones) throw new Error('Escribe el motivo del ajuste');
      for (const l of items) {
        if (l.modo === 'ENTRADA' && (num(l.cantidad) <= 0 || l.costo === '')) throw new Error('Las entradas necesitan cantidad y costo unitario');
        if (l.modo === 'SALIDA' && num(l.cantidad) <= 0) throw new Error('Las salidas necesitan cantidad');
        if (l.modo === 'CONTEO' && l.cantidad === '') throw new Error('Escribe la cantidad contada');
        if (l.modo === 'COSTO' && l.costo === '') throw new Error('Escribe el nuevo costo');
      }
      p.items = items.map(l => ({ producto_id: l.producto_id, modo: l.modo, cantidad: num(l.cantidad), costo: l.costo === '' ? null : num(l.costo) }));
      if (!await confirmar(`Se aplicará el ajuste a ${items.length} producto(s). Esta acción modifica el inventario.`, 'Aplicar ajuste')) return;
      res = chk(await sb.rpc('ajuste_inventario', { p }));
    } else {
      p.bodega_destino_id = +$('#aj-dest').value;
      if (p.bodega_id === p.bodega_destino_id) throw new Error('La bodega de origen y destino deben ser diferentes');
      if (items.some(l => num(l.cantidad) <= 0)) throw new Error('Hay productos sin cantidad');
      p.items = items.map(l => ({ producto_id: l.producto_id, cantidad: num(l.cantidad) }));
      res = chk(await sb.rpc('traslado_inventario', { p }));
    }
    toast(`${DT[tipo].n} ${res.numero} aplicado`);
    verDoc(res.id); formInventario(tipo);
  }, e.currentTarget);
  await loadInv();
}

async function histInventario() {
  const body = $('#aj-body');
  body.innerHTML = `<div class="card filters"><label>Desde<input type="date" id="h-desde" value="${hoy(-90)}"></label><label>Hasta<input type="date" id="h-hasta" value="${hoy()}"></label></div><div class="card" id="h-list"></div>`;
  let rows = [];
  const load = async () => {
    rows = await fetchAll(() => sb.from('documentos').select('*, bodega:bodegas!documentos_bodega_id_fkey(nombre), destino:bodegas!documentos_bodega_destino_id_fkey(nombre), perfiles(nombre)').in('tipo', ['AJU', 'TRA']).gte('fecha', $('#h-desde').value).lte('fecha', $('#h-hasta').value).order('id', { ascending: false }));
    $('#h-list').innerHTML = table([{ t: 'Número', f: r => `<b>${esc(r.numero)}</b>` }, { t: 'Tipo', f: r => DT[r.tipo].n }, { t: 'Fecha', f: r => fd(r.fecha) },
    { t: 'Bodega', f: r => esc(r.bodega?.nombre) + (r.destino ? ' → ' + esc(r.destino.nombre) : '') }, { t: 'Motivo', f: r => `<small>${esc(r.observaciones || '')}</small>` },
    { t: 'Valor', c: 'num', f: r => fm(r.total) }, { t: 'Usuario', f: r => esc(r.perfiles?.nombre || '') }, { t: 'Estado', f: r => badge(r.estado) },
    { t: '', c: 'acts', f: r => btnA('ver', r.id, 'Ver') + (r.estado !== 'ANULADA' ? btnA('anular', r.id, 'Anular', 'danger') : '') }], rows, 'No hay ajustes ni traslados en este periodo');
  };
  $('#h-desde').onchange = $('#h-hasta').onchange = () => run(load);
  $('#h-list').onclick = e => { const b = e.target.closest('[data-act]'); if (b) docAction(b.dataset.act, rows.find(x => x.id == b.dataset.id), () => run(load)); };
  await load();
}

/* ---------------------------------------------------------------------
   ADMINISTRACIÓN · USUARIOS Y ROLES
   --------------------------------------------------------------------- */
async function vUsuarios() {
  $('#main').innerHTML = pageHead('Usuarios y roles', 'Administrador: acceso total. Trabajador: ventas, compras, clientes, consultas e informes; no puede ajustar inventario, anular documentos ni cambiar productos o configuración.',
    `<button class="btn primary" id="u-new">Crear usuario</button>`) + `<div class="card" id="u-list"></div>`;
  let rows = [];
  const load = async () => {
    rows = await fetchAll(() => sb.from('perfiles').select('*').order('creado_en'));
    $('#u-list').innerHTML = table([{ t: 'Nombre', f: u => `<b>${esc(u.nombre)}</b>${u.id === S.perfil.id ? ' <small class="muted">(tú)</small>' : ''}` }, { t: 'Correo', k: 'email' },
    { t: 'Rol', f: u => badge(u.rol === 'admin' ? 'Admin' : 'Trabajador') }, { t: 'Estado', f: u => u.activo ? badge('Activo').replace('b-activo', 'b-aplicado') : badge('INACTIVO') },
    { t: 'Creado', f: u => fd(u.creado_en) }, { t: '', c: 'acts', f: u => btnA('edit', u.id, 'Editar') + btnA('reset', u.id, 'Enviar enlace de contraseña') }], rows);
  };
  const edit = u => {
    const m = modal(`Editar usuario`, `<div class="form-grid"><label class="span-all">Nombre<input id="ue-nom" value="${esc(u.nombre || '')}"></label>
      <label>Rol<select id="ue-rol"><option value="trabajador" ${u.rol === 'trabajador' ? 'selected' : ''}>Trabajador</option><option value="admin" ${u.rol === 'admin' ? 'selected' : ''}>Administrador</option></select></label>
      <label class="chk"><input type="checkbox" id="ue-act" ${u.activo ? 'checked' : ''}> Puede ingresar al sistema</label></div>`,
      { footer: `<button class="btn" data-close>Cancelar</button><button class="btn primary" id="ue-ok">Guardar cambios</button>` });
    $('#ue-ok', m).onclick = e => run(async () => {
      chk(await sb.from('perfiles').update({ nombre: $('#ue-nom', m).value.trim(), rol: $('#ue-rol', m).value, activo: $('#ue-act', m).checked }).eq('id', u.id));
      toast('Usuario actualizado'); m.remove(); load();
    }, e.currentTarget);
  };
  $('#u-new').onclick = () => {
    const m = modal('Crear usuario', `<div class="form-grid"><label class="span-all">Nombre completo<input id="un-nom"></label><label class="span-all">Correo<input type="email" id="un-em"></label>
      <label>Contraseña inicial<input type="text" id="un-pw" minlength="6" placeholder="Mínimo 6 caracteres"></label>
      <label>Rol<select id="un-rol"><option value="trabajador">Trabajador</option><option value="admin">Administrador</option></select></label>
      <p class="muted span-all"><small>Entrega la contraseña al usuario; podrá cambiarla con "Cambiar contraseña". Si en Supabase está activa la confirmación por correo, deberá confirmarlo antes de ingresar.</small></p></div>`,
      { footer: `<button class="btn" data-close>Cancelar</button><button class="btn primary" id="un-ok">Crear usuario</button>` });
    $('#un-ok', m).onclick = e => run(async () => {
      const nombre = $('#un-nom', m).value.trim(), email = $('#un-em', m).value.trim(), password = $('#un-pw', m).value;
      if (!nombre || !email) throw new Error('El nombre y el correo son obligatorios');
      if (password.length < 6) throw new Error('La contraseña debe tener al menos 6 caracteres');
      // Cliente temporal: crea la cuenta sin cerrar la sesión del administrador
      const tmp = window.supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_ANON_KEY, { auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false, storageKey: 'sc-tmp-signup' } });
      const data = chk(await tmp.auth.signUp({ email, password, options: { data: { nombre } } }));
      if (!data.user || (data.user.identities && data.user.identities.length === 0)) throw new Error('Ese correo ya está registrado');
      chk(await sb.from('perfiles').update({ nombre, rol: $('#un-rol', m).value, activo: true }).eq('id', data.user.id));
      toast(`Usuario ${email} creado`); m.remove(); load();
    }, e.currentTarget);
  };
  $('#u-list').onclick = e => {
    const b = e.target.closest('[data-act]'); if (!b) return;
    const u = rows.find(x => x.id === b.dataset.id);
    if (b.dataset.act === 'edit') edit(u);
    if (b.dataset.act === 'reset') run(async () => { chk(await sb.auth.resetPasswordForEmail(u.email, { redirectTo: location.href.split('#')[0] })); toast(`Enlace enviado a ${u.email}`); });
  };
  await load();
}

/* ---------------------------------------------------------------------
   ADMINISTRACIÓN · EMPRESA Y CONSECUTIVOS
   --------------------------------------------------------------------- */
async function vEmpresa() {
  const E = chk(await sb.from('empresa').select('*').eq('id', 1).single());
  const cons = chk(await sb.from('consecutivos').select('*').order('tipo'));
  const f = (k, l, cls = '') => `<label class="${cls}">${l}<input id="em-${k}" value="${esc(E[k] || '')}"></label>`;
  $('#main').innerHTML = pageHead('Empresa y consecutivos', 'Estos datos aparecen en el encabezado de todos los documentos impresos') +
    `<div class="card"><div class="form-grid">${f('nombre', 'Razón social', 'span2')}${f('nit', 'NIT')}${f('telefono', 'Teléfono')}${f('direccion', 'Dirección', 'span2')}${f('ciudad', 'Ciudad')}${f('email', 'Correo')}
      <label class="span-all">Texto de resolución (se imprime en facturas)<textarea id="em-resolucion" rows="2">${esc(E.resolucion || '')}</textarea></label>
      <label class="span-all">Pie de página de documentos<textarea id="em-pie_documento" rows="2">${esc(E.pie_documento || '')}</textarea></label></div>
      <div class="form-actions"><button class="btn primary" id="em-ok">Guardar datos de la empresa</button></div></div>
    <div class="card"><div class="card-h"><h3>Consecutivos de documentos</h3></div>
      ${table([{ t: 'Documento', f: c => DT[c.tipo]?.n || c.tipo }, { t: 'Prefijo', f: c => `<input data-pre="${c.tipo}" value="${esc(c.prefijo)}" style="width:110px">` },
      { t: 'Siguiente número', f: c => `<input type="number" min="1" data-sig="${c.tipo}" value="${c.siguiente}" style="width:120px">` }, { t: 'Próximo documento', f: c => `<b>${esc(c.prefijo + c.siguiente)}</b>` }], cons)}
      <div class="form-actions"><button class="btn primary" id="co-ok">Guardar consecutivos</button></div></div>`;
  $('#em-ok').onclick = e => run(async () => {
    const data = {}; ['nombre', 'nit', 'telefono', 'direccion', 'ciudad', 'email', 'resolucion', 'pie_documento'].forEach(k => data[k] = $('#em-' + k).value.trim() || null);
    if (!data.nombre) throw new Error('La razón social es obligatoria');
    S.empresa = chk(await sb.from('empresa').update(data).eq('id', 1).select().single());
    toast('Datos de la empresa guardados'); renderShell(); go('empresa'); route();
  }, e.currentTarget);
  $('#co-ok').onclick = e => run(async () => {
    for (const c of cons) {
      const prefijo = $(`[data-pre="${c.tipo}"]`).value.trim(), siguiente = parseInt($(`[data-sig="${c.tipo}"]`).value);
      if (!(siguiente > 0)) throw new Error(`Número inválido para ${DT[c.tipo]?.n}`);
      if (prefijo !== c.prefijo || siguiente !== +c.siguiente) chk(await sb.from('consecutivos').update({ prefijo, siguiente }).eq('tipo', c.tipo));
    }
    toast('Consecutivos guardados'); vEmpresa();
  }, e.currentTarget);
}

/* ---------------------------------------------------------------------
   INFORMES EN EXCEL
   --------------------------------------------------------------------- */
async function itemsDe(tipo, desde, hasta, bod) {
  return fetchAll(() => {
    let q = sb.from('documento_items').select('*, productos(codigo,nombre,categoria,unidad), documentos!inner(id,numero,fecha,tipo,estado,bodega_id,forma_pago,cliente_id,proveedor_id,referencia,clientes(nombre,documento),proveedores(nombre,documento))')
      .eq('documentos.tipo', tipo).neq('documentos.estado', 'ANULADA').gte('documentos.fecha', desde).lte('documentos.fecha', hasta).order('id');
    if (bod) q = q.eq('documentos.bodega_id', bod);
    return q;
  });
}
async function docsDe(tipo, desde, hasta, bod, conAnuladas = false) {
  return fetchAll(() => {
    let q = sb.from('documentos').select('*, clientes(nombre,documento), proveedores(nombre,documento), bodega:bodegas!documentos_bodega_id_fkey(nombre), perfiles(nombre)').eq('tipo', tipo).gte('fecha', desde).lte('fecha', hasta).order('fecha');
    if (!conAnuladas) q = q.neq('estado', 'ANULADA');
    if (bod) q = q.eq('bodega_id', bod);
    return q;
  });
}
const agrupar = (rows, keyFn, init, add) => { const m = {}; rows.forEach(r => { const k = keyFn(r); add(m[k] || (m[k] = init(r)), r); }); return Object.values(m); };

const REPS = [
  {
    id: 'ventas', t: 'Ventas detalladas', d: 'Facturas del periodo con su detalle por producto. Los administradores ven además costo y utilidad.', f: async ({ desde, hasta, bod }) => {
      const [docs, items] = await Promise.all([docsDe('FAC', desde, hasta, bod, true), itemsDe('FAC', desde, hasta, bod)]);
      excel(`Ventas_${desde}_a_${hasta}`, [
        { name: 'Facturas', rows: docs.map(r => ({ Factura: r.numero, Fecha: fd(r.fecha), Cliente: r.clientes?.nombre, Documento: r.clientes?.documento, Bodega: r.bodega?.nombre || '', 'Forma de pago': r.forma_pago, Subtotal: +r.subtotal, Descuento: +r.descuento, IVA: +r.iva, Total: +r.total, Saldo: +r.saldo, Estado: r.estado, Vendedor: r.perfiles?.nombre || '' })) },
        {
          name: 'Detalle', rows: items.map(i => {
            const row = { Factura: i.documentos.numero, Fecha: fd(i.documentos.fecha), Cliente: i.documentos.clientes?.nombre, Código: i.productos?.codigo, Producto: i.productos?.nombre, Categoría: i.productos?.categoria || '', Cantidad: +i.cantidad, 'Precio unitario': +i.precio, 'Desc %': +i.descuento_pct, 'Venta neta': +i.subtotal, IVA: +i.iva_valor, Total: +i.total };
            if (isAdmin()) { const c = r2((+i.costo_unitario || 0) * i.cantidad); Object.assign(row, { 'Costo unitario': r2(i.costo_unitario), 'Costo total': c, Utilidad: r2(i.subtotal - c), 'Margen %': +i.subtotal ? r2((i.subtotal - c) / i.subtotal * 100) : 0 }); }
            return row;
          })
        },
      ]);
    }
  },
  {
    id: 'vprod', t: 'Ventas por producto', d: 'Cantidades y valores vendidos por producto en el periodo, ordenados de mayor a menor.', f: async ({ desde, hasta, bod }) => {
      const items = await itemsDe('FAC', desde, hasta, bod);
      const rows = agrupar(items, i => i.producto_id, i => ({ Código: i.productos?.codigo, Producto: i.productos?.nombre, Categoría: i.productos?.categoria || '', Cantidad: 0, 'Venta neta': 0, IVA: 0, Total: 0, ...(isAdmin() ? { 'Costo total': 0, Utilidad: 0 } : {}) }),
        (a, i) => { a.Cantidad += +i.cantidad; a['Venta neta'] += +i.subtotal; a.IVA += +i.iva_valor; a.Total += +i.total; if (isAdmin()) { const c = (+i.costo_unitario || 0) * i.cantidad; a['Costo total'] += c; a.Utilidad += i.subtotal - c; } });
      rows.forEach(r => Object.keys(r).forEach(k => typeof r[k] === 'number' && (r[k] = r2(r[k]))));
      excel(`Ventas_por_producto_${desde}_a_${hasta}`, [{ name: 'Por producto', rows: rows.sort((a, b) => b['Venta neta'] - a['Venta neta']) }]);
    }
  },
  {
    id: 'vcli', t: 'Ventas por cliente', d: 'Total facturado, número de facturas y saldo pendiente de cada cliente en el periodo.', f: async ({ desde, hasta, bod }) => {
      const docs = await docsDe('FAC', desde, hasta, bod);
      const rows = agrupar(docs, r => r.cliente_id, r => ({ Cliente: r.clientes?.nombre, Documento: r.clientes?.documento, Facturas: 0, Subtotal: 0, IVA: 0, Total: 0, 'Saldo pendiente': 0 }),
        (a, r) => { a.Facturas++; a.Subtotal += +r.subtotal - +r.descuento; a.IVA += +r.iva; a.Total += +r.total; a['Saldo pendiente'] += +r.saldo; });
      excel(`Ventas_por_cliente_${desde}_a_${hasta}`, [{ name: 'Por cliente', rows: rows.sort((a, b) => b.Total - a.Total) }]);
    }
  },
  {
    id: 'inv', t: 'Inventario valorizado', d: 'Existencias actuales a costo promedio: detalle por bodega, consolidado por producto y resumen por bodega.', admin: false, f: async ({ bod }) => {
      const all = await cargarExistencias(), rows = all.filter(x => !bod || x.bodega_id == bod);
      const det = rows.filter(x => +x.cantidad !== 0).map(x => ({ Bodega: x.bodegas.nombre, Código: x.productos.codigo, Producto: x.productos.nombre, Categoría: x.productos.categoria || '', Unidad: x.productos.unidad, Cantidad: +x.cantidad, 'Costo promedio': r2(x.costo_promedio), 'Valor total': r2(x.cantidad * x.costo_promedio) }));
      const cons = consolidar(rows).filter(x => +x.cantidad !== 0).map(x => ({ Código: x.productos.codigo, Producto: x.productos.nombre, Categoría: x.productos.categoria || '', Cantidad: +x.cantidad, 'Costo promedio ponderado': r2(x.costo_promedio), 'Valor total': r2(x.valor), 'Stock mínimo': +x.productos.stock_minimo, 'Bajo mínimo': +x.productos.stock_minimo > 0 && x.cantidad < x.productos.stock_minimo ? 'Sí' : 'No' }));
      const res = agrupar(rows, x => x.bodega_id, x => ({ Bodega: x.bodegas.nombre, Referencias: 0, 'Valor total': 0 }), (a, x) => { if (+x.cantidad) a.Referencias++; a['Valor total'] = r2(a['Valor total'] + x.cantidad * x.costo_promedio); });
      excel('Inventario_valorizado', [{ name: 'Por bodega', rows: det }, { name: 'Consolidado', rows: cons }, { name: 'Resumen bodegas', rows: res }]);
    }
  },
  { id: 'cartera', t: 'Cartera por edades', d: 'Saldo de cada cliente distribuido por antigüedad de vencimiento, más el detalle por factura.', f: async () => excel('Cartera_por_edades', carteraHojas(await cargarCartera())) },
  {
    id: 'clientes', t: 'Clientes', d: 'Listado completo de clientes con condiciones de crédito y saldo actual en cartera.', f: async () => {
      const [cli, cart] = await Promise.all([fetchAll(() => sb.from('clientes').select('*').order('nombre')), cargarCartera()]);
      const sal = {}; cart.forEach(r => sal[r.cliente_id] = (sal[r.cliente_id] || 0) + +r.saldo);
      excel('Clientes', [{ name: 'Clientes', rows: cli.map(c => ({ 'Tipo doc': c.tipo_doc, Documento: c.documento, Nombre: c.nombre, Teléfono: c.telefono || '', Correo: c.email || '', Dirección: c.direccion || '', Ciudad: c.ciudad || '', 'Días crédito': c.dias_credito, 'Cupo crédito': +c.cupo_credito, 'Saldo cartera': sal[c.id] || 0, Activo: c.activo ? 'Sí' : 'No' })) }]);
    }
  },
  {
    id: 'compras', t: 'Compras detalladas', d: 'Compras del periodo con el detalle por producto y el costo de entrada al inventario.', f: async ({ desde, hasta, bod }) => {
      const [docs, items] = await Promise.all([docsDe('COM', desde, hasta, bod, true), itemsDe('COM', desde, hasta, bod)]);
      excel(`Compras_${desde}_a_${hasta}`, [
        { name: 'Compras', rows: docs.map(r => ({ Compra: r.numero, Fecha: fd(r.fecha), Proveedor: r.proveedores?.nombre, NIT: r.proveedores?.documento, 'Factura proveedor': r.referencia || '', Bodega: r.bodega?.nombre || '', Subtotal: +r.subtotal, Descuento: +r.descuento, IVA: +r.iva, Total: +r.total, Estado: r.estado })) },
        { name: 'Detalle', rows: items.map(i => ({ Compra: i.documentos.numero, Fecha: fd(i.documentos.fecha), Proveedor: i.documentos.proveedores?.nombre, Código: i.productos?.codigo, Producto: i.productos?.nombre, Cantidad: +i.cantidad, 'Costo unitario': +i.precio, 'Desc %': +i.descuento_pct, 'Costo neto': +i.subtotal, IVA: +i.iva_valor, Total: +i.total })) },
      ]);
    }
  },
  {
    id: 'ocpend', t: 'Órdenes de compra pendientes', d: 'Productos pedidos a proveedores que aún no han llegado, con la cantidad pendiente.', f: async ({ desde, hasta }) => {
      const items = (await fetchAll(() => sb.from('documento_items').select('*, productos(codigo,nombre), documentos!inner(numero,fecha,fecha_vencimiento,estado,tipo,proveedores(nombre))').eq('documentos.tipo', 'OC').in('documentos.estado', ['PENDIENTE', 'PARCIAL']).gte('documentos.fecha', desde).lte('documentos.fecha', hasta))).filter(i => +i.cantidad_recibida < +i.cantidad);
      excel('Ordenes_pendientes', [{ name: 'Pendientes', rows: items.map(i => ({ Orden: i.documentos.numero, Fecha: fd(i.documentos.fecha), 'Entrega esperada': fd(i.documentos.fecha_vencimiento), Proveedor: i.documentos.proveedores?.nombre, Código: i.productos?.codigo, Producto: i.productos?.nombre, Pedido: +i.cantidad, Recibido: +i.cantidad_recibida, Pendiente: r3(i.cantidad - i.cantidad_recibida), 'Costo unitario': +i.precio })) }]);
    }
  },
  {
    id: 'rempend', t: 'Remisiones por facturar', d: 'Mercancía despachada a clientes que todavía no se ha facturado.', f: async () => {
      const docs = await fetchAll(() => sb.from('documentos').select('*, clientes(nombre,documento), bodega:bodegas!documentos_bodega_id_fkey(nombre)').eq('tipo', 'REM').eq('estado', 'PENDIENTE').order('fecha'));
      excel('Remisiones_por_facturar', [{ name: 'Remisiones', rows: docs.map(r => ({ Remisión: r.numero, Fecha: fd(r.fecha), 'Días sin facturar': diasEntre(fd(r.fecha), hoy()), Cliente: r.clientes?.nombre, Documento: r.clientes?.documento, Bodega: r.bodega?.nombre, Total: +r.total })) }]);
    }
  },
  {
    id: 'entregas', t: 'Seguimiento de entregas', d: 'Remisiones y facturas del periodo con su estado de entrega, fecha y quién recibió.', f: async ({ desde, hasta, bod }) => {
      const rows = await cargarPedidos(desde, hasta, bod);
      excel(`Entregas_${desde}_a_${hasta}`, [{ name: 'Entregas', rows: rows.map(pedidoExcel) }]);
    }
  },
  {
    id: 'cot', t: 'Cotizaciones', d: 'Cotizaciones del periodo con su estado: abiertas, aprobadas, rechazadas o convertidas en venta.', f: async ({ desde, hasta }) => {
      const docs = await docsDe('COT', desde, hasta, null, true);
      excel(`Cotizaciones_${desde}_a_${hasta}`, [{ name: 'Cotizaciones', rows: docs.map(r => ({ Cotización: r.numero, Fecha: fd(r.fecha), 'Válida hasta': fd(r.fecha_vencimiento), Cliente: r.clientes?.nombre, Total: +r.total, Estado: r.estado, Asesor: r.perfiles?.nombre || '' })) }]);
    }
  },
  {
    id: 'kardex', t: 'Kardex del periodo', d: 'Todos los movimientos de inventario del periodo con saldos y costo promedio después de cada uno.', f: async ({ desde, hasta, bod }) => {
      const rows = await fetchAll(() => { let q = sb.from('kardex').select('*, productos(codigo,nombre), bodegas(nombre), perfiles(nombre)').gte('fecha', dIni(desde)).lte('fecha', dFin(hasta)).order('id'); if (bod) q = q.eq('bodega_id', bod); return q; });
      excel(`Kardex_${desde}_a_${hasta}`, [{ name: 'Kardex', rows: kardexRows(rows) }]);
    }
  },
];

async function vInformes() {
  $('#main').innerHTML = pageHead('Informes en Excel', 'Elige el periodo y, si aplica, la bodega. Cada informe se descarga como archivo .xlsx') +
    `<div class="card filters"><label>Desde<input type="date" id="i-desde" value="${inicioMes()}"></label><label>Hasta<input type="date" id="i-hasta" value="${hoy()}"></label>
      <label>Bodega<select id="i-bod"><option value="">Todas</option>${opts(S.bodegas, '')}</select></label></div>
    <div class="rep-grid">${REPS.map(r => `<div class="card rep"><h3>${r.t}</h3><p class="muted">${r.d}</p><button class="btn primary" data-rep="${r.id}">Descargar Excel</button></div>`).join('')}</div>`;
  $('.rep-grid').onclick = e => {
    const b = e.target.closest('[data-rep]'); if (!b) return;
    const R = REPS.find(r => r.id === b.dataset.rep);
    const ctx = { desde: $('#i-desde').value, hasta: $('#i-hasta').value, bod: $('#i-bod').value || null };
    if (ctx.desde > ctx.hasta) return toast('La fecha inicial es posterior a la final', 'err');
    const txt = b.textContent; b.textContent = 'Generando…';
    run(() => R.f(ctx), b).finally(() => { b.textContent = txt; });
  };
}
