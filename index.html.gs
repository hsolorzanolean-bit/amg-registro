// ============================================================
// AMG de Colombia SAS — Sistema de Registro de Turno
// Google Apps Script — sirve el HTML Y guarda los datos
// ============================================================

const SHEET_NAME = 'Form Responses 1';

// ── Sirve la página web ──────────────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutput(getHTML())
    .setTitle('AMG de Colombia SAS — Registro de turno')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0');
}

// ── Recibe los datos del formulario ─────────────────────────
function doPost(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data  = JSON.parse(e.postData.contents);

    // Encabezados si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Marca de tiempo',
        'Nombre del empleado',
        'Vehículo',
        'Tipo de registro',
        'Aprobador',
        'Motivo / Punto de recogida',
        'GPS Latitud',
        'GPS Longitud',
        'Precisión GPS (m)',
        'GPS Verificado',
        'Ver en Google Maps'
      ]);
      // Formato de encabezados
      const hdr = sheet.getRange(1, 1, 1, 11);
      hdr.setBackground('#0A3D47');
      hdr.setFontColor('#FFFFFF');
      hdr.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const now      = new Date();
    const mapsLink = (data.geoLat && data.geoLng)
      ? `https://maps.google.com/?q=${data.geoLat},${data.geoLng}`
      : 'No disponible';

    sheet.appendRow([
      now,
      data.nombre    || '',
      data.vehiculo  || '',
      data.tipo      || '',
      data.aprobador || '',
      data.motivo    || data.recogida || '',
      data.geoLat    || '',
      data.geoLng    || '',
      data.geoAcc    || '',
      data.geoLat ? 'Sí' : 'No',
      mapsLink
    ]);

    // Correos
    const tiempoStr = Utilities.formatDate(now, 'America/Bogota', 'dd/MM/yyyy HH:mm');
    const todosAprobadores = 'alugo@amgdecolombia.com,pmartinez@amgdecolombia.com,gestionhumana@amgdecolombia.com';
    const gestion          = 'gestionhumana@amgdecolombia.com';

    if (data.tipo === 'Tiempo extra' || data.tipo === 'Solicitud de recogida') {
      GmailApp.sendEmail(
        todosAprobadores,
        `[Aprobación requerida] ${data.tipo} — ${data.nombre}`,
        `Se ha recibido una solicitud que requiere aprobación.\n\nEmpleado: ${data.nombre}\nTipo: ${data.tipo}\nVehículo: ${data.vehiculo}\nAprobador seleccionado: ${data.aprobador}\nDetalle: ${data.motivo || data.recogida || 'N/A'}\nHora: ${tiempoStr}\nUbicación GPS: ${mapsLink}\n\n— Sistema de registro AMG de Colombia SAS`
      );
    } else {
      GmailApp.sendEmail(
        gestion,
        `[Registro de turno] ${data.tipo} — ${data.nombre}`,
        `Nuevo registro de turno recibido.\n\nEmpleado: ${data.nombre}\nTipo: ${data.tipo}\nVehículo: ${data.vehiculo}\nHora: ${tiempoStr}\nUbicación GPS: ${mapsLink}\n\n— Sistema de registro AMG de Colombia SAS`
      );
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── HTML del formulario ──────────────────────────────────────
function getHTML() {
  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
<title>AMG de Colombia SAS — Registro de turno</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:opsz,wght@9..40,400;9..40,500;9..40,600;9..40,700&family=DM+Mono:wght@400;500&display=swap');
:root{
  --teal-dark:#0A3D47;--teal:#156B6B;--teal-mid:#1D9E9E;
  --green:#D1FAE5;--green-dark:#065F46;--green-border:#6EE7B7;
  --red:#FEF2F2;--red-dark:#991B1B;--red-border:#FECACA;
  --amber:#FFFBEB;--amber-dark:#92400E;--amber-border:#FCD34D;
  --blue:#EFF6FF;--blue-dark:#1E40AF;--blue-border:#BFDBFE;
  --gray-50:#F8FAFC;--gray-100:#F1F5F9;--gray-200:#E2E8F0;
  --gray-400:#94A3B8;--gray-500:#64748B;--gray-800:#1E293B;--gray-900:#0F172A;
  --white:#FFFFFF;
}
*{box-sizing:border-box;margin:0;padding:0;-webkit-tap-highlight-color:transparent;}
body{font-family:'DM Sans',sans-serif;background:var(--gray-50);min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:1rem;}
.card{background:var(--white);border-radius:24px;box-shadow:0 8px 32px rgba(0,0,0,0.12);width:100%;max-width:420px;overflow:hidden;}
.header{background:linear-gradient(145deg,var(--teal-dark) 0%,var(--teal) 60%,var(--teal-mid) 100%);padding:1.5rem;position:relative;overflow:hidden;}
.header::before{content:'';position:absolute;top:-40px;right:-40px;width:120px;height:120px;background:rgba(255,255,255,0.05);border-radius:50%;}
.header-inner{position:relative;z-index:1;text-align:center;}
.header-badge{display:inline-block;background:rgba(255,255,255,0.12);border:1px solid rgba(255,255,255,0.2);border-radius:20px;padding:2px 10px;font-size:10px;color:rgba(255,255,255,0.8);margin-bottom:8px;letter-spacing:0.3px;}
.header h1{color:white;font-size:14px;font-weight:600;letter-spacing:0.4px;margin-bottom:4px;}
.header-clock{color:rgba(255,255,255,0.6);font-size:11px;font-family:'DM Mono',monospace;letter-spacing:0.5px;}
.body{padding:1.25rem;}
.geo-box{border-radius:14px;padding:10px 13px;margin-bottom:1rem;display:flex;align-items:center;gap:11px;border:1.5px solid;transition:all 0.5s ease;}
.geo-loading{background:#F0F9FF;border-color:#BAE6FD;}
.geo-ok{background:var(--green);border-color:var(--green-border);}
.geo-fail{background:var(--red);border-color:var(--red-border);}
.geo-icon{width:38px;height:38px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:17px;flex-shrink:0;}
.geo-loading .geo-icon{background:#E0F2FE;}
.geo-ok .geo-icon{background:#A7F3D0;}
.geo-fail .geo-icon{background:#FECACA;}
.geo-info h4{font-size:12px;font-weight:600;color:var(--gray-800);}
.geo-info p{font-size:10px;color:var(--gray-500);margin-top:1px;line-height:1.4;}
.geo-coords{font-family:'DM Mono',monospace;font-size:9px;color:var(--green-dark);background:#A7F3D0;padding:2px 7px;border-radius:5px;display:inline-block;margin-top:3px;}
.spinner{width:17px;height:17px;border:2px solid #BAE6FD;border-top-color:#0EA5E9;border-radius:50%;animation:spin 0.7s linear infinite;}
@keyframes spin{to{transform:rotate(360deg);}}
.fg{margin-bottom:12px;}
.fg label{font-size:11px;font-weight:600;color:var(--gray-500);display:block;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.4px;}
.sw{position:relative;}
.sw::after{content:'▾';position:absolute;right:12px;top:50%;transform:translateY(-50%);color:var(--gray-400);pointer-events:none;font-size:11px;}
select,input[type=text],textarea{width:100%;padding:10px 13px;border:1.5px solid var(--gray-200);border-radius:11px;font-family:'DM Sans',sans-serif;font-size:14px;color:var(--gray-800);background:var(--white);outline:none;transition:border-color 0.2s,box-shadow 0.2s;-webkit-appearance:none;appearance:none;}
select:focus,input:focus,textarea:focus{border-color:var(--teal-mid);box-shadow:0 0 0 3px rgba(29,158,158,0.1);}
textarea{resize:none;height:66px;font-size:13px;}
.tipo-label{font-size:11px;font-weight:600;color:var(--gray-500);margin-bottom:6px;text-transform:uppercase;letter-spacing:0.4px;}
.tipo-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-bottom:6px;}
.tbtn{padding:10px 4px;border-radius:11px;border:1.5px solid var(--gray-200);background:var(--white);font-family:'DM Sans',sans-serif;font-size:11px;font-weight:500;cursor:pointer;text-align:center;color:var(--gray-500);transition:all 0.15s;line-height:1.35;}
.tbtn.s-in{background:var(--green);color:var(--green-dark);border-color:var(--green-border);font-weight:600;}
.tbtn.s-out{background:var(--red);color:var(--red-dark);border-color:var(--red-border);font-weight:600;}
.tbtn.s-ot{background:var(--amber);color:var(--amber-dark);border-color:var(--amber-border);font-weight:600;}
.tbtn-pk{width:100%;padding:10px;border-radius:11px;border:1.5px solid var(--gray-200);background:var(--white);font-family:'DM Sans',sans-serif;font-size:11px;font-weight:500;cursor:pointer;color:var(--gray-500);transition:all 0.15s;display:none;margin-bottom:6px;}
.tbtn-pk.s-pk{background:var(--blue);color:var(--blue-dark);border-color:var(--blue-border);font-weight:600;}
.extra{display:none;}.extra.on{display:block;}
.notice{background:var(--blue);border:1px solid var(--blue-border);border-radius:10px;padding:9px 11px;font-size:11px;color:var(--blue-dark);line-height:1.55;margin-top:8px;margin-bottom:10px;}
.sbtn{width:100%;padding:13px;border-radius:13px;border:none;font-family:'DM Sans',sans-serif;font-size:15px;font-weight:700;cursor:pointer;margin-top:8px;display:none;transition:all 0.2s;letter-spacing:0.2px;}
.sbtn:active{transform:scale(0.98);}
.sbtn:disabled{opacity:0.55;cursor:not-allowed;transform:none;}
.sbtn.c-in{background:linear-gradient(135deg,#065F46,#059669);color:white;box-shadow:0 4px 14px rgba(5,150,105,0.3);}
.sbtn.c-out{background:linear-gradient(135deg,#991B1B,#DC2626);color:white;box-shadow:0 4px 14px rgba(220,38,38,0.3);}
.sbtn.c-ot{background:linear-gradient(135deg,#92400E,#D97706);color:white;box-shadow:0 4px 14px rgba(217,119,6,0.3);}
.sbtn.c-pk{background:linear-gradient(135deg,#1E40AF,#3B82F6);color:white;box-shadow:0 4px 14px rgba(59,130,246,0.3);}
.footer{font-size:10px;color:var(--gray-400);text-align:center;margin-top:14px;font-family:'DM Mono',monospace;}
.overlay{display:none;position:fixed;inset:0;background:rgba(15,23,42,0.65);align-items:center;justify-content:center;z-index:99;padding:1.5rem;backdrop-filter:blur(4px);}
.overlay.on{display:flex;}
.ov-card{background:white;border-radius:24px;padding:2rem 1.5rem;text-align:center;max-width:320px;width:100%;box-shadow:0 24px 60px rgba(0,0,0,0.2);animation:pop 0.35s cubic-bezier(.34,1.56,.64,1);}
@keyframes pop{from{transform:scale(0.7);opacity:0;}to{transform:scale(1);opacity:1;}}
.ov-icon{font-size:54px;margin-bottom:10px;}
.ov-title{font-size:18px;font-weight:700;color:var(--gray-900);margin-bottom:6px;}
.ov-sub{font-size:13px;color:var(--gray-500);line-height:1.55;}
.ov-time{font-family:'DM Mono',monospace;font-size:10px;color:var(--teal);margin-top:10px;}
.ov-geo{font-size:10px;color:var(--gray-400);margin-top:4px;}
.ov-new{margin-top:16px;padding:10px 20px;border-radius:10px;border:1.5px solid var(--gray-200);background:var(--gray-50);font-family:'DM Sans',sans-serif;font-size:13px;font-weight:500;cursor:pointer;color:#334155;}
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <div class="header-inner">
      <div class="header-badge">AMG DE COLOMBIA SAS</div>
      <h1>Registro de turno y vehículo</h1>
      <div class="header-clock" id="clock">--/--/---- --:--</div>
    </div>
  </div>
  <div class="body">
    <div class="geo-box geo-loading" id="geo-box">
      <div class="geo-icon" id="geo-icon"><div class="spinner"></div></div>
      <div class="geo-info">
        <h4 id="geo-title">Verificando ubicación...</h4>
        <p id="geo-desc">Acepta el permiso de GPS para validar el registro.</p>
      </div>
    </div>
    <div class="fg">
      <label>Vehículo asignado</label>
      <div class="sw">
        <select id="f-vehicle">
          <option value="">Selecciona...</option>
          <option>Nissan Frontier — IKT-256</option>
          <option>Nissan Frontier — LOQ-838</option>
          <option value="Sin vehículo asignado">Sin vehículo asignado</option>
        </select>
      </div>
    </div>
    <div class="fg">
      <label>Nombre del empleado</label>
      <input type="text" id="f-name" placeholder="Nombre completo" autocomplete="name">
    </div>
    <div class="fg">
      <div class="tipo-label">¿Qué necesitas registrar?</div>
      <div class="tipo-grid">
        <button class="tbtn" id="t-in"  onclick="setTipo('in')">Entrada<br>a turno</button>
        <button class="tbtn" id="t-out" onclick="setTipo('out')">Salida<br>de turno</button>
        <button class="tbtn" id="t-ot"  onclick="setTipo('ot')">Tiempo<br>extra</button>
      </div>
      <button class="tbtn-pk" id="t-pk" onclick="setTipo('pk')">Solicitar recogida — requiere aprobador</button>
    </div>
    <div class="extra" id="x-ot">
      <div class="fg">
        <label>Aprobador</label>
        <div class="sw">
          <select id="f-ap-ot">
            <option value="">Selecciona aprobador...</option>
            <option>Yury Trujillo</option>
            <option>Adriana Lugo</option>
            <option>Paola Martinez</option>
          </select>
        </div>
      </div>
      <div class="fg">
        <label>Motivo del tiempo extra</label>
        <textarea id="f-motivo" placeholder="Describe brevemente..."></textarea>
      </div>
    </div>
    <div class="extra" id="x-pk">
      <div class="fg">
        <label>Aprobador</label>
        <div class="sw">
          <select id="f-ap-pk">
            <option value="">Selecciona aprobador...</option>
            <option>Yury Trujillo</option>
            <option>Adriana Lugo</option>
            <option>Paola Martinez</option>
          </select>
        </div>
      </div>
      <div class="fg">
        <label>Punto de recogida</label>
        <input type="text" id="f-lugar" placeholder="Ej: Entrada principal obra Norte">
      </div>
      <div class="notice">Esta solicitud <strong>no registra entrada a turno.</strong> El turno inicia cuando llegues físicamente y hagas tu registro de entrada.</div>
    </div>
    <button class="sbtn" id="sbtn" onclick="enviar()">Registrar</button>
  </div>
</div>
<div class="footer">AMG de Colombia SAS · Registro con verificación GPS · 100% Google</div>
<div class="overlay" id="overlay">
  <div class="ov-card">
    <div class="ov-icon"  id="ov-icon">✅</div>
    <div class="ov-title" id="ov-title">Registro enviado</div>
    <div class="ov-sub"   id="ov-sub"></div>
    <div class="ov-time"  id="ov-time"></div>
    <div class="ov-geo"   id="ov-geo"></div>
    <button class="ov-new" onclick="nuevoRegistro()">Hacer otro registro</button>
  </div>
</div>
<script>
const SCRIPT_URL = ScriptApp ? ScriptApp.getService().getUrl() : window.location.href.split('?')[0];
let gLat=null,gLng=null,gAcc=null,tipo=null;
function pad(n){return String(n).padStart(2,'0');}
function ahora(){const d=new Date();return pad(d.getDate())+'/'+pad(d.getMonth()+1)+'/'+d.getFullYear()+' '+pad(d.getHours())+':'+pad(d.getMinutes());}
setInterval(()=>document.getElementById('clock').textContent=ahora(),1000);
document.getElementById('clock').textContent=ahora();
(function getGeo(){
  const box=document.getElementById('geo-box'),icon=document.getElementById('geo-icon'),title=document.getElementById('geo-title'),desc=document.getElementById('geo-desc');
  if(!navigator.geolocation){box.className='geo-box geo-fail';icon.textContent='⚠️';title.textContent='GPS no disponible';desc.textContent='El administrador verá que no se pudo verificar.';return;}
  navigator.geolocation.getCurrentPosition(pos=>{
    gLat=pos.coords.latitude.toFixed(6);gLng=pos.coords.longitude.toFixed(6);gAcc=Math.round(pos.coords.accuracy);
    box.className='geo-box geo-ok';icon.textContent='📍';title.textContent='Ubicación verificada ✓';
    desc.innerHTML='GPS activo y registrado.<br><span class="geo-coords">'+gLat+', '+gLng+' (±'+gAcc+'m)</span>';
  },()=>{box.className='geo-box geo-fail';icon.textContent='⚠️';title.textContent='Ubicación no disponible';desc.textContent='Activa el GPS. Quedará marcado sin verificación.';},{enableHighAccuracy:true,timeout:12000});
})();
document.getElementById('f-vehicle').addEventListener('change',function(){
  document.getElementById('t-pk').style.display=this.value==='Sin vehículo asignado'?'block':'none';
  if(tipo==='pk'&&this.value!=='Sin vehículo asignado')setTipo(null);
});
function setTipo(t){
  tipo=t;
  ['in','out','ot'].forEach(x=>document.getElementById('t-'+x).className='tbtn');
  document.getElementById('t-pk').className='tbtn-pk';
  ['x-ot','x-pk'].forEach(x=>document.getElementById(x).classList.remove('on'));
  const sb=document.getElementById('sbtn');
  if(!t){sb.style.display='none';return;}
  const cfg={in:{sel:'s-in',bc:'c-in',txt:'✓  Registrar entrada a turno'},out:{sel:'s-out',bc:'c-out',txt:'✓  Registrar salida de turno'},ot:{sel:'s-ot',bc:'c-ot',txt:'Enviar solicitud de tiempo extra'},pk:{sel:'s-pk',bc:'c-pk',txt:'Enviar solicitud de recogida'}}[t];
  if(t==='pk')document.getElementById('t-pk').classList.add('s-pk');
  else document.getElementById('t-'+t).classList.add(cfg.sel);
  if(t==='ot')document.getElementById('x-ot').classList.add('on');
  if(t==='pk')document.getElementById('x-pk').classList.add('on');
  sb.className='sbtn '+cfg.bc;sb.textContent=cfg.txt;sb.style.display='block';
}
async function enviar(){
  const nombre=document.getElementById('f-name').value.trim();
  const vehiculo=document.getElementById('f-vehicle').value;
  if(!nombre){alert('Por favor ingresa tu nombre completo.');return;}
  if(!vehiculo){alert('Por favor selecciona el vehículo.');return;}
  if(!tipo)return;
  if(tipo==='ot'&&!document.getElementById('f-ap-ot').value){alert('El tiempo extra requiere un aprobador.');return;}
  if(tipo==='pk'&&!document.getElementById('f-ap-pk').value){alert('La recogida requiere un aprobador.');return;}
  const tipoLabel={in:'Entrada a turno',out:'Salida de turno',ot:'Tiempo extra',pk:'Solicitud de recogida'}[tipo];
  const payload={nombre,vehiculo,tipo:tipoLabel,aprobador:tipo==='ot'?document.getElementById('f-ap-ot').value:tipo==='pk'?document.getElementById('f-ap-pk').value:'',motivo:tipo==='ot'?document.getElementById('f-motivo').value.trim():'',recogida:tipo==='pk'?document.getElementById('f-lugar').value.trim():'',geoLat:gLat,geoLng:gLng,geoAcc:gAcc};
  const sb=document.getElementById('sbtn');sb.disabled=true;sb.textContent='Enviando...';
  try{await fetch(window.location.href.split('?')[0],{method:'POST',body:JSON.stringify(payload)});}catch(e){}
  const iconM={in:'✅',out:'🔴',ot:'⏱️',pk:'🚗'};
  const subM={in:'Tu entrada a turno quedó registrada con GPS.',out:'Tu salida de turno quedó registrada con GPS.',ot:'Solicitud enviada. Pendiente de aprobación por '+payload.aprobador+'.',pk:'Solicitud enviada a '+payload.aprobador+'. Recuerda registrar tu entrada cuando llegues.'};
  document.getElementById('ov-icon').textContent=iconM[tipo];document.getElementById('ov-title').textContent=tipoLabel;document.getElementById('ov-sub').textContent=subM[tipo];document.getElementById('ov-time').textContent=ahora();document.getElementById('ov-geo').textContent=gLat?'GPS: '+gLat+', '+gLng:'Sin GPS — marcado para el administrador.';
  document.getElementById('overlay').classList.add('on');
}
function nuevoRegistro(){
  document.getElementById('overlay').classList.remove('on');
  ['f-name','f-motivo','f-lugar'].forEach(id=>document.getElementById(id).value='');
  ['f-vehicle','f-ap-ot','f-ap-pk'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('t-pk').style.display='none';
  const sb=document.getElementById('sbtn');sb.disabled=false;
  setTipo(null);
}
<\/script>
</body>
</html>`;
}
