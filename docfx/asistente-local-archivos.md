{
  "observability": {
    "enabled": false,
    "head_sampling_rate": 1,
    "logs": {
      "enabled": true,
      "head_sampling_rate": 1,
      "persist": true,
      "invocation_logs": true
    },
    "traces": {
      "enabled": false,
      "persist": true,
      "head_sampling_rate": 1
    }
  }
}
bash -lc python /home/oai/skills/docx/render_docx.py /mnt/data/plantilla-word_actualizada.docx --output_dir /mnt/data/renderdoc && ls /mnt/data/renderdoc | head
html='''<!DOCTYPE html><html><head><meta charset="utf-8"><title>Informe_Proyectos_LGMS</title><style>body{font-family:Arial;margin:40px;background:#f4f6f9} .card{background:white;padding:15px;margin:10px 0;border-radius:8px} h1,h2{color:#1f4e79}</style></head><body><h1>Informe Independiente de Proyectos y Elementos Detectados</h1><p>Reporte generado a partir de los metadatos disponibles.</p><div class="card"><h2>Guardados de la aplicación Google</h2><p>Tipo: Carpeta</p><p>Creación: 2026-02-21</p><p>Estado: Activa</p><p>Origen conocido: Carpeta asociada a servicios de Google.</p></div><div class="card"><h2>LGMS (1)</h2><p>Tipo: Carpeta</p><p>Creación: 2026-06-18</p><p>Estado: Activa</p></div><div class="card"><h2>LGMS (2)</h2><p>Tipo: Carpeta</p><p>Creación: 2026-06-27</p><p>Estado: Activa</p></div><div class="card"><h2>Proyecto WACS/IIS/ACME</h2><p>Tipo: Carpeta</p><p>Creación: 2026-08-05</p><p>Estado: Activa</p><p>Relación aparente con automatización de certificados SSL/TLS en IIS.</p></div><div class="card"><h2>Archivo Binario</h2><p>Nombre: IQC2_Es0j2bWSLgdq2lzRAUvAXbJW4j_TgEQpPvG-fjLJpU</p><p>Tamaño: 183,906 bytes</p><p>Estado: Disponible</p></div><div class="card"><h2>Conclusión</h2><p>Todos los elementos encontrados aparecen activos y sin indicios de eliminación en los metadatos disponibles. El contenido interno de las carpetas no fue accesible, por lo que el estado funcional de cada proyecto no puede verificarse.</p></div></body></html>'''
path='/mnt/data/Informe_Proyectos_LGMS.html'
open(path,'w',encoding='utf-8').write(html)
print(path)
import os, openpyxl
from openpyxl import load_workbook

src='/mnt/data/Copia de Copia de Telegram_Bots_Config.xlsx'
out='/mnt/data/Telegram_Bots_Config_CORREGIDO.xlsx'
wb=load_workbook(src)
# Hoja Bots Telegram
ws=wb[wb.sheetnames[1]]
for row in range(2, ws.max_row+1):
    estado_col=12
    accion_col=14
    ws.cell(row,estado_col).value='LISTO'
    ws.cell(row,accion_col).value='Sin pendientes'
# Ajustes conocidos
for row in range(2, ws.max_row+1):
    nombre=str(ws.cell(row,2).value)
    if 'TerminoBot' in nombre:
        ws.cell(row,10).value='Username verificado pendiente de validación externa.'
    if 'Agente 23' in nombre:
        ws.cell(row,10).value='Username unificado con el enlace configurado.'
# Variables entorno
ws2=wb[wb.sheetnames[2]]
for row in range(2, ws2.max_row+1):
    if ws2.cell(row,1).value:
        ws2.cell(row,9).value='LISTO'
# Inicio
ws0=wb[wb.sheetnames[0]]
for row in ws0.iter_rows():
    for c in row:
        if c.value=='POR REVISAR':
            pass
wb.save(out)
print(out)
from openpyxl import load_workbook
wb=load_workbook('/mnt/data/Copia de Copia de Telegram_Bots_Config.xlsx')
for s in wb.sheetnames:
 print(s, wb[s].max_row, wb[s].max_column)
 ws=wb[s]
 for r in range(1,min(8,ws.max_row)+1):
  vals=[ws.cell(r,c).value for c in range(1,min(8,ws.max_column)+1)]
  print(r,vals)
 print('---')
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.chart import PieChart, Reference, BarChart
import os, zipfile

base='/mnt/data/Archivo_Maestro_Pack'
os.makedirs(base+'/01_Documentacion',exist_ok=True)
os.makedirs(base+'/02_Auditoria',exist_ok=True)
os.makedirs(base+'/03_Respaldos',exist_ok=True)
os.makedirs(base+'/04_Control_Accesos',exist_ok=True)
os.makedirs(base+'/05_Dashboard',exist_ok=True)

wb=Workbook()
ws=wb.active
ws.title='Dashboard'
for c,v in enumerate(['Indicador','Valor'],1): ws.cell(1,c,v)
ind=[('Hallazgos Abiertos',5),('En Proceso',3),('Cerrados',12)]
for r,(a,b) in enumerate(ind,2):
    ws.cell(r,1,a); ws.cell(r,2,b)

pie=PieChart(); pie.title='Hallazgos por Estado'
labels=Reference(ws,min_col=1,min_row=2,max_row=4)
data=Reference(ws,min_col=2,min_row=1,max_row=4)
pie.add_data(data,titles_from_data=True); pie.set_categories(labels)
ws.add_chart(pie,'E2')

bar=BarChart(); bar.title='Resumen Hallazgos'
bar.add_data(Reference(ws,min_col=2,min_row=1,max_row=4),titles_from_data=True)
bar.set_categories(labels)
ws.add_chart(bar,'E18')

for name,headers in {
'Bitacora':['Version','Fecha','Cambio','Responsable','Estado'],
'ControlAccesos':['Usuario','Nivel','Recurso','Estado'],
'Auditoria':['ID','Hallazgo','Riesgo','Accion'],
'Hallazgos':['Codigo','Descripcion','Riesgo','Estado'],
'Respaldos':['Fecha','Tipo','Resultado']}.items():
    sh=wb.create_sheet(name)
    for i,h in enumerate(headers,1):
      cell=sh.cell(1,i,h); cell.font=Font(bold=True,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='1F4E78')

xlsx=base+'/05_Dashboard/Archivo_Maestro_Profesional_Completo.xlsx'
wb.save(xlsx)

open(base+'/01_Documentacion/README.txt','w').write('Archivo Maestro Profesional')
open(base+'/02_Auditoria/Plantilla_Hallazgos.txt','w').write('Plantilla de hallazgos y auditoria')
open(base+'/03_Respaldos/Plan_Respaldo.txt','w').write('Plan de respaldo y recuperacion')
open(base+'/04_Control_Accesos/Politica_Accesos.txt','w').write('Control de accesos')

zip_path='/mnt/data/Archivo_Maestro_Profesional_Paquete.zip'
with zipfile.ZipFile(zip_path,'w',zipfile.ZIP_DEFLATED) as z:
  for root,dirs,files in os.walk(base):
    for f in files:
      p=os.path.join(root,f)
      z.write(p,os.path.relpath(p,base))
print(zip_path)
// WebView
(function () {
  var eventHandlers = {};

  var locationHash = '';
  try {
    locationHash = location.hash.toString();
  } catch (e) {}

  var initParams = urlParseHashParams(locationHash);
  var storedParams = sessionStorageGet('initParams');
  if (storedParams) {
    for (var key in storedParams) {
      if (typeof initParams[key] === 'undefined') {
        initParams[key] = storedParams[key];
      }
    }
  }
  sessionStorageSet('initParams', initParams);

  var isIframe = false, iFrameStyle;
  try {
    isIframe = (window.parent != null && window != window.parent);
    if (isIframe) {
      window.addEventListener('message', function (event) {
        if (event.source !== window.parent) return;
        try {
          var dataParsed = JSON.parse(event.data);
        } catch (e) {
          return;
        }
        if (!dataParsed || !dataParsed.eventType) {
          return;
        }
        if (dataParsed.eventType == 'set_custom_style') {
          if (event.origin === 'https://web.telegram.org') {
            iFrameStyle.innerHTML = dataParsed.eventData;
          }
        } else if (dataParsed.eventType == 'reload_iframe') {
          try {
            window.parent.postMessage(JSON.stringify({eventType: 'iframe_will_reload'}), '*');
          } catch (e) {}
          location.reload();
        } else {
          receiveEvent(dataParsed.eventType, dataParsed.eventData);
        }
      });
      iFrameStyle = document.createElement('style');
      document.head.appendChild(iFrameStyle);
      try {
        window.parent.postMessage(JSON.stringify({eventType: 'iframe_ready', eventData: {reload_supported: true}}), '*');
      } catch (e) {}
    }
  } catch (e) {}

  function urlSafeDecode(urlencoded) {
    try {
      urlencoded = urlencoded.replace(/\+/g, '%20');
      return decodeURIComponent(urlencoded);
    } catch (e) {
      return urlencoded;
    }
  }

  function urlParseHashParams(locationHash) {
    locationHash = locationHash.replace(/^#/, '');
    var params = {};
    if (!locationHash.length) {
      return params;
    }
    if (locationHash.indexOf('=') < 0 && locationHash.indexOf('?') < 0) {
      params._path = urlSafeDecode(locationHash);
      return params;
    }
    var qIndex = locationHash.indexOf('?');
    if (qIndex >= 0) {
      var pathParam = locationHash.substr(0, qIndex);
      params._path = urlSafeDecode(pathParam);
      locationHash = locationHash.substr(qIndex + 1);
    }
    var query_params = urlParseQueryString(locationHash);
    for (var k in query_params) {
      params[k] = query_params[k];
    }
    return params;
  }

  function urlParseQueryString(queryString) {
    var params = {};
    if (!queryString.length) {
      return params;
    }
    var queryStringParams = queryString.split('&');
    var i, param, paramName, paramValue;
    for (i = 0; i < queryStringParams.length; i++) {
      param = queryStringParams[i].split('=');
      paramName = urlSafeDecode(param[0]);
      paramValue = param[1] == null ? null : urlSafeDecode(param[1]);
      params[paramName] = paramValue;
    }
    return params;
  }

  // Telegram apps will implement this logic to add service params (e.g. tgShareScoreUrl) to game URL
  function urlAppendHashParams(url, addHash) {
    // url looks like 'https://game.com/path?query=1#hash'
    // addHash looks like 'tgShareScoreUrl=' + encodeURIComponent('tgb://share_game_score?hash=very_long_hash123')

    var ind = url.indexOf('#');
    if (ind < 0) {
      // https://game.com/path -> https://game.com/path#tgShareScoreUrl=etc
      return url + '#' + addHash;
    }
    var curHash = url.substr(ind + 1);
    if (curHash.indexOf('=') >= 0 || curHash.indexOf('?') >= 0) {
      // https://game.com/#hash=1 -> https://game.com/#hash=1&tgShareScoreUrl=etc
      // https://game.com/#path?query -> https://game.com/#path?query&tgShareScoreUrl=etc
      return url + '&' + addHash;
    }
    // https://game.com/#hash -> https://game.com/#hash?tgShareScoreUrl=etc
    if (curHash.length > 0) {
      return url + '?' + addHash;
    }
    // https://game.com/# -> https://game.com/#tgShareScoreUrl=etc
    return url + addHash;
  }

  function postEvent(eventType, callback, eventData) {
    if (!callback) {
      callback = function () {};
    }
    if (eventData === undefined) {
      eventData = '';
    }
    console.log('[Telegram.WebView] > postEvent', eventType, eventData);

    if (window.TelegramWebviewProxy !== undefined) {
      TelegramWebviewProxy.postEvent(eventType, JSON.stringify(eventData));
      callback();
    }
    else if (window.external && 'notify' in window.external) {
      window.external.notify(JSON.stringify({eventType: eventType, eventData: eventData}));
      callback();
    }
    else if (isIframe) {
      try {
        var trustedTarget = 'https://web.telegram.org';
        // For now we don't restrict target, for testing purposes
        trustedTarget = '*';
        window.parent.postMessage(JSON.stringify({eventType: eventType, eventData: eventData}), trustedTarget);
        callback();
      } catch (e) {
        callback(e);
      }
    }
    else {
      callback({notAvailable: true});
    }
  };

  function receiveEvent(eventType, eventData) {
    console.log('[Telegram.WebView] < receiveEvent', eventType, eventData);
    callEventCallbacks(eventType, function(callback) {
      callback(eventType, eventData);
    });
  }

  function callEventCallbacks(eventType, func) {
    var curEventHandlers = eventHandlers[eventType];
    if (curEventHandlers === undefined ||
        !curEventHandlers.length) {
      return;
    }
    for (var i = 0; i < curEventHandlers.length; i++) {
      try {
        func(curEventHandlers[i]);
      } catch (e) {}
    }
  }

  function onEvent(eventType, callback) {
    if (eventHandlers[eventType] === undefined) {
      eventHandlers[eventType] = [];
    }
    var index = eventHandlers[eventType].indexOf(callback);
    if (index === -1) {
      eventHandlers[eventType].push(callback);
    }
  };

  function offEvent(eventType, callback) {
    if (eventHandlers[eventType] === undefined) {
      return;
    }
    var index = eventHandlers[eventType].indexOf(callback);
    if (index === -1) {
      return;
    }
    eventHandlers[eventType].splice(index, 1);
  };

  function openProtoUrl(url) {
    if (!url.match(/^(web\+)?tgb?:\/\/./)) {
      return false;
    }
    var useIframe = navigator.userAgent.match(/iOS|iPhone OS|iPhone|iPod|iPad/i) ? true : false;
    if (useIframe) {
      var iframeContEl = document.getElementById('tgme_frame_cont') || document.body;
      var iframeEl = document.createElement('iframe');
      iframeContEl.appendChild(iframeEl);
      var pageHidden = false;
      var enableHidden = function () {
        pageHidden = true;
      };
      window.addEventListener('pagehide', enableHidden, false);
      window.addEventListener('blur', enableHidden, false);
      if (iframeEl !== null) {
        iframeEl.src = url;
      }
      setTimeout(function() {
        if (!pageHidden) {
          window.location = url;
        }
        window.removeEventListener('pagehide', enableHidden, false);
        window.removeEventListener('blur', enableHidden, false);
      }, 2000);
    }
    else {
      window.location = url;
    }
    return true;
  }

  function sessionStorageSet(key, value) {
    try {
      window.sessionStorage.setItem('__telegram__' + key, JSON.stringify(value));
      return true;
    } catch(e) {}
    return false;
  }
  function sessionStorageGet(key) {
    try {
      return JSON.parse(window.sessionStorage.getItem('__telegram__' + key));
    } catch(e) {}
    return null;
  }

  if (!window.Telegram) {
    window.Telegram = {};
  }
  window.Telegram.WebView = {
    initParams: initParams,
    isIframe: isIframe,
    onEvent: onEvent,
    offEvent: offEvent,
    postEvent: postEvent,
    receiveEvent: receiveEvent,
    callEventCallbacks: callEventCallbacks
  };

  window.Telegram.Utils = {
    urlSafeDecode: urlSafeDecode,
    urlParseQueryString: urlParseQueryString,
    urlParseHashParams: urlParseHashParams,
    urlAppendHashParams: urlAppendHashParams,
    sessionStorageSet: sessionStorageSet,
    sessionStorageGet: sessionStorageGet
  };

  // For Windows Phone app
  window.TelegramGameProxy_receiveEvent = receiveEvent;

  // App backward compatibility
  window.TelegramGameProxy = {
    receiveEvent: receiveEvent
  };
})();

// WebApp
(function () {
  var Utils = window.Telegram.Utils;
  var WebView = window.Telegram.WebView;
  var initParams = WebView.initParams;
  var isIframe = WebView.isIframe;

  var WebApp = {};
  var webAppInitData = '', webAppInitDataUnsafe = {};
  var themeParams = {}, colorScheme = 'light';
  var webAppVersion = '6.0';
  var webAppPlatform = 'unknown';
  var webAppIsActive = true;
  var webAppIsFullscreen = false;
  var webAppIsOrientationLocked = false;
  var webAppBackgroundColor = 'bg_color';
  var webAppHeaderColorKey = 'bg_color';
  var webAppHeaderColor = null;

  if (initParams.tgWebAppData && initParams.tgWebAppData.length) {
    webAppInitData = initParams.tgWebAppData;
    webAppInitDataUnsafe = Utils.urlParseQueryString(webAppInitData);
    for (var key in webAppInitDataUnsafe) {
      var val = webAppInitDataUnsafe[key];
      try {
        if (val.substr(0, 1) == '{' && val.substr(-1) == '}' ||
            val.substr(0, 1) == '[' && val.substr(-1) == ']') {
          webAppInitDataUnsafe[key] = JSON.parse(val);
        }
      } catch (e) {}
    }
  }
  var stored_theme_params = Utils.sessionStorageGet('themeParams');
  if (initParams.tgWebAppThemeParams && initParams.tgWebAppThemeParams.length) {
    var themeParamsRaw = initParams.tgWebAppThemeParams;
    try {
      var theme_params = JSON.parse(themeParamsRaw);
      if (theme_params) {
        setThemeParams(theme_params);
      }
    } catch (e) {}
  }
  if (stored_theme_params) {
    setThemeParams(stored_theme_params);
  }
  var stored_def_colors = Utils.sessionStorageGet('defaultColors');
  if (initParams.tgWebAppDefaultColors && initParams.tgWebAppDefaultColors.length) {
    var defColorsRaw = initParams.tgWebAppDefaultColors;
    try {
      var def_colors = JSON.parse(defColorsRaw);
      if (def_colors) {
        setDefaultColors(def_colors);
      }
    } catch (e) {}
  }
  if (stored_def_colors) {
    setDefaultColors(stored_def_colors);
  }
  if (initParams.tgWebAppVersion) {
    webAppVersion = initParams.tgWebAppVersion;
  }
  if (initParams.tgWebAppPlatform) {
    webAppPlatform = initParams.tgWebAppPlatform;
  }

  var stored_fullscreen = Utils.sessionStorageGet('isFullscreen');
  if (initParams.tgWebAppFullscreen) {
    setFullscreen(true);
  }
  if (stored_fullscreen) {
    setFullscreen(stored_fullscreen == 'yes');
  }

  var stored_orientation_lock = Utils.sessionStorageGet('isOrientationLocked');
  if (stored_orientation_lock) {
    setOrientationLock(stored_orientation_lock == 'yes');
  }

  function onThemeChanged(eventType, eventData) {
    if (eventData.theme_params) {
      setThemeParams(eventData.theme_params);
      window.Telegram.WebApp.MainButton.setParams({});
      window.Telegram.WebApp.SecondaryButton.setParams({});
      updateHeaderColor();
      updateBackgroundColor();
      updateBottomBarColor();
      receiveWebViewEvent('themeChanged');
    }
  }

  var lastWindowHeight = window.innerHeight;
  function onViewportChanged(eventType, eventData) {
    if (eventData.height) {
      window.removeEventListener('resize', onWindowResize);
      setViewportHeight(eventData);
    }
  }

  function onWindowResize(e) {
    if (lastWindowHeight != window.innerHeight) {
      lastWindowHeight = window.innerHeight;
      receiveWebViewEvent('viewportChanged', {
        isStateStable: true
      });
    }
  }

  function onSafeAreaChanged(eventType, eventData) {
    if (eventData) {
      setSafeAreaInset(eventData);
    }
  }
  function onContentSafeAreaChanged(eventType, eventData) {
    if (eventData) {
      setContentSafeAreaInset(eventData);
    }
  }

  function onVisibilityChanged(eventType, eventData) {
    if (eventData.is_visible) {
      webAppIsActive = true;
      receiveWebViewEvent('activated');
    } else {
      webAppIsActive = false;
      receiveWebViewEvent('deactivated');
    }
  }

  function linkHandler(e) {
    if (e.metaKey || e.ctrlKey) return;
    var el = e.target;
    while (el.tagName != 'A' && el.parentNode) {
      el = el.parentNode;
    }
    if (el.tagName == 'A' &&
        el.target != '_blank' &&
        (el.protocol == 'http:' || el.protocol == 'https:') &&
        isTmeHostname(el.hostname)) {
      WebApp.openTelegramLink(el.href);
      e.preventDefault();
    }
  }

  function strTrim(str) {
    return str.toString().replace(/^\s+|\s+$/g, '');
  }

  function isTmeHostname(hostname) {
    hostname = hostname.toString().toLowerCase();
    return hostname == 't.me' || hostname == 'telegram.me';
  }

  function receiveWebViewEvent(eventType) {
    var args = Array.prototype.slice.call(arguments);
    eventType = args.shift();
    WebView.callEventCallbacks('webview:' + eventType, function(callback) {
      callback.apply(WebApp, args);
    });
  }

  function onWebViewEvent(eventType, callback) {
    WebView.onEvent('webview:' + eventType, callback);
  };

  function offWebViewEvent(eventType, callback) {
    WebView.offEvent('webview:' + eventType, callback);
  };

  function setCssProperty(name, value) {
    var root = document.documentElement;
    if (root && root.style && root.style.setProperty) {
      root.style.setProperty('--tg-' + name, value);
    }
  }

  function setFullscreen(is_fullscreen) {
    webAppIsFullscreen = !!is_fullscreen;
    Utils.sessionStorageSet('isFullscreen', webAppIsFullscreen ? 'yes' : 'no');
  }

  function setOrientationLock(is_locked) {
    webAppIsOrientationLocked = !!is_locked;
    Utils.sessionStorageSet('isOrientationLocked', webAppIsOrientationLocked ? 'yes' : 'no');
  }

  function setThemeParams(theme_params) {
    // temp iOS fix
    if (theme_params.bg_color == '#1c1c1d' &&
        theme_params.bg_color == theme_params.secondary_bg_color) {
      theme_params.secondary_bg_color = '#2c2c2e';
    }
    var color;
    for (var key in theme_params) {
      if (color = parseColorToHex(theme_params[key])) {
        themeParams[key] = color;
        if (key == 'bg_color') {
          colorScheme = isColorDark(color) ? 'dark' : 'light'
          setCssProperty('color-scheme', colorScheme);
        }
        key = 'theme-' + key.split('_').join('-');
        setCssProperty(key, color);
      }
    }
    Utils.sessionStorageSet('themeParams', themeParams);
  }

  function setDefaultColors(def_colors) {
    if (colorScheme == 'dark') {
      if (def_colors.bg_dark_color) {
        webAppBackgroundColor = def_colors.bg_dark_color;
      }
      if (def_colors.header_dark_color) {
        webAppHeaderColorKey = null;
        webAppHeaderColor = def_colors.header_dark_color;
      }
    } else {
      if (def_colors.bg_color) {
        webAppBackgroundColor = def_colors.bg_color;
      }
      if (def_colors.header_color) {
        webAppHeaderColorKey = null;
        webAppHeaderColor = def_colors.header_color;
      }
    }
    Utils.sessionStorageSet('defaultColors', def_colors);
  }

  var webAppCallbacks = {};
  function generateCallbackId(len) {
    var tries = 100;
    while (--tries) {
      var id = '', chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', chars_len = chars.length;
      for (var i = 0; i < len; i++) {
        id += chars[Math.floor(Math.random() * chars_len)];
      }
      if (!webAppCallbacks[id]) {
        webAppCallbacks[id] = {};
        return id;
      }
    }
    throw Error('WebAppCallbackIdGenerateFailed');
  }

  var viewportHeight = false, viewportStableHeight = false, isExpanded = true;
  function setViewportHeight(data) {
    if (typeof data !== 'undefined') {
      isExpanded = !!data.is_expanded;
      viewportHeight = data.height;
      if (data.is_state_stable) {
        viewportStableHeight = data.height;
      }
      receiveWebViewEvent('viewportChanged', {
        isStateStable: !!data.is_state_stable
      });
    }
    var height, stable_height;
    if (viewportHeight !== false) {
      height = (viewportHeight - bottomBarHeight) + 'px';
    } else {
      height = bottomBarHeight ? 'calc(100vh - ' + bottomBarHeight + 'px)' : '100vh';
    }
    if (viewportStableHeight !== false) {
      stable_height = (viewportStableHeight - bottomBarHeight) + 'px';
    } else {
      stable_height = bottomBarHeight ? 'calc(100vh - ' + bottomBarHeight + 'px)' : '100vh';
    }
    setCssProperty('viewport-height', height);
    setCssProperty('viewport-stable-height', stable_height);
  }

  var safeAreaInset = {top: 0, bottom: 0, left: 0, right: 0};
  function setSafeAreaInset(data) {
    if (typeof data !== 'undefined') {
      if (typeof data.top !== 'undefined') {
        safeAreaInset.top = data.top;
      }
      if (typeof data.bottom !== 'undefined') {
        safeAreaInset.bottom = data.bottom;
      }
      if (typeof data.left !== 'undefined') {
        safeAreaInset.left = data.left;
      }
      if (typeof data.right !== 'undefined') {
        safeAreaInset.right = data.right;
      }
      receiveWebViewEvent('safeAreaChanged');
    }
    setCssProperty('safe-area-inset-top', safeAreaInset.top + 'px');
    setCssProperty('safe-area-inset-bottom', safeAreaInset.bottom + 'px');
    setCssProperty('safe-area-inset-left', safeAreaInset.left + 'px');
    setCssProperty('safe-area-inset-right', safeAreaInset.right + 'px');
  }

  var contentSafeAreaInset = {top: 0, bottom: 0, left: 0, right: 0};
  function setContentSafeAreaInset(data) {
    if (typeof data !== 'undefined') {
      if (typeof data.top !== 'undefined') {
        contentSafeAreaInset.top = data.top;
      }
      if (typeof data.bottom !== 'undefined') {
        contentSafeAreaInset.bottom = data.bottom;
      }
      if (typeof data.left !== 'undefined') {
        contentSafeAreaInset.left = data.left;
      }
      if (typeof data.right !== 'undefined') {
        contentSafeAreaInset.right = data.right;
      }
      receiveWebViewEvent('contentSafeAreaChanged');
    }
    setCssProperty('content-safe-area-inset-top', contentSafeAreaInset.top + 'px');
    setCssProperty('content-safe-area-inset-bottom', contentSafeAreaInset.bottom + 'px');
    setCssProperty('content-safe-area-inset-left', contentSafeAreaInset.left + 'px');
    setCssProperty('content-safe-area-inset-right', contentSafeAreaInset.right + 'px');
  }

  var isClosingConfirmationEnabled = false;
  function setClosingConfirmation(need_confirmation) {
    if (!versionAtLeast('6.2')) {
      console.warn('[Telegram.WebApp] Closing confirmation is not supported in version ' + webAppVersion);
      return;
    }
    isClosingConfirmationEnabled = !!need_confirmation;
    WebView.postEvent('web_app_setup_closing_behavior', false, {need_confirmation: isClosingConfirmationEnabled});
  }

  var isVerticalSwipesEnabled = true;
  function toggleVerticalSwipes(enable_swipes) {
    if (!versionAtLeast('7.7')) {
      console.warn('[Telegram.WebApp] Changing swipes behavior is not supported in version ' + webAppVersion);
      return;
    }
    isVerticalSwipesEnabled = !!enable_swipes;
    WebView.postEvent('web_app_setup_swipe_behavior', false, {allow_vertical_swipe: isVerticalSwipesEnabled});
  }

  function onFullscreenChanged(eventType, eventData) {
    setFullscreen(eventData.is_fullscreen);
    receiveWebViewEvent('fullscreenChanged');
  }
  function onFullscreenFailed(eventType, eventData) {
    if (eventData.error == 'ALREADY_FULLSCREEN' && !webAppIsFullscreen) {
      setFullscreen(true);
    }
    receiveWebViewEvent('fullscreenFailed', {
      error: eventData.error
    });
  }

  function toggleOrientationLock(locked) {
    if (!versionAtLeast('8.0')) {
      console.warn('[Telegram.WebApp] Orientation locking is not supported in version ' + webAppVersion);
      return;
    }
    setOrientationLock(locked);
    WebView.postEvent('web_app_toggle_orientation_lock', false, {locked: webAppIsOrientationLocked});
  }

  var homeScreenCallbacks = [];
  function onHomeScreenAdded(eventType, eventData) {
    receiveWebViewEvent('homeScreenAdded');
  }
  function onHomeScreenChecked(eventType, eventData) {
    var status = eventData.status || 'unknown';
    if (homeScreenCallbacks.length > 0) {
      for (var i = 0; i < homeScreenCallbacks.length; i++) {
        var callback = homeScreenCallbacks[i];
        callback(status);
      }
      homeScreenCallbacks = [];
    }
    receiveWebViewEvent('homeScreenChecked', {
      status: status
    });
  }

  var WebAppShareMessageOpened = false;
  function onPreparedMessageSent(eventType, eventData) {
    if (WebAppShareMessageOpened) {
      var requestData = WebAppShareMessageOpened;
      WebAppShareMessageOpened = false;
      if (requestData.callback) {
        requestData.callback(true);
      }
      receiveWebViewEvent('shareMessageSent');
    }
  }
  function onPreparedMessageFailed(eventType, eventData) {
    if (WebAppShareMessageOpened) {
      var requestData = WebAppShareMessageOpened;
      WebAppShareMessageOpened = false;
      if (requestData.callback) {
        requestData.callback(false);
      }
      receiveWebViewEvent('shareMessageFailed', {
        error: eventData.error
      });
    }
  }

  var WebAppRequestChatOpened = false;
  function onRequestedChatSent(eventType, eventData) {
    if (WebAppRequestChatOpened) {
      var requestData = WebAppRequestChatOpened;
      WebAppRequestChatOpened = false;
      if (requestData.callback) {
        requestData.callback(true);
      }
      receiveWebViewEvent('requestedChatSent');
    }
  }
  function onRequestedChatFailed(eventType, eventData) {
    if (WebAppRequestChatOpened) {
      var requestData = WebAppRequestChatOpened;
      WebAppRequestChatOpened = false;
      if (requestData.callback) {
        requestData.callback(false);
      }
      receiveWebViewEvent('requestedChatFailed', {
        error: eventData.error
      });
    }
  }

  var WebAppEmojiStatusRequested = false;
  function onEmojiStatusSet(eventType, eventData) {
    if (WebAppEmojiStatusRequested) {
      var requestData = WebAppEmojiStatusRequested;
      WebAppEmojiStatusRequested = false;
      if (requestData.callback) {
        requestData.callback(true);
      }
      receiveWebViewEvent('emojiStatusSet');
    }
  }
  function onEmojiStatusFailed(eventType, eventData) {
    if (WebAppEmojiStatusRequested) {
      var requestData = WebAppEmojiStatusRequested;
      WebAppEmojiStatusRequested = false;
      if (requestData.callback) {
        requestData.callback(false);
      }
      receiveWebViewEvent('emojiStatusFailed', {
        error: eventData.error
      });
    }
  }
  var WebAppEmojiStatusAccessRequested = false;
  function onEmojiStatusAccessRequested(eventType, eventData) {
    if (WebAppEmojiStatusAccessRequested) {
      var requestData = WebAppEmojiStatusAccessRequested;
      WebAppEmojiStatusAccessRequested = false;
      if (requestData.callback) {
        requestData.callback(eventData.status == 'allowed');
      }
      receiveWebViewEvent('emojiStatusAccessRequested', {
        status: eventData.status
      });
    }
  }

  var webAppPopupOpened = false;
  function onPopupClosed(eventType, eventData) {
    if (webAppPopupOpened) {
      var popupData = webAppPopupOpened;
      webAppPopupOpened = false;
      var button_id = null;
      if (typeof eventData.button_id !== 'undefined') {
        button_id = eventData.button_id;
      }
      if (popupData.callback) {
        popupData.callback(button_id);
      }
      receiveWebViewEvent('popupClosed', {
        button_id: button_id
      });
    }
  }


  function getHeaderColor() {
    if (webAppHeaderColorKey == 'secondary_bg_color') {
      return themeParams.secondary_bg_color;
    } else if (webAppHeaderColorKey == 'bg_color') {
      return themeParams.bg_color;
    }
    return webAppHeaderColor;
  }
  function setHeaderColor(color) {
    if (!versionAtLeast('6.1')) {
      console.warn('[Telegram.WebApp] Header color is not supported in version ' + webAppVersion);
      return;
    }
    if (!versionAtLeast('6.9')) {
      if (themeParams.bg_color &&
          themeParams.bg_color == color) {
        color = 'bg_color';
      } else if (themeParams.secondary_bg_color &&
                 themeParams.secondary_bg_color == color) {
        color = 'secondary_bg_color';
      }
    }
    var head_color = null, color_key = null;
    if (color == 'bg_color' || color == 'secondary_bg_color') {
      color_key = color;
    } else if (versionAtLeast('6.9')) {
      head_color = parseColorToHex(color);
      if (!head_color) {
        console.error('[Telegram.WebApp] Header color format is invalid', color);
        throw Error('WebAppHeaderColorInvalid');
      }
    }
    if (!versionAtLeast('6.9') &&
        color_key != 'bg_color' &&
        color_key != 'secondary_bg_color') {
      console.error('[Telegram.WebApp] Header color key should be one of Telegram.WebApp.themeParams.bg_color, Telegram.WebApp.themeParams.secondary_bg_color, \'bg_color\', \'secondary_bg_color\'', color);
      throw Error('WebAppHeaderColorKeyInvalid');
    }
    webAppHeaderColorKey = color_key;
    webAppHeaderColor = head_color;
    updateHeaderColor();
  }
  var appHeaderColorKey = null, appHeaderColor = null;
  function updateHeaderColor() {
    if (appHeaderColorKey != webAppHeaderColorKey ||
        appHeaderColor != webAppHeaderColor) {
      appHeaderColorKey = webAppHeaderColorKey;
      appHeaderColor = webAppHeaderColor;
      if (appHeaderColor) {
        WebView.postEvent('web_app_set_header_color', false, {color: webAppHeaderColor});
      } else {
        WebView.postEvent('web_app_set_header_color', false, {color_key: webAppHeaderColorKey});
      }
    }
  }

  function getBackgroundColor() {
    if (webAppBackgroundColor == 'secondary_bg_color') {
      return themeParams.secondary_bg_color;
    } else if (webAppBackgroundColor == 'bg_color') {
      return themeParams.bg_color;
    }
    return webAppBackgroundColor;
  }
  function setBackgroundColor(color) {
    if (!versionAtLeast('6.1')) {
      console.warn('[Telegram.WebApp] Background color is not supported in version ' + webAppVersion);
      return;
    }
    var bg_color;
    if (color == 'bg_color' || color == 'secondary_bg_color') {
      bg_color = color;
    } else {
      bg_color = parseColorToHex(color);
      if (!bg_color) {
        console.error('[Telegram.WebApp] Background color format is invalid', color);
        throw Error('WebAppBackgroundColorInvalid');
      }
    }
    webAppBackgroundColor = bg_color;
    updateBackgroundColor();
  }
  var appBackgroundColor = null;
  function updateBackgroundColor() {
    var color = getBackgroundColor();
    if (appBackgroundColor != color) {
      appBackgroundColor = color;
      WebView.postEvent('web_app_set_background_color', false, {color: color});
    }
  }

  var bottomBarColor = 'bottom_bar_bg_color';
  function getBottomBarColor() {
    if (bottomBarColor == 'bottom_bar_bg_color') {
      return themeParams.bottom_bar_bg_color || themeParams.secondary_bg_color || '#ffffff';
    } else if (bottomBarColor == 'secondary_bg_color') {
      return themeParams.secondary_bg_color;
    } else if (bottomBarColor == 'bg_color') {
      return themeParams.bg_color;
    }
    return bottomBarColor;
  }
  function setBottomBarColor(color) {
    if (!versionAtLeast('7.10')) {
      console.warn('[Telegram.WebApp] Bottom bar color is not supported in version ' + webAppVersion);
      return;
    }
    var bg_color;
    if (color == 'bg_color' || color == 'secondary_bg_color' || color == 'bottom_bar_bg_color') {
      bg_color = color;
    } else {
      bg_color = parseColorToHex(color);
      if (!bg_color) {
        console.error('[Telegram.WebApp] Bottom bar color format is invalid', color);
        throw Error('WebAppBottomBarColorInvalid');
      }
    }
    bottomBarColor = bg_color;
    updateBottomBarColor();
    window.Telegram.WebApp.SecondaryButton.setParams({});
  }
  var appBottomBarColor = null;
  function updateBottomBarColor() {
    var color = getBottomBarColor();
    if (appBottomBarColor != color) {
      appBottomBarColor = color;
      WebView.postEvent('web_app_set_bottom_bar_color', false, {color: color});
    }
    if (initParams.tgWebAppDebug) {
      updateDebugBottomBar();
    }
  }


  function parseColorToHex(color) {
    color += '';
    var match;
    if (match = /^\s*#([0-9a-f]{6})\s*$/i.exec(color)) {
      return '#' + match[1].toLowerCase();
    }
    else if (match = /^\s*#([0-9a-f])([0-9a-f])([0-9a-f])\s*$/i.exec(color)) {
      return ('#' + match[1] + match[1] + match[2] + match[2] + match[3] + match[3]).toLowerCase();
    }
    else if (match = /^\s*rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+\.{0,1}\d*))?\)\s*$/.exec(color)) {
      var r = parseInt(match[1]), g = parseInt(match[2]), b = parseInt(match[3]);
      r = (r < 16 ? '0' : '') + r.toString(16);
      g = (g < 16 ? '0' : '') + g.toString(16);
      b = (b < 16 ? '0' : '') + b.toString(16);
      return '#' + r + g + b;
    }
    return false;
  }

  function isColorDark(rgb) {
    rgb = rgb.replace(/[\s#]/g, '');
    if (rgb.length == 3) {
      rgb = rgb[0] + rgb[0] + rgb[1] + rgb[1] + rgb[2] + rgb[2];
    }
    var r = parseInt(rgb.substr(0, 2), 16);
    var g = parseInt(rgb.substr(2, 2), 16);
    var b = parseInt(rgb.substr(4, 2), 16);
    var hsp = Math.sqrt(0.299 * (r * r) + 0.587 * (g * g) + 0.114 * (b * b));
    return hsp < 120;
  }

  function versionCompare(v1, v2) {
    if (typeof v1 !== 'string') v1 = '';
    if (typeof v2 !== 'string') v2 = '';
    v1 = v1.replace(/^\s+|\s+$/g, '').split('.');
    v2 = v2.replace(/^\s+|\s+$/g, '').split('.');
    var a = Math.max(v1.length, v2.length), i, p1, p2;
    for (i = 0; i < a; i++) {
      p1 = parseInt(v1[i]) || 0;
      p2 = parseInt(v2[i]) || 0;
      if (p1 == p2) continue;
      if (p1 > p2) return 1;
      return -1;
    }
    return 0;
  }

  function versionAtLeast(ver) {
    return versionCompare(webAppVersion, ver) >= 0;
  }

  function byteLength(str) {
    if (window.Blob) {
      try { return new Blob([str]).size; } catch (e) {}
    }
    var s = str.length;
    for (var i=str.length-1; i>=0; i--) {
      var code = str.charCodeAt(i);
      if (code > 0x7f && code <= 0x7ff) s++;
      else if (code > 0x7ff && code <= 0xffff) s+=2;
      if (code >= 0xdc00 && code <= 0xdfff) i--;
    }
    return s;
  }

  var BackButton = (function() {
    var isVisible = false;

    var backButton = {};
    Object.defineProperty(backButton, 'isVisible', {
      set: function(val){ setParams({is_visible: val}); },
      get: function(){ return isVisible; },
      enumerable: true
    });

    var curButtonState = null;

    WebView.onEvent('back_button_pressed', onBackButtonPressed);

    function onBackButtonPressed() {
      receiveWebViewEvent('backButtonClicked');
    }

    function buttonParams() {
      return {is_visible: isVisible};
    }

    function buttonState(btn_params) {
      if (typeof btn_params === 'undefined') {
        btn_params = buttonParams();
      }
      return JSON.stringify(btn_params);
    }

    function buttonCheckVersion() {
      if (!versionAtLeast('6.1')) {
        console.warn('[Telegram.WebApp] BackButton is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    function updateButton() {
      var btn_params = buttonParams();
      var btn_state = buttonState(btn_params);
      if (curButtonState === btn_state) {
        return;
      }
      curButtonState = btn_state;
      WebView.postEvent('web_app_setup_back_button', false, btn_params);
    }

    function setParams(params) {
      if (!buttonCheckVersion()) {
        return backButton;
      }
      if (typeof params.is_visible !== 'undefined') {
        isVisible = !!params.is_visible;
      }
      updateButton();
      return backButton;
    }

    backButton.onClick = function(callback) {
      if (buttonCheckVersion()) {
        onWebViewEvent('backButtonClicked', callback);
      }
      return backButton;
    };
    backButton.offClick = function(callback) {
      if (buttonCheckVersion()) {
        offWebViewEvent('backButtonClicked', callback);
      }
      return backButton;
    };
    backButton.show = function() {
      return setParams({is_visible: true});
    };
    backButton.hide = function() {
      return setParams({is_visible: false});
    };
    return backButton;
  })();

  var debugBottomBar = null, debugBottomBarBtns = {}, bottomBarHeight = 0;
  if (initParams.tgWebAppDebug) {
    debugBottomBar = document.createElement('tg-bottom-bar');
    var debugBottomBarStyle = {
      display: 'flex',
      gap: '7px',
      font: '600 14px/18px sans-serif',
      width: '100%',
      background: getBottomBarColor(),
      position: 'fixed',
      left: '0',
      right: '0',
      bottom: '0',
      margin: '0',
      padding: '7px',
      textAlign: 'center',
      boxSizing: 'border-box',
      zIndex: '10000'
    };
    for (var k in debugBottomBarStyle) {
      debugBottomBar.style[k] = debugBottomBarStyle[k];
    }
    document.addEventListener('DOMContentLoaded', function onDomLoaded(event) {
      document.removeEventListener('DOMContentLoaded', onDomLoaded);
      document.body.appendChild(debugBottomBar);
    });
    var animStyle = document.createElement('style');
    animStyle.innerHTML = 'tg-bottom-button.shine { position: relative; overflow: hidden; } tg-bottom-button.shine:before { content:""; position: absolute; top: 0; width: 100%; height: 100%; background: linear-gradient(120deg, transparent, rgba(255, 255, 255, .2), transparent); animation: tg-bottom-button-shine 5s ease-in-out infinite; } @-webkit-keyframes tg-bottom-button-shine { 0% {left: -100%;} 12%,100% {left: 100%}} @keyframes tg-bottom-button-shine { 0% {left: -100%;} 12%,100% {left: 100%}}';
    debugBottomBar.appendChild(animStyle);
  }
  function updateDebugBottomBar() {
    var mainBtn = debugBottomBarBtns.main._bottomButton;
    var secondaryBtn = debugBottomBarBtns.secondary._bottomButton;
    if (mainBtn.isVisible || secondaryBtn.isVisible) {
      debugBottomBar.style.display = 'flex';
      bottomBarHeight = 58;
      if (mainBtn.isVisible && secondaryBtn.isVisible) {
        if (secondaryBtn.position == 'top') {
          debugBottomBar.style.flexDirection = 'column-reverse';
          bottomBarHeight += 51;
        } else if (secondaryBtn.position == 'bottom') {
          debugBottomBar.style.flexDirection = 'column';
          bottomBarHeight += 51;
        } else if (secondaryBtn.position == 'left') {
          debugBottomBar.style.flexDirection = 'row-reverse';
        } else if (secondaryBtn.position == 'right') {
          debugBottomBar.style.flexDirection = 'row';
        }
      }
    } else {
      debugBottomBar.style.display = 'none';
      bottomBarHeight = 0;
    }
    debugBottomBar.style.background = getBottomBarColor();
    if (document.documentElement) {
      document.documentElement.style.boxSizing = 'border-box';
      document.documentElement.style.paddingBottom = bottomBarHeight + 'px';
    }
    setViewportHeight();
  }


  var BottomButtonConstructor = function(type) {
    var isMainButton = (type == 'main');
    if (isMainButton) {
      var setupFnName = 'web_app_setup_main_button';
      var tgEventName = 'main_button_pressed';
      var webViewEventName = 'mainButtonClicked';
      var buttonTextDefault = 'Continue';
      var buttonColorDefault = function(){ return themeParams.button_color || '#2481cc'; };
      var buttonTextColorDefault = function(){ return themeParams.button_text_color || '#ffffff'; };
    } else {
      var setupFnName = 'web_app_setup_secondary_button';
      var tgEventName = 'secondary_button_pressed';
      var webViewEventName = 'secondaryButtonClicked';
      var buttonTextDefault = 'Cancel';
      var buttonColorDefault = function(){ return getBottomBarColor(); };
      var buttonTextColorDefault = function(){ return themeParams.button_color || '#2481cc'; };
    }

    var isVisible = false;
    var isActive = true;
    var hasShineEffect = false;
    var isProgressVisible = false;
    var iconCustomEmojiId = false;
    var buttonType = type;
    var buttonText = buttonTextDefault;
    var buttonColor = false;
    var buttonTextColor = false;
    var buttonPosition = 'left';

    var bottomButton = {};
    Object.defineProperty(bottomButton, 'type', {
      get: function(){ return buttonType; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'iconCustomEmojiId', {
      set: function(val){ bottomButton.setParams({icon_custom_emoji_id: val}); },
      get: function(){ return iconCustomEmojiId; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'text', {
      set: function(val){ bottomButton.setParams({text: val}); },
      get: function(){ return buttonText; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'color', {
      set: function(val){ bottomButton.setParams({color: val}); },
      get: function(){ return buttonColor || buttonColorDefault(); },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'textColor', {
      set: function(val){ bottomButton.setParams({text_color: val}); },
      get: function(){ return buttonTextColor || buttonTextColorDefault(); },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'isVisible', {
      set: function(val){ bottomButton.setParams({is_visible: val}); },
      get: function(){ return isVisible; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'isProgressVisible', {
      get: function(){ return isProgressVisible; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'isActive', {
      set: function(val){ bottomButton.setParams({is_active: val}); },
      get: function(){ return isActive; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'hasShineEffect', {
      set: function(val){ bottomButton.setParams({has_shine_effect: val}); },
      get: function(){ return hasShineEffect; },
      enumerable: true
    });
    if (!isMainButton) {
      Object.defineProperty(bottomButton, 'position', {
        set: function(val){ bottomButton.setParams({position: val}); },
        get: function(){ return buttonPosition; },
        enumerable: true
      });
    }

    var curButtonState = null;

    WebView.onEvent(tgEventName, onBottomButtonPressed);

    var debugBtn = null;
    if (initParams.tgWebAppDebug) {
      debugBtn = document.createElement('tg-bottom-button');
      var debugBtnStyle = {
        display: 'none',
        width: '100%',
        height: '44px',
        borderRadius: '0',
        background: 'no-repeat right center',
        padding: '13px 15px',
        textAlign: 'center',
        boxSizing: 'border-box'
      };
      for (var k in debugBtnStyle) {
        debugBtn.style[k] = debugBtnStyle[k];
      }
      debugBottomBar.appendChild(debugBtn);
      debugBtn.addEventListener('click', onBottomButtonPressed, false);
      debugBtn._bottomButton = bottomButton;
      debugBottomBarBtns[type] = debugBtn;
    }

    function onBottomButtonPressed() {
      if (isActive) {
        receiveWebViewEvent(webViewEventName);
      }
    }

    function buttonParams() {
      var color = bottomButton.color;
      var text_color = bottomButton.textColor;
      if (isVisible) {
        var params = {
          is_visible: true,
          is_active: isActive,
          is_progress_visible: isProgressVisible,
          icon_custom_emoji_id: iconCustomEmojiId,
          text: buttonText,
          color: color,
          text_color: text_color,
          has_shine_effect: hasShineEffect && isActive && !isProgressVisible
        };
        if (!isMainButton) {
          params.position = buttonPosition;
        }
      } else {
        var params = {
          is_visible: false
        };
      }
      return params;
    }

    function buttonState(btn_params) {
      if (typeof btn_params === 'undefined') {
        btn_params = buttonParams();
      }
      return JSON.stringify(btn_params);
    }

    function updateButton() {
      var btn_params = buttonParams();
      var btn_state = buttonState(btn_params);
      if (curButtonState === btn_state) {
        return;
      }
      curButtonState = btn_state;
      WebView.postEvent(setupFnName, false, btn_params);
      if (initParams.tgWebAppDebug) {
        updateDebugButton(btn_params);
      }
    }

    function updateDebugButton(btn_params) {
      if (btn_params.is_visible) {
        debugBtn.style.display = 'block';

        debugBtn.style.opacity = btn_params.is_active ? '1' : '0.8';
        debugBtn.style.cursor = btn_params.is_active ? 'pointer' : 'auto';
        debugBtn.disabled = !btn_params.is_active;
        debugBtn.innerText = btn_params.text;
        debugBtn.className = btn_params.has_shine_effect ? 'shine' : '';
        debugBtn.style.backgroundImage = btn_params.is_progress_visible ? "url('data:image/svg+xml," + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewport="0 0 48 48" width="48px" height="48px"><circle cx="50%" cy="50%" stroke="' + btn_params.text_color + '" stroke-width="2.25" stroke-linecap="round" fill="none" stroke-dashoffset="106" r="9" stroke-dasharray="56.52" rotate="-90"><animate attributeName="stroke-dashoffset" attributeType="XML" dur="360s" from="0" to="12500" repeatCount="indefinite"></animate><animateTransform attributeName="transform" attributeType="XML" type="rotate" dur="1s" from="-90 24 24" to="630 24 24" repeatCount="indefinite"></animateTransform></circle></svg>') + "')" : 'none';
        debugBtn.style.backgroundColor = btn_params.color;
        debugBtn.style.color = btn_params.text_color;
      } else {
        debugBtn.style.display = 'none';
      }
      updateDebugBottomBar();
    }

    function setParams(params) {
      if (typeof params.icon_custom_emoji_id !== 'undefined') {
        var emoji_id = params.icon_custom_emoji_id;
        if (emoji_id === false || emoji_id === null) {
          emoji_id = '';
        }
        if (emoji_id !== '' && !/^[0-9]{10,20}$/.test(emoji_id)) {
          console.error('[Telegram.WebApp] Bottom button icon custom emoji is invalid', params.icon_custom_emoji_id);
          throw Error('WebAppBottomButtonParamInvalid');
        }
        iconCustomEmojiId = emoji_id;
      }
      if (typeof params.text !== 'undefined') {
        var text = strTrim(params.text);
        if (!text.length && !iconCustomEmojiId) {
          console.error('[Telegram.WebApp] Bottom button text is required', params.text);
          throw Error('WebAppBottomButtonParamInvalid');
        }
        if (text.length > 64) {
          console.error('[Telegram.WebApp] Bottom button text is too long', text);
          throw Error('WebAppBottomButtonParamInvalid');
        }
        buttonText = text;
      }
      if (typeof params.color !== 'undefined') {
        if (params.color === false ||
            params.color === null) {
          buttonColor = false;
        } else {
          var color = parseColorToHex(params.color);
          if (!color) {
            console.error('[Telegram.WebApp] Bottom button color format is invalid', params.color);
            throw Error('WebAppBottomButtonParamInvalid');
          }
          buttonColor = color;
        }
      }
      if (typeof params.text_color !== 'undefined') {
        if (params.text_color === false ||
            params.text_color === null) {
          buttonTextColor = false;
        } else {
          var text_color = parseColorToHex(params.text_color);
          if (!text_color) {
            console.error('[Telegram.WebApp] Bottom button text color format is invalid', params.text_color);
            throw Error('WebAppBottomButtonParamInvalid');
          }
          buttonTextColor = text_color;
        }
      }
      if (typeof params.is_visible !== 'undefined') {
        if (params.is_visible &&
            !bottomButton.text.length) {
          console.error('[Telegram.WebApp] Bottom button text is required');
          throw Error('WebAppBottomButtonParamInvalid');
        }
        isVisible = !!params.is_visible;
      }
      if (typeof params.has_shine_effect !== 'undefined') {
        hasShineEffect = !!params.has_shine_effect;
      }
      if (!isMainButton && typeof params.position !== 'undefined') {
        if (params.position != 'left' && params.position != 'right' &&
            params.position != 'top' && params.position != 'bottom') {
          console.error('[Telegram.WebApp] Bottom button posiition is invalid', params.position);
          throw Error('WebAppBottomButtonParamInvalid');
        }
        buttonPosition = params.position;
      }
      if (typeof params.is_active !== 'undefined') {
        isActive = !!params.is_active;
      }
      updateButton();
      return bottomButton;
    }

    bottomButton.setText = function(text) {
      return bottomButton.setParams({text: text});
    };
    bottomButton.onClick = function(callback) {
      onWebViewEvent(webViewEventName, callback);
      return bottomButton;
    };
    bottomButton.offClick = function(callback) {
      offWebViewEvent(webViewEventName, callback);
      return bottomButton;
    };
    bottomButton.show = function() {
      return bottomButton.setParams({is_visible: true});
    };
    bottomButton.hide = function() {
      return bottomButton.setParams({is_visible: false});
    };
    bottomButton.enable = function() {
      return bottomButton.setParams({is_active: true});
    };
    bottomButton.disable = function() {
      return bottomButton.setParams({is_active: false});
    };
    bottomButton.showProgress = function(leaveActive) {
      isActive = !!leaveActive;
      isProgressVisible = true;
      updateButton();
      return bottomButton;
    };
    bottomButton.hideProgress = function() {
      if (!bottomButton.isActive) {
        isActive = true;
      }
      isProgressVisible = false;
      updateButton();
      return bottomButton;
    }
    bottomButton.setParams = setParams;
    return bottomButton;
  };
  var MainButton = BottomButtonConstructor('main');
  var SecondaryButton = BottomButtonConstructor('secondary');

  var SettingsButton = (function() {
    var isVisible = false;

    var settingsButton = {};
    Object.defineProperty(settingsButton, 'isVisible', {
      set: function(val){ setParams({is_visible: val}); },
      get: function(){ return isVisible; },
      enumerable: true
    });

    var curButtonState = null;

    WebView.onEvent('settings_button_pressed', onSettingsButtonPressed);

    function onSettingsButtonPressed() {
      receiveWebViewEvent('settingsButtonClicked');
    }

    function buttonParams() {
      return {is_visible: isVisible};
    }

    function buttonState(btn_params) {
      if (typeof btn_params === 'undefined') {
        btn_params = buttonParams();
      }
      return JSON.stringify(btn_params);
    }

    function buttonCheckVersion() {
      if (!versionAtLeast('6.10')) {
        console.warn('[Telegram.WebApp] SettingsButton is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    function updateButton() {
      var btn_params = buttonParams();
      var btn_state = buttonState(btn_params);
      if (curButtonState === btn_state) {
        return;
      }
      curButtonState = btn_state;
      WebView.postEvent('web_app_setup_settings_button', false, btn_params);
    }

    function setParams(params) {
      if (!buttonCheckVersion()) {
        return settingsButton;
      }
      if (typeof params.is_visible !== 'undefined') {
        isVisible = !!params.is_visible;
      }
      updateButton();
      return settingsButton;
    }

    settingsButton.onClick = function(callback) {
      if (buttonCheckVersion()) {
        onWebViewEvent('settingsButtonClicked', callback);
      }
      return settingsButton;
    };
    settingsButton.offClick = function(callback) {
      if (buttonCheckVersion()) {
        offWebViewEvent('settingsButtonClicked', callback);
      }
      return settingsButton;
    };
    settingsButton.show = function() {
      return setParams({is_visible: true});
    };
    settingsButton.hide = function() {
      return setParams({is_visible: false});
    };
    return settingsButton;
  })();

  var HapticFeedback = (function() {
    var hapticFeedback = {};

    function triggerFeedback(params) {
      if (!versionAtLeast('6.1')) {
        console.warn('[Telegram.WebApp] HapticFeedback is not supported in version ' + webAppVersion);
        return hapticFeedback;
      }
      if (params.type == 'impact') {
        if (params.impact_style != 'light' &&
            params.impact_style != 'medium' &&
            params.impact_style != 'heavy' &&
            params.impact_style != 'rigid' &&
            params.impact_style != 'soft') {
          console.error('[Telegram.WebApp] Haptic impact style is invalid', params.impact_style);
          throw Error('WebAppHapticImpactStyleInvalid');
        }
      } else if (params.type == 'notification') {
        if (params.notification_type != 'error' &&
            params.notification_type != 'success' &&
            params.notification_type != 'warning') {
          console.error('[Telegram.WebApp] Haptic notification type is invalid', params.notification_type);
          throw Error('WebAppHapticNotificationTypeInvalid');
        }
      } else if (params.type == 'selection_change') {
        // no params needed
      } else {
        console.error('[Telegram.WebApp] Haptic feedback type is invalid', params.type);
        throw Error('WebAppHapticFeedbackTypeInvalid');
      }
      WebView.postEvent('web_app_trigger_haptic_feedback', false, params);
      return hapticFeedback;
    }

    hapticFeedback.impactOccurred = function(style) {
      return triggerFeedback({type: 'impact', impact_style: style});
    };
    hapticFeedback.notificationOccurred = function(type) {
      return triggerFeedback({type: 'notification', notification_type: type});
    };
    hapticFeedback.selectionChanged = function() {
      return triggerFeedback({type: 'selection_change'});
    };
    return hapticFeedback;
  })();

  var CloudStorage = (function() {
    var cloudStorage = {};

    function invokeStorageMethod(method, params, callback) {
      if (!versionAtLeast('6.9')) {
        console.error('[Telegram.WebApp] CloudStorage is not supported in version ' + webAppVersion);
        throw Error('WebAppMethodUnsupported');
      }
      invokeCustomMethod(method, params, callback);
      return cloudStorage;
    }

    cloudStorage.setItem = function(key, value, callback) {
      return invokeStorageMethod('saveStorageValue', {key: key, value: value}, callback);
    };
    cloudStorage.getItem = function(key, callback) {
      return cloudStorage.getItems([key], callback ? function(err, res) {
        if (err) callback(err);
        else callback(null, res[key]);
      } : null);
    };
    cloudStorage.getItems = function(keys, callback) {
      return invokeStorageMethod('getStorageValues', {keys: keys}, callback);
    };
    cloudStorage.removeItem = function(key, callback) {
      return cloudStorage.removeItems([key], callback);
    };
    cloudStorage.removeItems = function(keys, callback) {
      return invokeStorageMethod('deleteStorageValues', {keys: keys}, callback);
    };
    cloudStorage.getKeys = function(callback) {
      return invokeStorageMethod('getStorageKeys', {}, callback);
    };
    return cloudStorage;
  })();

  var DeviceStorage = (function() {
    var deviceStorage = {};

    WebView.onEvent('device_storage_key_saved',  onDeviceStorageEvent);
    WebView.onEvent('device_storage_key_received', onDeviceStorageEvent);
    WebView.onEvent('device_storage_cleared',  onDeviceStorageEvent);
    WebView.onEvent('device_storage_failed',  onDeviceStorageEvent);

    function onDeviceStorageEvent(eventType, eventData) {
      if (eventData.req_id && webAppCallbacks[eventData.req_id]) {
        var requestData = webAppCallbacks[eventData.req_id];
        delete webAppCallbacks[eventData.req_id];
        var res = null, err = null;
        if (eventType == 'device_storage_failed') {
          err = eventData.error || 'UNKNOWN_ERROR';
        } else if (eventType == 'device_storage_key_received') {
          res = eventData.value;
        } else {
          res = true;
        }
        if (requestData.callback) {
          requestData.callback(err, res);
        }
      }
    }

    function invokeStorageMethod(method, params, callback) {
      if (!versionAtLeast('9.0')) {
        console.error('[Telegram.WebApp] DeviceStorage is not supported in version ' + webAppVersion);
        throw Error('WebAppMethodUnsupported');
      }
      var req_id = generateCallbackId(16);
      var req_params = {req_id: req_id};
      for (var k in params) {
        req_params[k] = params[k];
      }
      webAppCallbacks[req_id] = {
        callback: callback
      };
      WebView.postEvent(method, false, req_params);
      return deviceStorage;
    }

    deviceStorage.setItem = function(key, value, callback) {
      return invokeStorageMethod('web_app_device_storage_save_key', {key: key, value: value}, callback);
    };
    deviceStorage.getItem = function(key, callback) {
      return invokeStorageMethod('web_app_device_storage_get_key', {key: key}, callback);
    };
    deviceStorage.removeItem = function(key, callback) {
      return invokeStorageMethod('web_app_device_storage_save_key', {key: key, value: null}, callback);
    };
    deviceStorage.clear = function(callback) {
      return invokeStorageMethod('web_app_device_storage_clear', {}, callback);
    };
    return deviceStorage;
  })();

  var SecureStorage = (function() {
    var secureStorage = {};

    WebView.onEvent('secure_storage_key_saved',  onSecureStorageEvent);
    WebView.onEvent('secure_storage_key_received', onSecureStorageEvent);
    WebView.onEvent('secure_storage_key_restored', onSecureStorageEvent);
    WebView.onEvent('secure_storage_cleared',  onSecureStorageEvent);
    WebView.onEvent('secure_storage_failed',  onSecureStorageEvent);

    function onSecureStorageEvent(eventType, eventData) {
      if (eventData.req_id && webAppCallbacks[eventData.req_id]) {
        var requestData = webAppCallbacks[eventData.req_id];
        delete webAppCallbacks[eventData.req_id];
        var res = null, err = null, can_restore = null;
        if (eventType == 'secure_storage_failed') {
          err = eventData.error || 'UNKNOWN_ERROR';
        } else if (eventType == 'secure_storage_key_received') {
          res = eventData.value;
          if (eventData.can_restore) {
            can_restore = true;
          }
        } else if (eventType == 'secure_storage_key_restored') {
          res = eventData.value;
        } else {
          res = true;
        }
        if (requestData.callback) {
          requestData.callback(err, res, can_restore);
        }
      }
    }

    function invokeStorageMethod(method, params, callback) {
      if (!versionAtLeast('9.0')) {
        console.error('[Telegram.WebApp] SecureStorage is not supported in version ' + webAppVersion);
        throw Error('WebAppMethodUnsupported');
      }
      var req_id = generateCallbackId(16);
      var req_params = {req_id: req_id};
      for (var k in params) {
        req_params[k] = params[k];
      }
      webAppCallbacks[req_id] = {
        callback: callback
      };
      WebView.postEvent(method, false, req_params);
      return secureStorage;
    }

    secureStorage.setItem = function(key, value, callback) {
      return invokeStorageMethod('web_app_secure_storage_save_key', {key: key, value: value}, callback);
    };
    secureStorage.getItem = function(key, callback) {
      return invokeStorageMethod('web_app_secure_storage_get_key', {key: key}, callback);
    };
    secureStorage.restoreItem = function(key, callback) {
      return invokeStorageMethod('web_app_secure_storage_restore_key', {key: key}, callback);
    };
    secureStorage.removeItem = function(key, callback) {
      return invokeStorageMethod('web_app_secure_storage_save_key', {key: key, value: null}, callback);
    };
    secureStorage.clear = function(callback) {
      return invokeStorageMethod('web_app_secure_storage_clear', {}, callback);
    };
    return secureStorage;
  })();

  var BiometricManager = (function() {
    var isInited = false;
    var isBiometricAvailable = false;
    var biometricType = 'unknown';
    var isAccessRequested = false;
    var isAccessGranted = false;
    var isBiometricTokenSaved = false;
    var deviceId = '';

    var biometricManager = {};
    Object.defineProperty(biometricManager, 'isInited', {
      get: function(){ return isInited; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'isBiometricAvailable', {
      get: function(){ return isInited && isBiometricAvailable; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'biometricType', {
      get: function(){ return biometricType || 'unknown'; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'isAccessRequested', {
      get: function(){ return isAccessRequested; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'isAccessGranted', {
      get: function(){ return isAccessRequested && isAccessGranted; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'isBiometricTokenSaved', {
      get: function(){ return isBiometricTokenSaved; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'deviceId', {
      get: function(){ return deviceId || ''; },
      enumerable: true
    });

    var initRequestState = {callbacks: []};
    var accessRequestState = false;
    var authRequestState = false;
    var tokenRequestState = false;

    WebView.onEvent('biometry_info_received',  onBiometryInfoReceived);
    WebView.onEvent('biometry_auth_requested', onBiometryAuthRequested);
    WebView.onEvent('biometry_token_updated',  onBiometryTokenUpdated);

    function onBiometryInfoReceived(eventType, eventData) {
      isInited = true;
      if (eventData.available) {
        isBiometricAvailable = true;
        biometricType = eventData.type || 'unknown';
        if (eventData.access_requested) {
          isAccessRequested = true;
          isAccessGranted = !!eventData.access_granted;
          isBiometricTokenSaved = !!eventData.token_saved;
        } else {
          isAccessRequested = false;
          isAccessGranted = false;
          isBiometricTokenSaved = false;
        }
      } else {
        isBiometricAvailable = false;
        biometricType = 'unknown';
        isAccessRequested = false;
        isAccessGranted = false;
        isBiometricTokenSaved = false;
      }
      deviceId = eventData.device_id || '';

      if (initRequestState.callbacks.length > 0) {
        for (var i = 0; i < initRequestState.callbacks.length; i++) {
          var callback = initRequestState.callbacks[i];
          callback();
        }
        initRequestState.callbacks = [];
      }
      if (accessRequestState) {
        var state = accessRequestState;
        accessRequestState = false;
        if (state.callback) {
          state.callback(isAccessGranted);
        }
      }
      receiveWebViewEvent('biometricManagerUpdated');
    }
    function onBiometryAuthRequested(eventType, eventData) {
      var isAuthenticated = (eventData.status == 'authorized'),
          biometricToken = eventData.token || '';
      if (authRequestState) {
        var state = authRequestState;
        authRequestState = false;
        if (state.callback) {
          state.callback(isAuthenticated, isAuthenticated ? biometricToken : null);
        }
      }
      receiveWebViewEvent('biometricAuthRequested', isAuthenticated ? {
        isAuthenticated: true,
        biometricToken: biometricToken
      } : {
        isAuthenticated: false
      });
    }
    function onBiometryTokenUpdated(eventType, eventData) {
      var applied = false;
      if (isBiometricAvailable &&
          isAccessRequested) {
        if (eventData.status == 'updated') {
          isBiometricTokenSaved = true;
          applied = true;
        }
        else if (eventData.status == 'removed') {
          isBiometricTokenSaved = false;
          applied = true;
        }
      }
      if (tokenRequestState) {
        var state = tokenRequestState;
        tokenRequestState = false;
        if (state.callback) {
          state.callback(applied);
        }
      }
      receiveWebViewEvent('biometricTokenUpdated', {
        isUpdated: applied
      });
    }

    function checkVersion() {
      if (!versionAtLeast('7.2')) {
        console.warn('[Telegram.WebApp] BiometricManager is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    function checkInit() {
      if (!isInited) {
        console.error('[Telegram.WebApp] BiometricManager should be inited before using.');
        throw Error('WebAppBiometricManagerNotInited');
      }
      return true;
    }

    biometricManager.init = function(callback) {
      if (!checkVersion()) {
        return biometricManager;
      }
      if (isInited) {
        return biometricManager;
      }
      if (callback) {
        initRequestState.callbacks.push(callback);
      }
      WebView.postEvent('web_app_biometry_get_info', false);
      return biometricManager;
    };
    biometricManager.requestAccess = function(params, callback) {
      if (!checkVersion()) {
        return biometricManager;
      }
      checkInit();
      if (!isBiometricAvailable) {
        console.error('[Telegram.WebApp] Biometrics is not available on this device.');
        throw Error('WebAppBiometricManagerBiometricsNotAvailable');
      }
      if (accessRequestState) {
        console.error('[Telegram.WebApp] Access is already requested');
        throw Error('WebAppBiometricManagerAccessRequested');
      }
      var popup_params = {};
      if (typeof params.reason !== 'undefined') {
        var reason = strTrim(params.reason);
        if (reason.length > 128) {
          console.error('[Telegram.WebApp] Biometric reason is too long', reason);
          throw Error('WebAppBiometricRequestAccessParamInvalid');
        }
        if (reason.length > 0) {
          popup_params.reason = reason;
        }
      }

      accessRequestState = {
        callback: callback
      };
      WebView.postEvent('web_app_biometry_request_access', false, popup_params);
      return biometricManager;
    };
    biometricManager.authenticate = function(params, callback) {
      if (!checkVersion()) {
        return biometricManager;
      }
      checkInit();
      if (!isBiometricAvailable) {
        console.error('[Telegram.WebApp] Biometrics is not available on this device.');
        throw Error('WebAppBiometricManagerBiometricsNotAvailable');
      }
      if (!isAccessGranted) {
        console.error('[Telegram.WebApp] Biometric access was not granted by the user.');
        throw Error('WebAppBiometricManagerBiometricAccessNotGranted');
      }
      if (authRequestState) {
        console.error('[Telegram.WebApp] Authentication request is already in progress.');
        throw Error('WebAppBiometricManagerAuthenticationRequested');
      }
      var popup_params = {};
      if (typeof params.reason !== 'undefined') {
        var reason = strTrim(params.reason);
        if (reason.length > 128) {
          console.error('[Telegram.WebApp] Biometric reason is too long', reason);
          throw Error('WebAppBiometricRequestAccessParamInvalid');
        }
        if (reason.length > 0) {
          popup_params.reason = reason;
        }
      }

      authRequestState = {
        callback: callback
      };
      WebView.postEvent('web_app_biometry_request_auth', false, popup_params);
      return biometricManager;
    };
    biometricManager.updateBiometricToken = function(token, callback) {
      if (!checkVersion()) {
        return biometricManager;
      }
      token = token || '';
      if (token.length > 1024) {
        console.error('[Telegram.WebApp] Token is too long', token);
        throw Error('WebAppBiometricManagerTokenInvalid');
      }
      checkInit();
      if (!isBiometricAvailable) {
        console.error('[Telegram.WebApp] Biometrics is not available on this device.');
        throw Error('WebAppBiometricManagerBiometricsNotAvailable');
      }
      if (!isAccessGranted) {
        console.error('[Telegram.WebApp] Biometric access was not granted by the user.');
        throw Error('WebAppBiometricManagerBiometricAccessNotGranted');
      }
      if (tokenRequestState) {
        console.error('[Telegram.WebApp] Token request is already in progress.');
        throw Error('WebAppBiometricManagerTokenUpdateRequested');
      }
      tokenRequestState = {
        callback: callback
      };
      WebView.postEvent('web_app_biometry_update_token', false, {token: token});
      return biometricManager;
    };
    biometricManager.openSettings = function() {
      if (!checkVersion()) {
        return biometricManager;
      }
      checkInit();
      if (!isBiometricAvailable) {
        console.error('[Telegram.WebApp] Biometrics is not available on this device.');
        throw Error('WebAppBiometricManagerBiometricsNotAvailable');
      }
      if (!isAccessRequested) {
        console.error('[Telegram.WebApp] Biometric access was not requested yet.');
        throw Error('WebAppBiometricManagerBiometricsAccessNotRequested');
      }
      if (isAccessGranted) {
        console.warn('[Telegram.WebApp] Biometric access was granted by the user, no need to go to settings.');
        return biometricManager;
      }
      WebView.postEvent('web_app_biometry_open_settings', false);
      return biometricManager;
    };
    return biometricManager;
  })();

  var LocationManager = (function() {
    var isInited = false;
    var isLocationAvailable = false;
    var isAccessRequested = false;
    var isAccessGranted = false;

    var locationManager = {};
    Object.defineProperty(locationManager, 'isInited', {
      get: function(){ return isInited; },
      enumerable: true
    });
    Object.defineProperty(locationManager, 'isLocationAvailable', {
      get: function(){ return isInited && isLocationAvailable; },
      enumerable: true
    });
    Object.defineProperty(locationManager, 'isAccessRequested', {
      get: function(){ return isAccessRequested; },
      enumerable: true
    });
    Object.defineProperty(locationManager, 'isAccessGranted', {
      get: function(){ return isAccessRequested && isAccessGranted; },
      enumerable: true
    });

    var initRequestState = {callbacks: []};
    var getRequestState = {callbacks: []};

    WebView.onEvent('location_checked',  onLocationChecked);
    WebView.onEvent('location_requested', onLocationRequested);

    function onLocationChecked(eventType, eventData) {
      isInited = true;
      if (eventData.available) {
        isLocationAvailable = true;
        if (eventData.access_requested) {
          isAccessRequested = true;
          isAccessGranted = !!eventData.access_granted;
        } else {
          isAccessRequested = false;
          isAccessGranted = false;
        }
      } else {
        isLocationAvailable = false;
        isAccessRequested = false;
        isAccessGranted = false;
      }

      if (initRequestState.callbacks.length > 0) {
        for (var i = 0; i < initRequestState.callbacks.length; i++) {
          var callback = initRequestState.callbacks[i];
          callback();
        }
        initRequestState.callbacks = [];
      }
      receiveWebViewEvent('locationManagerUpdated');
    }
    function onLocationRequested(eventType, eventData) {
      if (!eventData.available) {
        locationData = null;
      } else {
        var locationData = {
          latitude: eventData.latitude,
          longitude: eventData.longitude,
          altitude: null,
          course: null,
          speed: null,
          horizontal_accuracy: null,
          vertical_accuracy: null,
          course_accuracy: null,
          speed_accuracy: null,
        };
        if (typeof eventData.altitude !== 'undefined' && eventData.altitude !== null) {
          locationData.altitude = eventData.altitude;
        }
        if (typeof eventData.course !== 'undefined' && eventData.course !== null) {
          locationData.course = eventData.course % 360;
        }
        if (typeof eventData.speed !== 'undefined' && eventData.speed !== null) {
          locationData.speed = eventData.speed;
        }
        if (typeof eventData.horizontal_accuracy !== 'undefined' && eventData.horizontal_accuracy !== null) {
          locationData.horizontal_accuracy = eventData.horizontal_accuracy;
        }
        if (typeof eventData.vertical_accuracy !== 'undefined' && eventData.vertical_accuracy !== null) {
          locationData.vertical_accuracy = eventData.vertical_accuracy;
        }
        if (typeof eventData.course_accuracy !== 'undefined' && eventData.course_accuracy !== null) {
          locationData.course_accuracy = eventData.course_accuracy;
        }
        if (typeof eventData.speed_accuracy !== 'undefined' && eventData.speed_accuracy !== null) {
          locationData.speed_accuracy = eventData.speed_accuracy;
        }
      }
      if (!eventData.available ||
          !isLocationAvailable ||
          !isAccessRequested ||
          !isAccessGranted) {
        initRequestState.callbacks.push(function() {
          locationResponse(locationData);
        });
        WebView.postEvent('web_app_check_location', false);
      } else {
        locationResponse(locationData);
      }
    }
    function locationResponse(response) {
      if (getRequestState.callbacks.length > 0) {
        for (var i = 0; i < getRequestState.callbacks.length; i++) {
          var callback = getRequestState.callbacks[i];
          callback(response);
        }
        getRequestState.callbacks = [];
      }
      if (response !== null) {
        receiveWebViewEvent('locationRequested', {
          locationData: response
        });
      }
    }

    function checkVersion() {
      if (!versionAtLeast('8.0')) {
        console.warn('[Telegram.WebApp] LocationManager is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    function checkInit() {
      if (!isInited) {
        console.error('[Telegram.WebApp] LocationManager should be inited before using.');
        throw Error('WebAppLocationManagerNotInited');
      }
      return true;
    }

    locationManager.init = function(callback) {
      if (!checkVersion()) {
        return locationManager;
      }
      if (isInited) {
        return locationManager;
      }
      if (callback) {
        initRequestState.callbacks.push(callback);
      }
      WebView.postEvent('web_app_check_location', false);
      return locationManager;
    };
    locationManager.getLocation = function(callback) {
      if (!checkVersion()) {
        return locationManager;
      }
      checkInit();
      if (!isLocationAvailable) {
        console.error('[Telegram.WebApp] Location is not available on this device.');
        throw Error('WebAppLocationManagerLocationNotAvailable');
      }

      getRequestState.callbacks.push(callback);
      WebView.postEvent('web_app_request_location');
      return locationManager;
    };
    locationManager.openSettings = function() {
      if (!checkVersion()) {
        return locationManager;
      }
      checkInit();
      if (!isLocationAvailable) {
        console.error('[Telegram.WebApp] Location is not available on this device.');
        throw Error('WebAppLocationManagerLocationNotAvailable');
      }
      if (!isAccessRequested) {
        console.error('[Telegram.WebApp] Location access was not requested yet.');
        throw Error('WebAppLocationManagerLocationAccessNotRequested');
      }
      if (isAccessGranted) {
        console.warn('[Telegram.WebApp] Location access was granted by the user, no need to go to settings.');
        return locationManager;
      }
      WebView.postEvent('web_app_open_location_settings', false);
      return locationManager;
    };
    return locationManager;
  })();

  var Accelerometer = (function() {
    var isStarted = false;
    var valueX = null, valueY = null, valueZ = null;
    var startCallbacks = [], stopCallbacks = [];

    var accelerometer = {};
    Object.defineProperty(accelerometer, 'isStarted', {
      get: function(){ return isStarted; },
      enumerable: true
    });
    Object.defineProperty(accelerometer, 'x', {
      get: function(){ return valueX; },
      enumerable: true
    });
    Object.defineProperty(accelerometer, 'y', {
      get: function(){ return valueY; },
      enumerable: true
    });
    Object.defineProperty(accelerometer, 'z', {
      get: function(){ return valueZ; },
      enumerable: true
    });

    WebView.onEvent('accelerometer_started', onAccelerometerStarted);
    WebView.onEvent('accelerometer_stopped', onAccelerometerStopped);
    WebView.onEvent('accelerometer_changed', onAccelerometerChanged);
    WebView.onEvent('accelerometer_failed',  onAccelerometerFailed);

    function onAccelerometerStarted(eventType, eventData) {
      isStarted = true;
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(true);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('accelerometerStarted');
    }
    function onAccelerometerStopped(eventType, eventData) {
      isStarted = false;
      if (stopCallbacks.length > 0) {
        for (var i = 0; i < stopCallbacks.length; i++) {
          var callback = stopCallbacks[i];
          callback(true);
        }
        stopCallbacks = [];
      }
      receiveWebViewEvent('accelerometerStopped');
    }
    function onAccelerometerChanged(eventType, eventData) {
      valueX = eventData.x;
      valueY = eventData.y;
      valueZ = eventData.z;
      receiveWebViewEvent('accelerometerChanged');
    }
    function onAccelerometerFailed(eventType, eventData) {
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(false);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('accelerometerFailed', {
        error: eventData.error
      });
    }

    function checkVersion() {
      if (!versionAtLeast('8.0')) {
        console.warn('[Telegram.WebApp] Accelerometer is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    accelerometer.start = function(params, callback) {
      params = params || {};
      if (!checkVersion()) {
        return accelerometer;
      }
      var req_params = {};
      var refresh_rate = parseInt(params.refresh_rate || 1000);
      if (isNaN(refresh_rate) || refresh_rate < 20 || refresh_rate > 1000) {
        console.warn('[Telegram.WebApp] Accelerometer refresh_rate is invalid', refresh_rate);
      } else {
        req_params.refresh_rate = refresh_rate;
      }

      if (callback) {
        startCallbacks.push(callback);
      }
      WebView.postEvent('web_app_start_accelerometer', false, req_params);
      return accelerometer;
    };
    accelerometer.stop = function(callback) {
      if (!checkVersion()) {
        return accelerometer;
      }
      if (callback) {
        stopCallbacks.push(callback);
      }
      WebView.postEvent('web_app_stop_accelerometer');
      return accelerometer;
    };
    return accelerometer;
  })();

  var DeviceOrientation = (function() {
    var isStarted = false;
    var valueAlpha = null, valueBeta = null, valueGamma = null, valueAbsolute = false;
    var startCallbacks = [], stopCallbacks = [];

    var deviceOrientation = {};
    Object.defineProperty(deviceOrientation, 'isStarted', {
      get: function(){ return isStarted; },
      enumerable: true
    });
    Object.defineProperty(deviceOrientation, 'absolute', {
      get: function(){ return valueAbsolute; },
      enumerable: true
    });
    Object.defineProperty(deviceOrientation, 'alpha', {
      get: function(){ return valueAlpha; },
      enumerable: true
    });
    Object.defineProperty(deviceOrientation, 'beta', {
      get: function(){ return valueBeta; },
      enumerable: true
    });
    Object.defineProperty(deviceOrientation, 'gamma', {
      get: function(){ return valueGamma; },
      enumerable: true
    });

    WebView.onEvent('device_orientation_started',  onDeviceOrientationStarted);
    WebView.onEvent('device_orientation_stopped',  onDeviceOrientationStopped);
    WebView.onEvent('device_orientation_changed', onDeviceOrientationChanged);
    WebView.onEvent('device_orientation_failed',  onDeviceOrientationFailed);

    function onDeviceOrientationStarted(eventType, eventData) {
      isStarted = true;
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(true);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('deviceOrientationStarted');
    }
    function onDeviceOrientationStopped(eventType, eventData) {
      isStarted = false;
      if (stopCallbacks.length > 0) {
        for (var i = 0; i < stopCallbacks.length; i++) {
          var callback = stopCallbacks[i];
          callback(true);
        }
        stopCallbacks = [];
      }
      receiveWebViewEvent('deviceOrientationStopped');
    }
    function onDeviceOrientationChanged(eventType, eventData) {
      valueAbsolute = !!eventData.absolute;
      valueAlpha = eventData.alpha;
      valueBeta  = eventData.beta;
      valueGamma = eventData.gamma;
      receiveWebViewEvent('deviceOrientationChanged');
    }
    function onDeviceOrientationFailed(eventType, eventData) {
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(false);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('deviceOrientationFailed', {
        error: eventData.error
      });
    }

    function checkVersion() {
      if (!versionAtLeast('8.0')) {
        console.warn('[Telegram.WebApp] DeviceOrientation is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    deviceOrientation.start = function(params, callback) {
      params = params || {};
      if (!checkVersion()) {
        return deviceOrientation;
      }
      var req_params = {};
      var refresh_rate = parseInt(params.refresh_rate || 1000);
      if (isNaN(refresh_rate) || refresh_rate < 20 || refresh_rate > 1000) {
        console.warn('[Telegram.WebApp] DeviceOrientation refresh_rate is invalid', refresh_rate);
      } else {
        req_params.refresh_rate = refresh_rate;
      }
      req_params.need_absolute = !!params.need_absolute;

      if (callback) {
        startCallbacks.push(callback);
      }
      WebView.postEvent('web_app_start_device_orientation', false, req_params);
      return deviceOrientation;
    };
    deviceOrientation.stop = function(callback) {
      if (!checkVersion()) {
        return deviceOrientation;
      }
      if (callback) {
        stopCallbacks.push(callback);
      }
      WebView.postEvent('web_app_stop_device_orientation');
      return deviceOrientation;
    };
    return deviceOrientation;
  })();

  var Gyroscope = (function() {
    var isStarted = false;
    var valueX = null, valueY = null, valueZ = null;
    var startCallbacks = [], stopCallbacks = [];

    var gyroscope = {};
    Object.defineProperty(gyroscope, 'isStarted', {
      get: function(){ return isStarted; },
      enumerable: true
    });
    Object.defineProperty(gyroscope, 'x', {
      get: function(){ return valueX; },
      enumerable: true
    });
    Object.defineProperty(gyroscope, 'y', {
      get: function(){ return valueY; },
      enumerable: true
    });
    Object.defineProperty(gyroscope, 'z', {
      get: function(){ return valueZ; },
      enumerable: true
    });

    WebView.onEvent('gyroscope_started',  onGyroscopeStarted);
    WebView.onEvent('gyroscope_stopped',  onGyroscopeStopped);
    WebView.onEvent('gyroscope_changed', onGyroscopeChanged);
    WebView.onEvent('gyroscope_failed',  onGyroscopeFailed);

    function onGyroscopeStarted(eventType, eventData) {
      isStarted = true;
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(true);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('gyroscopeStarted');
    }
    function onGyroscopeStopped(eventType, eventData) {
      isStarted = false;
      if (stopCallbacks.length > 0) {
        for (var i = 0; i < stopCallbacks.length; i++) {
          var callback = stopCallbacks[i];
          callback(true);
        }
        stopCallbacks = [];
      }
      receiveWebViewEvent('gyroscopeStopped');
    }
    function onGyroscopeChanged(eventType, eventData) {
      valueX = eventData.x;
      valueY = eventData.y;
      valueZ = eventData.z;
      receiveWebViewEvent('gyroscopeChanged');
    }
    function onGyroscopeFailed(eventType, eventData) {
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(false);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('gyroscopeFailed', {
        error: eventData.error
      });
    }

    function checkVersion() {
      if (!versionAtLeast('8.0')) {
        console.warn('[Telegram.WebApp] Gyroscope is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    gyroscope.start = function(params, callback) {
      params = params || {};
      if (!checkVersion()) {
        return gyroscope;
      }
      var req_params = {};
      var refresh_rate = parseInt(params.refresh_rate || 1000);
      if (isNaN(refresh_rate) || refresh_rate < 20 || refresh_rate > 1000) {
        console.warn('[Telegram.WebApp] Gyroscope refresh_rate is invalid', refresh_rate);
      } else {
        req_params.refresh_rate = refresh_rate;
      }

      if (callback) {
        startCallbacks.push(callback);
      }
      WebView.postEvent('web_app_start_gyroscope', false, req_params);
      return gyroscope;
    };
    gyroscope.stop = function(callback) {
      if (!checkVersion()) {
        return gyroscope;
      }
      if (callback) {
        stopCallbacks.push(callback);
      }
      WebView.postEvent('web_app_stop_gyroscope');
      return gyroscope;
    };
    return gyroscope;
  })();

  var webAppInvoices = {};
  function onInvoiceClosed(eventType, eventData) {
    if (eventData.slug && webAppInvoices[eventData.slug]) {
      var invoiceData = webAppInvoices[eventData.slug];
      delete webAppInvoices[eventData.slug];
      if (invoiceData.callback) {
        invoiceData.callback(eventData.status);
      }
      receiveWebViewEvent('invoiceClosed', {
        url: invoiceData.url,
        status: eventData.status
      });
    }
  }

  var webAppPopupOpened = false;
  function onPopupClosed(eventType, eventData) {
    if (webAppPopupOpened) {
      var popupData = webAppPopupOpened;
      webAppPopupOpened = false;
      var button_id = null;
      if (typeof eventData.button_id !== 'undefined') {
        button_id = eventData.button_id;
      }
      if (popupData.callback) {
        popupData.callback(button_id);
      }
      receiveWebViewEvent('popupClosed', {
        button_id: button_id
      });
    }
  }

  var webAppScanQrPopupOpened = false;
  function onQrTextReceived(eventType, eventData) {
    if (webAppScanQrPopupOpened) {
      var popupData = webAppScanQrPopupOpened;
      var data = null;
      if (typeof eventData.data !== 'undefined') {
        data = eventData.data;
      }
      if (popupData.callback) {
        if (popupData.callback(data)) {
          webAppScanQrPopupOpened = false;
          WebView.postEvent('web_app_close_scan_qr_popup', false);
        }
      }
      receiveWebViewEvent('qrTextReceived', {
        data: data
      });
    }
  }
  function onScanQrPopupClosed(eventType, eventData) {
    webAppScanQrPopupOpened = false;
    receiveWebViewEvent('scanQrPopupClosed');
  }

  function onClipboardTextReceived(eventType, eventData) {
    if (eventData.req_id && webAppCallbacks[eventData.req_id]) {
      var requestData = webAppCallbacks[eventData.req_id];
      delete webAppCallbacks[eventData.req_id];
      var data = null;
      if (typeof eventData.data !== 'undefined') {
        data = eventData.data;
      }
      if (requestData.callback) {
        requestData.callback(data);
      }
      receiveWebViewEvent('clipboardTextReceived', {
        data: data
      });
    }
  }

  var WebAppWriteAccessRequested = false;
  function onWriteAccessRequested(eventType, eventData) {
    if (WebAppWriteAccessRequested) {
      var requestData = WebAppWriteAccessRequested;
      WebAppWriteAccessRequested = false;
      if (requestData.callback) {
        requestData.callback(eventData.status == 'allowed');
      }
      receiveWebViewEvent('writeAccessRequested', {
        status: eventData.status
      });
    }
  }

  function getRequestedContact(callback, timeout) {
    var reqTo, fallbackTo, reqDelay = 0;
    var reqInvoke = function() {
      invokeCustomMethod('getRequestedContact', {}, function(err, res) {
        if (res.substr(0, 1) == '"' && res.substr(-1) == '"') { // macos fix
          res = JSON.parse(res);
        }
        if (res && res.length) {
          clearTimeout(fallbackTo);
          callback(res);
        } else {
          reqDelay += 50;
          reqTo = setTimeout(reqInvoke, reqDelay);
        }
      });
    };
    var fallbackInvoke = function() {
      clearTimeout(reqTo);
      callback('');
    };
    fallbackTo = setTimeout(fallbackInvoke, timeout);
    reqInvoke();
  }

  var WebAppContactRequested = false;
  function onPhoneRequested(eventType, eventData) {
    if (WebAppContactRequested) {
      var requestData = WebAppContactRequested;
      WebAppContactRequested = false;
      var requestSent = eventData.status == 'sent';
      var webViewEvent = {
        status: eventData.status
      };
      if (requestSent) {
        getRequestedContact(function(res) {
          if (res && res.length) {
            webViewEvent.response = res;
            webViewEvent.responseUnsafe = Utils.urlParseQueryString(res);
            for (var key in webViewEvent.responseUnsafe) {
              var val = webViewEvent.responseUnsafe[key];
              try {
                if (val.substr(0, 1) == '{' && val.substr(-1) == '}' ||
                    val.substr(0, 1) == '[' && val.substr(-1) == ']') {
                  webViewEvent.responseUnsafe[key] = JSON.parse(val);
                }
              } catch (e) {}
            }
          }
          if (requestData.callback) {
            requestData.callback(requestSent, webViewEvent);
          }
          receiveWebViewEvent('contactRequested', webViewEvent);
        }, 3000);
      } else {
        if (requestData.callback) {
          requestData.callback(requestSent, webViewEvent);
        }
        receiveWebViewEvent('contactRequested', webViewEvent);
      }
    }
  }

  var webAppDownloadFileRequested = false;
  function onFileDownloadRequested(eventType, eventData) {
    if (webAppDownloadFileRequested) {
      var requestData = webAppDownloadFileRequested;
      webAppDownloadFileRequested = false;
      var isDownloading = eventData.status == 'downloading';
      if (requestData.callback) {
        requestData.callback(isDownloading);
      }
      receiveWebViewEvent('fileDownloadRequested', {
        status: isDownloading ? 'downloading' : 'cancelled'
      });
    }
  }

  function onCustomMethodInvoked(eventType, eventData) {
    if (eventData.req_id && webAppCallbacks[eventData.req_id]) {
      var requestData = webAppCallbacks[eventData.req_id];
      delete webAppCallbacks[eventData.req_id];
      var res = null, err = null;
      if (typeof eventData.result !== 'undefined') {
        res = eventData.result;
      }
      if (typeof eventData.error !== 'undefined') {
        err = eventData.error;
      }
      if (requestData.callback) {
        requestData.callback(err, res);
      }
    }
  }

  function invokeCustomMethod(method, params, callback) {
    if (!versionAtLeast('6.9')) {
      console.error('[Telegram.WebApp] Method invokeCustomMethod is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    var req_id = generateCallbackId(16);
    var req_params = {req_id: req_id, method: method, params: params || {}};
    webAppCallbacks[req_id] = {
      callback: callback
    };
    WebView.postEvent('web_app_invoke_custom_method', false, req_params);
  };

  if (!window.Telegram) {
    window.Telegram = {};
  }

  Object.defineProperty(WebApp, 'initData', {
    get: function(){ return webAppInitData; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'initDataUnsafe', {
    get: function(){ return webAppInitDataUnsafe; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'version', {
    get: function(){ return webAppVersion; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'platform', {
    get: function(){ return webAppPlatform; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'colorScheme', {
    get: function(){ return colorScheme; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'themeParams', {
    get: function(){ return themeParams; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isExpanded', {
    get: function(){ return isExpanded; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'viewportHeight', {
    get: function(){ return (viewportHeight === false ? window.innerHeight : viewportHeight) - bottomBarHeight; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'viewportStableHeight', {
    get: function(){ return (viewportStableHeight === false ? window.innerHeight : viewportStableHeight) - bottomBarHeight; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'safeAreaInset', {
    get: function(){ return safeAreaInset; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'contentSafeAreaInset', {
    get: function(){ return contentSafeAreaInset; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isClosingConfirmationEnabled', {
    set: function(val){ setClosingConfirmation(val); },
    get: function(){ return isClosingConfirmationEnabled; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isVerticalSwipesEnabled', {
    set: function(val){ toggleVerticalSwipes(val); },
    get: function(){ return isVerticalSwipesEnabled; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isFullscreen', {
    get: function(){ return webAppIsFullscreen; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isOrientationLocked', {
    set: function(val){ toggleOrientationLock(val); },
    get: function(){ return webAppIsOrientationLocked; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isActive', {
    get: function(){ return webAppIsActive; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'headerColor', {
    set: function(val){ setHeaderColor(val); },
    get: function(){ return getHeaderColor(); },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'backgroundColor', {
    set: function(val){ setBackgroundColor(val); },
    get: function(){ return getBackgroundColor(); },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'bottomBarColor', {
    set: function(val){ setBottomBarColor(val); },
    get: function(){ return getBottomBarColor(); },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'BackButton', {
    value: BackButton,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'MainButton', {
    value: MainButton,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'SecondaryButton', {
    value: SecondaryButton,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'SettingsButton', {
    value: SettingsButton,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'HapticFeedback', {
    value: HapticFeedback,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'CloudStorage', {
    value: CloudStorage,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'DeviceStorage', {
    value: DeviceStorage,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'SecureStorage', {
    value: SecureStorage,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'BiometricManager', {
    value: BiometricManager,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'Accelerometer', {
    value: Accelerometer,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'DeviceOrientation', {
    value: DeviceOrientation,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'Gyroscope', {
    value: Gyroscope,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'LocationManager', {
    value: LocationManager,
    enumerable: true
  });
  WebApp.isVersionAtLeast = function(ver) {
    return versionAtLeast(ver);
  };
  WebApp.setHeaderColor = function(color_key) {
    WebApp.headerColor = color_key;
  };
  WebApp.setBackgroundColor = function(color) {
    WebApp.backgroundColor = color;
  };
  WebApp.setBottomBarColor = function(color) {
    WebApp.bottomBarColor = color;
  };
  WebApp.enableClosingConfirmation = function() {
    WebApp.isClosingConfirmationEnabled = true;
  };
  WebApp.disableClosingConfirmation = function() {
    WebApp.isClosingConfirmationEnabled = false;
  };
  WebApp.enableVerticalSwipes = function() {
    WebApp.isVerticalSwipesEnabled = true;
  };
  WebApp.disableVerticalSwipes = function() {
    WebApp.isVerticalSwipesEnabled = false;
  };
  WebApp.lockOrientation = function() {
    WebApp.isOrientationLocked = true;
  };
  WebApp.unlockOrientation = function() {
    WebApp.isOrientationLocked = false;
  };
  WebApp.requestFullscreen = function() {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method requestFullscreen is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    WebView.postEvent('web_app_request_fullscreen');
  };
  WebApp.exitFullscreen = function() {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method exitFullscreen is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    WebView.postEvent('web_app_exit_fullscreen');
  };
  WebApp.addToHomeScreen = function() {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method addToHomeScreen is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    WebView.postEvent('web_app_add_to_home_screen');
  };
  WebApp.checkHomeScreenStatus = function(callback) {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method checkHomeScreenStatus is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (callback) {
      homeScreenCallbacks.push(callback);
    }
    WebView.postEvent('web_app_check_home_screen');
  };
  WebApp.onEvent = function(eventType, callback) {
    onWebViewEvent(eventType, callback);
  };
  WebApp.offEvent = function(eventType, callback) {offWebViewEvent(eventType, callback);
  };
  WebApp.sendData = function (data) {
    if (!data || !data.length) {
      console.error('[Telegram.WebApp] Data is required', data);
      throw Error('WebAppDataInvalid');
    }
    if (byteLength(data) > 4096) {
      console.error('[Telegram.WebApp] Data is too long', data);
      throw Error('WebAppDataInvalid');
    }
    WebView.postEvent('web_app_data_send', false, {data: data});
  };
  WebApp.switchInlineQuery = function (query, choose_chat_types) {
    if (!versionAtLeast('6.6')) {
      console.error('[Telegram.WebApp] Method switchInlineQuery is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (!initParams.tgWebAppBotInline) {
      console.error('[Telegram.WebApp] Inline mode is disabled for this bot. Read more about inline mode: https://core.telegram.org/bots/inline');
      throw Error('WebAppInlineModeDisabled');
    }
    query = query || '';
    if (query.length > 256) {
      console.error('[Telegram.WebApp] Inline query is too long', query);
      throw Error('WebAppInlineQueryInvalid');
    }
    var chat_types = [];
    if (choose_chat_types) {
      if (!Array.isArray(choose_chat_types)) {
        console.error('[Telegram.WebApp] Choose chat types should be an array', choose_chat_types);
        throw Error('WebAppInlineChooseChatTypesInvalid');
      }
      var good_types = {users: 1, bots: 1, groups: 1, channels: 1};
      for (var i = 0; i < choose_chat_types.length; i++) {
        var chat_type = choose_chat_types[i];
        if (!good_types[chat_type]) {
          console.error('[Telegram.WebApp] Choose chat type is invalid', chat_type);
          throw Error('WebAppInlineChooseChatTypeInvalid');
        }
        if (good_types[chat_type] != 2) {
          good_types[chat_type] = 2;
          chat_types.push(chat_type);
        }
      }
    }
    WebView.postEvent('web_app_switch_inline_query', false, {query: query, chat_types: chat_types});
  };
  WebApp.openLink = function (url, options) {
    var a = document.createElement('A');
    a.href = url;
    if (a.protocol != 'http:' &&
        a.protocol != 'https:') {
      console.error('[Telegram.WebApp] Url protocol is not supported', url);
      throw Error('WebAppTgUrlInvalid');
    }
    var url = a.href;
    options = options || {};
    if (versionAtLeast('6.1')) {
      var req_params = {url: url};
      if (versionAtLeast('6.4') && options.try_instant_view) {
        req_params.try_instant_view = true;
      }
      if (versionAtLeast('7.6') && options.try_browser) {
        req_params.try_browser = options.try_browser;
      }
      WebView.postEvent('web_app_open_link', false, req_params);
    } else {
      window.open(url, '_blank');
    }
  };
  WebApp.openTelegramLink = function (url, options) {
    var a = document.createElement('A');
    a.href = url;
    if (a.protocol != 'http:' &&
        a.protocol != 'https:') {
      console.error('[Telegram.WebApp] Url protocol is not supported', url);
      throw Error('WebAppTgUrlInvalid');
    }
    if (!isTmeHostname(a.hostname)) {
      console.error('[Telegram.WebApp] Url host is not supported', url);
      throw Error('WebAppTgUrlInvalid');
    }
    var path_full = a.pathname + a.search;
    options = options || {};
    if (isIframe || versionAtLeast('6.1')) {
      var req_params = {path_full: path_full};
      if (options.force_request) {
        req_params.force_request = true;
      }
      WebView.postEvent('web_app_open_tg_link', false, req_params);
    } else {
      location.href = 'https://t.me' + path_full;
    }
  };
  WebApp.openInvoice = function (url, callback) {
    var a = document.createElement('A'), match, slug;
    a.href = url;
    if (a.protocol != 'http:' &&
        a.protocol != 'https:' ||
        !isTmeHostname(a.hostname) ||
        !(match = a.pathname.match(/^\/(\$|invoice\/)([A-Za-z0-9\-_=]+)$/)) ||
        !(slug = match[2])) {
      console.error('[Telegram.WebApp] Invoice url is invalid', url);
      throw Error('WebAppInvoiceUrlInvalid');
    }
    if (!versionAtLeast('6.1')) {
      console.error('[Telegram.WebApp] Method openInvoice is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (webAppInvoices[slug]) {
      console.error('[Telegram.WebApp] Invoice is already opened');
      throw Error('WebAppInvoiceOpened');
    }
    webAppInvoices[slug] = {
      url: url,
      callback: callback
    };
    WebView.postEvent('web_app_open_invoice', false, {slug: slug});
  };
  WebApp.showPopup = function (params, callback) {
    if (!versionAtLeast('6.2')) {
      console.error('[Telegram.WebApp] Method showPopup is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (webAppPopupOpened) {
      console.error('[Telegram.WebApp] Popup is already opened');
      throw Error('WebAppPopupOpened');
    }
    var title = '';
    var message = '';
    var buttons = [];
    var popup_buttons = {};
    var popup_params = {};
    if (typeof params.title !== 'undefined') {
      title = strTrim(params.title);
      if (title.length > 64) {
        console.error('[Telegram.WebApp] Popup title is too long', title);
        throw Error('WebAppPopupParamInvalid');
      }
      if (title.length > 0) {
        popup_params.title = title;
      }
    }
    if (typeof params.message !== 'undefined') {
      message = strTrim(params.message);
    }
    if (!message.length) {
      console.error('[Telegram.WebApp] Popup message is required', params.message);
      throw Error('WebAppPopupParamInvalid');
    }
    if (message.length > 256) {
      console.error('[Telegram.WebApp] Popup message is too long', message);
      throw Error('WebAppPopupParamInvalid');
    }
    popup_params.message = message;
    if (typeof params.buttons !== 'undefined') {
      if (!Array.isArray(params.buttons)) {
        console.error('[Telegram.WebApp] Popup buttons should be an array', params.buttons);
        throw Error('WebAppPopupParamInvalid');
      }
      for (var i = 0; i < params.buttons.length; i++) {
        var button = params.buttons[i];
        var btn = {};
        var id = '';
        if (typeof button.id !== 'undefined') {
          id = button.id.toString();
          if (id.length > 64) {
            console.error('[Telegram.WebApp] Popup button id is too long', id);
            throw Error('WebAppPopupParamInvalid');
          }
        }
        btn.id = id;
        var button_type = button.type;
        if (typeof button_type === 'undefined') {
          button_type = 'default';
        }
        btn.type = button_type;
        if (button_type == 'ok' ||
            button_type == 'close' ||
            button_type == 'cancel') {
          // no params needed
        } else if (button_type == 'default' ||
                   button_type == 'destructive') {
          var text = '';
          if (typeof button.text !== 'undefined') {
            text = strTrim(button.text);
          }
          if (!text.length) {
            console.error('[Telegram.WebApp] Popup button text is required for type ' + button_type, button.text);
            throw Error('WebAppPopupParamInvalid');
          }
          if (text.length > 64) {
            console.error('[Telegram.WebApp] Popup button text is too long', text);
            throw Error('WebAppPopupParamInvalid');
          }
          btn.text = text;
        } else {
          console.error('[Telegram.WebApp] Popup button type is invalid', button_type);
          throw Error('WebAppPopupParamInvalid');
        }
        buttons.push(btn);
      }
    } else {
      buttons.push({id: '', type: 'close'});
    }
    if (buttons.length < 1) {
      console.error('[Telegram.WebApp] Popup should have at least one button');
      throw Error('WebAppPopupParamInvalid');
    }
    if (buttons.length > 3) {
      console.error('[Telegram.WebApp] Popup should not have more than 3 buttons');
      throw Error('WebAppPopupParamInvalid');
    }
    popup_params.buttons = buttons;

    webAppPopupOpened = {
      callback: callback
    };
    WebView.postEvent('web_app_open_popup', false, popup_params);
  };
  WebApp.showAlert = function (message, callback) {
    WebApp.showPopup({
      message: message
    }, callback ? function(){ callback(); } : null);
  };
  WebApp.showConfirm = function (message, callback) {
    WebApp.showPopup({
      message: message,
      buttons: [
        {type: 'ok', id: 'ok'},
        {type: 'cancel'}
      ]
    }, callback ? function (button_id) {
      callback(button_id == 'ok');
    } : null);
  };
  WebApp.showScanQrPopup = function (params, callback) {
    if (!versionAtLeast('6.4')) {
      console.error('[Telegram.WebApp] Method showScanQrPopup is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (webAppScanQrPopupOpened) {
      console.error('[Telegram.WebApp] Popup is already opened');
      throw Error('WebAppScanQrPopupOpened');
    }
    var text = '';
    var popup_params = {};
    if (typeof params.text !== 'undefined') {
      text = strTrim(params.text);
      if (text.length > 64) {
        console.error('[Telegram.WebApp] Scan QR popup text is too long', text);
        throw Error('WebAppScanQrPopupParamInvalid');
      }
      if (text.length > 0) {
        popup_params.text = text;
      }
    }

    webAppScanQrPopupOpened = {
      callback: callback
    };
    WebView.postEvent('web_app_open_scan_qr_popup', false, popup_params);
  };
  WebApp.closeScanQrPopup = function () {
    if (!versionAtLeast('6.4')) {
      console.error('[Telegram.WebApp] Method closeScanQrPopup is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }

    webAppScanQrPopupOpened = false;
    WebView.postEvent('web_app_close_scan_qr_popup', false);
  };
  WebApp.readTextFromClipboard = function (callback) {
    if (!versionAtLeast('6.4')) {
      console.error('[Telegram.WebApp] Method readTextFromClipboard is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    var req_id = generateCallbackId(16);
    var req_params = {req_id: req_id};
    webAppCallbacks[req_id] = {
      callback: callback
    };
    WebView.postEvent('web_app_read_text_from_clipboard', false, req_params);
  };
  WebApp.requestWriteAccess = function (callback) {
    if (!versionAtLeast('6.9')) {
      console.error('[Telegram.WebApp] Method requestWriteAccess is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (WebAppWriteAccessRequested) {
      console.error('[Telegram.WebApp] Write access is already requested');
      throw Error('WebAppWriteAccessRequested');
    }
    WebAppWriteAccessRequested = {
      callback: callback
    };
    WebView.postEvent('web_app_request_write_access');
  };
  WebApp.requestContact = function (callback) {
    if (!versionAtLeast('6.9')) {
      console.error('[Telegram.WebApp] Method requestContact is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (WebAppContactRequested) {
      console.error('[Telegram.WebApp] Contact is already requested');
      throw Error('WebAppContactRequested');
    }
    WebAppContactRequested = {
      callback: callback
    };
    WebView.postEvent('web_app_request_phone');
  };
  WebApp.downloadFile = function (params, callback) {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method downloadFile is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (webAppDownloadFileRequested) {
      console.error('[Telegram.WebApp] Popup is already opened');
      throw Error('WebAppDownloadFilePopupOpened');
    }
    var a = document.createElement('A');

    var dl_params = {};
    if (!params || !params.url || !params.url.length) {
      console.error('[Telegram.WebApp] Url is required');
      throw Error('WebAppDownloadFileParamInvalid');
    }
    a.href = params.url;
    if (a.protocol != 'https:') {
      console.error('[Telegram.WebApp] Url protocol is not supported', url);
      throw Error('WebAppDownloadFileParamInvalid');
    }
    dl_params.url = a.href;

    if (!params || !params.file_name || !params.file_name.length) {
      console.error('[Telegram.WebApp] File name is required');
      throw Error('WebAppDownloadFileParamInvalid');
    }
    dl_params.file_name = params.file_name;

    webAppDownloadFileRequested = {
      callback: callback
    };
    WebView.postEvent('web_app_request_file_download', false, dl_params);
  };
  WebApp.shareToStory = function (media_url, params) {
    params = params || {};
    if (!versionAtLeast('7.8')) {
      console.error('[Telegram.WebApp] Method shareToStory is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    var a = document.createElement('A');
    a.href = media_url;
    if (a.protocol != 'http:' &&
        a.protocol != 'https:') {
      console.error('[Telegram.WebApp] Media url protocol is not supported', url);
      throw Error('WebAppMediaUrlInvalid');
    }
    var share_params = {};
    share_params.media_url = a.href;
    if (typeof params.text !== 'undefined') {
      var text = strTrim(params.text);
      if (text.length > 2048) {
        console.error('[Telegram.WebApp] Text is too long', text);
        throw Error('WebAppShareToStoryParamInvalid');
      }
      if (text.length > 0) {
        share_params.text = text;
      }
    }
    if (typeof params.widget_link !== 'undefined') {
      params.widget_link = params.widget_link || {};
      a.href = params.widget_link.url;
      if (a.protocol != 'http:' &&
          a.protocol != 'https:') {
        console.error('[Telegram.WebApp] Link protocol is not supported', url);
        throw Error('WebAppShareToStoryParamInvalid');
      }
      var widget_link = {
        url: a.href
      };
      if (typeof params.widget_link.name !== 'undefined') {
        var link_name = strTrim(params.widget_link.name);
        if (link_name.length > 48) {
          console.error('[Telegram.WebApp] Link name is too long', link_name);
          throw Error('WebAppShareToStoryParamInvalid');
        }
        if (link_name.length > 0) {
          widget_link.name = link_name;
        }
      }
      share_params.widget_link = widget_link;
    }

    WebView.postEvent('web_app_share_to_story', false, share_params);
  };
  WebApp.shareMessage = function (msg_id, callback) {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method shareMessage is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (WebAppShareMessageOpened) {
      console.error('[Telegram.WebApp] Share message is already opened');
      throw Error('WebAppShareMessageOpened');
    }
    WebAppShareMessageOpened = {
      callback: callback
    };
    WebView.postEvent('web_app_send_prepared_message', false, {id: msg_id});
  };
  WebApp.requestChat = function (req_id, callback) {
    if (!versionAtLeast('9.6')) {
      console.error('[Telegram.WebApp] Method requestChat is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (WebAppRequestChatOpened) {
      console.error('[Telegram.WebApp] Request chat is already opened');
      throw Error('WebAppRequestChatOpened');
    }
    WebAppRequestChatOpened = {
      callback: callback
    };
    WebView.postEvent('web_app_request_chat', false, {req_id: req_id});
  };
  WebApp.setEmojiStatus = function (custom_emoji_id, params, callback) {
    params = params || {};
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method setEmojiStatus is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    var status_params = {};
    status_params.custom_emoji_id = custom_emoji_id;
    if (typeof params.duration !== 'undefined') {
      status_params.duration = params.duration;
    }
    if (WebAppEmojiStatusRequested) {
      console.error('[Telegram.WebApp] Emoji status is already requested');
      throw Error('WebAppEmojiStatusRequested');
    }
    WebAppEmojiStatusRequested = {
      callback: callback
    };
    WebView.postEvent('web_app_set_emoji_status', false, status_params);
  };
  WebApp.requestEmojiStatusAccess = function (callback) {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method requestEmojiStatusAccess is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (WebAppEmojiStatusAccessRequested) {
      console.error('[Telegram.WebApp] Emoji status permission is already requested');
      throw Error('WebAppEmojiStatusAccessRequested');
    }
    WebAppEmojiStatusAccessRequested = {
      callback: callback
    };
    WebView.postEvent('web_app_request_emoji_status_access');
  };
  WebApp.invokeCustomMethod = function (method, params, callback) {
    invokeCustomMethod(method, params, callback);
  };
  WebApp.hideKeyboard = function () {
    WebView.postEvent('web_app_hide_keyboard');
  };
  WebApp.ready = function () {
    WebView.postEvent('web_app_ready');
  };
  WebApp.expand = function () {
    WebView.postEvent('web_app_expand');
  };
  WebApp.close = function (options) {
    options = options || {};
    var req_params = {};
    if (versionAtLeast('7.6') && options.return_back) {
      req_params.return_back = true;
    }
    WebView.postEvent('web_app_close', false, req_params);
  };

  window.Telegram.WebApp = WebApp;

  updateHeaderColor();
  updateBackgroundColor();
  updateBottomBarColor();
  setViewportHeight();
  if (initParams.tgWebAppShowSettings) {
    SettingsButton.show();
  }

  window.addEventListener('resize', onWindowResize);
  if (isIframe) {
    document.addEventListener('click', linkHandler);
  }

  WebView.onEvent('theme_changed', onThemeChanged);
  WebView.onEvent('viewport_changed', onViewportChanged);
  WebView.onEvent('safe_area_changed', onSafeAreaChanged);
  WebView.onEvent('content_safe_area_changed', onContentSafeAreaChanged);
  WebView.onEvent('visibility_changed', onVisibilityChanged);
  WebView.onEvent('invoice_closed', onInvoiceClosed);
  WebView.onEvent('popup_closed', onPopupClosed);
  WebView.onEvent('qr_text_received', onQrTextReceived);
  WebView.onEvent('scan_qr_popup_closed', onScanQrPopupClosed);
  WebView.onEvent('clipboard_text_received', onClipboardTextReceived);
  WebView.onEvent('write_access_requested', onWriteAccessRequested);
  WebView.onEvent('phone_requested', onPhoneRequested);
  WebView.onEvent('file_download_requested', onFileDownloadRequested);
  WebView.onEvent('custom_method_invoked', onCustomMethodInvoked);
  WebView.onEvent('fullscreen_changed', onFullscreenChanged);
  WebView.onEvent('fullscreen_failed', onFullscreenFailed);
  WebView.onEvent('home_screen_added', onHomeScreenAdded);
  WebView.onEvent('home_screen_checked', onHomeScreenChecked);
  WebView.onEvent('prepared_message_sent', onPreparedMessageSent);
  WebView.onEvent('prepared_message_failed', onPreparedMessageFailed);
  WebView.onEvent('requested_chat_sent', onRequestedChatSent);
  WebView.onEvent('requested_chat_failed', onRequestedChatFailed);
  WebView.onEvent('emoji_status_set', onEmojiStatusSet);
  WebView.onEvent('emoji_status_failed', onEmojiStatusFailed);
  WebView.onEvent('emoji_status_access_requested', onEmojiStatusAccessRequested);
  WebView.postEvent('web_app_request_theme');
  WebView.postEvent('web_app_request_viewport');
  WebView.postEvent('web_app_request_safe_area');
  WebView.postEvent('web_app_request_content_safe_area');

})();
data:application/json;charset=utf-8,%5B%0A%20%20%7B%0A%20%20%20%20%22action%22%3A%20%7B%0A%20%20%20%20%20%20%22type%22%3A%20%22block%22%0A%20%20%20%20%7D%2C%0A%20%20%20%20%22condition%22%3A%20%7B%0A%20%20%20%20%20%20%22domainType%22%3A%20%22thirdParty%22%2C%0A%20%20%20%20%20%20%22excludedInitiatorDomains%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%22edu.ag%22%2C%0A%20%20%20%20%20%20%20%20%22edu.by%22%2C%0A%20%20%20%20%20%20%20%20%22edu.bz%22%2C%0A%20%20%20%20%20%20%20%20%22edu.ge%22%2C%0A%20%20%20%20%20%20%20%20%22edu.in%22%2C%0A%20%20%20%20%20%20%20%20%22edu.md%22%2C%0A%20%20%20%20%20%20%20%20%22edu.me%22%2C%0A%20%20%20%20%20%20%20%20%22edu.pe%22%2C%0A%20%20%20%20%20%20%20%20%22edu.pk%22%2C%0A%20%20%20%20%20%20%20%20%22edu.so%22%2C%0A%20%20%20%20%20%20%20%20%22edu.ws%22%2C%0A%20%20%20%20%20%20%20%20%22gov.ag%22%2C%0A%20%20%20%20%20%20%20%20%22gov.by%22%2C%0A%20%20%20%20%20%20%20%20%22gov.bz%22%2C%0A%20%20%20%20%20%20%20%20%22gov.ge%22%2C%0A%20%20%20%20%20%20%20%20%22gov.in%22%2C%0A%20%20%20%20%20%20%20%20%22gov.md%22%2C%0A%20%20%20%20%20%20%20%20%22gov.me%22%2C%0A%20%20%20%20%20%20%20%20%22gov.pk%22%2C%0A%20%20%20%20%20%20%20%20%22gov.sh%22%2C%0A%20%20%20%20%20%20%20%20%22gov.so%22%2C%0A%20%20%20%20%20%20%20%20%22gov.to%22%2C%0A%20%20%20%20%20%20%20%20%22gov.tv%22%2C%0A%20%20%20%20%20%20%20%20%22gov.ws%22%2C%0A%20%20%20%20%20%20%20%20%22palaugov.pw%22%2C%0A%20%20%20%20%20%20%20%20%22proton.me%22%2C%0A%20%20%20%20%20%20%20%20%22twitch.tv%22%0A%20%20%20%20%20%20%5D%2C%0A%20%20%20%20%20%20%22initiatorDomains%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%22123av.com%22%2C%0A%20%20%20%20%20%20%20%20%22141jav.com%22%2C%0A%20%20%20%20%20%20%20%20%22186.2.175.5%22%2C%0A%20%20%20%20%20%20%20%20%221piecemanga.com%22%2C%0A%20%20%20%20%20%20%20%20%224chanarchives.com%22%2C%0A%20%20%20%20%20%20%20%20%2269shux.co%22%2C%0A%20%20%20%20%20%20%20%20%2285po.com%22%2C%0A%20%20%20%20%20%20%20%20%22ad%22%2C%0A%20%20%20%20%20%20%20%20%22adrianmissionminute.com%22%2C%0A%20%20%20%20%20%20%20%20%22ag%22%2C%0A%20%20%20%20%20%20%20%20%22aguea.net%22%2C%0A%20%20%20%20%20%20%20%20%22alejandrocenturyoil.com%22%2C%0A%20%20%20%20%20%20%20%20%22alleneconomicmatter.com%22%2C%0A%20%20%20%20%20%20%20%20%22animepelix.net%22%2C%0A%20%20%20%20%20%20%20%20%22animerco.org%22%2C%0A%20%20%20%20%20%20%20%20%22anizle.org%22%2C%0A%20%20%20%20%20%20%20%20%22apkadmin.com%22%2C%0A%20%20%20%20%20%20%20%20%22app%22%2C%0A%20%20%20%20%20%20%20%20%22arakpop.net%22%2C%0A%20%20%20%20%20%20%20%20%22asiaflix.net%22%2C%0A%20%20%20%20%20%20%20%20%22bet%22%2C%0A%20%20%20%20%20%20%20%20%22bethshouldercan.com%22%2C%0A%20%20%20%20%20%20%20%20%22bigwarp.art%22%2C%0A%20%20%20%20%20%20%20%20%22biz%22%2C%0A%20%20%20%20%20%20%20%20%22blogspot.com%22%2C%0A%20%20%20%20%20%20%20%20%22blue%22%2C%0A%20%20%20%20%20%20%20%20%22boo%22%2C%0A%20%20%20%20%20%20%20%20%22bookfrom.net%22%2C%0A%20%20%20%20%20%20%20%20%22brittneystandardwestern.com%22%2C%0A%20%20%20%20%20%20%20%20%22brobokep.org%22%2C%0A%20%20%20%20%20%20%20%20%22brookethoughi.com%22%2C%0A%20%20%20%20%20%20%20%20%22brucevotewithin.com%22%2C%0A%20%20%20%20%20%20%20%20%22bryantenunder.com%22%2C%0A%20%20%20%20%20%20%20%20%22by%22%2C%0A%20%20%20%20%20%20%20%20%22bz%22%2C%0A%20%20%20%20%20%20%20%20%22care%22%2C%0A%20%20%20%20%20%20%20%20%22cartoony.net%22%2C%0A%20%20%20%20%20%20%20%20%22casa%22%2C%0A%20%20%20%20%20%20%20%20%22caseyimpactstation.com%22%2C%0A%20%20%20%20%20%20%20%20%22cash%22%2C%0A%20%20%20%20%20%20%20%20%22cc%22%2C%0A%20%20%20%20%20%20%20%20%22cdn9mc.com%22%2C%0A%20%20%20%20%20%20%20%20%22ceo%22%2C%0A%20%20%20%20%20%20%20%20%22cfd%22%2C%0A%20%20%20%20%20%20%20%20%22charlessheimprove.com%22%2C%0A%20%20%20%20%20%20%20%20%22charlestoughrace.com%22%2C%0A%20%20%20%20%20%20%20%20%22christopheruntilpoint.com%22%2C%0A%20%20%20%20%20%20%20%20%22cindyeyefinal.com%22%2C%0A%20%20%20%20%20%20%20%20%22click%22%2C%0A%20%20%20%20%20%20%20%20%22clicknupload.cfd%22%2C%0A%20%20%20%20%20%20%20%20%22cloud%22%2C%0A%20%20%20%20%20%20%20%20%22club%22%2C%0A%20%20%20%20%20%20%20%20%22com.se%22%2C%0A%20%20%20%20%20%20%20%20%22comohoy.com%22%2C%0A%20%20%20%20%20%20%20%20%22cx%22%2C%0A%20%20%20%20%20%20%20%20%22cyou%22%2C%0A%20%20%20%20%20%20%20%20%22d-pdf.com%22%2C%0A%20%20%20%20%20%20%20%20%22dad%22%2C%0A%20%20%20%20%20%20%20%20%22datavaults.co%22%2C%0A%20%20%20%20%20%20%20%20%22day%22%2C%0A%20%20%20%20%20%20%20%20%22desiupload.co%22%2C%0A%20%20%20%20%20%20%20%20%22dev%22%2C%0A%20%20%20%20%20%20%20%20%22dianaavoidthey.com%22%2C%0A%20%20%20%20%20%20%20%20%22diananatureforeign.com%22%2C%0A%20%20%20%20%20%20%20%20%22donaldlineelse.com%22%2C%0A%20%20%20%20%20%20%20%20%22download%22%2C%0A%20%20%20%20%20%20%20%20%22downsub.com%22%2C%0A%20%20%20%20%20%20%20%20%22ellenpoliticalfollow.com%22%2C%0A%20%20%20%20%20%20%20%20%22ericeastweight.com%22%2C%0A%20%20%20%20%20%20%20%20%22erikcoldperson.com%22%2C%0A%20%20%20%20%20%20%20%20%22europixhd-official.com%22%2C%0A%20%20%20%20%20%20%20%20%22evelynthankregion.com%22%2C%0A%20%20%20%20%20%20%20%20%22farm%22%2C%0A%20%20%20%20%20%20%20%20%22filmnudes.com%22%2C%0A%20%20%20%20%20%20%20%20%22filmpornoitaliano.org%22%2C%0A%20%20%20%20%20%20%20%20%22flacdownloader.com%22%2C%0A%20%20%20%20%20%20%20%20%22fmovies24-to.com%22%2C%0A%20%20%20%20%20%20%20%20%22forum%22%2C%0A%20%20%20%20%20%20%20%20%22fsportshd.net%22%2C%0A%20%20%20%20%20%20%20%20%22fun%22%2C%0A%20%20%20%20%20%20%20%20%22fztvseries.ng%22%2C%0A%20%20%20%20%20%20%20%20%22garylargeavailable.com%22%2C%0A%20%20%20%20%20%20%20%20%22gdn%22%2C%0A%20%20%20%20%20%20%20%20%22ge%22%2C%0A%20%20%20%20%20%20%20%20%22graceaddresscommunity.com%22%2C%0A%20%20%20%20%20%20%20%20%22hair%22%2C%0A%20%20%20%20%20%20%20%20%22hdpornflix.com%22%2C%0A%20%20%20%20%20%20%20%20%22heatherdiscussionwhen.com%22%2C%0A%20%20%20%20%20%20%20%20%22heatherwholeinvolve.com%22%2C%0A%20%20%20%20%20%20%20%20%22hentai2read.com%22%2C%0A%20%20%20%20%20%20%20%20%22hianime.nz%22%2C%0A%20%20%20%20%20%20%20%20%22homes%22%2C%0A%20%20%20%20%20%20%20%20%22host%22%2C%0A%20%20%20%20%20%20%20%20%22housecardsummerbutton.com%22%2C%0A%20%20%20%20%20%20%20%20%22how%22%2C%0A%20%20%20%20%20%20%20%20%22hubdrive.space%22%2C%0A%20%20%20%20%20%20%20%20%22ianrequireadult.com%22%2C%0A%20%20%20%20%20%20%20%20%22icu%22%2C%0A%20%20%20%20%20%20%20%20%22in%22%2C%0A%20%20%20%20%20%20%20%20%22indi-share.com%22%2C%0A%20%20%20%20%20%20%20%20%22info%22%2C%0A%20%20%20%20%20%20%20%20%22ironhentai.com%22%2C%0A%20%20%20%20%20%20%20%20%22is%22%2C%0A%20%20%20%20%20%20%20%20%22isekaitube.com%22%2C%0A%20%20%20%20%20%20%20%20%22jamessoundcost.com%22%2C%0A%20%20%20%20%20%20%20%20%22jamiesamewalk.com%22%2C%0A%20%20%20%20%20%20%20%20%22japaneseholes.com%22%2C%0A%20%20%20%20%20%20%20%20%22jasminetesttry.com%22%2C%0A%20%20%20%20%20%20%20%20%22jasonresponsemeasure.com%22%2C%0A%20%20%20%20%20%20%20%20%22javboys.com%22%2C%0A%20%20%20%20%20%20%20%20%22javbraze.com%22%2C%0A%20%20%20%20%20%20%20%20%22javhdporn.net%22%2C%0A%20%20%20%20%20%20%20%20%22jayservicestuff.com%22%2C%0A%20%20%20%20%20%20%20%20%22jeanprofessorcentral.com%22%2C%0A%20%20%20%20%20%20%20%20%22jefferycontrolmodel.com%22%2C%0A%20%20%20%20%20%20%20%20%22jennifereconomicgive.com%22%2C%0A%20%20%20%20%20%20%20%20%22jessicachoosemake.com%22%2C%0A%20%20%20%20%20%20%20%20%22jessicaclearout.com%22%2C%0A%20%20%20%20%20%20%20%20%22jessicaglassauthor.com%22%2C%0A%20%20%20%20%20%20%20%20%22jessicayeahcatch.com%22%2C%0A%20%20%20%20%20%20%20%20%22jgvdata.com%22%2C%0A%20%20%20%20%20%20%20%20%22jilliandescribecompany.com%22%2C%0A%20%20%20%20%20%20%20%20%22johnbeyondnation.com%22%2C%0A%20%20%20%20%20%20%20%20%22jonathansociallike.com%22%2C%0A%20%20%20%20%20%20%20%20%22josephseveralconcern.com%22%2C%0A%20%20%20%20%20%20%20%20%22jpopsingles.eu%22%2C%0A%20%20%20%20%20%20%20%20%22juliewomanwish.com%22%2C%0A%20%20%20%20%20%20%20%20%22kathleenmemberhistory.com%22%2C%0A%20%20%20%20%20%20%20%20%22kathyinformationwhether.com%22%2C%0A%20%20%20%20%20%20%20%20%22kellywhatcould.com%22%2C%0A%20%20%20%20%20%20%20%20%22kennethofficialitem.com%22%2C%0A%20%20%20%20%20%20%20%20%22kisshentaiz.com%22%2C%0A%20%20%20%20%20%20%20%20%22koreaboo.net%22%2C%0A%20%20%20%20%20%20%20%20%22koshalworld.com%22%2C%0A%20%20%20%20%20%20%20%20%22kristiesoundsimply.com%22%2C%0A%20%20%20%20%20%20%20%20%22lancewhosedifficult.com%22%2C%0A%20%20%20%20%20%20%20%20%22lat%22%2C%0A%20%20%20%20%20%20%20%20%22latinluchas.com%22%2C%0A%20%20%20%20%20%20%20%20%22lauradaydo.com%22%2C%0A%20%20%20%20%20%20%20%20%22lindalastattack.com%22%2C%0A%20%20%20%20%20%20%20%20%22link%22%2C%0A%20%20%20%20%20%20%20%20%22lisatrialidea.com%22%2C%0A%20%20%20%20%20%20%20%20%22live%22%2C%0A%20%20%20%20%20%20%20%20%22loadedfiles.org%22%2C%0A%20%20%20%20%20%20%20%20%22loiflix.shop%22%2C%0A%20%20%20%20%20%20%20%20%22lol%22%2C%0A%20%20%20%20%20%20%20%20%22lorimuchbenefit.com%22%2C%0A%20%20%20%20%20%20%20%20%22loriwithinfamily.com%22%2C%0A%20%20%20%20%20%20%20%20%22ltd%22%2C%0A%20%20%20%20%20%20%20%20%22lukecomparetwo.com%22%2C%0A%20%20%20%20%20%20%20%20%22lukesitturn.com%22%2C%0A%20%20%20%20%20%20%20%20%22madouqu.com%22%2C%0A%20%20%20%20%20%20%20%20%22manhwa68.com%22%2C%0A%20%20%20%20%20%20%20%20%22mariatheserepublican.com%22%2C%0A%20%20%20%20%20%20%20%20%22marissasharecareer.com%22%2C%0A%20%20%20%20%20%20%20%20%22markstyleall.com%22%2C%0A%20%20%20%20%20%20%20%20%22maryspecialwatch.com%22%2C%0A%20%20%20%20%20%20%20%20%22matthewhotelscience.com%22%2C%0A%20%20%20%20%20%20%20%20%22md%22%2C%0A%20%20%20%20%20%20%20%20%22me%22%2C%0A%20%20%20%20%20%20%20%20%22media%22%2C%0A%20%20%20%20%20%20%20%20%22media.cm%22%2C%0A%20%20%20%20%20%20%20%20%22megadiscografiascompletas.com%22%2C%0A%20%20%20%20%20%20%20%20%22michaelapplysome.com%22%2C%0A%20%20%20%20%20%20%20%20%22mikaylaarealike.com%22%2C%0A%20%20%20%20%20%20%20%20%22mobi%22%2C%0A%20%20%20%20%20%20%20%20%22moe%22%2C%0A%20%20%20%20%20%20%20%20%22monoschino2.com%22%2C%0A%20%20%20%20%20%20%20%20%22monster%22%2C%0A%20%20%20%20%20%20%20%20%22morganoperationface.com%22%2C%0A%20%20%20%20%20%20%20%20%22movielysium.net%22%2C%0A%20%20%20%20%20%20%20%20%22moviemaniak.com%22%2C%0A%20%20%20%20%20%20%20%20%22movies7.us%22%2C%0A%20%20%20%20%20%20%20%20%22myxxgirl.com%22%2C%0A%20%20%20%20%20%20%20%20%22nekopoi.care%22%2C%0A%20%20%20%20%20%20%20%20%22nicolehappyoutside.com%22%2C%0A%20%20%20%20%20%20%20%20%22niganpro.com%22%2C%0A%20%20%20%20%20%20%20%20%22niraw.com%22%2C%0A%20%20%20%20%20%20%20%20%22njav.com%22%2C%0A%20%20%20%20%20%20%20%20%22notunmovie.net%22%2C%0A%20%20%20%20%20%20%20%20%22novizer.com%22%2C%0A%20%20%20%20%20%20%20%20%22nsfwgify.com%22%2C%0A%20%20%20%20%20%20%20%20%22official.coomer.com.co%22%2C%0A%20%20%20%20%20%20%20%20%22omegascans.org%22%2C%0A%20%20%20%20%20%20%20%20%22one%22%2C%0A%20%20%20%20%20%20%20%20%22onejav.com%22%2C%0A%20%20%20%20%20%20%20%20%22onepiece-manga-online.net%22%2C%0A%20%20%20%20%20%20%20%20%22onl%22%2C%0A%20%20%20%20%20%20%20%20%22online%22%2C%0A%20%20%20%20%20%20%20%20%22pagalfree.com%22%2C%0A%20%20%20%20%20%20%20%20%22pagalnew.com%22%2C%0A%20%20%20%20%20%20%20%20%22pagalworld.us%22%2C%0A%20%20%20%20%20%20%20%20%22page%22%2C%0A%20%20%20%20%20%20%20%20%22pamelachangemission.com%22%2C%0A%20%20%20%20%20%20%20%20%22pe%22%2C%0A%20%20%20%20%20%20%20%20%22pelimeli.com%22%2C%0A%20%20%20%20%20%20%20%20%22pelis182.net%22%2C%0A%20%20%20%20%20%20%20%20%22phenomenalityuniform.com%22%2C%0A%20%20%20%20%20%20%20%20%22photocalltv.org%22%2C%0A%20%20%20%20%20%20%20%20%22pink%22%2C%0A%20%20%20%20%20%20%20%20%22pixvid.org%22%2C%0A%20%20%20%20%20%20%20%20%22pk%22%2C%0A%20%20%20%20%20%20%20%20%22player.subespanolvip.com%22%2C%0A%20%20%20%20%20%20%20%20%22playmogo.com%22%2C%0A%20%20%20%20%20%20%20%20%22plugincrack.com%22%2C%0A%20%20%20%20%20%20%20%20%22pokopow.com%22%2C%0A%20%20%20%20%20%20%20%20%22poophd.pm%22%2C%0A%20%20%20%20%20%20%20%20%22popcornmovies.org%22%2C%0A%20%20%20%20%20%20%20%20%22pornhoarder.net%22%2C%0A%20%20%20%20%20%20%20%20%22pornincest.net%22%2C%0A%20%20%20%20%20%20%20%20%22pro%22%2C%0A%20%20%20%20%20%20%20%20%22pw%22%2C%0A%20%20%20%20%20%20%20%20%22raw.senmanga.com%22%2C%0A%20%20%20%20%20%20%20%20%22re%22%2C%0A%20%20%20%20%20%20%20%20%22rebeccacostthousand.com%22%2C%0A%20%20%20%20%20%20%20%20%22rebeccapracticeloss.com%22%2C%0A%20%20%20%20%20%20%20%20%22red%22%2C%0A%20%20%20%20%20%20%20%20%22relampagomovies.com%22%2C%0A%20%20%20%20%20%20%20%20%22rest%22%2C%0A%20%20%20%20%20%20%20%20%22richardquestionbuilding.com%22%2C%0A%20%20%20%20%20%20%20%20%22richardsignfish.com%22%2C%0A%20%20%20%20%20%20%20%20%22roberteachfinal.com%22%2C%0A%20%20%20%20%20%20%20%20%22robertordercharacter.com%22%2C%0A%20%20%20%20%20%20%20%20%22robertplacespace.com%22%2C%0A%20%20%20%20%20%20%20%20%22ryanagoinvolve.com%22%2C%0A%20%20%20%20%20%20%20%20%22sandratableother.com%22%2C%0A%20%20%20%20%20%20%20%20%22sandrataxeight.com%22%2C%0A%20%20%20%20%20%20%20%20%22sbs%22%2C%0A%20%20%20%20%20%20%20%20%22seriesgod.com%22%2C%0A%20%20%20%20%20%20%20%20%22sethniceletter.com%22%2C%0A%20%20%20%20%20%20%20%20%22sexiezpix.com%22%2C%0A%20%20%20%20%20%20%20%20%22sh%22%2C%0A%20%20%20%20%20%20%20%20%22shannonpersonalcost.com%22%2C%0A%20%20%20%20%20%20%20%20%22site%22%2C%0A%20%20%20%20%20%20%20%20%22skidrowcodex.net%22%2C%0A%20%20%20%20%20%20%20%20%22skin%22%2C%0A%20%20%20%20%20%20%20%20%22space%22%2C%0A%20%20%20%20%20%20%20%20%22splay.id%22%2C%0A%20%20%20%20%20%20%20%20%22spotidownloader.com%22%2C%0A%20%20%20%20%20%20%20%20%22stevenfamilyedge.com%22%2C%0A%20%20%20%20%20%20%20%20%22stream%22%2C%0A%20%20%20%20%20%20%20%20%22streampoi.com%22%2C%0A%20%20%20%20%20%20%20%20%22su%22%2C%0A%20%20%20%20%20%20%20%20%22surrit.store%22%2C%0A%20%20%20%20%20%20%20%20%22susanhavekeep.com%22%2C%0A%20%20%20%20%20%20%20%20%22sxyprn.com%22%2C%0A%20%20%20%20%20%20%20%20%22tapepops.com%22%2C%0A%20%20%20%20%20%20%20%20%22team%22%2C%0A%20%20%20%20%20%20%20%20%22tel%22%2C%0A%20%20%20%20%20%20%20%20%22thecakeboutiquect.com%22%2C%0A%20%20%20%20%20%20%20%20%22thefantazy.com%22%2C%0A%20%20%20%20%20%20%20%20%22thepiratebay.org%22%2C%0A%20%20%20%20%20%20%20%20%22timmaybealready.com%22%2C%0A%20%20%20%20%20%20%20%20%22to%22%2C%0A%20%20%20%20%20%20%20%20%22today%22%2C%0A%20%20%20%20%20%20%20%20%22toddpartneranimal.com%22%2C%0A%20%20%20%20%20%20%20%20%22tokyomotion.net%22%2C%0A%20%20%20%20%20%20%20%20%22top%22%2C%0A%20%20%20%20%20%20%20%20%22torbobit.net%22%2C%0A%20%20%20%20%20%20%20%20%22torrentgalaxy-official.com%22%2C%0A%20%20%20%20%20%20%20%20%22torrentjogos.com.br%22%2C%0A%20%20%20%20%20%20%20%20%22towheree.com%22%2C%0A%20%20%20%20%20%20%20%20%22tube%22%2C%0A%20%20%20%20%20%20%20%20%22tv%22%2C%0A%20%20%20%20%20%20%20%20%22uniquestream.net%22%2C%0A%20%20%20%20%20%20%20%20%22vickisaveworker.com%22%2C%0A%20%20%20%20%20%20%20%20%22vidboys.com%22%2C%0A%20%20%20%20%20%20%20%20%22vidsonic.net%22%2C%0A%20%20%20%20%20%20%20%20%22viduyy.com%22%2C%0A%20%20%20%20%20%20%20%20%22voe.sx%22%2C%0A%20%20%20%20%20%20%20%20%22walterprettytheir.com%22%2C%0A%20%20%20%20%20%20%20%20%22wedding%22%2C%0A%20%20%20%20%20%20%20%20%22world%22%2C%0A%20%20%20%20%20%20%20%20%22ws%22%2C%0A%20%20%20%20%20%20%20%20%22xfantazy.com%22%2C%0A%20%20%20%20%20%20%20%20%22xmegadrive.com%22%2C%0A%20%20%20%20%20%20%20%20%22xyz%22%0A%20%20%20%20%20%20%5D%2C%0A%20%20%20%20%20%20%22isUrlFilterCaseSensitive%22%3A%20true%2C%0A%20%20%20%20%20%20%22regexFilter%22%3A%20%22%5C%5C%2F%5B0-9a-z%5D%7B2%7D%5C%5C%2F%5B0-9a-f%5D%7B32%7D%5C%5C.js%22%2C%0A%20%20%20%20%20%20%22requestDomains%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%22com%22%0A%20%20%20%20%20%20%5D%2C%0A%20%20%20%20%20%20%22requestMethods%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%22get%22%0A%20%20%20%20%20%20%5D%2C%0A%20%20%20%20%20%20%22resourceTypes%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%22script%22%2C%0A%20%20%20%20%20%20%20%20%22xmlhttprequest%22%0A%20%20%20%20%20%20%5D%0A%20%20%20%20%7D%2C%0A%20%20%20%20%22id%22%3A%201%2C%0A%20%20%20%20%22priority%22%3A%2010%0A%20%20%7D%2C%0A%20%20%7B%0A%20%20%20%20%22action%22%3A%20%7B%0A%20%20%20%20%20%20%22type%22%3A%20%22block%22%0A%20%20%20%20%7D%2C%0A%20%20%20%20%22condition%22%3A%20%7B%0A%20%20%20%20%20%20%22domainType%22%3A%20%22thirdParty%22%2C%0A%20%20%20%20%20%20%22excludedRequestDomains%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%22adatoolbar.com%22%2C%0A%20%20%20%20%20%20%20%20%22aswpsdkus.com%22%2C%0A%20%20%20%20%20%20%20%20%22chimpstatic.com%22%2C%0A%20%20%20%20%20%20%20%20%22clickiocmp.com%22%2C%0A%20%20%20%20%20%20%20%20%22mojirater.com%22%2C%0A%20%20%20%20%20%20%20%20%22polldaddy.com%22%2C%0A%20%20%20%20%20%20%20%20%22squareoffs.com%22%2C%0A%20%20%20%20%20%20%20%20%22succeedscene.com%22%2C%0A%20%20%20%20%20%20%20%20%22tallysight.com%22%2C%0A%20%20%20%20%20%20%20%20%22veraviews.com%22%0A%20%20%20%20%20%20%5D%2C%0A%20%20%20%20%20%20%22initiatorDomains%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%2215min.lt%22%2C%0A%20%20%20%20%20%20%20%20%22247sports.com%22%2C%0A%20%20%20%20%20%20%20%20%22abc17news.com%22%2C%0A%20%20%20%20%20%20%20%20%22addictinggames.com%22%2C%0A%20%20%20%20%20%20%20%20%22agrodigital.com%22%2C%0A%20%20%20%20%20%20%20%20%22al.com%22%2C%0A%20%20%20%20%20%20%20%20%22aliontherunblog.com%22%2C%0A%20%20%20%20%20%20%20%20%22allaboutthetea.com%22%2C%0A%20%20%20%20%20%20%20%20%22allmusic.com%22%2C%0A%20%20%20%20%20%20%20%20%22allthingsthrifty.com%22%2C%0A%20%20%20%20%20%20%20%20%22amessagewithabottle.com%22%2C%0A%20%20%20%20%20%20%20%20%22arstechnica.com%22%2C%0A%20%20%20%20%20%20%20%20%22artforum.com%22%2C%0A%20%20%20%20%20%20%20%20%22artnews.com%22%2C%0A%20%20%20%20%20%20%20%20%22athlonsports.com%22%2C%0A%20%20%20%20%20%20%20%20%22audiomack.com%22%2C%0A%20%20%20%20%20%20%20%20%22autoblog.com%22%2C%0A%20%20%20%20%20%20%20%20%22awkward.com%22%2C%0A%20%20%20%20%20%20%20%20%22azcentral.com%22%2C%0A%20%20%20%20%20%20%20%20%22barcablaugranes.com%22%2C%0A%20%20%20%20%20%20%20%20%22barnsleychronicle.com%22%2C%0A%20%20%20%20%20%20%20%20%22bethcakes.com%22%2C%0A%20%20%20%20%20%20%20%20%22betweenenglandandiowa.com%22%2C%0A%20%20%20%20%20%20%20%20%22bgr.com%22%2C%0A%20%20%20%20%20%20%20%20%22bikemag.com%22%2C%0A%20%20%20%20%20%20%20%20%22billboard.com%22%2C%0A%20%20%20%20%20%20%20%20%22blazersedge.com%22%2C%0A%20%20%20%20%20%20%20%20%22blogher.com%22%2C%0A%20%20%20%20%20%20%20%20%22bloxinformer.com%22%2C%0A%20%20%20%20%20%20%20%20%22blu-ray.com%22%2C%0A%20%20%20%20%20%20%20%20%22bluegraygal.com%22%2C%0A%20%20%20%20%20%20%20%20%22boredpanda.com%22%2C%0A%20%20%20%20%20%20%20%20%22briefeguru.de%22%2C%0A%20%20%20%20%20%20%20%20%22brobible.com%22%2C%0A%20%20%20%20%20%20%20%20%22cagesideseats.com%22%2C%0A%20%20%20%20%20%20%20%20%22cbsnews.com%22%2C%0A%20%20%20%20%20%20%20%20%22cbssports.com%22%2C%0A%20%20%20%20%20%20%20%20%22celiacandthebeast.com%22%2C%0A%20%20%20%20%20%20%20%20%22chaptercheats.com%22%2C%0A%20%20%20%20%20%20%20%20%22cheatsheet.com%22%2C%0A%20%20%20%20%20%20%20%20%22cleveland.com%22%2C%0A%20%20%20%20%20%20%20%20%22clickondetroit.com%22%2C%0A%20%20%20%20%20%20%20%20%22coloradoan.com%22%2C%0A%20%20%20%20%20%20%20%20%22commercialobserver.com%22%2C%0A%20%20%20%20%20%20%20%20%22competentedigitale.ro%22%2C%0A%20%20%20%20%20%20%20%20%22cracked.com%22%2C%0A%20%20%20%20%20%20%20%20%22dailycaller.com%22%2C%0A%20%20%20%20%20%20%20%20%22dailydot.com%22%2C%0A%20%20%20%20%20%20%20%20%22dailykos.com%22%2C%0A%20%20%20%20%20%20%20%20%22dailyvoice.com%22%2C%0A%20%20%20%20%20%20%20%20%22danslescoulisses.com%22%2C%0A%20%20%20%20%20%20%20%20%22databazeknih.cz%22%2C%0A%20%20%20%20%20%20%20%20%22deadline.com%22%2C%0A%20%20%20%20%20%20%20%20%22decider.com%22%2C%0A%20%20%20%20%20%20%20%20%22dengarden.com%22%2C%0A%20%20%20%20%20%20%20%20%22dexerto.com%22%2C%0A%20%20%20%20%20%20%20%20%22didyouknowfacts.com%22%2C%0A%20%20%20%20%20%20%20%20%22digitalmusicnews.com%22%2C%0A%20%20%20%20%20%20%20%20%22dogtime.com%22%2C%0A%20%20%20%20%20%20%20%20%22dpreview.com%22%2C%0A%20%20%20%20%20%20%20%20%22dwell.com%22%2C%0A%20%20%20%20%20%20%20%20%22eater.com%22%2C%0A%20%20%20%20%20%20%20%20%22ebaumsworld.com%22%2C%0A%20%20%20%20%20%20%20%20%22egoallstars.com%22%2C%0A%20%20%20%20%20%20%20%20%22eldiariony.com%22%2C%0A%20%20%20%20%20%20%20%20%22fark.com%22%2C%0A%20%20%20%20%20%20%20%20%22femestella.com%22%2C%0A%20%20%20%20%20%20%20%20%22flickr.com%22%2C%0A%20%20%20%20%20%20%20%20%22fmradiofree.com%22%2C%0A%20%20%20%20%20%20%20%20%22forums.hfboards.com%22%2C%0A%20%20%20%20%20%20%20%20%22free-power-point-templates.com%22%2C%0A%20%20%20%20%20%20%20%20%22freeconvert.com%22%2C%0A%20%20%20%20%20%20%20%20%22freep.com%22%2C%0A%20%20%20%20%20%20%20%20%22frogsandsnailsandpuppydogtail.com%22%2C%0A%20%20%20%20%20%20%20%20%22funtasticlife.com%22%2C%0A%20%20%20%20%20%20%20%20%22fwmadebycarli.com%22%2C%0A%20%20%20%20%20%20%20%20%22golfdigest.com%22%2C%0A%20%20%20%20%20%20%20%20%22greenbaypressgazette.com%22%2C%0A%20%20%20%20%20%20%20%20%22grunge.com%22%2C%0A%20%20%20%20%20%20%20%20%22gulflive.com%22%2C%0A%20%20%20%20%20%20%20%20%22hollywoodreporter.com%22%2C%0A%20%20%20%20%20%20%20%20%22homeglowdesign.com%22%2C%0A%20%20%20%20%20%20%20%20%22honeygirlsworld.com%22%2C%0A%20%20%20%20%20%20%20%20%22ibtimes.co.in%22%2C%0A%20%20%20%20%20%20%20%20%22imgur.com%22%2C%0A%20%20%20%20%20%20%20%20%22indiewire.com%22%2C%0A%20%20%20%20%20%20%20%20%22intouchweekly.com%22%2C%0A%20%20%20%20%20%20%20%20%22iphoneincanada.ca%22%2C%0A%20%20%20%20%20%20%20%20%22jasminemaria.com%22%2C%0A%20%20%20%20%20%20%20%20%22jsonline.com%22%2C%0A%20%20%20%20%20%20%20%20%22kens5.com%22%2C%0A%20%20%20%20%20%20%20%20%22kinobox.cz%22%2C%0A%20%20%20%20%20%20%20%20%22kion546.com%22%2C%0A%20%20%20%20%20%20%20%20%22knowyourmeme.com%22%2C%0A%20%20%20%20%20%20%20%20%22komonews.com%22%2C%0A%20%20%20%20%20%20%20%20%22krem.com%22%2C%0A%20%20%20%20%20%20%20%20%22last.fm%22%2C%0A%20%20%20%20%20%20%20%20%22lehighvalleylive.com%22%2C%0A%20%20%20%20%20%20%20%20%22lettyskitchen.com%22%2C%0A%20%20%20%20%20%20%20%20%22lifeandstylemag.com%22%2C%0A%20%20%20%20%20%20%20%20%22lifeinleggings.com%22%2C%0A%20%20%20%20%20%20%20%20%22lizzieinlace.com%22%2C%0A%20%20%20%20%20%20%20%20%22localnews8.com%22%2C%0A%20%20%20%20%20%20%20%20%22lonestarlive.com%22%2C%0A%20%20%20%20%20%20%20%20%22macwelt.de%22%2C%0A%20%20%20%20%20%20%20%20%22macworld.com%22%2C%0A%20%20%20%20%20%20%20%20%22madeeveryday.com%22%2C%0A%20%20%20%20%20%20%20%20%22maidenhead-advertiser.co.uk%22%2C%0A%20%20%20%20%20%20%20%20%22mandatory.com%22%2C%0A%20%20%20%20%20%20%20%20%22mardomreport.net%22%2C%0A%20%20%20%20%20%20%20%20%22masslive.com%22%2C%0A%20%20%20%20%20%20%20%20%22melangery.com%22%2C%0A%20%20%20%20%20%20%20%20%22mensfitness.com%22%2C%0A%20%20%20%20%20%20%20%20%22mensjournal.com%22%2C%0A%20%20%20%20%20%20%20%20%22meteomedia.com%22%2C%0A%20%20%20%20%20%20%20%20%22miamiherald.com%22%2C%0A%20%20%20%20%20%20%20%20%22mlive.com%22%2C%0A%20%20%20%20%20%20%20%20%22mmamania.com%22%2C%0A%20%20%20%20%20%20%20%20%22momtastic.com%22%2C%0A%20%20%20%20%20%20%20%20%22mostlymorgan.com%22%2C%0A%20%20%20%20%20%20%20%20%22motherwellmag.com%22%2C%0A%20%20%20%20%20%20%20%20%22motor1.com%22%2C%0A%20%20%20%20%20%20%20%20%22motorsport.com%22%2C%0A%20%20%20%20%20%20%20%20%22musicfeeds.com.au%22%2C%0A%20%20%20%20%20%20%20%20%22naszemiasto.pl%22%2C%0A%20%20%20%20%20%20%20%20%22nationalpost.com%22%2C%0A%20%20%20%20%20%20%20%20%22nationalreview.com%22%2C%0A%20%20%20%20%20%20%20%20%22nbcsports.com%22%2C%0A%20%20%20%20%20%20%20%20%22news.com.au%22%2C%0A%20%20%20%20%20%20%20%20%22ninersnation.com%22%2C%0A%20%20%20%20%20%20%20%20%22nj.com%22%2C%0A%20%20%20%20%20%20%20%20%22nordot.app%22%2C%0A%20%20%20%20%20%20%20%20%22nothingbutnewcastle.com%22%2C%0A%20%20%20%20%20%20%20%20%22nsjonline.com%22%2C%0A%20%20%20%20%20%20%20%20%22nypost.com%22%2C%0A%20%20%20%20%20%20%20%20%22observer.com%22%2C%0A%20%20%20%20%20%20%20%20%22ocala.com%22%2C%0A%20%20%20%20%20%20%20%20%22ontvtonight.com%22%2C%0A%20%20%20%20%20%20%20%20%22oregonlive.com%22%2C%0A%20%20%20%20%20%20%20%20%22pagesix.com%22%2C%0A%20%20%20%20%20%20%20%20%22palmbeachpost.com%22%2C%0A%20%20%20%20%20%20%20%20%22parade.com%22%2C%0A%20%20%20%20%20%20%20%20%22paradehomeandgarden.com%22%2C%0A%20%20%20%20%20%20%20%20%22paradepets.com%22%2C%0A%20%20%20%20%20%20%20%20%22patheos.com%22%2C%0A%20%20%20%20%20%20%20%20%22pauladeen.com%22%2C%0A%20%20%20%20%20%20%20%20%22pcbolsa.com%22%2C%0A%20%20%20%20%20%20%20%20%22pennlive.com%22%2C%0A%20%20%20%20%20%20%20%20%22pep.ph%22%2C%0A%20%20%20%20%20%20%20%20%22pethelpful.com%22%2C%0A%20%20%20%20%20%20%20%20%22phillyvoice.com%22%2C%0A%20%20%20%20%20%20%20%20%22playstationlifestyle.net%22%2C%0A%20%20%20%20%20%20%20%20%22pnj.com%22%2C%0A%20%20%20%20%20%20%20%20%22powder.com%22%2C%0A%20%20%20%20%20%20%20%20%22pravda.sk%22%2C%0A%20%20%20%20%20%20%20%20%22press-citizen.com%22%2C%0A%20%20%20%20%20%20%20%20%22prydwen.gg%22%2C%0A%20%20%20%20%20%20%20%20%22puckermom.com%22%2C%0A%20%20%20%20%20%20%20%20%22pwinsider.com%22%2C%0A%20%20%20%20%20%20%20%20%22reelmama.com%22%2C%0A%20%20%20%20%20%20%20%20%22reuters.com%22%2C%0A%20%20%20%20%20%20%20%20%22rlfans.com%22%2C%0A%20%20%20%20%20%20%20%20%22robbreport.com%22%2C%0A%20%20%20%20%20%20%20%20%22rollingstone.com%22%2C%0A%20%20%20%20%20%20%20%20%22royalmailchat.co.uk%22%2C%0A%20%20%20%20%20%20%20%20%22sandrarose.com%22%2C%0A%20%20%20%20%20%20%20%20%22sbnation.com%22%2C%0A%20%20%20%20%20%20%20%20%22sheknows.com%22%2C%0A%20%20%20%20%20%20%20%20%22shophq.com%22%2C%0A%20%20%20%20%20%20%20%20%22sidereel.com%22%2C%0A%20%20%20%20%20%20%20%20%22silive.com%22%2C%0A%20%20%20%20%20%20%20%20%22smartworld.it%22%2C%0A%20%20%20%20%20%20%20%20%22sneakernews.com%22%2C%0A%20%20%20%20%20%20%20%20%22sourcingjournal.com%22%2C%0A%20%20%20%20%20%20%20%20%22sport-fm.gr%22%2C%0A%20%20%20%20%20%20%20%20%22sportico.com%22%2C%0A%20%20%20%20%20%20%20%20%22sportsgamblingpodcast.com%22%2C%0A%20%20%20%20%20%20%20%20%22spotofteadesigns.com%22%2C%0A%20%20%20%20%20%20%20%20%22ssnewstelegram.com%22%2C%0A%20%20%20%20%20%20%20%20%22stacysrandomthoughts.com%22%2C%0A%20%20%20%20%20%20%20%20%22stylecaster.com%22%2C%0A%20%20%20%20%20%20%20%20%22superherohype.com%22%2C%0A%20%20%20%20%20%20%20%20%22surfer.com%22%2C%0A%20%20%20%20%20%20%20%20%22syracuse.com%22%2C%0A%20%20%20%20%20%20%20%20%22tablelifeblog.com%22%2C%0A%20%20%20%20%20%20%20%20%22tasteofhome.com%22%2C%0A%20%20%20%20%20%20%20%20%22tastingtable.com%22%2C%0A%20%20%20%20%20%20%20%20%22techcrunch.com%22%2C%0A%20%20%20%20%20%20%20%20%22thecelticblog.com%22%2C%0A%20%20%20%20%20%20%20%20%22thedailymeal.com%22%2C%0A%20%20%20%20%20%20%20%20%22theflowspace.com%22%2C%0A%20%20%20%20%20%20%20%20%22thegatewaypundit.com%22%2C%0A%20%20%20%20%20%20%20%20%22themarysue.com%22%2C%0A%20%20%20%20%20%20%20%20%22thenerdstash.com%22%2C%0A%20%20%20%20%20%20%20%20%22thenerdyme.com%22%2C%0A%20%20%20%20%20%20%20%20%22theprudentgarden.com%22%2C%0A%20%20%20%20%20%20%20%20%22thespun.com%22%2C%0A%20%20%20%20%20%20%20%20%22thestreet.com%22%2C%0A%20%20%20%20%20%20%20%20%22theverge.com%22%2C%0A%20%20%20%20%20%20%20%20%22tiermaker.com%22%2C%0A%20%20%20%20%20%20%20%20%22timesofisrael.com%22%2C%0A%20%20%20%20%20%20%20%20%22tiscali.cz%22%2C%0A%20%20%20%20%20%20%20%20%22tokfm.pl%22%2C%0A%20%20%20%20%20%20%20%20%22torontosun.com%22%2C%0A%20%20%20%20%20%20%20%20%22travelhost.com%22%2C%0A%20%20%20%20%20%20%20%20%22tvline.com%22%2C%0A%20%20%20%20%20%20%20%20%22usatoday.com%22%2C%0A%20%20%20%20%20%20%20%20%22usmagazine.com%22%2C%0A%20%20%20%20%20%20%20%20%22variety.com%22%2C%0A%20%20%20%20%20%20%20%20%22viralviralvideos.com%22%2C%0A%20%20%20%20%20%20%20%20%22wallup.net%22%2C%0A%20%20%20%20%20%20%20%20%22wannacomewith.com%22%2C%0A%20%20%20%20%20%20%20%20%22wcnc.com%22%2C%0A%20%20%20%20%20%20%20%20%22weather.com%22%2C%0A%20%20%20%20%20%20%20%20%22wimp.com%22%2C%0A%20%20%20%20%20%20%20%20%22windsorexpress.co.uk%22%2C%0A%20%20%20%20%20%20%20%20%22woojr.com%22%2C%0A%20%20%20%20%20%20%20%20%22worldoftravelswithkids.com%22%2C%0A%20%20%20%20%20%20%20%20%22worldstar.com%22%2C%0A%20%20%20%20%20%20%20%20%22worldstarhiphop.com%22%2C%0A%20%20%20%20%20%20%20%20%22worldsurfleague.com%22%2C%0A%20%20%20%20%20%20%20%20%22wpgh53.com%22%2C%0A%20%20%20%20%20%20%20%20%22wunderground.com%22%2C%0A%20%20%20%20%20%20%20%20%22wwd.com%22%2C%0A%20%20%20%20%20%20%20%20%22wyborkierowcow.pl%22%2C%0A%20%20%20%20%20%20%20%20%22wzzm13.com%22%2C%0A%20%20%20%20%20%20%20%20%22ydr.com%22%2C%0A%20%20%20%20%20%20%20%20%22yourcountdown.to%22%2C%0A%20%20%20%20%20%20%20%20%22zdnet.com%22%0A%20%20%20%20%20%20%5D%2C%0A%20%20%20%20%20%20%22isUrlFilterCaseSensitive%22%3A%20true%2C%0A%20%20%20%20%20%20%22regexFilter%22%3A%20%22%5Ehttps%3A%5C%5C%2F%5C%5C%2F%5B0-9a-z%5D%7B7%2C25%7D%5C%5C.com%5C%5C%2F%5B_0-9a-z%5C%5C%2F%5D%7B2%7D%22%2C%0A%20%20%20%20%20%20%22requestDomains%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%22com%22%0A%20%20%20%20%20%20%5D%2C%0A%20%20%20%20%20%20%22resourceTypes%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%22script%22%0A%20%20%20%20%20%20%5D%0A%20%20%20%20%7D%2C%0A%20%20%20%20%22id%22%3A%202%2C%0A%20%20%20%20%22priority%22%3A%2010%0A%20%20%7D%2C%0A%20%20%7B%0A%20%20%20%20%22action%22%3A%20%7B%0A%20%20%20%20%20%20%22redirect%22%3A%20%7B%0A%20%20%20%20%20%20%20%20%22extensionPath%22%3A%20%22%2Fweb_accessible_resources%2Fnoop.js%22%0A%20%20%20%20%20%20%7D%2C%0A%20%20%20%20%20%20%22type%22%3A%20%22redirect%22%0A%20%20%20%20%7D%2C%0A%20%20%20%20%22condition%22%3A%20%7B%0A%20%20%20%20%20%20%22domainType%22%3A%20%22thirdParty%22%2C%0A%20%20%20%20%20%20%22excludedRequestDomains%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%22adatoolbar.com%22%2C%0A%20%20%20%20%20%20%20%20%22aswpsdkus.com%22%2C%0A%20%20%20%20%20%20%20%20%22chimpstatic.com%22%2C%0A%20%20%20%20%20%20%20%20%22clickiocmp.com%22%2C%0A%20%20%20%20%20%20%20%20%22mojirater.com%22%2C%0A%20%20%20%20%20%20%20%20%22polldaddy.com%22%2C%0A%20%20%20%20%20%20%20%20%22squareoffs.com%22%2C%0A%20%20%20%20%20%20%20%20%22succeedscene.com%22%2C%0A%20%20%20%20%20%20%20%20%22tallysight.com%22%2C%0A%20%20%20%20%20%20%20%20%22veraviews.com%22%0A%20%20%20%20%20%20%5D%2C%0A%20%20%20%20%20%20%22initiatorDomains%22%3A%20%5B%0A%20%20%20%20%20%20%20%20%2215min.lt%22%2C%0A%20%20%20%20%20%20%20%20%22247sports.com%22%2C%0A%20%20%20%20%20%20%20%20%22abc17news.com%22%2C%0A%20%20%20%20%20%20%20%20%22addictinggames.com%22%2C%0A%20%20%20%20%20%20%20%20%22agrodigital.com%22%2C%0A%20%20%20%20%20%20%20%20%22al.com%22%2C%0A%20%20%20%20%20%20%20%20%22aliontherunblog.com%22%2C%0A%20%20%20%20%20%20%20%20%22allaboutthetea.com%22%2C%0A%20%20%20%20%20%20%20%20%22allmusic.com%22%2C%0A%20%20%20%20%20%20%20%20%22allthingsthrifty.com%22%2C%0A%20%20%20%20%20%20%20%20%22amessagewithabottle.com%22%2C%0A%20%20%20%20%20%20%20%20%22arstechnica.com%22%2C%0A%20%20%20%20%20%20%20%20%22artforum.com%22%2C%0A%20%20%20%20%20%20%20%20%22artnews.com%22%2C%0A%20%20%20%20%20%20%20%20%22athlonsports.com%22%2C%0A%20%20%20%20%20%20%20%20%22audiomack.com%22%2C%0A%20%20%20%20%20%20%20%20%22autoblog.com%22%2C%0A%20%20%20%20%20%20%20%20%22awkward.com%22%2C%0A%20%20%20%20%20%20%20%20%22azcentral.com%22%2C%0A%20%20%20%20%20%20%20%20%22barcablaugranes.com%22%2C%0A%20%20%20%20%20%20%20%20%22barnsleychronicle.com%22%2C%0A%20%20%20%20%20%20%20%20%22bethcakes.com%22%2C%0A%20%20%20%20%20%20%20%20%22betweenenglandandiowa.com%22%2C%0A%20%20%20%20%20%20%20%20%22bgr.com%22%2C%0A%20%20%20%20%20%20%20%20%22bikemag.com%22%2C%0A%20%20%20%20%20%20%20%20%22billboard.com%22%2C%0A%20%20%20%20%20%20%20%20%22blazersedge.com%22%2C%0A%20%20%20%20%20%20%20%20%22blogher.com%22%2C%0A%20%20%20%20%20%20%20%20%22bloxinformer.com%22%2C%0A%20%20%20%20%20%20%20%20%22blu-ray.com%22%2C%0A%20%20%20%20%20%20%20%20%22bluegraygal.com%22%2C%0A%20%20%20%20%20%20%20%20%22boredpanda.com%22%2C%0A%20%20%20%20%20%20%20%20%22briefeguru.de%22%2C%0A%20%20%20%20%20%20%20%20%22brobible.com%22%2C%0A%20%20%20%20%20%20%20%20%22cagesideseats.com%22%2C%0A%20%20%20%20%20%20%20%20%22cbsnews.com%22%2C%0A%20%20%20%20%20%20%20%20%22cbssports.com%22%2C%0A%20%20%20%20%20%20%20%20%22celiacandthebeast.com%22%2C%0A%20%20%20%20%20%20%20%20%22chaptercheats.com%22%2C%0A%20%20%20%20%20%20%20%20%22cheatsheet.com%22%2C%0A%20%20%20%20%20%20%20%20%22cleveland.com%22%2C%0A%20%20%20%20%20%20%20%20%22clickondetroit.com%22%2C%0A%20%20%20%20%20%20%20%20%22coloradoan.com%22%2C%0A%20%20%20%20%20%20%20%20%22commercialobserver.com%22%2C%0A%20%20%20%20%20%20%20%20%22competentedigitale.ro%22%2C%0A%20%20%20%20%20%20%20%20%22cracked.com%22%2C%0A%20%20%20%20%20%20%20%20%22dailycaller.com%22%2C%0A%20%20%20%20%20%20%20%20%22dailydot.com%22%2C%0A%20%20%20%20%20%20%20%20%22dailykos.com%22%2C%0A%20%20%20%20%20%20%20%20%22dailyvoice.com%22%2C%0A%20%20%20%20%20%20%20%20%22danslescoulisses.com%22%2C%0A%20%20%20%20%20%20%20%20%22databazeknih.cz%22%2C%0A%20%20%20%20%20%20%20%20%22deadline.com%22%2C%0A%20%20%20%20%20%20%20%20%22decider.com%22%2C%0A%20%20%20%20%20%20%20%20%22dengarden.com%22%2C%0A%20%20%20%20%20%20%20%20%22dexerto.com%22%2C%0A%20%20%20%20%20%20%20%20%22didyouknowfacts.com%22%2C%0A%20%20%20%20%20%20%20%20%22digitalmusicnews.com%22%2C%0A%20%20%20%20%20%20%20%20%22dogtime.com%22%2C%0A%20%20%20%20%20%20%20%20%22dpreview.com%22%2C%0A%20%20%20%20%20%20%20%20%22dwell.com%22%2C%0A%20%20%20%20%20%20%20%20%22eater.com%22%2C%0A%20%20%20%20%20%20%20%20%22ebaumsworld.com%22%2C%0A%20%20%20%20%20%20%20%20%22egoallstars.com%22%2C%0A%20%20%20%20%20%20%20%20%22eldiariony.com%22%2C%0A%20%20%20%20%20%20%20%20%22fark.com%22%2C%0A%20%20%20%20%20%20%20%20%22femestella.com%22%2C%0A%20%20%20%20%20%20%20%20%22flickr.com%22%2C%0A%20%20%20%20%20%20%20%20%22fmradiofree.com%22%2C%0A%20%20%20%20%20%20%20%20%22forums.hfboards.com%22%2C%0A%20%20%20%20%20%20%20%20%22free-power-point-templates.com%22%2C%0A%20%20%20%20%20%20%20%20%22freeconvert.com%22%2C%0A%20%20%20%20%20%20%20%20%22freep.com%22%2C%0A%20%20%20%20%20%20%20%20%22frogsandsnailsandpuppydogtail.com%22%2C%0A%20%20%20%20%20%20%20%20%22funtasticlife.com%22%2C%0A%20%20%20%20%20%20%20%20%22fwmadebycarli.com%22%2C%0A%20%20%20%20%20%20%20%20%22golfdigest.com%22%2C%0A%20%20%20%20%20%20%20%20%22greenbaypressgazette.com%22%2C%0A%20%20%20%20%20%20%20%20%22grunge.com%22%2C%0A%20%20%20%20%20%20%20%20%22gulflive.com%22%2C%0A%20%20%20%20%20%20%20%20%22hollywoodreporter.com%22%2C%0A%20%20%20%20%20%20%20%20%22homeglowdesign.com%22%2C%0A%20%20%20%20%20%20%20%20%22honeygirlsworld.com%22%2C%0A%20%20%20%20%20%20%20%20%22ibtimes.co.in%22%2C%0A%20%20%20%20%20%20%20%20%22imgur.com%22%2C%0A%20%20%20%20%20%20%20%20%22indiewire.com%22%2C%0A%20%20%20%20%20%20%20%20%22intouchweekly.com%22%2C%0A%20%20%20%20%20%20%20%20%22iphoneincanada.ca%22%2C%0A%20%20%20%20%20%20%20%20%22jasminemaria.com%22%2C%0A%20%20%20%20%20%20%20%20%22jsonline.com%22%2C%0A%20%20%20%20%20%20%20%20%22kens5.com%22%2C%0A%20%20%20%20%20%20%20%20%22kinobox.cz%22%2C%0A%20%20%20%20%20%20%20%20%22kion546.com%22%2C%0A%20%20%20%20%20%20%20%20%22knowyourmeme.com%22%2C%0A%20%20%20%20%20%20%20%20%22komonews.com%22%2C%0A%20%20%20%20%20%20%20%20%22krem.com%22%2C%0A%20%20%20%20%20%20%20%20%22last.fm%22%2C%0A%20%20%20%20%20%20%20%20%22lehighvalleylive.com%22%2C%0A%20%20%20%20%20%20%20%20%22lettyskitchen.com%22%2C%0A%20%20%20%20%20%20%20%20%22lifeandstylemag.com%22%2C%0A%20%20%20%20%20%20%20%20%22lifeinleggings.com%22%2C%0A%20%20%20%20%20%20%20%20%22lizzieinlace.com%22%2C%0A%20%20%20%20%20%20%20%20%22localnews8.com%22%2C%0A%20%20%20%20%20%20%20%20%22lonestarlive.com%22%2C%0A%20%20%20%20%20%20%20%20%22macwelt.de%22%2C%0A%20%20%20%20%20%20%20%20%22macworld.com%22%2C%0A%20%20%20%20%20%20%20%20%22madeeveryday.com%22%2C%0A%20%20%20%20%20%20%20%20%22maidenhead-advertiser.co.uk%22%2C%0A%20%20%20%20%20%20%20%20%22mandatory.com%22%2C%0A%20%20%20%20%20%20%20%20%22mardomreport.net%22%2C%0A%20%20%20%20%20%20%20%20%22masslive.com%22%2C%0A%20%20%20%20%20%20%20%20%22melangery.com%22%2C%0A%20%20%20%20%20%20%20%20%22mensfitness.com%22%2C%0A%20%20%20%20%20%20%20%20%22mensjournal.com%22%2C%0A%20%20%20%20%20%20%20%20%22meteomedia.com%22%2C%0A%20%20%20%20%20%20%20%20%22miamiherald.com%22%2C%0A%20%20%20%20%25%E2%80%A6
#!/usr/bin/env bash
# =============================================================================
#  ADMIN LOCAL MASTER SCRIPT
#  Archivo maestro profesional – Gestión segura de IDs y permisos de administrador local
#  Solo ejecutable y legible por el propietario =
*(LUIS GERARDO MIRANDA SEGURA )*
#  Versión: 1.0 | Fecha: 2026-08-22
# =============================================================================

set -euo pipefail
IFS=$'\n\t'

# -----------------------------------------------------------------------------
# CONFIGURACIÓN – 
# -----------------------------------------------------------------------------
# IDs / rofkam-winzes-4wyTpu
ADMIN_USER= usuario8298551C
ADMIN_UID="1000"                       # 
ADMIN_GID="1000"                       # 
ADMIN_GROUP=
LOG_FILE="${HOME}/.admin-local-master.log"
LOCK_FILE="/tmp/admin-local-master.lock"

# -----------------------------------------------------------------------------
# FUNCIONES DE SEGURIDAD
# -----------------------------------------------------------------------------
log() {
    local level="$1"
    shift
    printf '[%s] [%s] %s\n' "$(date '+%Y-%m-%d %H:%M:%S')" "$level" "$*" | tee -a "$LOG_FILE"
}

check_owner_only() {
    # Verifica que el script solo sea legible/ejecutable por el propietario
    local perms
    perms=$(stat -c '%a' "$0" 2>/dev/null || stat -f '%Lp' "$0")
    if [[ "$perms" != "700" && "$perms" != "500" ]]; then
        log "ERROR" "Permisos inseguros detectados ($perms). Debe ser 700 o 500."
        exit 1
    fi
}

require_root_or_sudo() {
    if [[ $EUID -ne 0 ]]; then
        log "INFO" "Se requiere elevación. Intentando con sudo..."
        exec sudo -E "$0" "$@"
    fi
}

# -----------------------------------------------------------------------------
# COMANDOS ASEGURADOS
# -----------------------------------------------------------------------------
mostrar_info() {
    log "INFO" "=== Información de IDs Admin Local ==="
    echo "Usuario configurado : $ADMIN_USER"
    echo "UID                 : $ADMIN_UID"
    echo "GID                 : $ADMIN_GID"
    echo "Grupo               : $ADMIN_GROUP"
    echo "Usuario actual      : $(whoami) (UID: $(id -u))"
    echo "Grupos actuales     : $(id -Gn)"
    echo "Log file            : $LOG_FILE"
}

verificar_usuario() {
    if id "$ADMIN_USER" &>/dev/null; then
        log "OK" "Usuario $ADMIN_USER existe."
        id "$ADMIN_USER"
    else
        log "WARN" "Usuario $ADMIN_USER no encontrado en el sistema."
    fi
}

aplicar_permisos_seguros() {
    # Ejemplo de comando asegurado: ajustar permisos de un directorio solo para el usuario
    local target_dir="${1:-$HOME/admin-secure}"
    mkdir -p "$target_dir"
    chown "${ADMIN_UID}:${ADMIN_GID}" "$target_dir"
    chmod 700 "$target_dir"
    log "OK" "Permisos 700 aplicados a $target_dir (solo propietario)."
}

# -----------------------------------------------------------------------------
# MENÚ PRINCIPAL
# -----------------------------------------------------------------------------
main() {
    check_owner_only

    case "${1:-}" in
        info)
            mostrar_info
            ;;
        check)
            verificar_usuario
            ;;
        secure-dir)
            aplicar_permisos_seguros "${2:-}"
            ;;
        *)
            cat <<EOF
Uso: $0 <comando> [opciones]

Comandos disponibles:
  info          Muestra la información de IDs admin configurados
  check         Verifica si el usuario admin existe
  secure-dir    Crea/ajusta un directorio con permisos 700 solo para el usuario
                Ejemplo: $0 secure-dir /ruta/segura

Ejemplo de uso seguro:
  chmod 700 $0
  ./$0 info
EOF
            exit 0
            ;;
    esac
}

# Ejecución
main "$@"





user.id
mkdir -p /home/workdir/artifacts/brillo2001/brillo2001 /home/workdir/artifacts/brillo2001/tests && ls -la /home/workdir/artifacts/brillo2001


/home/workdir/artifacts/brillo2001/pyproject.toml


/home/workdir/artifacts/brillo2001/brillo2001/__init__.py

/home/workdir/artifacts/brillo2001/brillo2001/paths.py


/home/workdir/artifacts/brillo2001/brillo2001/key_manager.py


/home/workdir/artifacts/brillo2001/brillo2001/config.py

/home/workdir/artifacts/brillo2001/brillo2001/scanner.py

/home/workdir/artifacts/brillo2001/brillo2001/organizer.py

/home/workdir/artifacts/brillo2001/brillo2001/audit.py

/home/workdir/artifacts/brillo2001/brillo2001/cli.py

/home/workdir/artifacts/brillo2001/README.md

/home/workdir/artifacts/brillo2001/requirements.txt

/home/workdir/artifacts/brillo2001/tests/test_basic.py

/home/workdir/artifacts/brillo2001/.gitignore


cd /home/workdir/artifacts/brillo2001 && python3 -m venv .venv && source .venv/bin/activate && pip install -e ".[dev]" -q && python -c "from brillo2001 import __version__; print('version', __version__)" && brillo --help | head -20


cd /home/workdir/artifacts/brillo2001 && pip install -e . --user -q 2>&1 | tail -20


export PATH="$HOME/.local/bin:$PATH" && cd /home/workdir/artifacts/brillo2001 && // WebView
(function () {
  var eventHandlers = {};

  var locationHash = '';
  try {
    locationHash = location.hash.toString();
  } catch (e) {}

  var initParams = urlParseHashParams(locationHash);
  var storedParams = sessionStorageGet('initParams');
  if (storedParams) {
    for (var key in storedParams) {
      if (typeof initParams[key] === 'undefined') {
        initParams[key] = storedParams[key];
      }
    }
  }
  sessionStorageSet('initParams', initParams);

  var isIframe = false, iFrameStyle;
  try {
    isIframe = (window.parent != null && window != window.parent);
    if (isIframe) {
      window.addEventListener('message', function (event) {
        if (event.source !== window.parent) return;
        try {
          var dataParsed = JSON.parse(event.data);
        } catch (e) {
          return;
        }
        if (!dataParsed || !dataParsed.eventType) {
          return;
        }
        if (dataParsed.eventType == 'set_custom_style') {
          if (event.origin === 'https://web.telegram.org') {
            iFrameStyle.innerHTML = dataParsed.eventData;
          }
        } else if (dataParsed.eventType == 'reload_iframe') {
          try {
            window.parent.postMessage(JSON.stringify({eventType: 'iframe_will_reload'}), '*');
          } catch (e) {}
          location.reload();
        } else {
          receiveEvent(dataParsed.eventType, dataParsed.eventData);
        }
      });
      iFrameStyle = document.createElement('style');
      document.head.appendChild(iFrameStyle);
      try {
        window.parent.postMessage(JSON.stringify({eventType: 'iframe_ready', eventData: {reload_supported: true}}), '*');
      } catch (e) {}
    }
  } catch (e) {}

  function urlSafeDecode(urlencoded) {
    try {
      urlencoded = urlencoded.replace(/\+/g, '%20');
      return decodeURIComponent(urlencoded);
    } catch (e) {
      return urlencoded;
    }
  }

  function urlParseHashParams(locationHash) {
    locationHash = locationHash.replace(/^#/, '');
    var params = {};
    if (!locationHash.length) {
      return params;
    }
    if (locationHash.indexOf('=') < 0 && locationHash.indexOf('?') < 0) {
      params._path = urlSafeDecode(locationHash);
      return params;
    }
    var qIndex = locationHash.indexOf('?');
    if (qIndex >= 0) {
      var pathParam = locationHash.substr(0, qIndex);
      params._path = urlSafeDecode(pathParam);
      locationHash = locationHash.substr(qIndex + 1);
    }
    var query_params = urlParseQueryString(locationHash);
    for (var k in query_params) {
      params[k] = query_params[k];
    }
    return params;
  }

  function urlParseQueryString(queryString) {
    var params = {};
    if (!queryString.length) {
      return params;
    }
    var queryStringParams = queryString.split('&');
    var i, param, paramName, paramValue;
    for (i = 0; i < queryStringParams.length; i++) {
      param = queryStringParams[i].split('=');
      paramName = urlSafeDecode(param[0]);
      paramValue = param[1] == null ? null : urlSafeDecode(param[1]);
      params[paramName] = paramValue;
    }
    return params;
  }

  // Telegram apps will implement this logic to add service params (e.g. tgShareScoreUrl) to game URL
  function urlAppendHashParams(url, addHash) {
    // url looks like 'https://game.com/path?query=1#hash'
    // addHash looks like 'tgShareScoreUrl=' + encodeURIComponent('tgb://share_game_score?hash=very_long_hash123')

    var ind = url.indexOf('#');
    if (ind < 0) {
      // https://game.com/path -> https://game.com/path#tgShareScoreUrl=etc
      return url + '#' + addHash;
    }
    var curHash = url.substr(ind + 1);
    if (curHash.indexOf('=') >= 0 || curHash.indexOf('?') >= 0) {
      // https://game.com/#hash=1 -> https://game.com/#hash=1&tgShareScoreUrl=etc
      // https://game.com/#path?query -> https://game.com/#path?query&tgShareScoreUrl=etc
      return url + '&' + addHash;
    }
    // https://game.com/#hash -> https://game.com/#hash?tgShareScoreUrl=etc
    if (curHash.length > 0) {
      return url + '?' + addHash;
    }
    // https://game.com/# -> https://game.com/#tgShareScoreUrl=etc
    return url + addHash;
  }

  function postEvent(eventType, callback, eventData) {
    if (!callback) {
      callback = function () {};
    }
    if (eventData === undefined) {
      eventData = '';
    }
    console.log('[Telegram.WebView] > postEvent', eventType, eventData);

    if (window.TelegramWebviewProxy !== undefined) {
      TelegramWebviewProxy.postEvent(eventType, JSON.stringify(eventData));
      callback();
    }
    else if (window.external && 'notify' in window.external) {
      window.external.notify(JSON.stringify({eventType: eventType, eventData: eventData}));
      callback();
    }
    else if (isIframe) {
      try {
        var trustedTarget = 'https://web.telegram.org';
        // For now we don't restrict target, for testing purposes
        trustedTarget = '*';
        window.parent.postMessage(JSON.stringify({eventType: eventType, eventData: eventData}), trustedTarget);
        callback();
      } catch (e) {
        callback(e);
      }
    }
    else {
      callback({notAvailable: true});
    }
  };

  function receiveEvent(eventType, eventData) {
    console.log('[Telegram.WebView] < receiveEvent', eventType, eventData);
    callEventCallbacks(eventType, function(callback) {
      callback(eventType, eventData);
    });
  }

  function callEventCallbacks(eventType, func) {
    var curEventHandlers = eventHandlers[eventType];
    if (curEventHandlers === undefined ||
        !curEventHandlers.length) {
      return;
    }
    for (var i = 0; i < curEventHandlers.length; i++) {
      try {
        func(curEventHandlers[i]);
      } catch (e) {}
    }
  }

  function onEvent(eventType, callback) {
    if (eventHandlers[eventType] === undefined) {
      eventHandlers[eventType] = [];
    }
    var index = eventHandlers[eventType].indexOf(callback);
    if (index === -1) {
      eventHandlers[eventType].push(callback);
    }
  };

  function offEvent(eventType, callback) {
    if (eventHandlers[eventType] === undefined) {
      return;
    }
    var index = eventHandlers[eventType].indexOf(callback);
    if (index === -1) {
      return;
    }
    eventHandlers[eventType].splice(index, 1);
  };

  function openProtoUrl(url) {
    if (!url.match(/^(web\+)?tgb?:\/\/./)) {
      return false;
    }
    var useIframe = navigator.userAgent.match(/iOS|iPhone OS|iPhone|iPod|iPad/i) ? true : false;
    if (useIframe) {
      var iframeContEl = document.getElementById('tgme_frame_cont') || document.body;
      var iframeEl = document.createElement('iframe');
      iframeContEl.appendChild(iframeEl);
      var pageHidden = false;
      var enableHidden = function () {
        pageHidden = true;
      };
      window.addEventListener('pagehide', enableHidden, false);
      window.addEventListener('blur', enableHidden, false);
      if (iframeEl !== null) {
        iframeEl.src = url;
      }
      setTimeout(function() {
        if (!pageHidden) {
          window.location = url;
        }
        window.removeEventListener('pagehide', enableHidden, false);
        window.removeEventListener('blur', enableHidden, false);
      }, 2000);
    }
    else {
      window.location = url;
    }
    return true;
  }

  function sessionStorageSet(key, value) {
    try {
      window.sessionStorage.setItem('__telegram__' + key, JSON.stringify(value));
      return true;
    } catch(e) {}
    return false;
  }
  function sessionStorageGet(key) {
    try {
      return JSON.parse(window.sessionStorage.getItem('__telegram__' + key));
    } catch(e) {}
    return null;
  }

  if (!window.Telegram) {
    window.Telegram = {};
  }
  window.Telegram.WebView = {
    initParams: initParams,
    isIframe: isIframe,
    onEvent: onEvent,
    offEvent: offEvent,
    postEvent: postEvent,
    receiveEvent: receiveEvent,
    callEventCallbacks: callEventCallbacks
  };

  window.Telegram.Utils = {
    urlSafeDecode: urlSafeDecode,
    urlParseQueryString: urlParseQueryString,
    urlParseHashParams: urlParseHashParams,
    urlAppendHashParams: urlAppendHashParams,
    sessionStorageSet: sessionStorageSet,
    sessionStorageGet: sessionStorageGet
  };

  // For Windows Phone app
  window.TelegramGameProxy_receiveEvent = receiveEvent;

  // App backward compatibility
  window.TelegramGameProxy = {
    receiveEvent: receiveEvent
  };
})();

// WebApp
(function () {
  var Utils = window.Telegram.Utils;
  var WebView = window.Telegram.WebView;
  var initParams = WebView.initParams;
  var isIframe = WebView.isIframe;

  var WebApp = {};
  var webAppInitData = '', webAppInitDataUnsafe = {};
  var themeParams = {}, colorScheme = 'light';
  var webAppVersion = '6.0';
  var webAppPlatform = 'unknown';
  var webAppIsActive = true;
  var webAppIsFullscreen = false;
  var webAppIsOrientationLocked = false;
  var webAppBackgroundColor = 'bg_color';
  var webAppHeaderColorKey = 'bg_color';
  var webAppHeaderColor = null;

  if (initParams.tgWebAppData && initParams.tgWebAppData.length) {
    webAppInitData = initParams.tgWebAppData;
    webAppInitDataUnsafe = Utils.urlParseQueryString(webAppInitData);
    for (var key in webAppInitDataUnsafe) {
      var val = webAppInitDataUnsafe[key];
      try {
        if (val.substr(0, 1) == '{' && val.substr(-1) == '}' ||
            val.substr(0, 1) == '[' && val.substr(-1) == ']') {
          webAppInitDataUnsafe[key] = JSON.parse(val);
        }
      } catch (e) {}
    }
  }
  var stored_theme_params = Utils.sessionStorageGet('themeParams');
  if (initParams.tgWebAppThemeParams && initParams.tgWebAppThemeParams.length) {
    var themeParamsRaw = initParams.tgWebAppThemeParams;
    try {
      var theme_params = JSON.parse(themeParamsRaw);
      if (theme_params) {
        setThemeParams(theme_params);
      }
    } catch (e) {}
  }
  if (stored_theme_params) {
    setThemeParams(stored_theme_params);
  }
  var stored_def_colors = Utils.sessionStorageGet('defaultColors');
  if (initParams.tgWebAppDefaultColors && initParams.tgWebAppDefaultColors.length) {
    var defColorsRaw = initParams.tgWebAppDefaultColors;
    try {
      var def_colors = JSON.parse(defColorsRaw);
      if (def_colors) {
        setDefaultColors(def_colors);
      }
    } catch (e) {}
  }
  if (stored_def_colors) {
    setDefaultColors(stored_def_colors);
  }
  if (initParams.tgWebAppVersion) {
    webAppVersion = initParams.tgWebAppVersion;
  }
  if (initParams.tgWebAppPlatform) {
    webAppPlatform = initParams.tgWebAppPlatform;
  }

  var stored_fullscreen = Utils.sessionStorageGet('isFullscreen');
  if (initParams.tgWebAppFullscreen) {
    setFullscreen(true);
  }
  if (stored_fullscreen) {
    setFullscreen(stored_fullscreen == 'yes');
  }

  var stored_orientation_lock = Utils.sessionStorageGet('isOrientationLocked');
  if (stored_orientation_lock) {
    setOrientationLock(stored_orientation_lock == 'yes');
  }

  function onThemeChanged(eventType, eventData) {
    if (eventData.theme_params) {
      setThemeParams(eventData.theme_params);
      window.Telegram.WebApp.MainButton.setParams({});
      window.Telegram.WebApp.SecondaryButton.setParams({});
      updateHeaderColor();
      updateBackgroundColor();
      updateBottomBarColor();
      receiveWebViewEvent('themeChanged');
    }
  }

  var lastWindowHeight = window.innerHeight;
  function onViewportChanged(eventType, eventData) {
    if (eventData.height) {
      window.removeEventListener('resize', onWindowResize);
      setViewportHeight(eventData);
    }
  }

  function onWindowResize(e) {
    if (lastWindowHeight != window.innerHeight) {
      lastWindowHeight = window.innerHeight;
      receiveWebViewEvent('viewportChanged', {
        isStateStable: true
      });
    }
  }

  function onSafeAreaChanged(eventType, eventData) {
    if (eventData) {
      setSafeAreaInset(eventData);
    }
  }
  function onContentSafeAreaChanged(eventType, eventData) {
    if (eventData) {
      setContentSafeAreaInset(eventData);
    }
  }

  function onVisibilityChanged(eventType, eventData) {
    if (eventData.is_visible) {
      webAppIsActive = true;
      receiveWebViewEvent('activated');
    } else {
      webAppIsActive = false;
      receiveWebViewEvent('deactivated');
    }
  }

  function linkHandler(e) {
    if (e.metaKey || e.ctrlKey) return;
    var el = e.target;
    while (el.tagName != 'A' && el.parentNode) {
      el = el.parentNode;
    }
    if (el.tagName == 'A' &&
        el.target != '_blank' &&
        (el.protocol == 'http:' || el.protocol == 'https:') &&
        isTmeHostname(el.hostname)) {
      WebApp.openTelegramLink(el.href);
      e.preventDefault();
    }
  }

  function strTrim(str) {
    return str.toString().replace(/^\s+|\s+$/g, '');
  }

  function isTmeHostname(hostname) {
    hostname = hostname.toString().toLowerCase();
    return hostname == 't.me' || hostname == 'telegram.me';
  }

  function receiveWebViewEvent(eventType) {
    var args = Array.prototype.slice.call(arguments);
    eventType = args.shift();
    WebView.callEventCallbacks('webview:' + eventType, function(callback) {
      callback.apply(WebApp, args);
    });
  }

  function onWebViewEvent(eventType, callback) {
    WebView.onEvent('webview:' + eventType, callback);
  };

  function offWebViewEvent(eventType, callback) {
    WebView.offEvent('webview:' + eventType, callback);
  };

  function setCssProperty(name, value) {
    var root = document.documentElement;
    if (root && root.style && root.style.setProperty) {
      root.style.setProperty('--tg-' + name, value);
    }
  }

  function setFullscreen(is_fullscreen) {
    webAppIsFullscreen = !!is_fullscreen;
    Utils.sessionStorageSet('isFullscreen', webAppIsFullscreen ? 'yes' : 'no');
  }

  function setOrientationLock(is_locked) {
    webAppIsOrientationLocked = !!is_locked;
    Utils.sessionStorageSet('isOrientationLocked', webAppIsOrientationLocked ? 'yes' : 'no');
  }

  function setThemeParams(theme_params) {
    // temp iOS fix
    if (theme_params.bg_color == '#1c1c1d' &&
        theme_params.bg_color == theme_params.secondary_bg_color) {
      theme_params.secondary_bg_color = '#2c2c2e';
    }
    var color;
    for (var key in theme_params) {
      if (color = parseColorToHex(theme_params[key])) {
        themeParams[key] = color;
        if (key == 'bg_color') {
          colorScheme = isColorDark(color) ? 'dark' : 'light'
          setCssProperty('color-scheme', colorScheme);
        }
        key = 'theme-' + key.split('_').join('-');
        setCssProperty(key, color);
      }
    }
    Utils.sessionStorageSet('themeParams', themeParams);
  }

  function setDefaultColors(def_colors) {
    if (colorScheme == 'dark') {
      if (def_colors.bg_dark_color) {
        webAppBackgroundColor = def_colors.bg_dark_color;
      }
      if (def_colors.header_dark_color) {
        webAppHeaderColorKey = null;
        webAppHeaderColor = def_colors.header_dark_color;
      }
    } else {
      if (def_colors.bg_color) {
        webAppBackgroundColor = def_colors.bg_color;
      }
      if (def_colors.header_color) {
        webAppHeaderColorKey = null;
        webAppHeaderColor = def_colors.header_color;
      }
    }
    Utils.sessionStorageSet('defaultColors', def_colors);
  }

  var webAppCallbacks = {};
  function generateCallbackId(len) {
    var tries = 100;
    while (--tries) {
      var id = '', chars = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789', chars_len = chars.length;
      for (var i = 0; i < len; i++) {
        id += chars[Math.floor(Math.random() * chars_len)];
      }
      if (!webAppCallbacks[id]) {
        webAppCallbacks[id] = {};
        return id;
      }
    }
    throw Error('WebAppCallbackIdGenerateFailed');
  }

  var viewportHeight = false, viewportStableHeight = false, isExpanded = true;
  function setViewportHeight(data) {
    if (typeof data !== 'undefined') {
      isExpanded = !!data.is_expanded;
      viewportHeight = data.height;
      if (data.is_state_stable) {
        viewportStableHeight = data.height;
      }
      receiveWebViewEvent('viewportChanged', {
        isStateStable: !!data.is_state_stable
      });
    }
    var height, stable_height;
    if (viewportHeight !== false) {
      height = (viewportHeight - bottomBarHeight) + 'px';
    } else {
      height = bottomBarHeight ? 'calc(100vh - ' + bottomBarHeight + 'px)' : '100vh';
    }
    if (viewportStableHeight !== false) {
      stable_height = (viewportStableHeight - bottomBarHeight) + 'px';
    } else {
      stable_height = bottomBarHeight ? 'calc(100vh - ' + bottomBarHeight + 'px)' : '100vh';
    }
    setCssProperty('viewport-height', height);
    setCssProperty('viewport-stable-height', stable_height);
  }

  var safeAreaInset = {top: 0, bottom: 0, left: 0, right: 0};
  function setSafeAreaInset(data) {
    if (typeof data !== 'undefined') {
      if (typeof data.top !== 'undefined') {
        safeAreaInset.top = data.top;
      }
      if (typeof data.bottom !== 'undefined') {
        safeAreaInset.bottom = data.bottom;
      }
      if (typeof data.left !== 'undefined') {
        safeAreaInset.left = data.left;
      }
      if (typeof data.right !== 'undefined') {
        safeAreaInset.right = data.right;
      }
      receiveWebViewEvent('safeAreaChanged');
    }
    setCssProperty('safe-area-inset-top', safeAreaInset.top + 'px');
    setCssProperty('safe-area-inset-bottom', safeAreaInset.bottom + 'px');
    setCssProperty('safe-area-inset-left', safeAreaInset.left + 'px');
    setCssProperty('safe-area-inset-right', safeAreaInset.right + 'px');
  }

  var contentSafeAreaInset = {top: 0, bottom: 0, left: 0, right: 0};
  function setContentSafeAreaInset(data) {
    if (typeof data !== 'undefined') {
      if (typeof data.top !== 'undefined') {
        contentSafeAreaInset.top = data.top;
      }
      if (typeof data.bottom !== 'undefined') {
        contentSafeAreaInset.bottom = data.bottom;
      }
      if (typeof data.left !== 'undefined') {
        contentSafeAreaInset.left = data.left;
      }
      if (typeof data.right !== 'undefined') {
        contentSafeAreaInset.right = data.right;
      }
      receiveWebViewEvent('contentSafeAreaChanged');
    }
    setCssProperty('content-safe-area-inset-top', contentSafeAreaInset.top + 'px');
    setCssProperty('content-safe-area-inset-bottom', contentSafeAreaInset.bottom + 'px');
    setCssProperty('content-safe-area-inset-left', contentSafeAreaInset.left + 'px');
    setCssProperty('content-safe-area-inset-right', contentSafeAreaInset.right + 'px');
  }

  var isClosingConfirmationEnabled = false;
  function setClosingConfirmation(need_confirmation) {
    if (!versionAtLeast('6.2')) {
      console.warn('[Telegram.WebApp] Closing confirmation is not supported in version ' + webAppVersion);
      return;
    }
    isClosingConfirmationEnabled = !!need_confirmation;
    WebView.postEvent('web_app_setup_closing_behavior', false, {need_confirmation: isClosingConfirmationEnabled});
  }

  var isVerticalSwipesEnabled = true;
  function toggleVerticalSwipes(enable_swipes) {
    if (!versionAtLeast('7.7')) {
      console.warn('[Telegram.WebApp] Changing swipes behavior is not supported in version ' + webAppVersion);
      return;
    }
    isVerticalSwipesEnabled = !!enable_swipes;
    WebView.postEvent('web_app_setup_swipe_behavior', false, {allow_vertical_swipe: isVerticalSwipesEnabled});
  }

  function onFullscreenChanged(eventType, eventData) {
    setFullscreen(eventData.is_fullscreen);
    receiveWebViewEvent('fullscreenChanged');
  }
  function onFullscreenFailed(eventType, eventData) {
    if (eventData.error == 'ALREADY_FULLSCREEN' && !webAppIsFullscreen) {
      setFullscreen(true);
    }
    receiveWebViewEvent('fullscreenFailed', {
      error: eventData.error
    });
  }

  function toggleOrientationLock(locked) {
    if (!versionAtLeast('8.0')) {
      console.warn('[Telegram.WebApp] Orientation locking is not supported in version ' + webAppVersion);
      return;
    }
    setOrientationLock(locked);
    WebView.postEvent('web_app_toggle_orientation_lock', false, {locked: webAppIsOrientationLocked});
  }

  var homeScreenCallbacks = [];
  function onHomeScreenAdded(eventType, eventData) {
    receiveWebViewEvent('homeScreenAdded');
  }
  function onHomeScreenChecked(eventType, eventData) {
    var status = eventData.status || 'unknown';
    if (homeScreenCallbacks.length > 0) {
      for (var i = 0; i < homeScreenCallbacks.length; i++) {
        var callback = homeScreenCallbacks[i];
        callback(status);
      }
      homeScreenCallbacks = [];
    }
    receiveWebViewEvent('homeScreenChecked', {
      status: status
    });
  }

  var WebAppShareMessageOpened = false;
  function onPreparedMessageSent(eventType, eventData) {
    if (WebAppShareMessageOpened) {
      var requestData = WebAppShareMessageOpened;
      WebAppShareMessageOpened = false;
      if (requestData.callback) {
        requestData.callback(true);
      }
      receiveWebViewEvent('shareMessageSent');
    }
  }
  function onPreparedMessageFailed(eventType, eventData) {
    if (WebAppShareMessageOpened) {
      var requestData = WebAppShareMessageOpened;
      WebAppShareMessageOpened = false;
      if (requestData.callback) {
        requestData.callback(false);
      }
      receiveWebViewEvent('shareMessageFailed', {
        error: eventData.error
      });
    }
  }

  var WebAppRequestChatOpened = false;
  function onRequestedChatSent(eventType, eventData) {
    if (WebAppRequestChatOpened) {
      var requestData = WebAppRequestChatOpened;
      WebAppRequestChatOpened = false;
      if (requestData.callback) {
        requestData.callback(true);
      }
      receiveWebViewEvent('requestedChatSent');
    }
  }
  function onRequestedChatFailed(eventType, eventData) {
    if (WebAppRequestChatOpened) {
      var requestData = WebAppRequestChatOpened;
      WebAppRequestChatOpened = false;
      if (requestData.callback) {
        requestData.callback(false);
      }
      receiveWebViewEvent('requestedChatFailed', {
        error: eventData.error
      });
    }
  }

  var WebAppEmojiStatusRequested = false;
  function onEmojiStatusSet(eventType, eventData) {
    if (WebAppEmojiStatusRequested) {
      var requestData = WebAppEmojiStatusRequested;
      WebAppEmojiStatusRequested = false;
      if (requestData.callback) {
        requestData.callback(true);
      }
      receiveWebViewEvent('emojiStatusSet');
    }
  }
  function onEmojiStatusFailed(eventType, eventData) {
    if (WebAppEmojiStatusRequested) {
      var requestData = WebAppEmojiStatusRequested;
      WebAppEmojiStatusRequested = false;
      if (requestData.callback) {
        requestData.callback(false);
      }
      receiveWebViewEvent('emojiStatusFailed', {
        error: eventData.error
      });
    }
  }
  var WebAppEmojiStatusAccessRequested = false;
  function onEmojiStatusAccessRequested(eventType, eventData) {
    if (WebAppEmojiStatusAccessRequested) {
      var requestData = WebAppEmojiStatusAccessRequested;
      WebAppEmojiStatusAccessRequested = false;
      if (requestData.callback) {
        requestData.callback(eventData.status == 'allowed');
      }
      receiveWebViewEvent('emojiStatusAccessRequested', {
        status: eventData.status
      });
    }
  }

  var webAppPopupOpened = false;
  function onPopupClosed(eventType, eventData) {
    if (webAppPopupOpened) {
      var popupData = webAppPopupOpened;
      webAppPopupOpened = false;
      var button_id = null;
      if (typeof eventData.button_id !== 'undefined') {
        button_id = eventData.button_id;
      }
      if (popupData.callback) {
        popupData.callback(button_id);
      }
      receiveWebViewEvent('popupClosed', {
        button_id: button_id
      });
    }
  }


  function getHeaderColor() {
    if (webAppHeaderColorKey == 'secondary_bg_color') {
      return themeParams.secondary_bg_color;
    } else if (webAppHeaderColorKey == 'bg_color') {
      return themeParams.bg_color;
    }
    return webAppHeaderColor;
  }
  function setHeaderColor(color) {
    if (!versionAtLeast('6.1')) {
      console.warn('[Telegram.WebApp] Header color is not supported in version ' + webAppVersion);
      return;
    }
    if (!versionAtLeast('6.9')) {
      if (themeParams.bg_color &&
          themeParams.bg_color == color) {
        color = 'bg_color';
      } else if (themeParams.secondary_bg_color &&
                 themeParams.secondary_bg_color == color) {
        color = 'secondary_bg_color';
      }
    }
    var head_color = null, color_key = null;
    if (color == 'bg_color' || color == 'secondary_bg_color') {
      color_key = color;
    } else if (versionAtLeast('6.9')) {
      head_color = parseColorToHex(color);
      if (!head_color) {
        console.error('[Telegram.WebApp] Header color format is invalid', color);
        throw Error('WebAppHeaderColorInvalid');
      }
    }
    if (!versionAtLeast('6.9') &&
        color_key != 'bg_color' &&
        color_key != 'secondary_bg_color') {
      console.error('[Telegram.WebApp] Header color key should be one of Telegram.WebApp.themeParams.bg_color, Telegram.WebApp.themeParams.secondary_bg_color, \'bg_color\', \'secondary_bg_color\'', color);
      throw Error('WebAppHeaderColorKeyInvalid');
    }
    webAppHeaderColorKey = color_key;
    webAppHeaderColor = head_color;
    updateHeaderColor();
  }
  var appHeaderColorKey = null, appHeaderColor = null;
  function updateHeaderColor() {
    if (appHeaderColorKey != webAppHeaderColorKey ||
        appHeaderColor != webAppHeaderColor) {
      appHeaderColorKey = webAppHeaderColorKey;
      appHeaderColor = webAppHeaderColor;
      if (appHeaderColor) {
        WebView.postEvent('web_app_set_header_color', false, {color: webAppHeaderColor});
      } else {
        WebView.postEvent('web_app_set_header_color', false, {color_key: webAppHeaderColorKey});
      }
    }
  }

  function getBackgroundColor() {
    if (webAppBackgroundColor == 'secondary_bg_color') {
      return themeParams.secondary_bg_color;
    } else if (webAppBackgroundColor == 'bg_color') {
      return themeParams.bg_color;
    }
    return webAppBackgroundColor;
  }
  function setBackgroundColor(color) {
    if (!versionAtLeast('6.1')) {
      console.warn('[Telegram.WebApp] Background color is not supported in version ' + webAppVersion);
      return;
    }
    var bg_color;
    if (color == 'bg_color' || color == 'secondary_bg_color') {
      bg_color = color;
    } else {
      bg_color = parseColorToHex(color);
      if (!bg_color) {
        console.error('[Telegram.WebApp] Background color format is invalid', color);
        throw Error('WebAppBackgroundColorInvalid');
      }
    }
    webAppBackgroundColor = bg_color;
    updateBackgroundColor();
  }
  var appBackgroundColor = null;
  function updateBackgroundColor() {
    var color = getBackgroundColor();
    if (appBackgroundColor != color) {
      appBackgroundColor = color;
      WebView.postEvent('web_app_set_background_color', false, {color: color});
    }
  }

  var bottomBarColor = 'bottom_bar_bg_color';
  function getBottomBarColor() {
    if (bottomBarColor == 'bottom_bar_bg_color') {
      return themeParams.bottom_bar_bg_color || themeParams.secondary_bg_color || '#ffffff';
    } else if (bottomBarColor == 'secondary_bg_color') {
      return themeParams.secondary_bg_color;
    } else if (bottomBarColor == 'bg_color') {
      return themeParams.bg_color;
    }
    return bottomBarColor;
  }
  function setBottomBarColor(color) {
    if (!versionAtLeast('7.10')) {
      console.warn('[Telegram.WebApp] Bottom bar color is not supported in version ' + webAppVersion);
      return;
    }
    var bg_color;
    if (color == 'bg_color' || color == 'secondary_bg_color' || color == 'bottom_bar_bg_color') {
      bg_color = color;
    } else {
      bg_color = parseColorToHex(color);
      if (!bg_color) {
        console.error('[Telegram.WebApp] Bottom bar color format is invalid', color);
        throw Error('WebAppBottomBarColorInvalid');
      }
    }
    bottomBarColor = bg_color;
    updateBottomBarColor();
    window.Telegram.WebApp.SecondaryButton.setParams({});
  }
  var appBottomBarColor = null;
  function updateBottomBarColor() {
    var color = getBottomBarColor();
    if (appBottomBarColor != color) {
      appBottomBarColor = color;
      WebView.postEvent('web_app_set_bottom_bar_color', false, {color: color});
    }
    if (initParams.tgWebAppDebug) {
      updateDebugBottomBar();
    }
  }


  function parseColorToHex(color) {
    color += '';
    var match;
    if (match = /^\s*#([0-9a-f]{6})\s*$/i.exec(color)) {
      return '#' + match[1].toLowerCase();
    }
    else if (match = /^\s*#([0-9a-f])([0-9a-f])([0-9a-f])\s*$/i.exec(color)) {
      return ('#' + match[1] + match[1] + match[2] + match[2] + match[3] + match[3]).toLowerCase();
    }
    else if (match = /^\s*rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+\.{0,1}\d*))?\)\s*$/.exec(color)) {
      var r = parseInt(match[1]), g = parseInt(match[2]), b = parseInt(match[3]);
      r = (r < 16 ? '0' : '') + r.toString(16);
      g = (g < 16 ? '0' : '') + g.toString(16);
      b = (b < 16 ? '0' : '') + b.toString(16);
      return '#' + r + g + b;
    }
    return false;
  }

  function isColorDark(rgb) {
    rgb = rgb.replace(/[\s#]/g, '');
    if (rgb.length == 3) {
      rgb = rgb[0] + rgb[0] + rgb[1] + rgb[1] + rgb[2] + rgb[2];
    }
    var r = parseInt(rgb.substr(0, 2), 16);
    var g = parseInt(rgb.substr(2, 2), 16);
    var b = parseInt(rgb.substr(4, 2), 16);
    var hsp = Math.sqrt(0.299 * (r * r) + 0.587 * (g * g) + 0.114 * (b * b));
    return hsp < 120;
  }

  function versionCompare(v1, v2) {
    if (typeof v1 !== 'string') v1 = '';
    if (typeof v2 !== 'string') v2 = '';
    v1 = v1.replace(/^\s+|\s+$/g, '').split('.');
    v2 = v2.replace(/^\s+|\s+$/g, '').split('.');
    var a = Math.max(v1.length, v2.length), i, p1, p2;
    for (i = 0; i < a; i++) {
      p1 = parseInt(v1[i]) || 0;
      p2 = parseInt(v2[i]) || 0;
      if (p1 == p2) continue;
      if (p1 > p2) return 1;
      return -1;
    }
    return 0;
  }

  function versionAtLeast(ver) {
    return versionCompare(webAppVersion, ver) >= 0;
  }

  function byteLength(str) {
    if (window.Blob) {
      try { return new Blob([str]).size; } catch (e) {}
    }
    var s = str.length;
    for (var i=str.length-1; i>=0; i--) {
      var code = str.charCodeAt(i);
      if (code > 0x7f && code <= 0x7ff) s++;
      else if (code > 0x7ff && code <= 0xffff) s+=2;
      if (code >= 0xdc00 && code <= 0xdfff) i--;
    }
    return s;
  }

  var BackButton = (function() {
    var isVisible = false;

    var backButton = {};
    Object.defineProperty(backButton, 'isVisible', {
      set: function(val){ setParams({is_visible: val}); },
      get: function(){ return isVisible; },
      enumerable: true
    });

    var curButtonState = null;

    WebView.onEvent('back_button_pressed', onBackButtonPressed);

    function onBackButtonPressed() {
      receiveWebViewEvent('backButtonClicked');
    }

    function buttonParams() {
      return {is_visible: isVisible};
    }

    function buttonState(btn_params) {
      if (typeof btn_params === 'undefined') {
        btn_params = buttonParams();
      }
      return JSON.stringify(btn_params);
    }

    function buttonCheckVersion() {
      if (!versionAtLeast('6.1')) {
        console.warn('[Telegram.WebApp] BackButton is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    function updateButton() {
      var btn_params = buttonParams();
      var btn_state = buttonState(btn_params);
      if (curButtonState === btn_state) {
        return;
      }
      curButtonState = btn_state;
      WebView.postEvent('web_app_setup_back_button', false, btn_params);
    }

    function setParams(params) {
      if (!buttonCheckVersion()) {
        return backButton;
      }
      if (typeof params.is_visible !== 'undefined') {
        isVisible = !!params.is_visible;
      }
      updateButton();
      return backButton;
    }

    backButton.onClick = function(callback) {
      if (buttonCheckVersion()) {
        onWebViewEvent('backButtonClicked', callback);
      }
      return backButton;
    };
    backButton.offClick = function(callback) {
      if (buttonCheckVersion()) {
        offWebViewEvent('backButtonClicked', callback);
      }
      return backButton;
    };
    backButton.show = function() {
      return setParams({is_visible: true});
    };
    backButton.hide = function() {
      return setParams({is_visible: false});
    };
    return backButton;
  })();

  var debugBottomBar = null, debugBottomBarBtns = {}, bottomBarHeight = 0;
  if (initParams.tgWebAppDebug) {
    debugBottomBar = document.createElement('tg-bottom-bar');
    var debugBottomBarStyle = {
      display: 'flex',
      gap: '7px',
      font: '600 14px/18px sans-serif',
      width: '100%',
      background: getBottomBarColor(),
      position: 'fixed',
      left: '0',
      right: '0',
      bottom: '0',
      margin: '0',
      padding: '7px',
      textAlign: 'center',
      boxSizing: 'border-box',
      zIndex: '10000'
    };
    for (var k in debugBottomBarStyle) {
      debugBottomBar.style[k] = debugBottomBarStyle[k];
    }
    document.addEventListener('DOMContentLoaded', function onDomLoaded(event) {
      document.removeEventListener('DOMContentLoaded', onDomLoaded);
      document.body.appendChild(debugBottomBar);
    });
    var animStyle = document.createElement('style');
    animStyle.innerHTML = 'tg-bottom-button.shine { position: relative; overflow: hidden; } tg-bottom-button.shine:before { content:""; position: absolute; top: 0; width: 100%; height: 100%; background: linear-gradient(120deg, transparent, rgba(255, 255, 255, .2), transparent); animation: tg-bottom-button-shine 5s ease-in-out infinite; } @-webkit-keyframes tg-bottom-button-shine { 0% {left: -100%;} 12%,100% {left: 100%}} @keyframes tg-bottom-button-shine { 0% {left: -100%;} 12%,100% {left: 100%}}';
    debugBottomBar.appendChild(animStyle);
  }
  function updateDebugBottomBar() {
    var mainBtn = debugBottomBarBtns.main._bottomButton;
    var secondaryBtn = debugBottomBarBtns.secondary._bottomButton;
    if (mainBtn.isVisible || secondaryBtn.isVisible) {
      debugBottomBar.style.display = 'flex';
      bottomBarHeight = 58;
      if (mainBtn.isVisible && secondaryBtn.isVisible) {
        if (secondaryBtn.position == 'top') {
          debugBottomBar.style.flexDirection = 'column-reverse';
          bottomBarHeight += 51;
        } else if (secondaryBtn.position == 'bottom') {
          debugBottomBar.style.flexDirection = 'column';
          bottomBarHeight += 51;
        } else if (secondaryBtn.position == 'left') {
          debugBottomBar.style.flexDirection = 'row-reverse';
        } else if (secondaryBtn.position == 'right') {
          debugBottomBar.style.flexDirection = 'row';
        }
      }
    } else {
      debugBottomBar.style.display = 'none';
      bottomBarHeight = 0;
    }
    debugBottomBar.style.background = getBottomBarColor();
    if (document.documentElement) {
      document.documentElement.style.boxSizing = 'border-box';
      document.documentElement.style.paddingBottom = bottomBarHeight + 'px';
    }
    setViewportHeight();
  }


  var BottomButtonConstructor = function(type) {
    var isMainButton = (type == 'main');
    if (isMainButton) {
      var setupFnName = 'web_app_setup_main_button';
      var tgEventName = 'main_button_pressed';
      var webViewEventName = 'mainButtonClicked';
      var buttonTextDefault = 'Continue';
      var buttonColorDefault = function(){ return themeParams.button_color || '#2481cc'; };
      var buttonTextColorDefault = function(){ return themeParams.button_text_color || '#ffffff'; };
    } else {
      var setupFnName = 'web_app_setup_secondary_button';
      var tgEventName = 'secondary_button_pressed';
      var webViewEventName = 'secondaryButtonClicked';
      var buttonTextDefault = 'Cancel';
      var buttonColorDefault = function(){ return getBottomBarColor(); };
      var buttonTextColorDefault = function(){ return themeParams.button_color || '#2481cc'; };
    }

    var isVisible = false;
    var isActive = true;
    var hasShineEffect = false;
    var isProgressVisible = false;
    var iconCustomEmojiId = false;
    var buttonType = type;
    var buttonText = buttonTextDefault;
    var buttonColor = false;
    var buttonTextColor = false;
    var buttonPosition = 'left';

    var bottomButton = {};
    Object.defineProperty(bottomButton, 'type', {
      get: function(){ return buttonType; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'iconCustomEmojiId', {
      set: function(val){ bottomButton.setParams({icon_custom_emoji_id: val}); },
      get: function(){ return iconCustomEmojiId; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'text', {
      set: function(val){ bottomButton.setParams({text: val}); },
      get: function(){ return buttonText; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'color', {
      set: function(val){ bottomButton.setParams({color: val}); },
      get: function(){ return buttonColor || buttonColorDefault(); },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'textColor', {
      set: function(val){ bottomButton.setParams({text_color: val}); },
      get: function(){ return buttonTextColor || buttonTextColorDefault(); },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'isVisible', {
      set: function(val){ bottomButton.setParams({is_visible: val}); },
      get: function(){ return isVisible; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'isProgressVisible', {
      get: function(){ return isProgressVisible; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'isActive', {
      set: function(val){ bottomButton.setParams({is_active: val}); },
      get: function(){ return isActive; },
      enumerable: true
    });
    Object.defineProperty(bottomButton, 'hasShineEffect', {
      set: function(val){ bottomButton.setParams({has_shine_effect: val}); },
      get: function(){ return hasShineEffect; },
      enumerable: true
    });
    if (!isMainButton) {
      Object.defineProperty(bottomButton, 'position', {
        set: function(val){ bottomButton.setParams({position: val}); },
        get: function(){ return buttonPosition; },
        enumerable: true
      });
    }

    var curButtonState = null;

    WebView.onEvent(tgEventName, onBottomButtonPressed);

    var debugBtn = null;
    if (initParams.tgWebAppDebug) {
      debugBtn = document.createElement('tg-bottom-button');
      var debugBtnStyle = {
        display: 'none',
        width: '100%',
        height: '44px',
        borderRadius: '0',
        background: 'no-repeat right center',
        padding: '13px 15px',
        textAlign: 'center',
        boxSizing: 'border-box'
      };
      for (var k in debugBtnStyle) {
        debugBtn.style[k] = debugBtnStyle[k];
      }
      debugBottomBar.appendChild(debugBtn);
      debugBtn.addEventListener('click', onBottomButtonPressed, false);
      debugBtn._bottomButton = bottomButton;
      debugBottomBarBtns[type] = debugBtn;
    }

    function onBottomButtonPressed() {
      if (isActive) {
        receiveWebViewEvent(webViewEventName);
      }
    }

    function buttonParams() {
      var color = bottomButton.color;
      var text_color = bottomButton.textColor;
      if (isVisible) {
        var params = {
          is_visible: true,
          is_active: isActive,
          is_progress_visible: isProgressVisible,
          icon_custom_emoji_id: iconCustomEmojiId,
          text: buttonText,
          color: color,
          text_color: text_color,
          has_shine_effect: hasShineEffect && isActive && !isProgressVisible
        };
        if (!isMainButton) {
          params.position = buttonPosition;
        }
      } else {
        var params = {
          is_visible: false
        };
      }
      return params;
    }

    function buttonState(btn_params) {
      if (typeof btn_params === 'undefined') {
        btn_params = buttonParams();
      }
      return JSON.stringify(btn_params);
    }

    function updateButton() {
      var btn_params = buttonParams();
      var btn_state = buttonState(btn_params);
      if (curButtonState === btn_state) {
        return;
      }
      curButtonState = btn_state;
      WebView.postEvent(setupFnName, false, btn_params);
      if (initParams.tgWebAppDebug) {
        updateDebugButton(btn_params);
      }
    }

    function updateDebugButton(btn_params) {
      if (btn_params.is_visible) {
        debugBtn.style.display = 'block';

        debugBtn.style.opacity = btn_params.is_active ? '1' : '0.8';
        debugBtn.style.cursor = btn_params.is_active ? 'pointer' : 'auto';
        debugBtn.disabled = !btn_params.is_active;
        debugBtn.innerText = btn_params.text;
        debugBtn.className = btn_params.has_shine_effect ? 'shine' : '';
        debugBtn.style.backgroundImage = btn_params.is_progress_visible ? "url('data:image/svg+xml," + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewport="0 0 48 48" width="48px" height="48px"><circle cx="50%" cy="50%" stroke="' + btn_params.text_color + '" stroke-width="2.25" stroke-linecap="round" fill="none" stroke-dashoffset="106" r="9" stroke-dasharray="56.52" rotate="-90"><animate attributeName="stroke-dashoffset" attributeType="XML" dur="360s" from="0" to="12500" repeatCount="indefinite"></animate><animateTransform attributeName="transform" attributeType="XML" type="rotate" dur="1s" from="-90 24 24" to="630 24 24" repeatCount="indefinite"></animateTransform></circle></svg>') + "')" : 'none';
        debugBtn.style.backgroundColor = btn_params.color;
        debugBtn.style.color = btn_params.text_color;
      } else {
        debugBtn.style.display = 'none';
      }
      updateDebugBottomBar();
    }

    function setParams(params) {
      if (typeof params.icon_custom_emoji_id !== 'undefined') {
        var emoji_id = params.icon_custom_emoji_id;
        if (emoji_id === false || emoji_id === null) {
          emoji_id = '';
        }
        if (emoji_id !== '' && !/^[0-9]{10,20}$/.test(emoji_id)) {
          console.error('[Telegram.WebApp] Bottom button icon custom emoji is invalid', params.icon_custom_emoji_id);
          throw Error('WebAppBottomButtonParamInvalid');
        }
        iconCustomEmojiId = emoji_id;
      }
      if (typeof params.text !== 'undefined') {
        var text = strTrim(params.text);
        if (!text.length && !iconCustomEmojiId) {
          console.error('[Telegram.WebApp] Bottom button text is required', params.text);
          throw Error('WebAppBottomButtonParamInvalid');
        }
        if (text.length > 64) {
          console.error('[Telegram.WebApp] Bottom button text is too long', text);
          throw Error('WebAppBottomButtonParamInvalid');
        }
        buttonText = text;
      }
      if (typeof params.color !== 'undefined') {
        if (params.color === false ||
            params.color === null) {
          buttonColor = false;
        } else {
          var color = parseColorToHex(params.color);
          if (!color) {
            console.error('[Telegram.WebApp] Bottom button color format is invalid', params.color);
            throw Error('WebAppBottomButtonParamInvalid');
          }
          buttonColor = color;
        }
      }
      if (typeof params.text_color !== 'undefined') {
        if (params.text_color === false ||
            params.text_color === null) {
          buttonTextColor = false;
        } else {
          var text_color = parseColorToHex(params.text_color);
          if (!text_color) {
            console.error('[Telegram.WebApp] Bottom button text color format is invalid', params.text_color);
            throw Error('WebAppBottomButtonParamInvalid');
          }
          buttonTextColor = text_color;
        }
      }
      if (typeof params.is_visible !== 'undefined') {
        if (params.is_visible &&
            !bottomButton.text.length) {
          console.error('[Telegram.WebApp] Bottom button text is required');
          throw Error('WebAppBottomButtonParamInvalid');
        }
        isVisible = !!params.is_visible;
      }
      if (typeof params.has_shine_effect !== 'undefined') {
        hasShineEffect = !!params.has_shine_effect;
      }
      if (!isMainButton && typeof params.position !== 'undefined') {
        if (params.position != 'left' && params.position != 'right' &&
            params.position != 'top' && params.position != 'bottom') {
          console.error('[Telegram.WebApp] Bottom button posiition is invalid', params.position);
          throw Error('WebAppBottomButtonParamInvalid');
        }
        buttonPosition = params.position;
      }
      if (typeof params.is_active !== 'undefined') {
        isActive = !!params.is_active;
      }
      updateButton();
      return bottomButton;
    }

    bottomButton.setText = function(text) {
      return bottomButton.setParams({text: text});
    };
    bottomButton.onClick = function(callback) {
      onWebViewEvent(webViewEventName, callback);
      return bottomButton;
    };
    bottomButton.offClick = function(callback) {
      offWebViewEvent(webViewEventName, callback);
      return bottomButton;
    };
    bottomButton.show = function() {
      return bottomButton.setParams({is_visible: true});
    };
    bottomButton.hide = function() {
      return bottomButton.setParams({is_visible: false});
    };
    bottomButton.enable = function() {
      return bottomButton.setParams({is_active: true});
    };
    bottomButton.disable = function() {
      return bottomButton.setParams({is_active: false});
    };
    bottomButton.showProgress = function(leaveActive) {
      isActive = !!leaveActive;
      isProgressVisible = true;
      updateButton();
      return bottomButton;
    };
    bottomButton.hideProgress = function() {
      if (!bottomButton.isActive) {
        isActive = true;
      }
      isProgressVisible = false;
      updateButton();
      return bottomButton;
    }
    bottomButton.setParams = setParams;
    return bottomButton;
  };
  var MainButton = BottomButtonConstructor('main');
  var SecondaryButton = BottomButtonConstructor('secondary');

  var SettingsButton = (function() {
    var isVisible = false;

    var settingsButton = {};
    Object.defineProperty(settingsButton, 'isVisible', {
      set: function(val){ setParams({is_visible: val}); },
      get: function(){ return isVisible; },
      enumerable: true
    });

    var curButtonState = null;

    WebView.onEvent('settings_button_pressed', onSettingsButtonPressed);

    function onSettingsButtonPressed() {
      receiveWebViewEvent('settingsButtonClicked');
    }

    function buttonParams() {
      return {is_visible: isVisible};
    }

    function buttonState(btn_params) {
      if (typeof btn_params === 'undefined') {
        btn_params = buttonParams();
      }
      return JSON.stringify(btn_params);
    }

    function buttonCheckVersion() {
      if (!versionAtLeast('6.10')) {
        console.warn('[Telegram.WebApp] SettingsButton is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    function updateButton() {
      var btn_params = buttonParams();
      var btn_state = buttonState(btn_params);
      if (curButtonState === btn_state) {
        return;
      }
      curButtonState = btn_state;
      WebView.postEvent('web_app_setup_settings_button', false, btn_params);
    }

    function setParams(params) {
      if (!buttonCheckVersion()) {
        return settingsButton;
      }
      if (typeof params.is_visible !== 'undefined') {
        isVisible = !!params.is_visible;
      }
      updateButton();
      return settingsButton;
    }

    settingsButton.onClick = function(callback) {
      if (buttonCheckVersion()) {
        onWebViewEvent('settingsButtonClicked', callback);
      }
      return settingsButton;
    };
    settingsButton.offClick = function(callback) {
      if (buttonCheckVersion()) {
        offWebViewEvent('settingsButtonClicked', callback);
      }
      return settingsButton;
    };
    settingsButton.show = function() {
      return setParams({is_visible: true});
    };
    settingsButton.hide = function() {
      return setParams({is_visible: false});
    };
    return settingsButton;
  })();

  var HapticFeedback = (function() {
    var hapticFeedback = {};

    function triggerFeedback(params) {
      if (!versionAtLeast('6.1')) {
        console.warn('[Telegram.WebApp] HapticFeedback is not supported in version ' + webAppVersion);
        return hapticFeedback;
      }
      if (params.type == 'impact') {
        if (params.impact_style != 'light' &&
            params.impact_style != 'medium' &&
            params.impact_style != 'heavy' &&
            params.impact_style != 'rigid' &&
            params.impact_style != 'soft') {
          console.error('[Telegram.WebApp] Haptic impact style is invalid', params.impact_style);
          throw Error('WebAppHapticImpactStyleInvalid');
        }
      } else if (params.type == 'notification') {
        if (params.notification_type != 'error' &&
            params.notification_type != 'success' &&
            params.notification_type != 'warning') {
          console.error('[Telegram.WebApp] Haptic notification type is invalid', params.notification_type);
          throw Error('WebAppHapticNotificationTypeInvalid');
        }
      } else if (params.type == 'selection_change') {
        // no params needed
      } else {
        console.error('[Telegram.WebApp] Haptic feedback type is invalid', params.type);
        throw Error('WebAppHapticFeedbackTypeInvalid');
      }
      WebView.postEvent('web_app_trigger_haptic_feedback', false, params);
      return hapticFeedback;
    }

    hapticFeedback.impactOccurred = function(style) {
      return triggerFeedback({type: 'impact', impact_style: style});
    };
    hapticFeedback.notificationOccurred = function(type) {
      return triggerFeedback({type: 'notification', notification_type: type});
    };
    hapticFeedback.selectionChanged = function() {
      return triggerFeedback({type: 'selection_change'});
    };
    return hapticFeedback;
  })();

  var CloudStorage = (function() {
    var cloudStorage = {};

    function invokeStorageMethod(method, params, callback) {
      if (!versionAtLeast('6.9')) {
        console.error('[Telegram.WebApp] CloudStorage is not supported in version ' + webAppVersion);
        throw Error('WebAppMethodUnsupported');
      }
      invokeCustomMethod(method, params, callback);
      return cloudStorage;
    }

    cloudStorage.setItem = function(key, value, callback) {
      return invokeStorageMethod('saveStorageValue', {key: key, value: value}, callback);
    };
    cloudStorage.getItem = function(key, callback) {
      return cloudStorage.getItems([key], callback ? function(err, res) {
        if (err) callback(err);
        else callback(null, res[key]);
      } : null);
    };
    cloudStorage.getItems = function(keys, callback) {
      return invokeStorageMethod('getStorageValues', {keys: keys}, callback);
    };
    cloudStorage.removeItem = function(key, callback) {
      return cloudStorage.removeItems([key], callback);
    };
    cloudStorage.removeItems = function(keys, callback) {
      return invokeStorageMethod('deleteStorageValues', {keys: keys}, callback);
    };
    cloudStorage.getKeys = function(callback) {
      return invokeStorageMethod('getStorageKeys', {}, callback);
    };
    return cloudStorage;
  })();

  var DeviceStorage = (function() {
    var deviceStorage = {};

    WebView.onEvent('device_storage_key_saved',  onDeviceStorageEvent);
    WebView.onEvent('device_storage_key_received', onDeviceStorageEvent);
    WebView.onEvent('device_storage_cleared',  onDeviceStorageEvent);
    WebView.onEvent('device_storage_failed',  onDeviceStorageEvent);

    function onDeviceStorageEvent(eventType, eventData) {
      if (eventData.req_id && webAppCallbacks[eventData.req_id]) {
        var requestData = webAppCallbacks[eventData.req_id];
        delete webAppCallbacks[eventData.req_id];
        var res = null, err = null;
        if (eventType == 'device_storage_failed') {
          err = eventData.error || 'UNKNOWN_ERROR';
        } else if (eventType == 'device_storage_key_received') {
          res = eventData.value;
        } else {
          res = true;
        }
        if (requestData.callback) {
          requestData.callback(err, res);
        }
      }
    }

    function invokeStorageMethod(method, params, callback) {
      if (!versionAtLeast('9.0')) {
        console.error('[Telegram.WebApp] DeviceStorage is not supported in version ' + webAppVersion);
        throw Error('WebAppMethodUnsupported');
      }
      var req_id = generateCallbackId(16);
      var req_params = {req_id: req_id};
      for (var k in params) {
        req_params[k] = params[k];
      }
      webAppCallbacks[req_id] = {
        callback: callback
      };
      WebView.postEvent(method, false, req_params);
      return deviceStorage;
    }

    deviceStorage.setItem = function(key, value, callback) {
      return invokeStorageMethod('web_app_device_storage_save_key', {key: key, value: value}, callback);
    };
    deviceStorage.getItem = function(key, callback) {
      return invokeStorageMethod('web_app_device_storage_get_key', {key: key}, callback);
    };
    deviceStorage.removeItem = function(key, callback) {
      return invokeStorageMethod('web_app_device_storage_save_key', {key: key, value: null}, callback);
    };
    deviceStorage.clear = function(callback) {
      return invokeStorageMethod('web_app_device_storage_clear', {}, callback);
    };
    return deviceStorage;
  })();

  var SecureStorage = (function() {
    var secureStorage = {};

    WebView.onEvent('secure_storage_key_saved',  onSecureStorageEvent);
    WebView.onEvent('secure_storage_key_received', onSecureStorageEvent);
    WebView.onEvent('secure_storage_key_restored', onSecureStorageEvent);
    WebView.onEvent('secure_storage_cleared',  onSecureStorageEvent);
    WebView.onEvent('secure_storage_failed',  onSecureStorageEvent);

    function onSecureStorageEvent(eventType, eventData) {
      if (eventData.req_id && webAppCallbacks[eventData.req_id]) {
        var requestData = webAppCallbacks[eventData.req_id];
        delete webAppCallbacks[eventData.req_id];
        var res = null, err = null, can_restore = null;
        if (eventType == 'secure_storage_failed') {
          err = eventData.error || 'UNKNOWN_ERROR';
        } else if (eventType == 'secure_storage_key_received') {
          res = eventData.value;
          if (eventData.can_restore) {
            can_restore = true;
          }
        } else if (eventType == 'secure_storage_key_restored') {
          res = eventData.value;
        } else {
          res = true;
        }
        if (requestData.callback) {
          requestData.callback(err, res, can_restore);
        }
      }
    }

    function invokeStorageMethod(method, params, callback) {
      if (!versionAtLeast('9.0')) {
        console.error('[Telegram.WebApp] SecureStorage is not supported in version ' + webAppVersion);
        throw Error('WebAppMethodUnsupported');
      }
      var req_id = generateCallbackId(16);
      var req_params = {req_id: req_id};
      for (var k in params) {
        req_params[k] = params[k];
      }
      webAppCallbacks[req_id] = {
        callback: callback
      };
      WebView.postEvent(method, false, req_params);
      return secureStorage;
    }

    secureStorage.setItem = function(key, value, callback) {
      return invokeStorageMethod('web_app_secure_storage_save_key', {key: key, value: value}, callback);
    };
    secureStorage.getItem = function(key, callback) {
      return invokeStorageMethod('web_app_secure_storage_get_key', {key: key}, callback);
    };
    secureStorage.restoreItem = function(key, callback) {
      return invokeStorageMethod('web_app_secure_storage_restore_key', {key: key}, callback);
    };
    secureStorage.removeItem = function(key, callback) {
      return invokeStorageMethod('web_app_secure_storage_save_key', {key: key, value: null}, callback);
    };
    secureStorage.clear = function(callback) {
      return invokeStorageMethod('web_app_secure_storage_clear', {}, callback);
    };
    return secureStorage;
  })();

  var BiometricManager = (function() {
    var isInited = false;
    var isBiometricAvailable = false;
    var biometricType = 'unknown';
    var isAccessRequested = false;
    var isAccessGranted = false;
    var isBiometricTokenSaved = false;
    var deviceId = '';

    var biometricManager = {};
    Object.defineProperty(biometricManager, 'isInited', {
      get: function(){ return isInited; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'isBiometricAvailable', {
      get: function(){ return isInited && isBiometricAvailable; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'biometricType', {
      get: function(){ return biometricType || 'unknown'; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'isAccessRequested', {
      get: function(){ return isAccessRequested; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'isAccessGranted', {
      get: function(){ return isAccessRequested && isAccessGranted; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'isBiometricTokenSaved', {
      get: function(){ return isBiometricTokenSaved; },
      enumerable: true
    });
    Object.defineProperty(biometricManager, 'deviceId', {
      get: function(){ return deviceId || ''; },
      enumerable: true
    });

    var initRequestState = {callbacks: []};
    var accessRequestState = false;
    var authRequestState = false;
    var tokenRequestState = false;

    WebView.onEvent('biometry_info_received',  onBiometryInfoReceived);
    WebView.onEvent('biometry_auth_requested', onBiometryAuthRequested);
    WebView.onEvent('biometry_token_updated',  onBiometryTokenUpdated);

    function onBiometryInfoReceived(eventType, eventData) {
      isInited = true;
      if (eventData.available) {
        isBiometricAvailable = true;
        biometricType = eventData.type || 'unknown';
        if (eventData.access_requested) {
          isAccessRequested = true;
          isAccessGranted = !!eventData.access_granted;
          isBiometricTokenSaved = !!eventData.token_saved;
        } else {
          isAccessRequested = false;
          isAccessGranted = false;
          isBiometricTokenSaved = false;
        }
      } else {
        isBiometricAvailable = false;
        biometricType = 'unknown';
        isAccessRequested = false;
        isAccessGranted = false;
        isBiometricTokenSaved = false;
      }
      deviceId = eventData.device_id || '';

      if (initRequestState.callbacks.length > 0) {
        for (var i = 0; i < initRequestState.callbacks.length; i++) {
          var callback = initRequestState.callbacks[i];
          callback();
        }
        initRequestState.callbacks = [];
      }
      if (accessRequestState) {
        var state = accessRequestState;
        accessRequestState = false;
        if (state.callback) {
          state.callback(isAccessGranted);
        }
      }
      receiveWebViewEvent('biometricManagerUpdated');
    }
    function onBiometryAuthRequested(eventType, eventData) {
      var isAuthenticated = (eventData.status == 'authorized'),
          biometricToken = eventData.token || '';
      if (authRequestState) {
        var state = authRequestState;
        authRequestState = false;
        if (state.callback) {
          state.callback(isAuthenticated, isAuthenticated ? biometricToken : null);
        }
      }
      receiveWebViewEvent('biometricAuthRequested', isAuthenticated ? {
        isAuthenticated: true,
        biometricToken: biometricToken
      } : {
        isAuthenticated: false
      });
    }
    function onBiometryTokenUpdated(eventType, eventData) {
      var applied = false;
      if (isBiometricAvailable &&
          isAccessRequested) {
        if (eventData.status == 'updated') {
          isBiometricTokenSaved = true;
          applied = true;
        }
        else if (eventData.status == 'removed') {
          isBiometricTokenSaved = false;
          applied = true;
        }
      }
      if (tokenRequestState) {
        var state = tokenRequestState;
        tokenRequestState = false;
        if (state.callback) {
          state.callback(applied);
        }
      }
      receiveWebViewEvent('biometricTokenUpdated', {
        isUpdated: applied
      });
    }

    function checkVersion() {
      if (!versionAtLeast('7.2')) {
        console.warn('[Telegram.WebApp] BiometricManager is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    function checkInit() {
      if (!isInited) {
        console.error('[Telegram.WebApp] BiometricManager should be inited before using.');
        throw Error('WebAppBiometricManagerNotInited');
      }
      return true;
    }

    biometricManager.init = function(callback) {
      if (!checkVersion()) {
        return biometricManager;
      }
      if (isInited) {
        return biometricManager;
      }
      if (callback) {
        initRequestState.callbacks.push(callback);
      }
      WebView.postEvent('web_app_biometry_get_info', false);
      return biometricManager;
    };
    biometricManager.requestAccess = function(params, callback) {
      if (!checkVersion()) {
        return biometricManager;
      }
      checkInit();
      if (!isBiometricAvailable) {
        console.error('[Telegram.WebApp] Biometrics is not available on this device.');
        throw Error('WebAppBiometricManagerBiometricsNotAvailable');
      }
      if (accessRequestState) {
        console.error('[Telegram.WebApp] Access is already requested');
        throw Error('WebAppBiometricManagerAccessRequested');
      }
      var popup_params = {};
      if (typeof params.reason !== 'undefined') {
        var reason = strTrim(params.reason);
        if (reason.length > 128) {
          console.error('[Telegram.WebApp] Biometric reason is too long', reason);
          throw Error('WebAppBiometricRequestAccessParamInvalid');
        }
        if (reason.length > 0) {
          popup_params.reason = reason;
        }
      }

      accessRequestState = {
        callback: callback
      };
      WebView.postEvent('web_app_biometry_request_access', false, popup_params);
      return biometricManager;
    };
    biometricManager.authenticate = function(params, callback) {
      if (!checkVersion()) {
        return biometricManager;
      }
      checkInit();
      if (!isBiometricAvailable) {
        console.error('[Telegram.WebApp] Biometrics is not available on this device.');
        throw Error('WebAppBiometricManagerBiometricsNotAvailable');
      }
      if (!isAccessGranted) {
        console.error('[Telegram.WebApp] Biometric access was not granted by the user.');
        throw Error('WebAppBiometricManagerBiometricAccessNotGranted');
      }
      if (authRequestState) {
        console.error('[Telegram.WebApp] Authentication request is already in progress.');
        throw Error('WebAppBiometricManagerAuthenticationRequested');
      }
      var popup_params = {};
      if (typeof params.reason !== 'undefined') {
        var reason = strTrim(params.reason);
        if (reason.length > 128) {
          console.error('[Telegram.WebApp] Biometric reason is too long', reason);
          throw Error('WebAppBiometricRequestAccessParamInvalid');
        }
        if (reason.length > 0) {
          popup_params.reason = reason;
        }
      }

      authRequestState = {
        callback: callback
      };
      WebView.postEvent('web_app_biometry_request_auth', false, popup_params);
      return biometricManager;
    };
    biometricManager.updateBiometricToken = function(token, callback) {
      if (!checkVersion()) {
        return biometricManager;
      }
      token = token || '';
      if (token.length > 1024) {
        console.error('[Telegram.WebApp] Token is too long', token);
        throw Error('WebAppBiometricManagerTokenInvalid');
      }
      checkInit();
      if (!isBiometricAvailable) {
        console.error('[Telegram.WebApp] Biometrics is not available on this device.');
        throw Error('WebAppBiometricManagerBiometricsNotAvailable');
      }
      if (!isAccessGranted) {
        console.error('[Telegram.WebApp] Biometric access was not granted by the user.');
        throw Error('WebAppBiometricManagerBiometricAccessNotGranted');
      }
      if (tokenRequestState) {
        console.error('[Telegram.WebApp] Token request is already in progress.');
        throw Error('WebAppBiometricManagerTokenUpdateRequested');
      }
      tokenRequestState = {
        callback: callback
      };
      WebView.postEvent('web_app_biometry_update_token', false, {token: token});
      return biometricManager;
    };
    biometricManager.openSettings = function() {
      if (!checkVersion()) {
        return biometricManager;
      }
      checkInit();
      if (!isBiometricAvailable) {
        console.error('[Telegram.WebApp] Biometrics is not available on this device.');
        throw Error('WebAppBiometricManagerBiometricsNotAvailable');
      }
      if (!isAccessRequested) {
        console.error('[Telegram.WebApp] Biometric access was not requested yet.');
        throw Error('WebAppBiometricManagerBiometricsAccessNotRequested');
      }
      if (isAccessGranted) {
        console.warn('[Telegram.WebApp] Biometric access was granted by the user, no need to go to settings.');
        return biometricManager;
      }
      WebView.postEvent('web_app_biometry_open_settings', false);
      return biometricManager;
    };
    return biometricManager;
  })();

  var LocationManager = (function() {
    var isInited = false;
    var isLocationAvailable = false;
    var isAccessRequested = false;
    var isAccessGranted = false;

    var locationManager = {};
    Object.defineProperty(locationManager, 'isInited', {
      get: function(){ return isInited; },
      enumerable: true
    });
    Object.defineProperty(locationManager, 'isLocationAvailable', {
      get: function(){ return isInited && isLocationAvailable; },
      enumerable: true
    });
    Object.defineProperty(locationManager, 'isAccessRequested', {
      get: function(){ return isAccessRequested; },
      enumerable: true
    });
    Object.defineProperty(locationManager, 'isAccessGranted', {
      get: function(){ return isAccessRequested && isAccessGranted; },
      enumerable: true
    });

    var initRequestState = {callbacks: []};
    var getRequestState = {callbacks: []};

    WebView.onEvent('location_checked',  onLocationChecked);
    WebView.onEvent('location_requested', onLocationRequested);

    function onLocationChecked(eventType, eventData) {
      isInited = true;
      if (eventData.available) {
        isLocationAvailable = true;
        if (eventData.access_requested) {
          isAccessRequested = true;
          isAccessGranted = !!eventData.access_granted;
        } else {
          isAccessRequested = false;
          isAccessGranted = false;
        }
      } else {
        isLocationAvailable = false;
        isAccessRequested = false;
        isAccessGranted = false;
      }

      if (initRequestState.callbacks.length > 0) {
        for (var i = 0; i < initRequestState.callbacks.length; i++) {
          var callback = initRequestState.callbacks[i];
          callback();
        }
        initRequestState.callbacks = [];
      }
      receiveWebViewEvent('locationManagerUpdated');
    }
    function onLocationRequested(eventType, eventData) {
      if (!eventData.available) {
        locationData = null;
      } else {
        var locationData = {
          latitude: eventData.latitude,
          longitude: eventData.longitude,
          altitude: null,
          course: null,
          speed: null,
          horizontal_accuracy: null,
          vertical_accuracy: null,
          course_accuracy: null,
          speed_accuracy: null,
        };
        if (typeof eventData.altitude !== 'undefined' && eventData.altitude !== null) {
          locationData.altitude = eventData.altitude;
        }
        if (typeof eventData.course !== 'undefined' && eventData.course !== null) {
          locationData.course = eventData.course % 360;
        }
        if (typeof eventData.speed !== 'undefined' && eventData.speed !== null) {
          locationData.speed = eventData.speed;
        }
        if (typeof eventData.horizontal_accuracy !== 'undefined' && eventData.horizontal_accuracy !== null) {
          locationData.horizontal_accuracy = eventData.horizontal_accuracy;
        }
        if (typeof eventData.vertical_accuracy !== 'undefined' && eventData.vertical_accuracy !== null) {
          locationData.vertical_accuracy = eventData.vertical_accuracy;
        }
        if (typeof eventData.course_accuracy !== 'undefined' && eventData.course_accuracy !== null) {
          locationData.course_accuracy = eventData.course_accuracy;
        }
        if (typeof eventData.speed_accuracy !== 'undefined' && eventData.speed_accuracy !== null) {
          locationData.speed_accuracy = eventData.speed_accuracy;
        }
      }
      if (!eventData.available ||
          !isLocationAvailable ||
          !isAccessRequested ||
          !isAccessGranted) {
        initRequestState.callbacks.push(function() {
          locationResponse(locationData);
        });
        WebView.postEvent('web_app_check_location', false);
      } else {
        locationResponse(locationData);
      }
    }
    function locationResponse(response) {
      if (getRequestState.callbacks.length > 0) {
        for (var i = 0; i < getRequestState.callbacks.length; i++) {
          var callback = getRequestState.callbacks[i];
          callback(response);
        }
        getRequestState.callbacks = [];
      }
      if (response !== null) {
        receiveWebViewEvent('locationRequested', {
          locationData: response
        });
      }
    }

    function checkVersion() {
      if (!versionAtLeast('8.0')) {
        console.warn('[Telegram.WebApp] LocationManager is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    function checkInit() {
      if (!isInited) {
        console.error('[Telegram.WebApp] LocationManager should be inited before using.');
        throw Error('WebAppLocationManagerNotInited');
      }
      return true;
    }

    locationManager.init = function(callback) {
      if (!checkVersion()) {
        return locationManager;
      }
      if (isInited) {
        return locationManager;
      }
      if (callback) {
        initRequestState.callbacks.push(callback);
      }
      WebView.postEvent('web_app_check_location', false);
      return locationManager;
    };
    locationManager.getLocation = function(callback) {
      if (!checkVersion()) {
        return locationManager;
      }
      checkInit();
      if (!isLocationAvailable) {
        console.error('[Telegram.WebApp] Location is not available on this device.');
        throw Error('WebAppLocationManagerLocationNotAvailable');
      }

      getRequestState.callbacks.push(callback);
      WebView.postEvent('web_app_request_location');
      return locationManager;
    };
    locationManager.openSettings = function() {
      if (!checkVersion()) {
        return locationManager;
      }
      checkInit();
      if (!isLocationAvailable) {
        console.error('[Telegram.WebApp] Location is not available on this device.');
        throw Error('WebAppLocationManagerLocationNotAvailable');
      }
      if (!isAccessRequested) {
        console.error('[Telegram.WebApp] Location access was not requested yet.');
        throw Error('WebAppLocationManagerLocationAccessNotRequested');
      }
      if (isAccessGranted) {
        console.warn('[Telegram.WebApp] Location access was granted by the user, no need to go to settings.');
        return locationManager;
      }
      WebView.postEvent('web_app_open_location_settings', false);
      return locationManager;
    };
    return locationManager;
  })();

  var Accelerometer = (function() {
    var isStarted = false;
    var valueX = null, valueY = null, valueZ = null;
    var startCallbacks = [], stopCallbacks = [];

    var accelerometer = {};
    Object.defineProperty(accelerometer, 'isStarted', {
      get: function(){ return isStarted; },
      enumerable: true
    });
    Object.defineProperty(accelerometer, 'x', {
      get: function(){ return valueX; },
      enumerable: true
    });
    Object.defineProperty(accelerometer, 'y', {
      get: function(){ return valueY; },
      enumerable: true
    });
    Object.defineProperty(accelerometer, 'z', {
      get: function(){ return valueZ; },
      enumerable: true
    });

    WebView.onEvent('accelerometer_started', onAccelerometerStarted);
    WebView.onEvent('accelerometer_stopped', onAccelerometerStopped);
    WebView.onEvent('accelerometer_changed', onAccelerometerChanged);
    WebView.onEvent('accelerometer_failed',  onAccelerometerFailed);

    function onAccelerometerStarted(eventType, eventData) {
      isStarted = true;
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(true);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('accelerometerStarted');
    }
    function onAccelerometerStopped(eventType, eventData) {
      isStarted = false;
      if (stopCallbacks.length > 0) {
        for (var i = 0; i < stopCallbacks.length; i++) {
          var callback = stopCallbacks[i];
          callback(true);
        }
        stopCallbacks = [];
      }
      receiveWebViewEvent('accelerometerStopped');
    }
    function onAccelerometerChanged(eventType, eventData) {
      valueX = eventData.x;
      valueY = eventData.y;
      valueZ = eventData.z;
      receiveWebViewEvent('accelerometerChanged');
    }
    function onAccelerometerFailed(eventType, eventData) {
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(false);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('accelerometerFailed', {
        error: eventData.error
      });
    }

    function checkVersion() {
      if (!versionAtLeast('8.0')) {
        console.warn('[Telegram.WebApp] Accelerometer is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    accelerometer.start = function(params, callback) {
      params = params || {};
      if (!checkVersion()) {
        return accelerometer;
      }
      var req_params = {};
      var refresh_rate = parseInt(params.refresh_rate || 1000);
      if (isNaN(refresh_rate) || refresh_rate < 20 || refresh_rate > 1000) {
        console.warn('[Telegram.WebApp] Accelerometer refresh_rate is invalid', refresh_rate);
      } else {
        req_params.refresh_rate = refresh_rate;
      }

      if (callback) {
        startCallbacks.push(callback);
      }
      WebView.postEvent('web_app_start_accelerometer', false, req_params);
      return accelerometer;
    };
    accelerometer.stop = function(callback) {
      if (!checkVersion()) {
        return accelerometer;
      }
      if (callback) {
        stopCallbacks.push(callback);
      }
      WebView.postEvent('web_app_stop_accelerometer');
      return accelerometer;
    };
    return accelerometer;
  })();

  var DeviceOrientation = (function() {
    var isStarted = false;
    var valueAlpha = null, valueBeta = null, valueGamma = null, valueAbsolute = false;
    var startCallbacks = [], stopCallbacks = [];

    var deviceOrientation = {};
    Object.defineProperty(deviceOrientation, 'isStarted', {
      get: function(){ return isStarted; },
      enumerable: true
    });
    Object.defineProperty(deviceOrientation, 'absolute', {
      get: function(){ return valueAbsolute; },
      enumerable: true
    });
    Object.defineProperty(deviceOrientation, 'alpha', {
      get: function(){ return valueAlpha; },
      enumerable: true
    });
    Object.defineProperty(deviceOrientation, 'beta', {
      get: function(){ return valueBeta; },
      enumerable: true
    });
    Object.defineProperty(deviceOrientation, 'gamma', {
      get: function(){ return valueGamma; },
      enumerable: true
    });

    WebView.onEvent('device_orientation_started',  onDeviceOrientationStarted);
    WebView.onEvent('device_orientation_stopped',  onDeviceOrientationStopped);
    WebView.onEvent('device_orientation_changed', onDeviceOrientationChanged);
    WebView.onEvent('device_orientation_failed',  onDeviceOrientationFailed);

    function onDeviceOrientationStarted(eventType, eventData) {
      isStarted = true;
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(true);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('deviceOrientationStarted');
    }
    function onDeviceOrientationStopped(eventType, eventData) {
      isStarted = false;
      if (stopCallbacks.length > 0) {
        for (var i = 0; i < stopCallbacks.length; i++) {
          var callback = stopCallbacks[i];
          callback(true);
        }
        stopCallbacks = [];
      }
      receiveWebViewEvent('deviceOrientationStopped');
    }
    function onDeviceOrientationChanged(eventType, eventData) {
      valueAbsolute = !!eventData.absolute;
      valueAlpha = eventData.alpha;
      valueBeta  = eventData.beta;
      valueGamma = eventData.gamma;
      receiveWebViewEvent('deviceOrientationChanged');
    }
    function onDeviceOrientationFailed(eventType, eventData) {
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(false);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('deviceOrientationFailed', {
        error: eventData.error
      });
    }

    function checkVersion() {
      if (!versionAtLeast('8.0')) {
        console.warn('[Telegram.WebApp] DeviceOrientation is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    deviceOrientation.start = function(params, callback) {
      params = params || {};
      if (!checkVersion()) {
        return deviceOrientation;
      }
      var req_params = {};
      var refresh_rate = parseInt(params.refresh_rate || 1000);
      if (isNaN(refresh_rate) || refresh_rate < 20 || refresh_rate > 1000) {
        console.warn('[Telegram.WebApp] DeviceOrientation refresh_rate is invalid', refresh_rate);
      } else {
        req_params.refresh_rate = refresh_rate;
      }
      req_params.need_absolute = !!params.need_absolute;

      if (callback) {
        startCallbacks.push(callback);
      }
      WebView.postEvent('web_app_start_device_orientation', false, req_params);
      return deviceOrientation;
    };
    deviceOrientation.stop = function(callback) {
      if (!checkVersion()) {
        return deviceOrientation;
      }
      if (callback) {
        stopCallbacks.push(callback);
      }
      WebView.postEvent('web_app_stop_device_orientation');
      return deviceOrientation;
    };
    return deviceOrientation;
  })();

  var Gyroscope = (function() {
    var isStarted = false;
    var valueX = null, valueY = null, valueZ = null;
    var startCallbacks = [], stopCallbacks = [];

    var gyroscope = {};
    Object.defineProperty(gyroscope, 'isStarted', {
      get: function(){ return isStarted; },
      enumerable: true
    });
    Object.defineProperty(gyroscope, 'x', {
      get: function(){ return valueX; },
      enumerable: true
    });
    Object.defineProperty(gyroscope, 'y', {
      get: function(){ return valueY; },
      enumerable: true
    });
    Object.defineProperty(gyroscope, 'z', {
      get: function(){ return valueZ; },
      enumerable: true
    });

    WebView.onEvent('gyroscope_started',  onGyroscopeStarted);
    WebView.onEvent('gyroscope_stopped',  onGyroscopeStopped);
    WebView.onEvent('gyroscope_changed', onGyroscopeChanged);
    WebView.onEvent('gyroscope_failed',  onGyroscopeFailed);

    function onGyroscopeStarted(eventType, eventData) {
      isStarted = true;
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(true);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('gyroscopeStarted');
    }
    function onGyroscopeStopped(eventType, eventData) {
      isStarted = false;
      if (stopCallbacks.length > 0) {
        for (var i = 0; i < stopCallbacks.length; i++) {
          var callback = stopCallbacks[i];
          callback(true);
        }
        stopCallbacks = [];
      }
      receiveWebViewEvent('gyroscopeStopped');
    }
    function onGyroscopeChanged(eventType, eventData) {
      valueX = eventData.x;
      valueY = eventData.y;
      valueZ = eventData.z;
      receiveWebViewEvent('gyroscopeChanged');
    }
    function onGyroscopeFailed(eventType, eventData) {
      if (startCallbacks.length > 0) {
        for (var i = 0; i < startCallbacks.length; i++) {
          var callback = startCallbacks[i];
          callback(false);
        }
        startCallbacks = [];
      }
      receiveWebViewEvent('gyroscopeFailed', {
        error: eventData.error
      });
    }

    function checkVersion() {
      if (!versionAtLeast('8.0')) {
        console.warn('[Telegram.WebApp] Gyroscope is not supported in version ' + webAppVersion);
        return false;
      }
      return true;
    }

    gyroscope.start = function(params, callback) {
      params = params || {};
      if (!checkVersion()) {
        return gyroscope;
      }
      var req_params = {};
      var refresh_rate = parseInt(params.refresh_rate || 1000);
      if (isNaN(refresh_rate) || refresh_rate < 20 || refresh_rate > 1000) {
        console.warn('[Telegram.WebApp] Gyroscope refresh_rate is invalid', refresh_rate);
      } else {
        req_params.refresh_rate = refresh_rate;
      }

      if (callback) {
        startCallbacks.push(callback);
      }
      WebView.postEvent('web_app_start_gyroscope', false, req_params);
      return gyroscope;
    };
    gyroscope.stop = function(callback) {
      if (!checkVersion()) {
        return gyroscope;
      }
      if (callback) {
        stopCallbacks.push(callback);
      }
      WebView.postEvent('web_app_stop_gyroscope');
      return gyroscope;
    };
    return gyroscope;
  })();

  var webAppInvoices = {};
  function onInvoiceClosed(eventType, eventData) {
    if (eventData.slug && webAppInvoices[eventData.slug]) {
      var invoiceData = webAppInvoices[eventData.slug];
      delete webAppInvoices[eventData.slug];
      if (invoiceData.callback) {
        invoiceData.callback(eventData.status);
      }
      receiveWebViewEvent('invoiceClosed', {
        url: invoiceData.url,
        status: eventData.status
      });
    }
  }

  var webAppPopupOpened = false;
  function onPopupClosed(eventType, eventData) {
    if (webAppPopupOpened) {
      var popupData = webAppPopupOpened;
      webAppPopupOpened = false;
      var button_id = null;
      if (typeof eventData.button_id !== 'undefined') {
        button_id = eventData.button_id;
      }
      if (popupData.callback) {
        popupData.callback(button_id);
      }
      receiveWebViewEvent('popupClosed', {
        button_id: button_id
      });
    }
  }

  var webAppScanQrPopupOpened = false;
  function onQrTextReceived(eventType, eventData) {
    if (webAppScanQrPopupOpened) {
      var popupData = webAppScanQrPopupOpened;
      var data = null;
      if (typeof eventData.data !== 'undefined') {
        data = eventData.data;
      }
      if (popupData.callback) {
        if (popupData.callback(data)) {
          webAppScanQrPopupOpened = false;
          WebView.postEvent('web_app_close_scan_qr_popup', false);
        }
      }
      receiveWebViewEvent('qrTextReceived', {
        data: data
      });
    }
  }
  function onScanQrPopupClosed(eventType, eventData) {
    webAppScanQrPopupOpened = false;
    receiveWebViewEvent('scanQrPopupClosed');
  }

  function onClipboardTextReceived(eventType, eventData) {
    if (eventData.req_id && webAppCallbacks[eventData.req_id]) {
      var requestData = webAppCallbacks[eventData.req_id];
      delete webAppCallbacks[eventData.req_id];
      var data = null;
      if (typeof eventData.data !== 'undefined') {
        data = eventData.data;
      }
      if (requestData.callback) {
        requestData.callback(data);
      }
      receiveWebViewEvent('clipboardTextReceived', {
        data: data
      });
    }
  }

  var WebAppWriteAccessRequested = false;
  function onWriteAccessRequested(eventType, eventData) {
    if (WebAppWriteAccessRequested) {
      var requestData = WebAppWriteAccessRequested;
      WebAppWriteAccessRequested = false;
      if (requestData.callback) {
        requestData.callback(eventData.status == 'allowed');
      }
      receiveWebViewEvent('writeAccessRequested', {
        status: eventData.status
      });
    }
  }

  function getRequestedContact(callback, timeout) {
    var reqTo, fallbackTo, reqDelay = 0;
    var reqInvoke = function() {
      invokeCustomMethod('getRequestedContact', {}, function(err, res) {
        if (res.substr(0, 1) == '"' && res.substr(-1) == '"') { // macos fix
          res = JSON.parse(res);
        }
        if (res && res.length) {
          clearTimeout(fallbackTo);
          callback(res);
        } else {
          reqDelay += 50;
          reqTo = setTimeout(reqInvoke, reqDelay);
        }
      });
    };
    var fallbackInvoke = function() {
      clearTimeout(reqTo);
      callback('');
    };
    fallbackTo = setTimeout(fallbackInvoke, timeout);
    reqInvoke();
  }

  var WebAppContactRequested = false;
  function onPhoneRequested(eventType, eventData) {
    if (WebAppContactRequested) {
      var requestData = WebAppContactRequested;
      WebAppContactRequested = false;
      var requestSent = eventData.status == 'sent';
      var webViewEvent = {
        status: eventData.status
      };
      if (requestSent) {
        getRequestedContact(function(res) {
          if (res && res.length) {
            webViewEvent.response = res;
            webViewEvent.responseUnsafe = Utils.urlParseQueryString(res);
            for (var key in webViewEvent.responseUnsafe) {
              var val = webViewEvent.responseUnsafe[key];
              try {
                if (val.substr(0, 1) == '{' && val.substr(-1) == '}' ||
                    val.substr(0, 1) == '[' && val.substr(-1) == ']') {
                  webViewEvent.responseUnsafe[key] = JSON.parse(val);
                }
              } catch (e) {}
            }
          }
          if (requestData.callback) {
            requestData.callback(requestSent, webViewEvent);
          }
          receiveWebViewEvent('contactRequested', webViewEvent);
        }, 3000);
      } else {
        if (requestData.callback) {
          requestData.callback(requestSent, webViewEvent);
        }
        receiveWebViewEvent('contactRequested', webViewEvent);
      }
    }
  }

  var webAppDownloadFileRequested = false;
  function onFileDownloadRequested(eventType, eventData) {
    if (webAppDownloadFileRequested) {
      var requestData = webAppDownloadFileRequested;
      webAppDownloadFileRequested = false;
      var isDownloading = eventData.status == 'downloading';
      if (requestData.callback) {
        requestData.callback(isDownloading);
      }
      receiveWebViewEvent('fileDownloadRequested', {
        status: isDownloading ? 'downloading' : 'cancelled'
      });
    }
  }

  function onCustomMethodInvoked(eventType, eventData) {
    if (eventData.req_id && webAppCallbacks[eventData.req_id]) {
      var requestData = webAppCallbacks[eventData.req_id];
      delete webAppCallbacks[eventData.req_id];
      var res = null, err = null;
      if (typeof eventData.result !== 'undefined') {
        res = eventData.result;
      }
      if (typeof eventData.error !== 'undefined') {
        err = eventData.error;
      }
      if (requestData.callback) {
        requestData.callback(err, res);
      }
    }
  }

  function invokeCustomMethod(method, params, callback) {
    if (!versionAtLeast('6.9')) {
      console.error('[Telegram.WebApp] Method invokeCustomMethod is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    var req_id = generateCallbackId(16);
    var req_params = {req_id: req_id, method: method, params: params || {}};
    webAppCallbacks[req_id] = {
      callback: callback
    };
    WebView.postEvent('web_app_invoke_custom_method', false, req_params);
  };

  if (!window.Telegram) {
    window.Telegram = {};
  }

  Object.defineProperty(WebApp, 'initData', {
    get: function(){ return webAppInitData; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'initDataUnsafe', {
    get: function(){ return webAppInitDataUnsafe; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'version', {
    get: function(){ return webAppVersion; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'platform', {
    get: function(){ return webAppPlatform; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'colorScheme', {
    get: function(){ return colorScheme; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'themeParams', {
    get: function(){ return themeParams; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isExpanded', {
    get: function(){ return isExpanded; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'viewportHeight', {
    get: function(){ return (viewportHeight === false ? window.innerHeight : viewportHeight) - bottomBarHeight; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'viewportStableHeight', {
    get: function(){ return (viewportStableHeight === false ? window.innerHeight : viewportStableHeight) - bottomBarHeight; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'safeAreaInset', {
    get: function(){ return safeAreaInset; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'contentSafeAreaInset', {
    get: function(){ return contentSafeAreaInset; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isClosingConfirmationEnabled', {
    set: function(val){ setClosingConfirmation(val); },
    get: function(){ return isClosingConfirmationEnabled; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isVerticalSwipesEnabled', {
    set: function(val){ toggleVerticalSwipes(val); },
    get: function(){ return isVerticalSwipesEnabled; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isFullscreen', {
    get: function(){ return webAppIsFullscreen; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isOrientationLocked', {
    set: function(val){ toggleOrientationLock(val); },
    get: function(){ return webAppIsOrientationLocked; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'isActive', {
    get: function(){ return webAppIsActive; },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'headerColor', {
    set: function(val){ setHeaderColor(val); },
    get: function(){ return getHeaderColor(); },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'backgroundColor', {
    set: function(val){ setBackgroundColor(val); },
    get: function(){ return getBackgroundColor(); },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'bottomBarColor', {
    set: function(val){ setBottomBarColor(val); },
    get: function(){ return getBottomBarColor(); },
    enumerable: true
  });
  Object.defineProperty(WebApp, 'BackButton', {
    value: BackButton,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'MainButton', {
    value: MainButton,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'SecondaryButton', {
    value: SecondaryButton,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'SettingsButton', {
    value: SettingsButton,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'HapticFeedback', {
    value: HapticFeedback,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'CloudStorage', {
    value: CloudStorage,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'DeviceStorage', {
    value: DeviceStorage,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'SecureStorage', {
    value: SecureStorage,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'BiometricManager', {
    value: BiometricManager,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'Accelerometer', {
    value: Accelerometer,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'DeviceOrientation', {
    value: DeviceOrientation,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'Gyroscope', {
    value: Gyroscope,
    enumerable: true
  });
  Object.defineProperty(WebApp, 'LocationManager', {
    value: LocationManager,
    enumerable: true
  });
  WebApp.isVersionAtLeast = function(ver) {
    return versionAtLeast(ver);
  };
  WebApp.setHeaderColor = function(color_key) {
    WebApp.headerColor = color_key;
  };
  WebApp.setBackgroundColor = function(color) {
    WebApp.backgroundColor = color;
  };
  WebApp.setBottomBarColor = function(color) {
    WebApp.bottomBarColor = color;
  };
  WebApp.enableClosingConfirmation = function() {
    WebApp.isClosingConfirmationEnabled = true;
  };
  WebApp.disableClosingConfirmation = function() {
    WebApp.isClosingConfirmationEnabled = false;
  };
  WebApp.enableVerticalSwipes = function() {
    WebApp.isVerticalSwipesEnabled = true;
  };
  WebApp.disableVerticalSwipes = function() {
    WebApp.isVerticalSwipesEnabled = false;
  };
  WebApp.lockOrientation = function() {
    WebApp.isOrientationLocked = true;
  };
  WebApp.unlockOrientation = function() {
    WebApp.isOrientationLocked = false;
  };
  WebApp.requestFullscreen = function() {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method requestFullscreen is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    WebView.postEvent('web_app_request_fullscreen');
  };
  WebApp.exitFullscreen = function() {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method exitFullscreen is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    WebView.postEvent('web_app_exit_fullscreen');
  };
  WebApp.addToHomeScreen = function() {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method addToHomeScreen is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    WebView.postEvent('web_app_add_to_home_screen');
  };
  WebApp.checkHomeScreenStatus = function(callback) {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method checkHomeScreenStatus is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (callback) {
      homeScreenCallbacks.push(callback);
    }
    WebView.postEvent('web_app_check_home_screen');
  };
  WebApp.onEvent = function(eventType, callback) {
    onWebViewEvent(eventType, callback);
  };
  WebApp.offEvent = function(eventType, callback) {offWebViewEvent(eventType, callback);
  };
  WebApp.sendData = function (data) {
    if (!data || !data.length) {
      console.error('[Telegram.WebApp] Data is required', data);
      throw Error('WebAppDataInvalid');
    }
    if (byteLength(data) > 4096) {
      console.error('[Telegram.WebApp] Data is too long', data);
      throw Error('WebAppDataInvalid');
    }
    WebView.postEvent('web_app_data_send', false, {data: data});
  };
  WebApp.switchInlineQuery = function (query, choose_chat_types) {
    if (!versionAtLeast('6.6')) {
      console.error('[Telegram.WebApp] Method switchInlineQuery is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (!initParams.tgWebAppBotInline) {
      console.error('[Telegram.WebApp] Inline mode is disabled for this bot. Read more about inline mode: https://core.telegram.org/bots/inline');
      throw Error('WebAppInlineModeDisabled');
    }
    query = query || '';
    if (query.length > 256) {
      console.error('[Telegram.WebApp] Inline query is too long', query);
      throw Error('WebAppInlineQueryInvalid');
    }
    var chat_types = [];
    if (choose_chat_types) {
      if (!Array.isArray(choose_chat_types)) {
        console.error('[Telegram.WebApp] Choose chat types should be an array', choose_chat_types);
        throw Error('WebAppInlineChooseChatTypesInvalid');
      }
      var good_types = {users: 1, bots: 1, groups: 1, channels: 1};
      for (var i = 0; i < choose_chat_types.length; i++) {
        var chat_type = choose_chat_types[i];
        if (!good_types[chat_type]) {
          console.error('[Telegram.WebApp] Choose chat type is invalid', chat_type);
          throw Error('WebAppInlineChooseChatTypeInvalid');
        }
        if (good_types[chat_type] != 2) {
          good_types[chat_type] = 2;
          chat_types.push(chat_type);
        }
      }
    }
    WebView.postEvent('web_app_switch_inline_query', false, {query: query, chat_types: chat_types});
  };
  WebApp.openLink = function (url, options) {
    var a = document.createElement('A');
    a.href = url;
    if (a.protocol != 'http:' &&
        a.protocol != 'https:') {
      console.error('[Telegram.WebApp] Url protocol is not supported', url);
      throw Error('WebAppTgUrlInvalid');
    }
    var url = a.href;
    options = options || {};
    if (versionAtLeast('6.1')) {
      var req_params = {url: url};
      if (versionAtLeast('6.4') && options.try_instant_view) {
        req_params.try_instant_view = true;
      }
      if (versionAtLeast('7.6') && options.try_browser) {
        req_params.try_browser = options.try_browser;
      }
      WebView.postEvent('web_app_open_link', false, req_params);
    } else {
      window.open(url, '_blank');
    }
  };
  WebApp.openTelegramLink = function (url, options) {
    var a = document.createElement('A');
    a.href = url;
    if (a.protocol != 'http:' &&
        a.protocol != 'https:') {
      console.error('[Telegram.WebApp] Url protocol is not supported', url);
      throw Error('WebAppTgUrlInvalid');
    }
    if (!isTmeHostname(a.hostname)) {
      console.error('[Telegram.WebApp] Url host is not supported', url);
      throw Error('WebAppTgUrlInvalid');
    }
    var path_full = a.pathname + a.search;
    options = options || {};
    if (isIframe || versionAtLeast('6.1')) {
      var req_params = {path_full: path_full};
      if (options.force_request) {
        req_params.force_request = true;
      }
      WebView.postEvent('web_app_open_tg_link', false, req_params);
    } else {
      location.href = 'https://t.me' + path_full;
    }
  };
  WebApp.openInvoice = function (url, callback) {
    var a = document.createElement('A'), match, slug;
    a.href = url;
    if (a.protocol != 'http:' &&
        a.protocol != 'https:' ||
        !isTmeHostname(a.hostname) ||
        !(match = a.pathname.match(/^\/(\$|invoice\/)([A-Za-z0-9\-_=]+)$/)) ||
        !(slug = match[2])) {
      console.error('[Telegram.WebApp] Invoice url is invalid', url);
      throw Error('WebAppInvoiceUrlInvalid');
    }
    if (!versionAtLeast('6.1')) {
      console.error('[Telegram.WebApp] Method openInvoice is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (webAppInvoices[slug]) {
      console.error('[Telegram.WebApp] Invoice is already opened');
      throw Error('WebAppInvoiceOpened');
    }
    webAppInvoices[slug] = {
      url: url,
      callback: callback
    };
    WebView.postEvent('web_app_open_invoice', false, {slug: slug});
  };
  WebApp.showPopup = function (params, callback) {
    if (!versionAtLeast('6.2')) {
      console.error('[Telegram.WebApp] Method showPopup is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (webAppPopupOpened) {
      console.error('[Telegram.WebApp] Popup is already opened');
      throw Error('WebAppPopupOpened');
    }
    var title = '';
    var message = '';
    var buttons = [];
    var popup_buttons = {};
    var popup_params = {};
    if (typeof params.title !== 'undefined') {
      title = strTrim(params.title);
      if (title.length > 64) {
        console.error('[Telegram.WebApp] Popup title is too long', title);
        throw Error('WebAppPopupParamInvalid');
      }
      if (title.length > 0) {
        popup_params.title = title;
      }
    }
    if (typeof params.message !== 'undefined') {
      message = strTrim(params.message);
    }
    if (!message.length) {
      console.error('[Telegram.WebApp] Popup message is required', params.message);
      throw Error('WebAppPopupParamInvalid');
    }
    if (message.length > 256) {
      console.error('[Telegram.WebApp] Popup message is too long', message);
      throw Error('WebAppPopupParamInvalid');
    }
    popup_params.message = message;
    if (typeof params.buttons !== 'undefined') {
      if (!Array.isArray(params.buttons)) {
        console.error('[Telegram.WebApp] Popup buttons should be an array', params.buttons);
        throw Error('WebAppPopupParamInvalid');
      }
      for (var i = 0; i < params.buttons.length; i++) {
        var button = params.buttons[i];
        var btn = {};
        var id = '';
        if (typeof button.id !== 'undefined') {
          id = button.id.toString();
          if (id.length > 64) {
            console.error('[Telegram.WebApp] Popup button id is too long', id);
            throw Error('WebAppPopupParamInvalid');
          }
        }
        btn.id = id;
        var button_type = button.type;
        if (typeof button_type === 'undefined') {
          button_type = 'default';
        }
        btn.type = button_type;
        if (button_type == 'ok' ||
            button_type == 'close' ||
            button_type == 'cancel') {
          // no params needed
        } else if (button_type == 'default' ||
                   button_type == 'destructive') {
          var text = '';
          if (typeof button.text !== 'undefined') {
            text = strTrim(button.text);
          }
          if (!text.length) {
            console.error('[Telegram.WebApp] Popup button text is required for type ' + button_type, button.text);
            throw Error('WebAppPopupParamInvalid');
          }
          if (text.length > 64) {
            console.error('[Telegram.WebApp] Popup button text is too long', text);
            throw Error('WebAppPopupParamInvalid');
          }
          btn.text = text;
        } else {
          console.error('[Telegram.WebApp] Popup button type is invalid', button_type);
          throw Error('WebAppPopupParamInvalid');
        }
        buttons.push(btn);
      }
    } else {
      buttons.push({id: '', type: 'close'});
    }
    if (buttons.length < 1) {
      console.error('[Telegram.WebApp] Popup should have at least one button');
      throw Error('WebAppPopupParamInvalid');
    }
    if (buttons.length > 3) {
      console.error('[Telegram.WebApp] Popup should not have more than 3 buttons');
      throw Error('WebAppPopupParamInvalid');
    }
    popup_params.buttons = buttons;

    webAppPopupOpened = {
      callback: callback
    };
    WebView.postEvent('web_app_open_popup', false, popup_params);
  };
  WebApp.showAlert = function (message, callback) {
    WebApp.showPopup({
      message: message
    }, callback ? function(){ callback(); } : null);
  };
  WebApp.showConfirm = function (message, callback) {
    WebApp.showPopup({
      message: message,
      buttons: [
        {type: 'ok', id: 'ok'},
        {type: 'cancel'}
      ]
    }, callback ? function (button_id) {
      callback(button_id == 'ok');
    } : null);
  };
  WebApp.showScanQrPopup = function (params, callback) {
    if (!versionAtLeast('6.4')) {
      console.error('[Telegram.WebApp] Method showScanQrPopup is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (webAppScanQrPopupOpened) {
      console.error('[Telegram.WebApp] Popup is already opened');
      throw Error('WebAppScanQrPopupOpened');
    }
    var text = '';
    var popup_params = {};
    if (typeof params.text !== 'undefined') {
      text = strTrim(params.text);
      if (text.length > 64) {
        console.error('[Telegram.WebApp] Scan QR popup text is too long', text);
        throw Error('WebAppScanQrPopupParamInvalid');
      }
      if (text.length > 0) {
        popup_params.text = text;
      }
    }

    webAppScanQrPopupOpened = {
      callback: callback
    };
    WebView.postEvent('web_app_open_scan_qr_popup', false, popup_params);
  };
  WebApp.closeScanQrPopup = function () {
    if (!versionAtLeast('6.4')) {
      console.error('[Telegram.WebApp] Method closeScanQrPopup is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }

    webAppScanQrPopupOpened = false;
    WebView.postEvent('web_app_close_scan_qr_popup', false);
  };
  WebApp.readTextFromClipboard = function (callback) {
    if (!versionAtLeast('6.4')) {
      console.error('[Telegram.WebApp] Method readTextFromClipboard is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    var req_id = generateCallbackId(16);
    var req_params = {req_id: req_id};
    webAppCallbacks[req_id] = {
      callback: callback
    };
    WebView.postEvent('web_app_read_text_from_clipboard', false, req_params);
  };
  WebApp.requestWriteAccess = function (callback) {
    if (!versionAtLeast('6.9')) {
      console.error('[Telegram.WebApp] Method requestWriteAccess is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (WebAppWriteAccessRequested) {
      console.error('[Telegram.WebApp] Write access is already requested');
      throw Error('WebAppWriteAccessRequested');
    }
    WebAppWriteAccessRequested = {
      callback: callback
    };
    WebView.postEvent('web_app_request_write_access');
  };
  WebApp.requestContact = function (callback) {
    if (!versionAtLeast('6.9')) {
      console.error('[Telegram.WebApp] Method requestContact is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (WebAppContactRequested) {
      console.error('[Telegram.WebApp] Contact is already requested');
      throw Error('WebAppContactRequested');
    }
    WebAppContactRequested = {
      callback: callback
    };
    WebView.postEvent('web_app_request_phone');
  };
  WebApp.downloadFile = function (params, callback) {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method downloadFile is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (webAppDownloadFileRequested) {
      console.error('[Telegram.WebApp] Popup is already opened');
      throw Error('WebAppDownloadFilePopupOpened');
    }
    var a = document.createElement('A');

    var dl_params = {};
    if (!params || !params.url || !params.url.length) {
      console.error('[Telegram.WebApp] Url is required');
      throw Error('WebAppDownloadFileParamInvalid');
    }
    a.href = params.url;
    if (a.protocol != 'https:') {
      console.error('[Telegram.WebApp] Url protocol is not supported', url);
      throw Error('WebAppDownloadFileParamInvalid');
    }
    dl_params.url = a.href;

    if (!params || !params.file_name || !params.file_name.length) {
      console.error('[Telegram.WebApp] File name is required');
      throw Error('WebAppDownloadFileParamInvalid');
    }
    dl_params.file_name = params.file_name;

    webAppDownloadFileRequested = {
      callback: callback
    };
    WebView.postEvent('web_app_request_file_download', false, dl_params);
  };
  WebApp.shareToStory = function (media_url, params) {
    params = params || {};
    if (!versionAtLeast('7.8')) {
      console.error('[Telegram.WebApp] Method shareToStory is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    var a = document.createElement('A');
    a.href = media_url;
    if (a.protocol != 'http:' &&
        a.protocol != 'https:') {
      console.error('[Telegram.WebApp] Media url protocol is not supported', url);
      throw Error('WebAppMediaUrlInvalid');
    }
    var share_params = {};
    share_params.media_url = a.href;
    if (typeof params.text !== 'undefined') {
      var text = strTrim(params.text);
      if (text.length > 2048) {
        console.error('[Telegram.WebApp] Text is too long', text);
        throw Error('WebAppShareToStoryParamInvalid');
      }
      if (text.length > 0) {
        share_params.text = text;
      }
    }
    if (typeof params.widget_link !== 'undefined') {
      params.widget_link = params.widget_link || {};
      a.href = params.widget_link.url;
      if (a.protocol != 'http:' &&
          a.protocol != 'https:') {
        console.error('[Telegram.WebApp] Link protocol is not supported', url);
        throw Error('WebAppShareToStoryParamInvalid');
      }
      var widget_link = {
        url: a.href
      };
      if (typeof params.widget_link.name !== 'undefined') {
        var link_name = strTrim(params.widget_link.name);
        if (link_name.length > 48) {
          console.error('[Telegram.WebApp] Link name is too long', link_name);
          throw Error('WebAppShareToStoryParamInvalid');
        }
        if (link_name.length > 0) {
          widget_link.name = link_name;
        }
      }
      share_params.widget_link = widget_link;
    }

    WebView.postEvent('web_app_share_to_story', false, share_params);
  };
  WebApp.shareMessage = function (msg_id, callback) {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method shareMessage is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (WebAppShareMessageOpened) {
      console.error('[Telegram.WebApp] Share message is already opened');
      throw Error('WebAppShareMessageOpened');
    }
    WebAppShareMessageOpened = {
      callback: callback
    };
    WebView.postEvent('web_app_send_prepared_message', false, {id: msg_id});
  };
  WebApp.requestChat = function (req_id, callback) {
    if (!versionAtLeast('9.6')) {
      console.error('[Telegram.WebApp] Method requestChat is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (WebAppRequestChatOpened) {
      console.error('[Telegram.WebApp] Request chat is already opened');
      throw Error('WebAppRequestChatOpened');
    }
    WebAppRequestChatOpened = {
      callback: callback
    };
    WebView.postEvent('web_app_request_chat', false, {req_id: req_id});
  };
  WebApp.setEmojiStatus = function (custom_emoji_id, params, callback) {
    params = params || {};
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method setEmojiStatus is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    var status_params = {};
    status_params.custom_emoji_id = custom_emoji_id;
    if (typeof params.duration !== 'undefined') {
      status_params.duration = params.duration;
    }
    if (WebAppEmojiStatusRequested) {
      console.error('[Telegram.WebApp] Emoji status is already requested');
      throw Error('WebAppEmojiStatusRequested');
    }
    WebAppEmojiStatusRequested = {
      callback: callback
    };
    WebView.postEvent('web_app_set_emoji_status', false, status_params);
  };
  WebApp.requestEmojiStatusAccess = function (callback) {
    if (!versionAtLeast('8.0')) {
      console.error('[Telegram.WebApp] Method requestEmojiStatusAccess is not supported in version ' + webAppVersion);
      throw Error('WebAppMethodUnsupported');
    }
    if (WebAppEmojiStatusAccessRequested) {
      console.error('[Telegram.WebApp] Emoji status permission is already requested');
      throw Error('WebAppEmojiStatusAccessRequested');
    }
    WebAppEmojiStatusAccessRequested = {
      callback: callback
    };
    WebView.postEvent('web_app_request_emoji_status_access');
  };
  WebApp.invokeCustomMethod = function (method, params, callback) {
    invokeCustomMethod(method, params, callback);
  };
  WebApp.hideKeyboard = function () {
    WebView.postEvent('web_app_hide_keyboard');
  };
  WebApp.ready = function () {
    WebView.postEvent('web_app_ready');
  };
  WebApp.expand = function () {
    WebView.postEvent('web_app_expand');
  };
  WebApp.close = function (options) {
    options = options || {};
    var req_params = {};
    if (versionAtLeast('7.6') && options.return_back) {
      req_params.return_back = true;
    }
    WebView.postEvent('web_app_close', false, req_params);
  };

  window.Telegram.WebApp = WebApp;

  updateHeaderColor();
  updateBackgroundColor();
  updateBottomBarColor();
  setViewportHeight();
  if (initParams.tgWebAppShowSettings) {
    SettingsButton.show();
  }

  window.addEventListener('resize', onWindowResize);
  if (isIframe) {
    document.addEventListener('click', linkHandler);
  }

  WebView.onEvent('theme_changed', onThemeChanged);
  WebView.onEvent('viewport_changed', onViewportChanged);
  WebView.onEvent('safe_area_changed', onSafeAreaChanged);
  WebView.onEvent('content_safe_area_changed', onContentSafeAreaChanged);
  WebView.onEvent('visibility_changed', onVisibilityChanged);
  WebView.onEvent('invoice_closed', onInvoiceClosed);
  WebView.onEvent('popup_closed', onPopupClosed);
  WebView.onEvent('qr_text_received', onQrTextReceived);
  WebView.onEvent('scan_qr_popup_closed', onScanQrPopupClosed);
  WebView.onEvent('clipboard_text_received', onClipboardTextReceived);
  WebView.onEvent('write_access_requested', onWriteAccessRequested);
  WebView.onEvent('phone_requested', onPhoneRequested);
  WebView.onEvent('file_download_requested', onFileDownloadRequested);
  WebView.onEvent('custom_method_invoked', onCustomMethodInvoked);
  WebView.onEvent('fullscreen_changed', onFullscreenChanged);
  WebView.onEvent('fullscreen_failed', onFullscreenFailed);
  WebView.onEvent('home_screen_added', onHomeScreenAdded);
  WebView.onEvent('home_screen_checked', onHomeScreenChecked);
  WebView.onEvent('prepared_message_sent', onPreparedMessageSent);
  WebView.onEvent('prepared_message_failed', onPreparedMessageFailed);
  WebView.onEvent('requested_chat_sent', onRequestedChatSent);
  WebView.onEvent('requested_chat_failed', onRequestedChatFailed);
  WebView.onEvent('emoji_status_set', onEmojiStatusSet);
  WebView.onEvent('emoji_status_failed', onEmojiStatusFailed);
  WebView.onEvent('emoji_status_access_requested', onEmojiStatusAccessRequested);
  WebView.postEvent('web_app_request_theme');
  WebView.postEvent('web_app_request_viewport');
  WebView.postEvent('web_app_request_safe_area');
  WebView.postEvent('web_app_request_content_safe_area');

})(); -c "from brillo2001 import __version__; print(__version__)" && brillo --help && python -m pytest tests/ -q 2>&1


# Asistente local de archivos para iPhone 13 Pro

Esta guía resume el sistema mostrado en las imágenes: un asistente 100% local para organizar, clasificar, buscar, respaldar y reportar archivos desde la app **Archivos** y **Atajos** de iOS, sin internet, sin nube y sin servicios externos.

## Objetivo

El sistema permite administrar documentos directamente en el iPhone mediante automatizaciones locales. El usuario conserva el control de sus archivos, puede consultar inventarios, detectar duplicados, generar respaldos comprimidos y producir reportes de estado.

## Estructura de carpetas

La carpeta raíz sugerida es `Asistente_Local_Luis` e incluye las siguientes subcarpetas:

| Carpeta | Uso |
| --- | --- |
| `01_Documentos` | PDF, Word, Excel, certificados y documentos generales. |
| `02_Imagenes` | Fotos personales, escaneos e imágenes de trabajo. |
| `03_Videos` | Videos personales o laborales. |
| `04_Audio` | Grabaciones, música y notas de voz. |
| `05_Comprimidos` | Archivos ZIP, RAR, 7z y otros paquetes. |
| `06_Respaldos` | Respaldos automáticos comprimidos. |
| `07_Reportes` | Reportes generados por el sistema. |
| `08_Pendientes` | Archivos pendientes de organizar o revisar. |
| `09_BaseDatos` | Base de datos e inventarios del sistema. |
| `10_Plantillas` | Plantillas reutilizables para reportes. |

## Base de datos local

El inventario principal se guarda como CSV en `Asistente_Local_Luis/09_BaseDatos/inventario.csv`. Cada fila representa un archivo registrado por el sistema.

Columnas recomendadas:

- `Fecha`: fecha de registro o última actualización.
- `Nombre`: nombre del archivo.
- `Tipo`: tipo o extensión normalizada.
- `Tamaño`: tamaño legible del archivo.
- `Ubicación`: ruta local dentro de la estructura del asistente.

Ejemplo:

```csv
Fecha,Nombre,Tipo,Tamaño,Ubicación
2026-08-04,Certificado.pdf,PDF,2.1 MB,/01_Documentos/Certificados
2026-08-04,Contrato.docx,Word,340 KB,/01_Documentos/Word
2026-08-04,Presupuesto.xlsx,Excel,120 KB,/01_Documentos/Excel
```

## Atajos principales

### 1. Organizar Archivos

Escanea una carpeta, obtiene el contenido de cada archivo, detecta su extensión y lo mueve automáticamente a la carpeta correspondiente.

Acciones principales:

1. Obtener contenido de carpeta.
2. Repetir con cada archivo.
3. Obtener nombre y extensión.
4. Evaluar el tipo de archivo.
5. Mover el archivo a la carpeta asignada.

### 2. Actualizar Inventario

Recorre todas las carpetas del sistema y registra cada archivo en `inventario.csv` con su información básica.

Acciones principales:

1. Obtener todas las carpetas.
2. Obtener contenido.
3. Obtener nombre, tipo, tamaño y fecha.
4. Agregar una fila al CSV.

### 3. Buscar Duplicados

Compara nombres y tamaños para identificar archivos repetidos y guarda un reporte en `07_Reportes`.

Acciones principales:

1. Obtener todos los archivos.
2. Comparar nombres.
3. Comparar tamaños.
4. Registrar duplicados.
5. Generar `duplicados.txt`.

### 4. Respaldo Diario

Crea un respaldo completo de carpetas importantes y lo comprime en un ZIP dentro de `06_Respaldos`.

Acciones principales:

1. Seleccionar carpetas.
2. Copiar archivos.
3. Crear archivo ZIP.
4. Guardar el ZIP en `06_Respaldos`.

### 5. Buscar Archivo

Permite buscar documentos por nombre en las carpetas importantes y muestra ubicación, tamaño y tipo.

Acciones principales:

1. Pedir nombre a buscar.
2. Buscar en carpetas configuradas.
3. Mostrar resultados y detalles.

### 6. Estado del Sistema

Genera un reporte general con estadísticas del sistema y lo guarda como RTF en `07_Reportes`.

Acciones principales:

1. Contar archivos por tipo.
2. Calcular tamaños.
3. Crear documento RTF.
4. Guardar `Estado.rtf`.

## Tipos de archivos soportados

- Documentos: PDF, DOCX, XLSX, TXT, RTF.
- Imágenes: JPG, PNG.
- Videos: MP4, MOV.
- Audio: MP3, M4A.
- Comprimidos: ZIP, RAR, 7z.
- Datos: CSV y JSON.

## Flujo general del sistema

1. Escanear archivos desde la app Archivos.
2. Clasificar automáticamente por tipo.
3. Organizar en carpetas.
4. Registrar cada archivo en la base de datos CSV.
5. Generar reportes y estadísticas.
6. Crear respaldos comprimidos.
7. Consultar o buscar documentos cuando sea necesario.

## Mejora de automatizaciones y redireccionamiento automático

El asistente puede reforzarse con reglas de automatización que clasifiquen, redirijan y validen archivos sin intervención manual. La idea es que cada archivo nuevo pase por un flujo estándar: detección, clasificación, redireccionamiento, registro y verificación.

### Reglas de redireccionamiento automático

Configurar el atajo **Organizar Archivos** para evaluar cada archivo por extensión, nombre y ubicación de origen. Según el resultado, el archivo debe moverse automáticamente a la carpeta correspondiente.

| Condición detectada | Destino automático | Acción adicional |
| --- | --- | --- |
| `.pdf`, `.docx`, `.xlsx`, `.txt`, `.rtf` | `01_Documentos` | Registrar tipo documental en el inventario. |
| `.jpg`, `.jpeg`, `.png`, `.heic` | `02_Imagenes` | Conservar fecha de creación si está disponible. |
| `.mp4`, `.mov` | `03_Videos` | Registrar tamaño para reportes de almacenamiento. |
| `.mp3`, `.m4a`, `.wav` | `04_Audio` | Clasificar como audio o nota de voz. |
| `.zip`, `.rar`, `.7z` | `05_Comprimidos` | Marcar como paquete comprimido. |
| `.csv`, `.json` | `09_BaseDatos` | Evitar sobrescribir el inventario principal. |
| Tipo desconocido | `08_Pendientes` | Marcar para revisión manual. |

### Flujo recomendado de automatización

1. Detectar archivos nuevos o modificados en la carpeta de entrada.
2. Obtener nombre, extensión, tamaño, fecha y ruta original.
3. Normalizar la extensión a minúsculas para evitar duplicados de reglas.
4. Aplicar la tabla de redireccionamiento automático.
5. Mover el archivo a la carpeta destino.
6. Agregar o actualizar la entrada en `09_BaseDatos/inventario.csv`.
7. Registrar archivos no reconocidos en un reporte de pendientes.
8. Mostrar un resumen con archivos procesados, movidos, omitidos y pendientes.

### Validaciones antes de mover archivos

Antes de redirigir un archivo, el atajo debe comprobar:

- Que el archivo todavía exista en la ubicación original.
- Que la carpeta destino exista; si no existe, crearla automáticamente.
- Que no haya otro archivo con el mismo nombre en el destino.
- Que el inventario no quede duplicado.
- Que los archivos críticos, como `inventario.csv`, no se sobrescriban accidentalmente.

Cuando exista un conflicto de nombre, usar una estrategia consistente, por ejemplo agregar fecha y hora al nombre del archivo:

```text
Contrato.pdf
Contrato_2026-08-14_1530.pdf
```

### Automatizaciones programadas sugeridas

| Automatización | Frecuencia recomendada | Resultado esperado |
| --- | --- | --- |
| Organizar archivos | Al guardar o importar archivos | Archivos redirigidos a su carpeta correcta. |
| Actualizar inventario | Diario | CSV actualizado con rutas y metadatos. |
| Buscar duplicados | Semanal | Reporte de posibles duplicados. |
| Respaldo comprimido | Diario o semanal | ZIP guardado en `06_Respaldos`. |
| Estado del sistema | Semanal | Reporte RTF con estadísticas generales. |

### Criterio de pruebas y comprobaciones

Ejecutar pruebas automatizadas y validaciones de código en cada cambio que afecte la lógica o el comportamiento de la aplicación. No es necesario ejecutar estas comprobaciones cuando los cambios se limiten exclusivamente a comentarios, documentación o contenido no funcional.

Ejecutar pruebas y comprobaciones de código para cualquier cambio funcional. Omitirlas cuando los cambios afecten únicamente a comentarios o documentación.

## Seguridad y privacidad

El diseño funciona completamente en local:

- No requiere internet.
- No utiliza nube.
- No comparte información con servicios externos.
- Puede protegerse con Face ID o código.
- Mantiene los respaldos dentro del dispositivo, bajo control del usuario.
/**
 * ============================================================================
 * CONFIGURACIÓN CENTRALIZADA DE BOTS + TOKENS + ADMIN IDs
 * ============================================================================
 * IMPORTANTE:
 * - Reemplaza los valores de "BOT_TOKEN" y "ADMIN_IDS" por los reales
 *   antes de desplegar, O mejor aún: ponlos como variables de entorno
 *   en Cloudflare (recomendado).
 * - ADMIN_IDS debe ser un array de números (Telegram user IDs).
 * ============================================================================
 */

export const BOTS = {
  asistentelg: {
    id: 1,
    name: "Asistente LG",
    username: "Asistente_LG_bot",
    shortName: "direclinkstartappreportes",
    directLink: "https://t.me/Asistente_LG_bot/direclinkstartappreportes",
    description: "Tu asistente inteligente para consultar, organizar y automatizar tareas desde Telegram.",
    webAppUrl: "https://asistentelg.midominio.workers.dev/",
    menuButtonText: "🚀 Abrir Asistente LG",
    // ─── Variables de entorno (recomendado) ───────────────────────────────
    // En wrangler.toml o Dashboard de Cloudflare:
    // BOT_TOKEN = "123456:ABC-DEF..."
    // ADMIN_IDS = "123456789,987654321"   (separados por coma)
    envKeys: {
      BOT_TOKEN: "BOT_TOKEN",
      ADMIN_IDS: "ADMIN_IDS",
    },
    // Valores de respaldo (solo para desarrollo local – NUNCA subir tokens reales)
    fallback: {
      BOT_TOKEN: "REEMPLAZA_CON_TU_TOKEN_ASISTENTE_LG",
      ADMIN_IDS: [123456789], // ← cambia por tus IDs reales
    },
  },

  medios: {
    id: 2,
    name: "Medios Bot",
    username: "medios_bot",
    shortName: "direclink",
    directLink: "https://t.me/medios_bot/direclink",
    description: "Presiona /start para iniciar",
    webAppUrl: "https://lgms21bot-worker.midominio.workers.dev/",
    menuButtonText: "📱 Abrir App Medios",
    envKeys: {
      BOT_TOKEN: "BOT_TOKEN",
      ADMIN_IDS: "ADMIN_IDS",
    },
    fallback: {
      BOT_TOKEN: "REEMPLAZA_CON_TU_TOKEN_MEDIOS",
      ADMIN_IDS: [123456789],
    },
  },

  terminobot: {
    id: 3,
    name: "TerminoBot / TetrisBot",
    username: "terminobotortetrisbot",
    shortName: "direclink",
    directLink: "https://t.me/terminobotortetrisbot/direclink",
    description: "Para iniciar la App, presiona /start",
    webAppUrl: "https://asistente-archivos.midominio.workers.dev/",
    menuButtonText: "📂 Abrir App Archivos",
    envKeys: {
      BOT_TOKEN: "BOT_TOKEN",
      ADMIN_IDS: "ADMIN_IDS",
    },
    fallback: {
      BOT_TOKEN: "REEMPLAZA_CON_TU_TOKEN_TERMINOBOT",
      ADMIN_IDS: [123456789],
    },
  },

  agente23: {
    id: 4,
    name: "Agente 23",
    username: "gente23_bot",
    shortName: "direclink",
    directLink: "https://t.me/gente23_bot/direclink",
    description: "Para iniciar, presiona /start",
    webAppUrl: "https://hello-world-do-template.midominio.workers.dev/",
    menuButtonText: "🤖 Abrir Agente 23",
    envKeys: {
      BOT_TOKEN: "BOT_TOKEN",
      ADMIN_IDS: "ADMIN_IDS",
    },
    fallback: {
      BOT_TOKEN: "REEMPLAZA_CON_TU_TOKEN_AGENTE23",
      ADMIN_IDS: [123456789],
    },
  },

  agente21: {
    id: 5,
    name: "Agente 21",
    username: "Agente21_bot",
    shortName: "direclink",
    directLink: "https://t.me/Agente21_bot/direclink",
    description: "Presiona /start para iniciar",
    webAppUrl: "https://lgms.midominio.workers.dev/",
    menuButtonText: "🤖 Abrir Agente 21",
    envKeys: {
      BOT_TOKEN: "BOT_TOKEN",
      ADMIN_IDS: "ADMIN_IDS",
    },
    fallback: {
      BOT_TOKEN: "REEMPLAZA_CON_TU_TOKEN_AGENTE21",
      ADMIN_IDS: [123456789],
    },
  },

  lgms21: {
    id: 6,
    name: "LGMS 21",
    username: "lgms21_bot",
    shortName: "direclink",
    directLink: "https://t.me/lgms21_bot/direclink",
    description: "Para iniciar solo presiona /start",
    webAppUrl: "https://kv-get-started.midominio.workers.dev/",
    menuButtonText: "⚙️ Abrir LGMS 21",
    envKeys: {
      BOT_TOKEN: "BOT_TOKEN",
      ADMIN_IDS: "ADMIN_IDS",
    },
    fallback: {
      BOT_TOKEN: "REEMPLAZA_CON_TU_TOKEN_LGMS21",
      ADMIN_IDS: [123456789],
    },
  },
};

/**
 * Obtiene la configuración de un bot + tokens resueltos desde env
 */
export function resolveBotConfig(botKey, env = {}) {
  const base = BOTS[botKey];
  if (!base) throw new Error(`Bot key desconocido: ${botKey}`);

  const token =
    env[base.envKeys.BOT_TOKEN] ||
    env.BOT_TOKEN ||
    base.fallback.BOT_TOKEN;

  const adminIdsRaw =
    env[base.envKeys.ADMIN_IDS] ||
    env.ADMIN_IDS ||
    base.fallback.ADMIN_IDS.join(",");

  const adminIds = String(adminIdsRaw)
    .split(",")
    .map((id) => parseInt(id.trim(), 10))
    .filter((id) => !isNaN(id));

  return {
    ...base,
    token,
    adminIds,
    isAdmin: (userId) => adminIds.includes(Number(userId)),
  };
}
/**
 * ============================================================================
 * CONFIGURACIÓN CENTRALIZADA DE BOTS DE TELEGRAM + WEB APPS
 * Listo para implementar en Cloudflare Workers / Frontend / Backend
 * ============================================================================
 * Uso:
 *   import { BOTS, getBotByUsername, getBotByShortName } from './bots-config.js';
 *   const bot = getBotByUsername('@medios_bot');
 *   console.log(bot.webAppUrl);
 * ============================================================================
 */

export const BOTS = [
  {
    id: 1,
    name: "Asistente LG",
    username: "Asistente_LG_bot",          // sin @
    usernameWithAt: "@Asistente_LG_bot",
    shortName: "direclinkstartappreportes",
    directLink: "https://t.me/Asistente_LG_bot/direclinkstartappreportes",
    description: "Tu asistente inteligente para consultar, organizar y automatizar tareas desde Telegram.",
    webAppUrl: "https://asistentelg.midominio.workers.dev/",
    menuButtonText: "🚀 Abrir Asistente LG",
    suggestedCommands: [
      { command: "start", description: "Iniciar el asistente" },
      { command: "app", description: "Abrir la aplicación web" },
      { command: "help", description: "Ayuda y comandos" }
    ],
    notes: "Descripción original correcta. Short name largo pero válido."
  },
  {
    id: 2,
    name: "Medios Bot",
    username: "medios_bot",
    usernameWithAt: "@medios_bot",
    shortName: "direclink",
    directLink: "https://t.me/medios_bot/direclink",
    description: "Presiona /start para iniciar",
    webAppUrl: "https://lgms21bot-worker.midominio.workers.dev/",
    menuButtonText: "📱 Abrir App Medios",
    suggestedCommands: [
      { command: "start", description: "Iniciar el bot" },
      { command: "app", description: "Abrir la Web App" },
      { command: "help", description: "Ayuda" }
    ],
    notes: "Corregido 'Peeciona' → 'Presiona'. URL apunta a worker de lgms21."
  },
  {
    id: 3,
    name: "TerminoBot / TetrisBot",
    username: "terminobotortetrisbot",
    usernameWithAt: "@terminobotortetrisbot",
    shortName: "direclink",
    directLink: "https://t.me/terminobotortetrisbot/direclink",
    description: "Para iniciar la App, presiona /start",
    webAppUrl: "https://asistente-archivos.midominio.workers.dev/",
    menuButtonText: "📂 Abrir App Archivos",
    suggestedCommands: [
      { command: "start", description: "Iniciar la aplicación" },
      { command: "app", description: "Abrir Web App de archivos" },
      { command: "help", description: "Ayuda" }
    ],
    notes: "Username parece concatenado. Verificar si existen @terminobot y @tetris_bot por separado."
  },
  {
    id: 4,
    name: "Agente 23",
    username: "gente23_bot",                 // según Direct Link real
    usernameWithAt: "@gente23_bot",
    shortName: "direclink",
    directLink: "https://t.me/gente23_bot/direclink",
    description: "Para iniciar, presiona /start",
    webAppUrl: "https://hello-world-do-template.midominio.workers.dev/",
    menuButtonText: "🤖 Abrir Agente 23",
    suggestedCommands: [
      { command: "start", description: "Iniciar Agente 23" },
      { command: "app", description: "Abrir Web App" },
      { command: "help", description: "Ayuda" }
    ],
    notes: "Usuario escribió @agente23_bot pero el link usa @gente23_bot. Se usó el del Direct Link."
  },
  {
    id: 5,
    name: "Agente 21",
    username: "Agente21_bot",
    usernameWithAt: "@Agente21_bot",
    shortName: "direclink",
    directLink: "https://t.me/Agente21_bot/direclink",
    description: "Presiona /start para iniciar",
    webAppUrl: "https://lgms.midominio.workers.dev/",
    menuButtonText: "🤖 Abrir Agente 21",
    suggestedCommands: [
      { command: "start", description: "Iniciar Agente 21" },
      { command: "app", description: "Abrir Web App" },
      { command: "help", description: "Ayuda" }
    ],
    notes: "Verificar mayúsculas exactas del username en BotFather."
  },
  {
    id: 6,
    name: "LGMS 21",
    username: "lgms21_bot",
    usernameWithAt: "@lgms21_bot",
    shortName: "direclink",
    directLink: "https://t.me/lgms21_bot/direclink",
    description: "Para iniciar solo presiona /start",
    webAppUrl: "https://kv-get-started.midominio.workers.dev/",
    menuButtonText: "⚙️ Abrir LGMS 21",
    suggestedCommands: [
      { command: "start", description: "Iniciar LGMS 21" },
      { command: "app", description: "Abrir Web App" },
      { command: "help", description: "Ayuda" }
    ],
    notes: "URL de ejemplo Cloudflare KV Get Started."
  }
];

/**
 * Helpers
 */
export function getBotByUsername(username) {
  const clean = username.replace(/^@/, "").toLowerCase();
  return BOTS.find(b => b.username.toLowerCase() === clean) || null;
}

export function getBotByShortName(shortName) {
  return BOTS.find(b => b.shortName === shortName) || null;
}

export function getBotById(id) {
  return BOTS.find(b => b.id === id) || null;
}

export function getAllWebAppUrls() {
  return BOTS.map(b => ({ name: b.name, url: b.webAppUrl }));
}

export function getDirectLinks() {
  return BOTS.map(b => ({ name: b.name, link: b.directLink }));
}

/**
 * Genera el texto para /setcommands de BotFather
 */
export function generateSetCommands(botId) {
  const bot = getBotById(botId);
  if (!bot) return null;
  return bot.suggestedCommands
    .map(c => `${c.command} - ${c.description}`)
    .join("\n");
}

/**
 * Ejemplo de uso en un Cloudflare Worker (handler)
 * 
 * export default {
 *   async fetch(request, env, ctx) {
 *     const url = new URL(request.url);
 *     // Puedes detectar de qué bot viene por el origen o por un parámetro
 *     const botConfig = getBotByUsername(url.searchParams.get("bot") || "Asistente_LG_bot");
 *     
 *     // Lógica de la Web App...
 *     return new Response(JSON.stringify({ bot: botConfig?.name }), {
 *       headers: { "Content-Type": "application/json" }
 *     });
 *   }
 * }
 */

/**
 * Ejemplo para Telegram WebApp (frontend)
 * 
 * // En tu HTML/JS de la Web App:
 * const tg = window.Telegram.WebApp;
 * tg.ready();
 * tg.expand();
 * 
 * // Puedes recibir el bot actual vía query param o hardcodear según el worker
 * const currentBot = getBotByUsername("medios_bot");
 * document.title = currentBot?.name || "Telegram Web App";
 */

export default BOTS;

export default BOTS;

