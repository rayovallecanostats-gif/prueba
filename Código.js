// PARTE 1: El puente de Google que alimentará tu web de GitHub
function doGet(e) {
  var ruta = (e && e.parameter && e.parameter.v) ? e.parameter.v : 'index';
  var termino = (e && e.parameter && e.parameter.q) ? e.parameter.q : '';
  
  // Si la petición pide "buscar", ejecuta la lógica y devuelve JSON (lo que GitHub entiende)
  if (ruta === 'buscar') {
    var resultados = buscarPublicacion(termino);
    return ContentService.createTextOutput(JSON.stringify(resultados))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Por si acaso entran directamente al enlace de Google, mantiene el comportamiento original
  var archivoHtml = 'Index';
  if (ruta === 'alineaciones') archivoHtml = 'Lineups';
  if (ruta === 'widget') archivoHtml = 'Widget';
  
  return HtmlService.createHtmlOutputFromFile(archivoHtml)
    .setTitle('Rayostats')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function procesarEstructuraTexto(texto) {
  if (!texto) return { titular: "", subtitulo: "", cuerpo: "" };
  var textoLimpio = texto.toString().trim();
  var lineas = textoLimpio.split("\n").map(function(l) { return l.trim(); }).filter(function(l) { return l !== ""; });
  var titular = "";
  var subtitulo = "";
  if (lineas.length > 1) {
    titular = lineas[0];
    subtitulo = lineas[1];
  } else {
    var frases = textoLimpio.split(".").map(function(f) { return f.trim(); }).filter(function(f) { return f !== ""; });
    titular = frases[0] ? frases[0] + "." : textoLimpio;
    subtitulo = frases[1] ? frases[1] + "." : titular;
  }
  return { titular: titular, subtitulo: subtitulo, cuerpo: textoLimpio };
}

function optimizarUrlImagen(urlOriginal, esPortada) {
  if (!urlOriginal) return "NO_IMAGE";
  var urlStr = urlOriginal.toString().trim();
  if (!urlStr.startsWith("http") || urlStr.indexOf("undefined") !== -1) return "NO_IMAGE";
  var ancho = esPortada ? 1400 : 1080;
  return 'https://wsrv.nl' + encodeURIComponent(urlStr) + '&w=' + ancho + '&q=100';
}

function buscarPublicacion(termino) {
  try {
    termino = termino ? termino.toString().trim() : "";
    var cacheKey = "busqueda_v11_" + encodeURIComponent(termino.toLowerCase());
    try {
      var cache = CacheService.getScriptCache();
      var cachedData = cache.get(cacheKey);
      if (cachedData) return JSON.parse(cachedData);
    } catch (eCacheRead) {}

    var idHoja = "1t7oPWd7lPTBnnW41RFeOtIRJiDcpacgLKY2e55xNJsg"; 
    var ss = SpreadsheetApp.openById(idHoja);
    var sheet = ss.getSheetByName("Articulos") || ss.getSheets()[0];
    var datos = sheet.getDataRange().getDisplayValues();
    if (!datos || datos.length < 2) return [];
    
    var cabecera = datos[0];
    var colTexto = 1;        
    var colFotoPrincipal = 7; 
    var indicesFotosCarrusel = [];
    var colUrlSA = -1;
    var colFechaRY = -1;

    for (var c = 0; c < cabecera.length; c++) {
      var nombreCol = cabecera[c] ? cabecera[c].toString().toLowerCase().trim() : "";
      if (nombreCol.indexOf("childposts/") !== -1 && nombreCol.indexOf("displayurl") !== -1) indicesFotosCarrusel.push(c);
      if (nombreCol === 'sa') colUrlSA = c;
      if (nombreCol === 'timestamp' || nombreCol === 'ry') colFechaRY = c;
    }

    var listaPostsProcesados = [];
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (!fila[colTexto] || fila[colTexto].toString().trim() === "") continue;
      
      var textoCompleto = fila[colTexto].toString();
      var estructura = procesarEstructuraTexto(textoCompleto);
      var textoFechaISO = (colFechaRY !== -1 && fila[colFechaRY]) ? fila[colFechaRY].toString().trim() : "";
      var tiempoMilisegundos = 0;
      if (textoFechaISO !== "") {
        var fechaObjeto = new Date(textoFechaISO);
        tiempoMilisegundos = !isNaN(fechaObjeto.getTime()) ? fechaObjeto.getTime() : 0;
      }
      if (tiempoMilisegundos === 0) tiempoMilisegundos = i; 
      
      var urlPost = (colUrlSA !== -1 && fila[colUrlSA]) ? fila[colUrlSA].toString().trim() : "#";
      var fotoMain = fila[colFotoPrincipal] ? fila[colFotoPrincipal].toString().trim() : "";

      listaPostsProcesados.push({
        titular: estructura.titular,
        subtitulo: estructura.subtitulo,
        texto: estructura.cuerpo,
        url: urlPost,
        fotoRaw: fotoMain,
        filaRaw: fila,
        tiempo: tiempoMilisegundos
      });
    }

    listaPostsProcesados.sort(function(a, b) { return b.tiempo - a.tiempo; });

    if (termino !== "") {
      var terminoBuscado = termino.toLowerCase();
      listaPostsProcesados = listaPostsProcesados.filter(function(post) {
        return post.texto.toLowerCase().indexOf(terminoBuscado) !== -1;
      });
    }

    var resultados = [];
    var limiteResultados = Math.min(listaPostsProcesados.length, 40);

    for (var idx = 0; idx < limiteResultados; idx++) {
      var item = listaPostsProcesados[idx];
      var esTop = (idx < 3);
      var galeriaFotos = [];
      var fotoMainOpt = optimizarUrlImagen(item.fotoRaw, esTop);
      if (fotoMainOpt !== "NO_IMAGE") galeriaFotos.push(fotoMainOpt);

      for (var k = 0; k < indicesFotosCarrusel.length; k++) {
        var urlChild = item.filaRaw[indicesFotosCarrusel[k]] ? item.filaRaw[indicesFotosCarrusel[k]].toString().trim() : "";
        var urlChildOpt = optimizarUrlImagen(urlChild, esTop);
        if (urlChildOpt !== "NO_IMAGE" && galeriaFotos.indexOf(urlChildOpt) === -1) galeriaFotos.push(urlChildOpt);
      }

      resultados.push({
        titular: item.titular,
        subtitulo: item.subtitulo,
        texto: item.texto,
        url: item.url,
        foto: galeriaFotos.length > 0 ? galeriaFotos[0] : "NO_IMAGE",
        fotos: galeriaFotos,
        tiempo: item.tiempo
      });
    }

    try {
      var cacheWrite = CacheService.getScriptCache();
      var jsonPayload = JSON.stringify(resultados);
      if (jsonPayload.length < 80000) cacheWrite.put(cacheKey, jsonPayload, 600);
    } catch (eCacheWrite) {}

    return resultados;
  } catch (error) {
    return [];
  }
}
