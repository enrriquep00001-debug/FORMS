/* =====================================================
   api.gs — Registro de Beneficiarios
   Blindado para 30+ usuarios simultáneos
   ===================================================== */

var CONFIG = {
  SHEET_ID: "179i-d0ByA-rPBUn1bzFWpBCERpWqIZq6js2UsTlHqkI",
  SHEETS:   { BENEFICIARIOS: "BENEFICIARIOS" },
  TZ:       Session.getScriptTimeZone()
};

var USUARIOS = {
  "admin":        { nombre: "Administrador",    password: "vivienda2026", rol: "admin",      activo: true,  puedeEditar: true,  puedeEliminar: true  },
  "capturista1":  { nombre: "Capturista 01",    password: "cap2026#01",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista2":  { nombre: "Capturista 02",    password: "cap2026#02",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista3":  { nombre: "Capturista 03",    password: "cap2026#03",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista4":  { nombre: "Capturista 04",    password: "cap2026#04",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista5":  { nombre: "Capturista 05",    password: "cap2026#05",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista6":  { nombre: "Capturista 06",    password: "cap2026#06",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista7":  { nombre: "Capturista 07",    password: "cap2026#07",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista8":  { nombre: "Capturista 08",    password: "cap2026#08",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista9":  { nombre: "Capturista 09",    password: "cap2026#09",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista10": { nombre: "Capturista 10",    password: "cap2026#10",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista11": { nombre: "Capturista 11",    password: "cap2026#11",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista12": { nombre: "Capturista 12",    password: "cap2026#12",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista13": { nombre: "Capturista 13",    password: "cap2026#13",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista14": { nombre: "Capturista 14",    password: "cap2026#14",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista15": { nombre: "Capturista 15",    password: "cap2026#15",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista16": { nombre: "Capturista 16",    password: "cap2026#16",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista17": { nombre: "Capturista 17",    password: "cap2026#17",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista18": { nombre: "Capturista 18",    password: "cap2026#18",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista19": { nombre: "Capturista 19",    password: "cap2026#19",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista20": { nombre: "Capturista 20",    password: "cap2026#20",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista21": { nombre: "Capturista 21",    password: "cap2026#21",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista22": { nombre: "Capturista 22",    password: "cap2026#22",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista23": { nombre: "Capturista 23",    password: "cap2026#23",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista24": { nombre: "Capturista 24",    password: "cap2026#24",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false },
  "capturista25": { nombre: "Capturista 25",    password: "cap2026#25",   rol: "capturista", activo: true,  puedeEditar: true,  puedeEliminar: false }
};

/* ─────────────────────────────────────────────
   UTILIDADES INTERNAS
───────────────────────────────────────────── */

// _coordStr: Oaxaca tiene 2 dígitos enteros en lat (15-18) y lng (93-98).
// Cuando Sheets borra el punto: 17.072128 → 17072128
// Corrección: siempre insertar punto después del dígito 2.
function _coordStr(val) {
  if (!val || val === "") return "";
  var s = String(val).replace(/[^0-9.\-]/g, "").trim();
  if (!s || s === "0") return "";
  var neg = (s[0] === "-");
  var digits = neg ? s.slice(1) : s;

  if (digits.indexOf(".") !== -1) {
    // Ya tiene punto — reconstruir limpio desde número
    var n = parseFloat(s);
    if (isNaN(n) || n === 0) return "";
    // Construir sin toFixed: multiplicar y reconstruir
    var abs = Math.abs(n);
    var str = String(Math.round(abs * 1000000));
    while (str.length < 8) str = "0" + str;
    return (neg ? "-" : "") + str.substring(0, 2) + "." + str.substring(2);
  }

  // Sin punto (guardado mal): insertar después del dígito 2
  while (digits.length < 8) digits = "0" + digits;
  return (neg ? "-" : "") + digits.substring(0, 2) + "." + digits.substring(2);
}

// Aplica formato texto (@) a las celdas de coordenadas de una fila específica.
// Esto evita que Sheets convierta "17.072128" a número 17072128.
// Cols 39=Latitud, 40=Longitud (base-1)
function _forzarTextoCoordsEnFila(sheet, fila) {
  try {
    sheet.getRange(fila, 39).setNumberFormat("@").setValue(
      sheet.getRange(fila, 39).getValue()
    );
    sheet.getRange(fila, 40).setNumberFormat("@").setValue(
      sheet.getRange(fila, 40).getValue()
    );
  } catch(e) {
    Logger.log("_forzarTextoCoordsEnFila error: " + e.message);
  }
}

function _getSheet() {
  var sheet = SpreadsheetApp
    .openById(CONFIG.SHEET_ID)
    .getSheetByName(CONFIG.SHEETS.BENEFICIARIOS);
  if (!sheet) throw new Error("No existe la hoja BENEFICIARIOS");
  return sheet;
}

// _coordToStr: igual que _coordStr pero recibe número — preserva signo negativo
function _coordToStr(n) {
  if (!n || isNaN(n) || n === 0) return "";
  return _coordStr((n < 0 ? "-" : "") + String(Math.abs(n)));
}

function _formatRow(row) {
  var COL_LAT = 38;
  var COL_LNG = 39;

  return row.map(function(cell, idx) {
    if (cell instanceof Date) {
      return Utilities.formatDate(cell, CONFIG.TZ, "dd/MM/yyyy HH:mm");
    }
    if (cell === null || cell === undefined || cell === "") return "";

    if (idx === COL_LAT || idx === COL_LNG) {
      var num;
      if (typeof cell === "number" && cell !== 0) {
        num = cell; // número nativo de Sheets (puede ser negativo ✅)
      } else {
        var s = String(cell).replace(/[^0-9.\-]/g, "").trim();
        if (!s) return "";
        if (s.indexOf(".") === -1 && s.replace("-","").length >= 7) {
          // Sin punto decimal — insertar después del dígito 2
          var neg = s[0] === "-";
          var d = neg ? s.slice(1) : s;
          s = (neg ? "-" : "") + d.substring(0, 2) + "." + d.substring(2, 8);
        }
        num = parseFloat(s);
      }
      var resultado = _coordToStr(num);
      // Longitud en Oaxaca siempre negativa: si llega positiva en rango 93-99, agregar signo
      if (idx === COL_LNG && resultado && resultado[0] !== "-") {
        var n = parseFloat(resultado);
        if (!isNaN(n) && n >= 93 && n <= 99) resultado = "-" + resultado;
      }
      return resultado;
    }

    return String(cell);
  });
}

// Quita acentos para búsqueda más tolerante
function _sinAcentos(str) {
  return String(str || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

/* ─────────────────────────────────────────────
   LOGIN
───────────────────────────────────────────── */

function validarLogin(usuario, password) {
  var u = USUARIOS[String(usuario).trim()];
  if (!u || !u.activo || u.password !== password) return { status: false };
  var config = obtenerConfig();
  return {
    status:           true,
    user:             usuario,
    nombre:           u.nombre,
    rol:              u.rol,
    puedeEditar:      u.puedeEditar   || false,
    puedeEliminar:    u.puedeEliminar || false,
    territorioActivo: u.rol === "admin" ? true : config.territorioActivo
  };
}

/* ─────────────────────────────────────────────
   GUARDAR BENEFICIARIO
   Lock de script: serializa escrituras de los 30 usuarios.
   appendRow es atómico en Sheets — el lock solo protege
   la generación del ID único y el número correlativo.
───────────────────────────────────────────── */

function guardarBeneficiario(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000); // espera hasta 12 seg si hay cola de usuarios
  } catch(e) {
    throw new Error("El sistema está ocupado, intenta en unos segundos.");
  }

  try {
    var sheet   = _getSheet();
    var lastRow = sheet.getLastRow();
    var noFila  = lastRow; // fila 1 = encabezado → registro 1 está en fila 2, su No.=1

    var ahora      = new Date();
    var idRegistro = Utilities.formatDate(ahora, CONFIG.TZ, "yyyyMMdd")
                   + "-" + String(noFila)
                   + "-" + String(Math.floor(Math.random() * 900) + 100);
    var fechaStr   = Utilities.formatDate(ahora, CONFIG.TZ, "dd/MM/yyyy HH:mm");

    /*
      MAPA DE COLUMNAS — NO modificar sin actualizar EF_IDX en index.html
      Col 1  = No.               Col 2  = Fecha             Col 3  = ID Registro
      Col 4  = PrimerApellido    Col 5  = SegundoApellido   Col 6  = Nombres
      Col 7  = CURP              Col 8  = Edad              Col 9  = Telefono
      Col 10 = PersonasVivienda  Col 11 = HablaLengua       Col 12 = LenguaIndigena
      Col 13 = PoblacionIndigena Col 14 = PoblacionAfro     Col 15 = TieneDiscapacidad
      Col 16 = TipoDiscapacidad  Col 17 = Calle             Col 18 = NumExterior
      Col 19 = NumInterior       Col 20 = Colonia           Col 21 = Localidad
      Col 22 = Municipio         Col 23 = CP                Col 24 = PersonasPrograma
      Col 25 = TipoVivienda      Col 26 = IngresoMensual    Col 27 = NumDormitorios
      Col 28 = MaterialPiso      Col 29 = MaterialTecho     Col 30 = MaterialMuros
      Col 31 = AguaPotable       Col 32 = Drenaje           Col 33 = Electricidad
      Col 34 = Gas               Col 35 = Internet          Col 36 = DanoEstructural
      Col 37 = DanoNatural       Col 38 = ZonaRiesgo        Col 39 = Latitud
      Col 40 = Longitud          Col 41 = JefeFamilia       Col 42 = HabitantesCarac
      Col 43 = Capturista        Col 44 = FotoFachada       Col 45 = FotoInterior
      Col 46 = FotoEntorno       Col 47 = TenenciaPredio    Col 48 = EspacioConstruccion
      Col 49 = CondicionesAcceso Col 50 = DocumentoPredio   Col 51 = TipoEmpleo
    */
    // Preparar coords como strings con punto decimal garantizado
    var latStr = _coordStr(data.latitud);
    var lngStr = _coordStr(data.longitud);

    sheet.appendRow([
      noFila,                          //  1
      fechaStr,                        //  2
      idRegistro,                      //  3
      data.primerApellido    || "",    //  4
      data.segundoApellido   || "",    //  5
      data.nombres           || "",    //  6
      data.curp              || "",    //  7
      data.edad              || "",    //  8
      data.telefono          || "",    //  9
      data.personasVivienda  || "",    // 10
      data.hablaLengua       || "",    // 11
      data.lenguaIndigena    || "",    // 12
      data.poblacionIndigena || "",    // 13
      data.poblacionAfro     || "",    // 14
      data.tieneDiscapacidad || "",    // 15
      data.tipoDiscapacidad  || "",    // 16
      data.calle             || "",    // 17
      data.numExterior       || "",    // 18
      data.numInterior       || "",    // 19
      data.colonia           || "",    // 20
      data.localidad         || "",    // 21
      data.municipio         || "",    // 22
      data.cp                || "",    // 23
      data.personasPrograma  || "",    // 24
      data.tipoVivienda      || "",    // 25
      data.ingresoMensual    || "",    // 26
      data.numDormitorios    || "",    // 27
      data.materialPiso      || "",    // 28
      data.materialTecho     || "",    // 29
      data.materialMuros     || "",    // 30
      data.aguaPotable       || "",    // 31
      data.drenaje           || "",    // 32
      data.electricidad      || "",    // 33
      data.gas               || "",    // 34
      data.internet          || "",    // 35
      data.danoEstructural   || "",    // 36
      data.danoNatural       || "",    // 37
      data.zonaRiesgo        || "",    // 38
      "",                              // 39 Latitud  → se escribe abajo como texto
      "",                              // 40 Longitud → se escribe abajo como texto
      data.jefeFamilia       || "",    // 41
      data.habitantesCarac   || "",    // 42
      data.capturista        || "",    // 43
      "",                              // 44 FotoFachada
      "",                              // 45 FotoInterior
      "",                              // 46 FotoEntorno
      data.tipoTenenciaPredio  || "",  // 47
      data.espacioConstruccion || "",  // 48
      data.condicionesAcceso   || "",  // 49
      data.documentoPredio     || "",  // 50
      data.tipoEmpleo          || "",  // 51
      data.nombreTitular       || "",  // 52
      data.telefonoTitular     || "",  // 53
      data.linkUbicacion       || "",  // 54
      data.llamadaRealizada    || ""   // 55
    ]);

    // Escribir coords por separado como texto puro — appendRow las convertiría a número
    var filaGuardada = sheet.getLastRow();
    if (latStr) {
      sheet.getRange(filaGuardada, 39).setNumberFormat("@").setValue(latStr);
    }
    if (lngStr) {
      sheet.getRange(filaGuardada, 40).setNumberFormat("@").setValue(lngStr);
    }

    return { idRegistro: idRegistro };

  } finally {
    lock.releaseLock();
  }
}

/* ─────────────────────────────────────────────
   GUARDAR REFERIDO (caso No Aplica)
   Se guarda en una hoja separada "REFERIDOS"
   con: fecha, nombre referido, teléfono, solicitante, capturista
───────────────────────────────────────────── */
function guardarReferido(data) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { throw new Error("Sistema ocupado, intenta de nuevo."); }
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    var sheet = ss.getSheetByName("REFERIDOS");
    if (!sheet) {
      // Crear la hoja si no existe con encabezados
      sheet = ss.insertSheet("REFERIDOS");
      sheet.appendRow(["Fecha","Nombre Referido","Teléfono Referido","Solicitante (quien hizo el trámite)","Capturista"]);
    }
    var fechaStr = Utilities.formatDate(new Date(), CONFIG.TZ, "dd/MM/yyyy HH:mm");
    sheet.appendRow([
      fechaStr,
      data.nombreReferido   || "",
      data.telefonoReferido || "",
      data.solicitante      || "",
      data.capturista       || ""
    ]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ─────────────────────────────────────────────
   GUARDAR FOTO EN DRIVE
   Sin lock: cada foto opera en su propia celda por ID único,
   no hay riesgo de colisión entre usuarios distintos.
───────────────────────────────────────────── */

function guardarFotoEnDrive(idRegistro, nombreBeneficiario, fotoObj) {
  var colFoto = { fachada: 44, interior: 45, entorno: 46 };
  var col = colFoto[fotoObj.tipo];
  if (!col) throw new Error("Tipo de foto no válido: " + fotoObj.tipo);

  // Guardar en Drive
  var bytes   = Utilities.base64Decode(fotoObj.base64);
  var blob    = Utilities.newBlob(bytes, fotoObj.mimeType,
                  nombreBeneficiario + "_" + fotoObj.tipo + "_" + idRegistro + ".jpg");
  var folder  = _obtenerCarpetaFotos();
  var archivo = folder.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = "https://drive.google.com/uc?id=" + archivo.getId();

  // Encontrar la fila por ID (solo lee columna 3, rapidísimo)
  var sheet  = _getSheet();
  var ultima = sheet.getLastRow();
  if (ultima < 2) throw new Error("Sin registros en la hoja.");

  var ids = sheet.getRange(2, 3, ultima - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(idRegistro)) {
      sheet.getRange(i + 2, col).setValue(url);
      return { ok: true };
    }
  }
  throw new Error("ID no encontrado al guardar foto: " + idRegistro);
}

function _obtenerCarpetaFotos() {
  var nombre = "Fotos_Beneficiarios";
  var iter   = DriveApp.getFoldersByName(nombre);
  if (iter.hasNext()) return iter.next();
  return DriveApp.createFolder(nombre);
}

// Edición reutiliza la misma lógica de guardado de foto
function guardarFotoEdicion(idRegistro, nombreBeneficiario, fotoObj) {
  return guardarFotoEnDrive(idRegistro, nombreBeneficiario, fotoObj);
}

/* ─────────────────────────────────────────────
   OBTENER ÚLTIMOS REGISTROS
   Lee exactamente N filas desde el final — nada más.
───────────────────────────────────────────── */

function obtenerUltimosRegistros(n) {
  var sheet   = _getSheet();
  var ultima  = sheet.getLastRow();
  if (ultima < 2) return [];

  var cantidad = Math.min(n || 1, ultima - 1);
  var inicio   = Math.max(2, ultima - cantidad + 1);
  var numFilas = ultima - inicio + 1;
  var cols     = sheet.getLastColumn() || 51;

  var data = sheet.getRange(inicio, 1, numFilas, cols).getValues();
  data.reverse(); // más nuevo primero
  return data.map(_formatRow);
}

/* ─────────────────────────────────────────────
   BUSCAR REGISTRO
   Solo busca en: PrimerApellido (col4), SegundoApellido (col5), Nombres (col6).
   Tolerante a acentos y mayúsculas.
   Devuelve máximo 30 resultados.
───────────────────────────────────────────── */

function buscarRegistro(query) {
  if (!query || !query.trim()) return [];
  var sheet  = _getSheet();
  var ultima = sheet.getLastRow();
  if (ultima < 2) return [];

  var q    = _sinAcentos(query.trim());
  var cols = sheet.getLastColumn() || 51;

  // Leer toda la hoja en UNA sola llamada a la API
  var data = sheet.getRange(2, 1, ultima - 1, cols).getValues();
  var res  = [];

  for (var i = 0; i < data.length; i++) {
    var r = data[i];

    // Solo busca en: PrimerApellido[3], SegundoApellido[4], Nombres[5]
    var haystack = _sinAcentos(
      String(r[3] || "") + " " +   // Primer apellido
      String(r[4] || "") + " " +   // Segundo apellido
      String(r[5] || "")            // Nombre(s)
    );

    if (haystack.indexOf(q) !== -1) {
      res.push(_formatRow(r));
      if (res.length >= 30) break; // máximo 30 resultados
    }
  }

  return res;
}

/* ─────────────────────────────────────────────
   EDITAR REGISTRO
   - Lock de script para serializar ediciones concurrentes
   - UN SOLO setValues() para toda la fila (40x más rápido)
   - Busca por ID de registro (col 3), nunca por número de fila
───────────────────────────────────────────── */

function editarRegistro(idRegistro, data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000);
  } catch(e) {
    throw new Error("El sistema está ocupado, intenta de nuevo.");
  }

  try {
    var sheet  = _getSheet();
    var ultima = sheet.getLastRow();
    if (ultima < 2) throw new Error("No hay registros.");

    var cols    = sheet.getLastColumn() || 51;
    var valores = sheet.getRange(2, 1, ultima - 1, cols).getValues();
    var filaIdx = -1;

    // Buscar fila por ID de registro (columna 3, índice [2])
    for (var i = 0; i < valores.length; i++) {
      if (String(valores[i][2]) === String(idRegistro)) {
        filaIdx = i;
        break;
      }
    }
    if (filaIdx === -1) throw new Error("Registro no encontrado: " + idRegistro);

    // Copiar fila actual y sobreescribir solo los campos editables
    var fila = valores[filaIdx].slice();

    // Preservar: fila[0]=No., fila[1]=Fecha, fila[2]=ID, fila[42]=capturista, fila[43-45]=fotos
    fila[3]  = data.primerApellido    !== undefined ? (data.primerApellido    || fila[3])  : fila[3];
    fila[4]  = data.segundoApellido   !== undefined ? (data.segundoApellido   || fila[4])  : fila[4];
    fila[5]  = data.nombres           !== undefined ? (data.nombres           || fila[5])  : fila[5];
    fila[6]  = data.curp              !== undefined ? (data.curp              || fila[6])  : fila[6];
    fila[7]  = data.edad              !== undefined ? data.edad              : fila[7];
    fila[8]  = data.telefono          !== undefined ? data.telefono          : fila[8];
    fila[9]  = data.personasVivienda  !== undefined ? data.personasVivienda  : fila[9];
    fila[10] = data.hablaLengua       !== undefined ? data.hablaLengua       : fila[10];
    fila[11] = data.lenguaIndigena    !== undefined ? data.lenguaIndigena    : fila[11];
    fila[12] = data.poblacionIndigena !== undefined ? data.poblacionIndigena : fila[12];
    fila[13] = data.poblacionAfro     !== undefined ? data.poblacionAfro     : fila[13];
    fila[14] = data.tieneDiscapacidad !== undefined ? data.tieneDiscapacidad : fila[14];
    fila[15] = data.tipoDiscapacidad  !== undefined ? data.tipoDiscapacidad  : fila[15];
    fila[16] = data.calle             !== undefined ? data.calle             : fila[16];
    fila[17] = data.numExterior       !== undefined ? data.numExterior       : fila[17];
    fila[18] = data.numInterior       !== undefined ? data.numInterior       : fila[18];
    fila[19] = data.colonia           !== undefined ? data.colonia           : fila[19];
    fila[20] = data.localidad         !== undefined ? data.localidad         : fila[20];
    fila[21] = data.municipio         !== undefined ? data.municipio         : fila[21];
    fila[22] = data.cp                !== undefined ? data.cp                : fila[22];
    fila[23] = data.personasPrograma  !== undefined ? data.personasPrograma  : fila[23];
    fila[24] = data.tipoVivienda      !== undefined ? data.tipoVivienda      : fila[24];
    fila[25] = data.ingresoMensual    !== undefined ? data.ingresoMensual    : fila[25];
    fila[26] = data.numDormitorios    !== undefined ? data.numDormitorios    : fila[26];
    fila[27] = data.materialPiso      !== undefined ? data.materialPiso      : fila[27];
    fila[28] = data.materialTecho     !== undefined ? data.materialTecho     : fila[28];
    fila[29] = data.materialMuros     !== undefined ? data.materialMuros     : fila[29];
    fila[30] = data.aguaPotable       !== undefined ? data.aguaPotable       : fila[30];
    fila[31] = data.drenaje           !== undefined ? data.drenaje           : fila[31];
    fila[32] = data.electricidad      !== undefined ? data.electricidad      : fila[32];
    fila[33] = data.gas               !== undefined ? data.gas               : fila[33];
    fila[34] = data.internet          !== undefined ? data.internet          : fila[34];
    fila[35] = data.danoEstructural   !== undefined ? data.danoEstructural   : fila[35];
    fila[36] = data.danoNatural       !== undefined ? data.danoNatural       : fila[36];
    fila[37] = data.zonaRiesgo        !== undefined ? data.zonaRiesgo        : fila[37];
    fila[38] = data.latitud  !== undefined ? _coordStr(data.latitud)  : fila[38];
    fila[39] = data.longitud !== undefined ? _coordStr(data.longitud) : fila[39];
    fila[40] = data.jefeFamilia       !== undefined ? data.jefeFamilia       : fila[40];
    fila[41] = data.habitantesCarac   !== undefined ? data.habitantesCarac   : fila[41];
    // fila[42] = capturista          → NO se toca
    // fila[43,44,45] = fotos         → NO se tocan aquí
    fila[46] = data.tipoTenenciaPredio  !== undefined ? data.tipoTenenciaPredio  : fila[46];
    fila[47] = data.espacioConstruccion !== undefined ? data.espacioConstruccion : fila[47];
    fila[48] = data.condicionesAcceso   !== undefined ? data.condicionesAcceso   : fila[48];
    fila[49] = data.documentoPredio     !== undefined ? data.documentoPredio     : fila[49];
    fila[50] = data.tipoEmpleo          !== undefined ? data.tipoEmpleo          : fila[50];
    fila[51] = data.nombreTitular       !== undefined ? data.nombreTitular       : fila[51];
    fila[52] = data.telefonoTitular     !== undefined ? data.telefonoTitular     : fila[52];
    fila[53] = data.linkUbicacion       !== undefined ? data.linkUbicacion       : fila[53];
    fila[54] = data.llamadaRealizada    !== undefined ? data.llamadaRealizada    : fila[54];

    // Guardar coords aparte — sacarlas del array principal para evitar autoconversión
    var latStr = data.latitud  !== undefined ? _coordStr(data.latitud)  : _coordStr(String(fila[38] || ""));
    var lngStr = data.longitud !== undefined ? _coordStr(data.longitud) : _coordStr(String(fila[39] || ""));
    fila[38] = "";  // se escribe abajo como texto
    fila[39] = "";  // se escribe abajo como texto

    // UNA sola escritura a Sheets para todos los demás campos
    sheet.getRange(filaIdx + 2, 1, 1, fila.length).setValues([fila]);

    // Coords como texto puro — separadas para evitar autoconversión
    if (latStr) sheet.getRange(filaIdx + 2, 39).setNumberFormat("@").setValue(latStr);
    if (lngStr) sheet.getRange(filaIdx + 2, 40).setNumberFormat("@").setValue(lngStr);

    return { ok: true };

  } finally {
    lock.releaseLock();
  }
}

/* ─────────────────────────────────────────────
   ELIMINAR REGISTRO
───────────────────────────────────────────── */

function eliminarRegistro(idRegistro) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch(e) {
    throw new Error("Sistema ocupado, intenta de nuevo.");
  }

  try {
    var sheet  = _getSheet();
    var ultima = sheet.getLastRow();
    if (ultima < 2) throw new Error("No hay registros.");

    // Lee solo la columna de IDs (col 3) — mínimo posible
    var ids = sheet.getRange(2, 3, ultima - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(idRegistro)) {
        sheet.deleteRow(i + 2);
        return { ok: true };
      }
    }
    throw new Error("Registro no encontrado: " + idRegistro);

  } finally {
    lock.releaseLock();
  }
}

/* ─────────────────────────────────────────────
   CATÁLOGOS
───────────────────────────────────────────── */

/* ─────────────────────────────────────────────
   CATÁLOGOS — hoja "Catálogos"
   Lee por NOMBRE de columna en la fila 1 (encabezado).
   Columnas esperadas:
     "Lengua indígena"              → lenguas
     "Nombre del Municipio"         → municipios
     "Clave Geoestadística Estatal" → clave estatal
     "Clave Geoestadística Municipal" → clave municipal
     "Nombre de Localidad"          → localidades
     "d_codigo"                     → código postal
     "d_asenta"                     → colonia/asentamiento
     "calles"                       → sugerencias de calles
───────────────────────────────────────────── */

function _getCataIdx() {
  var sheet  = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName("Catálogos");
  if (!sheet) throw new Error("No existe la hoja 'Catálogos'");
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idx = {};
  for (var c = 0; c < header.length; c++) {
    var h = String(header[c]).trim();
    if (h) idx[h] = c;
  }
  var data = sheet.getDataRange().getValues().slice(1); // sin encabezado
  return { idx: idx, data: data };
}

// ── Municipios ──────────────────────────────
function obtenerMunicipios() {
  try {
    var cat  = _getCataIdx();
    var cMun = cat.idx["Nombre del Municipio"];
    var cEst = cat.idx["Clave Geoestadística Estatal"];
    var cCve = cat.idx["Clave Geoestadística Municipal"];
    if (cMun === undefined) { Logger.log("Columna 'Nombre del Municipio' no encontrada"); return []; }
    var res = [], visto = {};
    for (var i = 0; i < cat.data.length; i++) {
      var nombre = String(cat.data[i][cMun] || "").trim();
      if (!nombre || visto[nombre]) continue;
      visto[nombre] = true;
      res.push({
        nombre: nombre,
        cveEst: cEst !== undefined ? String(cat.data[i][cEst] || "").trim() : "",
        cveMun: cCve !== undefined ? String(cat.data[i][cCve] || "").trim() : ""
      });
    }
    Logger.log("Municipios: " + res.length);
    return res;
  } catch(e) { Logger.log("obtenerMunicipios: " + e.message); return []; }
}

// ── Localidades + Colonias + CP + Municipio ─
// Devuelve [localidad, colonia, cp, municipio]
function obtenerCatalogoLocalidades() {
  try {
    var cat  = _getCataIdx();
    var cLoc = cat.idx["Nombre de Localidad"];
    var cCol = cat.idx["d_asenta"];
    var cCp  = cat.idx["d_codigo"];
    var cMun = cat.idx["Nombre del Municipio"];
    var res  = [];
    for (var i = 0; i < cat.data.length; i++) {
      var loc = cLoc !== undefined ? String(cat.data[i][cLoc] || "").trim() : "";
      var col = cCol !== undefined ? String(cat.data[i][cCol] || "").trim() : "";
      var cp  = cCp  !== undefined ? String(cat.data[i][cCp]  || "").trim() : "";
      var mun = cMun !== undefined ? String(cat.data[i][cMun] || "").trim() : "";
      if (!loc && !col) continue;
      res.push([loc, col, cp, mun]);
    }
    Logger.log("Localidades/Colonias: " + res.length);
    return res;
  } catch(e) { Logger.log("obtenerCatalogoLocalidades: " + e.message); return []; }
}

// ── Calles ───────────────────────────────────
function obtenerCatalogoCalles() {
  try {
    var cat  = _getCataIdx();
    var cCal = cat.idx["calles"];
    if (cCal === undefined) { Logger.log("Columna 'calles' no encontrada"); return []; }
    var res = [], visto = {};
    for (var i = 0; i < cat.data.length; i++) {
      var nombre = String(cat.data[i][cCal] || "").trim();
      if (!nombre || visto[nombre]) continue;
      visto[nombre] = true;
      res.push(nombre);
    }
    Logger.log("Calles: " + res.length);
    return res;
  } catch(e) { Logger.log("obtenerCatalogoCalles: " + e.message); return []; }
}

// ── Lenguas Indígenas ────────────────────────
function obtenerCatalogoLenguas() {
  try {
    var cat  = _getCataIdx();
    var cLen = cat.idx["Lengua indígena"];
    if (cLen === undefined) { Logger.log("Columna 'Lengua indígena' no encontrada"); return []; }
    var res = [], visto = {};
    for (var i = 0; i < cat.data.length; i++) {
      var nombre = String(cat.data[i][cLen] || "").trim();
      if (!nombre || visto[nombre]) continue;
      visto[nombre] = true;
      res.push(nombre);
    }
    Logger.log("Lenguas: " + res.length);
    return res;
  } catch(e) { Logger.log("obtenerCatalogoLenguas: " + e.message); return []; }
}

/* ─────────────────────────────────────────────
   ENTRADA WEB
───────────────────────────────────────────── */

/* ─────────────────────────────────────────────
   UTILIDAD — Ejecutar UNA VEZ para corregir coords existentes
   Desde el editor de Apps Script: Ejecutar → repararCoordenadas
───────────────────────────────────────────── */
function repararCoordenadas() {
  var sheet  = _getSheet();
  var ultima = sheet.getLastRow();
  if (ultima < 2) { Logger.log("Sin registros"); return; }

  // Aplicar formato texto a todas las celdas de lat/lng de una vez
  sheet.getRange(2, 39, ultima - 1, 2).setNumberFormat("@");

  // Ahora corregir los valores que llegaron sin punto (ej: 17072128 → 17.072128)
  var datos = sheet.getRange(2, 39, ultima - 1, 2).getValues();
  var corregidos = 0;

  for (var i = 0; i < datos.length; i++) {
    for (var c = 0; c < 2; c++) {
      var val = datos[i][c];
      if (!val || val === "") continue;

      var s = String(val).replace(/[^0-9.\-]/g, "").trim();
      if (!s || s === "0") continue;

      // Si ya tiene punto decimal, no tocar
      if (s.indexOf(".") !== -1) continue;

      // Sin punto — reconstruir: insertar punto después de 2 dígitos
      var neg = s[0] === "-";
      var digits = neg ? s.slice(1) : s;
      if (digits.length >= 7) {
        var corregido = (neg ? "-" : "") + digits.substring(0, 2) + "." + digits.substring(2);
        datos[i][c] = corregido;
        corregidos++;
      }
    }
  }

  if (corregidos > 0) {
    sheet.getRange(2, 39, ultima - 1, 2).setValues(datos);
    Logger.log("Coordenadas corregidas: " + corregidos + " celdas");
  } else {
    Logger.log("No se encontraron coordenadas sin punto — todo OK");
  }
}

/* ─────────────────────────────────────────────
   CONFIGURACIÓN DEL SISTEMA
   Usa PropertiesService para persistir entre sesiones.
   La clave "territorioActivo" controla si los capturistas
   pueden ver y llenar la sección de territorio.
───────────────────────────────────────────── */

function obtenerConfig() {
  var props = PropertiesService.getScriptProperties();
  return {
    territorioActivo: props.getProperty("territorioActivo") === "true"
  };
}

function guardarConfig(config) {
  // Solo admin puede cambiar configuración
  // (la verificación de rol se hace en el frontend con SESION)
  var props = PropertiesService.getScriptProperties();
  if (config.territorioActivo !== undefined) {
    props.setProperty("territorioActivo", config.territorioActivo ? "true" : "false");
  }
  return { ok: true, territorioActivo: props.getProperty("territorioActivo") === "true" };
}


/* ─────────────────────────────────────────────
   MIGRACIÓN — corregir coordenadas existentes
   Ejecutar UNA SOLA VEZ desde el editor de Apps Script:
     migrarCoordenadas()
   Recorre todas las filas y corrige "17072128" → "17.072128"
───────────────────────────────────────────── */
function migrarCoordenadas() {
  var sheet  = _getSheet();
  var ultima = sheet.getLastRow();
  if (ultima < 2) { Logger.log("Sin registros"); return; }

  // Primero aplicar formato texto a todas las celdas de coords
  sheet.getRange(2, 39, ultima - 1, 2).setNumberFormat("@");

  var rango  = sheet.getRange(2, 39, ultima - 1, 2);
  var valores = rango.getValues();
  var cambios = 0;

  for (var i = 0; i < valores.length; i++) {
    for (var c = 0; c < 2; c++) {
      var raw = valores[i][c];
      if (raw === null || raw === "" || raw === 0) continue;
      var s   = String(raw).replace(/[^0-9.\-]/g, "").trim();
      if (!s) continue;

      var corregido = "";
      if (s.indexOf(".") !== -1) {
        // Ya tiene punto — solo normalizar
        var n = parseFloat(s);
        if (!isNaN(n) && n !== 0) corregido = _coordToStr(n);
      } else if (s.replace("-","").length >= 7) {
        // Sin punto — reconstruir
        var neg = s[0] === "-";
        var d   = neg ? s.slice(1) : s;
        var rec = (neg ? "-" : "") + d.substring(0,2) + "." + d.substring(2);
        var n2  = parseFloat(rec);
        if (!isNaN(n2) && n2 !== 0) corregido = _coordToStr(n2);
      }

      if (corregido && corregido !== String(raw)) {
        valores[i][c] = corregido;
        cambios++;
      }
    }
  }

  if (cambios > 0) {
    rango.setValues(valores);
    Logger.log("Coordenadas corregidas: " + cambios + " celdas");
  } else {
    Logger.log("Todas las coordenadas ya estaban correctas.");
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Registro de Beneficiarios")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}