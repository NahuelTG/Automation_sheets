var libroCbba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2024"); // Reemplaza "2024" con el nombre correcto de tu hoja
var plazosPagosCbba = SpreadsheetApp.openByUrl(
  "https://docs.google.com/spreadsheets/d/1XcqOdbIIb5FZQiAn8hOA6u7y1UPEaxsBY-PluzfgJYs/edit?usp=sharing"
).getSheetByName("TESISTAS");

// Listas de abreviaturas por categorías
var categoriaA = [
  "arts. plásticas",
  "dis. gráfico",
  "art. musicales",
  "antro. y arqueología",
  "com. social",
  "sociología",
  "trab. social",
  "derecho",
  "cs. políticas",
  "rel. internacionales",
  "cs. información",
  "cs. educación",
  "filosofía",
  "historia",
  "lingüística",
  "literatura",
  "psicología",
  "turismo",
];
var categoriaB = [
  "arquitectura",
  "veterinaria",
  "ing. agronómica",
  "ing. agropecuaria",
  "adm. empresas",
  "contaduría",
  "economía",
  "marketing",
  "bioquímica",
  "farmacéutica",
  "ing. geográfica",
  "ing. geológica",
  "medicina",
  "enfermería",
  "nutrición",
  "tec. médica",
  "odontología",
  "aeronáutica",
  "cons. civiles",
  "elec. industrial",
  "telecomunicaciones",
  "ing. electromecánica",
  "mec. automotriz",
  "mec. industrial",
  "qui. industrial",
  "ing. topografía",
  "biología",
  "cs. químicas",
  "estadística",
  "física",
  "informática",
  "matemáticas",
  "industrial",
  "comercial",
];

var ingresoPrimero = 0;
var ingresoSegundo = 0;
var ingresoTercero = 0;

function actualizarDatos(filaEditada, codigo) {
  var nombreCliente = libroCbba.getRange("D" + filaEditada).getValue();
  var concepto = libroCbba.getRange("E" + filaEditada).getValue();
  var ingreso = libroCbba.getRange("G" + filaEditada).getValue();
  var tipoCuota = determinarCuotas(concepto, categoriaA, categoriaB);
  var nombreContrato =
    concepto.split(",").length > 1
      ? concepto.split(",")[1].trim().toUpperCase()
      : "";

  try {
    if (filaEditada < 2) {
      // No hacer nada si se edita en la cabecera
      return;
    }

    var cliente = Iteracion(codigo);
    var datosOtro = plazosPagosCbba
      .getRange("A3:A" + plazosPagosCbba.getLastRow())
      .getValues();
    var filaPlazos = 0;

    for (var j = 0; j < cliente.length; j++) {
      if (
        cliente[j].concepto.includes("1ra cuota") &&
        verificarSubCuota(cliente[j].concepto)
      ) {
        ingresoPrimero = ingresoPrimero + cliente[j].ingreso;
      }
      if (
        cliente[j].concepto.includes("2da cuota") &&
        verificarSubCuota(cliente[j].concepto)
      ) {
        ingresoSegundo = ingresoSegundo + cliente[j].ingreso;
      }
      if (
        cliente[j].concepto.includes("3ra cuota") &&
        verificarSubCuota(cliente[j].concepto)
      ) {
        ingresoTercero = ingresoTercero + cliente[j].ingreso;
      }
    }

    for (var j = 0; j < datosOtro.length; j++) {
      if (datosOtro[j][0] == codigo) {
        filaPlazos = j;
        if (concepto.includes("1ra cuota")) {
          plazosPagosCbba.getRange("B" + (j + 3)).setValue(nombreContrato);
          plazosPagosCbba.getRange("C" + (j + 3)).setValue(nombreCliente);
          plazosPagosCbba.getRange("F" + (j + 3)).setValue("PRIMERA CUOTA");
          plazosPagosCbba.getRange("G" + (j + 3)).setValue(tipoCuota.primera);
          plazosPagosCbba.getRange("H" + (j + 3)).setValue("SEGUNDA CUOTA");
          plazosPagosCbba.getRange("I" + (j + 3)).setValue(tipoCuota.segunda);
          plazosPagosCbba.getRange("J" + (j + 3)).setValue("ÚLTIMA CUOTA");
          if (!concepto.includes("descuento")) {
            plazosPagosCbba.getRange("K" + (j + 3)).setValue(tipoCuota.ultima);
          }
        }
        break;
      }
    }
    if (verificarSubCuota(concepto)) {
      if (concepto.includes("1ra cuota")) {
        verificarCuotas(
          filaPlazos + 3,
          plazosPagosCbba,
          ingresoPrimero,
          concepto,
          verificarSubCuota(concepto)
        );
      }
      if (concepto.includes("2da cuota")) {
        verificarCuotas(
          filaPlazos + 3,
          plazosPagosCbba,
          ingresoSegundo,
          concepto,
          verificarSubCuota(concepto)
        );
      }
      if (concepto.includes("3ra cuota")) {
        verificarCuotas(
          filaPlazos + 3,
          plazosPagosCbba,
          ingresoTercero,
          concepto,
          verificarSubCuota(concepto)
        );
      }
    } else {
      verificarCuotas(
        filaPlazos + 3,
        plazosPagosCbba,
        ingreso,
        concepto,
        verificarSubCuota(concepto)
      );
    }
  } catch (e) {
    Logger.log(e.toString());
  }
}

function Iteracion(codigo) {
  var datosOtroLibro = libroCbba
    .getRange("B2:G" + libroCbba.getLastRow())
    .getValues();
  var filasCoincidentes = [];

  if (codigo && codigo.match(/^SPACBBOL \d+$/)) {
    for (var j = 0; j < datosOtroLibro.length; j++) {
      var fila = datosOtroLibro[j];
      var codigoFila = fila[0]; // Código en la columna B

      if (codigoFila == codigo) {
        var concepto = fila[3]; // Concepto en la columna E
        var nombre = fila[2]; // Nombre en la columna D
        var ingreso = fila[5]; // Ingreso en la columna G

        // Almacenar la fila en el array de filas coincidentes
        filasCoincidentes.push({
          concepto: concepto,
          nombre: nombre,
          ingreso: ingreso,
        });
      }
    }
  }
  return filasCoincidentes;
}

function determinarCuotas(concepto, categoriaA, categoriaB) {
  var cuotaPrimera = 2000;
  var cuotaSegunda = 0;
  var cuotaUltima = 0;
  var conceptoMinusculas = concepto.toLowerCase(); // Convertir concepto a minúsculas para la comparación
  var partesConcepto = concepto.split(",");
  var carreraEncontrada = null;

  // Verificar categoría A
  for (var i = 0; i < categoriaA.length; i++) {
    if (conceptoMinusculas.includes(categoriaA[i])) {
      carreraEncontrada = categoriaA[i];
      break;
    }
  }
  // Si no se encontró en categoría A, verificar categoría B
  if (!carreraEncontrada) {
    for (var j = 0; j < categoriaB.length; j++) {
      if (conceptoMinusculas.includes(categoriaB[j])) {
        carreraEncontrada = categoriaB[j];
        break;
      }
    }
  }
  // Verificar si es maestría
  var esMaestria = partesConcepto.some((parte) =>
    parte.trim().includes("mae.")
  );
  // Determinar tipoCuota basado en la carrera encontrada
  if (carreraEncontrada) {
    if (esMaestria) {
      cuotaSegunda = categoriaA.includes(carreraEncontrada) ? 2500 : 2600;
      cuotaUltima = categoriaA.includes(carreraEncontrada) ? 3000 : 3400;
    } else {
      cuotaSegunda = categoriaA.includes(carreraEncontrada) ? 2400 : 2500;
      cuotaUltima = categoriaA.includes(carreraEncontrada) ? 2600 : 3000;
    }
  }

  return {
    primera: cuotaPrimera,
    segunda: cuotaSegunda,
    ultima: cuotaUltima,
  };
}

function verificarSubCuota(concepto) {
  var match = concepto.match(/(\w),/);
  var subCuota = false;
  // Si se encuentra una letra antes de la coma
  if (match) {
    var tipoConcepto = match[1];
    // Verificar si el carácter es una letra mayúscula (ASCII 65-90)
    if (tipoConcepto.charCodeAt(0) >= 65 && tipoConcepto.charCodeAt(0) <= 90) {
      subCuota = true;
    }
  }
  return subCuota;
}

function verificarCuotas(fila, plazosPagosCbba, ingreso, concepto, subCuota) {
  var valorColumnaPri = plazosPagosCbba.getRange("G" + fila).getValue();
  var valorColumnaSeg = plazosPagosCbba.getRange("I" + fila).getValue();
  var valorColumnaTer = plazosPagosCbba.getRange("K" + fila).getValue();
  if (subCuota) {
    if (concepto.includes("1ra cuota")) {
      verificarYActualizarM(fila, plazosPagosCbba, "1ra cuota");
      if (valorColumnaPri == ingresoPrimero) {
        // Pintar celdas de verde en el otro sheet
        plazosPagosCbba.getRange("F" + fila).setBackground("#7cd455");
        plazosPagosCbba.getRange("G" + fila).setBackground("#7cd455");
        dejarObservaciones(fila, plazosPagosCbba);
      } else {
        plazosPagosCbba.getRange("F" + fila).setBackground("#ffff00");
        plazosPagosCbba.getRange("G" + fila).setBackground("#ffff00");
      }
    }
    if (concepto.includes("2da cuota")) {
      verificarYActualizarM(fila, plazosPagosCbba, "2da cuota");
      if (valorColumnaSeg == ingresoSegundo) {
        plazosPagosCbba.getRange("H" + fila).setBackground("#7cd455");
        plazosPagosCbba.getRange("I" + fila).setBackground("#7cd455");
        dejarObservaciones(fila, plazosPagosCbba);
      } else {
        plazosPagosCbba.getRange("H" + fila).setBackground("#ffff00");
        plazosPagosCbba.getRange("I" + fila).setBackground("#ffff00");
      }
    }
    if (concepto.includes("3ra cuota")) {
      verificarYActualizarM(fila, plazosPagosCbba, "3ra cuota");
      if (valorColumnaTer == ingresoTercero) {
        plazosPagosCbba.getRange("J" + fila).setBackground("#7cd455");
        plazosPagosCbba.getRange("K" + fila).setBackground("#7cd455");
        dejarObservaciones(fila, plazosPagosCbba);
      } else {
        plazosPagosCbba.getRange("J" + fila).setBackground("#ffff00");
        plazosPagosCbba.getRange("K" + fila).setBackground("#ffff00");
      }
    }
  } else {
    if (concepto.includes("1ra cuota")) {
      if (valorColumnaPri == ingreso) {
        // Pintar celdas de verde en el otro sheet
        plazosPagosCbba.getRange("F" + fila).setBackground("#7cd455");
        plazosPagosCbba.getRange("G" + fila).setBackground("#7cd455");
      }
    }
    if (concepto.includes("2da cuota")) {
      if (valorColumnaSeg == ingreso) {
        // Pintar celdas de verde en el otro sheet
        plazosPagosCbba.getRange("H" + fila).setBackground("#7cd455");
        plazosPagosCbba.getRange("I" + fila).setBackground("#7cd455");
      }
    }
    if (concepto.includes("3ra cuota")) {
      if (valorColumnaTer == ingreso) {
        // Pintar celdas de verde en el otro sheet
        plazosPagosCbba.getRange("J" + fila).setBackground("#7cd455");
        plazosPagosCbba.getRange("K" + fila).setBackground("#7cd455");
      }
    }
  }
}

function dejarObservaciones(fila, plazosPagosCbba) {
  const celda = plazosPagosCbba.getRange("M" + fila);
  const textoOriginal = celda.getValue();

  // Buscar "canceló" y última coma
  const indiceCancelo = textoOriginal.indexOf("canceló");
  const indiceUltimaComa = textoOriginal.lastIndexOf(",");

  if (indiceCancelo !== -1 && indiceUltimaComa !== -1) {
    const textoAntesCancelo = textoOriginal.substring(0, indiceCancelo).trim();
    const textoDespuesUltimaComa = textoOriginal
      .substring(indiceUltimaComa + 1)
      .trim();
    const nuevoTexto = textoAntesCancelo + " " + textoDespuesUltimaComa;

    celda.setValue(nuevoTexto);
  }
}

function verificarYActualizarM(fila, plazosPagosCbba, nroCuota) {
  var columnaM = plazosPagosCbba.getRange("M" + fila);
  var valorColumnaM = columnaM.getValue().trim();

  if (nroCuota == "1ra cuota") {
    if (valorColumnaM === "") {
      var cuotaValue = obtenerValorCuota(nroCuota, fila, plazosPagosCbba);
      var diferencia = cuotaValue - ingresoPrimero;
      columnaM.setValue(
        "canceló Bs." + ingresoPrimero + ", faltan Bs." + diferencia + ","
      );
      columnaM.setBackground("#ffff00");
    } else {
      // Separar el texto en partes
      var partes = separarTexto(valorColumnaM);

      var parte1 = partes[0]; // "El tesista canceló Bs."
      var ingresoPasado = parseFloat(partes[1]); // 1500
      var parte3 = partes[2]; // ", faltan Bs."
      var parte4 = partes[3]; // ","

      var nuevaDiferencia =
        obtenerValorCuota(nroCuota, fila, plazosPagosCbba) - ingresoPrimero;
      // Reconstruir el texto actualizado
      var textoActualizado =
        parte1 + ingresoPrimero + parte3 + nuevaDiferencia + parte4;
      // Borrar contenido de la celda antes de actualizar
      columnaM.clearContent();
      // Actualizar la columna M con el texto actualizado
      columnaM.setValue(textoActualizado);
    }
  }
  if (nroCuota == "2da cuota") {
    if (valorColumnaM === "") {
      var cuotaValue = obtenerValorCuota(nroCuota, fila, plazosPagosCbba);
      var diferencia = cuotaValue - ingresoSegundo;
      columnaM.setValue(
        "canceló Bs." + ingresoSegundo + ", faltan Bs." + diferencia + ","
      );
      columnaM.setBackground("#ffff00");
    } else {
      if (
        valorColumnaM.includes("canceló") &&
        valorColumnaM.includes("faltan")
      ) {
        // Separar el texto en partes
        var partes = separarTexto(valorColumnaM);

        var parte1 = partes[0]; // "El tesista canceló Bs."
        var ingresoPasado = parseFloat(partes[1]); // 1500
        var parte3 = partes[2]; // ", faltan Bs."
        var parte4 = partes[3]; // ","

        var nuevaDiferencia =
          obtenerValorCuota(nroCuota, fila, plazosPagosCbba) - ingresoPrimero;
        // Reconstruir el texto actualizado
        var textoActualizado =
          parte1 + ingresoSegundo + parte3 + nuevaDiferencia + parte4;
        // Borrar contenido de la celda antes de actualizar
        columnaM.clearContent();
        // Actualizar la columna M con el texto actualizado
        columnaM.setValue(textoActualizado);
      } else {
        // Obtener el texto existente en la columna M
        var textoExistente = valorColumnaM;

        // Generar el nuevo texto
        var cuotaValue = obtenerValorCuota(nroCuota, fila, plazosPagosCbba);
        var diferencia = cuotaValue - ingresoSegundo;
        var nuevoTexto =
          "canceló Bs." + ingresoSegundo + ", faltan Bs." + diferencia + ",";

        // Concatenar el texto existente al final del nuevo texto
        var textoFinal = nuevoTexto + " " + textoExistente;

        // Borrar contenido de la celda antes de actualizar
        columnaM.clearContent();
        // Actualizar la columna M con el texto final
        columnaM.setValue(textoFinal);
      }
    }
  }
  if (nroCuota == "3ra cuota") {
    if (valorColumnaM === "") {
      var cuotaValue = obtenerValorCuota(nroCuota, fila, plazosPagosCbba);
      var diferencia = cuotaValue - ingresoTercero;
      columnaM.setValue(
        "canceló Bs." + ingresoTercero + ", faltan Bs." + diferencia + ","
      );
      columnaM.setBackground("#ffff00");
    } else {
      if (
        valorColumnaM.includes("canceló") &&
        valorColumnaM.includes("faltan")
      ) {
        // Separar el texto en partes
        var partes = separarTexto(valorColumnaM);

        var parte1 = partes[0]; // "El tesista canceló Bs."
        var ingresoPasado = parseFloat(partes[1]); // 1500
        var parte3 = partes[2]; // ", faltan Bs."
        var parte4 = partes[3]; // ","

        var nuevaDiferencia =
          obtenerValorCuota(nroCuota, fila, plazosPagosCbba) - ingresoTercero;
        // Reconstruir el texto actualizado
        var textoActualizado =
          parte1 + ingresoTercero + parte3 + nuevaDiferencia + parte4;
        // Borrar contenido de la celda antes de actualizar
        columnaM.clearContent();
        // Actualizar la columna M con el texto actualizado
        columnaM.setValue(textoActualizado);
      } else {
        // Obtener el texto existente en la columna M
        var textoExistente = valorColumnaM;

        // Generar el nuevo texto
        var cuotaValue = obtenerValorCuota(nroCuota, fila, plazosPagosCbba);
        var diferencia = cuotaValue - ingresoPrimero;
        var nuevoTexto =
          "canceló Bs." + ingresoPrimero + ", faltan Bs." + diferencia + ",";

        // Concatenar el texto existente al final del nuevo texto
        var textoFinal = nuevoTexto + " " + textoExistente;

        // Borrar contenido de la celda antes de actualizar
        columnaM.clearContent();
        // Actualizar la columna M con el texto final
        columnaM.setValue(textoFinal);
      }
    }
  }
}

function obtenerValorCuota(nroCuota, fila, plazosPagosCbba) {
  switch (nroCuota) {
    case "1ra cuota":
      return plazosPagosCbba.getRange("G" + fila).getValue();
    case "2da cuota":
      return plazosPagosCbba.getRange("I" + fila).getValue();
    case "3ra cuota":
      return plazosPagosCbba.getRange("K" + fila).getValue();
    default:
      return 0; // Valor por defecto o manejo de error
  }
}

function separarTexto(texto) {
  // Encontrar las posiciones de los puntos y comas
  var punto1 = texto.indexOf(".");
  var punto2 = texto.indexOf(".", punto1 + 1);
  var coma1 = texto.indexOf(",");
  var coma2 = texto.indexOf(",", coma1 + 1);

  // Extraer partes del texto
  var parte1 = texto.substring(0, punto1 + 1);
  var parte2 = texto.substring(punto1 + 1, coma1);
  var parte3 = texto.substring(coma1, punto2 + 1);
  var parte4 = texto.substring(coma2);

  return [parte1, parte2, parte3, parte4];
}

// Ejemplo de uso
function test() {
  actualizarDatos(803, "SPACBBOL 162");
}

function iter() {
  var resultado = Iteracion("SPACBBOL 196");
  Logger.log(resultado);
}
