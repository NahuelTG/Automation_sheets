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
];
function onEdit(event) {
  try {
    var filaEditada = event.range.getRow();
    var nombreCliente = libroCbba.getRange("D" + filaEditada).getValue();
    var concepto = libroCbba.getRange("E" + filaEditada).getValue();
    var codigo = libroCbba.getRange("B" + filaEditada).getValue();
    var ingreso = libroCbba.getRange("G" + filaEditada).getValue();
    var tipoCuota = determinarCuotas(concepto, categoriaA, categoriaB);
    var nombreContrato =
      concepto.split(",").length > 1
        ? concepto.split(",")[1].trim().toUpperCase()
        : "";

    if (filaEditada < 2) {
      // No hacer nada si se edita en la cabecera
      return;
    }

    if (codigo && codigo.match(/^SPACBBOL \d+$/)) {
      var datosOtro = plazosPagosCbba
        .getRange("A3:A" + plazosPagosCbba.getLastRow())
        .getValues();

      for (var j = 0; j < datosOtro.length; j++) {
        if (datosOtro[j][0] == codigo) {
          if (concepto.includes("1ra cuota")) {
            plazosPagosCbba.getRange("B" + (j + 3)).setValue(nombreContrato);
            plazosPagosCbba.getRange("C" + (j + 3)).setValue(nombreCliente);
            plazosPagosCbba.getRange("F" + (j + 3)).setValue("PRIMERA CUOTA");
            plazosPagosCbba.getRange("G" + (j + 3)).setValue(tipoCuota.primera);
            plazosPagosCbba.getRange("H" + (j + 3)).setValue("SEGUNDA CUOTA");
            plazosPagosCbba.getRange("I" + (j + 3)).setValue(tipoCuota.segunda);
            plazosPagosCbba.getRange("J" + (j + 3)).setValue("ÚLTIMA CUOTA");
            plazosPagosCbba.getRange("K" + (j + 3)).setValue(tipoCuota.ultima);
          }

          verificarCuotas(
            j + 3,
            plazosPagosCbba,
            ingreso,
            concepto,
            verificarSubCuota(concepto)
          );
          break;
        }
      }
    }
  } catch (e) {
    Logger.log(e.toString());
  }
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
  if (subCuota) {
    return "";
  } else {
    if (concepto.includes("1ra cuota")) {
      var valorColumnaG = plazosPagosCbba.getRange("G" + fila).getValue();
      if (valorColumnaG == ingreso) {
        // Pintar celdas de verde en el otro sheet
        plazosPagosCbba.getRange("F" + fila).setBackground("#7cd455");
        plazosPagosCbba.getRange("G" + fila).setBackground("#7cd455");
      }
    }
    if (concepto.includes("2da cuota")) {
      var valorColumnaG = plazosPagosCbba.getRange("I" + fila).getValue();
      if (valorColumnaG == ingreso) {
        // Pintar celdas de verde en el otro sheet
        plazosPagosCbba.getRange("H" + fila).setBackground("#7cd455");
        plazosPagosCbba.getRange("I" + fila).setBackground("#7cd455");
      }
    }
    if (concepto.includes("3ra cuota")) {
      var valorColumnaG = plazosPagosCbba.getRange("K" + fila).getValue();
      if (valorColumnaG == ingreso) {
        // Pintar celdas de verde en el otro sheet
        plazosPagosCbba.getRange("J" + fila).setBackground("#7cd455");
        plazosPagosCbba.getRange("K" + fila).setBackground("#7cd455");
      }
    }
  }
}

// Ejemplo de uso
function test() {
  actualizarDatos(781, "SPACBBOL 196");
}
