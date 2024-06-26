function onEdit(event) {
  try {
    var filaEditada = event.range.getRow();

    if (filaEditada < 2) {
      // No hacer nada si se edita en la cabecera
      return;
    }

    var libroCbba =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2024"); // Reemplaza "2024" con el nombre correcto de tu hoja
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

    var nombreCliente = libroCbba.getRange("D" + filaEditada).getValue();
    var concepto = libroCbba.getRange("E" + filaEditada).getValue();
    var codigo = libroCbba.getRange("B" + filaEditada).getValue();
    var tipoCuota = determinarCuotas(concepto, categoriaA, categoriaB);
    var nombreContrato =
      concepto.split(",").length > 1
        ? partesConcepto[1].trim().toUpperCase()
        : "";

    if (codigo && codigo.match(/^SPACBBOL \d+$/)) {
      var datosOtro = plazosPagosCbba
        .getRange("A3:A" + plazosPagosCbba.getLastRow())
        .getValues();
      var encontrado = false;

      for (var j = 0; j < datosOtro.length; j++) {
        if (datosOtro[j][0] == codigo) {
          encontrado = true;
          plazosPagosCbba.getRange("B" + (j + 3)).setValue(nombreContrato);
          plazosPagosCbba.getRange("C" + (j + 3)).setValue(nombreCliente);

          if (concepto.includes("1ra cuota")) {
            plazosPagosCbba.getRange("F" + (j + 3)).setValue("PRIMERA CUOTA");
            plazosPagosCbba.getRange("G" + (j + 3)).setValue(tipoCuota.primera);
            plazosPagosCbba.getRange("H" + (j + 3)).setValue("SEGUNDA CUOTA");
            plazosPagosCbba.getRange("I" + (j + 3)).setValue(tipoCuota.segunda);
            plazosPagosCbba.getRange("J" + (j + 3)).setValue("ÚLTIMA CUOTA");
            plazosPagosCbba.getRange("K" + (j + 3)).setValue(tipoCuota.ultima);
          }

          verificarPago(j + 3, plazosPagosCbba);
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

function verificarPago(fila, plazosPagosCbba) {
  var concepto = plazosPagosCbba
    .getRange("E" + fila)
    .getValue()
    .toLowerCase();
  if (concepto.includes("1er pago")) {
    var valorColumnaG = plazosPagosCbba.getRange("G" + fila).getValue();
    var primeraCuota = determinarCuotas(
      concepto,
      categoriaA,
      categoriaB
    ).primera;
    if (valorColumnaG == primeraCuota) {
      // Pintar celdas de verde en el otro sheet
      plazosPagosCbba.getRange("F" + fila).setBackground("green");
      plazosPagosCbba.getRange("G" + fila).setBackground("green");
    }
  }
}
