function onEdit(event) {
  try {
    var filaEditada = event.range.getRow();

    if (filaEditada < 2) {
      // No hacer nada si se edita en la cabecera
      return;
    }

    var libroActual = SpreadsheetApp.getActiveSpreadsheet();
    var sheetActual = libroActual.getSheetByName("2024"); // Reemplaza "2024" con el nombre correcto de tu hoja

    var libroOtro = SpreadsheetApp.openByUrl(
      "https://docs.google.com/spreadsheets/d/1XcqOdbIIb5FZQiAn8hOA6u7y1UPEaxsBY-PluzfgJYs/edit?usp=sharing"
    );
    var sheetOtro = libroOtro.getSheetByName("TESISTAS");

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

    var codigo = sheetActual.getRange("B" + filaEditada).getValue();
    var nombreContrato = sheetActual.getRange("D" + filaEditada).getValue();
    var concepto = sheetActual
      .getRange("E" + filaEditada)
      .getValue()
      .toLowerCase();

    if (codigo && codigo.match(/^SPACBBOL \d+$/)) {
      var datosOtro = sheetOtro
        .getRange("A3:A" + sheetOtro.getLastRow())
        .getValues();
      var encontrado = false;

      for (var j = 0; j < datosOtro.length; j++) {
        if (datosOtro[j][0] == codigo) {
          encontrado = true;
          var partesConcepto = concepto.split(",");
          var nombreContratoTrimmed =
            partesConcepto.length > 1
              ? partesConcepto[1].trim().toUpperCase()
              : "";

          sheetOtro.getRange("B" + (j + 3)).setValue(nombreContratoTrimmed);
          sheetOtro.getRange("C" + (j + 3)).setValue(nombreContrato); // Convertir a mayúsculas

          if (concepto.includes("1ra cuota")) {
            var cuotas = determinarCuotas(concepto, categoriaA, categoriaB);
            sheetOtro.getRange("F" + (j + 3)).setValue("PRIMERA CUOTA");
            sheetOtro.getRange("G" + (j + 3)).setValue(cuotas.primera);
            sheetOtro.getRange("H" + (j + 3)).setValue("SEGUNDA CUOTA");
            sheetOtro.getRange("I" + (j + 3)).setValue(cuotas.segunda);
            sheetOtro.getRange("J" + (j + 3)).setValue("ÚLTIMA CUOTA");
            sheetOtro.getRange("K" + (j + 3)).setValue(cuotas.ultima);
          }
          break;
        }
      }

      if (!encontrado) {
        var ultimaFilaOtro = sheetOtro.getLastRow() + 1;

        var partesConcepto = concepto.split(",");
        var nombreContratoTrimmed =
          partesConcepto.length > 1
            ? partesConcepto[1].trim().toUpperCase()
            : "";

        sheetOtro
          .getRange("A" + ultimaFilaOtro)
          .setValue(codigo)
          .setFontWeight("bold");
        sheetOtro
          .getRange("B" + ultimaFilaOtro)
          .setValue(nombreContratoTrimmed);
        sheetOtro.getRange("C" + ultimaFilaOtro).setValue(nombreContrato); // Convertir a mayúsculas

        if (concepto.includes("1ra cuota")) {
          var cuotas = determinarCuotas(concepto, categoriaA, categoriaB);
          sheetOtro.getRange("F" + ultimaFilaOtro).setValue("PRIMERA CUOTA");
          sheetOtro.getRange("G" + ultimaFilaOtro).setValue(cuotas.primera);
          sheetOtro.getRange("H" + ultimaFilaOtro).setValue("SEGUNDA CUOTA");
          sheetOtro.getRange("I" + ultimaFilaOtro).setValue(cuotas.segunda);
          sheetOtro.getRange("J" + ultimaFilaOtro).setValue("ÚLTIMA CUOTA");
          sheetOtro.getRange("K" + ultimaFilaOtro).setValue(cuotas.ultima);
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

  // Determinar cuotas basado en la carrera encontrada
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
