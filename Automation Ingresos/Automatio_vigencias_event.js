var libroCbba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2024"); // Reemplaza "2024" con el nombre correcto de tu hoja
var plazosPagosCbba = SpreadsheetApp.openByUrl(
  "https://docs.google.com/spreadsheets/d/1XcqOdbIIb5FZQiAn8hOA6u7y1UPEaxsBY-PluzfgJYs/edit?usp=sharing"
).getSheetByName("TESISTAS");

// Listas de abreviaturas por categorías
var categoriaA = [
  "plásticas",
  "plasticas",
  "art.",
  "gráfico",
  "grafico",
  "musica",
  "musical",
  "musicales",
  "arqueología",
  "arqueologia",
  "antropologia",
  "antropología",
  "social",
  "comunicación",
  "comunicacion",
  "sociología",
  "sociologia",
  "derecho",
  "políticas",
  "politicas",
  "política",
  "politica",
  "internacionales",
  "información",
  "informacion",
  "educación",
  "educacion",
  "filosofía",
  "filosofia",
  "historia",
  "lingüística",
  "linguistica",
  "literatura",
  "psicología",
  "psicologia",
  "turismo",
];
var categoriaB = [
  "arquitectura",
  "veterinaria",
  "agronómica",
  "agronomica",
  "agronomia",
  "agronomía",
  "agropecuaria",
  "empresas",
  "empresa",
  "administración",
  "administracion",
  "contaduría",
  "contaduria",
  "economía",
  "economia",
  "marketing",
  "bioquímica",
  "bioquimica",
  "farmacéutica",
  "farmaceutica",
  "geográfica",
  "geografica",
  "geológica",
  "geologica",
  "medicina",
  "enfermería",
  "enfermeria",
  "nutrición",
  "nutricion",
  "médica",
  "medica",
  "medico",
  "odontología",
  "odontologia",
  "aeronáutica",
  "aeronautica",
  "civiles",
  "civil",
  "industrial",
  "telecomunicaciones",
  "electromecánica",
  "electromecanica",
  "automotriz",
  "topografía",
  "topografia",
  "biología",
  "biologia",
  "químicas",
  "quimicas",
  "quimica",
  "química",
  "estadística",
  "estadistica",
  "física",
  "fisica",
  "informática",
  "informatica",
  "matemáticas",
  "matematicas",
  "matemática",
  "matematica",
];
var ingresos = {
  primero: 0,
  segundo: 0,
  tercero: 0,
};
function onEdit(e) {
  try {
    const range = e.range;
    if (!libroCbba || range.getSheet().getName() !== "2024") {
      return;
    }

    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();

    for (let row = 0; row < numRows; row++) {
      const filaEditada = range.getRow() + row;

      if (filaEditada < 2) {
        // No hacer nada si se edita en la cabecera
        continue;
      }

      // Obtener los valores de las columnas B, D, E y G de la fila actual
      const codigo = libroCbba.getRange("B" + filaEditada).getValue();
      var nombreCliente = libroCbba.getRange("D" + filaEditada).getValue();
      var concepto = libroCbba.getRange("E" + filaEditada).getValue();
      var ingreso = libroCbba.getRange("G" + filaEditada).getValue();
      var tipoCuota = determinarCuotas(concepto, categoriaA, categoriaB);
      var nombreContrato = obtenerNombreContrato(concepto);

      // Verificar si las columnas B, D, E y G están llenas
      if (!codigo || !nombreCliente || !concepto || !ingreso) {
        continue; // Si alguna está vacía, saltar a la siguiente iteración
      }

      for (let col = 0; col < numCols; col++) {
        const cell = range.getCell(row + 1, col + 1);

        if (
          cell.getColumn() === 2 ||
          cell.getColumn() === 4 ||
          cell.getColumn() === 5 ||
          cell.getColumn() === 7
        ) {
          // Si se edita la columna B (2), D (4), E (5), verificar y ejecutar el código
          var cliente = obtenerCliente(codigo);
          actualizarIngresos(cliente);

          var filaPlazos = buscarFilaPlazos(codigo);
          if (filaPlazos > 0) {
            actualizarPlazos(
              filaPlazos,
              nombreContrato,
              nombreCliente,
              concepto,
              tipoCuota
            );
          }
        }
        if (verificarSubCuota(concepto)) {
          actualizarCuotasSubCuota(filaPlazos + 3, concepto);
        } else {
          actualizarCuotas(filaPlazos + 3, ingreso, concepto);
        }
      }
    }
  } catch (e) {
    Logger.log(e.toString());
  }
}

function obtenerNombreContrato(concepto) {
  var partes = concepto.split(",");
  if (partes.length > 1) {
    var nombreContrato = partes[1].trim();
    var indiceParentesis = nombreContrato.indexOf("(");
    if (indiceParentesis !== -1) {
      nombreContrato = nombreContrato.substring(0, indiceParentesis).trim();
    }
    return nombreContrato.toUpperCase();
  }
  return "";
}

function obtenerCliente(codigo) {
  var hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var clientes = [];

  hojas.forEach((hoja) => {
    var datos = hoja.getRange("B2:G" + hoja.getLastRow()).getValues();
    var clientesEnHoja = datos
      .filter((fila) => fila[0] === codigo)
      .map((fila) => ({
        concepto: fila[3],
        nombre: fila[2],
        ingreso: fila[5],
      }));
    clientes = clientes.concat(clientesEnHoja);
  });

  return clientes;
}

function actualizarIngresos(cliente) {
  ingresos = { primero: 0, segundo: 0, tercero: 0 };

  cliente.forEach((item) => {
    if (item.concepto.includes("1ra cuota") && verificarSubCuota(item.concepto))
      ingresos.primero += item.ingreso;
    if (item.concepto.includes("2da cuota") && verificarSubCuota(item.concepto))
      ingresos.segundo += item.ingreso;
    if (item.concepto.includes("3ra cuota") && verificarSubCuota(item.concepto))
      ingresos.tercero += item.ingreso;
  });
}

function buscarFilaPlazos(codigo) {
  var datos = plazosPagosCbba
    .getRange("A3:A" + plazosPagosCbba.getLastRow())
    .getValues();
  for (var j = 0; j < datos.length; j++) {
    if (datos[j][0] == codigo) return j;
  }
  return -1;
}

function actualizarPlazos(
  filaPlazos,
  nombreContrato,
  nombreCliente,
  concepto,
  tipoCuota
) {
  if (concepto.includes("1ra cuota")) {
    var celdaF = plazosPagosCbba.getRange("F" + (filaPlazos + 3));
    if (celdaF.getValue() === "") {
      plazosPagosCbba.getRange("B" + (filaPlazos + 3)).setValue(nombreContrato);
      plazosPagosCbba.getRange("C" + (filaPlazos + 3)).setValue(nombreCliente);
      celdaF.setValue("PRIMERA CUOTA");
      plazosPagosCbba
        .getRange("G" + (filaPlazos + 3))
        .setValue(tipoCuota.primera);
      plazosPagosCbba
        .getRange("H" + (filaPlazos + 3))
        .setValue("SEGUNDA CUOTA");
      plazosPagosCbba
        .getRange("I" + (filaPlazos + 3))
        .setValue(tipoCuota.segunda);
      plazosPagosCbba.getRange("J" + (filaPlazos + 3)).setValue("ÚLTIMA CUOTA");
      if (!concepto.includes("descuento")) {
        plazosPagosCbba
          .getRange("K" + (filaPlazos + 3))
          .setValue(tipoCuota.ultima);
      }
    }
  }
}

function determinarCuotas(concepto, categoriaA, categoriaB) {
  var cuotaPrimera = 2000;
  var cuotaSegunda = 0;
  var cuotaUltima = 0;
  var conceptoMinusculas = concepto.toLowerCase();
  var partesConcepto = concepto.split(",");
  var carreraEncontrada = encontrarCarrera(
    conceptoMinusculas,
    categoriaA,
    categoriaB
  );
  var esMaestria = partesConcepto.some((parte) =>
    parte.trim().includes("mae.")
  );

  if (carreraEncontrada) {
    if (esMaestria) {
      cuotaSegunda = categoriaA.includes(carreraEncontrada) ? 2500 : 2600;
      cuotaUltima = categoriaA.includes(carreraEncontrada) ? 3000 : 3400;
    } else {
      cuotaSegunda = categoriaA.includes(carreraEncontrada) ? 2400 : 2500;
      cuotaUltima = categoriaA.includes(carreraEncontrada) ? 2600 : 3000;
    }
  }

  return { primera: cuotaPrimera, segunda: cuotaSegunda, ultima: cuotaUltima };
}

function encontrarCarrera(concepto, categoriaA, categoriaB) {
  return (
    categoriaA.find((cat) => concepto.includes(cat)) ||
    categoriaB.find((cat) => concepto.includes(cat)) ||
    null
  );
}

function verificarSubCuota(concepto) {
  var match = concepto.match(/(\w),/);
  if (match) {
    var tipoConcepto = match[1];
    return tipoConcepto.charCodeAt(0) >= 65 && tipoConcepto.charCodeAt(0) <= 90;
  }
  return false;
}

function actualizarCuotas(fila, ingreso, concepto) {
  if (concepto.includes("1ra cuota")) actualizarCelda(fila, "F", "G", ingreso);
  if (concepto.includes("2da cuota")) actualizarCelda(fila, "H", "I", ingreso);
  if (concepto.includes("3ra cuota")) actualizarCelda(fila, "J", "K", ingreso);
}

function actualizarCelda(fila, colConcepto, colValor, ingreso) {
  var valorCelda = plazosPagosCbba.getRange(colValor + fila).getValue();
  if (valorCelda == ingreso) {
    plazosPagosCbba.getRange(colConcepto + fila).setBackground("#7cd455");
    plazosPagosCbba.getRange(colValor + fila).setBackground("#7cd455");
  }
}

function actualizarCuotasSubCuota(fila, concepto) {
  if (concepto.includes("1ra cuota"))
    actualizarSubCuota(fila, "1ra cuota", ingresos.primero);
  if (concepto.includes("2da cuota"))
    actualizarSubCuota(fila, "2da cuota", ingresos.segundo);
  if (concepto.includes("3ra cuota"))
    actualizarSubCuota(fila, "3ra cuota", ingresos.tercero);
}

function actualizarSubCuota(fila, nroCuota, ingresoTotal) {
  verificarYActualizarM(fila, nroCuota, ingresoTotal);
  var colConcepto, colValor;
  if (nroCuota === "1ra cuota") [colConcepto, colValor] = ["F", "G"];
  if (nroCuota === "2da cuota") [colConcepto, colValor] = ["H", "I"];
  if (nroCuota === "3ra cuota") [colConcepto, colValor] = ["J", "K"];

  var valorCelda = plazosPagosCbba.getRange(colValor + fila).getValue();
  if (valorCelda == ingresoTotal) {
    plazosPagosCbba.getRange(colConcepto + fila).setBackground("#7cd455");
    plazosPagosCbba.getRange(colValor + fila).setBackground("#7cd455");
    dejarObservaciones(fila);
  } else {
    plazosPagosCbba.getRange(colConcepto + fila).setBackground("#ffff00");
    plazosPagosCbba.getRange(colValor + fila).setBackground("#ffff00");
  }
}

function verificarYActualizarM(fila, nroCuota, ingresoTotal) {
  var columnaM = plazosPagosCbba.getRange("M" + fila);
  var valorColumnaM = columnaM.getValue().trim();
  var cuotaValue = obtenerValorCuota(nroCuota, fila);
  var diferencia = cuotaValue - ingresoTotal;
  var textoActualizado;

  if (valorColumnaM === "") {
    textoActualizado = `canceló Bs.${ingresoTotal}, faltan Bs.${diferencia},`;
  } else {
    var partes = separarTexto(valorColumnaM);
    if (partes[0] == "" && partes[1] == "" && partes[2] == "") {
      textoActualizado = `canceló Bs.${ingresoTotal}, faltan Bs.${diferencia},${partes[3]}`;
    } else {
      textoActualizado = `${partes[0]} ${ingresoTotal}${partes[2]}${diferencia}${partes[3]}`;
    }
  }

  columnaM.clearContent();
  columnaM.setValue(textoActualizado);
  columnaM.setBackground("#ffff00");
}

function obtenerValorCuota(nroCuota, fila) {
  switch (nroCuota) {
    case "1ra cuota":
      return plazosPagosCbba.getRange("G" + fila).getValue();
    case "2da cuota":
      return plazosPagosCbba.getRange("I" + fila).getValue();
    case "3ra cuota":
      return plazosPagosCbba.getRange("K" + fila).getValue();
    default:
      return 0;
  }
}

function dejarObservaciones(fila) {
  var celda = plazosPagosCbba.getRange("M" + fila);
  var textoOriginal = celda.getValue();
  var indiceCancelo = textoOriginal.indexOf("canceló");
  var indiceUltimaComa = textoOriginal.lastIndexOf(",");

  if (indiceCancelo !== -1 && indiceUltimaComa !== -1) {
    var textoAntesCancelo = textoOriginal.substring(0, indiceCancelo).trim();
    var textoDespuesUltimaComa = textoOriginal
      .substring(indiceUltimaComa + 1)
      .trim();
    var nuevoTexto = textoAntesCancelo + " " + textoDespuesUltimaComa;
    celda.setValue(nuevoTexto);
  }
}

function separarTexto(texto) {
  var punto1 = texto.indexOf(".");
  var punto2 = texto.indexOf(".", punto1 + 1);
  var coma1 = texto.indexOf(",");
  var coma2 = texto.indexOf(",", coma1 + 1);

  var parte1 = texto.substring(0, punto1 + 1);
  var parte2 = texto.substring(punto1 + 1, coma1);
  var parte3 = texto.substring(coma1, punto2 + 1);
  var parte4 = texto.substring(coma2);

  return [parte1, parte2, parte3, parte4];
}
