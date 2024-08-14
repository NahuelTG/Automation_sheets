function onFormSubmit(e) {
  const sheetCasos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respuestas de formulario 1");
  const lastRow = sheetCasos.getLastRow();
  actualizarContratos(lastRow);
}

function actualizarContratos(row) {
  const sheetCasos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respuestas de formulario 1");
  const sheetContratos = SpreadsheetApp.openByUrl(
    "https://docs.google.com/spreadsheets/d/1vmbRCWVd1zaA0825-bVxwkCNfvM6JlITnfM5D8ztEZk/edit?usp=sharing"
  ).getSheetByName("INFORME DE CONTRATOS COMBOS");

  const dataCasos = sheetCasos.getRange(row, 1, 1, sheetCasos.getLastColumn()).getValues()[0];
  const dataContratos = sheetContratos.getDataRange().getValues();
  
  const mes = dataCasos[1];                  // Mes
  const caso = dataCasos[2];                 // Código del contrato
  const situacion = dataCasos[3];            // SITUACIÓN
  const estado = dataCasos[4];               // ESTADO DEL DOCUMENTO
  const comunicacion = dataCasos[5];         // COMUNICACIÓN ENTRE TESISTA Y REDACTORES
  const observaciones = dataCasos[6];        // OBSERVACIONES
  const recomendaciones = dataCasos[7];      // RECOMENDACIONES
  
  // Buscar y actualizar en la hoja de contratos
  for (let j = 1; j < dataContratos.length; j++) {
    const codigo = dataContratos[j][0];      // Código del contrato en la hoja de contratos
    
    if (codigo.includes(caso.toString())) {
      let mesColumna;
      switch(mes) {
        case "Junio":
          mesColumna = 4; // Columna D
          break;
        case "Julio":
          mesColumna = 5; // Columna E
          break;
        case "Agosto":
          mesColumna = 6; // Columna F
          break;
        default:
          continue;
      }
      
      const contenido = `SITUACIÓN O HISTORIA: ${situacion}\n` +
                        `ESTADO DEL DOCUMENTO: ${estado}\n` +
                        `COMUNICACIÓN ENTRE TESISTA Y REDACTORES: ${comunicacion}\n` +
                        `OBSERVACIONES: ${observaciones}\n` +
                        `RECOMENDACIONES: ${recomendaciones}`;
      
      sheetContratos.getRange(j + 1, mesColumna).setValue(contenido);
      break;
    }
  }
}
