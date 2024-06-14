function ActualizarVigenciasTesistasDesplegableOtro() {
  // Obteniendo la hoja "Tesistas" donde se aloja el script
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tesistas");
  const data = sheet.getDataRange().getValues();

  // Fecha actual
  const currentDate = new Date();

  // Recorriendo los datos de la hoja "Tesistas" a partir de la fila 3
  for (let i = 2; i < data.length; i++) {
    // Comenzar en 2 para omitir los encabezados
    let dateCell = data[i][3]; // Columna "D" (índice 3) - Fecha de vencimiento

    if (dateCell) {
      let expiryDate = new Date(dateCell);
      let timeDiff = expiryDate - currentDate;
      let diffInDays = timeDiff / (1000 * 60 * 60 * 24);

      // Condiciones de marcado basadas en la fecha de vencimiento
      if (diffInDays < 0) {
        sheet.getRange(i + 1, 4).setBackground("red"); // Pintar de rojo si la fecha ya pasó
      } else if (diffInDays <= 14) {
        sheet.getRange(i + 1, 4).setBackground("orange"); // Pintar de naranja si la fecha está a 14 días o menos
      } else {
        // Limpiar el color de fondo si ninguna condición se cumple
        sheet.getRange(i + 1, 4).setBackground("green");
      }
    }
  }
}
