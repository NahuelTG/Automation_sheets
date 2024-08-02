function ActualizarVigencias() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SUELTOS");
  const data = sheet.getDataRange().getValues();
  const currentDate = new Date();

  // Función para convertir fecha en texto a objeto Date (formato DD/MM/YYYY)
  function parseDate(dateStr) {
    if (!dateStr) return null; // Retorna null si la cadena está vacía o indefinida

    // Si dateStr es un objeto Date, retornarlo directamente
    if (dateStr instanceof Date) return dateStr;

    // Si dateStr es una cadena de texto en formato "Thu Dec 28 2023 00:00:00 GMT-0400 (Bolivia Time)"
    if (!isNaN(Date.parse(dateStr))) return new Date(dateStr);

    // Si dateStr es una cadena de texto en formato DD/MM/YYYY
    if (typeof dateStr === "string") {
      const parts = dateStr.split("/");
      if (parts.length !== 3) return null; // Retorna null si la cadena no tiene el formato esperado
      const day = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1; // Los meses en JavaScript son base 0
      const year = parseInt(parts[2]);
      return new Date(year, month, day);
    }

    return null; // Retorna null si el formato no es reconocido
  }

  // Función para actualizar el color de la celda según la fecha
  function updateCellColor(rowIndex, colIndex) {
    let dateCell = data[rowIndex][colIndex];
    if (dateCell) {
      let expiryDate = parseDate(dateCell);
      if (expiryDate) {
        let timeDiff = expiryDate - currentDate;
        let diffInDays = timeDiff / (1000 * 60 * 60 * 24);

        if (diffInDays < 0) {
          sheet.getRange(rowIndex + 1, colIndex + 1).setBackground("#666666"); // Plomo oscuro si la fecha ya pasó
        } else if (diffInDays <= 1) {
          sheet.getRange(rowIndex + 1, colIndex + 1).setBackground("red"); // Rojo si la fecha está a 1 día o menos
        } else if (diffInDays <= 3) {
          sheet.getRange(rowIndex + 1, colIndex + 1).setBackground("orange"); // Naranja si la fecha está a 3 días o menos
        } else {
          sheet.getRange(rowIndex + 1, colIndex + 1).setBackground(null); // Limpiar el color de fondo si ninguna condición se cumple
        }
      } else {
        sheet.getRange(rowIndex + 1, colIndex + 1).setBackground(null); // Limpiar el color de fondo si la fecha no es válida
      }
    }
  }

  for (let i = 1; i < data.length; i++) {
    // Empezar desde la fila 2 (índice 1)
    updateCellColor(i, 12); // Columna "M" (índice 12)
    updateCellColor(i, 13); // Columna "N" (índice 13)
  }
}
