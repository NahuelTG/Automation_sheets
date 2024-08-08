function ActualizarVigencias() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TESISTAS");
  const data = sheet.getDataRange().getValues();
  const currentDate = new Date();

  // Función para convertir fecha en texto a objeto Date
  function parseDate(dateStr) {
    if (!dateStr) return null; // Retorna null si la cadena está vacía o indefinida
    const monthNames = [
      "enero",
      "febrero",
      "marzo",
      "abril",
      "mayo",
      "junio",
      "julio",
      "agosto",
      "septiembre",
      "octubre",
      "noviembre",
      "diciembre",
    ];
    const parts = dateStr.split(" ");
    if (parts.length !== 5) return null; // Retorna null si la cadena no tiene el formato esperado
    const day = parseInt(parts[0]);
    const month = parts[2].toLowerCase();
    const year = parseInt(parts[4]);
    const monthIndex = monthNames.indexOf(month);
    if (monthIndex === -1) return null; // Retorna null si el mes no es válido
    return new Date(year, monthIndex, day);
  }

  for (let i = 2; i < data.length; i++) {
    let dateCell = data[i][3]; // Columna "D" (índice 3) - Fecha de vencimiento
    if (dateCell) {
      let expiryDate = parseDate(dateCell);
      if (expiryDate) {
        let timeDiff = expiryDate - currentDate;
        let diffInDays = timeDiff / (1000 * 60 * 60 * 24);

        if (diffInDays < 0) {
          sheet.getRange(i + 1, 4).setBackground("#787878"); // Pintar de gris si la fecha ya pasó
        } else if (diffInDays <= 7) {
          sheet.getRange(i + 1, 4).setBackground("red"); // Pintar de rojo si la fecha está a 7 días o menos
        } else if (diffInDays <= 30) {
          sheet.getRange(i + 1, 4).setBackground("orange"); // Pintar de naranja si la fecha está a 30 días o menos
        } else if (diffInDays <= 60) {
          sheet.getRange(i + 1, 4).setBackground("#bd00ff"); // Pintar de lila si la fecha está a 60 días o menos
        } else {
          //Holi XD
        }
      } else {
        sheet.getRange(i + 1, 4).setBackground(null); // Limpiar el color de fondo si la fecha no es válida
      }
    }
  }
}
