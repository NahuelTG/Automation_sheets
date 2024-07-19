function ActualizarCumpleaños() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const currentDate = new Date();

  // Función para convertir fecha en texto a objeto Date
  function parseDate(dateStr) {
    if (!dateStr || typeof dateStr !== "string") return null; // Retorna null si la cadena está vacía, indefinida o no es una cadena
    const parts = dateStr.split("/");
    if (parts.length !== 3) return null; // Retorna null si la cadena no tiene el formato esperado
    const day = parseInt(parts[0]);
    const month = parseInt(parts[1]) - 1; // Mes en JavaScript es 0-indexado
    const year = parseInt(parts[2]);
    return new Date(year, month, day);
  }

  for (let i = 2; i < data.length; i++) {
    let dateCell = data[i][3]; // Columna "D" (índice 3) - Fecha de nacimiento
    if (dateCell) {
      let birthDate = parseDate(dateCell);
      if (birthDate) {
        let birthDay = birthDate.getDate();
        let birthMonth = birthDate.getMonth();
        let currentYear = currentDate.getFullYear();
        let nextBirthday = new Date(currentYear, birthMonth, birthDay);

        // Si el cumpleaños ya pasó este año, considera el próximo año
        if (nextBirthday < currentDate) {
          nextBirthday.setFullYear(currentYear + 1);
        }

        let timeDiff = nextBirthday - currentDate;
        let diffInDays = timeDiff / (1000 * 60 * 60 * 24);

        if (diffInDays < 0) {
          sheet.getRange(i + 1, 4).setBackground("red"); // Pintar de rojo si la fecha ya pasó
        } else if (diffInDays <= 20) {
          sheet.getRange(i + 1, 4).setBackground("orange"); // Pintar de naranja si la fecha está a 20 días o menos
        } else {
          sheet.getRange(i + 1, 4).setBackground(null); // Limpiar el color de fondo si ninguna condición se cumple
        }
      } else {
        sheet.getRange(i + 1, 4).setBackground(null); // Limpiar el color de fondo si la fecha no es válida
      }
    }
  }
}
