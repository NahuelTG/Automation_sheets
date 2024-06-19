function LowerCases() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2023");
  const data = sheet.getDataRange().getValues();

  // Iterar sobre las filas del data
  for (let i = 0; i < data.length; i++) {
    // Obtener el texto de la columna E (índice 4)
    let text = data[i][4];

    // Comprobar si el texto no es nulo o vacío
    if (text) {
      // Dividir el texto por la primera aparición de "-"
      let hyphenIndex = text.indexOf("-");
      let firstPart, remainingPart;

      if (hyphenIndex !== -1) {
        firstPart = text.substring(0, hyphenIndex);
        remainingPart = text.substring(hyphenIndex).toUpperCase();
      } else {
        firstPart = text;
        remainingPart = "";
      }

      // Dividir la parte antes del guion por la primera aparición de ","
      let commaIndex = firstPart.indexOf(",");
      if (commaIndex !== -1) {
        let beforeComma = firstPart.substring(0, commaIndex + 1);
        let afterComma = firstPart.substring(commaIndex + 1).trim();

        // Procesar la parte antes de la coma: primera letra en mayúscula, el resto se mantiene igual
        if (beforeComma) {
          beforeComma =
            beforeComma.charAt(0).toUpperCase() + beforeComma.slice(1);
        }

        // Procesar la parte después de la coma: primera letra en mayúscula, el resto se mantiene igual
        if (afterComma) {
          afterComma = afterComma.charAt(0).toUpperCase() + afterComma.slice(1);
        }

        // Reunir las partes
        firstPart = beforeComma + " " + afterComma;
      } else {
        // Procesar la primera parte si no hay coma: primera letra en mayúscula, el resto se mantiene igual
        firstPart = firstPart.charAt(0).toUpperCase() + firstPart.slice(1);
      }

      // Reunir las partes
      data[i][4] = firstPart + remainingPart;
    }
  }

  // Establecer los valores procesados de nuevo en la hoja
  sheet.getDataRange().setValues(data);
}
