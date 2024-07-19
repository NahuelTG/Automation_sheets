function ActualizarMinusculas(e) {
  const range = e.range;

  // Comprobar si la edición se realizó en la columna E (índice 5)
  if (range.getSheet().getName() === "2024" && range.getColumn() === 5) {
    let text = range.getValue();

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

        // Procesar la parte después de la coma: primera letra en mayúscula, el resto en minúsculas
        if (afterComma) {
          afterComma =
            afterComma.charAt(0).toUpperCase() +
            afterComma.slice(1).toLowerCase();
        }

        // Reunir las partes
        firstPart = beforeComma + " " + afterComma;
      } else {
        // Procesar la primera parte si no hay coma: primera letra en mayúscula, el resto se mantiene igual
        firstPart = firstPart.charAt(0).toUpperCase() + firstPart.slice(1);
      }

      // Reunir las partes
      text = firstPart + remainingPart;

      // Establecer el valor procesado de nuevo en la celda editada
      range.setValue(text);
    }
  }
}
function ActualizarMinusculasAuto(e) {
  const sheet = e.source.getSheetByName("2024");
  const range = e.range;

  // Comprobar si la edición se realizó en la hoja "2024"
  if (sheet && range.getSheet().getName() === "2024") {
    // Iterar sobre cada celda en el rango editado
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();

    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numCols; col++) {
        // Comprobar si la columna es la E (índice 5)
        if (range.getCell(row + 1, col + 1).getColumn() === 5) {
          let cell = range.getCell(row + 1, col + 1);
          let text = cell.getValue();

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

              // Procesar la parte después de la coma: primera letra en mayúscula, el resto en minúsculas
              if (afterComma) {
                afterComma =
                  afterComma.charAt(0).toUpperCase() +
                  afterComma.slice(1).toLowerCase();
              }

              // Reunir las partes
              firstPart = beforeComma + " " + afterComma;
            } else {
              // Procesar la primera parte si no hay coma: primera letra en mayúscula, el resto se mantiene igual
              firstPart =
                firstPart.charAt(0).toUpperCase() + firstPart.slice(1);
            }

            // Reunir las partes
            text = firstPart + remainingPart;

            // Establecer el valor procesado de nuevo en la celda editada
            cell.setValue(text);
          }
        }
      }
    }
  }
}

function myFunction() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2024");
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

        // Procesar la parte después de la coma: primera letra en mayúscula, el resto en minúsculas
        if (afterComma) {
          afterComma =
            afterComma.charAt(0).toUpperCase() +
            afterComma.slice(1).toLowerCase();
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
