function ActualizarMayusculasAuto(e) {
  const sheet = e.source.getSheetByName("TESISTAS");
  const range = e.range;

  // Comprobar si la edición se realizó en la hoja "2024"
  if (sheet && range.getSheet().getName() === "TESISTAS") {
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();

    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numCols; col++) {
        let cell = range.getCell(row + 1, col + 1);
        // Comprobar si la edición se realizó en la columna B (índice 2)
        if (cell.getColumn() === 2 && cell.getRow() >= 2) {
          let text = cell.getValue();

          if (text) {
            // Convertir todo el texto a mayúsculas
            cell.setValue(text.toUpperCase());
          }
        }
      }
    }
  }
}
