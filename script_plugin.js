function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Configurações do placeholder e células de controle
  const placeholderText = "Digite sua anotação aqui...";
  const mainCell = sheet.getRange("A1:B1");
  
  // Verifica se a célula editada é a célula A1 com o placeholder
  if (range.getA1Notation() === "A1") {
    if (range.getValue() !== placeholderText && range.getValue() !== "") {
      const lastRow = sheet.getLastRow() + 1;
      sheet.getRange(lastRow, 2).setValue(range.getValue());  // Coluna B
      sheet.getRange(lastRow, 1).insertCheckboxes();          // Coluna A
      sheet.getRange(lastRow, 3).setValue(new Date());        // Coluna C
      
      // Limpa o valor da célula A1 e restaura o placeholder
      mainCell.setValue(placeholderText);
      mainCell.setFontStyle("italic").setFontColor("#888888");
    } else if (range.getValue() === "") {
      // Restaura o placeholder se a célula A1 estiver em branco
      mainCell.setValue(placeholderText);
      mainCell.setFontStyle("italic").setFontColor("#888888");
    }
  }

  // Regras de timestamp e manipulação de checkbox
  if (range.getColumn() === 2 && range.getRow() > 1) {
    // Insere um checkbox automaticamente se a coluna B foi preenchida
    const checkboxCell = sheet.getRange(range.getRow(), 1);
    if (!checkboxCell.isChecked() && !checkboxCell.getValue()) {
      checkboxCell.insertCheckboxes();
    }
    sheet.getRange(range.getRow(), 3).setValue(new Date());  // Atualiza timestamp na coluna C
  } else if (range.getColumn() === 1 && range.getRow() > 1 && range.isChecked()) {
    sheet.getRange(range.getRow(), 4).setValue(new Date());  // Timestamp na coluna D
    
    // Calcula a diferença entre timestamps nas colunas D e C
    const creationTimestamp = sheet.getRange(range.getRow(), 3).getValue();
    const completionTimestamp = sheet.getRange(range.getRow(), 4).getValue();
    if (creationTimestamp) {
      const duration = (completionTimestamp - creationTimestamp) / (1000 * 60 * 60);
      sheet.getRange(range.getRow(), 5).setValue(duration.toFixed(2) + " horas");
    }
  }

  // Aplica as regras de ordenação após cada edição
  sortSheet(sheet);
}

// Função para ordenar a planilha conforme as regras
function sortSheet(sheet) {
  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5); // Linhas de dados
  const data = dataRange.getValues();

  // Regras de ordenação
  data.sort((a, b) => {
    // Regra 1: Checkbox desmarcado e coluna B com valor
    if (!a[0] && a[1] && b[0] && b[1]) return -1;
    if (!b[0] && b[1] && a[0] && a[1]) return 1;
    
    // Regra 2: Checkbox marcado e coluna B com valor
    if (a[0] && a[1] && b[0] && !b[1]) return -1;
    if (b[0] && b[1] && a[0] && !a[1]) return 1;
    
    // Regra 3: Checkbox marcado e coluna B vazia
    if (a[0] && !a[1] && !b[0] && !b[1]) return -1;
    if (!a[0] && !a[1] && b[0] && !b[1]) return 1;
    
    // Regra 4: Checkbox desmarcado e coluna B vazia
    if (!a[0] && !a[1] && b[0] && b[1]) return -1;
    if (a[0] && b[0] && !b[1]) return 1;
    
    // Regra 5: Ordenar as linhas desmarcadas pela data mais recente na coluna C
    if (!a[0] && a[1] && b[0] && !b[1]) {
      if (a[2] > b[2]) return -1;
      if (a[2] < b[2]) return 1;
    }
    
    // Regra 6: Ordenar as linhas marcadas pela data mais recente na coluna C
    if (a[0] && b[0]) {
      if (a[2] > b[2]) return -1;
      if (a[2] < b[2]) return 1;
    }

    return 0;
  });

  dataRange.setValues(data); // Atualiza a planilha com a ordem classificada
}

// Função para configurar o placeholder inicial
function setInitialPlaceholder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const placeholderText = "Digite sua anotação aqui...";
  const mainCell = sheet.getRange("A1:B1");

  mainCell.merge();
  mainCell.setValue(placeholderText);
  mainCell.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  mainCell.setFontStyle("italic").setFontColor("#888888").setHorizontalAlignment("center");
}
