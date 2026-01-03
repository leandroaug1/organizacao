function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Gestão Doméstica ERP')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function conectarPlanilha() {
  // ✅ AQUI ESTÁ A CORREÇÃO: O ID DA SUA PLANILHA JÁ ESTÁ INSERIDO
  const idPlanilha = "1GdTjx0Yw7ezee3HW4_u5Oner_jnAph2SOiP89hyI8sE"; 
  return SpreadsheetApp.openById(idPlanilha);
}

function getDados() {
  const ss = conectarPlanilha();
  const sheet = ss.getSheetByName('Organizacao');
  
  if (!sheet) return {}; 

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {}; 
  
  // Lê 4 colunas: A=Dia, B=Tarefa, C=Status, D=Tempo
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  
  const dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"];
  let organizado = {};
  
  dias.forEach(dia => {
    organizado[dia] = data
      .filter(row => row[0] === dia)
      .map(row => {
        let status = row[2];
        let estaFeito = (status === true || status === "TRUE" || status === "true" || status === "VERDADEIRO");
        
        // Processamento do Timer (Coluna D)
        let temposRaw = row[3]; 
        let tempos = [];
        
        if (temposRaw) {
           tempos = temposRaw.toString().split(',');
        }

        return { 
          tarefa: row[1], 
          feito: estaFeito, 
          dia: row[0], 
          tempos: tempos 
        };
      });
  });
  
  return organizado;
}

function alternarStatus(tarefaNome, diaNome) {
  const ss = conectarPlanilha();
  const sheet = ss.getSheetByName('Organizacao');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == diaNome && data[i][1] == tarefaNome) {
      let valorAtual = data[i][2];
      let novoStatus = !(valorAtual === true || valorAtual === "TRUE" || valorAtual === "true" || valorAtual === "VERDADEIRO");
      
      sheet.getRange(i + 1, 3).setValue(novoStatus);
      return novoStatus;
    }
  }
}

function resetarSemana() {
  const ss = conectarPlanilha();
  const sheet = ss.getSheetByName('Organizacao');
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 3, lastRow - 1).setValue(false);
  }
}
