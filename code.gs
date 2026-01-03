function doGet(e) {
  // Recebe os parâmetros da URL
  const op = e.parameter.op;
  
  // Encaminha para a função correta e retorna JSON
  if (op === 'getDados') {
    return outputJSON(getDados());
  } else if (op === 'alternarStatus') {
    const tarefa = e.parameter.tarefa;
    const dia = e.parameter.dia;
    return outputJSON(alternarStatus(tarefa, dia));
  } else if (op === 'resetarSemana') {
    return outputJSON(resetarSemana());
  }

  // Resposta padrão se aceder ao link sem parâmetros
  return ContentService.createTextOutput("API do ERP Doméstico Online está ativa.");
}

// Função auxiliar para formatar a resposta como JSON
function outputJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function conectarPlanilha() {
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
        // Converte diversos formatos de "verdadeiro" para booleano real
        let estaFeito = (status === true || String(status).toLowerCase() === "true" || status === "VERDADEIRO");
        
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
    // Compara dia (coluna 0) e tarefa (coluna 1)
    if (data[i][0] == diaNome && data[i][1] == tarefaNome) {
      let valorAtual = data[i][2];
      let novoStatus = !(valorAtual === true || String(valorAtual).toLowerCase() === "true" || valorAtual === "VERDADEIRO");
      
      // Atualiza na planilha
      sheet.getRange(i + 1, 3).setValue(novoStatus);
      return { status: "sucesso", novoValor: novoStatus };
    }
  }
  return { status: "erro", mensagem: "Tarefa não encontrada" };
}

function resetarSemana() {
  const ss = conectarPlanilha();
  const sheet = ss.getSheetByName('Organizacao');
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    // Define toda a coluna C (Status) como false
    sheet.getRange(2, 3, lastRow - 1).setValue(false);
  }
  return { status: "sucesso" };
}
