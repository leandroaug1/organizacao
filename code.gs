// Função principal que recebe os pedidos do site (GET)
function doGet(e) {
  const op = e.parameter.op;
  
  // Roteamento simples
  if (op === 'getDados') {
    return respostaJSON(getDados());
  } else if (op === 'alternarStatus') {
    const tarefa = e.parameter.tarefa;
    const dia = e.parameter.dia;
    return respostaJSON(alternarStatus(tarefa, dia));
  } else if (op === 'resetarSemana') {
    return respostaJSON(resetarSemana());
  }

  // Se não tiver parametros, retorna mensagem simples
  return ContentService.createTextOutput("API do ERP Doméstico Online");
}

// Função auxiliar para responder JSON corretamente (CORS)
function respostaJSON(dados) {
  return ContentService.createTextOutput(JSON.stringify(dados))
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

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"];
  let organizado = {};
  
  dias.forEach(dia => {
    organizado[dia] = data
      .filter(row => row[0] === dia)
      .map(row => {
        let status = row[2];
        let estaFeito = (status === true || String(status).toLowerCase() === "true" || status === "VERDADEIRO");
        let temposRaw = row[3]; 
        let tempos = temposRaw ? temposRaw.toString().split(',') : [];

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
  
  // Procura a linha correta para editar
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == diaNome && data[i][1] == tarefaNome) {
      let valorAtual = data[i][2];
      let novoStatus = !(valorAtual === true || String(valorAtual).toLowerCase() === "true" || valorAtual === "VERDADEIRO");
      sheet.getRange(i + 1, 3).setValue(novoStatus);
      return { status: "sucesso", novoValor: novoStatus };
    }
  }
  return { status: "erro", msg: "Tarefa não encontrada" };
}

function resetarSemana() {
  const ss = conectarPlanilha();
  const sheet = ss.getSheetByName('Organizacao');
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 3, lastRow - 1).setValue(false);
  }
  return { status: "resetado" };
}
