function rastrearStatus() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow(); 
  
  Logger.log("Iniciando o processo de rastreamento...");

  for (var linha = 2; linha <= lastRow; linha++) {
    Logger.log("Processando linha " + linha + "...");

    var statusAtual = sheet.getRange("B" + linha).getValue(); 
    if (statusAtual === "ENTREGUE") {
      Logger.log("Linha " + linha + " já está com status ENTREGUE. Pulando...");
      continue; 
    }

    Logger.log("Iniciando rastreio para o conjunto 1 na linha " + linha);
    rastrearCodigo(sheet, linha, "A", "B", "C");
    
    Logger.log("Substituindo texto na linha " + linha + "...");
    substituirTexto(sheet, ["B"], linha); 

    Logger.log("Linha " + linha + " processada com sucesso.");
    Logger.log("Aguardando 5 segundos antes de prosseguir para a próxima linha...");
    Utilities.sleep(5000); 
  }

  Logger.log("Processo de rastreamento concluído.");
}

function rastrearCodigo(sheet, linha, colunaCodigo, colunaStatus, colunaDataEntrega) {
  var codigoRastreio = sheet.getRange(colunaCodigo + linha).getValue().toString().trim();
  var statusAtual = sheet.getRange(colunaStatus + linha).getValue();
  
  Logger.log("Verificando código de rastreio na linha " + linha + ": " + codigoRastreio);

  if (!codigoRastreio || codigoRastreio === "") {
    Logger.log("Código de rastreio vazio na linha " + linha + ". Pulando...");
    return;
  }

  if (statusAtual === "EXTRAVIO") {
    Logger.log("Código " + codigoRastreio + " já marcado como EXTRAVIO. Rastreamento interrompido.");
    return;
  }

  Logger.log("Rastreando objeto com código: " + codigoRastreio);
  var status = rastrearObjeto(codigoRastreio);

  if (status.status === "Status não encontrado" && statusAtual && statusAtual !== "Status não encontrado") {
    Logger.log("Status não encontrado, mas preservando o status anterior: " + statusAtual);
    return;
  }

  if (status.status === "Status: Objeto não localizado no fluxo postal") {
    Logger.log("Código " + codigoRastreio + " marcado como EXTRAVIO.");
    sheet.getRange(colunaStatus + linha).setValue("EXTRAVIO");
    sheet.getRange(colunaDataEntrega + linha).setValue("Não aplicável");
    return; 
  }

  Logger.log("Status encontrado: " + status.status);
  Logger.log("Data de entrega: " + status.dataEntrega);

  sheet.getRange(colunaStatus + linha).setValue(status.status);
  sheet.getRange(colunaDataEntrega + linha).setValue(status.dataEntrega);

  var dataHoraAtualizacao = new Date().toLocaleString();
  sheet.getRange(colunaStatus + linha).setNote("Última atualização: " + dataHoraAtualizacao);
}

function rastrearObjeto(codigoRastreio) {
  var url = "https://www.linkcorreios.com.br/?id=" + codigoRastreio;
  Logger.log("Fazendo requisição para: " + url);
  
  try {
    var response = UrlFetchApp.fetch(url);
    var html = response.getContentText();
    var $ = Cheerio.load(html);

    // Encontra a seção de status
    var statusSection = $('.linha_status').first(); 
    if (statusSection.length > 0) {
      var status = statusSection.find('li').eq(0).text().trim(); 
      var dataHora = statusSection.find('li').eq(2).text().trim(); 

      // Limpa o texto do status e da data
      status = status.replace(/^Status:\s*/, ""); 
      dataHora = dataHora.replace(/^Data:\s*/, ""); 

      // Verifica se o status indica entrega
      if (status.toLowerCase().includes("entregue")) {
        Logger.log("Objeto entregue. Data de entrega: " + dataHora.split(" | ")[0]);
        return {
          status: status,
          dataEntrega: dataHora.split(" | ")[0] // Retorna apenas a data
        };
      } else {
        Logger.log("Objeto ainda não entregue. Status: " + status);
        return {
          status: status,
          dataEntrega: "Aguardando entrega"
        };
      }
    } else {
      Logger.log("Status não encontrado para o código: " + codigoRastreio);
      return {
        status: "Status não encontrado",
        dataEntrega: "Aguardando entrega"
      };
    }
  } catch (e) {
    Logger.log("Erro ao rastrear objeto: " + e.toString());
    return {
      status: "Erro ao rastrear",
      dataEntrega: "Erro ao rastrear"
    };
  }
}

function substituirTexto(sheet, colunas, linha) {
  var substituicoes = {
    "Status: Objeto em transferência - por favor aguarde": "EM TRÂNSITO",
    "Status: Objeto entregue ao destinatário": "ENTREGUE",
    "Status: Objeto entregue ao remetente": "DEVOLUÇÃO",
    "Status: Solicitação de suspensão de entrega ao destinatário": "DEVOLUÇÃO",
    "Status: Objeto saiu para entrega ao destinatário": "EM TRÂNSITO",
    "Status: Objeto aguardando retirada no endereço indicado": "AGUARDANDO RETIRADA",
    "Status: Objeto ainda não chegou à unidade": "EM TRÂNSITO",
    "Status: Objeto devolvido aos Correios": "DEVOLUÇÃO",
    "Status: Objeto não entregue - cliente recusou-se a receber o objeto": "DEVOLUÇÃO",
    "Status: Objeto encaminhado para retirada no endereço indicado": "AGUARDANDO RETIRADA",
    "Objeto entregue ao destinatário":"ENTREGUE",
    "Objeto em transferência - por favor aguarde":"EM TRÂNSITO"
  };

  colunas.forEach(function(coluna) {
    var celula = sheet.getRange(coluna + linha);
    var valor = celula.getValue().toString().trim(); 

    
    if (substituicoes[valor]) {
      Logger.log("Substituindo '" + valor + "' por '" + substituicoes[valor] + "' na linha " + linha);
      celula.setValue(substituicoes[valor]); 
    } else {
      Logger.log("Nenhuma substituição encontrada para o valor: " + valor);
    }
  });
}

function configurarIntervalo() {
  ScriptApp.newTrigger('rastrearStatus')
    .timeBased()
    .everyMinutes(30)
    .create();
}
