function corrigirNomesSabado() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var contador = 0;
  
  // Itera sobre todas as guias
  sheets.forEach(function(sheet) {
    var nomeAtual = sheet.getName();
    
    // Verifica se o nome contém "SABADO" (case sensitive, maiúsculo conforme seu padrão)
    if (nomeAtual.indexOf("SABADO") !== -1) {
      
      // Cria o novo nome substituindo apenas a palavra errada
      var novoNome = nomeAtual.replace("SABADO", "SÁBADO");
      
      // Verifica se já não existe uma aba com o nome correto para evitar erro
      if (!ss.getSheetByName(novoNome)) {
        try {
          sheet.setName(novoNome);
          contador++;
          Logger.log("Renomeado: " + nomeAtual + " -> " + novoNome);
        } catch (e) {
          Logger.log("Erro ao tentar renomear '" + nomeAtual + "': " + e.message);
        }
      } else {
        Logger.log("Pulei '" + nomeAtual + "' pois '" + novoNome + "' já existe.");
      }
    }
  });
  
  SpreadsheetApp.getUi().alert("✅ Processo finalizado!\n\n" + contador + " guias foram corrigidas de SABADO para SÁBADO.");
}
