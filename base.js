function criarAbaComBaseNoModeloECopiarDadosEmSequenciaComNomePersonalizado() {
    
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    
    
    var abaModelo = planilha.getSheetByName("Modelo");
    
    
    var abaOrigem = planilha.getSheetByName("Sheet");  
    
    
    if (!abaModelo || !abaOrigem) {
      SpreadsheetApp.getUi().alert('Verifique se as abas "Modelo" e "Sheet" (Origem) existem.');
      return;
    }
    
    
    var dadosOrigem = abaOrigem.getRange("A1:B44").getValues();  
    
    
    for (var i = 0; i < dadosOrigem.length; i++) {
      
      var novaAba = abaModelo.copyTo(planilha);
      
      
      var valorE3 = dadosOrigem[i][0];  
      var valorE4 = dadosOrigem[i][1];  
      
      
      novaAba.getRange("E3").setValue(valorE3);
      novaAba.getRange("E4").setValue(valorE4);
      
      
      try {
        novaAba.setName(valorE3);  
      } catch (e) {
        
        SpreadsheetApp.getUi().alert('Erro ao renomear aba: ' + valorE3 + '. Verifique se o nome já existe ou é inválido.');
        novaAba.setName('NovaAba' + (i + 1));  
      }
    }
    
    
    SpreadsheetApp.getUi().alert('Novas abas criadas e renomeadas com os valores de E3!');
  }
  