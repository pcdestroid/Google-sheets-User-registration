function indiceColuna(x,y,z){
  //exemplo: indiceColuna("texto a procurar","na linha","na Planilha")
  let index = z.getDataRange().getValues()[y - 1].indexOf(x);
  return index+1
}

function pegarUsuariosTransporte() {

  let todosUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuario');

  let usuariosTransporte = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('users_form_transporte');

  var linhaInicial = 2

  var coluna = indiceColuna('Solicitante_transporte',linhaInicial,todosUsuarios)
  
  usuariosTransporte.getRangeList(['A2:A1000']).clear({ contentsOnly: true, skipFilteredRows: true });//Limpar lista
  
  for (let x = linhaInicial; x <= todosUsuarios.getLastRow(); x = x + 1) {
    Logger.log(todosUsuarios.getRange(x,indiceColuna('Usuario',linhaInicial,todosUsuarios)).getValue())
    if (b = todosUsuarios.getRange(x,coluna).getValue() === 'sim')

    usuariosTransporte.getRange(usuariosTransporte.getLastRow()+1,indiceColuna('usuario',1,usuariosTransporte)).setValue(todosUsuarios.getRange(x,indiceColuna('Usuario',linhaInicial,todosUsuarios)).getValue())
    
    }

}
