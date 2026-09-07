/**
 * INVENTÁRIO DO GRUPO — porta de entrada das contagens.
 * A app envia para aqui cada contagem fechada e este script escreve-a
 * na folha "INVENTÁRIO — Mestre" (pasta 02 Contagens).
 *
 * Instalação: correr instalar() uma vez, autorizar, e depois
 * Implementar > Nova implementação > Aplicação Web
 * (Executar como: eu | Quem tem acesso: Qualquer pessoa).
 */
var XLSX_ID  = '1eMNYpi2nr2RQcceoA-9lpke50MyUzpJ8';   // INVENTARIO - Mestre.xlsx
var PASTA_ID = '1iyIUHBspvWLrW3EpmIiDgDx8Vz_YE3cY';   // pasta 02 Contagens
var NOME     = 'INVENTÁRIO — Mestre';
var CODIGO   = '2530';

function P_(){ return PropertiesService.getScriptProperties(); }
function J_(o){ return ContentService.createTextOutput(JSON.stringify(o))
  .setMimeType(ContentService.MimeType.JSON); }

function instalar(){
  var ss = folha_();
  Logger.log('Folha pronta: ' + ss.getUrl());
  return ss.getUrl();
}

function folha_(){
  var id = P_().getProperty('FOLHA_ID');
  if(id){ try{ return SpreadsheetApp.openById(id); }catch(e){} }
  var r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + XLSX_ID + '/copy', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
    payload: JSON.stringify({name: NOME, parents: [PASTA_ID],
      mimeType: 'application/vnd.google-apps.spreadsheet'})
  });
  var novo = JSON.parse(r.getContentText());
  if(!novo.id) throw new Error('Não deu para criar a folha: ' + r.getContentText());
  var ss = SpreadsheetApp.openById(novo.id);
  var lanc = ss.getSheetByName('Lançamentos');
  if(lanc && String(lanc.getRange(5,11).getValue()) === 'exemplo') lanc.deleteRow(5);
  P_().setProperty('FOLHA_ID', novo.id);
  return ss;
}

function doGet(e){
  var id = P_().getProperty('FOLHA_ID');
  return J_({ok:true, ligado:true, folha:id ? 'https://docs.google.com/spreadsheets/d/' + id : null});
}

function doPost(e){
  try{
    var d = JSON.parse(e.postData.contents);
    if(String(d.codigo) !== CODIGO) return J_({ok:false, erro:'codigo'});
    if(!d.linhas || !d.linhas.length) return J_({ok:false, erro:'contagem vazia'});
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try{ return J_(registar_(d)); } finally{ lock.releaseLock(); }
  }catch(err){ return J_({ok:false, erro:String(err)}); }
}

function registar_(d){
  var ss = folha_();
  var lanc = ss.getSheetByName('Lançamentos');
  var ult = lanc.getLastRow();
  if(ult >= 5){
    var ids = lanc.getRange(5,11,ult-4,1).getValues();
    for(var i=0;i<ids.length;i++)
      if(String(ids[i][0]) === String(d.id)) return {ok:true, duplicado:true, contados:d.contados};
  }
  var quando = Utilities.formatDate(new Date(),'Europe/Lisbon','yyyy-MM-dd HH:mm');
  var qtd_ = function(l){ return (l.quantidade === null || l.quantidade === undefined) ? l.parcelas : l.quantidade; };
  var linhas = d.linhas.map(function(l){
    return [d.data, d.armazem, d.seccao, l.categoria, l.produto, l.unidade || '',
            qtd_(l), l.observacoes || '', d.contador || '', quando, d.id];
  });
  lanc.getRange(Math.max(ult+1,5), 1, linhas.length, 11).setValues(linhas);

  var sh = ss.getSheetByName(d.armazem + ' - ' + d.seccao), novos = 0;
  if(sh){
    var ultL = Math.max(sh.getLastRow(),4), ultC = Math.max(sh.getLastColumn(),3);
    var cab = sh.getRange(4,1,1,ultC).getValues()[0], col = 0;
    for(var c=3;c<cab.length;c++){
      var v = cab[c];
      var t = (v instanceof Date) ? Utilities.formatDate(v,'Europe/Lisbon','yyyy-MM-dd') : String(v).trim();
      if(t === d.data){ col = c+1; break; }
    }
    if(!col){
      col = ultC + 1;
      sh.getRange(4,col).setNumberFormat('@').setValue(d.data).setBackground('#1f3864')
        .setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
      sh.setColumnWidth(col,95);
    }
    var mapa = {}, arts = ultL >= 5 ? sh.getRange(5,1,ultL-4,2).getValues() : [];
    for(var r=0;r<arts.length;r++)
      if(arts[r][1]) mapa[(String(arts[r][0]).trim()+'|'+String(arts[r][1]).trim()).toLowerCase()] = r+5;
    var prox = ultL < 5 ? 5 : ultL + 1;
    d.linhas.forEach(function(l){
      var k = (String(l.categoria).trim()+'|'+String(l.produto).trim()).toLowerCase();
      var linha = mapa[k];
      if(!linha){
        linha = prox++;
        sh.getRange(linha,1,1,3).setValues([[l.categoria, l.produto, l.unidade || '']]);
        sh.getRange(linha,3).setBackground('#fff2cc');
        mapa[k] = linha; novos++;
      }
      var cel = sh.getRange(linha,col).setValue(qtd_(l)).setHorizontalAlignment('right');
      if(l.observacoes) cel.setBackground('#fff2cc').setNote(l.observacoes);
    });
  }
  return {ok:true, contados:d.linhas.length, artigosNovos:novos};
}
