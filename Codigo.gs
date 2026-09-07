/**
 * INVENTÁRIO DO GRUPO — porta de entrada das contagens.
 * A app (https://groupmultifund.github.io/inventario/) envia para aqui cada
 * contagem fechada e este script escreve-a na folha "INVENTÁRIO — Mestre",
 * na pasta 02 Contagens do Drive.
 *
 * Instalação: correr instalar() uma vez e autorizar. Depois
 * Implementar > Nova implementação > Aplicação Web
 * (Executar como: eu | Quem tem acesso: Qualquer pessoa).
 */
var PASTA_ID = '1iyIUHBspvWLrW3EpmIiDgDx8Vz_YE3cY';   // pasta 02 Contagens
var NOME     = 'INVENTÁRIO — Mestre';
var CODIGO   = 'Sabiocrescimento2026';   // palavra-passe da app
var CATALOGO = 'https://groupmultifund.github.io/inventario/catalogo.json';
var AZUL     = '#1f3864';
var COLS = ['Data contagem','Armazém','Secção','Categoria','Produto','Unidade',
            'Quantidade','Observações','Contado por','Registado em','ID lançamento'];

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
  return criar_();
}

function criar_(){
  var cat = JSON.parse(UrlFetchApp.fetch(CATALOGO).getContentText());
  var ss = SpreadsheetApp.create(NOME);
  DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(PASTA_ID));

  var leia = ss.getSheets()[0].setName('Leia-me');
  var txt = [['INVENTÁRIO — Grupo (folha mestre)'],[''],
    ['Os gerentes lançam a contagem na app do telemóvel e ela entra aqui sozinha.'],
    ['Uma folha por armazém e secção: as colunas A a C são a lista oficial de artigos'],
    ['e cada contagem acrescenta uma coluna nova com a data.'],
    ['A folha "Lançamentos" tem o histórico linha a linha — é a que serve para análise.'],[''],
    ['Para acrescentar um artigo permanente, escreva a linha nas colunas A a C da folha do armazém.'],
    ['Célula vazia = artigo não contado. Zero = contado e não há.'],
    ['Não escrever à mão nas colunas com data.'],
    ['Fundo amarelo = artigo acrescentado pelo gerente durante a contagem, ou célula com observação.']];
  leia.getRange(1,1,txt.length,1).setValues(txt).setFontFamily('Arial');
  leia.getRange(1,1).setFontSize(15).setFontWeight('bold').setFontColor(AZUL);
  leia.setColumnWidth(1,640);

  cat.forEach(function(sec){
    var sh = ss.insertSheet(sec.armazem + ' - ' + sec.seccao);
    sh.getRange(1,1).setValue('INVENTÁRIO — ' + sec.armazem + ' · ' + sec.seccao)
      .setFontSize(13).setFontWeight('bold').setFontColor(AZUL);
    sh.getRange(2,1).setValue('Colunas A-C: lista oficial de artigos. Colunas de data: preenchidas pela app.')
      .setFontSize(9).setFontStyle('italic').setFontColor('#808080');
    sh.getRange(4,1,1,3).setValues([['Categoria','Produto','Unidade']])
      .setBackground(AZUL).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    if(sec.itens.length){
      sh.getRange(5,1,sec.itens.length,3).setValues(sec.itens.map(function(it){
        return [it.categoria, it.produto, it.unidade || ''];
      }));
    }
    sh.setColumnWidth(1,190); sh.setColumnWidth(2,260); sh.setColumnWidth(3,70);
    sh.setFrozenRows(4); sh.setFrozenColumns(3);
  });

  var lanc = ss.insertSheet('Lançamentos');
  lanc.getRange(1,1).setValue('LANÇAMENTOS DE INVENTÁRIO — histórico')
    .setFontSize(13).setFontWeight('bold').setFontColor(AZUL);
  lanc.getRange(2,1).setValue('Uma linha por artigo contado. Preenchido pela app.')
    .setFontSize(9).setFontStyle('italic').setFontColor('#808080');
  lanc.getRange(4,1,1,COLS.length).setValues([COLS])
    .setBackground(AZUL).setFontColor('#ffffff').setFontWeight('bold');
  lanc.setFrozenRows(4);
  lanc.setColumnWidth(4,180); lanc.setColumnWidth(5,230);
  lanc.setColumnWidth(8,220); lanc.setColumnWidth(11,240);

  P_().setProperty('FOLHA_ID', ss.getId());
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
    if(d.acao === 'entrar') return J_({ok:true, entrada:true});
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
  lanc.getRange(Math.max(ult+1,5), 1, linhas.length, COLS.length).setValues(linhas);

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
      sh.getRange(4,col).setNumberFormat('@').setValue(d.data).setBackground(AZUL)
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
