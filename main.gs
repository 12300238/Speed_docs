function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Speed docs')
    .addItem('Générer le document', 'cree_contra')
    .addItem('Vérifier les donnés', 'verif')
    .addToUi();
}

function verif(){
  let feuille = SpreadsheetApp.getActiveSpreadsheet();
  let sheet_variable = getTab(feuille.getUrl(), 'Variables');
  let contrat_sheet = getTab(feuille.getUrl(), 'Structure du contrat');
  let model_sheet = getTab(feuille.getUrl(), 'Configuration');
  let feuille_model = feuille.getSheetByName('Configuration')


  /*initialisation de la feuille log*/
  var feuille_Log = feuille.getSheetByName("Log");
  feuille_Log.clear();
  feuille_Log.setFrozenRows(1);
  feuille_Log.getRange('A1').setValue("feuille");
  feuille_Log.getRange('B1').setValue("objet verifier");
  feuille_Log.getRange('C1').setValue("Etat");
  feuille_Log.getRange('1:1').setFontWeight('bold');
  feuille_Log.getRange('2:4').setBackground(null);

  /*verifier URL dossier model*/
  feuille_Log.getRange('A2').setValue("Variables");
  feuille_Log.getRange('B2').setValue("Dossier de travail");
  let response = UrlFetchApp.fetch(sheet_variable[2][1],{muteHttpExceptions:true});

  if(response.getResponseCode() == 200) {
    feuille_Log.getRange('C2').setValue("valide").setBackground('#6aa84f').setFontColor('#000000');
  }
  else{
    feuille_Log.getRange('C2').setValue('Invalide: probleme avec l\'url du Dossier de travail').setBackground('#ff0000').setFontColor('#000000');
  }


  /*verification clé valeur*/
  feuille_Log.getRange('A3').setValue("Variables");
  feuille_Log.getRange('B3').setValue("clé valeur");
  let i = 0;
  while(i<sheet_variable.length){
    if(! sheet_variable[i][2] == ''){
      if(sheet_variable[i][1] == ''){
        feuille_Log.getRange('C3').setValue('Invalide: clé '+sheet_variable[i][2]+' sans valeur (ligne: '+(i + 1)+')').setBackground('#ff0000').setFontColor('#000000');
        break;
      }
      feuille_Log.getRange('C3').setValue("valide").setBackground('#6aa84f').setFontColor('#000000');
    }
    i++;
  }


  /*verifiaction selecteur*/
  feuille_Log.getRange('A4').setValue("Structure du contrat");
  feuille_Log.getRange('B4').setValue("Selecteur");

  let oui = false;
  i=1;
  while(i<contrat_sheet.length){
    if(contrat_sheet[i][0]=='Oui'){
      oui = true
    }
    else if (contrat_sheet[i][0]==''){
      console.log(i)
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Structure du contrat').getRange(i+1,1).setValue("Non");
    }
    i++;
  }

  if(oui == true){
    feuille_Log.getRange('C4').setValue("valide").setBackground('#6aa84f').setFontColor('#000000');
  }
  else{
    feuille_Log.getRange('C4').setValue("Attention: tout les selecteur sont a Non").setBackground('#ff9700').setFontColor('#000000');
  }


  /*verfication des type des fichier*/
  feuille_Log.getRange('A5').setValue("Structure du contrat");
  feuille_Log.getRange('B5').setValue("Type des fichier");
  i=1;
  while(i<contrat_sheet.length){
    if(! contrat_sheet[i][5]==''){
      if(contrat_sheet[i][6]==''){
        feuille_Log.getRange('C5').setValue('invalide: type de l\'onglet '+contrat_sheet[i][5]+' sans valeur (ligne: '+(i + 1)+')').setBackground('#ff0000').setFontColor('#000000');
        break;
      }
    }
    feuille_Log.getRange('C5').setValue("valide").setBackground('#6aa84f').setFontColor('#000000');
    i++;
  }


  /*vérification des url des fichier model*/
  feuille_Log.getRange('A6').setValue("Configuration");
  feuille_Log.getRange('B6').setValue("URL des model");

  i=1;
  let ok = true;
  let invalide = 0;
  while(i<model_sheet.length){

    if(! model_sheet[i][2] == ''){
      response = UrlFetchApp.fetch(model_sheet[i][2],{muteHttpExceptions:true});
      
      if(response.getResponseCode() == 200 || response.getResponseCode() == 401)  {
        feuille_model.getRange(i+1,5).setValue("Ok").setBackground('#6aa84f').setFontColor('#000000');
        model_sheet[i][4] = 'val';
      }
      else{
        ok = false;
        invalide = i+1;
        feuille_model.getRange(i+1,5).setValue('URL invalide').setBackground('#ff0000').setFontColor('#000000');
      }
    }
    else{feuille_model.getRange(i+1,5).setValue('').setBackground('#ffffff').setFontColor('#000000');}
    i++
  }

  if(ok){
    feuille_Log.getRange('C6').setValue("valide").setBackground('#6aa84f').setFontColor('#000000');
  }
  else{
    feuille_Log.getRange('C6').setValue('Invalide: l\'url ligne '+invalide+' n\'est pas reconu').setBackground('#ff0000').setFontColor('#000000');
  }
  
  /*vérification des tab des fichier model*/
  feuille_Log.getRange('A7').setValue("Configuration");
  feuille_Log.getRange('B7').setValue("onglet des model");

  i=2;
  ok = true;
  invalide = 0;
  while(i<model_sheet.length){

    if(! model_sheet[i][2] == ''){
      if(model_sheet[i][4] == 'val'){
        if(! model_sheet[i][3] == ''){
          if(! getTabsname(DocumentApp.openByUrl(model_sheet[i][2]).getTabs()).includes(model_sheet[i][3])){
            ok = false;
            invalide = i+1;
            feuille_model.getRange(i+1,5).setValue('longlet n\'a pas été trouver dans le document').setBackground('#ff0000').setFontColor('#000000');
          }
        }
        else{
          ok = false;
          invalide = i+1;
          feuille_model.getRange(i+1,5).setValue('onglet manquant').setBackground('#ff0000').setFontColor('#000000');
        }
      }
    }
    else{
      if(! model_sheet[i][3] == ''){
        ok = false;
        invalide = i+1;
        feuille_model.getRange(i+1,5).setValue('lien manquant').setBackground('#ff0000').setFontColor('#000000');
      }
      else{feuille_model.getRange(i+1,5).setValue('').setBackground('#ffffff').setFontColor('#000000');}
    }
    i++
  }

  if(ok){
    feuille_Log.getRange('C7').setValue("valide").setBackground('#6aa84f').setFontColor('#000000');
  }
  else{
    feuille_Log.getRange('C7').setValue('Invalide: probleme avec l\'onglet ligne '+invalide).setBackground('#ff0000').setFontColor('#000000');
  }
}

function cree_contra() {
  let feuille = SpreadsheetApp.getActiveSpreadsheet();
  let tab_sheet = getTab(feuille.getUrl(), 'Structure du contrat');
  let sheet_variable = getTab(feuille.getUrl(), 'Variables');

  let tab_utile = [];
  
  let verif = false;
  let mauvais = '';
  let url = '';
  let num_Art = 1;
  let num_Anx = 1;

  tab_sheet.forEach(function(ligne){
    if(ligne[0] == 'Oui'){
      verif == true;
      verif = recupertation(ligne[4], ligne[5])
      if(verif == false){
        mauvais = ligne[5];
        return;
      }
      if(verif == null){
        url = ligne[5];
        return
      }

      if(ligne[6] == 'Article'){
        ajout_de_titre('Article', ligne[1], num_Art, verif);
        num_Art++;
      }
      if(ligne[6] == 'Annexe'){
        ajout_de_titre('Annexe', ligne[1], num_Anx, verif);
        num_Art++;
      }
      tab_utile.push(verif);
    }
  });

  if(!mauvais == ''){
    let msg1 = "erreure, ellement: "
    Browser.msgBox(msg1.concat(mauvais," non trouver dans le doc model"));
    return;
  }

  if(!url == ''){
    let msg1 = "erreure, url du document contenant l'ellement: "
    Browser.msgBox(msg1.concat(url," non valide"));
    return;
  }

  let cname = sheet_variable[5][1];
  let cal = new Date();
  let jour = cal.getDate();
  let moi = cal.getMonth() + 1;
  let year = cal.getFullYear();
  let houre = cal.getHours();
  let min = cal.getMinutes();

  tot= ''+cname+year+moi+jour+houre+min;

  let document_vide = DocumentApp.create(tot);

  insertion(document_vide, tab_utile);

  let key = [];
  let val = [];

  sheet_variable.forEach(function(ligne){
    if(! ligne[2] == ''){
      key.push(ligne[2]);
      val.push(ligne[1]);
    }
  });

  replacevar(document_vide,key, val);

  let folderid = sheet_variable[2][1];
  folderid = folderid.split('/');
  folderid = folderid[folderid.length-1];

  try{
    let file = DriveApp.getFileById(document_vide.getId());
    var folder = DriveApp.getFolderById(folderid);
    file.moveTo(folder);
  }
  catch(e){
    console.log('echeque');
    let filever = DriveApp.getFilesByName(tot);
    var folder = DriveApp.getFolderById(folderid);

    if(filever.hasNext()){
      filever.next().moveTo(folder);
    }
  }
}

/**
 * crée un document et renvoy son objet
 * @param {string} nom_du_doc: le nom du document
 * @return {objet} le document du contrat
 */
function creationDocument(nom_du_doc) {
  let new_doc = DocumentApp.create(nom_du_doc);

  return new_doc;
}

/**
 * prend un document model et nom d'onglet et recupert le body qui est a se nom
 * @param{string} l'url du model
 * @param{string} le nom de l'onglet shouaiter
 * @return{body} l'onglet shouaiter
 * @return{false} l'objet na pas été trouver
 */
function recupertation(model, tab){
  let onglets;
  try{
    onglets = DocumentApp.openByUrl(model).getTabs();
  } catch(error){
    return null;
  }
  let find = false;
  let onglet_utile;

  let i = 0;
  while(i<onglets.length && find == false){
    if(onglets[i].getTitle() == tab){
      find = true;
      onglet_utile = onglets[i].asDocumentTab().getBody().copy();
    }
    i++;
  }

  if(find == false){
    return false;
  }
  return onglet_utile;
}

/**
 * met les dans un document les donné des body fournie
 * @param{objet body} le document ou on insert les body
 * @param{list_objet} la liste des body a inséré
 */
function insertion(new_doc, tab_utile){
  tab_utile.forEach(function(tab){
    let i=0;
    while(i<tab.getNumChildren()){
      if(tab.getChild(i).getType() == 'PARAGRAPH'){
        new_doc.getBody().appendParagraph(tab.getChild(i).copy());
      }
      else if(tab.getChild(i).getType() == 'LIST_ITEM'){
        new_doc.getBody().appendListItem(tab.getChild(i).copy());
      }
      else if(tab.getChild(i).getType() == 'TABLE'){
        new_doc.getBody().appendTable(tab.getChild(i).copy());
      }
      i++;
    }
  });
}

/**
 * ajoute des éditeur et des lecteur au document en paramétre
 * @param {objet} new_doc: le document a manipuler
 * @param {list_string} editeur: les adresse mail des editeur
 * @param {list_string} lecteur: les adresse mail des lecteur
 */
function droit_accer(new_doc, editeur, lecteur){
  editeur.forEach(function(edit){
    new_doc.addEditor(edit);
  });

  lecteur.forEach(function(lec){
    new_doc.addViewer(lec);
  });
}

/**
 * ajoute un element titre dans un document
 *
 * @param{string} le type de section
 * @param{string} le tritre
 * @param{int} le numero de l'ellement ou -1 si il n'ent a pas
 * @param{object body} le body ou ajouté le titre
 * @return {}
 */
function ajout_de_titre(sectype, titre, num, body){
  let title = sectype+' '+num+' – '+titre;
  body.insertParagraph(0, title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var titre = body.getParagraphs()[0].editAsText();
  titre.setForegroundColor('#c00000').setFontFamily('Arial').setBold(true).setFontSize(11);
}

/**
 * prend un dossier un tableau de clé, un tableau de valeur et dans le fichier remplace le clé par les valeur du même index
 * 
 * @param{document} le document ou fair les changement
 * @param{tab_string} le nom des valuer a remplacer
 * @param{tab_string} les nouvelles valeur
 */
function replacevar(document_vide,key, val){
  if(key.length == val.length){
    let i=0;
    while(i<key.length){
      document_vide.replaceText(key[i], val[i]);
      i++;
    }
  }
  else{
    console.log("il n'y pas autent de clé que de valeur");
    Browser.msgBox("il n'y pas autent de clé que de valeur sur la feuille variable");
  }
}

/**
 * prend en une feuille et renvois un tableau avec toute ces case utile
 * @param {string} : l'url du classeur
 * @param {string} : le nom de la feuille utiliser
 * @return {srting_tab} : les vale de ma feuille sous forme de double tableau
 */
function getTab(feuille_url, nom_feuille) {
  var sheet = SpreadsheetApp.openByUrl(feuille_url).getSheetByName(nom_feuille);

  let range = sheet.getDataRange();
  let tab = range.getValues();
  return tab;
}

/**
 * cette fonctoin renvois les nom des tabs donner
 * @param {list_tab} la liste des tabe dont on veut le nom
 * @return {list_string} le tableau avec le nom des tab
 */
function getTabsname(tabs) {
  let listTab = [];
  tabs.forEach(function(tab){
    listTab.push(tab.getTitle());
  })
  return listTab;
}

/**
 * prend un tableau de tableau en paramétre et retourne la valeur de la case spécifié
 * @param {tab} : le tableau ou cherhcer
 * @param {int} : ligne
 * @param {int} : colone
 * @return {string} : l'element chercher
 */
function recupeCase(tab, ligne, colone) {
  return tab[ligne][colone];
}

/**
 * prend une un tableau et si le premier ellement est une coix return le deuxieme sinon return false
 * @param {tab} : la le tableau
 * @return {string} : le deuxieme ellement du tab
 * @return {false}
 */
function useList(ligne) {
  if( ligne[0] == 'x'){
    return ligne[1];
  }
  return false;
}
