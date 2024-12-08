function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('a definni')
    .addItem('Générer le document', 'cree_contra')
    .addItem('Vérifier les donnés', 'verif')
    .addToUi();
}

function verif(){
  let feuille = SpreadsheetApp.getActiveSpreadsheet();
  let sheet_variable = bibliho_sheet.getTab(feuille.getUrl(), 'Variables');
  let contrat_sheet = bibliho_sheet.getTab(feuille.getUrl(), 'Structure du contrat');
  let model_sheet = bibliho_sheet.getTab(feuille.getUrl(), 'Configuration');


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
    feuille_Log.getRange('C2').setValue('Invalide: probleme avec l\'url du Dossier de travail').setBackground('#ff0000').setFontColor('#ffffff');
  }


  /*verification clé valeur*/
  feuille_Log.getRange('A3').setValue("Variables");
  feuille_Log.getRange('B3').setValue("clé valeur");
  let i = 0;
  while(i<sheet_variable.length){
    if(! sheet_variable[i][2] == ''){
      if(sheet_variable[i][1] == ''){
        feuille_Log.getRange('C3').setValue('Invalide: clé '+sheet_variable[i][2]+' sans valeur (ligne: '+(i - 1)+')').setBackground('#ff0000').setFontColor('#000000');
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
        feuille_Log.getRange('C5').setValue('invalide: type de l\'onglet '+sheet_variable[i][5]+' sans valeur (ligne: '+i-1+')').setBackground('#ff0000').setFontColor('#000000');
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
  while(i<model_sheet.length){

    if(! model_sheet[i][2] == ''){
      response = UrlFetchApp.fetch(model_sheet[i][2],{muteHttpExceptions:true});

      if(response.getResponseCode() == 200) {
        model_sheet.getRange(i,4).setValue("Ok").setBackground('#6aa84f').setFontColor('#000000');
      }
      else{
        ok = false;
        model_sheet.getRange(i,4).setValue('URL invalide').setBackground('#ff0000').setFontColor('#ffffff');
      }
    }
    i++
  }

  if(ok){
    feuille_Log.getRange('C6').setValue("valide").setBackground('#6aa84f').setFontColor('#000000');
  }
  else{
    feuille_Log.getRange('C6').setValue('Invalide: ').setBackground('#ff0000').setFontColor('#ffffff');
  }

  
}

function cree_contra() {
  let feuille = SpreadsheetApp.getActiveSpreadsheet();
  let tab_sheet = bibliho_sheet.getTab(feuille.getUrl(), 'Structure du contrat');
  let sheet_variable = bibliho_sheet.getTab(feuille.getUrl(), 'Variables');

  let tab_utile = [];
  
  let verif = false;
  let mauvais = '';
  let url = '';
  let num_Art = 1;
  let num_Anx = 1;

  tab_sheet.forEach(function(ligne){
    if(ligne[0] == 'Oui'){
      verif == true;
      verif = CrationDeContra.recupertation(ligne[4], ligne[5])
      if(verif == false){
        mauvais = ligne[5];
        return;
      }
      if(verif == null){
        url = ligne[5];
        return
      }

      if(ligne[6] == 'Article'){
        CrationDeContra.ajout_de_titre('Article', ligne[1], num_Art, verif);
        num_Art++;
      }
      if(ligne[6] == 'Annexe'){
        CrationDeContra.ajout_de_titre('Annexe', ligne[1], num_Anx, verif);
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

  CrationDeContra.insertion(document_vide, tab_utile);

  let key = [];
  let val = [];

  sheet_variable.forEach(function(ligne){
    if(! ligne[2] == ''){
      key.push(ligne[2]);
      val.push(ligne[1]);
    }
  });

  CrationDeContra.replacevar(document_vide,key, val);

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
