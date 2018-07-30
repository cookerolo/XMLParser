// Create and log an XML representation of the threads in your Gmail inbox.

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Send XML Email', functionName: 'doGet'},
  ];
  spreadsheet.addMenu('Scripts', menuItems);
}

function doGet() {
  
  //Get values from spreadsheet by spreadsheet ID (find it on the URL)
  var spreadsheet = SpreadsheetApp.openById("1GSrvY-qwT0HkjiZBuiBMUMvfoJj6Nyiy51LQ4IAulIY").getSheets()[0];
  
  //Select the range of interest in your spreadsheet
  var range = spreadsheet.getRange("A2:AA3000");
  
  //Iterate through each column to get all field values on the spreadsheet
  var values = range.getValues();
  var lastRow = spreadsheet.getLastRow();
  var root = XmlService.createElement("Avisos");
  for (i=0; i<lastRow-1 ;i++) {
    var posicion = values[i][13];
    var emp_idempresa = values[i][14];
    var emp_token = values[i][15];
    var txAccion = values[i][16];
    var idPlanPublicacion = values[i][17];    
    var txDescripcion = values[i][12];
    var txAviso2Url = values[i][18];
    var idPais = values[i][19];
    var idArea= values[i][20];
    var idTipoTrabajo= values[i][21];
    var txPuesto= values[i][22];
    var idIndustria= values[i][23];
    var txCiudad= values[i][24];
    var txZona= values[i][25];


    //Create "job" container for all job fields and iterate through jobs
    var child = XmlService.createElement("aviso");
      child.addContent(XmlService.createElement("DatosAdicionales").setAttribute("idPosicion", posicion).setAttribute("emp_idempresa",emp_idempresa).setAttribute("emp_token", emp_token))
      child.addContent(XmlService.createElement("txAccion").setText(txAccion))
      child.addContent(XmlService.createElement("idPlanPublicacion").setText(idPlanPublicacion))
      child.addContent(XmlService.createElement("txAviso2Url").setText(txAviso2Url))
      child.addContent(XmlService.createElement("idPais").setText(idPais))
      child.addContent(XmlService.createElement("idArea").setText(idArea))
      child.addContent(XmlService.createElement("idTipoTrabajo").setText(idTipoTrabajo))
      child.addContent(XmlService.createElement("txPuesto").setText(txPuesto))
      child.addContent(XmlService.createElement("idIndustria").setText(idIndustria))
      child.addContent(XmlService.createElement("txDescripcion").setText(txDescripcion))
      child.addContent(XmlService.createElement("txCiudad").setText(txCiudad))
      child.addContent(XmlService.createElement("txZona").setText(txZona))

    root.addContent(child);
  }
  
  var document = XmlService.createDocument(root);  
  var xml = XmlService.getPrettyFormat().format(document);  
  emailNotifier(xml,txCiudad);
}
function emailNotifier(body, city) {
  var d = new Date();
  var subject = 'XML file - ' + city + ' - ' + d.getDate() + '/' + (d.getMonth() + 1);
  var emailBody = body;
  var email = GmailApp.createDraft("victor.cooke@crossover.com", subject, emailBody);
  email.send();
}

