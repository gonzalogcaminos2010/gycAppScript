function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestión de Stock')
      .addItem('Crear Corona', 'showFormCoronas')
      .addItem('Crear Pedido', 'showFormPedidos')
      .addToUi();
}


function showFormPedidos() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('FormularioRemitos.html')
    .setWidth(400) // puedes ajustar estas dimensiones
    .setHeight(500); // puedes ajustar estas dimensiones
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Formulario Pedidos');
}

function showFormCoronas() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Form.html')
    .setWidth(400) // puedes ajustar estas dimensiones
    .setHeight(500); // puedes ajustar estas dimensiones
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Formulario Coronas');
}


function createCorona(coronaData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('reemplazos');

  var lastRow = sheet.getLastRow(); // Cambiamos sheetRegistro por sheet
  var lastPedidoNumber = lastRow > 1 ? sheet.getRange(lastRow, 6).getValue() : 0; // Cambiamos el nombre de la variable a lastPedidoNumber
  var newPedidoNumber = lastPedidoNumber + 1;

  var sheetCoronas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('coronas_modelo');
  lastRow = sheetCoronas.getLastRow();
  var range = sheetCoronas.getRange(1, 1, lastRow, 4); 
  var values = range.getValues();
  
  var fecha = coronaData[0];
  var productos = coronaData[1];
  var responsable = coronaData[2];
  var observaciones = coronaData[3];

  productos.forEach(function(productData) {
    var producto = productData[0];
    var cantidad = productData[1];

for (var i = 0; i < values.length; i++) {
  if (values[i][0] == producto) {
    values[i][1] -= cantidad;
    values[i][2] = cantidad; // Asentamos la cantidad en la columna 'Egreso1' sin acumularla
    if (values[i][1] <= 0) {
      SpreadsheetApp.getUi().alert('Advertencia: Se agotó el stock de ' + producto);
    }
    break;
  }
}


    // Añadir la fila
    sheet.appendRow([fecha, producto, cantidad, responsable, observaciones, newPedidoNumber]); // Utilizamos newPedidoNumber
  });

  // Actualizar los valores en la hoja de cálculos 'coronas_modelo'
  range.setValues(values);
}




function crearRemitoYEnviar(data, productosList, newPedidoNumber) {
  var templateId = '1ErNC4Tl6xOTanldfMvx6jSuGxccCYErkDYK5x-NENbk';
  var file = DriveApp.getFileById(templateId);
  var copy = file.makeCopy('Remito Nº ' + newPedidoNumber + ' para ' + data[1]);
  var copyId = copy.getId();
  var docCopy = DocumentApp.openById(copyId);
  
  var body = docCopy.getBody();
  body.replaceText('\\{remitoNumero\\}', newPedidoNumber.toString());

  var productosText = productosList.map(function(productData) {
    return productData[0] + ': ' + productData[1];
  }).join(', ');

  body.replaceText('\\{producto\\}', productosText);
  body.replaceText('\\{responsable\\}', data[2]);
  body.replaceText('\\{observaciones\\}', data[3]);

  docCopy.saveAndClose();
  Utilities.sleep(3000);

  var pdfFile = DriveApp.getFileById(copy.getId()).getAs('application/pdf');

  var folder = DriveApp.getFolderById('1MgXXOaWifVW-bRB0A_cgdzfL3z26s5mL');
  var newFile = folder.createFile(pdfFile);

  DriveApp.getFileById(copy.getId()).setTrashed(true);

  var email = 'gonzalogabrielcaminos@gmail.com'; 
  var subject = 'Pedido Nº ' + newPedidoNumber + ' para ' + data[1];
  var body = 'Aquí tienes el remito Nº ' + newPedidoNumber + ' para ' + data[1];
  GmailApp.sendEmail(email, subject, body, {
    attachments: [newFile]
  });
}


function getCoronaList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('coronas_modelo');
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1); // Asume que los datos de las coronas están en la columna 1 y comienza en la fila 2
  var values = range.getValues();
  
  // Flatten the 2D array to 1D
  var coronas = values.reduce(function(a, b) {
    return a.concat(b);
  }, []);

  return coronas;
}


