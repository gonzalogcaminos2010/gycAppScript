var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('reemplazos');

function createPedido(data) {
  var sheetRegistro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('remitos');
  
  var lastRow = sheetRegistro.getLastRow();
  var lastRemitoNumber = lastRow > 1 ? sheetRegistro.getRange(lastRow, 5).getValue() : 0;
  var newRemitoNumber = lastRemitoNumber + 1;
  
  var sheetRemitos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('inventario_general');
  lastRow = sheetRemitos.getLastRow();
  var range = sheetRemitos.getRange(1, 1, lastRow, 4); // Ahora la columna 4 es 'Ingresos'
  var values = range.getValues();
  
  data[1].forEach(function(productData) {
    var producto = productData[0];
    var cantidad = productData[1];

    for (var i = 0; i < values.length; i++) {
      if (values[i][0] == producto) {
        values[i][1] -= cantidad;
        values[i][2] = values[i][2] ? values[i][2] + cantidad : cantidad; // Egreso
        // Para 'Ingresos', puedes modificar la lógica según tus necesidades.
        if (values[i][1] <= 0) {
          SpreadsheetApp.getUi().alert('Advertencia: Se agotó el stock de ' + producto);
        }
        break;
      }
    }
  });

  range.setValues(values);
  var remitoRows = []; // Declaración del nuevo array.
  data[1].forEach(function(productData) {
    var row = [data[0], productData[0], productData[1], data[2], data[3], newRemitoNumber];
    remitoRows.push(row); // Aquí agregamos la fila al array.
    sheetRegistro.appendRow(row);
  });

  // Llama a crearRemitoYEnviar una vez con el array completo de remitoRows y el número de pedido.
  crearRemitoYEnviar(remitoRows, newRemitoNumber, data[0]);
}


function getProductList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('inventario_general');
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
  // Filtra cualquier celda vacía
  var productList = data.flat().filter(function (item) {
    return item !== '';
  });
  return productList;
}

function crearRemitoYEnviar(remitoRows, newRemitoNumber) {
  var templateId = '1ErNC4Tl6xOTanldfMvx6jSuGxccCYErkDYK5x-NENbk';
  var file = DriveApp.getFileById(templateId);
  var copy = file.makeCopy('Remito Nº ' + newRemitoNumber);
  var copyId = copy.getId();
  var docCopy = DocumentApp.openById(copyId);
  
  var body = docCopy.getBody();

  // Reemplazar los placeholders una vez al inicio.
  body.replaceText('\\{remitoNumero\\}', newRemitoNumber.toString());
  body.replaceText('\\{responsable\\}', remitoRows[0][3]);
  body.replaceText('\\{observaciones\\}', remitoRows[0][4]);

  // Crear un string que contenga la lista de productos.
  var productList = '';
  remitoRows.forEach(function(row, index) {
    productList += row[1] + ': ' + row[2] + '\n';
  });

  // Reemplazar {productos} con la lista de productos.
  body.replaceText('\\{productos\\}', productList);

  // Guardamos y cerramos el documento.
  docCopy.saveAndClose();
  
  // Agregamos un retraso de 3 segundos para asegurar que los cambios se hayan propagado.
  Utilities.sleep(3000);

  // Convertimos la copia en un archivo PDF.
  var pdfFile = DriveApp.getFileById(copy.getId()).getAs('application/pdf');

  // Movemos el archivo PDF a la carpeta especificada.
  var folder = DriveApp.getFolderById('1MgXXOaWifVW-bRB0A_cgdzfL3z26s5mL');
  var newFile = folder.createFile(pdfFile);

  // Trasladamos la copia del documento a la papelera.
  DriveApp.getFileById(copy.getId()).setTrashed(true);

  var email = 'gonzalogabrielcaminos@gmail.com'; 
  var subject = 'Pedido Nº ' + newRemitoNumber;
  var body = 'Aquí tienes el remito Nº ' + newRemitoNumber;
  GmailApp.sendEmail(email, subject, body, {
    attachments: [newFile]
  });
}

