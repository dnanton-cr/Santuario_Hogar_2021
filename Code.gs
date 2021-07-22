   /*
  Explicacion de lo que se debe obtener en crearFormSH():
  
  El formulario a crear tendrá:
  Favor escoger el grupo al que pertenece: y dar al botón "Next"
  Favor escoger el nombre de su pareja, y posteriormente emitir las opiniones de las parejas de su grupo.
  
  Proceso programado:
  Crear formulario, o abrirlo si ya existe.
  Borrar el formulario, para reescribirlo.
  Abrir las hojas que tienen el nombre Grupo en su pestaña, generar una opción en el formulario inicial y después de escoger ese grupo, brincar al subformulario relacionado a ese grupo, y además leer el contenido de cada hoja
  Escribir el contenido de cada hoja en un subformulario como opciones de una línea de texto.
  */

function crearFormSH() { // crear el formulario completamente
  var he = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.create("Formulario Santuario Hogar - Schoenstatt - IISem2020");
  form.setDescription("Formulario para recopilar las opiniones de las parejas que están participando en el taller para Santuario Hogar");
  var formUrl = form.getPublishedUrl();
  var formEditUrl = form.getEditUrl();
  Logger.log("Dirección para usar formulario: ",formUrl,"Dirección para editar formulario: ",formEditUrl);
  return formEditUrl;
}

function editarFormSH() { // editar el formulario creado anteriormente, y volver a cambiar los espacios
  var he = SpreadsheetApp.getActiveSpreadsheet();
  var form = FormApp.openByUrl(he.getFormUrl());
  // var form = FormApp.openByUrl("https://docs.google.com/forms/d/1bvoVSvLp69IVhuoQgZU_j6NpXGaRair19gcHQKp9few/edit");
  // form.setCollectEmail(false); // pedir el correo para actualizar y luego poder enviar resultados, valor debe estar en true para activarlo
  var formLen = form.getItems().length;
  Logger.log("Dirección del formulario: ",form,". Tamaño del formulario: ",formLen);
  // limpiar el formulario
  for (var i = 0; i < formLen; i++) {
    form.deleteItem(0);
  }
  
  //traer solo las hojas que tienen "Parejas" en el titulo de la pestaña
  var hGrupo = he.getSheets()
    .filter(function(sheet) {return sheet.getName().match(/Grupo/gi);});
    //Logger.log(hGrupo.length,hGrupo[0].getName(),hGrupo[1].getName());

  // creo una lista de opciones multiples conocida como "multiple choice items"
  var escogeGrupo = form.addMultipleChoiceItem()
    .setTitle('Grupos de Parejas - Taller Santuario Hogar Schoenstatt 2020');
    
  var gruposOpciones = [];
  for(var i = 0; i < hGrupo.length; i++) {
    var nombreGrupo = hGrupo[i].getName();
    
    var seccionGrupo = form.addPageBreakItem()
      .setTitle(nombreGrupo)
      .setGoToPage(FormApp.PageNavigationType.SUBMIT);
    
    var nombresParejas = hGrupo[i].getDataRange().getValues();
    
    var parejaOpina = nombresParejas;
    var parejasOpciones = [];
    var item = form.addListItem();
    for (var k = 1; k < parejaOpina.length; k++) { // incluye todos los nombres de las parejas del grupo, para escoger 1 pareja, equivalente a la pareja que emite opinion        
      item.setTitle("Favor escoger el nombre de su pareja, para opinar sobre las demás");
      item.isRequired();
      parejasOpciones.push(parejaOpina[k][0]);
      item.setChoiceValues(parejasOpciones);
    }
    
    for(var j = 1; j < nombresParejas.length; j++) { // inicio el contador en 1 para no leer el primer renglon que tiene solo el encabezado de la columna 
      form.addTextItem()
      .setTitle(nombresParejas[j][0] + ' - ' + nombreGrupo) // pongo como nombre del texto la celda del renglon en cuestion, pero siempre solo de la 1era columna
      .setHelpText('Favor opinar 3 virtudes sobre la pareja del grupo, separar las virtudes con coma (por ejemplo: virtud 1, virtud 2, virtud 3)');
    }
    // llenos las opciones de la lista de grupos al inicio del formulario, para luego ir a la pagina del grupo con las familias
    gruposOpciones.push(escogeGrupo.createChoice(nombreGrupo, seccionGrupo));
  }
  //asigna las opciones de los grupos en la variable escogeGrupo x escoger
  escogeGrupo.setChoices(gruposOpciones);
}

   /*
  Explicacion de lo que se debe obtener en crearReportSH():
  
  El correo a enviar se debe ver algo asi:
  Hola <Nombre de la pareja evaluada>, a continuación las opiniones que sus compañeros del grupo han emitido. 
  Opinión emitida por la pareja <nombre de la pareja que emite opinion>
  y asi sucesivamente.
  
  Proceso programado:
  Leer celda cabecera en renglón1, columna variante: Opiniones emitidas para la pareja (nombre en columna)
  Leer los renglones con respecto a cuando la columna no está vacía, y traer el nombre de la columna ParejaQueOpina, y el valor de la opinión. Texto: La pareja "Esposa y Esposo" opinaron "virtudes descritas en cada celda de la columna"
  */

function enviarReportSH() {
  var he = SpreadsheetApp.getActiveSpreadsheet();
  var hDatos = he.getSheetByName("Enviar"); // se debe cambiar el nombre de esta hoja por la que tenga los resultados de las encuestas
  var datosHojaPrueba = hDatos.getDataRange().getValues(); // trae los datos de las encuestas recogidas
  var hDatosGrupos = he.getSheetByName("Consolidado"); // identifica y ubica hoja Datos Grupos
  var datosHojaGrupos = hDatosGrupos.getDataRange().getValues(); // trae los datos de la hoja con todas las parejas y los grupos
  var hRespuestas = he.getSheetByName("Respuestas"); // hoja de trabajo para escribir los resultados de las encuestas en formato x renglones, luego se podria enviar directo a correo sin pasar por esta hoja
  var respuestas = []; // creo lista para obtener toda la informaci[on de las parejas y sus opiniones
  
  // var dirEmailPruebas = "dnanton@gmail.com, castilloraque@gmail.com, dnanton@gruasgmt.com";
  var dirEmailPruebas = "dnanton@gmail.com";
  var nombrePareja = "nombre de la pareja";
  var temaEmail = "Opinion de las parejas";
  var cuerpoEmail = "Nombres de las parejas y sus opiniones";
  // var htmlMachote = HtmlService.createTemplateFromFile("cuerpoCorreo");  
  
  var r = c = 1; // inicializo las variables para escribir en la hoja de practica
  hRespuestas.clear(); // borra los datos de la hoja de respuestas
  hRespuestas.activate(); // muestar la hoja de respuestas aun si antes estaba en otra hoja
  
  hRespuestas.getRange(r,c).setValue("Correo Enviado"); // pone el encabezado de la columna, 1er renglón
  c = c + 1;
  hRespuestas.getRange(r,c).setValue("Nombre Pareja"); // pone el encabezado de la columna, 1er renglón
  c = c + 1;
  hRespuestas.getRange(r,c).setValue("Email Pareja"); // pone el encabezado de la columna, 1er renglón
  c = c + 1;
  hRespuestas.getRange(r,c).setValue("Opiniones Pareja"); // pone el encabezado de la columna, 1er renglón
  r = r + 1;

  for (var col = 3; col < hDatos.getLastColumn(); col++) { // variable inicia en 3 porque es en la columna D que se tiene el primer nombre de la pareja evaluada
    nombrePareja = datosHojaPrueba[0][col]; // capturo el nombre de la pareja de la que se está opinando para buscar correo electronico
    
    for (var k = 1; k < datosHojaGrupos.length; k++) { // obtiene direccion de email de la pareja
      if (nombrePareja == datosHojaGrupos[k][0]) {
        dirEmail = datosHojaGrupos[k][4]+", "+datosHojaGrupos[k][6];
      }
    }
    
    c=1;
    hRespuestas.getRange(r,c).setValue("Enviado");
    c = c + 1;
    hRespuestas.getRange(r,c).setValue(nombrePareja);
    c = c + 1;
    hRespuestas.getRange(r,c).setValue(dirEmail);
    c = c + 1;
    
    var respuestas = avanzaRenglon(he,col,c,r);
    
    const htmlTemplate = HtmlService.createTemplateFromFile("correoParejas");
    htmlTemplate.h1 = nombrePareja;
    htmlTemplate.valoresRespuestas = respuestas;
    const htmlForEmail = htmlTemplate.evaluate().getContent();
    // var options = { htmlBody: htmlForEmail, bcc: "dnanton@gmail.com, castilloraque@gmail.com" }    // https://spreadsheet.dev/send-email-from-google-sheets
    // para aprender y aplicar estilos al html usar este sitio https://www.w3schools.com/html/html_css.asp
    var options = { htmlBody: htmlForEmail }    // https://spreadsheet.dev/send-email-from-google-sheets
    MailApp.sendEmail(dirEmail, nombrePareja, "Favor abrir el correo con un cliente que permita HTML", options);
    r = r + 1;
  }
  SpreadsheetApp.flush();
}

function avanzaRenglon(he,col,c,r) {
  var respuestas = [];
  var hPrueba = he.getSheetByName("Enviar");
  var datosHojaPrueba = hPrueba.getDataRange().getValues();
  var hRespuestas = he.getSheetByName("Respuestas");
  var valorRespuesta = [];
      
  for (var reng = 1; reng < datosHojaPrueba.length; reng++) { // Ciclo que avanzara por cada renglon de la hoja, inicia despues de encabezado = renglon 1
      valorRespuesta = datosHojaPrueba[reng][2]+" opinan de ustedes: "+datosHojaPrueba[reng][col];
      respuestas.push(valorRespuesta); // incluye en string las respuestas de las parejas, cada una separadas por una coma
      hRespuestas.getRange(r,c).setValue(respuestas[reng-1]); // escribe en la columna, la opinión emitida
      c = c + 1;
  }
  return respuestas; 
}

function crearReportSH_Old() {
  var he = SpreadsheetApp.getActiveSpreadsheet();
  var hPrueba = he.getSheetByName("Form Responses 1"); // nombre de hoja con los resultados de las encuestas
  var hDatosGrupos = he.getSheetByName("Consolidado"); // identifica y ubica hoja Consolidado
  var hRespuestas = he.getSheetByName("Respuestas"); // hoja de trabajo para escribir los resultados de las encuestas en formato x renglones, luego se podria enviar directo a correo sin pasar por esta hoja
  var datosHojaPrueba = hPrueba.getDataRange().getValues(); // trae los datos de las encuestas recogidas
  var datosHojaGrupos = hDatosGrupos.getDataRange().getValues(); // trae los datos de la hoja con todas las parejas y los grupos
  var r = c = 1; // inicializo las variables para escribir en la hoja de practica
  var respuestas = []; // creo lista para obtener toda la informaci[on de las parejas y sus opiniones
  
  var dirEmail = "dnanton@gmail.com";
  var nombrePareja = "nombre de la pareja";
  var temaEmail = "Opinion de las parejas";
  var cuerpoEmail = "Nombres de las parejas y sus opiniones";
  
  hRespuestas.clear(); // borra los datos de la hoja de respuestas
  hRespuestas.activate(); // muestra la hoja de respuestas aun si antes estaba en otra hoja
  hRespuestas.getRange(r,c).setValue("Correo Enviado"); // pone el encabezado de la columna, 1er renglón
  c = c + 1;
  hRespuestas.getRange(r,c).setValue("Nombre Pareja"); // pone el encabezado de la columna, 1er renglón
  c = c + 1;
  hRespuestas.getRange(r,c).setValue("Email Pareja"); // pone el encabezado de la columna, 1er renglón
  c = c + 1;
  hRespuestas.getRange(r,c).setValue("Opiniones Pareja"); // pone el encabezado de la columna, 1er renglón
  r = r + 1;
  
  for (var col = 4; col < hPrueba.getLastColumn(); col++) { // variable inicia en 4 porque es en la columna D que se tiene el primer nombre de la pareja evaluada
    nombrePareja = datosHojaPrueba[0][col]; // capturo el nombre de la pareja de la que se está opinando para buscar correo electronico
    
    for (var k = 1; k < datosHojaGrupos.length; k++) { // obtiene direccion de email de la pareja
      if (nombrePareja == datosHojaGrupos[k][0]) {
        dirEmail = datosHojaGrupos[k][4]+";"+datosHojaGrupos[k][6];
      }
    }
    
    c=1;
    hRespuestas.getRange(r,c).setValue("No enviado");
    c = c + 1;
    hRespuestas.getRange(r,c).setValue(nombrePareja);
    c = c + 1;
    hRespuestas.getRange(r,c).setValue(dirEmail);
    c = c + 1;
    
    var respuestas = avanzaRenglon(col,c,r);
    
    const htmlTemplate = HtmlService.createTemplateFromFile("correoParejas");
    htmlTemplate.h1 = nombrePareja;
    htmlTemplate.valoresRespuestas = respuestas;
    const htmlForEmail = htmlTemplate.evaluate().getContent();
    Logger.log(htmlForEmail);
    
    // MailApp.sendEmail("dnanton@gmail.com", nombrePareja, "Favor abrir el correo con un cliente que permita HTML", { htmlBody: htmlForEmail });
    r = r + 1;
    
  }
}

function avanzaRenglon_Old(col,c,r) {
  var respuestas = [];
  var he = SpreadsheetApp.getActiveSpreadsheet();
  var hPrueba = he.getSheetByName("Form Responses 1");
  var datosHojaPrueba = hPrueba.getDataRange().getValues();
  var hRespuestas = he.getSheetByName("Respuestas");
  
  for (var reng = 1; reng < datosHojaPrueba.length; reng++) { // Ciclo que avanzara por cada renglon de la hoja, inicia despues de encabezado = renglon 1
    if (datosHojaPrueba[reng][col] !== "") {
      var valorRespuesta = datosHojaPrueba[reng][3]+" opinan de ustedes: "+datosHojaPrueba[reng][col]; // no se puede fijar a columna 3, hay que buscar la primera columna en el renglon con datos 
      respuestas.push(valorRespuesta); // students.push(studentValues[i].join(' '));
      // escribe en la columna, la opinión emitida
      hRespuestas.getRange(r,c).setValue(respuestas[reng-1]);
      c = c + 1;
    }
  }
  return respuestas; 
}

// https://www.youtube.com/watch?v=fx6quWRC4l0 para entender como crear el correo con la información en html
