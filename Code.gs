/***********************Generar Menú****************************** */

function onOpen(e) {
  var startingsheet = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getUi().createMenu("Camionetas")
  .addItem('Ejecutar', 'mostrarBarra')
  .addToUi();
}

//*Barra lateral
function mostrarBarra(){
  var html = HtmlService.createHtmlOutputFromFile('barraLateral')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle("Informes camionetas")
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

/***************************************************************** */

/*
function generarPrueba(){
global("TLAXCALA","16IEXma1KmGLYshoUpoBrpBpgyXt8cV24");
}
*/





function global(argumentoTerritorio,argumentoIdCarpeta){
 /************************Traer Datos Base Conjunta*******************************************************/  
  
  //traer la hoja de cálculo de donde salen los datos por su id
  var archivoExterno =SpreadsheetApp.openById("1sX-TPywUPCllQV_OhSZlf-h6R6PwwZEtF5lGMo-OVls");
  
  // traer las hojas del archivo externo
  var hojaConjunta= archivoExterno.getSheetByName("base_conjunta");
  var hojaMatch= archivoExterno.getSheetByName("IDs");


  //traer las ultimas filas y columnas base conjunta
  var ultimaFilaConjunta= hojaConjunta.getLastRow();
  var ultimaColumnaConjunta= hojaConjunta.getLastColumn();

  //traer las ultimas filas y columnas IDs
  var ultimaFilaMatch= hojaMatch.getLastRow();
  var ultimaColumnaMatch= hojaMatch.getLastColumn();

  /*************************Hacer Match Filtrado***********************************************************/
  //traer todos los valores
  var arregloMatchCompleto= hojaMatch.getRange(1,1, ultimaFilaMatch,2).getValues();

  // condicionar solo traer de acuerdo a territorio
  var arregloMatchSemi= arregloMatchCompleto.filter(function(item){
  return item[1]==argumentoTerritorio; // Iteracion
  });

  // traer solo la primer columna de ese arreglo semi
  var arregloMatch=[];
  
  for(var z=0;z<= arregloMatchSemi.length-1;z++){
    var unoPorUno= arregloMatchSemi[z][0];
    arregloMatch.push(unoPorUno);
  }



  /***********************************  crear folders y subfolder **************************/
    var folderTraido=DriveApp.getFolderById(argumentoIdCarpeta); 
    var subFolder= folderTraido.createFolder(argumentoTerritorio); //Crear subfolder nombre del subfolder *
    var idSubFolder= subFolder.getId();

  /************************* Hacer filtro de base Conjunta**************************************************** */ 
  var datos_originales= hojaConjunta.getRange(1,1,ultimaFilaConjunta,ultimaColumnaConjunta).getValues();

    for (i=0; i<=arregloMatch.length-1; i++){

      /*hacer el filtro mediante ciertas condiciones*/
      var datos_filtrados= datos_originales.filter(function(item){
      return item[1]==arregloMatch[i]; // Iteracion
      });
      /**/




      /*********************filtrar solo las columnas que me interesan slice push******************/ //TUTORIAL
        var nuevoArreglo=[];
        for(var k=0;k<= datos_filtrados.length-1;k++){
        var unoPorUno= datos_filtrados[k].slice(6,11);
        nuevoArreglo.push(unoPorUno);
        }
      /**/

      /************************añadir un dia a la fecha***********

        for (w=0;w<=nuevoArreglo.length-1;w++) {
        columna=0;
        var fechaAnterior= new Date(nuevoArreglo[w][columna]);

        //SUMARLE 24 HORAS
        var milisegundosUnDia = 1000 * 60 * 60 * 24;
        var nuevaFecha = new Date(fechaAnterior.getTime() + milisegundosUnDia);
        
        //Cambiarle formato
        //var nuevaFecha = Utilities.formatDate(nuevaFecha, 'America/Chicago', 'dd/MM/yyyy');

        //reemplazar
        nuevoArreglo[w].splice(0,1,nuevaFecha);
  
        }
        ********************* */


      /***************************copia del archivo*********************************************** */
        nombreCopia=(datos_filtrados[0][3]);
      
      /*************************AQUI TIENE QUE IR EL MALDITO CAMBIO***************************************/

        var tamanio= nuevoArreglo.length;

        if(tamanio<=4){
          documentoCopiado= DriveApp.getFileById("1vkzCH4NZq5ZqIbWED-rGx92JRWptYRa5ojNHfFcTlHY").makeCopy(nombreCopia);
        } else if( tamanio<=8){
          documentoCopiado= DriveApp.getFileById("1L4bm-NmuIWnEbd61QN3rNJ5HV45-1TMGFytyzy3uyBM").makeCopy(nombreCopia);
        } else if(tamanio<=12){
          documentoCopiado= DriveApp.getFileById("14Vx955tnUXGwMULn3dHOK-lrgXB-7oQ2kCNiA41GlBo").makeCopy(nombreCopia);
        } else{
          documentoCopiado= DriveApp.getFileById("1hydEOzqbUkiUGCjeJQlaXeF_J4WyNq-6kOxrrcny8LQ").makeCopy(nombreCopia);
        }

        
  
        var idNuevoDocumento = (documentoCopiado.getId());

      /**/


      /********************************Traer la hoja************************************************* */
        //traer la hoja de cálculo Plantilla por su id
        var archivoPlantilla =SpreadsheetApp.openById(idNuevoDocumento);

        // traer las hojas del archivo Plantilla
        var hojaPlantilla= archivoPlantilla.getSheetByName("Hoja1");
      /**/


      /************************************pegar valores *************************************************/
  
        //Pegar la tabla
        var rangoAPegar= hojaPlantilla.getRange(14,1, nuevoArreglo.length,nuevoArreglo[0].length);
        rangoAPegar.setValues(nuevoArreglo);

        //pegar vehículo
        var rangoAPegar= hojaPlantilla.getRange(7,2);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][11]);

        //pegar el territorio
        var rangoAPegar= hojaPlantilla.getRange(7,5);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][4]);
        
        //pegar placas
        var rangoAPegar= hojaPlantilla.getRange(8,2);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][5]);

        //pegar monedero
        var rangoAPegar= hojaPlantilla.getRange(9,5);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][2]);
         
        //pegar serie NIV
        var rangoAPegar= hojaPlantilla.getRange(9,2);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][12]);

   /****************pegar hasta abajo****************/
        //pegar el resguardante
        var rangoAPegar= hojaPlantilla.getRange(1,2);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][3]);
      
      
        //pegar el jud
        var rangoAPegar= hojaPlantilla.getRange(1,1);
        // arreglo[0][3] [fila] [columna]
        rangoAPegar.setValue(datos_filtrados[0][13]);



   
      /*************************************************Mover archivo */


      //var archivo = DriveApp.getFileById("1ITt4o6ePYun2-iyezxQQeIaDDNJcVuWeMAEJ_dw1vAE"); //mover archivo
      DriveApp.getFolderById(idSubFolder).addFile(documentoCopiado);

      /**/

      /***************************************PONER LOS DATOS EN LA HOJA */

      var libro =SpreadsheetApp.getActive();
      var hojaDeTrabajo= libro.getSheetByName('Sheet1');
      hojaDeTrabajo.appendRow([nombreCopia,datos_filtrados[0][4],arregloMatch[i],new Date(),nuevoArreglo.length]);

    } //aquì termina el for

imprimir();

}// aqui termina la funcion global


function imprimir(){
	Browser.msgBox("Fin de la función");

}

/**
function enviarPrueba(){
  enviarCorreo("TLAXCALA","valdes.gam@gmail.com",16IEXma1KmGLYshoUpoBrpBpgyXt8cV24);
}
**/
/*******************************funcion  enviar a correo ****************************/

function enviarCorreo(argumentoCarpetaEnviar,argumentoMail,argumentoIdCarpetaPrincipal) {

  var folderTraido=DriveApp.getFolderById(argumentoIdCarpetaPrincipal); //traer la carpeta principal
  
       
   var contents1 = folderTraido.getFolders();//traer los subfolders
   contador1=0;
      while (contents1.hasNext()) {
      var file1 = contents1.next(); //recorre los subfolders
      contador1++;

       data1 = [
            file1.getName(), //traer el nombre y el ID
            file1.getId(),
        ];

        /*imprimir los ID*/
        if(data1[0]==argumentoCarpetaEnviar){ //Si el nombre es igual al que se le pide
        //console.log(data1[0]);
        //console.log(data1[1]);
        //console.log(argumentoMail);
        var idSubfolder1=data1[1]; //entonces genera una nueva variable el id del subfolder
        }else{
          console.log("vale verga");
        }


    };


   var folder = DriveApp.getFolderById(idSubfolder1); //se va al subfolder
   var contents = folder.getFiles(); //traer todos los archivos

   

   var contador = 0;
   var file;


   var nuevoArreglo=[] //los va a poner en este arreglo


   while (contents.hasNext()) {
    var file = contents.next();
    contador++;

       data = [
            file.getName(),//traer su nombre y ID
            file.getId(),
        ];


        
        //console.log(data[1]);
        //nuevoArreglo.push(data[1]);
        
        var archivo1 = DriveApp.getFileById(data[1]);
        nuevoArreglo.push(archivo1); //los va adjuntando
    };



    var mensaje = "Estimados (as) Jefes de departamento \n\n Por medio de la presente, hago llegar las bitácoras correspondientes al mes de junio de su territorio. \n\n Lo anterior, con la finalidad de que puedan validarse y, en caso de no existir inconveniente, proceder a ser firmada por los usuarios. Es importante notificar si existe algún error, a fin de solventar en tiempo y forma el inconveniente. \n\n Cabe mencionar que dichas bitácoras deberán ser previamente firmadas y escaneadas en formato PDF y ser enviadas en los próximos 5 días hábiles después de la entrega, a través del siguiente link: \n\n Link de carga: \n\n  https://forms.gle/WfGNjUZTZk8aXqqH9 \n\n Agradecemos su atención, quedamos a sus órdenes. \n\n Saludos cordiales. \n\n En un bosque se bifurcaron dos caminos, y yo... Yo tomé el menos transitado. Esto marcó toda la diferencia. R.F"


    GmailApp.sendEmail(argumentoMail, "BITÁCORA O. CENTRAL JUNIO", mensaje,{attachments:nuevoArreglo});
    Browser.msgBox("Se ha enviado el correo");



   };












