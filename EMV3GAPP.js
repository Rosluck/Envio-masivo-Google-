function EnvioV3(){
var hojainfo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("informacion");
var hojaBD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD");
var wordID ="1-Th9iVqvQdDw7XpMXwbor1e8BKOUsqB2";
var plantillaID = "1371nKq3iG8FAg_ui6Cp2Jq4dFOG6kmPGjCceGC1oWgQ";
var carpetaword = DriveApp.getFolderById(wordID);

// LLAMADO DE VALORES CORREO ------------------------------------------------------------------------------------------------------------------------
  var plantillaA =  hojainfo.getRange('G1').getValue();   //ASUNTO 
  var plantillaM =  hojainfo.getRange('G2').getValue();    //MENSAJE
  var fecha = hojainfo.getRange('C2').getValue();

//-----------------------------------------------------------
var nombresEmpresas =[];
var cEmpresas = [];
var archivoPlantilla = [];


var copiarArchivoPlantilla = [];
var copiaID = [];

//Creaccion de valores
for (var reconocimiento=1; reconocimiento <= hojainfo.getLastRow(); reconocimiento++){
    const empresas = hojainfo.getRange(reconocimiento, 1).getValue();
    const correos = hojainfo.getRange(reconocimiento,2).getValue();

archivoPlantilla [reconocimiento]= DriveApp.getFileById(plantillaID);
  nombresEmpresas [reconocimiento] = empresas;
  cEmpresas[reconocimiento] = correos;
}

//creacion de archivos de envio de correo
for (var reconocimiento2= 1 ; reconocimiento2 < 92 ; reconocimiento2++){
     copiarArchivoPlantilla[reconocimiento2] = archivoPlantilla[reconocimiento2].makeCopy(carpetaword);
    copiaID[reconocimiento2] =copiarArchivoPlantilla[reconocimiento2].getId();
}




//Reconocimiento de la base de datos--------------------------------------------------------------------------------------------


var ingresoE = [];
var ordenE =[];
var rutE = [];
var pacienteE =[]
var procedenciaE = [];
var conveniosE = [];
var resultadoE =[];

var contadorgeneral = 1;

//PERMANENTE

var ingresoP = [];
var ordenP = [];
var rutP = [];
var pacienteP = [];
var procedenciaP = [];
var convenioP = [];
var resultadoP =[];

var contador2= 2;
for (var contador =1 ; contador < 93; contador++){
for (var reconocimiento3=1; reconocimiento3 <= hojaBD.getLastRow(); reconocimiento3++){

    const informacion = hojaBD.getRange(reconocimiento3, 16).getValue();

        if (informacion == nombresEmpresas[contador2]){

        var ingreso = hojaBD.getRange(reconocimiento3, 1).getDisplayValue();
        var orden = hojaBD.getRange(reconocimiento3, 5).getValue();
        var rut = hojaBD.getRange(reconocimiento3, 6).getValue();
        var paciente = hojaBD.getRange(reconocimiento3, 7).getValue();
        var procedencia = hojaBD.getRange(reconocimiento3, 15).getValue();
        var convenios = hojaBD.getRange(reconocimiento3, 16).getValue();
        var resultado = hojaBD.getRange(reconocimiento3, 20).getValue();
        


        ingresoE[contadorgeneral]= "\n"+ingreso;
        ordenE[contadorgeneral]="\n"+orden;
        rutE[contadorgeneral]="\n"+rut;
        pacienteE[contadorgeneral]="\n"+paciente;
        procedenciaE[contadorgeneral]="\n"+procedencia;
        conveniosE[contadorgeneral]="\n"+convenios;
        resultadoE[contadorgeneral]="\n"+resultado;
         
         contadorgeneral = contadorgeneral+1;



        }

} //FIN DE RECONOCIMIENTO
contador2 = contador2 + 1;
ingresoP[contador]=ingresoE;
ordenP[contador]=ordenE;
rutP[contador]=rutE;
pacienteP[contador]=pacienteE;
procedenciaP[contador]=procedenciaE;
convenioP[contador]=conveniosE;
resultadoP[contador]=resultadoE;

contadorgeneral = 1;
} //fin del contador




  Logger.log(copiarArchivoPlantilla);
  Logger.log(copiaID);
  Logger.log(ingresoP);
}
