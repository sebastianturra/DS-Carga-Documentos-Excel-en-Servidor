<?php
include_once('Modelo/Conexion.php');
include_once('Modelo/Excelcarga.php');

$excelcarga = new excelcarga();

$lastvalue = $excelcarga->getlastvalue();

//SE CONECTA LA BASE DE DATOS
$enlace = mysqli_connect("localhost", "root", "", "irmaemp");

if (!$enlace) {
  echo "<div display='none'>
    <script type='text/javascript'>
        console.log('<br>Error: No se pudo conectar a MySQL.<br>');
    </script>
</div>";
    exit;
}

//ESCRIBE EN LA CONSOLA QUE SE CONECTO CORRECTAMENTE.
echo "<div display='none'>
    <script type='text/javascript'>
        console.log('<br>Éxito: Se realizó una conexión apropiada a MySQL! La base de datos mi_bd es genial.<br>');
    </script>
</div>";

?>

<div class="container">
<h2>Cargar e importar archivo excel a MySQL</h2>
<form name="importa" method="post" action="" enctype="multipart/form-data" >
  <div class="col-xs-4">
    <div class="form-group">
      <input type="file" class="filestyle" data-buttonText="Seleccione archivo" name="excel">
    </div>
  </div>
  <div class="col-xs-2">
    <input class="btn btn-default btn-file" type='submit' name='enviar'  value="Importar"  />
  </div>
  <input type="hidden" value="upload" name="action" />
  <input type="hidden" value="usuarios" name="mod">
  <input type="hidden" value="masiva" name="acc">
</form>
</div>

<?php 
extract($_POST);
if (isset($_POST['action'])) {
$action=$_POST['action'];
}

if (isset($action) == "upload"){
//SE CARGA EL ARCHIVO EXCEL
$archivo = $_FILES['excel']['name'];
$tipo = $_FILES['excel']['type'];
//SE AGREGA UN PREFIJO PARA IDENTIFICARLO.
$destino = "cop_".$archivo;
if (copy($_FILES['excel']['tmp_name'],$destino)) echo "<br>Archivo Cargado Con Éxito<br>";
else echo "<br>Error Al Cargar el Archivo<br>";
        
if (file_exists ("cop_".$archivo)){ 
// Llamamos las clases necesarias PHPEcel 
require_once('phpexcel/Classes/PHPExcel.php');
require_once('phpexcel/Classes/PHPExcel/Reader/Excel2007.php');                  
// Cargando la hoja de excel
$objReader = new PHPExcel_Reader_Excel2007();
$objPHPExcel = $objReader->load("cop_".$archivo);
$objFecha = new PHPExcel_Shared_Date();       
// Asignamon la hoja de excel activa
$objPHPExcel->setActiveSheetIndex(0);

//SE OBTIENE LA FILA Y LA COLUMNA MAS ALTA.
$columnas = $objPHPExcel->setActiveSheetIndex(0)->getHighestColumn();
$filas = $objPHPExcel->setActiveSheetIndex(0)->getHighestRow();

//Creamos un array con todos los datos del Excel importado
$a=0;

for ($i=2;$i<=$filas;$i++){
                        
                        $_DATOS_EXCEL[$a]['FACT_CODIGO'] = $objPHPExcel->getActiveSheet()->getCell('F'.$i)->getCalculatedValue();
                        $_DATOS_EXCEL[$a]['ESTF_CODIGO']= '3';
                        $_DATOS_EXCEL[$a]['TIPF_CODIGO'] = '1';
                        $_DATOS_EXCEL[$a]['TIPS_CODIGO'] = '1';
                        $_DATOS_EXCEL[$a]['FACT_FECHAING'] = $objPHPExcel->getActiveSheet()->getCell('G'.$i)->getCalculatedValue();
                        $_DATOS_EXCEL[$a]['FACT_FECHAING'] = date('Y-m-d',PHPExcel_Shared_Date::ExcelToPHP($objPHPExcel->getActiveSheet()->getCell('G'.$i)->getCalculatedValue()));
                        ////////////////////////////////////////////////////////////////////////////
//NOMBRE ARCHIVO PDF DE LA COTIZACION
                        /*$nombrepdf="FACT_".$_DATOS_EXCEL[$a]['FACT_CODIGO'].$_DATOS_EXCEL[$a]['FACT_FECHAING'].".pdf";*/
                        $_DATOS_EXCEL[$a]['ARCHF_NOMBRE'] = "FACT_".$_DATOS_EXCEL[$a]['FACT_CODIGO'].$_DATOS_EXCEL[$a]['FACT_FECHAING'].".pdf";
                        $_DATOS_EXCEL[$a]['ARCHF_USERNOM'] = "";
                        $_DATOS_EXCEL[$a]['ARCHF_FECHASUBIDA'] = $_DATOS_EXCEL[$a]['FACT_FECHAING'];
                      //  echo "<br>".$_DATOS_EXCEL[$a]['ARCHF_NOMBRE']."<br>";
//FECHA DE VENCIMIENTO   
                        $_DATOS_EXCEL[$a]['FECHA_PAGO_VENCIMIENTO']= date("Y-m-d", strtotime($_DATOS_EXCEL[$a]['FACT_FECHAING']."+ 1 month"));
                        ////////////////////////////////////////////////////////////////////////////
                        $_DATOS_EXCEL[$a]['FACT_NUMORDEN'] = '-';
                        $_DATOS_EXCEL[$a]['FACT_VNETO']= $objPHPExcel->getActiveSheet()->getCell('L'.$i)->getCalculatedValue();
                        $_DATOS_EXCEL[$a]['FACT_IVA'] = $objPHPExcel->getActiveSheet()->getCell('M'.$i)->getCalculatedValue();
                        $_DATOS_EXCEL[$a]['FACT_TOTAL'] = $objPHPExcel->getActiveSheet()->getCell('N'.$i)->getCalculatedValue();
                        $_DATOS_EXCEL[$a]['FACT_DESCRIPCION'] = '-';
                        $_DATOS_EXCEL[$a]['FACT_EXCENTO'] = 'NO';
                        $_DATOS_EXCEL[$a]['FACT_RUTRS']= $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getCalculatedValue();
                        $_DATOS_EXCEL[$a]['FACT_LUGAR'] = '-';
                        $_DATOS_EXCEL[$a]['FACT_NOMBRERS']= $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getCalculatedValue();
                        $_DATOS_EXCEL[$a]['FACT_EMIREC'] = 'EMITIDA';
                        $_DATOS_EXCEL[$a]['FACT_CONTACTO'] = '-';
                        $_DATOS_EXCEL[$a]['FACT_CORREO'] = '-';
                        $_DATOS_EXCEL[$a]['FACT_FONO'] = '0';
                        $_DATOS_EXCEL[$a]['FACT_FVENC']= $_DATOS_EXCEL[$a]['FECHA_PAGO_VENCIMIENTO'];
                        $_DATOS_EXCEL[$a]['FACT_FPAG'] = $_DATOS_EXCEL[$a]['FECHA_PAGO_VENCIMIENTO'];
                        $_DATOS_EXCEL[$a]['FACT_FORMPAG'] = 'CRED';
                        $_DATOS_EXCEL[$a]['FACT_COBRANZA'] = $_DATOS_EXCEL[$a]['FECHA_PAGO_VENCIMIENTO'];
          
                        $a++;

                    }       
                    $errores=0;
                   /* var_dump($_DATOS_EXCEL); */

$fact_id=$lastvalue;
for($i=0; $i<count($_DATOS_EXCEL); $i++){
  $numerofolio = $excelcarga->idverification($_DATOS_EXCEL[$i]['FACT_CODIGO']);
  if($numerofolio == 1){
    echo "<br>Ese folio NO esta registrado: ".$_DATOS_EXCEL[$i]['FACT_CODIGO']."<br>";
    /*$fact_id=999;*/
    $_DATOS_EXCEL[$i]['FACT_ID'] = $fact_id;

//VARIABLE DE CAMPO                   
        $campo=0;     

//SE CREAN LAS SENTENCIAS SQL DE SUBIDA DE FACTURAS
          $ingfac = "INSERT INTO `facturacion`(
          `FACT_ID`, 
          `FACT_CODIGO`, 
          `ESTF_CODIGO`, 
          `TIPF_CODIGO`, 
          `TIPS_CODIGO`, 
          `FACT_FECHAING`, 
          `FACT_NUMORDEN`, 
          `FACT_VNETO`, 
          `FACT_IVA`, 
          `FACT_TOTAL`, 
          `FACT_DESCRIPCION`, 
          `FACT_EXCENTO`, 
          `FACT_RUTRS`, 
          `FACT_LUGAR`, 
          `FACT_NOMBRERS`, 
          `FACT_EMIREC`, 
          `FACT_CONTACTO`, 
          `FACT_CORREO`, 
          `FACT_FONO`, 
          `FACT_FVENC`, 
          `FACT_FPAG`, 
          `FACT_FORMPAG`, 
          `FACT_COBRANZA`) 
          VALUES ";

                  $ingfac.="('".$_DATOS_EXCEL[$i]['FACT_ID']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_CODIGO']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['ESTF_CODIGO']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['TIPF_CODIGO']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['TIPS_CODIGO']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_FECHAING']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_NUMORDEN']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_VNETO']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_IVA']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_TOTAL']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_DESCRIPCION']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_EXCENTO']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_RUTRS']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_LUGAR']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_NOMBRERS']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_EMIREC']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_CONTACTO']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_CORREO']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_FONO']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_FVENC']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_FPAG']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_FORMPAG']."','";
                  $ingfac.=$_DATOS_EXCEL[$i]['FACT_COBRANZA']."')";

                //  echo "<br>".$ingfac."<br>";
                      
           $result = $enlace->query($ingfac);
             if (!$result){ echo "<div display='none'>
              <script type='text/javascript'>
                  console.log('<br>Error al insertar registro factura<br>');
              </script>
          </div>".$campo;$errores+=1;}     

/////////////////////////////////////////////////////////////////////////////////////////////
  $ingarcfac = "INSERT INTO `archivos_facturas`(
    `FACT_ID`, 
    `ARCHF_NOMBRE`, 
    `ARCHF_USERNOM`, 
    `ARCHF_FECHASUBIDA`) 
    VALUES";

$ingarcfac.="('".$_DATOS_EXCEL[$i] ['FACT_ID']."','";      
$ingarcfac.=$_DATOS_EXCEL[$i]['ARCHF_NOMBRE']."','";
$ingarcfac.=$_DATOS_EXCEL[$i]['ARCHF_USERNOM']."','";
$ingarcfac.=$_DATOS_EXCEL[$i]['ARCHF_FECHASUBIDA']."')";

//  echo "<br>".$ingarcfac."<br>";

$result = $enlace->query($ingarcfac);
if (!$result){ echo "<div display='none'>
      <script type='text/javascript'>
          console.log('<br>Error al insertar archivo factura<br>');
      </script></div>".$campo;$errores+=1;}   
     
     $fact_id++;

    }
  }
   
///////////////////////////////////////////////////////////////////////////////////
      echo "<br><hr> <div class='col-xs-12'>
        <div class='form-group'>
          <strong><center>ARCHIVO IMPORTADO CON EXITO, EN TOTAL LOS REGISTROS Y $errores ERRORES</center></strong>
        </div>
      </div> 
      <br> ";  
      echo "<meta http-equiv='refresh' content='4; url=index.php'>";

//Borramos el archivo que esta en el servidor con el prefijo cop_
//si por algun motivo no cargo el archivo cop_ 
                    unlink($destino);  
                }/*ESTE VA PARA EL FILE EXIST.*/
                else{
                    echo "<br>Primero debes cargar el archivo con extencion .xlisx<br>";
                    echo "<meta http-equiv='refresh' content='4; url=index.php'>";
                }
            } //ESTE VA PARA UPLOAD.
?>