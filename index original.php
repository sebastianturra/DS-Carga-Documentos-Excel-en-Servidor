<?php
include_once('Modelo/Conexion.php');
include_once('Modelo/Excelcarga.php');

$excelcarga = new excelcarga();

$lastvalue = $excelcarga->getlastvalue();

//SE CONECTA LA BASE DE DATOS
$enlace = mysqli_connect("localhost", "root", "", "irmaemp");

if (!$enlace) {
    echo "<br>Error: No se pudo conectar a MySQL.<br>";
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

if (isset($action)== "upload"){
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
                        $_DATOS_EXCEL[$a]['epru_id'] = $objPHPExcel->getActiveSheet()->getCell('A'.$i)->getCalculatedValue();
                        $_DATOS_EXCEL[$a]['epru_nom'] = $objPHPExcel->getActiveSheet()->getCell('B'.$i)->getCalculatedValue();
                        $_DATOS_EXCEL[$a]['epru_fecnac']= date('Y-m-d',PHPExcel_Shared_Date::ExcelToPHP($objPHPExcel->getActiveSheet()->getCell('C'.$i)->getCalculatedValue()));
                        $_DATOS_EXCEL[$a]['epru_numfav'] = $objPHPExcel->getActiveSheet()->getCell('D'.$i)->getCalculatedValue();
                        $_DATOS_EXCEL[$a]['epru_numfolio'] = $objPHPExcel->getActiveSheet()->getCell('E'.$i)->getCalculatedValue();
                        
                        $nombrepdf="FACT_".$_DATOS_EXCEL[$a]['epru_numfolio'].$_DATOS_EXCEL[$a]['epru_fecnac'].".pdf";
                       // echo $nombrepdf;
                        $_DATOS_EXCEL[$a]['epru_fecven']= date("Y-m-d", strtotime($_DATOS_EXCEL[$a]['epru_fecnac']."+ 1 month"));
                        $a++;
                    }       
                    $errores=0;

                   // var_dump($_DATOS_EXCEL);

//VARIABLE DE CAMPO                   
        $campo=0;     

//SE CREAN LAS SENTENCIAS SQL DE SUBIDA DE FACTURAS.
        for($i=0; $i< count($_DATOS_EXCEL); $i++){
          $sql = "INSERT INTO `excelprueba`(
            `epru_id`
          , `epru_nom`
          , `epru_fecnac`
          , `epru_numfav`
          , `epru_valdef`
          , `epru_numfolio`) VALUES ";
                  $sql.="('".$_DATOS_EXCEL[$i]['epru_id']."','";
                  $sql.=$_DATOS_EXCEL[$i]['epru_nom']."','";
                  $sql.=$_DATOS_EXCEL[$i]['epru_fecnac']."','";
                  $sql.=$_DATOS_EXCEL[$i]['epru_fecven']."','";
                  $sql.=$_DATOS_EXCEL[$i]['epru_numfav']."','";
                  $sql.="2','";
                  $sql.=$_DATOS_EXCEL[$i]['epru_numfolio']."')";

                  echo $sql;
                      
             $result = $enlace->query($sql);
             if (!$result){ echo "<br>Error al insertar registro<br>".$campo;$errores+=1;}
/////////////////////////////////////////////////////////////////////////   
          }
                    echo "<br><hr> <div class='col-xs-12'>
        <div class='form-group'>
          <strong><center>ARCHIVO IMPORTADO CON EXITO, EN TOTAL LOS REGISTROS Y $errores ERRORES</center></strong>
</div>
</div> <br> ";  
//Borramos el archivo que esta en el servidor con el prefijo cop_
//si por algun motivo no cargo el archivo cop_ 
                    unlink($destino);  
                }/*ESTE VA PARA EL FILE EXIST.*/
                else{
                    echo "<br>Primero debes cargar el archivo con extencion .xlisx<br>";
                }
            } //ESTE VA PARA UPLOAD.
        ?>