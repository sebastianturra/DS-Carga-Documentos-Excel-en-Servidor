<?php

class Conexion {
    
    private $host = "localhost";
    private $user = "root";
    private $pass = "";
    private $bd = "irmaemp";
    public $con;
    
    /*
    private $host = "localhost";
    private $user = "root";
    private $pass = "";
    private $bd = "salitrera_irma";
*/
    public function conectar() {
        $con = new mysqli($this->host, $this->user, $this->pass, $this->bd);
        if ($con->connect_errno) {
            echo "Lo sentimos, este sitio web está experimentando problemas.";
            echo "Error: problema al conectarse a MySQL debido a: \n";
            echo "Errno: " . $con->connect_errno . "\n";
            echo "Error: " . $con->connect_error . "\n";
            exit;
        } else {
           // echo"Conexion OK";
        }
        $acentos = $con->query("SET NAMES 'utf8'");
        return $con;
    }

    public function desconectar() {
          
        $con = new mysqli($this->host, $this->user, $this->pass, $this->bd);
        if ($this->con) {
          echo "llego aca";
            mysql_close($this->$con);
        }else{
         // echo "estaba desconectado";
        }
    }

  /*  public function pruebadb(){
      $con = new mysqli($this->host, $this->user, $this->pass, $this->bd);
      $tabla = "usuarios";
      $sql="SELECT COUNT(*) FROM usuarios";
      $query = mysqli_query($con,$sql);
      if ($query == 0){
      echo "Sentencia incorrecta llamado a tabla: $query";
      }else {
      echo "Funciona";
      }
      }
   */  
}

//$prueba =new Conexion();
//$prueba ->pruebadb();
?>