#cs ----------------------------------------------------------------------------
AutoIt Version: 3.3.14.5
Author:         Agustin Moreno
Script Function:
 Template AutoIt script.
#ce ----------------------------------------------------------------------------
; Script Start - Add your code below here
#include <IE.au3>
#include <ScreenCapture.au3>
#include <Misc.au3>
#include <Array.au3>
#include <Date.au3>
#include <MsgBoxConstants.au3>
#include "Functions.au3"
#include "Json.au3"
#include <ImageSearch.au3>


Local $error_Log = "D:\theeye\desktop\log\error.log"
Local $ejecucionLog = "D:\theeye\desktop\log\ejecucionLog.log"
Local $fecha_ejecucion = @MDAY & "-" & @MON & "-" & @YEAR &"  "& @HOUR & ":" & @MIN & ":" & @SEC
Local $imgSolArgentina = "D:\Theeye\desktop\images\solargentina.bmp"
Local $admCamTerceros = "D:\Theeye\desktop\images\amd1.bmp"
Local $fechadesde = "D:\Theeye\desktop\images\fechadesde.bmp"
Local $fechahasta = "D:\Theeye\desktop\images\fechahasta.bmp"
Local $descarga = "D:\Theeye\desktop\images\descarga.bmp"
Local $descargaexcel = "D:\Theeye\desktop\images\descargaexcel1.bmp"
Local $guardar = "D:\Theeye\desktop\images\guardar.bmp"
Local $aplicarfiltro = "D:\Theeye\desktop\images\aplicarfiltro.bmp"
Local $yesterday = @MDAY - 1&"/"& @MON&"/"&@YEAR
local $flechaGuardar = "D:\Theeye\desktop\images\flechaguardar.bmp"
local $deseareemplazarlo = "D:\Theeye\desktop\images\deseareemplazarlo.bmp"
Local $status= ""
Local $pathDestino = "D:\theeye\conciliacion\files";"D:\Theeye\desktop\descargas\export"
Local $ingresar = "D:\Theeye\desktop\images\ingresar.bmp"
Local $iteracionesAdmTerceros = 10
Local $seleccioneUbicacion = "D:\Theeye\desktop\images\seleccioneUbicacion.bmp"
$jsonFilePath = "D:\theeye\desktop\data\estructuraPagosEmerix.json"

if Not FileExists($jsonFilePath) Then
	  $status = "failure"
	  Local $mensajeError =  $fecha_ejecucion & " No se encuentro el JSON solicitado"
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	  Exit
 EndIf

$JsonPlainFile = FileRead($jsonFilePath)
if @error Then
   $status= "failure"
   Local $mensajeError =  $fecha_ejecucion & " Error de lectura del archivo "&$jsonFilePath&
   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
   Exit
EndIf
Local $oStatusBases = Json_decode($JsonPlainFile)
Local $opagos_emerixArray = Json_get($oStatusBases,'["pagos_emerix"]')



if Not (IsArray($CmdLine)) Then
	  $status= "failure"
	 Local $mensajeError =  $fecha_ejecucion & " No se ingresaron los parametros de manera correcta"& @CRLF
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	 Exit
EndIf
if Not (UBound($CmdLine) > 3) Then
      	 $status= "failure"
		 Local $mensajeError =  $fecha_ejecucion & " Numero incorrecto de parametros"
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		Exit
EndIf

Local $fechaDesdePrmtr = $CmdLine[1]
Local $fechaHastaPrmtr = $CmdLine[2]
Local $EntornoPrmtr = $CmdLine[3]




Local $user = "user1"
Local $pass = "user1"



If (FileExists($ejecucionLog)) Then
	  	FileWrite($ejecucionLog, $fecha_ejecucion & " Se ejecuto downloadExcel "& @CRLF)
 EndIf

;cierro todas las instancias de IE.
RunWait('taskkill /F /IM "iexplore.exe"')

;Ingreso al sitio

    if imageExists($seleccioneUbicacion, 10) then
		 Send("{F11}")
   EndIf


sleep(6000)
if $EntornoPrmtr = "DEV" Then
   Local $oIE = _IECreate("http://dstst04/tandem/login/source.asp")
   if imageExists($imgSolArgentina, 10) then
	  clickOnImage($imgSolArgentina)
	  if @error Then
		   FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$imgSolArgentina& @CRLF)
			$status= "failure"
			Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$imgSolArgentina& @CRLF
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		   RunWait('taskkill /F /IM "iexplore.exe"')
		   Exit
	  EndIf
	  FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso al sitio de Emerix"& @CRLF)
   Else
	   FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$imgSolArgentina& " no encontro el elemento para hacer click "& @CRLF)
	   $status= "failure"
	   Local $mensajeError = $fecha_ejecucion & " la imagen "&$imgSolArgentina& " no encontro el elemento para hacer click "
	   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	   RunWait('taskkill /F /IM "iexplore.exe"')
	   Exit
	EndIf
   sleep(1000)


   send($user)
   send ("{TAB}")
   send($pass)
   send ("{TAB}")
   send ("{ENTER}")

   sleep(4000)
   $iteracionesAdmTerceros=41
ElseIf $EntornoPrmtr = "PROD" Then
   Local $oIE = _IECreate("http://bcweb11:90/tandem/login/source.asp")
   if imageExists($imgSolArgentina, 10) then
	  clickOnImage($imgSolArgentina)
	  Sleep(5000)
	  if imageExists($seleccioneUbicacion, 10) then
		 Send("{F11}")
	  EndIf
	  if @error Then
		   FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$imgSolArgentina& @CRLF)
			$status= "failure"
			Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$imgSolArgentina& @CRLF
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		   RunWait('taskkill /F /IM "iexplore.exe"')
		   Exit
	  EndIf
	  FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso al sitio de Emerix"& @CRLF)
   Else
	   FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$imgSolArgentina& " no encontro el elemento para hacer click "& @CRLF)
	   $status= "failure"
	   Local $mensajeError = $fecha_ejecucion & " la imagen "&$imgSolArgentina& " no encontro el elemento para hacer click "
	   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	   RunWait('taskkill /F /IM "iexplore.exe"')
	   Exit
	EndIf
   sleep(4000)
EndIf

For $j=1 To $iteracionesAdmTerceros

 send ("{TAB}")

 sleep (250)

Next

sleep (2000)
send ("{ENTER}")
sleep (2000)



if imageExists($fechadesde, 10) then
   clickOnImage($fechadesde)
   if @error Then
		 FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$fechadesde& @CRLF)
	     $status= "failure"
		 Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$fechadesde& @CRLF
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		RunWait('taskkill /F /IM "iexplore.exe"')
		Exit
   EndIf
   FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso a adm de canales de terceros "& @CRLF)
   send ("{TAB}")
   send($fechaDesdePrmtr)
Else
	FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$fechadesde& " no encontro el elemento para hacer click "& @CRLF)
	$status= "failure"
    Local $mensajeError = $fecha_ejecucion & " la imagen "&$fechadesde& " no encontro el elemento para hacer click "
    ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
 EndIf



 if imageExists($fechahasta, 10) then
   clickOnImage($fechahasta)
   if @error Then
		 FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$fechahasta& @CRLF)
	     $status= "failure"
		 Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$fechahasta
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		RunWait('taskkill /F /IM "iexplore.exe"')
		Exit
   EndIf
   FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso a adm de canales de terceros "& @CRLF)
   send ("{TAB}")
   send($fechaHastaPrmtr)
Else
	FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$fechadesde& " no encontro el elemento para hacer click "& @CRLF)
	$status= "failure"
    Local $mensajeError = $fecha_ejecucion & " la imagen "&$fechadesde& " no encontro el elemento para hacer click "
    ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
 EndIf




sleep (2000)


if imageExists($aplicarfiltro, 10) then
   clickOnImage($aplicarfiltro)
   if @error Then
	  	FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$aplicarfiltro& @CRLF)
		$status= "failure"
		 Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$aplicarfiltro
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		RunWait('taskkill /F /IM "iexplore.exe"')
		Exit
   EndIf
   FileWrite($ejecucionLog, $fecha_ejecucion & " Se aplico el filtro "& @CRLF)
   send ("{TAB}")
   send($fechaDesde)
Else
	FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$aplicarfiltro& " no encontro el elemento para hacer click "& @CRLF)
	$status= "failure"
    Local $mensajeError = $fecha_ejecucion & " la imagen "&$aplicarfiltro& " no encontro el elemento para hacer click "
	ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
 EndIf

sleep (2000)


if imageExists($descarga, 10) then
   clickOnImage($descarga)
   if @error Then
	  	FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$descarga& @CRLF)
		$status= "failure"
		 Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$descarga
	   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')

		RunWait('taskkill /F /IM "iexplore.exe"')
		Exit
   EndIf
Else
	FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$descarga& " no encontro el elemento para hacer click "& @CRLF)
		$status= "failure"
	  Local $mensajeError =  $fecha_ejecucion & " la imagen "&$descarga& " no encontro el elemento para hacer click "
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
 EndIf
Sleep(1000)
MsgBox($MB_SYSTEMMODAL, "Title","sarasa", 1)

if imageExists($descargaexcel, 4) then
   clickOnImage($descargaexcel)
   if @error Then
	  	FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$descargaexcel& @CRLF)
		 $status= "failure"
		 Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$descargaexcel
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		RunWait('taskkill /F /IM "iexplore.exe"')
		Exit
   EndIf
Else
	FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$descargaexcel& " no encontro el elemento para hacer click "& @CRLF)
	 $status= "failure"
	  Local $mensajeError =  $fecha_ejecucion & " la imagen "&$descargaexcel& " no encontro el elemento para hacer click "
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
 EndIf

ConsoleWrite("Esperando que se pueda guardar el archivo"& @CRLF)
;sleep (60000)

while Not imageExists($flechaGuardar, 10)
   Sleep(1000)
   Local $contador = 1
   if $contador > 120 Then
	  $status= "failure"
	  Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$flechaGuardar
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	  RunWait('taskkill /F /IM "iexplore.exe"')
	  Exit
   EndIf
   $contador = $contador+1
WEnd


if imageExists($flechaGuardar, 10) then
   clickOnImage($flechaGuardar)
   if @error Then
	  	FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$flechaGuardar& @CRLF)
		 $status= "failure"
		 Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$flechaGuardar
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		RunWait('taskkill /F /IM "iexplore.exe"')
		Exit
   EndIf
	send("{DOWN}")
    send("{ENTER}")
	sleep(3000)
	send($pathDestino&@MDAY & "-" & @MON & "-" & @YEAR&".xls")
    send("{ENTER}")

Else
	FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$flechaGuardar& " no encontro el elemento para hacer click "& @CRLF)
	 $status= "failure"
	  Local $mensajeError =  $fecha_ejecucion & " la imagen "&$flechaGuardar& " no encontro el elemento para hacer click "
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
 EndIf

sleep(2000)
if imageExists($deseareemplazarlo, 10) then
   send("{LEFT}")
   send("{ENTER}")
EndIf
sleep(2000)

 Local $iFileExists = FileExists($pathDestino&@MDAY & "-" & @MON & "-" & @YEAR&".xls")

if $iFileExists Then
   FileWrite($ejecucionLog, $fecha_ejecucion & " Se logro descargar el archivo xls en la carpeta de descargas "& @CRLF)
   $status = "success"
   Local $mensaje = $fecha_ejecucion & " Se logro descargar el archivo xls en la carpeta de descargas "
   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensaje & '"]}')
   	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
Else
   	FileWrite($error_Log, $fecha_ejecucion & " no se logro encontrar el archivo "& @CRLF)
	$status= "failure"
	  Local $mensajeError =  $fecha_ejecucion & " no se logro encontrar el archivo "
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
EndIf


