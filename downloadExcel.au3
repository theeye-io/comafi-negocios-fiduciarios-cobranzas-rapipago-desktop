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
local $seleccioneUbicacionPROD = "D:\Theeye\desktop\images\seleccioneUbicacionPROD.bmp"


Local $status= ""

Local $ingresar = "D:\Theeye\desktop\images\ingresar.bmp"
Local $iteracionesAdmTerceros = 10
Local $seleccioneUbicacion = "D:\Theeye\desktop\images\seleccioneUbicacion.bmp"
$jsonFilePath = "D:\theeye\desktop\data\estructuraPagosEmerix.json"

if Not FileExists($jsonFilePath) Then
	  $status = "failure"
	  Local $mensajeError =  $fecha_ejecucion & " No se encuentro el JSON solicitado"
	  $mensajeError = StringReplace($mensajeError,"\" ,"/")
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	  Exit
 EndIf

$JsonPlainFile = FileRead($jsonFilePath)
if @error Then
   $status= "failure"
   Local $mensajeError =  $fecha_ejecucion & " Error de lectura del archivo "&$jsonFilePath&
   $mensajeError = StringReplace($mensajeError,"\" ,"/")
   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
   Exit
EndIf
Local $oStatusBases = Json_decode($JsonPlainFile)
Local $opagos_emerixArray = Json_get($oStatusBases,'["pagos_emerix"]')



if Not (IsArray($CmdLine)) Then
	  $status= "failure"
	 Local $mensajeError =  $fecha_ejecucion & " No se ingresaron los parametros de manera correcta"& @CRLF
	 $mensajeError = StringReplace($mensajeError,"\" ,"/")
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	 Exit
EndIf
if Not (UBound($CmdLine) > 4) Then
      	 $status= "failure"
		 Local $mensajeError =  $fecha_ejecucion & " Numero incorrecto de parametros"
		 $mensajeError = StringReplace($mensajeError,"\" ,"/")
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		Exit
EndIf

Local $fechaDesdePrmtr = $CmdLine[1]
Local $fechaHastaPrmtr = $CmdLine[2]
Local $EntornoPrmtr = $CmdLine[3]
Local $pathDestino = $CmdLine[4]

ToolTip($pathDestino&"\emerix"&@MDAY & @MON & @YEAR&".xls",0,0)
Sleep(1000)


Local $user = "user1" ; DEV
Local $pass = "user1" ; DEV



If (FileExists($ejecucionLog)) Then
	  	FileWrite($ejecucionLog, $fecha_ejecucion & " Se ejecuto downloadExcel "& @CRLF)
 EndIf

;cierro todas las instancias de IE.
RunWait('taskkill /F /IM "iexplore.exe"')

;Ingreso al sitio

   If $EntornoPrmtr = "DEV" Then
	  $urlEntorno = "http://dstst04"
   EndIf
   If $EntornoPrmtr = "PROD" Then
	  $urlEntorno = "http://bcweb11:90"
   EndIf


sleep(6000)
if $EntornoPrmtr = "DEV" Then
   Local $oIE = _IECreate("http://dstst04/tandem/login/source.asp")
   _IELoadWait($oIE)
   WinSetState(_IEPropertyGet($oIE, "frm"), "", @SW_MAXIMIZE)
   Sleep(4000)
  $oIE = _IEAttach("Seleccione")
   If WinExists("Seleccione") Then
	  WinActivate("Seleccione")
	  ToolTip("Ventana de Login encontrada",0,0)
	  Sleep(2000)
   Else
	  Return False
   EndIf

   Sleep(5000)
   if imageExists($seleccioneUbicacion, 10) then
		 Send("{F11}")
   EndIf
   if imageExists($imgSolArgentina, 10) then
	  clickOnImage($imgSolArgentina)
	  if @error Then
		   FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$imgSolArgentina& @CRLF)
			$status= "failure"
			Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$imgSolArgentina& @CRLF
			$mensajeError = StringReplace($mensajeError,"\" ,"/")
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		   RunWait('taskkill /F /IM "iexplore.exe"')
		   Exit
	  EndIf
	  FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso al sitio de Emerix"& @CRLF)
   Else
	   FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$imgSolArgentina& " no encontro el elemento para hacer click "& @CRLF)
	   $status= "failure"
	   Local $mensajeError = $fecha_ejecucion & " la imagen "&$imgSolArgentina& " no encontro el elemento para hacer click "
	   $mensajeError = StringReplace($mensajeError,"\" ,"/")
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
   WinSetState(_IEPropertyGet($oIE, "frm"), "", @SW_MAXIMIZE)
   _IELoadWait($oIE)
   Sleep(4000)
  $oIE = _IEAttach("Seleccione")
   If WinExists("Seleccione") Then
	  WinActivate("Seleccione")
	  ToolTip("Ventana de Login encontrada",0,0)
	  Sleep(2000)
   EndIf
   MouseMove(716,478)
   MouseClick($MOUSE_CLICK_LEFT)
      Sleep(5000)
   if imageExists($seleccioneUbicacionPROD, 10) then
	  Send("{F11}")
   EndIf
   if imageExists($imgSolArgentina, 10) then
	  clickOnImage($imgSolArgentina)
	  Sleep(5000)

	  if @error Then
		   FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$imgSolArgentina& @CRLF)
			$status= "failure"
			Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$imgSolArgentina& @CRLF
			$mensajeError = StringReplace($mensajeError,"\" ,"/")
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		   RunWait('taskkill /F /IM "iexplore.exe"')
		   Exit
	  EndIf
	  FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso al sitio de Emerix"& @CRLF)
   Else
	   FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$imgSolArgentina& " no encontro el elemento para hacer click "& @CRLF)
	   $status= "failure"
	   Local $mensajeError = $fecha_ejecucion & " la imagen "&$imgSolArgentina& " no encontro el elemento para hacer click "
	   $mensajeError = StringReplace($mensajeError,"\" ,"/")
	   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	   RunWait('taskkill /F /IM "iexplore.exe"')
	   Exit
	EndIf
   sleep(4000)
EndIf

ToolTip($urlEntorno,0,0)

$oIE = _IEAttach("Revisión de Cuentas")
Local $titles = _IETagNameGetCollection($oIE, "a")
Local $ingresarText
For $title in $titles
   $href_value = $title.GetAttribute("href")
   if($href_value = $urlEntorno&"/tandem/operacional/addins/circuitos_adm/int_ng_cobros_adm_pmafi.asp") Then
	  $ingresarText = $title
   ExitLoop
   EndIf
Next
_IEAction($ingresarText, "click")
sleep(4000)



if imageExists($fechadesde, 10) then
   clickOnImage($fechadesde)
   if @error Then
		 FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$fechadesde& @CRLF)
	     $status= "failure"
		 Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$fechadesde& @CRLF
		 $mensajeError = StringReplace($mensajeError,"\" ,"/")
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
	$mensajeError = StringReplace($mensajeError,"\" ,"/")
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
		 $mensajeError = StringReplace($mensajeError,"\" ,"/")
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
	$mensajeError = StringReplace($mensajeError,"\" ,"/")
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
		 $mensajeError = StringReplace($mensajeError,"\" ,"/")
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
	$mensajeError = StringReplace($mensajeError,"\" ,"/")
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
		 $mensajeError = StringReplace($mensajeError,"\" ,"/")
	   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')

		RunWait('taskkill /F /IM "iexplore.exe"')
		Exit
   EndIf
Else
	FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$descarga& " no encontro el elemento para hacer click "& @CRLF)
		$status= "failure"
	  Local $mensajeError =  $fecha_ejecucion & " la imagen "&$descarga& " no encontro el elemento para hacer click "
	  $mensajeError = StringReplace($mensajeError,"\" ,"/")
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
 EndIf
Sleep(1000)
MsgBox($MB_SYSTEMMODAL, "Title","sarasa", 1)

Sleep(2000)

if imageExists($descargaexcel, 4) then
   clickOnImage($descargaexcel)
   if @error Then
	  	FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$descargaexcel& @CRLF)
		 $status= "failure"
		 Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$descargaexcel
		 $mensajeError = StringReplace($mensajeError,"\" ,"/")
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		RunWait('taskkill /F /IM "iexplore.exe"')
		Exit
   EndIf
Else
	FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$descargaexcel& " no encontro el elemento para hacer click "& @CRLF)
	 $status= "failure"
	  Local $mensajeError =  $fecha_ejecucion & " la imagen "&$descargaexcel& " no encontro el elemento para hacer click "
	  $mensajeError = StringReplace($mensajeError,"\" ,"/")
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
 EndIf

ConsoleWrite("Esperando que se pueda guardar el archivo"& @CRLF)
;sleep (60000)
Local $contador = 1
while Not imageExists($flechaGuardar, 10)
   Sleep(1000)
   if $contador > 30 Then
	  $status= "failure"
	  Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$flechaGuardar
	  $mensajeError = StringReplace($mensajeError,"\" ,"/")
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
		 $mensajeError = StringReplace($mensajeError,"\" ,"/")
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		RunWait('taskkill /F /IM "iexplore.exe"')
		Exit
   EndIf
	send("{DOWN}")
    send("{ENTER}")
	sleep(3000)
	send($pathDestino&"\emerix"&@MDAY &"-"& @MON&"-"& @YEAR&".xls")
    send("{ENTER}")

Else
	FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$flechaGuardar& " no encontro el elemento para hacer click "& @CRLF)
	 $status= "failure"
	  Local $mensajeError =  $fecha_ejecucion & " la imagen "&$flechaGuardar& " no encontro el elemento para hacer click "
	  $mensajeError = StringReplace($mensajeError,"\" ,"/")
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
 EndIf

sleep(2000)
if imageExists($deseareemplazarlo, 10) then
   send("{LEFT}")
   send("{ENTER}")
EndIf
sleep(5000)

 Local $iFileExists = FileExists($pathDestino&"\emerix"&@MDAY&"-" & @MON&"-" & @YEAR&".xls")
 ToolTip($iFileExists,0,0)
 Sleep(10000)

if $iFileExists == 1 Then
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
	  $mensajeError = StringReplace($mensajeError,"\" ,"/")
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	RunWait('taskkill /F /IM "iexplore.exe"')
	Exit
EndIf


