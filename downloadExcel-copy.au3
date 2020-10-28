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
Local $pathDestino = "D:\Theeye\desktop\descargas\export"


if Not (IsArray($CmdLine)) Then
	  $status= "failure"
	 Local $mensajeError =  $fecha_ejecucion & " No se ingresaron los parametros de manera correcta"& @CRLF
	  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	 Exit
EndIf
if Not (UBound($CmdLine) > 2) Then
      	 $status= "failure"
		 Local $mensajeError =  $fecha_ejecucion & " Numero incorrecto de parametros"
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		Exit
EndIf

Local $fechaDesdePrmtr = $CmdLine[1]
Local $fechaHastaPrmtr = $CmdLine[2]


ConsoleWrite("parametros: "& $CmdLine[0] & $CmdLine[1])


If (FileExists($ejecucionLog)) Then
	  	FileWrite($ejecucionLog, $fecha_ejecucion & " Se ejecuto downloadExcel "& @CRLF)
 EndIf

;cierro todas las instancias de IE.
RunWait('taskkill /F /IM "iexplore.exe"')

;Ingreso al sitio
Local $oIE = _IECreate("http://bcweb11:90/tandem/login/source.asp")
;Send ("{F11}")

sleep(2000)

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

For $j=1 To 10

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


