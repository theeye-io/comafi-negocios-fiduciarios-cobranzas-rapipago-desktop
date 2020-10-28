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

;AutoIt3.exe D:\Theeye\desktop\downloadExcel3.au3 06/10/2020 16/10/2020 DEV


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
Local $ingresar = "D:\Theeye\desktop\images\ingresar.bmp"
Local $nrodocumento = "D:\Theeye\desktop\images\nrodocumento.bmp"
Local $noseencontrarondatos = "D:\Theeye\desktop\images\noseencontrarondatos.bmp"
Local $acreditado = "D:\Theeye\desktop\images\acreditado.bmp"
Local $distribucion = "D:\Theeye\desktop\images\distribucion.bmp"
Local $nocontribuirpagoanulado = "D:\Theeye\desktop\images\nocontribuirpagoanulado.bmp"
Local $agenciaNFComafi = "D:\Theeye\desktop\images\agenciaNFComafi.bmp"
Local $cambiar = "D:\Theeye\desktop\images\cambiar.bmp"
Local $flechitaSelectAgencia = "D:\Theeye\desktop\images\flechitaSelectAgencia.bmp"
Local $flechitaSelectTipoPago = "D:\Theeye\desktop\images\flechitaSelectTipoPago.bmp"
Local $cambiar2 = "D:\Theeye\desktop\images\cambiar2.bmp"
Local $pagoParcial = "D:\Theeye\desktop\images\pagoparcial.bmp"
Local $NumeroDocumento = "D:\Theeye\desktop\images\nrodocumento.bmp"
Local $nfComafi = "D:\Theeye\desktop\images\nfComafi.bmp"
Local $nfComafi2 = "D:\Theeye\desktop\images\nfComafi2.bmp"
Local $96 = "D:\Theeye\desktop\images\96.bmp"
Local $962 = "D:\Theeye\desktop\images\962.bmp"
Local $cuentaLockeada = "D:\Theeye\desktop\images\cuentaLockeada.bmp"
Local $cotinuarCuentaLockeada = "D:\Theeye\desktop\images\cotinuarCuentaLockeada.bmp"
Local $addins = "D:\Theeye\desktop\images\addins.bmp"
Local $cobros = "D:\Theeye\desktop\images\cobros.bmp"
Local $estadoIngresado = "D:\Theeye\desktop\images\estadoIngresado.bmp"
Local $checkSelect = "D:\Theeye\desktop\images\checkSelect.bmp"
Local $cambioDeEstado = "D:\Theeye\desktop\images\cambioEstadoGrabado.bmp"
Local $aceptar = "D:\Theeye\desktop\images\aceptar.bmp"
Local $ingresoDeCobros = "D:\Theeye\desktop\images\ingresoDeCobros.bmp"
Local $cliente = "D:\Theeye\desktop\images\cliente.bmp"
Local $tipoDePago = "D:\Theeye\desktop\images\tipoDePago.bmp"
Local $pagoparcialselect = "D:\Theeye\desktop\images\pagoparcialselect.bmp"
Local $importeCobroCancelacion = "D:\Theeye\desktop\images\importeCobroCancelacion.bmp"
Local $importeCobroPagoParcial = "D:\Theeye\desktop\images\importeCobroPagoParcial.bmp"
Local $pagoCancelacionCompleto = "D:\Theeye\desktop\images\pagoCancelacionCompleto.bmp"
Local $otrosConceptos = "D:\Theeye\desktop\images\otrosConceptos.bmp"
Local $importeCobroCancelacion2 = "D:\Theeye\desktop\images\importeCobroCancelacion2.bmp"
Local $seleccioneUbicacion = "D:\Theeye\desktop\images\seleccioneUbicacion.bmp"
Local $grabacionExitosa = "D:\Theeye\desktop\images\grabacionExitosa.bmp"
Local $procesadoPagosCorrectamente = "D:\Theeye\desktop\images\procesadoPagosCorrectamente.bmp"



;si es acreditado y NF-COMAFI como agencia siguien con el otro, si no cumple con ninguno de los dos requisitos hay que procesarlo
;si al distribuir como pago parcial o cancelatorio no se encuentre datos no aparezca producto osea ninguna fila se lo tomaba "Cobro sin saldo" estado_original: cobro sin saldo ccargarlo manualmente

Local $iteracionesAdmTerceros = 10
$jsonFilePath = "D:\theeye\desktop\data\estructuraPagosEmerix.json"

if Not FileExists($jsonFilePath) Then
	  $status= "failure"
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
Local $oEstado_lote = Json_get($oStatusBases,'["estado_lote"]')

if $oEstado_lote <>  "pending" Then
   $status= "success"
   Local $mensajeError =  $fecha_ejecucion & " conciliaciones completas"
   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
   Exit
EndIf


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

 If (FileExists($ejecucionLog)) Then
   FileWrite($ejecucionLog, $fecha_ejecucion & " Se ejecuto downloadExcel "& @CRLF)
EndIf

For  $opagos_emerix  In $opagos_emerixArray

   Local $fechaDeConvenio = Json_get($opagos_emerix,'["fecha_de_convenio"]')
   Local $fecha_de_pago_emerix = Json_get($opagos_emerix,'["fecha_de_pago_emerix"]')
   Local $DNI = Json_get ($opagos_emerix,'["DNI"]')
   Local $estado_original = Json_get ($opagos_emerix,'["estado_original"]')
   Local $tipo_de_pago = Json_get ($opagos_emerix,'["tipo_de_pago"]')
   Local $agencia = Json_get ($opagos_emerix,'["agencia"]')
   Local $moneda = Json_get ($opagos_emerix,'["moneda"]')
   Local $importe_emerix = Json_get ($opagos_emerix,'["importe_emerix"]')
   Local $valor_cuota = Json_get ($opagos_emerix,'["valor_cuota"]')
   Local $cant_total_cuotas = Json_get ($opagos_emerix,'["cant_total_cuotas"]')
   Local $tipo_actividad = "Parcial";Json_get ($opagos_emerix,'["tipo_actividad"]')
   Local $suma_importe_cobros_web_emerix = Json_get ($opagos_emerix,'["suma_importe_cobros_web_emerix"]')
   Local $estado_carga = Json_get ($opagos_emerix,'["estado_carga"]')

   if $estado_carga = "Acreditado NF-COMAFI" Then
	   Json_Put($opagos_emerix, ".error", "Ya se acredito pago")
	  ContinueLoop
   EndIf


   Local $user = "user1"
   Local $pass = "user1"
      ;cierro todas las instancias de IE.
   RunWait('taskkill /F /IM "iexplore.exe"')

   ;Ingreso al sitio

    Local $tipoDePago = "Parcial";$tipo_actividad;"Cancelatorio"
    Local $urlEntorno = ""
    If $EntornoPrmtr = "DEV" Then
	  $urlEntorno = "http://dstst04"
	  Local $oIE = _IECreate($urlEntorno&"/tandem/login/source.asp")
	  _IELoadWait($oIE)
	  Sleep(2000)
	  $oIE = _IEAttach("Seleccione Ubicación")
	  if imageExists($seleccioneUbicacion, 10) then
		 Send("{F11}")
	  EndIf

	  ;///////CLICK A TEST///////////
	  Local $images = _IETagNameGetCollection($oIE, "img")
	  Local $imageSolTest
	  sleep(6000)
	  For $image in $images
		 $src_value = $image.GetAttribute("src")
		 if($src_value = $urlEntorno&"/tandem/operacional/images/general/argentina.gif") Then
			$imageSolTest = $image
		 ExitLoop
		 EndIf
	  Next
	  _IEAction($imageSolTest, "click")
	  Sleep(4000)
	  ;///////Ingresamos usuario y contraseña///////////
	  send($user)
	  send ("{TAB}")
	  send($pass)
	  send ("{TAB}")
	  Send ("{ENTER}")
	  Sleep(10000)
	  ;///////CLICK A Adm canales de terceros///////////
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

   ;//////////ENTORNO A PROD//////////////
    ElseIf $EntornoPrmtr = "PROD" Then
	   $urlEntorno = "http://bcweb11:90"
	  Local $oIE = _IECreate($urlEntorno&"/tandem/login/source.asp")
	  _IELoadWait($oIE)
	  Sleep(2000)
	  if imageExists($seleccioneUbicacion, 10) then
		 Send("{F11}")
	  EndIf
	  $oIE = _IEAttach("Seleccione Ubicación")
	  Local $images = _IETagNameGetCollection($oIE, "img")
	  Local $imageSolTest
	  For $image in $images
		 $src_value = $image.GetAttribute("src")
		 if($src_value = $urlEntorno&"/tandem/operacional/images/general/argentina.gif") Then
			$imageSolTest = $image
		 ExitLoop
		 EndIf
	  Next
	  _IEAction($imageSolTest, "click")

	  sleep(5000)
	  ;///////CLICK A Adm canales de terceros///////////
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
    EndIf

   ;////////INGRESAMOS FECHA DESDE//////////////
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
	  send("06/10/2020");$fecha_de_pago_emerix)
    Else
	   FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$fechadesde& " no encontro el elemento para hacer click "& @CRLF)
	   $status= "failure"
	   Local $mensajeError = $fecha_ejecucion & " la imagen "&$fechadesde& " no encontro el elemento para hacer click "
	   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	   RunWait('taskkill /F /IM "iexplore.exe"')
	   Exit
    EndIf


   ;////////INGRESAMOS FECHA HASTA//////////////
    If imageExists($fechahasta, 10) then
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


	;//////////////////INGRESAMOS DOCUMENTO/////////////////////////

    if imageExists($nrodocumento, 10) then
	  clickOnImage($nrodocumento)
	  if @error Then
			FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$nrodocumento& @CRLF)
			$status= "failure"
			Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$nrodocumento
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		   RunWait('taskkill /F /IM "iexplore.exe"')
		   Exit
	  EndIf
	  FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso a adm de canales de terceros "& @CRLF)
	  send ("{TAB}")
	  send($DNI)
    Else
	   FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$nrodocumento& " no encontro el elemento para hacer click "& @CRLF)
	   $status= "failure"
	   Local $mensajeError = $fecha_ejecucion & " la imagen "&$nrodocumento& " no encontro el elemento para hacer click "
	   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	   RunWait('taskkill /F /IM "iexplore.exe"')
	   Exit
	EndIf
    sleep (1000)



    ;////////Click en filtro////////////////
    $oIE = _IEAttach("Emerix Tandem")
    Local $images = _IETagNameGetCollection($oIE, "a")
	Local $imageAplicarfiltro
	  For $image in $images
		 $id_value = $image.GetAttribute("id")
		 if($id_value = "ctl00_ContentPlaceHolder1_FrameContainer1_cmdFiltrar_lnkLink") Then
			$imageAplicarfiltro = $image
		 ExitLoop
		 EndIf
	  Next
	  _IEAction($imageAplicarfiltro, "click")
   sleep (5000)

   ;///////////Se notifica si no se encontro ningun resultado////////////
    if imageExists($noseencontrarondatos, 10) then
	  Json_Put($opagos_emerix, ".error", "No se encontraron Datos")
	 ContinueLoop
    EndIf


   ;///////////Se verifica que este en agencia NF-comafi////////////
    if Not imageExists($agenciaNFComafi,10) Then
	  _cambiarAgencia()
    EndIf
    Sleep(2000)

	  ;///////////Se verifica que este En estado  Ingresado/////////////
   if Not imageExists($estadoIngresado,10)Then
	  if imageExists($checkSelect,10)Then
		 clickOnImage($checkSelect)
	  EndIf

	  Sleep(2000)
	  if imageExists($cambiar,10) then
		 clickOnImage($cambiar)
	  EndIf

	  Sleep(5000)
	  if imageExists($cambioDeEstado,10) Then
		 clickOnImage($cambioDeEstado)
		  If imageExists($aceptar,10) Then
			clickOnImage($aceptar)
		  EndIf
	  EndIf
   EndIf

    ;///////////Si es pago parcial, se verifican las cuotas para pasarlo a cancelatorio////////////
   If $tipoDePago == "Parcial" Then
	  sleep (2000)
	  ;click a revision
	  $oIE = _IEAttach("Emerix Tandem")
	  Local $images = _IETagNameGetCollection($oIE, "a")
	  Local $imageRevision
	  For $image in $images
		 $href_value = $image.GetAttribute("href")
		 if($href_value = $urlEntorno&"/tandem/operacional/revision/revision.asp") Then
			$imageRevision = $image
			ExitLoop
			EndIf
	  Next
	  _IEAction($imageRevision, "click")
	  sleep (5000)
	  ;click a documentos
	  $oIE = _IEAttach("Revisión de Cuentas")
	  Local $images = _IETagNameGetCollection($oIE, "a")
	  Local $imageRevision
	  For $image in $images
		 $title_value = $image.GetAttribute("title")
		 if($title_value = "Buscar por Nº Documento") Then
			$imageRevision = $image
			ExitLoop
			EndIf
	  Next
	  _IEAction($imageRevision, "click")

	  Sleep(3000)

	  ;escribimos documento
	  $oIE = _IEAttach("Revisión de Cuentas")
	  Local $images = _IETagNameGetCollection($oIE, "input")
	  Local $imageRevision
	  For $image in $images
		$title_value = $image.GetAttribute("name")
		if($title_value = "P_NRO_DOC") Then
		   $imageRevision = $image
			ExitLoop
		EndIf
	  Next
	  _IEAction($imageRevision, "click")
	  sleep (3000)
	  send($DNI)
	  sleep (1000)
	  send("{ENTER}")
	  sleep (3000)

	  ;hacemos click en N° Cliente de NF-AgenciaComafi
	  $oIE = _IEAttach("Revisión de Cuentas")
	  Local $oTable = _IETableGetCollection($oIE,28)
	  Local  $aTableData = _IETableWriteToArray($oTable, True)
	    ;_ArrayDisplay($aTableData,'sometitle'&$i)
	  Local $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
	  Local $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.

	  if imageExists($nfComafi2, 10) then
	  clickOnImage($nfComafi2)
		 if imageExists($nfComafi, 10) then
			clickOnImage($nfComafi)
			if @error Then
				  FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$nfComafi& @CRLF)
				  $status= "failure"
				  Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$nfComafi
				  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
				 ;RunWait('taskkill /F /IM "iexplore.exe"')
				 ;Exit
			EndIf
			FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso a adm de canales de terceros "& @CRLF)
		 Else
			 FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$nfComafi& " no encontro el elemento para hacer click "& @CRLF)
			 $status= "failure"
			 Local $mensajeError = $fecha_ejecucion & " la imagen "&$nfComafi& " no encontro el elemento para hacer click "
			 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			; RunWait('taskkill /F /IM "iexplore.exe"')
			; Exit
		 EndIf
	  EndIf
	  Sleep(2000)
	  if imageExists($962, 10) then
	  clickOnImage($962)

		 if imageExists($96, 10) then
			clickOnImage($96)
			if @error Then
				  FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$96& @CRLF)
				  $status= "failure"
				  Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$96
				  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
				 ;RunWait('taskkill /F /IM "iexplore.exe"')
				 ;Exit
			EndIf
			FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso a adm de canales de terceros "& @CRLF)
		 Else
			 FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$96& " no encontro el elemento para hacer click "& @CRLF)
			 $status= "failure"
			 Local $mensajeError = $fecha_ejecucion & " la imagen "&$96& " no encontro el elemento para hacer click "
			 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			; RunWait('taskkill /F /IM "iexplore.exe"')
			; Exit
		 EndIf
	  EndIf
	  Sleep(5000)


	  ;verificamos si aparece el cartel de cuenta lockeada
	  if imageExists($cuentaLockeada, 10) then
	  clickOnImage($cuentaLockeada)
		 if imageExists($cotinuarCuentaLockeada,10) Then
			clickOnImage($cotinuarCuentaLockeada)
		 EndIf
	  EndIf

	  ;Addins
	  Sleep(2000)
	  if imageExists($addins, 10) then
		 clickOnImage($addins)
		 if @error Then
			   FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$addins& @CRLF)
			   $status= "failure"
			   Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$addins
			   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			  RunWait('taskkill /F /IM "iexplore.exe"')
			  Exit
		 EndIf
		 FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso a adm de canales de terceros "& @CRLF)
	  Else
		  FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$addins& " no encontro el elemento para hacer click "& @CRLF)
		  $status= "failure"
		  Local $mensajeError = $fecha_ejecucion & " la imagen "&$addins& " no encontro el elemento para hacer click "
		  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		  RunWait('taskkill /F /IM "iexplore.exe"')
		  Exit
	   EndIf

	  ;cobros
	  if imageExists($cobros, 10) then
		 clickOnImage($cobros)
		 if @error Then
			   FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$cobros& @CRLF)
			   $status= "failure"
			   Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$cobros
			   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			  RunWait('taskkill /F /IM "iexplore.exe"')
			  Exit
		 EndIf
		 FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso a adm de canales de terceros "& @CRLF)
	  Else
		  FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$cobros& " no encontro el elemento para hacer click "& @CRLF)
		  $status= "failure"
		  Local $mensajeError = $fecha_ejecucion & " la imagen "&$cobros& " no encontro el elemento para hacer click "
		  ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		  RunWait('taskkill /F /IM "iexplore.exe"')
		  Exit
	   EndIf
	   Sleep(2000)

	  ;VERIFICAMOS SI EXISTE ALGUN COBRO, SI NO EXISTE SE PASA A CANCELATORIO

	  Local $oIE = _IEAttach("Cobros")
	  Sleep(3000)
		 ;if Not imageExists($noseencontrarondatos,10) Then
	  if Not imageExists($noseencontrarondatos,10) Then
		 Local $oTable = _IETableGetCollection($oIE,6)
		 Local  $aTableData = _IETableWriteToArray($oTable, True)
				 ; _ArrayDisplay($aTableData,'sometitle')

		 Local $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
		 Local $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.
		 ;MsgBox($MB_SYSTEMMODAL, "Table Info", "There are " & $iRows & $iCols & " tables on the page")
		 Local $totalPagos = 0
		 Local $valorImporte = 0
		 For $i = 0 To $iRows -1
			For $j = 0 To $iCols -1
			   if $j = 2 Then
				  Local $Importe = $aTableData[$i][$j]
				  if $Importe <> "Importe" Then
					 $Importe = StringReplace($Importe,",",".")
				  EndIf
				  $totalPagos = $totalPagos + Number($Importe);$aTableData[$i][$j],3)
			   EndIf
			Next
		 Next

		 Local $importeTotal = $valor_cuota * $cant_total_cuotas
		 Local $importeCobradoyEmerix = $totalPagos + $importe_emerix
		 Local $cuentaFinalDeImportes = $importeTotal - $importeCobradoyEmerix
		 MsgBox("aaa",$cuentaFinalDeImportes,$cuentaFinalDeImportes)

		 if $cuentaFinalDeImportes > 0 Then
			   $tipoDePago = "Pago Parcial"
		 EndIf
	  EndIf
	  Sleep(2000)
	  Send("^w")
	  ;VOLVEMOS A ADMN CANALES DE TERCEROS Y VAMOS AL CAMINO FELIZ
	  ;///////CLICK A Adm canales de terceros///////////
	  $oIE = _IEAttach("Ficha de la Persona")
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

	  ;////////INGRESAMOS FECHA//////////////
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
		 send("06/10/2020");$fecha_de_pago_emerix)
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

	  Sleep
	   ;//////////////////INGRESAMOS DOCUMENTO/////////////////////////

	  if imageExists($nrodocumento, 10) then
		 clickOnImage($nrodocumento)
		 if @error Then
			FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$nrodocumento& @CRLF)
			$status= "failure"
			Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$nrodocumento
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			RunWait('taskkill /F /IM "iexplore.exe"')
			Exit
		 EndIf
		 FileWrite($ejecucionLog, $fecha_ejecucion & " Ingreso a adm de canales de terceros "& @CRLF)
		 send ("{TAB}")
		 send ("{TAB}")
		 send($DNI)
	  Else
		 FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$nrodocumento& " no encontro el elemento para hacer click "& @CRLF)
		 $status= "failure"
		 Local $mensajeError = $fecha_ejecucion & " la imagen "&$nrodocumento& " no encontro el elemento para hacer click "
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		 RunWait('taskkill /F /IM "iexplore.exe"')
		 Exit
	  EndIf
	  sleep (1000)



	  ;////////Click en filtro////////////////
	  $oIE = _IEAttach("Emerix Tandem")
	  Local $images = _IETagNameGetCollection($oIE, "a")
	  Local $imageAplicarfiltro
	  For $image in $images
		 $id_value = $image.GetAttribute("id")
		 if($id_value = "ctl00_ContentPlaceHolder1_FrameContainer1_cmdFiltrar_lnkLink") Then
			$imageAplicarfiltro = $image
			ExitLoop
		 EndIf
	  Next
	  _IEAction($imageAplicarfiltro, "click")
	  sleep (5000)

	  ;///////////Se notifica si no se encontro ningun resultado////////////
	  if imageExists($noseencontrarondatos, 10) then
		 Json_Put($opagos_emerix, ".error", "No se encontraron Datos")
		 ContinueLoop
	  EndIf
   EndIf

   ;////////////////////////////////////////////
   ;CAMINO FELIZ
   ;click en distribucion

   ;$ingresoDeCobros

   if imageExists($distribucion,10) then
	  clickOnImage($distribucion)
	  Sleep(1000)
	  if imageExists($nocontribuirpagoanulado,10) Then
		 Json_Put($opagos_emerix, ".error", "No se puede contribuir un pago anulado o acreditado")
		 ContinueLoop
	  EndIf
   EndIf

   Sleep(5000)


   MouseMove(787, 420)
   MouseClick($MOUSE_CLICK_LEFT)

   Sleep(3000)
   if ($tipoDePago == "Cancelatorio") Then
	  send("Cancelacion")
	  send("{ENTER}")
	  send("{TAB}")
	  send("{ENTER}")
   Else
	  Send("Pago")
	  send("{ENTER}")
	  send("{TAB}")
	  send("{ENTER}")
   EndIf


   Sleep(5000)

   ;click en boton distribuir
   ;Ingreso de Cobros
   ;click en check de Select
   $oIE = _IEAttach("Ingreso de Cobros","title",2)
   If @error = $_IEStatus_NoMatch Then
	  MsgBox("title","falla el match",2)
   EndIf
   Local $images = _IETagNameGetCollection($oIE, "input")
   Local $distribuir
   For $image in $images
	  $title_value = $image.GetAttribute("name")
	  if($title_value = "chkCol") Then
			$distribuir = $image
			_IEAction($distribuir, "click")
	  EndIf
   Next
   Sleep(2000)

   ;//////////SI TIENE UNA SOLA FILA EN EL INPUT DE IMPORTE COBRO SE LE PONE importe_emerix. Sino se hace Prorrateo de importe por producto/////
   $oIE = _IEAttach("Ingreso de Cobros","title",2)
   Local $oTable = _IETableGetCollection($oIE,20)
   Local  $aTableData = _IETableWriteToArray($oTable, True)
   Local $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
   Local $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.

   Sleep(2000)
   If $iRows == 2 Then
	  if $tipoDePago == "Cancelatorio" Then
		 if imageExists($importeCobroCancelacion, 10) then
			clickOnImage($importeCobroCancelacion)
			If @error Then
				FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$importeCobroCancelacion& @CRLF)
				$status= "failure"
				Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$importeCobroCancelacion& @CRLF
				ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			    RunWait('taskkill /F /IM "iexplore.exe"')
			    Exit
			EndIf
		 Else
			   FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$importeCobroCancelacion& " no encontro el elemento para hacer click "& @CRLF)
			   $status= "failure"
			   Local $mensajeError = $fecha_ejecucion & " la imagen "&$importeCobroCancelacion& " no encontro el elemento para hacer click "
			   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			   RunWait('taskkill /F /IM "iexplore.exe"')
			   Exit
		 EndIf
			Sleep(1000)
			send ("{TAB}")
			Sleep(1000)
			send ("{TAB}")
			Send($importe_emerix)
	   EndIf

	  if $tipoDePago == "Pago Parcial" Then
		 if imageExists($importeCobroPagoParcial, 10) then
			clickOnImage($importeCobroPagoParcial)
			If @error Then
				FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$importeCobroPagoParcial& @CRLF)
				$status= "failure"
				Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$importeCobroPagoParcial& @CRLF
				ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			    RunWait('taskkill /F /IM "iexplore.exe"')
			    Exit
			EndIf
		 Else
			   FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$importeCobroPagoParcial& " no encontro el elemento para hacer click "& @CRLF)
			   $status= "failure"
			   Local $mensajeError = $fecha_ejecucion & " la imagen "&$importeCobroPagoParcial& " no encontro el elemento para hacer click "
			   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			   RunWait('taskkill /F /IM "iexplore.exe"')
			   Exit
			EndIf
			Sleep(1000)
			send ("{TAB}")
			Sleep(1000)
			send ("{TAB}")
			Sleep(1000)
			send ("{TAB}")
			Send($importe_emerix)
	  EndIf


	  If imageExists($otrosConceptos, 10) then
		 clickOnImage($otrosConceptos)
		 if @error Then
			FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$otrosConceptos& @CRLF)
			$status= "failure"
			Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$otrosConceptos& @CRLF
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			RunWait('taskkill /F /IM "iexplore.exe"')
			Exit
		 EndIf
	  Else
		 FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$otrosConceptos& " no encontro el elemento para hacer click "& @CRLF)
		 $status= "failure"
		 Local $mensajeError = $fecha_ejecucion & " la imagen "&$otrosConceptos& " no encontro el elemento para hacer click "
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		 RunWait('taskkill /F /IM "iexplore.exe"')
		Exit
	  EndIf


	  ;grabamos
	  $oIE = _IEAttach("Ingreso de Cobros","title",2)
	  Local $images = _IETagNameGetCollection($oIE, "a")
	  Local $imageAplicarfiltro
	  For $image in $images
		 $id_value = $image.GetAttribute("id")
		 if($id_value = "cmdGrabar_lnkLink") Then
			$imageAplicarfiltro = $image
			ExitLoop
		 EndIf
	  Next
	  _IEAction($imageAplicarfiltro, "click")

	  ;si sale grabacion exitosa




   EndIf

   ;////////////////PORRATEO/////////////////
   if $tipoDePago == "Cancelatorio" Then
		 If $iRows > 2 Then
			Local $totalPagos = 0
			Local $rows = $iRows -1
			Local $arrayCapital[$rows]

			For $i = 0 To $iRows -1
			   For $j = 0 To $iCols -1
				  if $j = 6 Then
					 Local $Capital = $aTableData[$i][$j]
					 if $Capital <> "Capital" Then
						$Capital = StringReplace($Capital,",",".")
						Number($Capital)
						_ArrayPush($arrayCapital,  Number($Capital))
						 If $Capital == "0,00" Then
							  Local $images = _IETagNameGetCollection($oIE, "input")
							  Local $distribuir
							  For $image in $images
								 $title_value = $image.GetAttribute("name")
								 if($title_value = "chkCol") Then
									   $distribuir = $image
									   _IEAction($distribuir, "click")
								 EndIf
							  Next
						 EndIf
					 EndIf
				   $totalPagos = $totalPagos + Number($Capital);$aTableData[$i][$j],3)
				  EndIf
			   Next
			Next



			if imageExists($importeCobroCancelacion, 10) then
			   clickOnImage($importeCobroCancelacion)
			   if @error Then
					 FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$importeCobroCancelacion& @CRLF)
					 $status= "failure"
					 Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$importeCobroCancelacion& @CRLF
					 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
					RunWait('taskkill /F /IM "iexplore.exe"')
					Exit
			   EndIf
			Else
				FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$importeCobroCancelacion& " no encontro el elemento para hacer click "& @CRLF)
				$status= "failure"
				Local $mensajeError = $fecha_ejecucion & " la imagen "&$importeCobroCancelacion& " no encontro el elemento para hacer click "
				ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
				RunWait('taskkill /F /IM "iexplore.exe"')
				Exit
			EndIf
			Sleep(2000)
			;calculamos  el importe
			For $i = 0 To $Rows -1
			   Local $resultadoEnPorcentaje = 100*$arrayCapital[$i]/$totalPagos
			   Local $importeCobro = $resultadoEnPorcentaje * $importe_emerix/100
			   Sleep(1000)
			   send ("{TAB}")
			    Sleep(1000)
			   send ("{TAB}")
			   Send($importeCobro)
			Next
			Sleep(1000)

			If imageExists($otrosConceptos, 10) then
			   clickOnImage($otrosConceptos)
			   if @error Then
					 FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$otrosConceptos& @CRLF)
					 $status= "failure"
					 Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$otrosConceptos& @CRLF
					 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
					RunWait('taskkill /F /IM "iexplore.exe"')
					Exit
			   EndIf
			Else
				FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$otrosConceptos& " no encontro el elemento para hacer click "& @CRLF)
				$status= "failure"
				Local $mensajeError = $fecha_ejecucion & " la imagen "&$otrosConceptos& " no encontro el elemento para hacer click "
				ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
				RunWait('taskkill /F /IM "iexplore.exe"')
				Exit
			 EndIf
			;grabamos
			$oIE = _IEAttach("Ingreso de Cobros","title",2)
			Local $images = _IETagNameGetCollection($oIE, "a")
			Local $imageAplicarfiltro
			For $image in $images
			   $id_value = $image.GetAttribute("id")
			   if($id_value = "cmdGrabar_lnkLink") Then
				  $imageAplicarfiltro = $image
				  ExitLoop
			   EndIf
			Next
			_IEAction($imageAplicarfiltro, "click")




		 EndIf
   EndIf


   if $tipoDePago == "Pago Parcial" Then
		 If $iRows > 2 Then
			Local $totalPagos = 0
			Local $rows = $iRows -1
			Local $arrayCapital[$rows]

			For $i = 0 To $iRows -1
			   For $j = 0 To $iCols -1
				  if $j = 4 Then
					 Local $Capital = $aTableData[$i][$j]
					 if $Capital <> "Capital" Then
					  $Capital = StringReplace($Capital,",",".")
					  Number($Capital)
					  _ArrayPush($arrayCapital,  Number($Capital))
					 EndIf
				   $totalPagos = $totalPagos + Number($Capital);$aTableData[$i][$j],3)
				  EndIf
			   Next
			Next

			if imageExists($importeCobroPagoParcial, 10) then
			   clickOnImage($importeCobroPagoParcial)
			   if @error Then
					 FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$importeCobroPagoParcial& @CRLF)
					 $status= "failure"
					 Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$importeCobroPagoParcial& @CRLF
					 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
					RunWait('taskkill /F /IM "iexplore.exe"')
					Exit
			   EndIf
			Else
				FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$importeCobroPagoParcial& " no encontro el elemento para hacer click "& @CRLF)
				$status= "failure"
				Local $mensajeError = $fecha_ejecucion & " la imagen "&$importeCobroPagoParcial& " no encontro el elemento para hacer click "
				ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
				RunWait('taskkill /F /IM "iexplore.exe"')
				Exit
			EndIf
			Sleep(2000)
			;calculamos  el importe
			For $i = 0 To $Rows -1
			   Local $resultadoEnPorcentaje = 100*$arrayCapital[$i]/$totalPagos
			   Local $importeCobro = $resultadoEnPorcentaje * $importe_emerix/100
			   Sleep(1000)
			   send ("{TAB}")
			    Sleep(1000)
			   send ("{TAB}")
			   Sleep(1000)
			   send ("{TAB}")
			   Send($importeCobro)
			Next
			Sleep(1000)

			If imageExists($otrosConceptos, 10) then
			   clickOnImage($otrosConceptos)
			   if @error Then
					 FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$otrosConceptos& @CRLF)
					 $status= "failure"
					 Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$otrosConceptos& @CRLF
					 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
					RunWait('taskkill /F /IM "iexplore.exe"')
					Exit
			   EndIf
			Else
				FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$otrosConceptos& " no encontro el elemento para hacer click "& @CRLF)
				$status= "failure"
				Local $mensajeError = $fecha_ejecucion & " la imagen "&$otrosConceptos& " no encontro el elemento para hacer click "
				ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
				RunWait('taskkill /F /IM "iexplore.exe"')
				Exit
			 EndIf
			;grabamos
			$oIE = _IEAttach("Ingreso de Cobros","title",2)
			Local $images = _IETagNameGetCollection($oIE, "a")
			Local $imageAplicarfiltro
			For $image in $images
			   $id_value = $image.GetAttribute("id")
			   if($id_value = "cmdGrabar_lnkLink") Then
				  $imageAplicarfiltro = $image
				  ExitLoop
			   EndIf
			Next
			_IEAction($imageAplicarfiltro, "click")
		 EndIf
   EndIf

;prorateo calculo 100*capital/sumadetodos los capitales = resultado en porcentaje
;resultado en porcentaje*importe_emerix/100
    ;si sale grabacion exitosa le damos a aceptar y vamos a procesarlo.
	Sleep(3000)

   If imageExists($grabacionExitosa, 10) then
		 clickOnImage($grabacionExitosa)
		 Send("{ENTER}")
		 Sleep(3000)
		 if @error Then
			FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$grabacionExitosa& @CRLF)
			$status= "failure"
			Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$grabacionExitosa& @CRLF
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			RunWait('taskkill /F /IM "iexplore.exe"')
			Exit
		 EndIf
   Else
		 FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$grabacionExitosa& " no encontro el elemento para hacer click "& @CRLF)
		 $status= "failure"
		 Local $mensajeError = $fecha_ejecucion & " la imagen "&$grabacionExitosa& " no encontro el elemento para hacer click "
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		 RunWait('taskkill /F /IM "iexplore.exe"')
		 Exit
   EndIf


   ;VERIFICAMOS QUE ESTE VALIDADO Y LO PROCESAMOS
    $oIE = _IEAttach("Emerix Tandem")

   Local $oTable = _IETableGetCollection($oIE,46)
   Local  $aTableData = _IETableWriteToArray($oTable, True)
   $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
   $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.

	  For $i = 0 To $iRows -1
		 For $j = 0 To $iCols -1
			if $j = 11 Then
			   Local $Importe = $aTableData[$i][$j]
			   if $Importe == "VALIDADO" Then
					 If @error = $_IEStatus_NoMatch Then
						MsgBox("title","falla el match",2)
					 EndIf
					 Local $images = _IETagNameGetCollection($oIE, "input")
					 Local $distribuir
					 For $image in $images
						$title_value = $image.GetAttribute("name")
						if($title_value = "chkCol") Then
							  $distribuir = $image
							  ExitLoop
						EndIf
					 Next
					 _IEAction($distribuir, "click")
			   EndIf
			EndIf
		 Next
	  Next
	  Sleep(3000)
	  Local $images = _IETagNameGetCollection($oIE, "a")
	  Local $distribuir
	  For $image in $images
		 $title_value = $image.GetAttribute("id")
		 if($title_value = "ctl00_ContentPlaceHolder1_BTN_Procesar_lnkLink") Then
			$distribuir = $image
			ExitLoop
		 EndIf
	  Next
	  _IEAction($distribuir, "click")

   Sleep(3000)

   If imageExists($procesadoPagosCorrectamente, 10) then
		 clickOnImage($procesadoPagosCorrectamente)
		 Send("{ENTER}")
		 Sleep(3000)
		 if @error Then
			FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$procesadoPagosCorrectamente& @CRLF)
			$status= "failure"
			Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$procesadoPagosCorrectamente& @CRLF
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			RunWait('taskkill /F /IM "iexplore.exe"')
			Exit
		 EndIf
   Else
		 FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$procesadoPagosCorrectamente& " no encontro el elemento para hacer click "& @CRLF)
		 $status= "failure"
		 Local $mensajeError = $fecha_ejecucion & " la imagen "&$procesadoPagosCorrectamente& " no encontro el elemento para hacer click "
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		 RunWait('taskkill /F /IM "iexplore.exe"')
		 Exit
   EndIf

    Json_Put($opagos_emerix, ".estado_carga", "Finished")

Next
Exit



#cs ----------------------------------------------------------------------------

    sleep (10000)


   if imageExists($noseencontrarondatos, 10) then
	  Json_Put($opagos_emerix, ".error", "No se encontraron Datos")
	  ContinueLoop
   EndIf




   if Not imageExists($agenciaNFComafi,10) Then
	  _cambiarAgencia()
   EndIf


   if imageExists($distribucion,10) then
	  clickOnImage($distribucion)
	  Sleep(1000)
	  if imageExists($nocontribuirpagoanulado,10) Then
		 Json_Put($opagos_emerix, ".error", "No se puede contribuir un pago anulado o acreditado")
		 ContinueLoop
	  EndIf
   EndIf


#ce ----------------------------------------------------------------------------





saveJsonStatus($jsonFilePath, $opagos_emerix)


Func _cambiarAgencia()
   if imageExists($cambiar,10) then
     clickOnImage($cambiar)
	  if @error Then
		   FileWrite($error_Log, $fecha_ejecucion & " no existe la imagen "&$cambiar& @CRLF)
		   $status= "failure"
			Local $mensajeError = $fecha_ejecucion & " no existe la imagen "&$cambiar
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		   RunWait('taskkill /F /IM "iexplore.exe"')
		   Exit
	  EndIf
   Else
	   FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$cambiar& " no encontro el elemento para hacer click "& @CRLF)
	   $status= "failure"
	   Local $mensajeError = $fecha_ejecucion & " la imagen "&$cambiar& " no encontro el elemento para hacer click "
	   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	   RunWait('taskkill /F /IM "iexplore.exe"')
	   Exit
   EndIf

   Send ("{TAB}")
   Sleep(500)

   Switch $agencia
	  Case "NF-COMAFI"
		  send ("NF")
	  Case "Bot de Cobranza"
		 send ("BOT")
	  Case "Bot Live Person"
		 send ("BOT")
		 Sleep(1000)
		 Send("{DOWN}")
	  Case Else
		 send ("NF")
   EndSwitch

	sleep (500)
	Send ("{TAB}")

   Send ("{ENTER}")


EndFunc





Func saveJsonStatus($JsonfileName, $oJson)

    $outputFile = $JsonfileName

    $JSON = Json_encode($oJson)

    IF (FileExists ($outputFile)) Then
        FileDelete($outputFile)
    EndIf

    If Not (FileWrite($outputFile,$JSON)) Then
			;LogError(@error & "Cannot save JsonFile")
			FileWrite($error_Log, $fecha_ejecucion & " error al guardar el archivo   "& $filename&@CRLF)
			;MsgBox(0,"Error Saving File", $JsonfileName, 5)
            Return False
    EndIf

EndFunc
