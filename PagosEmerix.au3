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
If _Singleton("PagosEmerix", 1) = 0 Then
   $status = "failure"
   ConsoleWrite('{"state":"' & $status & '", "data":["' & "falla singleton" & '"]}')
   Exit
EndIf


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
Local $nopuedeingresarimportevalorcero = "D:\Theeye\desktop\images\nopuedeingresarimportevalorcero.bmp"
Local $admcanalesdeterceros = "D:\Theeye\desktop\images\admcanalesdeterceros.bmp"
Local $falloCambiarEstado ="D:\Theeye\desktop\images\falloCambiarEstado.bmp"
Local $flechitaCobroCancelacion = "D:\Theeye\desktop\images\flechitaCobroCancelacion.bmp"


;si es acreditado y NF-COMAFI como agencia siguien con el otro, si no cumple con ninguno de los dos requisitos hay que procesarlo

;si al distribuir como pago parcial o cancelatorio no se encuentre datos no aparezca producto osea ninguna fila se lo tomaba "Cobro sin saldo" estado_original: cobro sin saldo ccargarlo manualmente
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

Local $EntornoPrmtr = $CmdLine[1]
Local $fileJsonPrmtr = $CmdLine[2]
$jsonFilePath = "D:\Theeye\desktop\data\estructuraPagosEmerix.json";$fileJsonPrmtr

   ToolTip($jsonFilePath, 0, 0)
   Sleep(2000) ; Sleep to give tooltip time to display

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
   Local $tipo_actividad = Json_get ($opagos_emerix,'["tipo_actividad"]')
   Local $suma_importe_cobros_web_emerix = Json_get ($opagos_emerix,'["suma_importe_cobros_web_emerix"]')
   Local $estado_carga = Json_get ($opagos_emerix,'["estado_carga"]')
   ToolTip($fechaDeConvenio, 0, 0)
   Sleep(1000) ; Sleep to give tooltip time to display



   ToolTip($DNI, 0, 0)
   Sleep(1000) ; Sleep to give tooltip time to display

  Local $tipoDePago = $tipo_actividad
   ToolTip($tipoDePago, 0, 0)
   Sleep(1000) ; Sleep to give tooltip time to display

   if $estado_carga = "Acreditado NF-COMAFI" Then
	   Json_Put($opagos_emerix, ".error", "Ya se acredito pago")
	  ContinueLoop
   EndIf
   If $EntornoPrmtr = "DEV" Then
	  $urlEntorno = "http://dstst04"
   EndIf
   If $EntornoPrmtr = "PROD" Then
	  $urlEntorno = "http://bcweb11:90"
   EndIf
   ToolTip($urlEntorno, 0, 0)
   Sleep(1000) ; Sleep to give tooltip time to display
   ;Login
   $checklogin = _login($urlEntorno)
   if $checklogin == False Then
	  ContinueLoop 1
   EndIf
   ;Click a Adm Canales de Terceros
   Local $titleAttach = "Revisión de Cuentas"
   _AdmDeTerceros($urlEntorno,$titleAttach)

   ;Ingresamos fecha desde
   $contador = 1
   while Not imageExists($fechadesde, 10)
	  Sleep(1000)

	  if $contador > 45 Then
		 $status= "failure"
		 Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$flechaGuardar
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		 Json_Put($opagos_emerix, ".error", "Error en encontrar imagen"&$fechadesde)
		 _saveStatus()
		 ContinueLoop
	  EndIf
	  $contador = $contador+1
   WEnd
   _clickInImage($fechadesde)
   send ("{TAB}")
   send($fecha_de_pago_emerix)

    ;Ingresamos fecha hasta
  $contador = 1
   while Not imageExists($fechahasta, 10)
	  Sleep(1000)
	  if $contador > 45 Then
		 $status= "failure"
		 Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$flechaGuardar
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		 Json_Put($opagos_emerix, ".error", "Error en encontrar imagen"&$fechahasta)
		 _saveStatus()
		 ContinueLoop
	  EndIf
	  $contador = $contador+1
   WEnd
   _clickInImage($fechahasta)
   send ("{TAB}")
   send($fecha_de_pago_emerix)

   $contador = 1
   ;Ingresamos Documento
   while Not imageExists($nrodocumento, 10)
	  Sleep(1000)

	  if $contador > 45 Then
		 $status= "failure"
		 Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$flechaGuardar
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		 Json_Put($opagos_emerix, ".error", "Error en encontrar imagen"&$nrodocumento)
		 _saveStatus()
		 ContinueLoop
	  EndIf
	  $contador = $contador+1
   WEnd
    _clickInImage($nrodocumento)
   send ("{TAB}")
   send($DNI)

   ;Filtramos
   _filtroEnAdmTerceros()

   ;///////////Se notifica si no se encontro ningun resultado////////////
    if imageExists($noseencontrarondatos, 10) then
	  Json_Put($opagos_emerix, ".error", "No se encontraron Datos")
	  _saveStatus()
	 ContinueLoop
    EndIf




   ;///////////Se verifica que este En estado  Ingresado/////////////
   $verEstado = _verificarEstado()
   if $verEstado == False Then
	  Json_Put($opagos_emerix, ".error", "Error al cambiar estado de ingreso Alert: 'Solo se pueden hacer reingresar pagos acreditados en el mismo dia' ")
	  _saveStatus()
	  ContinueLoop
   EndIf

   ;///////////Se verifica que este en agencia NF-comafi////////////
	 $cambiarAgencia =  _cambiarAgencia()
	 if $cambiarAgencia == "no se encontro cambio de agencia" Then
		 Json_Put($opagos_emerix, ".error", "no se encontro cambio de agencia")
		 _saveStatus()
		 ContinueLoop
	  EndIf
    Sleep(2000)



    ;///////////Si es pago parcial, se verifican las cuotas para pasarlo a cancelatorio////////////
   If $tipoDePago == "Parcial" Then
	  sleep(10000)
	  $tipoDePago = _verificarTipoDePago()
	  if $tipoDePago == False Then
		 ContinueLoop 1
	  EndIf
	  if $tipoDePago == "Error en encontrar imagen _verificarTipoDePago " Then
		 Json_Put($opagos_emerix, ".error", "Error en encontrar imagen _verificarTipoDePago ")
		 _saveStatus()
		 ContinueLoop
	  EndIf
	  ;///////////Se notifica si no se encontro ningun resultado////////////
	  if imageExists($noseencontrarondatos, 10) then
		 Json_Put($opagos_emerix, ".error", "No se encontraron Datos")
		 _saveStatus()
		 ContinueLoop
	  EndIf
   EndIf
   ;////////////////////////////////////////////
   ;CAMINO FELIZ
	  Sleep(2000)
	  $titleAttach = "Ficha de la Persona"
	   ;Click a Adm Canales de Terceros
	  _AdmDeTerceros($urlEntorno,$titleAttach)

	  ;Ingresamos fecha desde
	  $contador = 1
	  while Not imageExists($fechadesde, 10)
		 Sleep(1000)
		 if $contador > 45 Then
			$status= "failure"
			Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$fechadesde
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			Json_Put($opagos_emerix, ".error", "Error en encontrar imagen"&$fechadesde)
			_saveStatus()
			ContinueLoop
		 EndIf
		 $contador = $contador+1
	  WEnd
	  _clickInImage($fechadesde)
	  send ("{TAB}")
	  send($fecha_de_pago_emerix)

	  $contador = 1
	  ;Ingresamos fecha hasta
	  while Not imageExists($fechahasta, 10)
		 Sleep(1000)
		 if $contador > 45 Then
			$status= "failure"
			Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$fechahasta
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			Json_Put($opagos_emerix, ".error", "Error en encontrar imagen"&$fechahasta)
			_saveStatus()
			ContinueLoop
		 EndIf
		 $contador = $contador+1
	  WEnd
	  _clickInImage($fechahasta)
	  send ("{TAB}")
	  send($fecha_de_pago_emerix)

	  $contador = 1
	  ;Ingresamos Documento
	  while Not imageExists($nrodocumento, 10)
		 Sleep(1000)
		 if $contador > 45 Then
			$status= "failure"
			Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$nrodocumento
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			Json_Put($opagos_emerix, ".error", "Error en encontrar imagen"&$nrodocumento)
			_saveStatus()
			ContinueLoop
		 EndIf
		 $contador = $contador+1
	  WEnd
	  _clickInImage($nrodocumento)
	  send ("{TAB}")
	  send($DNI)

   ;Filtramos
   _filtroEnAdmTerceros()
   ;Ingresamos a cobros
	  $verificarCobros = _ingresoACobros($tipoDePago)

	  if $verificarCobros == "fail" Then
		 Json_Put($estado_carga, ".error", "No se puede contribuir un pago anulado o acreditado")
		 _saveStatus()
		 ContinueLoop
	  EndIf


   ;Verificamos si tiene alguna capital tiene 0.00
   $verificarUnaFilaCapitalCero = _verificarUnaFilaCapitalCero()
   if $verificarUnaFilaCapitalCero == True Then
	  Json_Put($opagos_emerix, ".error", "Verificar este pago de manera manual")
	 _saveStatus()
	  ContinueLoop
   EndIf

   $verificarCero = _verificarCapitalCero()
   ToolTip($verificarCero&"verificarCero", 0, 0)
   ;MsgBox("","",$verificarCero&"verificarCero")
   Sleep(1000) ; Sleep to give tooltip time to display
   $verificarPorrateo = ""
   If $verificarCero == True Then
	   Sleep(3000)
	  _clickInImage($grabacionExitosa)
	  Send("{ENTER}")
	  Sleep(3000)
	  _verificacionDeValidez()

   Else
	  ;si tiene una sola fila en el input de importe cobro se le pone importe_emerix
	   ToolTip($tipoDePago, 0, 0)
	   ;MsgBox("","",$tipoDePago&"linea 387")
	  _cobroConUnaSolaFila($tipoDePago)
	  ;//////////SI TIENE UNA SOLA FILA EN EL INPUT DE IMPORTE COBRO SE LE PONE importe_emerix. Sino se hace Prorrateo de importe por producto/////
	  $verificarPorrateo = _porrateo($tipoDePago)
	  Sleep(3000)
	  _clickInImage($grabacionExitosa)
	  Send("{ENTER}")
	  Sleep(3000)
	  _verificacionDeValidez()
   EndIf
   if $verificarPorrateo == "no se puede ingresar importe con valor cero" Then
			Json_Put($opagos_emerix, ".error", "no se puede ingresar importe con valor cero")
			_saveStatus()
			ContinueLoop
   EndIf

   If $verificarCero == "No es valido" Then
	  Json_Put($opagos_emerix, ".error", "No se pudo cargar en distribucion por 0 ")
	  _saveStatus()
	  ContinueLoop
   EndIf
   Sleep(3000)

   Json_Put($opagos_emerix, ".estado_carga", "done")
  _saveStatus()


Next

Func _saveStatus()

   Json_Put($oStatusBases, ".pagos_emerix",$opagos_emerixArray)
   saveJsonStatus($jsonFilePath, $oStatusBases)

EndFunc
Json_Put($oStatusBases, ".estado_lote", "done")

saveJsonStatus($jsonFilePath, $oStatusBases)
  ;cierro todas las instancias de IE.
RunWait('taskkill /F /IM "iexplore.exe"')

Func _verificarEstado()
	   ;VERIFICAMOS QUE ESTE VALIDADO Y LO PROCESAMOS
   $oIE = _IEAttach("Emerix Tandem")
   Local $oTable = _IETableGetCollection($oIE,46)
   Local  $aTableData = _IETableWriteToArray($oTable, True)
   $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
   $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.

	  For $i = 0 To $iRows -1
		 For $j = 0 To $iCols -1
			if $j = 11 Then
			   Local $Estado = $aTableData[$i][$j]
			   if $Estado <> "Estado" Then
				  if $Estado <> "INGRESADO" Then
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
						_clickInImage($cambiar2)
						Sleep(2000)
						if imageExists($cambioDeEstado,10) Then
						   _clickInImage($cambioDeEstado)
						   Sleep(3000)
							If imageExists($aceptar,10) Then
							  _clickInImage($aceptar)
							  Return True
						   EndIf
						EndIf
						Sleep(2000)
						if imageExists($falloCambiarEstado,10) Then
						   Return False
						EndIf
				  Else
					 Return True
				  EndIf
			   EndIf
			EndIf
		 Next
	  Next

EndFunc

Func _verificarUnaFilaCapitalCero()
   $oIE = _IEAttach("Ingreso de Cobros","title",2)
   If @error = $_IEStatus_NoMatch Then
	  MsgBox("title","falla el match",2)
   EndIf
   Local $oTable = _IETableGetCollection($oIE,20)
   Local  $aTableData = _IETableWriteToArray($oTable, True)
   Local $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
   Local $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.
   Local $capitalCero = False

   ;Hacemos un recorrido y verificamos que contenga "0.00" en la tabla. Si tiene 0.00 devuelve true.
   If $iRows == 2 Then
	  Local $totalPagos = 0
	  Local $rows = $iRows -1
	  For $i = 0 To $iRows -1
		 For $j = 0 To $iCols -1
			if $j = 6 Then
			   Local $Capital = $aTableData[$i][$j]
			   if $Capital <> "Capital" Then
				  if $Capital == "0,00" Then
					 $capitalCero = True
					 Return $capitalCero
				  EndIf
			   EndIf
			EndIf
		 Next
	  Next
   EndIf
   Return $capitalCero
EndFunc

Func _cambiarAgencia()
    Local $contador = 1
   while Not imageExists($cambiar, 10)
	  Sleep(1000)
	  if $contador > 45 Then
		 $status= "failure"
		 Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$cambiar
		 ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
		 Return "no se encontro cambio de agencia"
	  EndIf
	  $contador = $contador+1
   WEnd
   _clickInImage($cambiar)

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


Func _cobroConUnaSolaFila($tipoDePago)

	  ;click en check de Select
	  $oIE = _IEAttach("Ingreso de Cobros","title",2)
		 ToolTip("_cobroConUnaSolaFila", 0, 0)
	  If @error = $_IEStatus_NoMatch Then
		 MsgBox("title","falla el match",2)
	  EndIf

	  Sleep(2000)
	  Local $oTable = _IETableGetCollection($oIE,20)
	  Local  $aTableData = _IETableWriteToArray($oTable, True)
	  Local $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
	  Local $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.
	  Sleep(2000)
	  If $iRows == 2 Then
		 Local $images = _IETagNameGetCollection($oIE, "input")
		 Local $distribuir
		 For $image in $images
			$title_value = $image.GetAttribute("name")
			if($title_value = "chkCol") Then
				  $distribuir = $image
				  _IEAction($distribuir, "click")
			EndIf
		 Next
		 ToolTip($tipoDePago&"en _cobroConUnaSolaFila", 0, 0)
		 Sleep(1000) ; Sleep to give tooltip time to display
		 if $tipoDePago == "Cancelatorio" Then
			Sleep(2000)
			_clickInImage($importeCobroCancelacion)
			Sleep(1000)
			send ("{TAB}")
			Sleep(1000)
			send ("{TAB}")
			send($importe_emerix)
		 EndIf

		 ToolTip($tipoDePago&"en _cobroConUnaSolaFila", 0, 0)

		 if $tipoDePago == "Pago Parcial" Then
			   Sleep(2000)
			   _clickInImage($importeCobroPagoParcial)
			   Sleep(1000)
			   send ("{TAB}")
			   Sleep(1000)
			   send ("{TAB}")
			   Sleep(1000)
			   send ("{TAB}")
			   Send($importe_emerix)
		 EndIf
		 _clickInImage($otrosConceptos)


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
   EndFunc




Func _verificarCapitalCero()
   $oIE = _IEAttach("Ingreso de Cobros","title",2)
   If @error = $_IEStatus_NoMatch Then
	  MsgBox("title","falla el match",2)
   EndIf
   Local $oTable = _IETableGetCollection($oIE,20)
   Local  $aTableData = _IETableWriteToArray($oTable, True)
   Local $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
   Local $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.
   Local $capitalCero = False
   ;_ArrayDisplay($aTableData)

   ;Hacemos un recorrido y verificamos que contenga "0.00" en la tabla. Si tiene 0.00 devuelve true.
   If $iRows > 2 Then
	  Local $totalPagos = 0
	  Local $rows = $iRows -1
	  For $i = 0 To $iRows -1
		 For $j = 0 To $iCols -1
			if $j = 4 Then
			   Local $Capital = $aTableData[$i][$j]
			   if $Capital <> "Capital" Then
				  if $Capital == "0,00" Then
					 $capitalCero = True
				  EndIf
			   EndIf
			EndIf
		 Next
	  Next
   EndIf


   If $capitalCero == True Then
	  If $iRows > 2 Then
		 Local $totalPagos = 0
		 Local $rows = $iRows -1
		 For $i = 0 To $iRows -1
			For $j = 0 To $iCols -1
			   if $j = 4 Then
				  Local $Capital = $aTableData[$i][$j]
				  if $Capital <> "Capital" Then
					 if $Capital == "0,00" Then
						_clickInImage($importeCobroPagoParcial)
						send("{TAB}")
						;send ("{SPACE}")
						Sleep(1000)
						send("{TAB}")
						Sleep(1000)
						send("{TAB}")
						Sleep(1000)
						send("{TAB}")
						Sleep(1000)
						send ("{SPACE}")
						Sleep(1000)
						send("{TAB}")
						Sleep(1000)
						send("{TAB}")
						Sleep(1000)
						Send($importe_emerix)
						Sleep(5000)
						_clickInImage($otrosConceptos)
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
						Return $capitalCero
					 Else
						_clickInImage($importeCobroPagoParcial)
						send("{TAB}")
						send ("{SPACE}")
						Sleep(1000)
						send("{TAB}")
						Sleep(1000)
						send("{TAB}")
						Sleep(1000)
						Send($importe_emerix)
						 Sleep(5000)
						_clickInImage($otrosConceptos)
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
						Return $capitalCero
					 EndIf
				  EndIf

			   EndIf
			Next
		 Next
	  EndIf
   EndIf

   Return $capitalCero
EndFunc






Func _porrateo($tipoDePago)

   ;click en check de Select
	  Sleep(5000)

	  $oIE = _IEAttach("Ingreso de Cobros","title",2)
	  Local $oTable = _IETableGetCollection($oIE,20)
	  Local  $aTableData = _IETableWriteToArray($oTable, True)
	  Local $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
	  Local $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.

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
	  Sleep(1000) ; Sleep to give tooltip time to display
	  if $tipoDePago == "Cancelatorio" Then
			$oIE = _IEAttach("Ingreso de Cobros","title",2)
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
					 ToolTip($totalPagos, 0, 0)
					 Sleep(1000) ; Sleep to give tooltip time to display
					 EndIf
				  Next
			   Next
			   ToolTip("porrateo", 0, 0)
			   Sleep(3000)
			   _clickInImage($importeCobroCancelacion)
			   Sleep(1000)
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
			   _clickInImage($otrosConceptos)
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
			MsgBox("","",$iRows)
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
			   _clickInImage($importeCobroPagoParcial)
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
			   _clickInImage($otrosConceptos)
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

			   Sleep(3000)
			   If imageExists($nopuedeingresarimportevalorcero,10) Then
				  return "no se puede ingresar importe con valor cero"
			   EndIf
			EndIf
	  EndIf

   EndFunc


Func _AdmDeTerceros($urlEntorno, $titleAttach)
     $oIE = _IEAttach($titleAttach)
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
 EndFunc


Func _CalcularExcedente()
   	  $oIE = _IEAttach("Ingreso de Cobros","title",2)
	  Local $oTable = _IETableGetCollection($oIE,20)
	  Local  $aTableData = _IETableWriteToArray($oTable, True)
	  Local $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
	  Local $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.
	  Local $totalPagos = 0
	  Local $rows = $iRows -1

	  For $i = 0 To $iRows -1
		 For $j = 0 To $iCols -1
			if $j = 5 Then
			   Local $Intereses = $aTableData[$i][$j]
			   if $Intereses <> "Intereses" Then
				  StringReplace($Intereses,",",".")
				  Number($Intereses)
				  Local $excedente = $importe_emerix - $Intereses
				  if $excedente > 0 Then
					 Return $excedente
				  EndIf
				  If $excedente <= 0 Then
					 Return "No Excedente"
				  EndIf
			   EndIf
			EndIf
		 Next
	  Next
EndFunc


Func _login($urlEntorno)
   Local $user = "user1"
   Local $pass = "user1"
      ;cierro todas las instancias de IE.
   RunWait('taskkill /F /IM "iexplore.exe"')

   ;Ingreso al sitio
    Local $tipoDePago = $tipo_actividad;"Cancelatorio"
    If $EntornoPrmtr = "DEV" Then
	  Local $oIE = _IECreate($urlEntorno&"/tandem/login/source.asp")
	  Sleep(2000)
	  _IELoadWait($oIE)
	  Sleep(2000)

	  if imageExists($seleccioneUbicacion, 10) then
		 Send("{F11}")
	  EndIf

	  ;///////CLICK A TEST///////////
	  Sleep(10000)
	  $oIE = _IEAttach("Seleccione Ubicación")
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
	  	If @error Then
		  return False
		EndIf

	  Sleep(4000)
	  ;///////Ingresamos usuario y contraseña///////////
	  send($user)
	  send ("{TAB}")
	  send($pass)
	  send ("{TAB}")
	  Send ("{ENTER}")
	  Sleep(10000)


   ;//////////ENTORNO A PROD//////////////
    ElseIf $EntornoPrmtr = "PROD" Then
	  Local $oIE = _IECreate($urlEntorno&"/tandem/login/source.asp")
	  _IELoadWait($oIE)
	  Sleep(6000)
	  if imageExists($seleccioneUbicacion, 10) then
		 Send("{F11}")
	  EndIf
	  Sleep(10000)
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
   EndIf




EndFunc


Func _clickInImage($imagePath)
   if imageExists($imagePath, 10) then
	  clickOnImage($imagePath)
	  FileWrite($ejecucionLog, $fecha_ejecucion & "  "& @CRLF)
   Else
	   FileWrite($error_Log, $fecha_ejecucion & " la imagen "&$imagePath& " no encontro el elemento para hacer click "& @CRLF)
	   $status= "failure"
	   Local $mensajeError = $fecha_ejecucion & " la imagen "&$imagePath& " no encontro el elemento para hacer click "
	   ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
	   ;RunWait('taskkill /F /IM "iexplore.exe"')
	   ;Exit
   EndIf
EndFunc

Func _filtroEnAdmTerceros()
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
EndFunc

Func _verificarTipoDePago()
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
	  sleep (8000)
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

	  ;hacemos click en N° Cliente
	  $oIE = _IEAttach("Revisión de Cuentas")
	  Local $oTable = _IETableGetCollection($oIE,28)
	  Local  $aTableData = _IETableWriteToArray($oTable, True)
	  ;_ArrayDisplay($aTableData,'sometitle')
	  Local $iRows = UBound($aTableData, $UBOUND_ROWS) ; Total number of rows.
	  Local $iCols = UBound($aTableData, $UBOUND_COLUMNS) ; Total number of columns.
	  Sleep(5000)

	  For $i = 0 To $iRows -1
		 For $j = 0 To $iCols -1
			if $j = 8 Then
			   Local $Cartera = $aTableData[$i][$j]
					 $oDivs = _IETagNameGetCollection($oIE, "div")
					 Sleep(5000)
					 For $oDiv In $oDivs -1
						   If IsObj($oDiv) Then
							  If $oDiv.className = "TablaDinamicaFormatoLink" Then
									_IEAction($oDiv, "click")
									If @error Then
										Return False
									EndIf
									ExitLoop
							  EndIf
						   EndIf
					 Next
			EndIf
		 Next
	  Next




	 ; _clickInImage($nfComafi2)
	  ;_clickInImage($962)
	  Sleep(7000)


	  ;verificamos si aparece el cartel de cuenta lockeada
	  if imageExists($cuentaLockeada, 10) then
	  _clickInImage($cuentaLockeada)
		 if imageExists($cotinuarCuentaLockeada,10) Then
			_clickInImage($cotinuarCuentaLockeada)
		 EndIf
	  EndIf
	  Sleep(7000)

	  ;Addins
	  while Not imageExists($addins, 10)
		 Sleep(1000)
		 Local $contador = 1
		 if $contador > 120 Then
			$status= "failure"
			Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$addins
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			return "Error en encontrar imagen _verificarTipoDePago "
		 EndIf
		 $contador = $contador+1
	  WEnd
	  _clickInImage($addins)
	  ;Cobros
	  while Not imageExists($cobros, 10)
		 Sleep(1000)
		 Local $contador = 1
		 if $contador > 45 Then
			$status= "failure"
			Local $mensajeError =  $fecha_ejecucion & " no existe la imagen "&$cobros
			ConsoleWrite('{"state":"' & $status & '", "data":["' & $mensajeError & '"]}')
			return "Error en encontrar imagen _verificarTipoDePago"
		 EndIf
		 $contador = $contador+1
	  WEnd
	  _clickInImage($cobros)

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
		 ;MsgBox("aaa",$cuentaFinalDeImportes,$cuentaFinalDeImportes)
		 ToolTip($cuentaFinalDeImportes, 0, 0)
		 Sleep(1000) ; Sleep to give tooltip time to display

		 if $cuentaFinalDeImportes > 0 Then
			$tipoDePago = "Pago Parcial"
		 EndIf

		 if $cuentaFinalDeImportes < 0 Then
		   $tipoDePago = "Cancelatorio"
		EndIf

	  Else

		 $tipoDePago = "Cancelatorio"

	  EndIf
	  Sleep(2000)
	  Send("^w")
	  ToolTip($tipoDePago, 0, 0)
	  Sleep(1000) ; Sleep to give tooltip time to display
	  Return $tipoDePago
EndFunc


Func _ingresoACobros($tipoDePago)
	  ;$ingresoDeCobros
	  if imageExists($distribucion,10) then
		 _clickInImage($distribucion)
		 Sleep(1000)
		 if imageExists($nocontribuirpagoanulado,10) Then
			Return "fail"
		 EndIf
	  EndIf

	  Sleep(2000)


	  MouseMove(787, 420)
	  MouseClick($MOUSE_CLICK_LEFT)

	  Sleep(1000)
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


	  ;click en boton distribuir
	  ;Ingreso de Cobros
	  ;click en check de Select
	  Sleep(5000)
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
			   ;_IEAction($distribuir, "click")
		 EndIf
	  Next
	  Sleep(2000)
EndFunc

Func _verificacionDeValidez()
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
	  _clickInImage($procesadoPagosCorrectamente)
	   Send("{ENTER}")
	   Sleep(3000)
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
