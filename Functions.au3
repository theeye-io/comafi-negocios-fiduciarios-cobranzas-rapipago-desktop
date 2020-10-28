#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------
#include-once
#include "ImageSearch.au3"
#include <Array.au3>
#include <File.au3>
#include <ScreenCapture.au3>
#include <Clipboard.au3>
#include <Misc.au3>
#include <FileConstants.au3>
#include <WinAPIFiles.au3>
#include <date.au3>

; Script Start - Add your code below here
Global $debug= false
Global $globalRunningStatus ="c:\theeye\MEP\logs\status.log"

Func clickOnImage($image, $timeout = 5, $xOffset = 0, $yOffset=0)

	$x1=0
	$y1=0
	Local $result = _WaitForImageSearch($image,$timeout,1,$x1,$y1,15)

	If ($result=1) Then
		MouseMove($x1 + $xOffset, $y1 + $yOffset)
		MouseClick("main")
		Sleep(2000)
		Return True
	Else
		Return False
	Endif

EndFunc

Func imageExists($image, $timeout = 5)

	$x1=0
	$y1=0
	Local $result = _WaitForImageSearch($image,$timeout,1,$x1,$y1,15)

	If ($result=1) Then
		Return True
	Else
		Return False
	Endif

EndFunc

Func createDirIfNotExists($dir)

	 If Not (FileExists($dir)) Then
			ConsoleWrite('Creating directory ' & $dir)
			DirCreate($dir)
			If Not(FileExists($dir)) Then
				 ConsoleWrite('Failed to create directory' & @LF)
				 ConsoleWrite('failure')
				 Exit
			EndIf
	 EndIf

EndFunc

;Func stringRemoveSpaces($str)
;
;	 Local $s = StringStripWS($str, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
;	 Return $s
;
;EndFunc


Func GetQueueFiles($queuePath)
    ; List all the queued files to process
    Local $aFileList = _FileListToArray($queuePath, "*", 1)
    If @error = 1 Then
        MsgBox(Null,"", "Path was invalid.")
        Exit
    EndIf
    If @error = 4 Then
        MsgBox(Null, "", "No file(s) were found.")
        Exit
    EndIf
    ; Display the results returned by _FileListToArray.
    ;_ArrayDisplay($aFileList, "$aFileList")

	Return $aFileList
EndFunc

Func saveErrorLog($description, $errorPath, $statusPath = $globalRunningStatus)

	Local $hFile = FileOpen($errorPath & "\MEPExecution.log", 1)
	Local $hStatusFile = FileOpen($statusPath, 2)

	if @error Then
	EndIf
	; ERROR

	If ($debug) then
		ConsoleWrite($description & @CRLF)
	Endif

	; Send("{ENTER}")
	Sleep(500)
	; Sacamos screenshot y guardamos en array de errores
	Local $fileErrorScreenshot = $errorPath & "\" & "error_" & "_" & @MDAY & "-" & @MON & "-" & @YEAR & "_" & @HOUR & "-" & @MIN & "-" & @MSEC & ".jpg"
	_ScreenCapture_Capture($fileErrorScreenshot)
	_FileWriteLog($hFile, $description & " File: " & $fileErrorScreenshot)
	FileWrite($hStatusFile, "failure")

	if @error Then
	EndIf
	Sleep(500)
	FileClose($hFile)
	; Send("{TAB}")
	; Send("{ENTER}")
EndFunc

Func saveInfoLog($description, $errorPath, $statusPath = $globalRunningStatus)

	Local $hFile = FileOpen($errorPath & "\MEPExecution.log", 1)
	Local $hStatusFile = FileOpen($statusPath, 2)

	if @error Then
	EndIf
	; ERROR

	If ($debug) then
		ConsoleWrite($description & @CRLF)
	Endif

	; Send("{ENTER}")
	Sleep(500)
	; Sacamos screenshot y guardamos en array de errores
	Local $fileErrorScreenshot = $errorPath & "\" & "error_" & "_" & @MDAY & "-" & @MON & "-" & @YEAR & "_" & @HOUR & "-" & @MIN & "-" & @MSEC & ".jpg"
	_FileWriteLog($hFile, "INFO: " & $description)

	if @error Then
	EndIf
	Sleep(500)
	FileClose($hFile)
	; Send("{TAB}")
	; Send("{ENTER}")
EndFunc

Func getSelectedText()
    ClipPut("")
    Local $copied = ""
    sleep(500)
    Send("^c")
    Sleep(2000)
    $copied = ClipGet()
    sleep(2250)
    Return $copied
EndFunc