;;#######################################################################################
;# Archivo	: pdfprint.au3
;# Fecha	: 2016/06/09
;# Autor	: pmoracho@gmail.com
;#######################################################################################
; Ejecucion:
;
;	printfile <archivo con lista de pdf´s o tiff> <path al gswin32c.exe>
;	
;	Formato de la lista:
;		archivo,tamaño papel
;	Por ej.
;		archivo1.pdf, a4
;		archivo2.pdf, legal
; 
; AutoIt: https://www.autoitscript.com/site/autoit/
; Para convertir a EXE: Aut2Exe\Aut2exe.exe (hay versiones de 32 y 64 bits)

#include-once
#include <Array.au3>
#include <File.au3>
#include <String.au3>
#include <Word.au3>
#include "SetLocalPrinter.au3"

;#######################################################################################
;# BEGIN
;#######################################################################################
Global Const $Tiff2Pdf			= @ScriptDir & "\tiff2pdf.exe"
Global $FileList[1][4]
Global $PrintList
Global $LogFile
Global $Debug					= False
Global $DefaultPrinter			= ""
Global $DefaultPaperSize		= "A4"
Local  $GsExe					= ""

if $CmdLine[0] < 1 Then
	if $Debug Then 
		$PrintList = "pdfprint.txt"
		$GsExe = "\\PLUSPMDESA\Mecanus\MecanusPM\LOG\..\Tools\GS\bin\gswin32c.exe"
	Else
		echo( "Ejecutar: pdfprint <archivo de la lista> <Path al GhostScript>")
		Exit
	EndIf
Else 
	$PrintList 	= $CmdLine[1]
	if $CmdLine[0] >= 2 Then
		$GsExe 	= $CmdLine[2]
	EndIf
EndIf

if Not FileExists($PrintList) then
	Echo("Error: No existe la lista de archivo a imprimir: " & $PrintList)
	Exit
Endif

if Not FileExists($GsExe) then
	Echo("Error: No se ha encontrado el GhostScript [" & $GsExe & "]")
	Exit
Endif


$DefaultPrinter = SelectDefaultPrinter()
$LogFile =  $PrintList & ".log"
FileDelete($LogFile)

ProcesarLista($PrintList)
ImprimirConGhostScript($GsExe)

Exit
;#######################################################################################
;# END
;#######################################################################################

Func Echo( $strMensajeUsuario )

   $strMensaje = 	"Printfile - Impresión via shellExec de una lista de archivos" & @CR
   $strMensaje	= 	$strMensaje	& @CR  & $strMensajeUsuario
   MsgBox(0,"Printfile", $strMensaje)

EndFunc
 

;#######################################################################################
;# Log
;#######################################################################################
Func FileLog($Data, $FileName = -1, $TimeStamp = False)
    If $FileName == -1 Then $FileName = @ScriptDir & '\Log.txt'
    $hFile = FileOpen($FileName, 1)
    If $hFile <> -1 Then
        If $TimeStamp = True Then $Data = _Now() & ' - ' & $Data
        FileWriteLine($hFile, $Data)
        FileClose($hFile)
    EndIf
EndFunc

;#######################################################################################
;# ProcesarLista
;# Procesa lista e imprime cada archivo
;#######################################################################################
Func ProcesarLista( $strListaFile )
   
	Local	$strFileToPrint
	Local	$TempFile
	Local	$strPaperSize
	Local	$strOrientation
	Local	$strDocument
	Local 	$input[5]

	;##########################################################################
	;# Cargo la lista en un array 
	;##########################################################################
	$i = 0
	$hFile = FileOpen($strListaFile)
	While 1
		; Read the next line
		; And exit the loop when we get to EOF
		$linea = FileReadLine($hFile)
		If @error = -1 Then ExitLoop

		$input = StringSplit($linea, ",")
		$strFileToPrint = StringStripWS($input[1],3)
		$strPaperSize = StringStripWS($input[2],3)
		$strOrientation = StringStripWS($input[3],3)
		$strDocument = StringStripWS($input[4],4)

		If FileExists($strFileToPrint) Then
			$ext = StringLower( StringRight($strFileToPrint,3))
			;##########################################################################
			;# Si es un Tif se convierte a PDF por que no funciona con ShellExecute
			;# la cagada es que las temporales quedan, no se pueden borrar por que
			;# el shellexecute es asincronico
			;##########################################################################
			if $ext = "tif" Then
				$nfile			= formatZeros( "0000", $i + 1 )
				$TempFile		= @TempDir & "\" & $nfile & ".Printfile.pdf" 
				$ConvertJob 	= $Tiff2Pdf & " " & Chr(34) & $strFileToPrint & Chr(34) & " -o " &  $TempFile
				FileDelete($TempFile)
				RunWait($ConvertJob, @WindowsDir, @SW_HIDE)
				;Echo($ConvertJob)
				if FileExists($TempFile) Then
					AddFile($TempFile, $strPaperSize, $strOrientation, $strDocument)
				Else
					Echo("Imposible generar: " & $TempFile)
				EndIf			
		 	Else
				AddFile($strFileToPrint, $strPaperSize, $strOrientation, $strDocument)
		 	EndIf
		else
			Echo( "No existe el archivo " & $strFileToPrint )
		EndIf
		$i = $i + 1
	WEnd
	; Close the file
	FileClose($hFile)
	;Borrar la lista
	if $Debug = False Then
		FileDelete($strListaFile)
	EndIf

EndFunc

;#######################################################################################
;# ImprimirConGhostScript
;# Imprime usando el Adobe reader que debe estar instalado en el equipo
;#######################################################################################
Func ImprimirConGhostScript( $GsExe )

	Local $p
	Local $f
	Local $o
	Local $d
	local $PrintJob
	local $item[2]
	local $w
	local $h
	local $sPapersize
	local $pid
	local $data
	local $linesep
	local $ext

	$linesep = "===============================================================================================================================================================================================================================================================" 

	For	$i = 0 to UBound($FileList) - 1
		$f = $FileList[$i][0]
		$p = $FileList[$i][1]
		$o = $FileList[$i][2]
		$d = $FileList[$i][3]

		; Check: http://ghostscript.com/doc/8.54/Use.htm#Output_resolution
		; Check: http://ghostscript.com/doc/8.54/Use.htm#Known_paper_sizes

		if $p = "legal" Then
			$w = "612"
			$h = "1008"
		else
			$w = "595"
			$h = "842"
		EndIf

		if $o = "landscape" Then
			$sPapersize = " -dDEVICEWIDTHPOINTS=" & $h & " -dDEVICEHEIGHTPOINTS=" & $w & " "
		else
			$sPapersize = " -dDEVICEWIDTHPOINTS=" & $w & " -dDEVICEHEIGHTPOINTS=" & $h & " "
		EndIf

		if $Debug = True Then
			Echo( "Imprimiendo " & $f & "en papel " & $p & "..." )
		EndIf

		if Not FileExists( $f ) Then
			Echo( "No existe el archivo " & $f )
		Else
			$ext = StringLower( StringRight($f,3))

			FileLog($linesep, $LogFile)
			FileLog("Imprimiendo " & $d, $LogFile)
			FileLog($linesep, $LogFile)

			if $ext = "pdf" Then 
				Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
				Local $aPathSplit = _PathSplit($GsExe, $sDrive, $sDir, $sFileName, $sExtension)
				Local $Lib =  $sDrive & $sDir & "..\Lib"
				Local $Font =  $sDrive & $sDir & "..\Fonts"

				EnvSet( "GS_PATH", $sDir )
				EnvSet( "GS_LIB", $Lib )
				EnvSet( "GS_FONT", $Font )

				$PrintJob =	$GsExe & " " & " -dBATCH -dNOPAUSE -dNoCancel " & $sPapersize & " -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -sDEVICE=mswinpr2 " & "-sOutputFile=""%printer%" & $DefaultPrinter & """ -I" & $sDir & " -I" & $Lib & " -I" &  $Font & " " & $f

				FileLog($PrintJob, $LogFile)

				;ClipPut($PrintJob)
				;RunWait($PrintJob, @WindowsDir, @SW_HIDE)
				;RunWait($PrintJob, "", @SW_HIDE)

				$pid = Run(@ComSpec & " /c " & $PrintJob, "", @SW_HIDE, 2)
				$data = ""
				While Not @error
					$data &= StdOutRead($pid)
				Wend

				FileLog("Output:",$LogFile)
				FileLog($data, $LogFile)
				FileLog("", $LogFile)

				if $Debug = True Then
					echo($PrintJob)
				EndIf
			else
				if $ext = "doc" Then
					FileLog("Impresión de documento Word", $LogFile)

					Local $oWord = _Word_Create(False)
					If @error Then 
						FileLog("Error creating a new Word application object. @error = " & @error & ", @extended = " & @extended, $LogFile)
					Else
						Local $oDoc = _Word_DocOpen($oWord, $f, Default, Default, True)

						If Not @error Then
							Local $sActivePrinter = $oDoc.Application.ActivePrinter
							_Word_DocPrint($oDoc)
							If @error Then
								FileLog("Error al imprimir el documento word: " & $f, $LogFile)
							EndIf
							FileLog("Impresión exitosa del documento word: " & $f & " impresora: " & $sActivePrinter, $LogFile)
						Else
							FileLog("Error al abrir el documento word: " & $f & " impresora: " & $sActivePrinter, $LogFile)
						EndIf
					Endif
				Endif
			Endif
		EndIf
		Sleep(3000)
	next
	Sleep(5000)

EndFunc

Func formatZeros($strZeros, $yourInteger)
	return  StringRight($strZeros & $yourInteger, StringLen($strZeros) + 1)
EndFunc

Func AddFile($f, $s, $o, $d)
   If $FileList[0][0]=""  Then
	  $FileList[UBound($FileList) - 1][0] = $f
	  $FileList[UBound($FileList) - 1][1] = $s
	  $FileList[UBound($FileList) - 1][2] = $o
	  $FileList[UBound($FileList) - 1][3] = $d
   else
	  ReDim $FileList[UBound($FileList) + 1][4]
	  $FileList[UBound($FileList) - 1][0] = $f
	  $FileList[UBound($FileList) - 1][1] = $s
	  $FileList[UBound($FileList) - 1][2] = $o
	  $FileList[UBound($FileList) - 1][3] = $d
   endif
EndFunc
