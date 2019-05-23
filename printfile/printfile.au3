;#######################################################################################
;# Archivo	: Printfile.vbs
;# Que hace?: Impresión via shellExec de una lista de archivos pasados por parámetros
;# Fecha	: 2013/09/11
;# Autor	: pmoracho
;#######################################################################################
; Ejecuion:
;
; 1. Usando el Adobe Acrobat Reader local o aplicación instalada para manejar acrhivos PDF
;		printfile <archivo con lista de pdf´s o tiff>
; 
; 2. Usando el GhostScript
;		printfile <archivo con lista de pdf´s o tiff> <path al gswin32c.exe>
;
; AutoIt: https://www.autoitscript.com/site/autoit/
; Para convertir a EXE: Aut2Exe\Aut2exe.exe (hay versiones de 32 y 64 bits)

#include-once
#include <Array.au3>
#include <File.au3>
#include "SetLocalPrinter.au3"

;#######################################################################################
;# BEGIN
;#######################################################################################
; 1. Chequear parametros
; 2. Verificar si la lista existe
; 3. Solicitar impresora
; 4. Recorrer la lista 
; 5.1 Si no existe el archivo Mensaje de error	
; 5.2 Si existe imprimirlo

Global Const $Tiff2Pdf			= @ScriptDir & "\tiff2pdf.exe"
Global $FileList[1]
Global $PrintList
Global $Debug					= False
Global $DefaultPrinter			= ""

if $CmdLine[0] < 1 Then
   if $Debug Then 
	  $PrintList = "c:\tmp\prueba.txt"
   Else
	  echo( "Ejecutar: Printfile <archivo de la lista> <GhostScript command si usamos esta herramienta>")
	  Exit
   EndIf
Else 
   $PrintList 	= $CmdLine[1]
EndIf

if Not FileExists($PrintList) then
	Echo("Error: No existe la lista de archivo a imprimir: " & $PrintList)
	Exit
Endif

$DefaultPrinter = SelectDefaultPrinter()

ProcesarLista($PrintList)

if $CmdLine[0] = 1 Then
	ImprimirConAdobe()
else
	ImprimirConGhostScript($CmdLine[2])
Endif

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
;# ProcesarLista
;# Procesa lista e imprime cada archivo
;#######################################################################################
Func ProcesarLista( $strListaFile )
   
	Local	$strFileToPrint

	;##########################################################################
	;# Cargo la lista en un array 
	;##########################################################################
	$i = 0
	$hFile = FileOpen($strListaFile)
	While 1
		; Read the next line
		$strFileToPrint = FileReadLine($hFile)

		; And exit the loop when we get to EOF
		If @error = -1 Then ExitLoop

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
				RunWait($ConvertJob, @WindowsDir, @SW_HIDE)
				;Echo($ConvertJob)
				if FileExists($TempFile) Then
					AddFile($TempFile)
				Else
					Echo("Imposible generar: " & $TempFile)
				EndIf			
		 	Else
				AddFile($strFileToPrint)
		 	EndIf
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

	For	$f in $FileList
		if $Debug = True Then
			Echo( "Imprimiendo " & $f & "..." )
		EndIf
		if Not FileExists( $f ) Then
			Echo( "No existe el archivo " & $f )
		Else
			Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
			Local $aPathSplit = _PathSplit($GsExe, $sDrive, $sDir, $sFileName, $sExtension)
			Local $Lib =  $sDrive & $sDir & "..\Lib"
			Local $Font =  $sDrive & $sDir & "..\Fonts"

			;@set GS_ROOT=\\pluspmdesa\Mecanus\MecanusPM\Tools\GS
			;@set GS_PATH=%GS_ROOT%\bin
			;@set GS=%GS_PATH%\gswin32c.exe
			;@set GS_LIB=%GS_ROOT%\Lib
			;@set GS_FONT=%GS_ROOT%\Fonts
			;@set PRINTERNAME=%1
			;@set FILENAME=%2
			;%GS% -dQUIET -dBATCH -dNOPAUSE -dNoCancel -sPAPERSIZE=a4 -sDEVICE=mswinpr2 -sOutputFile=\\spool\%printer%%PRINTERNAME% -I%GS_PATH% -I%GS_LIB% -I%GS_FONT% %FILENAME%

			EnvSet( "GS_PATH", $sDir )
			EnvSet( "GS_LIB", $Lib )
			EnvSet( "GS_FONT", $Font )

			$PrintJob =	$GsExe & " " & "-dQUIET -dBATCH -dNOPAUSE -dNoCancel -sPAPERSIZE=a4 -sDEVICE=mswinpr2 " & "-sOutputFile=""%printer%" & $DefaultPrinter & """ -I" & $sDir & " -I" & $Lib & " -I" &  $Font & " " & $f

			ClipPut($PrintJob)
			RunWait($PrintJob, @WindowsDir, @SW_HIDE)
			if $Debug = True Then
				echo($PrintJob)
			EndIf
		EndIf
		Sleep(3000)
	next
	Sleep(5000)

EndFunc

;#######################################################################################
;# ImprimirConAdobe
;# Imprime usando el Adobe reader que debe estar instalado en el equipo
;#######################################################################################
Func ImprimirConAdobe( )

   If WinExists ("[CLASS:AcrobatSDIWindow]") Then
	  $AdobeYaAbierto	= True
   Else
	  $AdobeYaAbierto	= False
   EndIf

   $CerrarAdobe	= False
   For	$f in $FileList
	  if $Debug = True Then
		 Echo( "Imprimiendo " & $f & "..." )
	  EndIf
	  if Not FileExists( $f ) Then
		 Echo( "No existe el archivo " & $f )
	  Else
		if _FilePrint($f, @SW_HIDE) = 1 Then
			$ext = StringLower( StringRight($f,3))
			if $ext = "pdf" Then 
			   $CerrarAdobe	= True
			EndIf
		 Else
		 	echo("Error al imprimir el archivo: " & $f)
		 EndIf
	  EndIf
	  Sleep(3000)
   next
   Sleep(5000)
   ;if $CerrarAdobe = True Then Echo( "$CerrarAdobe = True" )
   ;if $AdobeYaAbierto = True Then Echo( "$AdobeYaAbierto = True" )
   if $CerrarAdobe and not $AdobeYaAbierto Then
	  WinWait("[CLASS:AcrobatSDIWindow]", "", 10)
	  WinClose("[CLASS:AcrobatSDIWindow]", "")
	  WinKill("[CLASS:AcrobatSDIWindow]", "")
   endif

EndFunc


Func formatZeros($strZeros, $yourInteger)
	return  StringRight($strZeros & $yourInteger, StringLen($strZeros) + 1)
EndFunc

Func AddFile($f)
   If $FileList[0]=""  Then
	  $FileList[UBound($FileList) - 1] = $f
   else
	  ReDim $FileList[UBound($FileList) + 1]
	  $FileList[UBound($FileList) - 1] = $f
   endif
EndFunc
