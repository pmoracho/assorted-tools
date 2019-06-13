;;#######################################################################################
;# Archivo	: runbatch.au3
;# Fecha	: 2019/06/13
;# Autor	: pmoracho@gmail.com
;#######################################################################################
; Ejecucion:
;
;	runbatch.au3 <archivo batch a ejecutar>
;	
; 
; AutoIt: https://www.autoitscript.com/site/autoit/
; Para convertir a EXE: Aut2Exe\Aut2exe.exe (hay versiones de 32 y 64 bits)

#include-once
#include <Constants.au3>
#include <GUIConstantsEx.au3>
#include <Misc.au3>
#include <WindowsConstants.au3>
#include <ProgressConstants.au3>
#include <StaticConstants.au3>

#include <Array.au3>
#include <File.au3>
#include <String.au3>
#include <Word.au3>

;#######################################################################################
;# BEGIN
;#######################################################################################
Global $LogFile
Global $Debug					= False
Global $BatFile
Global $WinTitle    = "Ejecutando proceso.."
Global $iProgressBar
Global $hGUI,$Progress_1,$Icon_2,$Label_3

if $CmdLine[0] < 1 Then
	Echo( "Ejecutar: runbatch <archivo bat a ejecutar>")
	Exit
Else 
	$BatFile 	= $CmdLine[1]
EndIf

if Not FileExists($BatFile) then
	Echo("Error: No existe el archivo: " & $BatFile)
	Exit
Endif

Ejecutar()

Exit

;#######################################################################################
;# END
;#######################################################################################


Func Ejecutar()

    
    create_form($WinTitle, "Se está procesando su solicitud. Aguarde por favor...")
    ;main_loop()

    
EndFunc   ;==>Example

Func RunBatch($BatFile)
    Local $iReturn = RunWait( $BatFile, "", @SW_HIDE)
    WinClose($WinTitle)
    ;MsgBox($MB_SYSTEMMODAL, "", "The return code  was: " & $iReturn)
EndFunc


Func Echo( $strMensajeUsuario )

   $strMensaje = 	"Runbatch - Ejecución de una rachivo por lotes (BAT)" & @CR
   $strMensaje	= 	$strMensaje	& @CR  & $strMensajeUsuario
   MsgBox(0,"Runbatch", $strMensaje)

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


func create_form($WinTitle, $Label)
    $hGUI=GuiCreate($WinTitle, 466, 130, -1, -1, $DS_MODALFRAME)
    $Label_2 = GuiCtrlCreateLabel($Label, 20, 20, 420, 40,bitor($SS_CENTER,0,0))
    $iProgressBar = GuiCtrlCreateProgress(20, 70, 420, 20,  $PBS_MARQUEE)
    GUISetState(@SW_SHOW)

    _ProgressMarquee_Start($iProgressBar) ;Start
    RunBatch($BatFile)
    _ProgressMarquee_Stop($iProgressBar) ;Stop
    GUIDelete($hGUI)
    
endfunc

func main_loop()

    Local $Run = True
    While 1


        $msg = GuiGetMsg()
        Select
            Case $msg = $GUI_EVENT_CLOSE
                ExitLoop
            Case Else
        EndSelect
    WEnd

endfunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _ProgressMarquee_Start
; Description ...: Start the marquee effect
; Syntax ........: _ProgressMarquee_Start($iControlID)
; Parameters ....: $iControlID          - ControlID of a progressbar using the $PBS_MARQUEE style
; Return values .: Success - True
;                  Failure - False
; Author ........: guinness
; Example .......: Yes
; ===============================================================================================================================
Func _ProgressMarquee_Start($iControlID)
    Return GUICtrlSendMsg($iControlID, $PBM_SETMARQUEE, True, 50)
EndFunc   ;==>_ProgressMarquee_Start

; #FUNCTION# ====================================================================================================================
; Name ..........: _ProgressMarquee_Stop
; Description ...: Stop the marquee effect
; Syntax ........: _ProgressMarquee_Stop($iControlID[, $bReset = False])
; Parameters ....: $iControlID          - ControlID of a progressbar using the $PBS_MARQUEE style
;                  $bReset              - [optional] Reset the progressbar, True - Reset or False - Don't reset. Default is False
; Return values .: Success - True
;                  Failure - False
; Author ........: guinness
; Example .......: Yes
; ===============================================================================================================================
Func _ProgressMarquee_Stop($iControlID, $bReset = False)
    Local $bReturn = GUICtrlSendMsg($iControlID, $PBM_SETMARQUEE, False, 50)
    If $bReturn And $bReset Then
        GUICtrlSetStyle($iControlID, $PBS_MARQUEE)
    EndIf

    Return $bReturn
EndFunc   ;==>_ProgressMarquee_Stop

