#include <Array.au3> 
#include <GUIConstants.au3>

Global $regwriteprinters

Global $printers
Global $printercount
Global $regprinters

Global $currentprinter
Global $printer_list[1]
Global $printer_list_ext[1]
global $printer_radio_array[1]
global $imprimante
global $Finish
;SelectDefaultPrinter()

;===============================================================================
; Function Name:    SelectDefaultPrinter()
; Description:      
; Parameter(s):     
; Requirement(s):   none.
; Return Value(s):  String
; Author(s):        pmoracho@gmail.com
;===============================================================================
Func SelectDefaultPrinter()

   Local $msg
   Local $Printer

   $regprinters = "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Devices"
   $regwriteprinters = "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows"
   $currentprinter = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Windows\","Device")

   dim $i = 1
   dim $erreur_reg = false
   while not $erreur_reg
	   $imprimante = RegEnumVal($regprinters, $i)
	   $erreur_reg = @error
	   if not $erreur_reg then
		   _ArrayAdd($printer_list,$imprimante)
		   _ArrayAdd($printer_list_ext, $imprimante & "," & RegRead($regprinters,$imprimante))
	   endif
	   $i = $i + 1
   wend
   _ArrayDelete($printer_list,0) 
   _ArrayDelete($printer_list_ext,0) 

   if ubound($printer_list) >= 2 then ;if more that 2 printers available, we show the dialog
	   dim $groupheight = (ubound($printer_list) + 1) * 30
	   dim $guiheight = $groupheight + 50
	   dim $buttontop = $groupheight + 20
	   ;Opt("GUIOnEventMode", 1)
	   GUICreate("Seleccione la impresora a utilizar", 400, $guiheight)
	   dim $font = "Verdana"
	   GUISetFont (10, 400, 0, $font)
	   ;GUISetOnEvent($GUI_EVENT_CLOSE, "CLOSEClicked")
	   GUISetFont (10, 400, 0, $font)
	   GUICtrlCreateGroup("",10, 10, 380, $groupheight)
	   dim $position_vertical = 0
	   For $i=0 to ubound($printer_list)-1 step 1
		   GUISetFont (10, 400, 0, $font)
		   $position_vertical = $position_vertical + 30
		   $radio = GUICtrlCreateRadio ($printer_list[$i], 20, $position_vertical, 350, 20)
		   _ArrayAdd($printer_radio_array,$radio)
		   If $currentprinter = $printer_list_ext[$i] Then
			   GUICtrlSetState ($radio,$GUI_CHECKED)
		   endif
	   next
	   _ArrayDelete($printer_radio_array,0)
	   GUISetFont (10, 400, 0, $font)
	   $okbutton = GUICtrlCreateButton("OK", 10, $buttontop, 50, 25)
	   GUICtrlSetState ( $okbutton, $GUI_FOCUS )
	   GUICtrlSetOnEvent($okbutton, "SelectDefaultPrinter_OKButton")
	   ;set enter
	   GUISetState ()
	   $Finish = False
	   While 1
		 $msg = GUIGetMsg()
		 Select
			Case $msg = $GUI_EVENT_CLOSE
			   Exit
			Case $msg = $okbutton
			   for $i=0 to ubound($printer_radio_array)-1 step 1
				  If GUICtrlRead($printer_radio_array[$i])=1 Then
					 RegWrite($regwriteprinters, "Device", "REG_SZ", $printer_list_ext[$i])
					 $Printer = $printer_list[$i]
			   		 ExitLoop
				  endif
			   next
			   ExitLoop
		 EndSelect

	   WEnd
	   GUIDelete()
	   return $Printer
   Else
		return $printer_list[1]
   endif
   
endfunc

