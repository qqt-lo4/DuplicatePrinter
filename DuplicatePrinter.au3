#Include "UDF/PrintMgr.au3"
#Include <Array.au3>
;~ $aPrinterList = _PrintMgr_EnumPrinter()
;~ _ArrayDisplay($aPrinterList)

; List all installed printers
#Include <Array.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Select Printer", 546, 42, 2545, 469)
$ComboPrinter = GUICtrlCreateCombo("", 8, 8, 233, 21, BitOR($CBS_DROPDOWN,$CBS_AUTOHSCROLL))
$InputNewName = GUICtrlCreateInput("", 248, 8, 233, 21)
$ButtonOK = GUICtrlCreateButton("OK", 488, 8, 49, 21)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Func __PrintMgr_NewPrinterName($sPrinterName)
	Return $sPrinterName & "-Copy"
EndFunc

Func _PrintMgr_DuplicatePrinter($sPrinterName, $sDuplicatedPrinterName)
	Local $_aPrinters = _PrintMgr_EnumPrinterProperties($sPrinterName)
	If (IsArray($_aPrinters) And (UBound($_aPrinters, $UBOUND_ROWS) = 1)) Then
		$_sDriverName = $_aPrinters[0][31]
		$_sPortName = $_aPrinters[0][60]
		Return _PrintMgr_AddPrinter($sDuplicatedPrinterName, $_sDriverName, $_sPortName)
	EndIf
EndFunc

; Fill the combo box
Local $_aPrinters = _PrintMgr_EnumPrinter()
If (IsArray($_aPrinters)) Then
	If ((UBound($_aPrinters) >= 1) And ($_aPrinters[0] > 0)) Then
		Local $_sItems = _ArrayToString($_aPrinters, "|", 1)
		GUICtrlSetData($ComboPrinter, $_sItems)
	EndIf
EndIf

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $ComboPrinter
			Local $_sSelectedPrinter = GUICtrlRead($ComboPrinter)
			GUICtrlSetData($InputNewName, __PrintMgr_NewPrinterName($_sSelectedPrinter))
		Case $GUI_EVENT_CLOSE
			Exit
		Case $ButtonOK
			Local $_sSelectedPrinter = GUICtrlRead($ComboPrinter)
			Local $_sNewPrinterName = GUICtrlRead($InputNewName)
			If (_ArraySearch($_aPrinters, $_sNewPrinterName) = -1) Then
				Local $_bResult = _PrintMgr_DuplicatePrinter($_sSelectedPrinter, $_sNewPrinterName)
				Local $_sMsgBoxResult = ($_bResult = 1) ? "Successful" : "Failure"
				MsgBox(0, "Operation Result", $_sMsgBoxResult)
				Exit
			Else
				MsgBox($MB_OK + $MB_ICONERROR, "Error", "This printer name already exists.")
			EndIf
	EndSwitch
WEnd
