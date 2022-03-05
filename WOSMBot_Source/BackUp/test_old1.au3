#include <GUIConstantsEx.au3>
#include <GuiListBox.au3>
#include <MsgBoxConstants.au3>
#include <WindowsConstants.au3>
#include <Array.au3>
#include <Constants.au3>
#include <WinAPI.au3>
#include <AutoItConstants.au3>
#include <StringConstants.au3>
#include <WinAPISysWin.au3>
#include <Date.au3>
#include <GUIToolTip.au3>
#include <GUIConstantsEx.au3>
#include <Process.au3>
#include <File.au3>
#include <APISubiektWin.au3>

czySubiektUruchomiony2()

Func czySubiektUruchomiony2()
$hWnd = MWinGetHandle()
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Wydania magazynowe", "ins_hyperlink", ", o stanie:", "Static")
			$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
			ControlFocus($uchwyt, "", "")
			ControlClick($uchwyt, "", "")
	$hWnd = zlapSubiekta()

	$uchwyt = WinGetHandle("[CLASS:Afx:58480000:800]");ControlGetHandle("","Afx:58480000:800",2001)

	While $uchwyt <> 0

	$uchwyt2 = _WinAPI_GetWindow($uchwyt, $GW_CHILD)
			If _WinAPI_GetClassName($uchwyt2) = "ListBox" AND _WinAPI_GetDlgCtrlID($uchwyt2) = 2001 Then
				ConsoleWrite($uchwyt & " " & _WinAPI_GetClassName($uchwyt) & " " & _WinAPI_GetDlgCtrlID($uchwyt) & @LF)
				ConsoleWrite("__ " & $uchwyt2 & " " & _WinAPI_GetClassName($uchwyt2) & " " & _WinAPI_GetDlgCtrlID($uchwyt2) & @LF)
			EndIf
	WEnd
EndFunc   ;==>czyS

