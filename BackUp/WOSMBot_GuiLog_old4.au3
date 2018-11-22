#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
Opt("GUIOnEventMode", 1)
#Region ### START Koda GUI section ### Form=
$WOSMBot_GuiLog = GUICreate("WOSMBot - Status", 335, 200,@DesktopWidth - 640, @DesktopHeight - 300, -1, BitOR($WS_EX_TOPMOST,$WS_EX_WINDOWEDGE))
$Label1 = GUICtrlCreateLabel("Zakoncz - F6", 216, 16, 99, 24)
GUICtrlSetFont($Label1, 12, 400, 0, "MS Sans Serif")
$Label2 = GUICtrlCreateLabel("Pauza - F4", 16, 16, 84, 24)
GUICtrlSetFont($Label2, 12, 400, 0, "MS Sans Serif")
$Group1 = GUICtrlCreateGroup("Status", 8, 40, 313, 145)
$fldStatus7 = GUICtrlCreateLabel("fldStatus7", 24, 160, 292, 17)
$fldStatus6 = GUICtrlCreateLabel("fldStatus6", 24, 144, 292, 17)
$fldStatus5 = GUICtrlCreateLabel("fldStatus5", 24, 128, 292, 17)
$fldStatus4 = GUICtrlCreateLabel("fldStatus4", 24, 112, 292, 17)
$fldStatus3 = GUICtrlCreateLabel("fldStatus3", 24, 96, 292, 17)
$fldStatus2 = GUICtrlCreateLabel("fldStatus2", 24, 80, 292, 17)
$fldStatus1 = GUICtrlCreateLabel("fldStatus1", 24, 64, 292, 17)
GUICtrlCreateGroup("", -99, -99, 1, 1)
#EndRegion ### END Koda GUI section ###


Func uruchomLog()

	GUISetState(@SW_SHOW,$WOSMBot_GuiLog)
EndFunc
Func dodajStatus($text)

	GUICtrlSetData($fldStatus7,GUICtrlRead($fldStatus6))
	GUICtrlSetData($fldStatus6,GUICtrlRead($fldStatus5))
	GUICtrlSetData($fldStatus5,GUICtrlRead($fldStatus4))
	GUICtrlSetData($fldStatus4,GUICtrlRead($fldStatus3))
	GUICtrlSetData($fldStatus3,GUICtrlRead($fldStatus2))
	GUICtrlSetData($fldStatus2,GUICtrlRead($fldStatus1))
	GUICtrlSetData($fldStatus1,$text)
EndFunc
