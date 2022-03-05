;@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
;@#####################################################################################################################@
;@###############################@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@##################################@
;@###############################@													@##################################@
;@###############################@		Program do odkladania i wywolywania			@##################################@
;@###############################@		skutkow magazynowych w programie 			@##################################@
;@###############################@		Subiekt GT									@##################################@
;@###############################@													@##################################@
;@###############################@						Mateusz Błaszczak	2018	@##################################@
;@###############################@						mateuszblaszczakb@gmail.com	@##################################@
;@###############################@													@##################################@
;@###############################@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@##################################@
;@#####################################################################################################################@
;@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>
#include <Date.au3>
Opt("GUIOnEventMode", 1)
#Region ### START Koda GUI section ### Form=
$guiLog = GUICreate("Bot WOSM - LOG", 597, 293, @DesktopWidth - 640, @DesktopHeight - 300, -1, BitOR($WS_EX_TOPMOST, $WS_EX_WINDOWEDGE))
GUISetOnEvent($GUI_EVENT_CLOSE, "guiLogClose")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "guiLogMinimize")
GUISetOnEvent($GUI_EVENT_MAXIMIZE, "guiLogMaximize")
GUISetOnEvent($GUI_EVENT_RESTORE, "guiLogRestore")
$btnPauza = GUICtrlCreateButton("Pauza", 8, 8, 75, 25)
GUICtrlSetOnEvent($btnPauza, "btnPauzaClick")
$btnStop = GUICtrlCreateButton("Stop", 96, 8, 75, 25)
GUICtrlSetOnEvent($btnStop, "btnStopClick")
$edtLog = GUICtrlCreateEdit("", 8, 40, 577, 241,$ES_MULTILINE)
GUICtrlSetData($edtLog, "")
GUICtrlSetOnEvent($edtLog, "edtLogChange")


#EndRegion ### END Koda GUI section ###
Func uruchomLog()
	_GUICtrlEdit_SetLimitText($guiLog, 999999999)

	GUISetState(@SW_SHOW, $guiLog)
EndFunc   ;==>uruchomLog

Func dodajLog($text)
	_GUICtrlEdit_AppendText($edtLog, "[" & _NowDate() & " " & _NowTime() & "] " & $text & @LF)
EndFunc   ;==>dodajLog


Func btnPauzaClick()

EndFunc   ;==>btnPauzaClick
Func btnStopClick()

EndFunc   ;==>btnStopClick
Func edtLogChange()

EndFunc   ;==>edtLogChange
Func guiLogClose()
	Exit
EndFunc   ;==>guiLogClose
Func guiLogMaximize()

EndFunc   ;==>guiLogMaximize
Func guiLogMinimize()

EndFunc   ;==>guiLogMinimize
Func guiLogRestore()

EndFunc   ;==>guiLogRestore


