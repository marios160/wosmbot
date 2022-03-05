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
$guiLog = GUICreate("Bot WOSM - LOG", 597, 293, -1011, 130, -1, BitOR($WS_EX_TOPMOST,$WS_EX_WINDOWEDGE))
GUISetOnEvent($GUI_EVENT_CLOSE, "guiLogClose")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "guiLogMinimize")
GUISetOnEvent($GUI_EVENT_MAXIMIZE, "guiLogMaximize")
GUISetOnEvent($GUI_EVENT_RESTORE, "guiLogRestore")
$btnPauza = GUICtrlCreateButton("Pauza", 8, 8, 75, 25)
GUICtrlSetOnEvent($btnPauza, "btnPauzaClick")
$btnStop = GUICtrlCreateButton("Stop", 96, 8, 75, 25)
GUICtrlSetOnEvent($btnStop, "btnStopClick")
$edtLog = GUICtrlCreateEdit("", 8, 40, 577, 241)
GUICtrlSetData($edtLog, "")
GUICtrlSetOnEvent($edtLog, "edtLogChange")

#EndRegion ### END Koda GUI section ###

Func uruchomLog()
GUISetState(@SW_SHOW,$guiLog)


EndFunc

Func dodajLog($text)
	_GUICtrlEdit_AppendText($edtLog,"[" & _NowDate() & " " &  _NowTime() & "] " & $text & @LF)
EndFunc


Func btnPauzaClick()

EndFunc
Func btnStopClick()

EndFunc
Func edtLogChange()

EndFunc
Func guiLogClose()
	Exit
EndFunc
Func guiLogMaximize()

EndFunc
Func guiLogMinimize()

EndFunc
Func guiLogRestore()

EndFunc


