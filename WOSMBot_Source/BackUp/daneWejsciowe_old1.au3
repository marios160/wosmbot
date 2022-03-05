#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <_sql.au3>

Opt("GUIOnEventMode", 1)
#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("Form1", 326, 149, -1011, 155)
GUISetOnEvent($GUI_EVENT_CLOSE, "Form1Close")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "Form1Minimize")
GUISetOnEvent($GUI_EVENT_MAXIMIZE, "Form1Maximize")
GUISetOnEvent($GUI_EVENT_RESTORE, "Form1Restore")
$Label1 = GUICtrlCreateLabel("Serwer:", 8, 8, 49, 20)
GUICtrlSetFont($Label1, 10, 400, 0, "MS Sans Serif")
GUICtrlSetOnEvent($Label1, "Label1Click")
$fldSerwer = GUICtrlCreateInput("", 88, 8, 193, 21)
GUICtrlSetOnEvent($fldSerwer, "fldSerwerChange")
$Label2 = GUICtrlCreateLabel("Uzytkownik:", 8, 32, 74, 20)
GUICtrlSetFont($Label2, 10, 400, 0, "MS Sans Serif")
GUICtrlSetOnEvent($Label2, "Label2Click")
$Label3 = GUICtrlCreateLabel("Haslo:", 8, 56, 43, 20)
GUICtrlSetFont($Label3, 10, 400, 0, "MS Sans Serif")
GUICtrlSetOnEvent($Label3, "Label3Click")
$Label4 = GUICtrlCreateLabel("Podmiot:", 8, 80, 57, 20)
GUICtrlSetFont($Label4, 10, 400, 0, "MS Sans Serif")
GUICtrlSetOnEvent($Label4, "Label4Click")
$fldUzytkownik = GUICtrlCreateInput("", 88, 32, 193, 21)
GUICtrlSetOnEvent($fldUzytkownik, "fldUzytkownikChange")
$fldHaslo = GUICtrlCreateInput("", 88, 56, 193, 21)
GUICtrlSetOnEvent($fldHaslo, "fldHasloChange")
$fldPodmiot = GUICtrlCreateInput("", 88, 80, 193, 21)
GUICtrlSetOnEvent($fldPodmiot, "fldPodmiotChange")
$btnSprawdzPolaczenie = GUICtrlCreateButton("Sprawdz polaczenie", 8, 112, 123, 25)
GUICtrlSetOnEvent($btnSprawdzPolaczenie, "btnSprawdzPolaczenieClick")
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	Sleep(100)
WEnd

Func btnSprawdzPolaczenieClick()
	Local $ServerAddress = "XERO\INSERTGT"
Local $ServerUserName = "sa"
Local $ServerPassword = ""
Local $DatabaseName = "Arctus_Test_BotSM"


_SQL_RegisterErrorHandler() ;register the error handler to prevent hard crash on COM error
$OADODB = _SQL_Startup()
If $OADODB = $SQL_ERROR Then MsgBox(0 + 16 + 262144, "Error", _SQL_GetErrMsg())
If _sql_Connect(-1, $ServerAddress, $DatabaseName, $ServerUserName, $ServerPassword) = $SQL_ERROR Then
	MsgBox(0 + 16 + 262144, "Blad", "Brak polaczenia z serwerem")
	_SQL_Close()
	Exit
EndIf
Local $fullSQL
If _Sql_GetTableAsString(-1, "SELECT [mag_Symbol],[mag_Nazwa] FROM [Arctus_Test_BotSM].[dbo].[sl_Magazyn]", $fullSQL) = $SQL_OK Then
	ConsoleWrite($fullSQL)
Else
	MsgBox(0 + 16 + 262144, "SQL Error", _SQL_GetErrMsg())
EndIf

EndFunc
Func fldHasloChange()

EndFunc
Func fldPodmiotChange()

EndFunc
Func fldSerwerChange()

EndFunc
Func fldUzytkownikChange()

EndFunc
Func Form1Close()
Exit
EndFunc
Func Form1Maximize()

EndFunc
Func Form1Minimize()

EndFunc
Func Form1Restore()

EndFunc
Func Label1Click()

EndFunc
Func Label2Click()

EndFunc
Func Label3Click()

EndFunc
Func Label4Click()

EndFunc
