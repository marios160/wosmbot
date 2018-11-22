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
#include <ColorConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <_sql.au3>
#include <File.au3>




Opt("GUIOnEventMode", 1)
#Region ### START Koda GUI section ### Form=
$okienko = GUICreate("Wprowadź potrzebne dane", 392, 198, -1, -1)
GUISetOnEvent($GUI_EVENT_CLOSE, "oknoClose")

$Label1 = GUICtrlCreateLabel("Serwer:", 8, 8, 49, 20)
GUICtrlSetFont($Label1, 10, 400, 0, "MS Sans Serif")

$fldSerwer = GUICtrlCreateInput(EnvGet("USERDOMAIN") & "\INSERTGT", 184, 8, 193, 21)
GUICtrlSetOnEvent($fldSerwer, "fldSerwerChange")
$Label2 = GUICtrlCreateLabel("Uzytkownik (Baza danych):", 8, 32, 170, 20)
GUICtrlSetFont($Label2, 10, 400, 0, "MS Sans Serif")
$Label3 = GUICtrlCreateLabel("Haslo (Baza danych):", 8, 56, 139, 20)
GUICtrlSetFont($Label3, 10, 400, 0, "MS Sans Serif")
$Label4 = GUICtrlCreateLabel("Podmiot:", 8, 80, 57, 20)
GUICtrlSetFont($Label4, 10, 400, 0, "MS Sans Serif")
$Label5 = GUICtrlCreateLabel("Uzytkownik:", 8, 104, 81, 20)
GUICtrlSetFont($Label5, 10, 400, 0, "MS Sans Serif")
$Label6 = GUICtrlCreateLabel("Haslo:", 8, 128, 57, 20)
GUICtrlSetFont($Label6, 10, 400, 0, "MS Sans Serif")
$fldUzytkownikBaza = GUICtrlCreateInput("sa", 184, 32, 193, 21)
GUICtrlSetOnEvent($fldUzytkownikBaza, "fldUzytkownikBazaChange")
$fldHasloBaza = GUICtrlCreateInput("", 184, 56, 193, 21)
GUICtrlSetOnEvent($fldHasloBaza, "fldHasloBazaChange")
$fldPodmiot = GUICtrlCreateInput("", 184, 80, 193, 21)
GUICtrlSetOnEvent($fldPodmiot, "fldPodmiotChange")
$fldUzytkownik = GUICtrlCreateInput("Szef", 184, 104, 193, 21)
GUICtrlSetOnEvent($fldUzytkownik, "fldUzytkownikChange")
$fldHaslo = GUICtrlCreateInput("", 184, 128, 193, 21)
GUICtrlSetOnEvent($fldHaslo, "fldHasloChange")
$btnSprawdzPolaczenie = GUICtrlCreateButton("Sprawdz polaczenie", 8, 160, 123, 25, $BS_DEFPUSHBUTTON)
GUICtrlSetOnEvent($btnSprawdzPolaczenie, "btnSprawdzPolaczenieClick")
$labStatus = GUICtrlCreateLabel("", 136, 168, 200, 20)



Func uruchomOkienko()
	If FileExists(@ScriptDir & "\wosmbot.conf") Then
		Local $dane
		_FileReadToArray(@ScriptDir & "\wosmbot.conf",$dane)
		GUICtrlSetData($fldSerwer,$dane[1])
		GUICtrlSetData($fldUzytkownikBaza,$dane[2])
		GUICtrlSetData($fldHasloBaza,$dane[3])
		GUICtrlSetData($fldPodmiot,$dane[4])
		GUICtrlSetData($fldUzytkownik,$dane[5])
		GUICtrlSetData($fldHaslo,$dane[6])
	EndIf
	GUISetState(@SW_SHOW, $okienko)
	#EndRegion ### END Koda GUI section ###

	Global $polaczenieOK = False
	Global $wylacz = False
	;btnSprawdzPolaczenieClick()
	;btnSprawdzPolaczenieClick()
	While Not $wylacz
		Sleep(100)
	WEnd
	Local $dane[6] = [GUICtrlRead($fldSerwer), GUICtrlRead($fldUzytkownikBaza), GUICtrlRead($fldHasloBaza), GUICtrlRead($fldPodmiot), GUICtrlRead($fldUzytkownik), GUICtrlRead($fldHaslo)]
	zapiszDane($dane)
	Return $dane
EndFunc   ;==>uruchomOkienko

Func loguj($text)
	_FileWriteLog(@ScriptDir & "\ErrorLog.log", $text)
EndFunc   ;==>loguj

Func btnSprawdzPolaczenieClick()

	If Not $polaczenieOK Then

		Local $ServerAddress = GUICtrlRead($fldSerwer)
		Local $ServerUserName = GUICtrlRead($fldUzytkownikBaza)
		Local $ServerPassword = GUICtrlRead($fldHasloBaza)
		Local $DatabaseName = GUICtrlRead($fldPodmiot)
		If $ServerAddress = "" Or $ServerUserName = "" Or $DatabaseName = "" Then
			MsgBox(0, "Error", "Brak potrzebnych danych!")
			Return
		EndIf

		GUICtrlSetColor($labStatus, $COLOR_BLACK)
		GUICtrlSetData($labStatus, "Sprawdzam połączenie...")
		_SQL_RegisterErrorHandler() ;register the error handler to prevent hard crash on COM error
		$OADODB = _SQL_Startup()
		If $OADODB = $SQL_ERROR Then
			loguj(_SQL_GetErrMsg())
			GUICtrlSetData($labStatus, "Brak połączenia")
			GUICtrlSetColor($labStatus, $COLOR_RED)
			Return
		EndIf
		If _sql_Connect(-1, $ServerAddress, $DatabaseName, $ServerUserName, $ServerPassword) = $SQL_ERROR Then
			loguj(_SQL_GetErrMsg())
			GUICtrlSetData($labStatus, "Brak połączenia")
			GUICtrlSetColor($labStatus, $COLOR_RED)
			Return
		EndIf
		Local $fullSQL
		If _Sql_GetTableAsString(-1, "SELECT [mag_Symbol],[mag_Nazwa] FROM [" & $DatabaseName & "].[dbo].[sl_Magazyn]", $fullSQL) = $SQL_OK Then
			GUICtrlSetData($labStatus, "Nawiązano połączenie")
			GUICtrlSetColor($labStatus, $COLOR_GREEN)
			GUICtrlSetData($btnSprawdzPolaczenie, "Zapisz")
			$polaczenieOK = True
		Else
			loguj(_SQL_GetErrMsg())
			GUICtrlSetData($labStatus, "Brak połączenia")
			GUICtrlSetColor($labStatus, $COLOR_RED)
		EndIf
	Else

		$wylacz = True
		Return
	EndIf
	ControlGetFocus($btnSprawdzPolaczenie)

EndFunc   ;==>btnSprawdzPolaczenieClick

Func fldHasloChange()

EndFunc   ;==>fldHasloChange
Func fldHasloBazaChange()

EndFunc   ;==>fldHasloBazaChange
Func fldPodmiotChange()

EndFunc   ;==>fldPodmiotChange
Func fldSerwerChange()

EndFunc   ;==>fldSerwerChange
Func fldUzytkownikChange()

EndFunc   ;==>fldUzytkownikChange
Func fldUzytkownikBazaChange()

EndFunc   ;==>fldUzytkownikBazaChange

Func zapiszDane(ByRef $dane)
	_FileWriteFromArray(@ScriptDir & "\wosmbot.conf", $dane,Default,Default,@CRLF)
	GUIDelete($okienko)
EndFunc   ;==>zapiszDane
Func oknoClose()
	GUIDelete($okienko)
	Exit
EndFunc   ;==>oknoClose

