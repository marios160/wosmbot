#include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiComboBoxEx.au3>
#include <GuiComboBox.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <_sql.au3>
#include <daneWejsciowe.au3>
#include <ComboConstants.au3>
#include <listaDokumentow.au3>
#include <BotOSM.au3>
;#include <BotWSM.au3>
Opt("GUIOnEventMode", 1)

$subiektxml = '<?xml version="1.0" encoding="windows-1250"?>' & @CR & @LF & _
'<cfg>' & @CR & @LF & _
'	<startup>'



	$hF = FileOpen(EnvGet("ProgramData") & "\InsERT\InsERT GT\Subiekt.xml")

	For $i = 1 to _FileCountLines($hF)
		$line = FileReadLine($hF, $i)
		If StringInStr($line,"<version>") > 0 Then
			$wersja = StringMid($line,StringInStr($line,"<version>")+9,StringInStr($line,"</version>")-StringInStr($line,"<version>") - 9)
			ExitLoop
		EndIf
	Next
	FileClose($hF)

$subiektxml2 = '	</startup>' & @CR & @LF & _
	'	<menu_view>' & @CR & @LF & _
		'		<menu_popup image="2" label="&amp;Sprzedaż">' & @CR & @LF & _
			'			<menu_item style="0" progid="Subiekt.WidokPA.1" accel_key="2" accel_modifier="1" params="" visible="1" label="Sprzedaż &amp;detaliczna"/>' & @CR & @LF & _
			'			<menu_item style="0" progid="Subiekt.WidokZW.1" params="" visible="1" label="&amp;Zwroty detaliczne"/>' & @CR & @LF & _
		'		</menu_popup>' & @CR & @LF & _
		'		<menu_popup image="4" label="&amp;Magazyn">' & @CR & @LF & _
			'			<menu_item style="0" progid="Subiekt.WidokMagW.1" accel_key="4" accel_modifier="1" params="" visible="1" label="&amp;Wydania magazynowe"/>' & @CR & @LF & _
			'			<menu_item style="0" progid="Subiekt.WidokMagP.1" accel_key="5" accel_modifier="1" params="" visible="1" label="&amp;Przyjęcia magazynowe"/>' & @CR & @LF & _
		'		</menu_popup>' & @CR & @LF & _
	'	</menu_view>' & @CR & @LF & _
	'	<version>' & $wersja & '</version>' & @CR & @LF & _
	'</cfg>'



;#############################################################################################################################
;#############################################################################################################################
;#############################################################################################################################
;#############################################################################################################################

#Region ### START Koda GUI section ### Form=C:\Users\Mateusz Błaszczak\Desktop\Bramka\()AUTOIT\GUI\Form1_1.kxf
$Form1 = GUICreate("Bot do wywoływania/odkładania skutków magazynowych", 662, 687, -1036, 116)
GUISetOnEvent($GUI_EVENT_CLOSE, "Form1Close")
GUISetOnEvent($GUI_EVENT_MINIMIZE, "Form1Minimize")
GUISetOnEvent($GUI_EVENT_MAXIMIZE, "Form1Maximize")
GUISetOnEvent($GUI_EVENT_RESTORE, "Form1Restore")
$Group1 = GUICtrlCreateGroup("Czynność", 8, 0, 225, 73)
$radioOdkladanie = GUICtrlCreateRadio("Odkładanie skutków magazynowych", 18, 22, 193, 17)
GUICtrlSetState($radioOdkladanie, $GUI_CHECKED)
GUICtrlSetOnEvent($radioOdkladanie, "radioOdkladanieClick")
$radioWywolywanie = GUICtrlCreateRadio("Wywoływanie skutków magazynowych", 18, 46, 209, 17)
GUICtrlSetOnEvent($radioWywolywanie, "radioWywolywanieClick")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group2 = GUICtrlCreateGroup("Data", 8, 80, 225, 97)
$dataOd = GUICtrlCreateDate("2018/05/25 10:10:35", 48, 104, 162, 21,$DTS_SHORTDATEFORMAT)
GUICtrlSetOnEvent($dataOd, "dataOdChange")
$dataDo = GUICtrlCreateDate("2018/05/01 10:10:39", 48, 136, 162, 21,$DTS_SHORTDATEFORMAT)
GUICtrlSetOnEvent($dataDo, "dataDoChange")
$Label1 = GUICtrlCreateLabel("Od:", 24, 104, 21, 17)
GUICtrlSetOnEvent($Label1, "Label1Click")
$Label2 = GUICtrlCreateLabel("Do:", 24, 136, 21, 17)
GUICtrlSetOnEvent($Label2, "Label2Click")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group3 = GUICtrlCreateGroup("Dokumenty", 8, 248, 225, 233)
$chckSprzedazDetaliczna = GUICtrlCreateCheckbox("Sprzedaż detaliczna", 24, 272, 121, 17)
GUICtrlSetOnEvent($chckSprzedazDetaliczna, "chckSprzedazDetalicznaClick")
$chckWydMag = GUICtrlCreateCheckbox("Wydania magazynowe", 24, 296, 129, 17)
GUICtrlSetOnEvent($chckWydMag, "chckWydMagClick")
$chckWydMagRozchodWewnetrzny = GUICtrlCreateCheckbox("Rozchód wewnętrzny", 40, 320, 129, 17)
GUICtrlSetOnEvent($chckWydMagRozchodWewnetrzny, "chckWydMagRozchodWewnetrznyClick")
$chckWydMagPrzesuniecieMiedzymagazynowe = GUICtrlCreateCheckbox("Przesunięcie międzymagazynowe", 40, 344, 177, 17)
GUICtrlSetOnEvent($chckWydMagPrzesuniecieMiedzymagazynowe, "chckWydMagPrzesuniecieMiedzymagazynoweClick")
$chckPrzyjMag = GUICtrlCreateCheckbox("Przyjęcia magazynowe", 24, 368, 129, 17)
GUICtrlSetOnEvent($chckPrzyjMag, "chckPrzyjMagClick")
$chckPrzyjMagPrzychodWewnetrzny = GUICtrlCreateCheckbox("Przychód wewnętrzny", 40, 392, 153, 17)
GUICtrlSetOnEvent($chckPrzyjMagPrzychodWewnetrzny, "chckPrzyjMagPrzychodWewnetrznyClick")
$chckZwrotyDetaliczne = GUICtrlCreateCheckbox("Zwroty detaliczne", 24, 416, 105, 17)
GUICtrlSetOnEvent($chckZwrotyDetaliczne, "chckZwrotyDetaliczneClick")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$lstKolejnosc = GUICtrlCreateList("", 310, 32, 340, 448, BitOR($LBS_NOTIFY,$WS_HSCROLL,$WS_VSCROLL,$WS_BORDER))
GUICtrlSetData($lstKolejnosc, "")
GUICtrlSetOnEvent($lstKolejnosc, "lstKolejnoscClick")
$Label3 = GUICtrlCreateLabel("Kolejność", 310, 8, 50, 17)
GUICtrlSetOnEvent($Label3, "Label3Click")
$btnWGore = GUICtrlCreateButton("W górę", 240, 32, 67, 25)
GUICtrlSetFont($btnWGore, 8, 800, 0, "MS Sans Serif")
GUICtrlSetOnEvent($btnWGore, "btnWGoreClick")
$btnWDol = GUICtrlCreateButton("W dół", 240, 56, 67, 25)
GUICtrlSetFont($btnWDol, 8, 800, 0, "MS Sans Serif")
GUICtrlSetOnEvent($btnWDol, "btnWDolClick")
$btnWyczysc = GUICtrlCreateButton("Wyczyść", 240, 90, 67, 25)
GUICtrlSetFont($btnWyczysc, 8, 800, 0, "MS Sans Serif")
GUICtrlSetOnEvent($btnWyczysc, "btnWyczyscClick")
$Group4 = GUICtrlCreateGroup("Magazyn", 8, 184, 225, 57)
$cmbMagazyn = GUICtrlCreateCombo("", 15, 208, 201, 25)
GUICtrlSetOnEvent($cmbMagazyn, "cmbMagazynChange")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$btnStart = GUICtrlCreateButton("Start", 8, 488, 75, 25)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
GUICtrlSetOnEvent(-1, "btnStartClick")
$btnPauza = GUICtrlCreateButton("Pauza", 88, 488, 67, 25)
GUICtrlSetOnEvent(-1, "btnPauzaClick")
$btnStop = GUICtrlCreateButton("Stop", 160, 488, 67, 25)
GUICtrlSetOnEvent(-1, "btnStopClick")
$edtLogi = GUICtrlCreateEdit("", 8, 520, 641, 153)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetOnEvent(-1, "edtLogiChange")
$chckZaznaczWszystkie = GUICtrlCreateCheckbox("Zaznacz wszystkie", 24, 448, 137, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
GUICtrlSetOnEvent(-1, "chckZaznaczWszystkieClick")
GUICtrlCreateGroup("", -99, -99, 1, 1)
#EndRegion ### END Koda GUI section ###


Global $checkBoxes[8] = [$chckSprzedazDetaliczna,$chckWydMag,$chckWydMagRozchodWewnetrzny, _
						$chckWydMagPrzesuniecieMiedzymagazynowe, $chckPrzyjMag, _
						$chckPrzyjMagPrzychodWewnetrzny,$chckZwrotyDetaliczne,$chckZaznaczWszystkie]
Global $_ListaDokumentow = ""
Global $_Magazyny[0][0]
Global $aktMag = 0
Global $aktMagTxt = ""
Global $_ListaDokumentowTab[0]
Global $_ListaDokumentowBot[0]
Global $_ListaMagazynowBot[0]
;		    0 	  czy     1
Global $_OdkladanieCzyWywolywanie = 0
Global $_DataOd = ""
Global $_DataDo = ""
Global $_Podmiot = ""
Global $_SciezkaUruchomieniaSubiekta = ""



$dane = uruchomOkienko()

GUISetState(@SW_SHOW,$Form1)
Local $ServerAddress = $dane[0]
Local $ServerUserName = $dane[1]
Local $ServerPassword = $dane[2]
Local $DatabaseName = $dane[3]
Local $uzytkownik = $dane[4]
Local $haslo = $dane[5]
Local $_Podmiot = $dane[3]
Local $_nazwaOknaSubiekta = $_Podmiot & " na serwerze " & $ServerAddress & " - Subiekt GT"


	$hFile = FileOpen("SubiektBot.xml", 2 + 512)
	FileWrite($hFile, $subiektxml & @CR & @LF)
	FileWrite($hFile, '		<sql_server>' & $ServerAddress & '</sql_server>' & @CR & @LF)
	FileWrite($hFile, '		<auth_mode>MIXED</auth_mode>' & @CR & @LF)
	FileWrite($hFile, '		<sql_login encrypted="0">' & $ServerUserName & '\' & $ServerPassword  & '</sql_login>' & @CR & @LF)
	FileWrite($hFile, '		<database>' & $DatabaseName & '</database>' & @CR & @LF)
	FileWrite($hFile, '		<login encrypted="0">' & $uzytkownik & '\' & $haslo & '</login>' & @CR & @LF)
	FileWrite($hFile, $subiektxml2 & @CR & @LF)
	FileClose($hFile)

	If FileExists(EnvGet("ProgramFiles(x86)") & "\InsERT\InsERT GT\Subiekt.exe") Then
		$pathSubiekt = EnvGet("ProgramFiles(x86)") & "\InsERT\InsERT GT"
	ElseIf FileExists(EnvGet("ProgramFilesW6432") & "\InsERT\InsERT GT\Subiekt.exe") Then
		$pathSubiekt = EnvGet("ProgramFilesW6432") & "\InsERT\InsERT GT"
	Else
		Do
			$odp = MsgBox(1,"Nie znaleziono programu 'Subiekt GT'", "Wybierz folder, w którym zainstalowano program 'Subiekt.exe'")
			If $odp = 2 Then
				Exit
			EndIf
			$pathSubiekt = FileSelectFolder("Wybierz folder gdzie zainstalowano 'Subiekt.exe'",EnvGet("ProgramFiles(x86)"))
		Until FileExists($pathSubiekt & "\subiekt.exe") = 1
	EndIf
	$_SciezkaUruchomieniaSubiekta = $pathSubiekt & '\Subiekt.exe "'	& @ScriptDir & '\SubiektBot.xml"'

		If $ServerAddress = "" or $ServerUserName = "" Or $DatabaseName = "" Then
			MsgBox(0,"Error","Brak potrzebnych danych!")
			Exit
		EndIf

		_SQL_RegisterErrorHandler() ;register the error handler to prevent hard crash on COM error
		$OADODB = _SQL_Startup()
		If $OADODB = $SQL_ERROR Then
			loguj(_SQL_GetErrMsg())
			Exit
		EndIf
		If _sql_Connect(-1, $ServerAddress, $DatabaseName, $ServerUserName, $ServerPassword) = $SQL_ERROR Then
			loguj(_SQL_GetErrMsg())
			Exit
		EndIf
		Local $fullSQL
		If _Sql_GetTableAsString(-1, "SELECT [mag_Symbol],[mag_Nazwa] FROM [Arctus_Test_BotSM].[dbo].[sl_Magazyn]", $fullSQL," - ",0) = $SQL_OK Then
			$polaczenieOK = True
			$magTmp = StringSplit($fullSQL,@LF)
			ReDim $_Magazyny[$magTmp[0]][8]
			ReDim $_ListaMagazynowBot[$magTmp[0]]
			For $i = 0 To $magTmp[0] - 1
				For $j = 0 To 7
					$_Magazyny[$i][$j] = 4
				Next
				If $magTmp[$i + 1] <> "" Then $_ListaMagazynowBot[$i] = $magTmp[$i + 1]
			Next
			$fullSQL = StringReplace($fullSQL,@LF,"|")
			GUICtrlSetData($cmbMagazyn,$fullSQL)
			$aktMagTxt = StringLeft($fullSQL,3)
			_GUICtrlComboBox_SetCurSel($cmbMagazyn,0)
		Else
			loguj(_SQL_GetErrMsg())
			Exit
		EndIf
		_SQL_Close($OADODB)

While 1
	Sleep(100)
WEnd


;#############################################################################################################################
;#############################################################################################################################
;#############################################################################################################################
;#############################################################################################################################

;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne
Func btnStartClick()


	$_DataOd = GUICtrlRead($dataOd)
	$_DataDo = GUICtrlRead($dataDo)
	ReDim $_ListaDokumentowBot[UBound($_ListaDokumentowTab)]
	For $i = 0 to UBound($_ListaDokumentowTab) - 1
		$el = StringSplit($_ListaDokumentowTab[$i],"[",2)

		Switch(StringStripWS($el[0],  BitOR($STR_STRIPLEADING,$STR_STRIPTRAILING)))
			Case "Sprzedaż detaliczna"
				$_ListaDokumentowBot[$i] = "PA"
			Case "Przchód wewnętrzny (Przyjęcia magazynowe)"
				$_ListaDokumentowBot[$i] = "PW"
			Case "Przesunięcie międzymagazynowe (Wydania magazynowe)"
				$_ListaDokumentowBot[$i] = "MM"
			Case "Rozchód wewnętrzny (Wydania magazynowe)"
				$_ListaDokumentowBot[$i] = "RW"
			Case "Zwroty detaliczne"
				$_ListaDokumentowBot[$i] = "ZW"
			Case Else
				$_ListaDokumentowBot[$i] = $el[0]
		EndSwitch
	Next
	Run($_SciezkaUruchomieniaSubiekta,"")

	While czySubiektUruchomiony() <> 0
		Sleep (1000)
	WEnd
	Sleep(2000)

	Local $daneBot[3] = [$_Podmiot,$_DataOd,$_DataDo]
	If GUICtrlRead($radioOdkladanie) = $GUI_CHECKED Then
		startOSMBot($daneBot,$_ListaDokumentowBot, $_ListaMagazynowBot)
	Else
		;startWSMBot($daneBot,$_ListaDokumentowBot, $_ListaMagazynowBot)
	EndIf

EndFunc

;-----------------------------------------------------------------------------------------------------------------------------

Func cmbMagazynChange()
	$aktMag = _GUICtrlComboBox_GetCurSel($cmbMagazyn)
	$aktMagTxt = StringLeft(GUICtrlRead($cmbMagazyn),3)
	For $i = 0 To 7
		GUICtrlSetState($checkBoxes[$i], $_Magazyny[$aktMag][$i])
	Next
EndFunc

;-----------------------------------------------------------------------------------------------------------------------------
Func btnWyczyscClick()
	ReDim $_ListaDokumentowTab[0]
	GUICtrlSetData($lstKolejnosc, $_ListaDokumentowTab)
	For $j = 0 To UBound($_Magazyny) - 1
		For $i = 0 To 7
			$_Magazyny[$j][$i] = 4
			GUICtrlSetState($checkBoxes[$i], $_Magazyny[$j][$i])
		Next
	Next

EndFunc

;-----------------------------------------------------------------------------------------------------------------------------


Func btnWGoreClick()
	$item = _GUICtrlListBox_GetCurSel($lstKolejnosc)
	If $item = -1 Then Return
	$selStr = $_ListaDokumentowTab[$item]
	$_ListaDokumentowTab = moveUp($item,$_ListaDokumentowTab)
	GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
	GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
	$item = _ArraySearch($_ListaDokumentowTab,$selStr,0,$item - 1,0,0,0)
	ConsoleWrite($item & @LF)
	_GUICtrlListBox_SelectString($lstKolejnosc, $selStr,$item - 1)
;~ 	ConsoleWrite($item)
;~ 	$items = StringSplit($_ListaDokumentow, "|", $STR_NOCOUNT)
;~ 	If $item > 0 Then
;~ 		$_ListaDokumentow = ""
;~ 		$tmp = $items[$item]
;~ 		$items[$item] = $items[$item - 1]
;~ 		$items[$item - 1] = $tmp
;~ 		For $i = 0 To UBound($items) - 2
;~ 			$_ListaDokumentow = $_ListaDokumentow & $items[$i] & "|"
;~ 		Next
;~ 		GUICtrlSetData($lstKolejnosc, "")
;~ 		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
;~ 		_GUICtrlListBox_SelectString($lstKolejnosc, $tmp, $item - 1)
;~ 	EndIf
EndFunc   ;==>btnWGoreClick

;-----------------------------------------------------------------------------------------------------------------------------

Func btnWDolClick()
	$item = _GUICtrlListBox_GetCurSel($lstKolejnosc)
	If $item = -1 Then Return
	ConsoleWrite($item)
	$selStr = $_ListaDokumentowTab[$item]
	$_ListaDokumentowTab = moveDown($item,$_ListaDokumentowTab)
	GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
	GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
	$item = _ArraySearch($_ListaDokumentowTab,$selStr,$item)
	ConsoleWrite($item & @LF)
	_GUICtrlListBox_SelectString($lstKolejnosc, $selStr,$item -1 )

;~ 	$items = StringSplit($_ListaDokumentow, "|", $STR_NOCOUNT)
;~ 	If $item < UBound($items) - 2 Then
;~ 		$_ListaDokumentow = ""
;~ 		$tmp = $items[$item]
;~ 		$items[$item] = $items[$item + 1]
;~ 		$items[$item + 1] = $tmp
;~ 		For $i = 0 To UBound($items) - 2
;~ 			$_ListaDokumentow = $_ListaDokumentow & $items[$i] & "|"
;~ 		Next
;~ 		GUICtrlSetData($lstKolejnosc, "")
;~ 		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
;~ 		_GUICtrlListBox_SelectString($lstKolejnosc, $tmp, $item - 1)
;~ 	EndIf
EndFunc   ;==>btnWDolClick

;-----------------------------------------------------------------------------------------------------------------------------

Func chckPrzyjMagClick()
	If GUICtrlRead($chckPrzyjMag) = $GUI_CHECKED Then
		GUICtrlSetState($chckPrzyjMagPrzychodWewnetrzny, $GUI_CHECKED)
		chckPrzyjMagPrzychodWewnetrznyClick()
	Else
		GUICtrlSetState($chckPrzyjMagPrzychodWewnetrzny, $GUI_UNCHECKED)
		chckPrzyjMagPrzychodWewnetrznyClick()
	EndIf
EndFunc   ;==>chckPrzyjMagClick

;-----------------------------------------------------------------------------------------------------------------------------

Func chckPrzyjMagPrzychodWewnetrznyClick()
	If GUICtrlRead($chckPrzyjMagPrzychodWewnetrzny) = $GUI_CHECKED Then
		;$_ListaDokumentow = $_ListaDokumentow & "Przchód wewnętrzny (Przyjęcia magazynowe)|"
		$_ListaDokumentowTab = addEl("     Przchód wewnętrzny (Przyjęcia magazynowe) [" & $aktMagTxt & "]",$_ListaDokumentowTab)
		GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
		GUICtrlSetState($chckPrzyjMag, $GUI_CHECKED)
		$_Magazyny[$aktMag][4] = 1
		$_Magazyny[$aktMag][5] = 1
	Else
		;$_ListaDokumentow = StringReplace($_ListaDokumentow, "Przchód wewnętrzny (Przyjęcia magazynowe)|", "")
		$_ListaDokumentowTab = removeEl("     Przchód wewnętrzny (Przyjęcia magazynowe) [" & $aktMagTxt & "]",$_ListaDokumentowTab)
		GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
		GUICtrlSetState($chckPrzyjMag, $GUI_UNCHECKED)
		$_Magazyny[$aktMag][4] = 4
		$_Magazyny[$aktMag][5] = 4
	EndIf
EndFunc   ;==>chckPrzyjMagPrzychodWewnetrznyClick

;-----------------------------------------------------------------------------------------------------------------------------

Func chckSprzedazDetalicznaClick()

	If GUICtrlRead($chckSprzedazDetaliczna) = $GUI_CHECKED Then
		;$_ListaDokumentow = $_ListaDokumentow & "Sprzedaż detaliczna|"
		$_ListaDokumentowTab = addEl("     Sprzedaż detaliczna [" & $aktMagTxt & "]",$_ListaDokumentowTab)
		GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
		$_Magazyny[$aktMag][0] = 1
	Else
		;$_ListaDokumentow = StringReplace($_ListaDokumentow, "Sprzedaż detaliczna|", "")
		$_ListaDokumentowTab = removeEl("     Sprzedaż detaliczna [" & $aktMagTxt & "]",$_ListaDokumentowTab)
		GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
		$_Magazyny[$aktMag][0] = 4
	EndIf
EndFunc   ;==>chckSprzedazDetalicznaClick

;-----------------------------------------------------------------------------------------------------------------------------

Func chckWydMagClick()
	If GUICtrlRead($chckWydMag) = $GUI_CHECKED Then
		GUICtrlSetState($chckWydMagRozchodWewnetrzny, $GUI_CHECKED)
		GUICtrlSetState($chckWydMagPrzesuniecieMiedzymagazynowe, $GUI_CHECKED)
		chckWydMagPrzesuniecieMiedzymagazynoweClick()
		chckWydMagRozchodWewnetrznyClick()
	Else
		GUICtrlSetState($chckWydMagRozchodWewnetrzny, $GUI_UNCHECKED)
		GUICtrlSetState($chckWydMagPrzesuniecieMiedzymagazynowe, $GUI_UNCHECKED)
		chckWydMagPrzesuniecieMiedzymagazynoweClick()
		chckWydMagRozchodWewnetrznyClick()
	EndIf
EndFunc   ;==>chckWydMagClick

;-----------------------------------------------------------------------------------------------------------------------------

Func chckWydMagPrzesuniecieMiedzymagazynoweClick()
	If GUICtrlRead($chckWydMagPrzesuniecieMiedzymagazynowe) = $GUI_CHECKED Then
		;$_ListaDokumentow = $_ListaDokumentow & "Przesunięcie międzymagazynowe (Wydania magazynowe)|"
		$_ListaDokumentowTab = addEl("     Przesunięcie międzymagazynowe (Wydania magazynowe) [" & $aktMagTxt & "]",$_ListaDokumentowTab)
		GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
		GUICtrlSetState($chckWydMag, $GUI_CHECKED)
		$_Magazyny[$aktMag][1] = 1
		$_Magazyny[$aktMag][3] = 1
	Else
		;$_ListaDokumentow = StringReplace($_ListaDokumentow, "Przesunięcie międzymagazynowe (Wydania magazynowe)|", "")
		$_ListaDokumentowTab = removeEl("     Przesunięcie międzymagazynowe (Wydania magazynowe) [" & $aktMagTxt & "]",$_ListaDokumentowTab)
		GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
		$_Magazyny[$aktMag][3] = 4
		If GUICtrlRead($chckWydMagRozchodWewnetrzny) = $GUI_UNCHECKED Then
			GUICtrlSetState($chckWydMag, $GUI_UNCHECKED)
			$_Magazyny[$aktMag][1] = 4
		EndIf
	EndIf
EndFunc   ;==>chckWydMagPrzesuniecieMiedzymagazynoweClick

;-----------------------------------------------------------------------------------------------------------------------------

Func chckWydMagRozchodWewnetrznyClick()
	If GUICtrlRead($chckWydMagRozchodWewnetrzny) = $GUI_CHECKED Then
		;$_ListaDokumentow = $_ListaDokumentow & "Rozchód wewnętrzny (Wydania magazynowe)|"
		$_ListaDokumentowTab = addEl("     Rozchód wewnętrzny (Wydania magazynowe) [" & $aktMagTxt & "]",$_ListaDokumentowTab)
		GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
		GUICtrlSetState($chckWydMag, $GUI_CHECKED)
		$_Magazyny[$aktMag][1] = 1
		$_Magazyny[$aktMag][2] = 1
	Else
		;$_ListaDokumentow = StringReplace($_ListaDokumentow, "Rozchód wewnętrzny (Wydania magazynowe)|", "")
		$_ListaDokumentowTab = removeEl("     Rozchód wewnętrzny (Wydania magazynowe) [" & $aktMagTxt & "]",$_ListaDokumentowTab)
		GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
		$_Magazyny[$aktMag][2] = 4
		If GUICtrlRead($chckWydMagPrzesuniecieMiedzymagazynowe) = $GUI_UNCHECKED Then
			GUICtrlSetState($chckWydMag, $GUI_UNCHECKED)
			$_Magazyny[$aktMag][1] = 4
		EndIf

	EndIf
EndFunc   ;==>chckWydMagRozchodWewnetrznyClick

;-----------------------------------------------------------------------------------------------------------------------------

Func chckZwrotyDetaliczneClick()
	If GUICtrlRead($chckZwrotyDetaliczne) = $GUI_CHECKED Then
		;$_ListaDokumentow = $_ListaDokumentow & "Zwroty detaliczne|"
		$_ListaDokumentowTab = addEl("     Zwroty detaliczne [" & $aktMagTxt & "]",$_ListaDokumentowTab)
		GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
		$_Magazyny[$aktMag][6] = 1
	Else
		;$_ListaDokumentow = StringReplace($_ListaDokumentow, "Zwroty detaliczne|", "")
		$_ListaDokumentowTab = removeEl("     Zwroty detaliczne [" & $aktMagTxt & "]",$_ListaDokumentowTab)
		GUICtrlSetData($lstKolejnosc, "")
		;GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetData($lstKolejnosc, toString($_ListaDokumentowTab))
		$_Magazyny[$aktMag][6] = 4
	EndIf

EndFunc   ;==>chckZwrotyDetaliczneClick

;-----------------------------------------------------------------------------------------------------------------------------

Func chckZaznaczWszystkieClick()
	$status = 4
	If GUICtrlRead($chckZaznaczWszystkie) = $GUI_CHECKED Then
		$status = 1
	Else
		$status = 4
	EndIf

		For $i = 0 To 7
			$_Magazyny[$aktMag][$i] = $status
			GUICtrlSetState($checkBoxes[$i], $_Magazyny[$aktMag][$i])
		Next
		chckSprzedazDetalicznaClick()
		chckWydMagClick()
		chckPrzyjMagClick()
		chckZwrotyDetaliczneClick()
EndFunc

;-----------------------------------------------------------------------------------------------------------------------------

Func czySubiektUruchomiony()

	$hWnd = zlapSubiekta()

	$uchwyt = szukajUchwytu($hWnd, "Wybór serwera, użytkownika i hasła", "#32770")
	If $uchwyt <> 0 Then
		Return 1 ; serwer
	EndIf

	$uchwyt = szukajUchwytu($hWnd, "Wybór podmiotu", "#32770")
	If $uchwyt <> 0 Then
		Return 2 ; podmiot
	EndIf

	$uchwyt = szukajUchwytu($hWnd, "Wybierz użytkownika, z którym chcesz rozpocząć pracę", "Static")
	If $uchwyt <> 0 Then
		Return 3 ; uzytkownik
	EndIf

	$uchwyt = szukajUchwytu($hWnd, "Trwa logowanie, proszę czekać", "#32770")
	If $uchwyt <> 0 Then
		Return 4 ; uzytkownik
	EndIf

	$uchwyt = szukajUchwytu($hWnd, "", "ins_svcwindow")
	If $uchwyt <> 0 Then
		Return 0 ;subiekt
	Else
		Return -1 ;brak
	EndIf

EndFunc   ;==>wykryjEtapLogowania


;-----------------------------------------------------------------------------------------------------------

Func zlapSubiekta()
	$hWnd = WinGetHandle($_nazwaOknaSubiekta)
	$uchwyt = szukajUchwytu($hWnd, "Magazyn:", "ins_combohyperlink")
	If $uchwyt <> 0 Then
		Return $hWnd
	EndIf

	$hWnd = WinGetHandle("Subiekt GT")
	$uchwyt = szukajUchwytu($hWnd, "Magazyn:", "ins_combohyperlink")
	If $uchwyt <> 0 Then
		Return $hWnd
	EndIf
	Return 0
EndFunc   ;==>uchwycSubiekta






Func btnPauzaClick()

EndFunc




Func btnStopClick()

EndFunc
Func edtLogiChange()
EndFunc

Func dataDoChange()

EndFunc   ;==>dataDoChange

Func dataOdChange()

EndFunc   ;==>dataOdChange

Func Form1Close()
	GUIDelete()
	Exit
EndFunc   ;==>Form1Close

Func Form1Maximize()

EndFunc   ;==>Form1Maximize
Func Form1Minimize()

EndFunc   ;==>Form1Minimize
Func Form1Restore()

EndFunc   ;==>Form1Restore
Func Label1Click()

EndFunc   ;==>Label1Click
Func Label2Click()

EndFunc   ;==>Label2Click
Func Label3Click()

EndFunc   ;==>Label3Click
Func lstKolejnoscClick()

EndFunc   ;==>lstKolejnoscClick
Func radioOdkladanieClick()

EndFunc   ;==>radioOdkladanieClick
Func radioWywolywanieClick()

EndFunc   ;==>radioWywolywanieClick


