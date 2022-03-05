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
;#include <APISubiektWin.au3>

Func startOSMBot($dane, $dokumenty, $_ListaMagazynow)
	HotKeySet("{F6}", "zamknij")
	HotKeySet("{PAUSE}", "pauza")
	HotKeySet("{F4}", "pauzaSzybka")
;~ 	$odp = $IDNO
;~ 	While $odp == $IDNO
;~ 		$odp = MsgBox($MB_YESNOCANCEL, "Start", "Proszę włączyć na liscie zwrotow kolumne Dokument korygowany oraz na każdej liście ustawić sortowanie od najmłodszych dokumentów do najstarszych!" & @LF & "Czy chcesz rozpocząć pracę?")
;~ 	WEnd
;~ 	If $odp == $IDCANCEL Then
;~ 		Exit
;~ 	EndIf

;~ 	$odp = MsgBox($MB_YESNO, "Czyszczenie loga", "Czy chcesz wyczyscic 'logOSM.log'?")
;~ 	If $odp == $IDYES Then
	$hF = FileOpen("logOSM.log", 2)
	FileClose($hF)
;~ 	EndIf


	;;-----------Inicjalizacja zmiennych
	Global $_PAUSE = False
	Global $_DATA = $dane[1]
	Global $_ENDDATA = $dane[2]
	Global $_PODMIOT = $dane[0]

	;;-----------Petla programowa,,,,,,,,,,
	logs("START BotWSM")
	While $_DATA <> $_ENDDATA
		logs("**************************************************************************************")
		logs("*************************************" & $_DATA & " **************************************")
		logs("**************************************************************************************")
		For $i = 0 To UBound($dokumenty) - 1
			If StringLen($dokumenty[$i]) = 3 Then
				wybierzMagazyn($_ListaMagazynow.Item($dokumenty[$i]))
			Else
				wykonujListe($dokumenty[$i])
			EndIf
			Sleep(1000)
		Next
		$_DATA = _DateAdd('d', -1, $_DATA)
	WEnd
	logs("KONIEC BotOSM")
	wyslijSms("KONIEC BotOSM")
	Exit
EndFunc   ;==>startOSMBot

;#####################################################################################################################
;#####################################################################################################################
;################################													##################################
;################################						FUNKCJE						##################################
;################################													##################################
;#####################################################################################################################
;#####################################################################################################################

;---------------------------------------------------------------------------------------------------------------------
;Funkcja wykonuje wszystkie instrukcje potrzebne do wywołania skutkow magazynowych wszystkich pozycji na konkretnej liscie
;Przyjmuje nazwę listy na ktorej działamy
Func wykonujListe($lista)
	$lista = wybierzListe($lista)
	Sleep(500)
	ustawOkres($_DATA, $lista)
	Sleep(500)
	logs("------------------------------- START " & $lista & " ----------------------------------------")
	If $_PAUSE = True Then
		pauzaLoop()
	EndIf
	If Not czyPusta($lista) Then
		Do
			sendGXWNDEndHard($lista)
			$endpos = kopiujWiersz($lista)
			$curpos = ""
			sendGXWNDHomeHard($lista)
			While $endpos <> $curpos
				$czyDaSieOdlozyc = False
				$curpos = kopiujWiersz($lista)
				If StringInStr($curpos, $STR_WYWOLUJE_SKUTEK) <> 0 Then
					While Not $czyDaSieOdlozyc
						$curpos2 = kopiujWiersz($lista)
						If StringInStr($curpos2, $STR_WYWOLUJE_SKUTEK) <> 0 And Not $czyDaSieOdlozyc Then
							$czyDaSieOdlozyc = odlozSkutekMagazynowy($lista)
						Else
							$czyDaSieOdlozyc = True
						EndIf
					WEnd
				EndIf
				Sleep(10)
				If $_PAUSE = True Then
					pauzaLoop()
					ExitLoop
				EndIf
				sendGXWNDDown($lista)
			WEnd
			If StringInStr(kopiujListe($lista), $STR_WYWOLUJE_SKUTEK) <> 0 Then
				StringReplace(kopiujListe($lista), $STR_WYWOLUJE_SKUTEK, "xxx")
				logs("OMINIETO PRODUKT! " & "; " & " znaleziono: " & @extended)
			EndIf
		Until StringInStr(kopiujListe($lista), $STR_WYWOLUJE_SKUTEK) = 0
	EndIf
	logs("------------------------------- KONIEC " & $lista & " ----------------------------------------")
EndFunc   ;==>wykonujListe



;~ ;---------------------------------------------------------------------------------------------------------------------
;~ ;Funkcja wywołuje skutek magazynowy
;~ ;Przyjmuje nazwę listy na ktorej działamy
;~ ;
;~ ;PA - paragony
;~ ;FA - faktury i rachunki zakupu
;~ ;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;~ ;RW - rozchod wewnetrzny - wydania magazynowe
;~ ;PW - przychod wewnetrzny - przyjecia magazynowe
;~ ;ZW - zwroty detaliczne
;~ Func odlozSkutek($lista)
;~ 	$okno = ""
;~ 	$paragon = ""
;~ 	;KOREKTY ZAKUPU
;~ 	$tmp = StringSplit(ClipGet(), "	", 1)
;~ 	Switch $lista
;~ 		Case "PA"
;~ 			Send("{LALT}OM")
;~ 			$okno = "Paragon"
;~ 			$paragon = $tmp[6]
;~ 		Case "FA"
;~ 			Send("{LALT}OM")
;~ 			$okno = "Faktura VAT zakupu"
;~ 			$paragon = $tmp[5]
;~ 		Case "KFZ"
;~ 			Send("{LALT}OM")
;~ 			$okno = "Korekta fakturu VAT zakupu"
;~ 			$paragon = $tmp[5]
;~ 		Case "MM"
;~ 			Send("{LALT}OM")
;~ 			$okno = "Przesunięcie międzymagazynowe"
;~ 			$paragon = $tmp[4]
;~ 		Case "RW"
;~ 			Send("{LALT}OM")
;~ 			$okno = "Rozchód wewnętrzny"
;~ 			$paragon = $tmp[4]
;~ 		Case "PW"
;~ 			Send("{LALT}OM")
;~ 			$okno = "Przychód wewnętrzny"
;~ 			$paragon = $tmp[4]
;~ 		Case "ZW"
;~ 			Send("{LALT}OM")
;~ 			$okno = "Zwrot ze sprzedaży detalicznej"
;~ 			$paragon = $tmp[4]
;~ 	EndSwitch
;~ 	$var = False
;~ 	;logs($paragon & " ")
;~ 	$licznik = 0
;~ 	While WinGetHandle($okno) = 0 And Not $var And $licznik < 3000 ;ControlFocus($t,"","") = 0 Or WinGetHandle("Subiekt GT") <> 0
;~ 		$var = oknoSubiekt($lista)
;~ 		$licznik = $licznik + 1
;~ 		;ConsoleWrite(WinGetHandle($okno) & " * " & $var & " * " )
;~ 	WEnd
;~ 	$licznik = 0
;~ 	While WinGetHandle($okno) <> 0 And Not $var And $licznik < 3000 ;ControlFocus($t,"","") = 0 Or WinGetHandle("Subiekt GT") <> 0
;~ 		$var = oknoSubiekt($lista)
;~ 		$licznik = $licznik + 1
;~ 		;ConsoleWrite(WinGetHandle($okno) & " - " & $var & " - " )
;~ 	WEnd
;~ 	;logs(@LF)
;~ 	If $var Then
;~ 		logs("Nie odlozono skutku (brak towaru): " & $paragon)
;~ 	Else
;~ 		logs("Odlozono skutek: " & $paragon)
;~ 	EndIf
;~ 	Return $var
;~ EndFunc   ;==>odlozSkutek

;~ ;---------------------------------------------------------------------------------------------------------------------
;~ ;Funkcja obsluguje okienko brakującego towaru. Zapisuje do pliku wszystkie pozycje z listy a nastepnie zamyka okienko
;~ ;przyjmuje parametr okreslajacy czy dany produkt juz został zapisany oraz nazwe przerabianej listy
;~ Func zapiszBrakTowaru($lista)
;~ 	$brak = WinGetHandle("[REGEXPTITLE:(?i)(.* - brak towaru w magazynie)]")
;~ 	$uchwyt = szukajUchwytu($brak, "", "GXWND")
;~ 	$paragon = ""
;~ 	$data = ""
;~ 	ControlFocus($uchwyt, "", "")
;~ 	$tmp = StringSplit(ClipGet(), "	", 1)
;~ 	Switch $lista
;~ 		Case "PA"
;~ 			$data = $tmp[3]
;~ 			$paragon = $tmp[6]
;~ 		Case "FA"
;~ 			$data = $tmp[3]
;~ 			$paragon = $tmp[5]
;~ 		Case "KFZ"
;~ 			$data = $tmp[3]
;~ 			$paragon = $tmp[5]
;~ 		Case "MM"
;~ 			$data = $tmp[3]
;~ 			$paragon = $tmp[4]
;~ 		Case "RW"
;~ 			$data = $tmp[3]
;~ 			$paragon = $tmp[4]
;~ 		Case "PW"
;~ 			$data = $tmp[3]
;~ 			$paragon = $tmp[4]
;~ 		Case "ZW"
;~ 			$data = $tmp[3]
;~ 			$paragon = $tmp[4]
;~ 	EndSwitch
;~ 	$hFile = FileOpen("OSMbraktowaru.txt", 1)
;~ 	FileWrite($hFile, $paragon & ";" & $data & @LF)
;~ 	$iloscLiniiBrakow = $iloscLiniiBrakow + 1
;~ 	kopiujListe()
;~ 	$tmp2 = StringSplit(ClipGet(), @LF, 1)
;~ 	For $i = 2 To UBound($tmp2) - 1
;~ 		FileWrite($hFile, $tmp2[$i])
;~ 		$iloscLiniiBrakow = $iloscLiniiBrakow + 1
;~ 	Next
;~ 	FileClose($hFile)
;~ 	logs("Zapisano brak towaru: " & $paragon)
;~ 	$uchwyt = szukajUchwytu($brak, "Zamknij", "Button")
;~ 	ControlClick($uchwyt, "", "")
;~ EndFunc   ;==>zapiszBrakTowaru
Func logs($text)
	dodajLog($text)
	_FileWriteLog("logOSM.log", "[" & $_DATA & "] " & $text)
	ConsoleWrite("[" & $_DATA & "] " & $text & @LF)
	;Run('SmsEagle.exe "' & "[" & $_DATA & "] " & $text & @LF & '"', "", @SW_HIDE)
EndFunc   ;==>logs

;~ ;---------------------------------------------------------------------------------------------------------------------
;~ ;Funkcja obsluguje okienko potwierdzenia wywolania skutku
;~ Func oknoSubiekt($lista)
;~ 	$var = False
;~ 	$hSub = WinGetHandle("Subiekt GT")
;~ 	If $hSub <> 0 Then
;~ 		$uchwyt = szukajUchwytu($hSub, "&Tak", "Button")
;~ 		If $uchwyt = 0 Then
;~ 			$uchwyt = szukajUchwytu($hSub, "OK", "Button")
;~ 			If $uchwyt = 0 Then
;~ 				$uchwyt = szukajUchwytu($hSub, "Pomoc", "Button")
;~ 				If $uchwyt = 0 Then
;~ 					If _WinAPI_GetClassName($hSub) = "BalloonHelpClass" Then
;~ 						Send("{ESC}")
;~ 						$var = True
;~ 					EndIf
;~ 				Else
;~ 					$uchwyt = szukajUchwytu($hSub, "Anuluj", "Button")
;~ 				EndIf
;~ 			Else
;~ 				ControlClick($uchwyt, "", "")
;~ 				$var = True
;~ 			EndIf
;~ 		Else
;~ 			ControlClick($uchwyt, "", "")
;~ 		EndIf
;~ 	ElseIf WinGetHandle("[REGEXPTITLE:(?i)(.* - brak towaru w magazynie)]") <> 0 Then
;~ 		zapiszBrakTowaru($lista)
;~ 		$var = True
;~ 	EndIf
;~ 	Return $var
;~ EndFunc   ;==>oknoSubiekt

;~ ;---------------------------------------------------------------------------------------------------------------------
;~ ;Funkcja pobiera liste zwrotow detalicznych i zapisuje do tablicy $_PARAGONY
;~ Func pobierzListeZwrotow($magazyn)
;~ 	logs("========================= START " & $magazyn & " =============================")
;~ 	wybierzListe("ZW")
;~ 	;------Ustawianie daty-------------------------------------
;~ 	$hWnd = WinGetHandle('[REGEXPTITLE:(?i)(.* - Subiekt GT)]')
;~ 	$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Zwroty detaliczne", "ins_hyperlink", "Dokumenty z okresu:", "Static")
;~ 	$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
;~ 	ControlFocus($uchwyt, "", "")
;~ 	ControlClick($uchwyt, "", "")
;~ 	Send("{END}")
;~ 	Send("{UP}{UP}{UP}{UP}{UP}{UP}")
;~ 	Send("{ENTER}")
;~ 	;---------------------
;~ 	$mag = 0
;~ 	For $i = 0 To UBound($_ListaZwrotow) - 1
;~ 		If $_ListaZwrotow[$i][0] = $magazyn Then
;~ 			$mag = $i
;~ 			ExitLoop
;~ 		EndIf
;~ 	Next
;~ 	$_ListaZwrotow[$mag][1] = ObjCreate("Scripting.Dictionary")
;~ 	If Not czyPusta("ZW") Then
;~ 		fokusNaListeClick("ZW")
;~ 		kopiujListe()
;~ 		$tabLinii = StringSplit(ClipGet(), @LF, 1)
;~ 		For $i = 2 To $tabLinii[0] - 1
;~ 			$tmp = StringSplit($tabLinii[$i], "	", 1)
;~ 			$_ListaZwrotow[$mag].Add($tmp[6], $i - 2)
;~ 			logs($tmp[4] & " - " & $tmp[6])
;~ 		Next
;~ 	EndIf
;~ 	logs("========================= KONIEC " & $magazyn & " =============================")
;~ EndFunc   ;==>pobierzListeZwrotow

;~ ;---------------------------------------------------------------------------------------------------------------------
;~ ;Funkcja sprawdza czy na liscie zwrotow detalicznych znajduje sie paragon podany w parametrze
;~ ;

;~ Func sprawdzListeZwrotow($nazwa)
;~ 	$magazyn = ""
;~ 	$magTab = StringSplit($nazwa, "/", 1)

;~ 	Switch $magTab[2]
;~ 		Case "MAG"
;~ 			$magazyn = $MAG
;~ 		Case "ABR"
;~ 			$magazyn = $ABR
;~ 		Case "KUN"
;~ 			$magazyn = $KUN
;~ 		Case "TAR"
;~ 			$magazyn = $TAR
;~ 	EndSwitch
;~ 	If $_PARAGONY[$magazyn].Count < 1 Then
;~ 		Return False
;~ 	EndIf
;~ 	If $_PARAGONY[$magazyn].Exists($nazwa) Then
;~ 		$_PARAGONY[$magazyn].Remove($nazwa)
;~ 		Return True
;~ 	Else
;~ 		Return False
;~ 	EndIf
;~ EndFunc   ;==>sprawdzListeZwrotow

;~ ;---------------------------------------------------------------------------------------------------------------------
;~ ;Funkcja sprawdza czy do aktualnie przerabianego paragonu jest jakis zwrot detaliczny
;~ Func zwrotDetaliczny($paragon)
;~ 	$lista = "ZW"
;~ 	wybierzListe($lista)
;~ 	If Not czyPusta($lista) Then
;~ 		fokusNaListeClick($lista)
;~ 		homeBtn()
;~ 		endBtn()
;~ 		kopiujPozycje()
;~ 		$endpos = ClipGet()
;~ 		$curpos = ""
;~ 		homeBtn()
;~ 		While $endpos <> $curpos
;~ 			$curpos = kopiujPozycje()
;~ 			If StringInStr($curpos, "wywołuje skutek magazynowy") <> 0 Then
;~ 				If StringInStr($curpos, $paragon) <> 0 Then
;~ 					$i = 0
;~ 					Do
;~ 						odlozSkutek($lista)
;~ 						Sleep(1000)
;~ 						$i = $i + 1
;~ 					Until StringInStr(kopiujPozycje(), "wywołuje skutek magazynowy") = 0 Or $i >= 3
;~ 					If $i >= 3 Then
;~ 						logs("Nie udalo sie odlozyc skutku zwrotu: " & $paragon)
;~ 					EndIf
;~ 					ExitLoop
;~ 				EndIf
;~ 			EndIf
;~ 			downBtn()
;~ 		WEnd
;~ 	EndIf
;~ 	wybierzListe("PA")
;~ 	;homeBtn()
;~ EndFunc   ;==>zwrotDetaliczny



;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Zamyka program po wcisnieciu BREAK(ctrl+break)
Func zamknij()
	Exit
EndFunc   ;==>zamknij

;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Wstrzymuje działanie programu
Func pauza()
	If $_PAUSE = False Then
		$_PAUSE = True
		logs("PAUZA........." & @LF)
		;pauzaLoop()
	Else
		$_PAUSE = False
		logs("START........." & @LF)
		;ControlFocus($t, "", "")
	EndIf
EndFunc   ;==>pauza
Func pauzaLoop()
	While $_PAUSE
		Sleep(1000)
	WEnd
EndFunc   ;==>pauzaLoop
Func pauzaSzybka()
	If $_PAUSE = False Then
		$_PAUSE = True
		logs("PAUZA........." & @LF)
		pauzaLoop()
	Else
		$_PAUSE = False
		logs("START........." & @LF)
;		ControlFocus($t, "", "")
	EndIf
EndFunc   ;==>pauzaSzybka




;---------------------------------------------------------------------------------------------------------------------
Func wyslijSms($tresc)
	Run('SmsEagle.exe "' & $tresc & '"', "", @SW_HIDE)
EndFunc   ;==>wyslijSms

;---------------------------------------------------------------------------------------------------------------------
Func sprawdzDate($data)
	If StringRegExp($data, "20[0-2][0-9]/[0-1][0-9]/[0-3][0-9]") <> 1 Then
		Return False
	EndIf
	If _DateTimeFormat($data, 0) = "" Then
		Return False
	EndIf
	Return True
EndFunc   ;==>sprawdzDate


;---------------------------------------------------------------------------------------------------------------------
;
;~ Func przerwaNaBackup($podmiot)
;~ 	logs("")
;~ 	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
;~ 	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ PRZERWA NA BACKUP $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
;~ 	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
;~ 	logs("")
;~ 	$procTab = ProcessList("Subiekt.exe")
;~ 	$subiektPath = _WinAPI_GetProcessFileName($procTab[1][1])
;~ 	ProcessClose("Subiekt.exe")
;~ 	Do
;~ 		oknoArchiwizacja()
;~ 	Until oknoArchiwizacja() = 0 And uchwycSubiekta() = 0
;~ 	Do
;~ 		Sleep(1000)
;~ 	Until @HOUR = $godz And @MIN >= $minut2

;~ 	Run($subiektPath, "", @SW_SHOWMAXIMIZED)
;~ 	While wykryjEtapLogowania() = "brak"
;~ 		Sleep(1000)
;~ 	WEnd

;~ 	$obsluzono = False
;~ 	While $obsluzono = False
;~ 		$obsluzono = obsluzEtapLogowania(wykryjEtapLogowania(), $podmiot)
;~ 		Sleep(2000)
;~ 	WEnd
;~ 	logs("")
;~ 	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
;~ 	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ KONIEC PRZERWY NA BACKUP $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
;~ 	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
;~ 	logs("")
;~ EndFunc   ;==>przerwaNaBackup


;~ ;---------------------------------------------------------------------------------------------------------------------
;~ Func obsluzEtapLogowania($etap, $podmiot)
;~ 	$hWnd = uchwycSubiekta()
;~ 	Switch $etap
;~ 		Case "podmiot"
;~ 			$uchwyt = fokusNaListePodmiotow()
;~ 			$pos = kopiujPozycje()
;~ 			While StringStripCR(StringStripWS($pos, $STR_STRIPALL)) <> $podmiot
;~ 				downBtn()
;~ 				$pos = kopiujPozycje()
;~ 			WEnd
;~ 			$uchwyt = szukajUchwytu($hWnd, "OK", "Button")
;~ 			ControlClick($uchwyt, "", "")
;~ 		Case "uzytkownik"
;~ 			$uchwyt = szukajUchwytu($hWnd, "", "ComboBox")
;~ 			ControlFocus($uchwyt, "", "")
;~ 			Send("M")
;~ 			$uchwyt = szukajUchwytu($hWnd, "OK", "Button")
;~ 			ControlClick($uchwyt, "", "")
;~ 		Case "zalogowano"
;~ 			If StringInStr(WinGetTitle($hWnd), $podmiot & " na serwerze") = 0 Then
;~ 				Send("{CTRLDOWN}{F4}")
;~ 				Send("{CTRLUP}")
;~ 			Else
;~ 				Return True
;~ 			EndIf
;~ 		Case Else
;~ 			restartujSubiekta("Subiekt GT")
;~ 	EndSwitch
;~ 	Return False
;~ EndFunc   ;==>obsluzEtapLogowania

;~ ;-----------------------------------------------------------------------------------------------------------
;~ Func uchwycSubiekta()
;~ 	$hWnd = WinGetHandle($t)
;~ 	$uchwyt = szukajUchwytu($hWnd, "Magazyn:", "ins_combohyperlink")
;~ 	If $uchwyt <> 0 Then
;~ 		Return $hWnd
;~ 	EndIf
;~ 	$hWnd = WinGetHandle("Subiekt GT")
;~ 	$uchwyt = szukajUchwytu($hWnd, "Magazyn:", "ins_combohyperlink")
;~ 	If $uchwyt <> 0 Then
;~ 		Return $hWnd
;~ 	EndIf
;~ 	Return 0
;~ EndFunc   ;==>uchwycSubiekta

;~ ;-----------------------------------------------------------------------------------------------------------
;~ Func restartujSubiekta($nazwa)
;~ 	$procTab = ProcessList("Subiekt.exe")
;~ 	$subiektPath = _WinAPI_GetProcessFileName($procTab[1][1])
;~ 	ProcessClose("Subiekt.exe")
;~ 	Do
;~ 		oknoArchiwizacja()
;~ 	Until oknoArchiwizacja() = 0 And (WinGetHandle($nazwa) = 0 Or uchwycSubiekta())
;~ 	Sleep(5000)
;~ 	Run($subiektPath)
;~ 	While wykryjEtapLogowania() = "brak"
;~ 		Sleep(1000)
;~ 	WEnd
;~ EndFunc   ;==>restartujSubiekta

;~ ;-----------------------------------------------------------------------------------------------------------
;~ Func wykryjEtapLogowania()
;~ 	$hWnd = uchwycSubiekta()
;~ 	$uchwyt = szukajUchwytu($hWnd, "Wybór podmiotu", "#32770")
;~ 	If $uchwyt <> 0 Then
;~ 		Return "podmiot"
;~ 	EndIf
;~ 	$uchwyt = szukajUchwytu($hWnd, "Wybierz użytkownika, z którym chcesz rozpocząć pracę", "Static")
;~ 	If $uchwyt <> 0 Then
;~ 		Return "uzytkownik"
;~ 	EndIf
;~ 	$uchwyt = szukajUchwytu($hWnd, "", "ins_svcwindow")
;~ 	If $uchwyt <> 0 Then
;~ 		Return "zalogowano"
;~ 	Else
;~ 		Return "brak"
;~ 	EndIf
;~ EndFunc   ;==>wykryjEtapLogowania

;~ ;-----------------------------------------------------------------------------------------------------------
;~ Func oknoArchiwizacja()
;~ 	$hSub = WinGetHandle("Subiekt GT")
;~ 	If $hSub <> 0 Then
;~ 		$uchwyt = szukajUchwytu($hSub, "&Nie", "Button")
;~ 		ControlClick($uchwyt, "", "")
;~ 	EndIf
;~ 	Return $hSub
;~ EndFunc   ;==>oknoArchiwizacja

;~ ;---------------------------------------------------------------------------------------------------------------------
;~ ;Złapanie fokusa na liste
;~ Func fokusNaListePodmiotow()
;~ 	Local $uchwyt = 0
;~ 	Local $hWnd = uchwycSubiekta()
;~ 	$uchwyt = szukajUchwytu($hWnd, "", "GXWND")
;~ 	$tmp = ControlGetPos($uchwyt, "", "")
;~ 	ControlClick($uchwyt, "", "", "left", 1, $tmp[0], 22)
;~ 	$pos = kopiujPozycje()
;~ 	If $pos <> "" Then
;~ 		Return $uchwyt
;~ 	EndIf
;~ 	$i = 0
;~ 	Do
;~ 		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i)
;~ 		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 1)
;~ 		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 2)
;~ 		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 3)
;~ 		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 4)
;~ 		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 5)
;~ 		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 6)
;~ 		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 7)
;~ 		$pos = kopiujPozycje()
;~ 		$i = $i + 8
;~ 	Until $pos <> "" Or $i >= $tmp[3]
;~ 	Return $uchwyt
;~ EndFunc   ;==>fokusNaListePodmiotow
