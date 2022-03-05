;@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
;@#####################################################################################################################@
;@###############################@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@##################################@
;@###############################@													@##################################@
;@###############################@		Program do odkladania i wywolywania			@##################################@
;@###############################@		skutkow magazynowych w programie 			@##################################@
;@###############################@		Subiekt GT									@##################################@
;@###############################@													@##################################@
;@###############################@						Mateusz Blaszczak	2018	@##################################@
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




HotKeySet("{F6}", "zamknij")
HotKeySet("{PAUSE}", "pauza")
HotKeySet("{F4}", "pauzaSzybka")

$odp = $IDNO
While $odp == $IDNO
$odp = MsgBox($MB_YESNOCANCEL, "Start", "Proszê w³¹czyæ na liscie zwrotow kolumne Dokument korygowany oraz na ka¿dej liœcie ustawiæ sortowanie od najm³odszych dokumentów do najstarszych!" & @LF & "Czy chcesz rozpocz¹æ pracê?")
WEnd
If $odp == $IDCANCEL Then
	Exit
EndIf

$odp = MsgBox($MB_YESNO, "Czyszczenie loga", "Czy chcesz wyczyscic 'logWSM.log'?")
If $odp == $IDYES Then
	$hF = FileOpen("logWSM.log", 2)
	FileClose($hF)
EndIf


$odp = MsgBox($MB_YESNO, "Czyszczenie brakuj¹cych towarów", "Czy chcesz wyczyscic 'WSMbraktowaru.txt'?")
If $odp == $IDYES Then
	$hF = FileOpen("WSMbraktowaru.txt", 2)
	FileClose($hF)
EndIf
;_______________________________________Inicjalizacja zmiennych______________________________________________
Global $_PARAGONY[4]
Global $_ZWROTY
Global $listaBrakujacychTowarow = ObjCreate("Scripting.Dictionary")
Global Enum $MAG = 0, $ABR = 1, $KUN = 2, $TAR = 3
Global $_PAUSE = False
Global Const $t = "[REGEXPTITLE:(?i)(.* - Subiekt GT)]"
Global $ileBrakow = 0
Global $jestZwrot = False
Global $nazwaParagonuZwrotu = ""
Global $_DATA = "2018/01/02"
Global $_PODMIOT = "ACTUS_2011"
Global $godz1 = 23
Global $godz2 = 00
Global $minut1 = 45
Global $minut2 = 15

Do
	$_DATA = InputBox("Data", "WprowadŸ datê pocz¹tkow¹" & @LF & "(format: yyyy/mm/dd)")
	If @error = 1 Then
		Exit
	EndIf
	If Not sprawdzDate($_DATA) Then
		MsgBox($MB_OK, "Error", "Data niepoprawna")
	EndIf
Until sprawdzDate($_DATA)

Global $_ENDDATA = "2018/06/18"
Do
	$_ENDDATA = InputBox("Data", "WprowadŸ datê koñcow¹" & @LF & "(format: yyyy/mm/dd)")
	If @error = 1 Then
		Exit
	EndIf
	If Not sprawdzDate($_ENDDATA) Then
		MsgBox($MB_OK, "Error", "Data niepoprawna")
	EndIf
Until sprawdzDate($_ENDDATA)

$_ENDDATA = _DateAdd('d', 1, $_ENDDATA)
Global $iloscLiniiBrakow = 0
$hWnd = WinGetHandle($t)
ControlFocus($t, "", "")

;____________Petla programowa_______________________________________________________________________________
logs("START BotWSM")
;~ logs("Pobieranie listy zwrotow")
;~ magazyn("MAG")
;~ pobierzListeZwrotow($MAG)
;~ magazyn("ABR")
;~ pobierzListeZwrotow($ABR)
;~ magazyn("KUN")
;~ pobierzListeZwrotow($KUN)
;~ magazyn("TAR")
;~ pobierzListeZwrotow($TAR)
logs("START PETLI")
$licznikg = 0



While $_DATA <> $_ENDDATA
	If $licznikg > 10 Then
		;wyslijSms($_DATA & " - " & $iloscLiniiBrakow)
		$licznikg = 0
	EndIf

	If WinGetHandle("[REGEXPTITLE:(?i)(Informacja o nowej wersji*)]") <> 0 Then
		$x = szukajUchwytu(WinGetHandle("[REGEXPTITLE:(?i)(Informacja o nowej wersji*)]"), "&Zamknij", "Button")
		ControlClick($x, "", "")
	EndIf

	$licznikg = $licznikg + 1
	$listaBrakujacychTowarow = 0
	$listaBrakujacychTowarow = ObjCreate("Scripting.Dictionary")
	wykonujMAG("MAG")
	wykonujABR("ABR")
	wykonujKUN("KUN")
	wykonujTAR("TAR")
	$_DATA = _DateAdd('d', 1, $_DATA)
	If @HOUR = $godz1 And @MIN >= $minut1 Then
		przerwaNaBackup($_PODMIOT)

		$hWnd = WinGetHandle($t)
		ControlFocus($t, "", "")
	EndIf
WEnd
logs("KONIEC BotWSM")
wyslijSms("KONIEC BotWSM")

Exit

;#####################################################################################################################
;#####################################################################################################################
;################################													##################################
;################################						FUNKCJE						##################################
;################################													##################################
;#####################################################################################################################
;#####################################################################################################################

;---------------------------------------------------------------------------------------------------------------------
;Funkcja zmienia magazyn na ktorym mamy dzia³aæ oraz uruchamia wywo³ywania skutkow na ka¿dej liscie
;Przyjmuje nazwê listy na ktorej dzia³amy
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne

Func wykonujMAG($magazyn)
	magazyn($magazyn)
	logs("#################################### START " & $magazyn & " ###########################################")
;~ Sleep(1000)
;~    wykonujListe("FA")
	Sleep(500)
	wykonujListe("PW")
	Sleep(500)
	wykonujListe("MM")
	Sleep(500)
	wykonujListeZwrotow("ZW")
	Sleep(500)
	wykonujListe("PA")
	Sleep(500)
	wykonujListe("RW")
	Sleep(500)

	logs("#################################### KONIEC " & $magazyn & " ###########################################")
EndFunc   ;==>wykonujMAG
Func wykonujABR($magazyn)
	magazyn($magazyn)
	logs("#################################### START " & $magazyn & " ###########################################")
;~ Sleep(1000)
;~    wykonujListe("FA")
	Sleep(500)
	wykonujListe("PW")
	Sleep(500)
	wykonujListe("MM")
	Sleep(500)
	wykonujListeZwrotow("ZW")
	Sleep(500)
	wykonujListe("PA")
	Sleep(500)
	wykonujListe("RW")
	Sleep(500)

	logs("#################################### KONIEC " & $magazyn & " ###########################################")
EndFunc   ;==>wykonujABR
Func wykonujKUN($magazyn)
	magazyn($magazyn)
	logs("#################################### START " & $magazyn & " ###########################################")
	Sleep(500)
	wykonujListe("MM")
	logs("#################################### KONIEC " & $magazyn & " ###########################################")
EndFunc   ;==>wykonujKUN
Func wykonujTAR($magazyn)
	If $_DATA <> "2018/03/10" And $_DATA <> "2018/03/11" Then
		Return
	EndIf
	magazyn($magazyn)
	logs("#################################### START " & $magazyn & " ###########################################")
	Sleep(500)
	wykonujListe("PA")


	logs("#################################### KONIEC " & $magazyn & " ###########################################")
EndFunc   ;==>wykonujTAR

;---------------------------------------------------------------------------------------------------------------------
;Funkcja wykonuje wszystkie instrukcje potrzebne do wywo³ania skutkow magazynowych wszystkich pozycji na konkretnej liscie
;Przyjmuje nazwê listy na ktorej dzia³amy
;



Func wykonujListe($lista)
	wybierzListe($lista)
	Sleep(500)
	ustawOkres($_DATA, $lista)
	Sleep(500)
	logs("------------------------------- START " & $lista & " ----------------------------------------")
	If $_PAUSE = True Then
		pauzaLoop()

	EndIf
	If Not czyPusta($lista) Then
		Do
			$ileBrakow = 0
			fokusNaListeClick($lista)

			homeBtn()
			endBtn()
			kopiujPozycje()
			$endpos = ClipGet()
			$curpos = ""
			homeBtn()
			$i = 0
			While $endpos <> $curpos
				fokusNaListe($lista)
				$czyBrakTowaru = False
				Do
					;logs("1 $curpos = kopiujPozycje()")
					$curpos = kopiujPozycjeWiersz()
				Until StringInStr($curpos, "skutek") <> 0 Or $curpos = "STOP"
				If $curpos = "STOP" Then
					$brak = WinGetHandle("[REGEXPTITLE:(?i)(.* - brak towaru w magazynie)]")
					$uchwyt = szukajUchwytu($brak, "Zamknij", "Button")
					ControlClick($uchwyt, "", "")
				Else
					If StringInStr($curpos, "ony skutek magazynowy") <> 0 Then
						While Not $czyBrakTowaru
							;logs("2 $curpos2 = kopiujPozycje()")
							$curpos2 = kopiujPozycje()
							If StringInStr($curpos2, "ony skutek magazynowy") <> 0 And Not $czyBrakTowaru Then
								$czyBrakTowaru = wywolajSkutek($lista)
							Else
								$czyBrakTowaru = True
							EndIf
						WEnd
					EndIf
				EndIf
				Sleep(10)
				If $_PAUSE = True Then
					pauzaLoop()
					ExitLoop
				EndIf
				ControlFocus($t, "", "")
				downBtn()
			WEnd
			fokusNaListe($lista)
			While kopiujListe() = ""
				Sleep(10)
			WEnd
			StringReplace(ClipGet(), "ony skutek magazynowy", "xxx")
			$ileZnaleziono = @extended
			If $ileZnaleziono <> $ileBrakow Then
				logs("OMINIETO PRODUKT! " & "; Brakuje: " & $ileBrakow & ", znaleziono: " & @extended)
			EndIf
		Until $ileZnaleziono = $ileBrakow
	EndIf
	logs("------------------------------- KONIEC " & $lista & " ----------------------------------------")
EndFunc   ;==>wykonujListe

;---------------------------------------------------------------------------------------------------------------------
;Funkcja wykonuje wszystkie instrukcje potrzebne do wywo³ania skutkow magazynowych wszystkich pozycji na konkretnej liscie
;Przyjmuje nazwê listy na ktorej dzia³amy
;



Func wykonujListeZwrotow($lista)
	wybierzListe($lista)
	Sleep(500)
	ustawOkres($_DATA, $lista)
	Sleep(500)
	$_ZWROTY = ObjCreate("Scripting.Dictionary")
	logs("------------------------------- START " & $lista & " ----------------------------------------")
	If $_PAUSE = True Then
		pauzaLoop()

	EndIf
	If Not czyPusta($lista) Then
		Do

			$ileBrakow = 0
			fokusNaListeClick($lista)

			homeBtn()
			endBtn()
			kopiujPozycje()
			$endpos = ClipGet()
			$curpos = ""
			homeBtn()
			$i = 0
			While $endpos <> $curpos
				fokusNaListe($lista)
				$czyBrakTowaru = False
				Do
					;logs("1 $curpos = kopiujPozycje()")
					$curpos = kopiujPozycjeWiersz()
				Until StringInStr($curpos, "skutek") <> 0 Or $curpos = "STOP"
				If $curpos = "STOP" Then
					$brak = WinGetHandle("[REGEXPTITLE:(?i)(.* - brak towaru w magazynie)]")
					$uchwyt = szukajUchwytu($brak, "Zamknij", "Button")
					ControlClick($uchwyt, "", "")
				Else
					If StringInStr($curpos, "ony skutek magazynowy") <> 0 Then
						;$xyz = 0
						While Not $czyBrakTowaru
							;logs("2 $curpos2 = kopiujPozycje()")
							$curpos2 = kopiujPozycje()
							If StringInStr($curpos2, "ony skutek magazynowy") <> 0 And Not $czyBrakTowaru Then
								$tmp = StringSplit($curpos2, "	", 1)
								If $tmp[9] <> StringReplace($_DATA, '/', '-') Then
									$czyBrakTowaru = wywolajSkutek($lista)
									If $czyBrakTowaru = True Then
										$ileBrakow = $ileBrakow + 1
									EndIf
								Else
									;MsgBox($IDOK,"asd","asd")
									$_ZWROTY.Add($tmp[6], $ileBrakow)

									$czyBrakTowaru = True
									$ileBrakow = $ileBrakow + 1
								EndIf
							Else

								$czyBrakTowaru = True
							EndIf
							;$xyz = $xyz + 1
						WEnd
;~
					EndIf
				EndIf
				Sleep(10)
				If $_PAUSE = True Then
					pauzaLoop()
					ExitLoop
				EndIf
				ControlFocus($t, "", "")
				downBtn()
			WEnd
			fokusNaListe($lista)
			While kopiujListe() = ""
				Sleep(10)
			WEnd
			StringReplace(ClipGet(), "ony skutek magazynowy", "xxx")
			$ileZnaleziono = @extended
			If $ileZnaleziono <> $ileBrakow Then
				logs("OMINIETO PRODUKT! " & "; Brakuje: " & $ileBrakow & ", znaleziono: " & @extended)
			EndIf
		Until $ileZnaleziono = $ileBrakow
	EndIf
	logs("------------------------------- KONIEC " & $lista & " ----------------------------------------")
EndFunc   ;==>wykonujListeZwrotow


;-----------------------------------------------------------------------------------------------------------------------

;---------------------------------------------------------------------------------------------------------------------
;Funkcja wywo³uje skutek magazynowy
;Przyjmuje nazwê listy na ktorej dzia³amy
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne

Func wywolajSkutek($lista)
	$czyBrakTowaru = False
	$okno = ""
	$paragon = ""
	$tmp = StringSplit(ClipGet(), "	", 1)
	Switch $lista
		Case "PA"
			Send("{LALT}OW")
			$okno = "Paragon"
			$paragon = $tmp[6]
		Case "FA"
			Send("{LALT}OW")
			$okno = "Faktura VAT zakupu"
			$paragon = $tmp[5]
		Case "MM"
			Send("{LALT}OS{ENTER}")
			$okno = "Przesuniêcie miêdzymagazynowe"
			$paragon = $tmp[4]
		Case "RW"
			Send("{LALT}OS{ENTER}")
			$okno = "Rozchód wewnêtrzny"
			$paragon = $tmp[4]
		Case "PW"
			Send("{LALT}OS")
			$okno = "Przychód wewnêtrzny"
			$paragon = $tmp[4]
		Case "ZW"
			Send("{LALT}OW")
			$okno = "Zwrot ze sprzeda¿y detalicznej"
			$paragon = $tmp[4]
	EndSwitch
	$var = False
	;logs($paragon & " ")
	$licznik = 0
	While WinGetHandle($okno) = 0 And Not $var And $licznik < 3000 ;ControlFocus($t,"","") = 0 Or WinGetHandle("Subiekt GT") <> 0
		$var = oknoSubiekt($lista)
		$licznik = $licznik + 1
		;ConsoleWrite(WinGetHandle($okno) & " * " & $var & " * " )
	WEnd
	$licznik = 0
	While WinGetHandle($okno) <> 0 And Not $var And $licznik < 3000 ;ControlFocus($t,"","") = 0 Or WinGetHandle("Subiekt GT") <> 0
		$var = oknoSubiekt($lista)
		$licznik = $licznik + 1
		;ConsoleWrite(WinGetHandle($okno) & " - " & $var & " - " )
	WEnd
	;logs(@LF)
	If $lista = "PA" Then
		If $_ZWROTY.Count > 0 And $_ZWROTY.Exists($paragon) Then
			logs("Znaleziono zwrot")
			zwrotDetaliczny($paragon)
			$_ZWROTY.Remove($paragon)
		EndIf
	EndIf

	If $var Then
		logs("Nie wywolano skutku (brak towaru): " & $paragon)
	Else
		logs("Wywolano skutek: " & $paragon)
	EndIf
	Return $var
EndFunc   ;==>wywolajSkutek

;-----------------------------------------------------------------------------------------------------------------------


;---------------------------------------------------------------------------------------------------------------------
;Funkcja obsluguje okienko brakuj¹cego towaru. Zapisuje do pliku wszystkie pozycje z listy a nastepnie zamyka okienko
;przyjmuje parametr okreslajacy czy dany produkt juz zosta³ zapisany oraz nazwe przerabianej listy

Func zapiszBrakTowaru($lista)
	$brak = WinGetHandle("[REGEXPTITLE:(?i)(.* - brak towaru w magazynie)]")
	$uchwyt = szukajUchwytu($brak, "", "GXWND")
	$paragon = ""
	$data = ""

	ControlFocus($uchwyt, "", "")
	$tmp = StringSplit(ClipGet(), "	", 1)
	Switch $lista
		Case "PA"
			$data = $tmp[3]
			$paragon = $tmp[6]
		Case "FA"
			$data = $tmp[3]
			$paragon = $tmp[5]
		Case "MM"
			$data = $tmp[2]
			$paragon = $tmp[4]
		Case "RW"
			$data = $tmp[2]
			$paragon = $tmp[4]
		Case "PW"
			$data = $tmp[2]
			$paragon = $tmp[4]
		Case "ZW"
			$data = $tmp[2]
			$paragon = $tmp[4]
	EndSwitch
	If Not $listaBrakujacychTowarow.Exists($paragon) Then
		$hFile = FileOpen("WSMbraktowaru.txt", 1)
		FileWrite($hFile, $paragon & ";" & $data & @LF)
		$iloscLiniiBrakow = $iloscLiniiBrakow + 1
		kopiujListe()
		$tmp2 = StringSplit(ClipGet(), @LF, 1)
		For $i = 2 To UBound($tmp2) - 1
			FileWrite($hFile, $tmp2[$i])
			$iloscLiniiBrakow = $iloscLiniiBrakow + 1
		Next
		FileClose($hFile)
		$listaBrakujacychTowarow.Add($paragon, $listaBrakujacychTowarow.Count)
		If $lista <> "ZW" Then
			$ileBrakow = $ileBrakow + 1
		EndIf
		logs("Zapisano brak towaru: " & $paragon)
	Else
		logs("Proba zduplikowania brakujacego towaru: " & $paragon)
		If $lista <> "ZW" Then
			$ileBrakow = $ileBrakow + 1
		EndIf
	EndIf
	$uchwyt = szukajUchwytu($brak, "Zamknij", "Button")
	ControlClick($uchwyt, "", "")

EndFunc   ;==>zapiszBrakTowaru

Func logs($text)
	_FileWriteLog("logWSM.log", "[" & $_DATA & "] " & $text)
	ConsoleWrite("[" & $_DATA & "] " & $text & @LF)
EndFunc   ;==>logs

;---------------------------------------------------------------------------------------------------------------------
;Funkcja obsluguje okienko potwierdzenia wywolania skutku
;


Func oknoSubiekt($lista)
	$var = False
	$hSub = WinGetHandle("Subiekt GT")

	If $hSub <> 0 Then
		$uchwyt = szukajUchwytu($hSub, "&Tak", "Button")
		If $uchwyt = 0 Then
			$uchwyt = szukajUchwytu($hSub, "OK", "Button")
			If $uchwyt = 0 Then
				$uchwyt = szukajUchwytu($hSub, "Pomoc", "Button")
				If $uchwyt = 0 Then
					If _WinAPI_GetClassName($hSub) = "BalloonHelpClass" Then
						Send("{ESC}")
						$var = True
					EndIf
				Else
					$uchwyt = szukajUchwytu($hSub, "Anuluj", "Button")
				EndIf
			Else
				ControlClick($uchwyt, "", "")
				$var = True
			EndIf
		Else
			ControlClick($uchwyt, "", "")
		EndIf
	ElseIf WinGetHandle("[REGEXPTITLE:(?i)(.* - brak towaru w magazynie)]") <> 0 Then
		;MsgBox($IDOK,"Brak towaru","Brak towaru")
		zapiszBrakTowaru($lista)
		$var = True
	EndIf
	Return $var
EndFunc   ;==>oknoSubiekt

;---------------------------------------------------------------------------------------------------------------------
;Funkcja pobiera liste zwrotow detalicznych i zapisuje do tablicy $_PARAGONY
;

Func pobierzListeZwrotow($magazyn)
	logs("========================= START " & $magazyn & " =============================")
	wybierzListe("ZW")
	;------Ustawianie daty-------------------------------------
	$hWnd = WinGetHandle('[REGEXPTITLE:(?i)(.* - Subiekt GT)]')
	$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Zwroty detaliczne", "ins_hyperlink", "Dokumenty z okresu:", "Static")
	$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	ControlFocus($uchwyt, "", "")
	ControlClick($uchwyt, "", "")
	Send("{END}")
	Send("{UP}{UP}{UP}{UP}{UP}{UP}")
	Send("{ENTER}")
	;---------------------
	$_PARAGONY[$magazyn] = ObjCreate("Scripting.Dictionary")
	If Not czyPusta("ZW") Then
		fokusNaListe("ZW")
		kopiujListe()
		$tabLinii = StringSplit(ClipGet(), @LF, 1)

		For $i = 2 To $tabLinii[0] - 1
			$tmp = StringSplit($tabLinii[$i], "	", 1)
			$_PARAGONY[$magazyn].Add($tmp[6], $i - 2)
			logs($tmp[4] & " - " & $tmp[6])
		Next
	EndIf
	logs("========================= KONIEC " & $magazyn & " =============================")
EndFunc   ;==>pobierzListeZwrotow

;---------------------------------------------------------------------------------------------------------------------
;Funkcja sprawdza czy na liscie zwrotow detalicznych znajduje sie paragon podany w parametrze
;

Func sprawdzListeZwrotow($nazwa)
	$magazyn = ""
	$magTab = StringSplit($nazwa, "/", 1)

	Switch $magTab[2]
		Case "MAG"
			$magazyn = $MAG
		Case "ABR"
			$magazyn = $ABR
		Case "KUN"
			$magazyn = $KUN
		Case "TAR"
			$magazyn = $TAR
	EndSwitch
	If $_PARAGONY[$magazyn].Count < 1 Then
		Return False
	EndIf
	If $_PARAGONY[$magazyn].Exists($nazwa) Then
		$_PARAGONY[$magazyn].Remove($nazwa)
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>sprawdzListeZwrotow

;---------------------------------------------------------------------------------------------------------------------
;Funkcja sprawdza czy do aktualnie przerabianego paragonu jest jakis zwrot detaliczny
;

Func zwrotDetaliczny($paragon)

	$lista = "ZW"
	wybierzListe($lista)
	If Not czyPusta($lista) Then
		fokusNaListeClick($lista)
		homeBtn()
		endBtn()
		kopiujPozycje()
		$endpos = ClipGet()
		$curpos = ""
		homeBtn()
		While $endpos <> $curpos
			Do
				$curpos = kopiujPozycjeWiersz()
			Until StringInStr($curpos, "skutek") <> 0 Or $curpos = "STOP"
			If StringInStr($curpos, "ony skutek magazynowy") <> 0 Then
				If StringInStr($curpos, $paragon) <> 0 Then
					$i = 0
					Do
						While StringInStr(kopiujPozycje(), "ony skutek magazynowy") > 0 And $i < 3
							wywolajSkutek($lista)
							;ConsoleWrite("ZWROT" & @LF)
							Sleep(1000)
							$i = $i + 1
						WEnd
					Until StringInStr(kopiujPozycje(), "ony skutek magazynowy") = 0 Or $i >= 3
					If $i >= 3 Then
						logs("Nie udalo sie wywolac skutku zwrotu: " & $paragon)
					EndIf
					ExitLoop
				EndIf

			EndIf
			downBtn()
		WEnd
	EndIf
	wybierzListe("PA")
	;homeBtn()
EndFunc   ;==>zwrotDetaliczny

;---------------------------------------------------------------------------------------------------------------------
;Funkcja uruchamia listy dokumentów
;Przyjmuje nazwê listy któr¹ chcemy wyœwietliæ
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne

Func wybierzListe($lista)
	ControlFocus($t, "", "")
	Switch $lista
		Case "PA"
			Send("{ALTDOWN}2")
			Send("{ALTUP}")
		Case "FA"
			Send("{ALTDOWN}3")
			Send("{ALTUP}")
		Case "MM"
			ControlFocus($t, "", "")
			Send("{ALTDOWN}4")
			Send("{ALTUP}")
			$hWnd = WinGetHandle('[REGEXPTITLE:(?i)(.* - Subiekt GT)]')
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Wydania magazynowe", "ins_hyperlink", ", typu:", "Static")
			$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
			ControlFocus($uchwyt, "", "")
			ControlClick($uchwyt, "", "")
			Send("{END}{ENTER}")
		Case "RW"
			Send("{ALTDOWN}4")
			Send("{ALTUP}")
			$hWnd = WinGetHandle('[REGEXPTITLE:(?i)(.* - Subiekt GT)]')
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Wydania magazynowe", "ins_hyperlink", ", typu:", "Static")
			$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
			ControlFocus($uchwyt, "", "")
			ControlClick($uchwyt, "", "")
			Send("{END}{UP}{ENTER}")
		Case "PW"
			Send("{ALTDOWN}5")
			Send("{ALTUP}")
		Case "ZW"
			Send("{LALT}WSZ")

	EndSwitch

EndFunc   ;==>wybierzListe

;---------------------------------------------------------------------------------------------------------------------
;Sprawdza czy lista jest pusta
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne

Func czyPusta($lista)
	fokusNaListe($lista)
	While kopiujListe() = ""
		Sleep(10)
	WEnd

	$tmp = StringSplit(ClipGet(), @LF)
	If $tmp[0] < 2 Then
		Return True
	ElseIf StringInStr(ClipGet(), "ony skutek magazynowy") = 0 Then
		Return True
	EndIf
	If StringLen($tmp[2]) = 0 Then
		logs("PUSTA LISTA")
		Return True
	Else
		Return False
	EndIf
	Return False
EndFunc   ;==>czyPusta

;---------------------------------------------------------------------------------------------------------------------
;Z³apanie fokusa na liste dokumentow
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne

Func fokusNaListeClick($nazwa)
	Local $lista = ""
	Switch $nazwa
		Case "PA"
			$lista = "Sprzeda¿ detaliczna"
		Case "FA"
			$lista = "Faktury zakupu"
		Case "MM"
			$lista = "Wydania magazynowe"
		Case "RW"
			$lista = "Wydania magazynowe"
		Case "PW"
			$lista = "Przyjêcia magazynowe"
		Case "ZW"
			$lista = "Zwroty detaliczne"
	EndSwitch
	Local $uchwyt = 0
	Local $hWnd = szukajUchwytu(WinGetHandle('[REGEXPTITLE:(?i)(.* - Subiekt GT)]'), $lista, "ins_hyperlink")
	$hWnd = _WinAPI_GetAncestor($hWnd)
	$hWnd = _WinAPI_GetAncestor($hWnd)
	$tmp = ControlGetPos($hWnd, "", "")
	$uchwyt = szukajUchwytu($hWnd, "", "GXWND")
	ControlClick($uchwyt, "", "", "left", 1, $tmp[0], 22)
	$pos = kopiujPozycje()
	If $pos <> "" Then
		Return $uchwyt
	EndIf
	$i = 0
	Do
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 1)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 2)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 3)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 4)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 5)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 6)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 7)
		$pos = kopiujPozycje()
		$i = $i + 8
	Until $pos <> "" Or $i >= $tmp[3]
	Return $uchwyt
EndFunc   ;==>fokusNaListeClick

;---------------------------------------------------------------------------------------------------------------------
;Z³apanie fokusa na liste dokumentow
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne

Func fokusNaListe($nazwa)
	Local $lista = ""
	Switch $nazwa
		Case "PA"
			$lista = "Sprzeda¿ detaliczna"
		Case "FA"
			$lista = "Faktury zakupu"
		Case "MM"
			$lista = "Wydania magazynowe"
		Case "RW"
			$lista = "Wydania magazynowe"
		Case "PW"
			$lista = "Przyjêcia magazynowe"
		Case "ZW"
			$lista = "Zwroty detaliczne"
	EndSwitch
	Local $uchwyt = 0

	Local $hWnd = szukajUchwytu(WinGetHandle('[REGEXPTITLE:(?i)(.* - Subiekt GT)]'), $lista, "ins_hyperlink")
	$hWnd = _WinAPI_GetAncestor($hWnd)
	$hWnd = _WinAPI_GetAncestor($hWnd)
	$uchwyt = szukajUchwytu($hWnd, "", "GXWND")
	$tmp = ControlGetPos($uchwyt, "", "")
	ControlFocus($uchwyt, "", "")

	Return $uchwyt
EndFunc   ;==>fokusNaListe

;---------------------------------------------------------------------------------------------------------------------
;£apanie uchwytu na konkretn¹ kontrolkê. Jeœli kontrolki siê powtarzaj¹ to zwróci uchwyt
;na pierwsz¹ z³apan¹

Func szukajUchwytu($hWnd, $text, $class)
	Local $uchwyt = 0
	Local $hChild = _WinAPI_GetWindow($hWnd, $GW_CHILD)
	While $hChild

		If _WinAPI_GetWindowText($hChild) = $text And _WinAPI_GetClassName($hChild) = $class Then
			$uchwyt = $hChild
			ExitLoop
		EndIf
		$uchwyt = szukajUchwytu($hChild, $text, $class)
		If $uchwyt <> 0 Then
			ExitLoop
		EndIf
		$hChild = _WinAPI_GetWindow($hChild, $GW_HWNDNEXT)
	WEnd
	Return $uchwyt
EndFunc   ;==>szukajUchwytu

;---------------------------------------------------------------------------------------------------------------------
;Wybór magazynu
;Przyjmuje skrócon¹ nazwe magazynu

Func magazyn($nazwa)
	$hWnd = WinGetHandle('[REGEXPTITLE:(?i)(.* - Subiekt GT)]')
	$uchwyt = pobierzUchwytKontrolkiOjca($hWnd, "", "#32770", "Magazyn:", "ins_combohyperlink")
	ControlFocus($uchwyt, "", "")
	ControlClick($uchwyt, "", "")
	Send("{HOME}")
	Switch $nazwa
		Case "TAR"
			downBtn()
			downBtn()
			downBtn()
			downBtn()
		Case "KUN"
			downBtn()
			downBtn()
		Case "MAG"
			downBtn()
			downBtn()
			downBtn()
		Case "ABR"
	EndSwitch
	Send("{ENTER}")
EndFunc   ;==>magazyn

;---------------------------------------------------------------------------------------------------------------------
;Ustawia filtr wyswietlaj¹cy dokumenty tylko z danego dnia
;Przyjmuje date któr¹ chcemy ustawiæ, nazwe listy na ktorej chcemy ustawiæ okres
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne

Func ustawOkres($data, $lista)

	$dataTab = StringReplace($data, '/', '-')
	$hWnd = WinGetHandle('[REGEXPTITLE:(?i)(.* - Subiekt GT)]')
	$uchwyt = 0
	Switch $lista
		Case "PA"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Sprzeda¿ detaliczna", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "FA"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Faktury zakupu", "ins_hyperlink", ", z okresu:", "Static")
		Case "MM"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Wydania magazynowe", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "RW"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Wydania magazynowe", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "PW"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Przyjêcia magazynowe", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "ZW"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Zwroty detaliczne", "ins_hyperlink", "Dokumenty z okresu:", "Static")
	EndSwitch

	$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	ControlFocus($uchwyt, "", "")
	ControlClick($uchwyt, "", "")
	Send("{END}{ENTER}")
	Do
		$hWnd = WinGetHandle('Dowolny okres')
		ClipPut($dataTab)
		$uchwyt = ControlGetHandle("", "", "ins_dateedit1")
		ControlFocus($uchwyt, "", "")
		Send("{CTRLDOWN}V")
		Send("{CTRLUP}")
		;Send($dataTab[2] & $dataTab[1] & $dataTab[0])
		$uchwyt = ControlGetHandle("", "", "ins_dateedit2")
		ControlFocus($uchwyt, "", "")
		;Send($dataTab[2] & $dataTab[1] & $dataTab[0])

		Send("{CTRLDOWN}V")
		Send("{CTRLUP}")
		Send("{ENTER}")
		Sleep(1000)
		$hWnd = WinGetHandle('Dowolny okres')
	Until $hWnd = 0
EndFunc   ;==>ustawOkres

;---------------------------------------------------------------------------------------------------------------------
;£apie fokus kontrolek do filtrowania dokumentów; typ dokumentu, okres itd
;Ogólnie ³apie kontrolke $child która wystêpuje po kontrolce $bro na tej samej ga³êzi (brat)
;Pobiera uchwyt okna, nazwe kontrolki brata, klase kontrolki brata
;nazwe kontrolki docelowej, klase kontrolki docelowej

Func pobierzUchwytKontrolkiListy($hWnd, $broNM, $broCL, $childNM, $childCL)
	Local $uchwyt = 0
	Local $hChild = _WinAPI_GetWindow($hWnd, $GW_CHILD)
	While $hChild

		If _WinAPI_GetWindowText($hChild) = $broNM And _WinAPI_GetClassName($hChild) = $broCL Then
			$hChild = _WinAPI_GetAncestor($hChild)

			$hChild = _WinAPI_GetWindow($hChild, $GW_HWNDNEXT)

			$hChild = _WinAPI_GetWindow($hChild, $GW_CHILD)

			While $hChild
				If _WinAPI_GetWindowText($hChild) = $childNM And _WinAPI_GetClassName($hChild) = $childCL Then
					$uchwyt = $hChild
					ExitLoop
				EndIf
				$hChild = _WinAPI_GetWindow($hChild, $GW_HWNDNEXT)
			WEnd
			ExitLoop
		EndIf
		$uchwyt = pobierzUchwytKontrolkiListy($hChild, $broNM, $broCL, $childNM, $childCL)
		If $uchwyt <> 0 Then
			ExitLoop
		EndIf
		$hChild = _WinAPI_GetWindow($hChild, $GW_HWNDNEXT)
	WEnd

	Return $uchwyt
EndFunc   ;==>pobierzUchwytKontrolkiListy

;---------------------------------------------------------------------------------------------------------------------
;£apie kontrolkê w ga³êzi danego rodzica
;przyjmuje uchwyt okna, nazwe kontrolki rodzica, klase kontrolki rodzica,
;nazwe kontrolki docelowej, klase kontrolki docelowej

Func pobierzUchwytKontrolkiOjca($hWnd, $parentNM, $parentCL, $childNM, $childCL)
	Local $uchwyt1
	Local $hChild = _WinAPI_GetWindow($hWnd, $GW_CHILD)
	While $hChild

		If _WinAPI_GetWindowText($hChild) = $parentNM And _WinAPI_GetClassName($hChild) = $parentCL Then
			$uchwyt1 = pobierzUchwytKontrolkiOjca($hChild, $parentNM, $parentCL, $childNM, $childCL)
			ExitLoop
		EndIf
		If _WinAPI_GetWindowText($hChild) = $childNM And _WinAPI_GetClassName($hChild) = $childCL Then
			$uchwyt1 = $hChild
			ExitLoop

		EndIf
		$uchwyt1 = pobierzUchwytKontrolkiOjca($hChild, $parentNM, $parentCL, $childNM, $childCL)
		If $uchwyt1 <> 0 Then
			ExitLoop
		EndIf
		$hChild = _WinAPI_GetWindow($hChild, $GW_HWNDNEXT)
	WEnd
	Return $uchwyt1
EndFunc   ;==>pobierzUchwytKontrolkiOjca


;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Zamyka program po wcisnieciu BREAK(ctrl+break)

Func zamknij()
	Exit
EndFunc   ;==>zamknij

;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Wstrzymuje dzia³anie programu

Func pauza()

	If $_PAUSE = False Then
		$_PAUSE = True
		logs("PAUZA........." & @LF)
		;pauzaLoop()
	Else
		$_PAUSE = False
		logs("START........." & @LF)
		ControlFocus($t, "", "")
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
		ControlFocus($t, "", "")
	EndIf
EndFunc   ;==>pauzaSzybka


;---------------------------------------------------------------------------------------------------------------------
;
Func kopiujPozycjeWiersz()
	ClipPut("MATEUSZBLASZCZAKP")
	$i = 0
	Do
		While ClipGet() = "MATEUSZBLASZCZAKP" And $i < 200
			Send("{CTRLDOWN}{SHIFTDOWN}Y")
			Send("{CTRLUP}{SHIFTUP}")
			Sleep(10)
			$i = $i + $i
		WEnd
	Until ClipGet() <> "MATEUSZBLASZCZAKP" Or $i >= 200
	If $i >= 200 Then
		logs("STOP")
		Return "STOP"
	EndIf
	Return ClipGet()
EndFunc   ;==>kopiujPozycjeWiersz

Func kopiujPozycje()
	ClipPut("MATEUSZBLASZCZAKP")
	Do
		While ClipGet() = "MATEUSZBLASZCZAKP"
			Send("{CTRLDOWN}{SHIFTDOWN}Y")
			Send("{CTRLUP}{SHIFTUP}")
			Sleep(10)
		WEnd
	Until ClipGet() <> "MATEUSZBLASZCZAKP"
	Return ClipGet()
EndFunc   ;==>kopiujPozycje


;---------------------------------------------------------------------------------------------------------------------
;

Func kopiujListe()
	ClipPut("MATEUSZBLASZCZAKL")
	Do
		While ClipGet() = "MATEUSZBLASZCZAKL"
			Send("{CTRLDOWN}{SHIFTDOWN}C")
			Send("{CTRLUP}{SHIFTUP}")
			Sleep(10)
		WEnd
	Until ClipGet() <> "MATEUSZBLASZCZAKL"
	Return ClipGet()
EndFunc   ;==>kopiujListe

;---------------------------------------------------------------------------------------------------------------------
;

Func homeBtn()
	Send("{CTRLDOWN}{HOME}")
	Send("{CTRLUP}")
EndFunc   ;==>homeBtn

;---------------------------------------------------------------------------------------------------------------------
;

Func endBtn()
	Send("{CTRLDOWN}{END}")
	Send("{CTRLUP}")
EndFunc   ;==>endBtn

;---------------------------------------------------------------------------------------------------------------------
;

Func downBtn()
	Send("{DOWN}")
EndFunc   ;==>downBtn


;---------------------------------------------------------------------------------------------------------------------
;

Func wyslijSms($tresc)
	Run('SmsEagle.exe "' & $tresc & '"', "", @SW_HIDE)
EndFunc   ;==>wyslijSms

;---------------------------------------------------------------------------------------------------------------------
;

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
Func przerwaNaBackup($podmiot)
	logs("")
	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ PRZERWA NA BACKUP $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
	logs("")
	$procTab = ProcessList("Subiekt.exe")
	$subiektPath = _WinAPI_GetProcessFileName($procTab[1][1])
	ProcessClose("Subiekt.exe")
	Do
		oknoArchiwizacja()
	Until oknoArchiwizacja() = 0 And uchwycSubiekta() = 0
	Do
		Sleep(1000)
	Until @HOUR = $godz2 And @MIN >= $minut2

	Run($subiektPath, "", @SW_SHOWMAXIMIZED)


	While wykryjEtapLogowania() = "brak"
		If WinGetHandle("[REGEXPTITLE:(?i)(Informacja o nowej wersji*)]") <> 0 Then
			$x = szukajUchwytu(WinGetHandle("[REGEXPTITLE:(?i)(Informacja o nowej wersji*)]"), "&Zamknij", "Button")
			ControlClick($x, "", "")
		EndIf
		Sleep(1000)
	WEnd

	$obsluzono = False
	While $obsluzono = False
		$obsluzono = obsluzEtapLogowania(wykryjEtapLogowania(), $podmiot)
		Sleep(2000)
	WEnd
	logs("")
	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ KONIEC PRZERWY NA BACKUP $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
	logs("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$")
	logs("")
EndFunc   ;==>przerwaNaBackup


;---------------------------------------------------------------------------------------------------------------------
;

Func obsluzEtapLogowania($etap, $podmiot)
	$hWnd = uchwycSubiekta()
	Switch $etap

		Case "podmiot"
			$uchwyt = fokusNaListePodmiotow()
			$pos = kopiujPozycje()
			While StringStripCR(StringStripWS($pos, $STR_STRIPALL)) <> $podmiot
				downBtn()
				$pos = kopiujPozycje()
			WEnd
			$uchwyt = szukajUchwytu($hWnd, "OK", "Button")
			ControlClick($uchwyt, "", "")
		Case "uzytkownik"
			$uchwyt = szukajUchwytu($hWnd, "", "ComboBox")
			ControlFocus($uchwyt, "", "")
			Send("M")
			$uchwyt = szukajUchwytu($hWnd, "OK", "Button")
			ControlClick($uchwyt, "", "")
		Case "zalogowano"
			If StringInStr(WinGetTitle($hWnd), $podmiot & " na serwerze") = 0 Then
				Send("{CTRLDOWN}{F4}")
				Send("{CTRLUP}")
			Else
				Return True
			EndIf
		Case Else
			restartujSubiekta("Subiekt GT")

	EndSwitch
	Return False
EndFunc   ;==>obsluzEtapLogowania

;-----------------------------------------------------------------------------------------------------------

Func uchwycSubiekta()
	$hWnd = WinGetHandle($t)
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

;-----------------------------------------------------------------------------------------------------------


Func restartujSubiekta($nazwa)
	$procTab = ProcessList("Subiekt.exe")
	$subiektPath = _WinAPI_GetProcessFileName($procTab[1][1])

	ProcessClose("Subiekt.exe")
	Do
		oknoArchiwizacja()
	Until oknoArchiwizacja() = 0 And (WinGetHandle($nazwa) = 0 Or uchwycSubiekta())
	Sleep(5000)
	Run($subiektPath)
	While wykryjEtapLogowania() = "brak"
		Sleep(1000)
	WEnd
EndFunc   ;==>restartujSubiekta

;-----------------------------------------------------------------------------------------------------------

Func wykryjEtapLogowania()

	$hWnd = uchwycSubiekta()
	$uchwyt = szukajUchwytu($hWnd, "Wybór podmiotu", "#32770")
	If $uchwyt <> 0 Then
		Return "podmiot"
	EndIf

	$uchwyt = szukajUchwytu($hWnd, "Wybierz u¿ytkownika, z którym chcesz rozpocz¹æ pracê", "Static")
	If $uchwyt <> 0 Then
		Return "uzytkownik"
	EndIf

	$uchwyt = szukajUchwytu($hWnd, "", "ins_svcwindow")
	If $uchwyt <> 0 Then
		Return "zalogowano"
	Else
		Return "brak"
	EndIf

EndFunc   ;==>wykryjEtapLogowania

;-----------------------------------------------------------------------------------------------------------

Func oknoArchiwizacja()
	$hSub = WinGetHandle("Subiekt GT")
	If $hSub <> 0 Then
		$uchwyt = szukajUchwytu($hSub, "&Nie", "Button")
		ControlClick($uchwyt, "", "")
	EndIf
	Return $hSub
EndFunc   ;==>oknoArchiwizacja

;---------------------------------------------------------------------------------------------------------------------
;Z³apanie fokusa na liste


Func fokusNaListePodmiotow()
	Local $uchwyt = 0
	Local $hWnd = uchwycSubiekta()
	$uchwyt = szukajUchwytu($hWnd, "", "GXWND")
	$tmp = ControlGetPos($uchwyt, "", "")
	ControlClick($uchwyt, "", "", "left", 1, $tmp[0], 22)
	$pos = kopiujPozycje()
	If $pos <> "" Then
		Return $uchwyt
	EndIf
	$i = 0
	Do
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 1)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 2)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 3)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 4)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 5)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 6)
		ControlClick($uchwyt, "", "", "left", 1, $tmp[0], $i + 7)
		$pos = kopiujPozycje()
		$i = $i + 8
	Until $pos <> "" Or $i >= $tmp[3]
	Return $uchwyt
EndFunc   ;==>fokusNaListePodmiotow
