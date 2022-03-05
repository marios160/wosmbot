;#####################################################################################################################
;#####################################################################################################################
;################################													##################################
;################################		Bot do ODK£ADANIA skutków magazynowuch		##################################
;################################				w programie Subiekt GT				##################################
;################################													##################################
;################################								Mateusz B³aszczak	##################################
;#####################################################################################################################
;#####################################################################################################################
;#####################################################################################################################
;#####################################################################################################################
;#####################################################################################################################
;#####################################################################################################################
;Przed w³¹czeniem bota nale¿y
;- na kazdej lisicie ustawic sotrowanie po dacie malej¹co, od najm³odszych dokumentów do najstarszych
;- na liscie zwrotow wlaczyc kolumne Dokument korygowany
;- na liscie brak towaru wlaczyc kolumne symbol
;#####################################################################################################################
;#####################################################################################################################
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


Func startOSMBot($dane,$dokumenty)
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

	$odp = MsgBox($MB_YESNO, "Czyszczenie loga", "Czy chcesz wyczyscic 'logOSM.log'?")
	If $odp == $IDYES Then
		$hF = FileOpen("logOSM.log", 2)
		FileClose($hF)
	EndIf


	;;-----------Inicjalizacja zmiennych
	Global $_PARAGONY[4]
	Global Enum $MAG = 0, $ABR = 1, $KUN = 2, $TAR = 3
	Global $_PAUSE = False
	Global Const $t = "[REGEXPTITLE:(?i)(.* - Subiekt GT)]"
	Global $_DATA = $dane[1]
	Global $_ENDDATA = $dane[2]
	Global $iloscLiniiBrakow = 0
	Global Const $HTTP_
	STATUS_OK = 200
	Global $_PODMIOT = $dane[0]
	Global $godz = 23
	Global $minut1 = 45
	Global $minut2 = 15

	$hWnd = WinGetHandle($t)
	ControlFocus($t, "", "")

	;;-----------Petla programowa,,,,,,,,,,
	logs("START BotWSM")
	logs("Pobieranie listy zwrotow")

	While $_DATA <> $_ENDDATA
	For $i = 0 to UBound($dokumenty) - 1
		If StringLen($dokumenty[$i]) = 3 Then
			magazyn($dokumenty[$i])
		Else
			wykonujListe($dokumenty[$i])
		EndIf
		Sleep(1000)
	Next
	WEnd



	magazyn("TAR")
	pobierzListeZwrotow($TAR)
	magazyn("KUN")
	pobierzListeZwrotow($KUN)
	magazyn("ABR")
	pobierzListeZwrotow($ABR)
	magazyn("MAG")
	pobierzListeZwrotow($MAG)
	logs("START PETLI")
	$licznikg = 0
	While $_DATA <> $_ENDDATA
		If $licznikg > 10 Then
			wyslijSms($_DATA & " - " & $iloscLiniiBrakow)
			$licznikg = 0
		EndIf
		$licznikg = $licznikg + 1
		logs("**************************************************************************************")
		logs("*************************************" & $_DATA & " **************************************")
		logs("**************************************************************************************")
		wykonujMagazyn1("TAR")
		;wykonujMagazyn1("KUN")
		wykonujMagazyn1("ABR")
		wykonujMagazyn1("MAG")

		;wykonujMagazyn2("TAR")
		wykonujMagazyn2("KUN")
		wykonujMagazyn2("ABR")
		wykonujMagazyn2("MAG")
		logs($_DATA & " - " & $iloscLiniiBrakow & @LF)
		$_DATA = _DateAdd('d', -1, $_DATA)
		If @HOUR = $godz And @MIN >= $minut1 Then
			przerwaNaBackup($_PODMIOT)
			$hWnd = WinGetHandle($t)
			ControlFocus($t, "", "")
		EndIf
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
;Funkcja zmienia magazyn na ktorym mamy dzia³aæ oraz uruchamia wywo³ywania skutkow na ka¿dej liscie
;Przyjmuje nazwê listy na ktorej dzia³amy
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne

Func wykonujMagazyn1($magazyn)
	magazyn($magazyn)
	logs("#################################### START 1 " & $magazyn & " ###########################################")

	Sleep(1000)
	wykonujListe("ZW")
	Sleep(1000)
	wykonujListe("PA")
	;Sleep(1000)
	;  wykonujListe("FA")
	;Sleep(1000)
	logs("#################################### KONIEC 1 " & $magazyn & " ###########################################")
EndFunc   ;==>wykonujMagazyn1


Func wykonujMagazyn2($magazyn)
	magazyn($magazyn)
	logs("#################################### START 2 " & $magazyn & " ###########################################")
	Switch $magazyn
		Case "MAG"
			Sleep(1000)
			wykonujListe("RW")
			Sleep(1000)
			wykonujListe("MM")
			Sleep(1000)
			wykonujListe("PW")
		Case "ABR"
			Sleep(1000)
			wykonujListe("RW")
			Sleep(1000)
			wykonujListe("MM")
			Sleep(1000)
			wykonujListe("PW")
		Case "KUN"
			Sleep(1000)
			wykonujListe("MM")
	EndSwitch
;~ Sleep(1000)
;~    wykonujListe("KFZ")
	;Sleep(1000)
	;  wykonujListe("FA")
	;Sleep(1000)
	logs("#################################### KONIEC 2 " & $magazyn & " ###########################################")
EndFunc   ;==>wykonujMagazyn2

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
			fokusNaListeClick($lista)
			homeBtn()
			endBtn()
			kopiujPozycje()
			$endpos = ClipGet()
			$curpos = ""
			homeBtn()
			While $endpos <> $curpos
				fokusNaListe($lista)
				$czyDaSieOdlozyc = False
				Do
					$curpos = kopiujPozycje()
				Until StringInStr($curpos, "skutek") <> 0
				;logs("MATI " & $curpos)
				If StringInStr($curpos, "wywo³uje skutek magazynowy") <> 0 Then
					While Not $czyDaSieOdlozyc
						$curpos2 = kopiujPozycje()
						;logs("ZTAG " & $czyDaSieOdlozyc & " " & $curpos2)
						If StringInStr($curpos2, "wywo³uje skutek magazynowy") <> 0 And Not $czyDaSieOdlozyc Then
							$czyDaSieOdlozyc = odlozSkutek($lista)
							If $lista = "FA" Then
								$czyDaSieOdlozyc = True
							EndIf
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
				ControlFocus($t, "", "")
				downBtn()
;~ 		 while $curpos = kopiujPozycje() And $curpos <> $endpos
;~ 			Sleep(1)
;~ 		 WEnd
			WEnd
			If StringInStr(kopiujListe(), "wywo³uje") <> 0 Then
				StringReplace(kopiujListe(), "wywo³uje skutek magazynowy", "xxx")
				logs("OMINIETO PRODUKT! " & "; " & " znaleziono: " & @extended)
			EndIf

		Until StringInStr(kopiujListe(), "wywo³uje") = 0
	EndIf
	logs("------------------------------- KONIEC " & $lista & " ----------------------------------------")

EndFunc   ;==>wykonujListe



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

Func odlozSkutek($lista)
	$okno = ""
	$paragon = ""
	;KOREKTY ZAKUPU
	$tmp = StringSplit(ClipGet(), "	", 1)
	Switch $lista
		Case "PA"
			Send("{LALT}OM")
			$okno = "Paragon"
			$paragon = $tmp[6]
		Case "FA"
			Send("{LALT}OM")
			$okno = "Faktura VAT zakupu"
			$paragon = $tmp[5]
		Case "KFZ"
			Send("{LALT}OM")
			$okno = "Korekta fakturu VAT zakupu"
			$paragon = $tmp[5]
		Case "MM"
			Send("{LALT}OM")
			$okno = "Przesuniêcie miêdzymagazynowe"
			$paragon = $tmp[4]
		Case "RW"
			Send("{LALT}OM")
			$okno = "Rozchód wewnêtrzny"
			$paragon = $tmp[4]
		Case "PW"
			Send("{LALT}OM")
			$okno = "Przychód wewnêtrzny"
			$paragon = $tmp[4]
		Case "ZW"
			Send("{LALT}OM")
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
	If $var Then
		logs("Nie odlozono skutku (brak towaru): " & $paragon)
	Else
		logs("Odlozono skutek: " & $paragon)
	EndIf
	Return $var
EndFunc   ;==>odlozSkutek

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
		Case "KFZ"
			$data = $tmp[3]
			$paragon = $tmp[5]
		Case "MM"
			$data = $tmp[3]
			$paragon = $tmp[4]
		Case "RW"
			$data = $tmp[3]
			$paragon = $tmp[4]
		Case "PW"
			$data = $tmp[3]
			$paragon = $tmp[4]
		Case "ZW"
			$data = $tmp[3]
			$paragon = $tmp[4]
	EndSwitch
	$hFile = FileOpen("OSMbraktowaru.txt", 1)
	FileWrite($hFile, $paragon & ";" & $data & @LF)
	$iloscLiniiBrakow = $iloscLiniiBrakow + 1
	kopiujListe()
	$tmp2 = StringSplit(ClipGet(), @LF, 1)
	For $i = 2 To UBound($tmp2) - 1
		FileWrite($hFile, $tmp2[$i])
		$iloscLiniiBrakow = $iloscLiniiBrakow + 1
	Next
	FileClose($hFile)
	logs("Zapisano brak towaru: " & $paragon)
	$uchwyt = szukajUchwytu($brak, "Zamknij", "Button")
	ControlClick($uchwyt, "", "")
EndFunc   ;==>zapiszBrakTowaru


Func logs($text)
	_FileWriteLog("logOSM.log", "[" & $_DATA & "] " & $text)
	ConsoleWrite("[" & $_DATA & "] " & $text & @LF)
	;Run('SmsEagle.exe "' & "[" & $_DATA & "] " & $text & @LF & '"', "", @SW_HIDE)
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
		fokusNaListeClick("ZW")
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
			$curpos = kopiujPozycje()
			If StringInStr($curpos, "wywo³uje skutek magazynowy") <> 0 Then
				If StringInStr($curpos, $paragon) <> 0 Then
					$i = 0
					Do
						odlozSkutek($lista)
						Sleep(1000)
						$i = $i + 1
					Until StringInStr(kopiujPozycje(), "wywo³uje skutek magazynowy") = 0 Or $i >= 3
					If $i >= 3 Then
						logs("Nie udalo sie odlozyc skutku zwrotu: " & $paragon)
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
		Case "KFZ"
			Send("{LALT}WZK")
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
	kopiujListe()
	$tmp = StringSplit(ClipGet(), @LF)
	If $tmp[0] < 2 Then Return False
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
		Case "KFZ"
			$lista = "Korekty zakupu"
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
		Case "KFZ"
			$lista = "Korekty zakupu"
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
		Case "KFZ"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Korekty zakupu", "ins_hyperlink", ", z okresu:", "Static")
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

Func kopiujPozycje()
	ControlFocus($t, "", "")
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
	ControlFocus($t, "", "")
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
	Until @HOUR = $godz And @MIN >= $minut2

	Run($subiektPath, "", @SW_SHOWMAXIMIZED)
	While wykryjEtapLogowania() = "brak"
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
