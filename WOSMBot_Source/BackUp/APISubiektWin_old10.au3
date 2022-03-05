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
#include <GUIConstantsEx.au3>
#include <GuiListBox.au3>
#include <GuiListView.au3>
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
#include <GuiMenu.au3>
#include <Misc.au3>


;-----LISTY SKROTY-----
Global $PA = "PA"
Global $PW = "PW"
Global $MM = "MM"
Global $ZW = "ZW"
Global $FA = "FA"
Global $RW = "RW"
Global $KFZ = "KFZ"
;------LISTY-----
Global $LST_SPRZEDAZ_DETALICZNA = "Sprzedaż detaliczna"
Global $LST_WYDANIA_MAGAZYNOWE = "Wydania magazynowe"
Global $LST_FAKTURY_ZAKUPU = "Faktury zakupu"
Global $LST_KOREKTY_ZAKUPU = "Korekty zakupu"
Global $LST_PRZYJECIA_MAGAZYNOWE = "Przyjęcia magazynowe"
Global $LST_ZWROTY_DETALICZNE = "Zwroty detaliczne"
;-----KLASY-----
Global $CLS_INS_HYPERLINK = "ins_hyperlink"
Global $CLS_INS_COMBOHYPERLINK = "ins_combohyperlink"
Global $CLS_INS_GRIDCNT = "ins_gridcnt"
Global $CLS_STATIC = "Static"
Global $CLS_BUTTON = "Button"
Global $CLS_AFXWND80 = "AfxWnd80"
Global $CLS_GXWND = "GXWND"
Global $CLS_32770 = "#32770"
;----FILTRY----
Global $FLT_DOKUMENTY_Z_OKRESU = "Dokumenty z okresu:"
Global $FLT_Z_OKRESU = ", z okresu:"
Global $FLT_TYPU = ", typu:"
Global $FLT_FILTR = "Filtr: "
Global $FLT_DOKUMENTY_WG = "Dokumenty wg:"
Global $FLT_O_STANIE_NP = "o stanie:"
Global $FLT_O_STANIE = ", o stanie:"
Global $FLT_O_KATEGORII = ", o kategorii:"
Global $FLT_O_KATEGORII_NS = ",o kategorii:"
Global $FLT_O_KATEGORII_NP = "o kategorii:"
Global $FLT_Z_FLAGA = ", z flagą:"
Global $FLT_Z_FLAGA_NS = ",z flagą:"
Global $FLT_RODZAJ_ZWROTU = ",rodzaj zwrotu:"
;----POZYCJE FILTROW----
Global $POZ_DOWOLNY_OKRES = "dowolny okres"
Global $POZ_BIEZACY_ROK = "bieżący rok"
Global $POZ_BIEZACY_DZIEN = "bieżący dzień"
Global $POZ_NIEOKRESLONY = "(nieokreślony)"
Global $POZ_WSZYSTKIE = "(wszystkie)"
Global $POZ_DOWOLNA = "(dowolna)"
Global $POZ_BRAK = "(brak)"
Global $POZ_DATY_WYSTAWIENIA = "daty wystawienia"
Global $POZ_DATY_OTRZYMANIA = "daty otrzymania"
Global $POZ_WYDANIE_ZEWNETRZNE = "wydanie zewnętrzne"
Global $POZ_ROZCHOD_WEWNETRZNY = "rozchód wewnętrzny"
Global $POZ_PRZESUNIECIE_MIEDZYMAGAZYNOWE = "przesunięcie międzymagazynowe"
Global $POZ_PRZYJECIE_ZEWNETRZNE = "przyjęcie zewnętrzne"
Global $POZ_PRZYCHOD_WEWNETRZNY = "przychód wewnętrzny"
;-------KONTROLKI----------
Global $CTR_MAGAZYN = "Magazyn:"
;--------MENU--------------
Global $MNU_PODMIOT = "&Podmiot"
Global $MNU_WIDOK = "&Widok"
Global $MNU_DODAJ = "&Dodaj"
Global $MNU_PARAGON = "P&aragon"
Global $MNU_OPERACJE = "&Operacje"
Global $MNU_NARZEDZIA = "&Narzędzia"
Global $MNU_POMOC = "Pomo&c"
;------SUBMENU------------
Global $SUB_ODLOZ_SKUTEK_MAGAZYNOWY = "Odłóż skutek &magazynowy"
Global $SUB_WYWOLAJ_SKUTEK_MAGAZYNOWY_W = "&Wywołaj skutek magazynowy"
Global $SUB_WYWOLAJ_SKUTEK_MAGAZYNOWY_S = "Wywołaj &skutek magazynowy"
Global $SUB_SPRZEDAZ = "&Sprzedaż"
Global $SUB_ZAKUP = "&Zakup"
Global $SUB_MAGAZYN = "&Magazyn"
Global $SUB_SPRZEDAZ_DETALICZNA = "Sprzedaż &detaliczna"
Global $SUB_WYDANIA_MAGAZYNOWE = "&Wydania magazynowe"
Global $SUB_FAKTURY_ZAKUPU = "&Faktury zakupu"
Global $SUB_PRZYJECIA_MAGAZYNOWE = "&Przyjęcia magazynowe"
Global $SUB_ZWROTY_DETALICZNE = "&Zwroty detaliczne"
;------HWND CTRL---------------
Global $HWND_OKRES = 0
Global $HWND_TYPU = 1
Global $HWND_KATEGORIA = 2
Global $HWND_FLAGA = 3
Global $HWND_STAN = 4
Global $HWND_WEDLUG = 5
Global $HWND_RODZAJ = 6
;------NAPISY------------------
Global $STR_WYWOLUJE_SKUTEK = "wywołuje skutek magazynowy"
Global $STR_ODLOZONY_SKUTEK = "odłożony skutek magazynowy"
Global $SUBIEKT = "Subiekt"
Global $_UchwytyKontrolek = ObjCreate("Scripting.Dictionary") ;lista z uchwytami do kontrolek

Global $_AKTUALNY_STAN[3]		;tablica z waznymi danymi ktore ciagle sie zmieniaja; magazyn,lista,data
Global $ServerAddress = ""
Global $ServerUserName = ""
Global $ServerPassword = ""
Global $DatabaseName = ""
Global $_RESTART = False
Global $_PAUZA = False
Global $_OWSkutek = ""		;uzywana do okreslania czy lista jest pusta (funkcja czyPusta() w APISubiektWin)
Global $_SciezkaUruchomieniaSubiekta = "" ;sciezka do programu subiekta ustawiana w funkcji sprawdzCzyJestSubiekt
Global $_nazwaOknaSubiekta = "" ;nazwa okna programu subiekta ustawiana w funkcji utworzPlikKonfiguracyjny
Global $GuiOczekiwanie = 0 ;uchwyt okienka oczekiwania

;---------------------------------------------------------------------------------------------------------------------
; wykrywanie komunikatu potwierdzenia zmiany skutku magazynowego
Func znajdzKomunikat()
	logs('@@ (132) :(' & @MIN & ':' & @SEC & ') znajdzKomunikat()' & @CR) ;### Function Trace
	$hWnd = WinGetHandle("[TITLE:Subiekt GT;CLASS:#32770]")
	$uchwyt = _WinAPI_GetWindow($hWnd, $GW_CHILD)
	While StringInStr(_WinAPI_GetWindowText($uchwyt), "Czy na pewno wykonać operację zmiany skutku magazynowego?") = 0 And $uchwyt <> 0
		$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	WEnd
	$uchwyt = _WinAPI_GetWindow($hWnd, $GW_CHILD)
	While StringInStr(_WinAPI_GetWindowText($uchwyt), "&Tak") = 0 And $uchwyt <> 0
		$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	WEnd

	ControlSend($uchwyt, "", "", "T")
	Return $hWnd
EndFunc   ;==>znajdzKomunikat

;---------------------------------------------------------------------------------------------------------------------
;Wywoluje skutek magazynowy aktualnie zaznaczonego rekordu
Func wywolajSkutekMagazynowy($lista)
	logs('@@ (150) :(' & @MIN & ':' & @SEC & ') wywolajSkutekMagazynowy()' & @CR) ;### Function Trace


	$nazwaOkna = ""
	$listaKopiuj = ""
	Switch ($lista)
		Case $LST_SPRZEDAZ_DETALICZNA
			$nazwaOkna = "Paragon"
			$listaKopiuj = $LST_SPRZEDAZ_DETALICZNA
		Case $LST_FAKTURY_ZAKUPU
			$nazwaOkna = "Faktura VAT zakupu"
			$listaKopiuj = $LST_FAKTURY_ZAKUPU
		Case $LST_ZWROTY_DETALICZNE
			$nazwaOkna = "Zwrot ze sprzedaży detalicznej"
			$listaKopiuj = $LST_ZWROTY_DETALICZNE
		Case $POZ_PRZESUNIECIE_MIEDZYMAGAZYNOWE & " (" & $LST_WYDANIA_MAGAZYNOWE & ")"
			$nazwaOkna = "Przesunięcie międzymagazynowe"
			$listaKopiuj = $LST_WYDANIA_MAGAZYNOWE
		Case $POZ_ROZCHOD_WEWNETRZNY & " (" & $LST_WYDANIA_MAGAZYNOWE & ")"
			$nazwaOkna = "Rozchód wewnętrzny"
			$listaKopiuj = $LST_WYDANIA_MAGAZYNOWE
		Case $POZ_PRZYCHOD_WEWNETRZNY & " (" & $LST_PRZYJECIA_MAGAZYNOWE & ")"
			$nazwaOkna = "Przychód wewnętrzny"
			$listaKopiuj = $LST_PRZYJECIA_MAGAZYNOWE
	EndSwitch
	$wiersz = False
	$kopiuj = ""
	Do
		$kopiuj = kopiujWierszSzybko($listaKopiuj)
	Until $kopiuj <> ""
	$j = 0
	Do
		$i = 0
		$var = False
		If $j > 3 Then	;jesli ponmad 3 razy nie udalo sie wywolac skutku to znaczy ze jest jakis blad
			If ControlFocus(uchwytSubiekta(), "", "") = 0 Then ; jesli nie mozemy zlapac fokusa to znaczy ze wyskoczyl nieznany komunikat w subiekcie
				MsgBox(0, "Koniec", "Koniec")
				Exit
			Else				;w przeciwnym razie oslablo polaczenie z baza i trzeba zrestartowac subiekta
				$j = 0
				$_RESTART = True
				Return False
			EndIf
		EndIf
		;zaleznie od listy inaczej jest wywolywanie skutku
		If $listaKopiuj = $LST_PRZYJECIA_MAGAZYNOWE Or $listaKopiuj = $LST_WYDANIA_MAGAZYNOWE Then
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_OPERACJE, $SUB_WYWOLAJ_SKUTEK_MAGAZYNOWY_S)
		Else
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_OPERACJE, $SUB_WYWOLAJ_SKUTEK_MAGAZYNOWY_W)
		EndIf
		Do	;w petli czekamy na rozne komunikaty
			$okno = WinGetHandle($nazwaOkna)	;szukamy komunikatu ze skutek jest wywolywany
			If $var Then				;jesli znaleziono wczesniej taki komunikat
				If $okno = 0 Then		;jesli juz zniknal
					$wiersz = True		;to znaczy ze wywolano skutek
					ExitLoop
				EndIf
			Else						;jesli nie znaleziono wczesniej komunikatu o wywolywaniu skutku
				If $okno <> 0 Then		; jesli teraz znaleziono
					$var = True			;to zapisujemy ze znaleziono
				EndIf
			EndIf
			znajdzKomunikat()			;wykrywanie komunikatu o potwierdzeniu wywolania jesli sie pojawi
			$brakTowaru = WinGetHandle("[REGEXPTITLE:(?i)(.* - brak towaru w magazynie)]")		;wykrywanie braku towaru
			If $brakTowaru <> 0 Then
				zapiszBrakTowaru($lista)	;jak brak towaru to trzeba zapisac do pliku
			EndIf
			$i = $i + 1
		Until $i >= 500 Or StringInStr(kopiujWierszSzybko($listaKopiuj), $STR_WYWOLUJE_SKUTEK) > 0
		$j = $j + 1
	Until $wiersz Or StringInStr(kopiujWierszSzybko($listaKopiuj), $STR_WYWOLUJE_SKUTEK) > 0
	Return True
EndFunc   ;==>wywolajSkutekMagazynowy

;---------------------------------------------------------------------------------------------------------------------
;Odklada skutek magazynowy aktualnie zaznaczonego rekordu
;Wszystko odbywa sie analogicznie jak przy wywolywaniu skutku oprocz zapisywania brakujacego towaru
Func odlozSkutekMagazynowy($lista)
	logs('@@ (222) :(' & @MIN & ':' & @SEC & ') odlozSkutekMagazynowy()' & @CR) ;### Function Trace
	$nazwaOkna = ""
	$listaKopiuj = ""
	Switch ($lista)
		Case $LST_SPRZEDAZ_DETALICZNA
			$nazwaOkna = "Paragon"
			$listaKopiuj = $LST_SPRZEDAZ_DETALICZNA
		Case $LST_FAKTURY_ZAKUPU
			$nazwaOkna = "Faktura VAT zakupu"
			$listaKopiuj = $LST_FAKTURY_ZAKUPU
		Case $LST_ZWROTY_DETALICZNE
			$nazwaOkna = "Zwrot ze sprzedaży detalicznej"
			$listaKopiuj = $LST_ZWROTY_DETALICZNE
		Case $POZ_PRZESUNIECIE_MIEDZYMAGAZYNOWE & " (" & $LST_WYDANIA_MAGAZYNOWE & ")"
			$nazwaOkna = "Przesunięcie międzymagazynowe"
			$listaKopiuj = $LST_WYDANIA_MAGAZYNOWE
		Case $POZ_ROZCHOD_WEWNETRZNY & " (" & $LST_WYDANIA_MAGAZYNOWE & ")"
			$nazwaOkna = "Rozchód wewnętrzny"
			$listaKopiuj = $LST_WYDANIA_MAGAZYNOWE
		Case $POZ_PRZYCHOD_WEWNETRZNY & " (" & $LST_PRZYJECIA_MAGAZYNOWE & ")"
			$nazwaOkna = "Przychód wewnętrzny"
			$listaKopiuj = $LST_PRZYJECIA_MAGAZYNOWE
	EndSwitch
	$wiersz = False
	$j = 0
	Do
		$i = 0
		$var = False
		If $j > 3 Then
			If ControlFocus(uchwytSubiekta(), "", "") = 0 Then
				MsgBox(0, "Koniec", "Koniec")
				Exit
			Else
				$j = 0
				$_RESTART = True
				Return False
			EndIf
		EndIf
		WinMenuSelectItem(uchwytSubiekta(), "", $MNU_OPERACJE, $SUB_ODLOZ_SKUTEK_MAGAZYNOWY)
		Do
			$okno = WinGetHandle($nazwaOkna)
			logs($nazwaOkna)
			If $var Then
				If $okno = 0 Then
					$wiersz = True
					ExitLoop
				EndIf
			Else
				If $okno <> 0 Then
					$var = True
				EndIf
			EndIf
			znajdzKomunikat()
			$i = $i + 1
		Until $i >= 500 Or StringInStr(kopiujWierszSzybko($listaKopiuj), $STR_ODLOZONY_SKUTEK) > 0
		$j = $j + 1
	Until $wiersz Or StringInStr(kopiujWierszSzybko($listaKopiuj), $STR_ODLOZONY_SKUTEK) > 0
	;logs(kopiujWierszSzybko($listaKopiuj))
	Return True
EndFunc   ;==>odlozSkutekMagazynowy


;---------------------------------------------------------------------------------------------------------------------
;Ustawia filtry listy na domyslne
Func wyczyscFiltry($lista)
	logs('@@ (285) :(' & @MIN & ':' & @SEC & ') wyczyscFiltry()' & @CR) ;### Function Trace
	For $key In $_UchwytyKontrolek.Item($lista)
		ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($lista).Item($key)), "SetCurrentSelection", 0)
	Next
EndFunc   ;==>wyczyscFiltry

;---------------------------------------------------------------------------------------------------------------------
; pobiera numer paragonu ze stringa

Func getNumer($rekord)
	logs('@@ (295) :(' & @MIN & ':' & @SEC & ') getNumer()' & @CR) ;### Function Trace
	$tmp1 = StringSplit($rekord, "	", 2)
	Return $tmp1[2]
EndFunc   ;==>getNumer

;---------------------------------------------------------------------------------------------------------------------
; zapisujemy logi

Func logs($text)
	;dodajStatus($text)
	_FileWriteLog("logWOSM.log", $text)
	ConsoleWrite($text & @LF)
EndFunc   ;==>logs

;---------------------------------------------------------------------------------------------------------------------
;Uruchamia liste dokumentow uzywajac skrotu nazwy
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne
Func wybierzListeNN($lista)
	logs('@@ (315) :(' & @MIN & ':' & @SEC & ') wybierzListeNN()' & @CR) ;### Function Trace
	Switch ($lista)
		Case $PA
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_SPRZEDAZ_DETALICZNA)
			czekajNaZaladowanieListy($lista)
		Case $FA
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_ZAKUP, $SUB_FAKTURY_ZAKUPU)
			czekajNaZaladowanieListy($lista)
		Case $MM
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			czekajNaZaladowanieListy($lista)
			wybierzFiltr($LST_WYDANIA_MAGAZYNOWE, $HWND_TYPU, $POZ_PRZESUNIECIE_MIEDZYMAGAZYNOWE)
		Case $RW
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			czekajNaZaladowanieListy($lista)
			wybierzFiltr($LST_WYDANIA_MAGAZYNOWE, $HWND_TYPU, $POZ_ROZCHOD_WEWNETRZNY)
		Case $PW
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_PRZYJECIA_MAGAZYNOWE)
			czekajNaZaladowanieListy($lista)
			wybierzFiltr($LST_PRZYJECIA_MAGAZYNOWE, $HWND_TYPU, $POZ_PRZYCHOD_WEWNETRZNY)
		Case $ZW
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_ZWROTY_DETALICZNE)
			czekajNaZaladowanieListy($lista)
	EndSwitch
	czekajNaZaladowanieListy($lista)
EndFunc   ;==>wybierzListeNN


;---------------------------------------------------------------------------------------------------------------------
;Uruchamia liste dokumentow; ustawia filtry jesli podamy je w liscie key=nazwa filtra, item=pozycja(nazwa)
Func wybierzListe($lista, $filtry = ObjCreate("Scripting.Dictionary"))
	logs('@@ (346) :(' & @MIN & ':' & @SEC & ') wybierzListe()' & @CR) ;### Function Trace
	$listaBezFiltra = $lista
	Switch ($lista)
		Case $LST_SPRZEDAZ_DETALICZNA
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_SPRZEDAZ_DETALICZNA)
			czekajNaZaladowanieListy($lista)
		Case $LST_FAKTURY_ZAKUPU
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_ZAKUP, $SUB_FAKTURY_ZAKUPU)
			czekajNaZaladowanieListy($lista)
		Case $LST_WYDANIA_MAGAZYNOWE
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			czekajNaZaladowanieListy($lista)
		Case $LST_PRZYJECIA_MAGAZYNOWE
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_PRZYJECIA_MAGAZYNOWE)
			czekajNaZaladowanieListy($lista)
		Case $LST_ZWROTY_DETALICZNE
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_ZWROTY_DETALICZNE)
			czekajNaZaladowanieListy($lista)
		Case $POZ_PRZESUNIECIE_MIEDZYMAGAZYNOWE & " (" & $LST_WYDANIA_MAGAZYNOWE & ")"
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			czekajNaZaladowanieListy($LST_WYDANIA_MAGAZYNOWE)
			wybierzFiltr($LST_WYDANIA_MAGAZYNOWE, $HWND_TYPU, $POZ_PRZESUNIECIE_MIEDZYMAGAZYNOWE)
			$listaBezFiltra = $LST_WYDANIA_MAGAZYNOWE
		Case $POZ_ROZCHOD_WEWNETRZNY & " (" & $LST_WYDANIA_MAGAZYNOWE & ")"
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			czekajNaZaladowanieListy($LST_WYDANIA_MAGAZYNOWE)
			wybierzFiltr($LST_WYDANIA_MAGAZYNOWE, $HWND_TYPU, $POZ_ROZCHOD_WEWNETRZNY)
			$listaBezFiltra = $LST_WYDANIA_MAGAZYNOWE
		Case $POZ_PRZYCHOD_WEWNETRZNY & " (" & $LST_PRZYJECIA_MAGAZYNOWE & ")"
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_PRZYJECIA_MAGAZYNOWE)
			czekajNaZaladowanieListy($LST_PRZYJECIA_MAGAZYNOWE)
			wybierzFiltr($LST_PRZYJECIA_MAGAZYNOWE, $HWND_TYPU, $POZ_PRZYCHOD_WEWNETRZNY)
			$listaBezFiltra = $LST_PRZYJECIA_MAGAZYNOWE
	EndSwitch
	If $filtry.count > 0 Then
		For $key In $filtry
			wybierzFiltr($listaBezFiltra, $key, $filtry.Item($key))
		Next
	EndIf
	czekajNaZaladowanieListy($listaBezFiltra)
	Return $listaBezFiltra
EndFunc   ;==>wybierzListe


;---------------------------------------------------------------------------------------------------------------------
;Lapanie uchwytu na konkretna kontrolke. Jesli kontrolki sie powtarzaja to zwróci uchwyt
;na pierwsza zlapana
Func szukajUchwytu($hWnd, $text, $class)
	;logs('@@ (394) :(' & @MIN & ':' & @SEC & ') szukajUchwytu()' & @CR) ;### Function Trace
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
;Lapanie uchwytu glownego okna subiekta
Func uchwytSubiekta()
	logs('@@ (415) :(' & @MIN & ':' & @SEC & ') uchwytSubiekta()' & @CR) ;### Function Trace
	Return HWnd($_UchwytyKontrolek.Item($SUBIEKT))
EndFunc   ;==>uchwytSubiekta


;---------------------------------------------------------------------------------------------------------------------
;Ustawia filtr wyswietlajacy dokumenty tylko z danego dnia
;Przyjmuje date która chcemy ustawic, nazwe listy na ktorej chcemy ustawic okres
Func ustawOkres($data, $lista)
	logs('@@ (424) :(' & @MIN & ':' & @SEC & ') ustawOkres()' & @CR) ;### Function Trace
	dodajStatus("Ustawianie daty: " & $data)
	$dataTab = StringSplit($data, '-', 2)
	ControlFocus(uchwytSubiekta(), "", "")
	Do
		Run(@AutoItExe & ' /AutoIt3ExecuteLine "ControlCommand ( '''','''', hwnd(' & $_UchwytyKontrolek.Item($lista).Item($HWND_OKRES) & '), ''SelectString'', ''' & $POZ_DOWOLNY_OKRES & ''' )"')
		$i = 0
		Do
			$hWnd = WinGetHandle('Dowolny okres')
			Sleep(10)
			$i = $i + 1
		Until $hWnd <> 0 Or $i > 20
	Until $hWnd <> 0
	Do
		$uchwyt = szukajUchwytu($hWnd, "Data &początkowa:", $CLS_STATIC)
		$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
		ControlSend($uchwyt, "", "", $dataTab[2] & $dataTab[1] & $dataTab[0])
		$uchwyt = szukajUchwytu($hWnd, "Data &końcowa:", $CLS_STATIC)
		$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
		ControlSend($uchwyt, "", "", $dataTab[2] & $dataTab[1] & $dataTab[0])
		$uchwyt = szukajUchwytu($hWnd, "OK", $CLS_BUTTON)
		ControlClick($uchwyt, "", "")
	Until WinGetHandle('Dowolny okres') = 0
EndFunc   ;==>ustawOkres

;---------------------------------------------------------------------------------------------------------------------
;Szuka uchwytu dopoki nie znajdzie
Func pobierzUchwytKontrolkiListyPetla($hWnd, $broNM, $broCL, $childNM, $childCL)
	$wynik = 0
	While $wynik = 0
		$wynik = pobierzUchwytKontrolkiListy($hWnd, $broNM, $broCL, $childNM, $childCL)
	WEnd
	Return $wynik
EndFunc   ;==>pobierzUchwytKontrolkiListyPetla

;---------------------------------------------------------------------------------------------------------------------
;Lapie fokus kontrolek do filtrowania dokumentów; typ dokumentu, okres itd
;Ogólnie lapie kontrolke $child która wystepuje po kontrolce $bro na tej samej galezi (brat)
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
;Sprawdza czy lista jest pusta
Func czyPusta($lista)
	logs('@@ (497) :(' & @MIN & ':' & @SEC & ') czyPusta()' & @CR) ;### Function Trace
	Do
		$x = kopiujListe($lista)
	Until StringLen($x) > 0
	If StringInStr($x, $_OWSkutek) = 0 Then
		logs("Wszystko odlozone")
		Return True
	EndIf
	$tmp = StringSplit(ClipGet(), @LF)
	If $tmp[0] < 2 Then Return False
	If StringLen($tmp[2]) = 0 Then
		logs("PUSTA")
		Return True
	Else
		Return False
	EndIf
	Return False
EndFunc   ;==>czyPusta


;-----------------------------------------------------------------------------------------------------------------------------
;Sprawdza na ktorym etapie uruchamiania jest subiekt
Func czySubiektUruchomiony($_nazwaOknaSubiekta)
	logs('@@ (520) :(' & @MIN & ':' & @SEC & ') czySubiektUruchomiony()' & @CR) ;### Function Trace
	$hWnd = zlapSubiekta($_nazwaOknaSubiekta)
	$uchwyt = 0
	If $uchwyt <> 0 Then
		Return 4 ; uzytkownik
	EndIf
	$uchwyt = szukajUchwytu($hWnd, "Wybór serwera, użytkownika i hasła", $CLS_32770)
	If $uchwyt <> 0 Then
		Return 1 ; serwer
	EndIf
	$uchwyt = szukajUchwytu($hWnd, "Wybór podmiotu", $CLS_32770)
	If $uchwyt <> 0 Then
		Return 2 ; podmiot
	EndIf
	$uchwyt = szukajUchwytu($hWnd, "Wybierz użytkownika, z którym chcesz rozpocząć pracę", $CLS_STATIC)
	If $uchwyt <> 0 Then
		Return 3 ; uzytkownik
	EndIf
	$uchwyt = szukajUchwytu($hWnd, "Trwa logowanie, proszę czekać", $CLS_32770)
	If $uchwyt <> 0 Then
		Return 4 ; uzytkownik
	EndIf
	$uchwyt = szukajUchwytu($hWnd, "", "ins_svcwindow")
	If $uchwyt <> 0 Then
		$hWnd = WinGetHandle('[REGEXPTITLE:(?i)(Informacja o nowej wersji: *)]')
		If $hWnd <> 0 Then
			$uchwyt = szukajUchwytu($hWnd, "&Zamknij", $CLS_BUTTON)
			ControlClick($uchwyt, "", "")
		EndIf
		Return 0 ;subiekt
	Else
		Return -1 ;brak
	EndIf
EndFunc   ;==>czySubiektUruchomiony


;-----------------------------------------------------------------------------------------------------------
;lapie okno subiekta, przed zalogowaniem jest inna nazwa okna
Func zlapSubiekta($_nazwaOknaSubiekta)
	logs('@@ (559) :(' & @MIN & ':' & @SEC & ') zlapSubiekta()' & @CR) ;### Function Trace
	$hWnd = WinGetHandle($_nazwaOknaSubiekta)
	$uchwyt = szukajUchwytu($hWnd, "Magazyn:", $CLS_INS_COMBOHYPERLINK)
	If $uchwyt <> 0 Then
		Return $hWnd
	EndIf
	$hWnd = WinGetHandle("Subiekt GT")
	$uchwyt = szukajUchwytu($hWnd, "Magazyn:", $CLS_INS_COMBOHYPERLINK)
	If $uchwyt <> 0 Then
		Return $hWnd
	EndIf
	Return 0
EndFunc   ;==>zlapSubiekta

;---------------------------------------------------------------------------------------------------------------------
;ustawia podany filtr
Func wybierzFiltr($lista, $filtr, $pozycja = "")
	logs('@@ (576) :(' & @MIN & ':' & @SEC & ') wybierzFiltr()' & @CR) ;### Function Trace
	If $pozycja = "" Then
		ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($lista).Item($filtr)), "SetCurrentSelection", $pozycja)
	Else
		ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($lista).Item($filtr)), "SelectString", $pozycja)
	EndIf
EndFunc   ;==>wybierzFiltr

;---------------------------------------------------------------------------------------------------------------------
;ustawia podany magazyn
Func wybierzMagazyn($magazyn)
	logs('@@ (587) :(' & @MIN & ':' & @SEC & ') wybierzMagazyn()' & @CR) ;### Function Trace
	dodajStatus("Wybieranie magazynu: " & $magazyn)
	ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($CTR_MAGAZYN)), "SelectString", $magazyn)
EndFunc   ;==>wybierzMagazyn

;---------------------------------------------------------------------------------------------------------------------
;pobiera liste magazynow z subiekta
Func pobierzListeMagazynow()
	$uchwyt = HWnd($_UchwytyKontrolek.Item($CTR_MAGAZYN))
	$count = _GUICtrlListBox_GetCount($uchwyt) - 2
	Local $magazyn[$count]
	For $i = 0 To $count - 1
		$magazyn[$i] = _GUICtrlListBox_GetText($uchwyt, $i)
	Next
	Return $magazyn
EndFunc   ;==>pobierzListeMagazynow


;---------------------------------------------------------------------------------------------------------------------
;pobiera uchwyt listy wyboru magazynow
Func pobierzUchwytMagazynow()
	$i = 0
	While $i < 10
		$uchwyt = szukajUchwytu(uchwytSubiekta(), $CTR_MAGAZYN, $CLS_INS_COMBOHYPERLINK)
		ControlFocus($uchwyt, "", "")
		ControlClick($uchwyt, "", "")
		$uchwyt = WinGetHandle("[REGEXPCLASS:(Afx:.*.:800)]")
		$uchwyt2 = _WinAPI_GetWindow($uchwyt, $GW_CHILD)
		If _WinAPI_GetClassName($uchwyt2) = "ListBox" And _WinAPI_GetDlgCtrlID($uchwyt2) = 2001 Then
			Send("{ESC}")
			Return $uchwyt2
		EndIf
		Sleep(1000)
	WEnd
	; Jesli cos sie zle zaladowalo to trzeba jeszcze raz uruchomic bota
	MsgBox($MB_TOPMOST, "Błąd", "Wystąpił błąd przy pobieraniu uchwytów" & @LF & "Uruchom program WOSMBot.exe od nowa!")
	Exit
	Return 0
EndFunc   ;==>pobierzUchwytMagazynow

;---------------------------------------------------------------------------------------------------------------------
;usuwa separatory z filtrow zeby dalo sie je wybrac po uchwycie
Func usunSeparatory($uchwyt)
	$count = _GUICtrlListBox_GetCount($uchwyt)
	For $i = 0 To $count - 1
		If _GUICtrlListBox_GetText($uchwyt, $i) = "" Then
			ControlCommand("", "", $uchwyt, "DelString", $i)
		EndIf
	Next
EndFunc   ;==>usunSeparatory


;---------------------------------------------------------------------------------------------------------------------
;pobiera uchwyt do filtra
Func zapiszUchwytFiltry($lista, $filtr)
	$i = 0
	While $i < 10
		otworzFiltr($lista, $filtr)
		$uchwyt = WinGetHandle("[REGEXPCLASS:(Afx:.*.:800)]")
		$uchwyt2 = _WinAPI_GetWindow($uchwyt, $GW_CHILD)
		If _WinAPI_GetClassName($uchwyt2) = "ListBox" And _WinAPI_GetDlgCtrlID($uchwyt2) = 2001 Then
			usunSeparatory($uchwyt2)
			ControlCommand("", "", $uchwyt2, "SelectString", $POZ_BIEZACY_DZIEN)
			Send("{ESC}")
			Return $uchwyt2
		EndIf
		Sleep(1000)
	WEnd
	MsgBox($MB_TOPMOST, "Błąd", "Wystąpił błąd przy pobieraniu uchwytów" & @LF & "Uruchom program WOSMBot.exe od nowa!")
	Exit
	Return 0
EndFunc   ;==>zapiszUchwytFiltry

;---------------------------------------------------------------------------------------------------------------------
;pobiera uchwyt do listy dokumentow
Func zapiszUchwytGXWND($lista)
	$i = 0
	While $i < 10
		$uchwyt = szukajUchwytu(uchwytSubiekta(), $lista, $CLS_INS_HYPERLINK)
		$uchwyt = _WinAPI_GetAncestor($uchwyt, $GA_PARENT)
		While _WinAPI_GetClassName($uchwyt) <> $CLS_INS_GRIDCNT And $uchwyt <> 0
			$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
		WEnd
		$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_CHILD)
		If _WinAPI_GetClassName($uchwyt) = $CLS_GXWND Then
			Return $uchwyt
		EndIf
		Sleep(1000)
	WEnd
	MsgBox($MB_TOPMOST, "Błąd", "Wystąpił błąd przy pobieraniu uchwytów" & @LF & "Uruchom program WOSMBot.exe od nowa!")
	Exit
	Return 0
EndFunc   ;==>zapiszUchwytGXWND


;---------------------------------------------------------------------------------------------------------------------
;otwiera liste filtra
Func otworzFiltr($lista, $filtr)
	$hWnd = uchwytSubiekta()
	$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $lista, $CLS_INS_HYPERLINK, $filtr, $CLS_STATIC)
	$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	ControlFocus($uchwyt, "", "")
	ControlClick($uchwyt, "", "")
	Return $uchwyt
EndFunc   ;==>otworzFiltr

;---------------------------------------------------------------------------------------------------------------------
;wciskanie przyciskow, skrotow na liscie dokumentow
Func sendGXWNDDown($lista)
	logs('@@ (682) :(' & @MIN & ':' & @SEC & ') sendGXWNDDown()' & @CR) ;### Function Trace
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{DOWN}")
EndFunc   ;==>sendGXWNDDown
Func sendGXWNDUp($lista)
	logs('@@ (686) :(' & @MIN & ':' & @SEC & ') sendGXWNDUp()' & @CR) ;### Function Trace
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{UP}")
EndFunc   ;==>sendGXWNDUp
Func sendGXWNDHomeHard($lista)
	logs('@@ (690) :(' & @MIN & ':' & @SEC & ') sendGXWNDHomeHard()' & @CR) ;### Function Trace
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{CTRLDOWN}{HOME}{CTRLUP}")
EndFunc   ;==>sendGXWNDHomeHard
Func sendGXWNDEndHard($lista)
	logs('@@ (694) :(' & @MIN & ':' & @SEC & ') sendGXWNDEndHard()' & @CR) ;### Function Trace
	Send("{SHIFTUP}")
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "^{END}")
EndFunc   ;==>sendGXWNDEndHard
Func sendGXWNDHome($lista)
	logs('@@ (699) :(' & @MIN & ':' & @SEC & ') sendGXWNDHome()' & @CR) ;### Function Trace
	Send("{SHIFTUP}")
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{HOME}")
EndFunc   ;==>sendGXWNDHome
Func sendGXWNDEnd($lista)
	logs('@@ (704) :(' & @MIN & ':' & @SEC & ') sendGXWNDEnd()' & @CR) ;### Function Trace
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{END}")
EndFunc   ;==>sendGXWNDEnd
Func sendGXWNDEnter($lista)
	logs('@@ (708) :(' & @MIN & ':' & @SEC & ') sendGXWNDEnter()' & @CR) ;### Function Trace
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{ENTER}")
EndFunc   ;==>sendGXWNDEnter
Func sendGXWNDEsc($lista)
	logs('@@ (712) :(' & @MIN & ':' & @SEC & ') sendGXWNDEsc()' & @CR) ;### Function Trace
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{ESC}")
EndFunc   ;==>sendGXWNDEsc

;---------------------------------------------------------------------------------------------------------------------
;kopiuje tresc rekordu
Func kopiujWiersz($lista)
	logs('@@ (719) :(' & @MIN & ':' & @SEC & ') kopiujWiersz()' & @CR) ;### Function Trace
	$j = 0
	ControlFocus(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "")
	ClipPut("XEDEX_W")
	While ClipGet() = "XEDEX_W" And $j < 500
		$i = 0
		Send("{CTRLDOWN}{SHIFTDOWN}Y{SHIFTUP}{CTRLUP}")
		While ClipGet() = "" And $i < 10
			Sleep(1)
			$i = $i + 1
		WEnd
		$j = $j + 1
	WEnd
	;trzeba odcisnac shift lub control
	While _IsPressed(10)
		Sleep(10)
		logs('>Error code: ' & @error & @CRLF & @CRLF & '@@ Trace(715) :    		Send("{SHIFTUP}")' & @CRLF) ;### Trace Console
		Send("{SHIFTUP}")
	WEnd
	While _IsPressed(11)
		Sleep(10)
		logs('>Error code: ' & @error & @CRLF & @CRLF & '@@ Trace(715) :    		Send("{CTRLUP}")' & @CRLF) ;### Trace Console
		Send("{CTRLUP}")
	WEnd
	Return ClipGet()
EndFunc   ;==>kopiujWiersz

;---------------------------------------------------------------------------------------------------------------------
;kopiuje tresc rekordu
Func kopiujWierszSzybko($lista)
	logs('@@ (745) :(' & @MIN & ':' & @SEC & ') kopiujWierszSzybko()' & @CR) ;### Function Trace
	ControlFocus(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "")
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{CTRLDOWN}{SHIFTDOWN}Y{SHIFTUP}{CTRLUP}")
	While _IsPressed(10)
		Sleep(10)
		logs('>Error code: ' & @error & @CRLF & @CRLF & '@@ Trace(715) :    		Send("{SHIFTUP}")' & @CRLF) ;### Trace Console
		Send("{SHIFTUP}")
	WEnd
	While _IsPressed(11)
		Sleep(10)
		logs('>Error code: ' & @error & @CRLF & @CRLF & '@@ Trace(715) :    		Send("{CTRLUP}")' & @CRLF) ;### Trace Console
		Send("{CTRLUP}")
	WEnd
	Return ClipGet()
EndFunc   ;==>kopiujWierszSzybko

;---------------------------------------------------------------------------------------------------------------------
;kopiuje cala liste
Func kopiujListe($lista)
	logs('@@ (761) :(' & @MIN & ':' & @SEC & ') kopiujListe()' & @CR) ;### Function Trace
	ClipPut("XEDEX_L")
	While ClipGet() = "XEDEX_L"
		$i = 0
		ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "^+{C}")
		While ClipGet() = "" And $i < 10
			Sleep(1)
			$i = $i + 1
		WEnd
	WEnd
	Return ClipGet()
EndFunc   ;==>kopiujListe

;---------------------------------------------------------------------------------------------------------------------
; czekamy az lista sie cala zaladuje
Func czekajNaZaladowanieListy($lista)
	logs('@@ (777) :(' & @MIN & ':' & @SEC & ') czekajNaZaladowanieListy()' & @CR) ;### Function Trace

	ClipPut("XEDEX_Z")
	While ClipGet() = "XEDEX_Z"
		$uchwyt = zapiszUchwytGXWND($lista)
		ControlSend($uchwyt, "", "", "^+{Y}")
		; czasem moze wyskoczyc komunikat o aktualizacji
		komunikatAktualizacji()
	WEnd
	Return ClipGet()
EndFunc   ;==>czekajNaZaladowanieListy

;---------------------------------------------------------------------------------------------------------------------
;fokus na liste
Func fokusNaGXWND($lista)
	logs('@@ (791) :(' & @MIN & ':' & @SEC & ') fokusNaGXWND()' & @CR) ;### Function Trace
	$hWnd = 0
	While $hWnd = 0
		ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{CTRLDOWN}{SHIFTDOWN}k{CTRLUP}{SHIFTUP}")
		$hWnd = WinGetHandle("Lista kolumn")
	WEnd
	$uchwyt = 0
	Do
		$uchwyt = szukajUchwytu($hWnd, "OK", $CLS_BUTTON)
		ControlClick($uchwyt, "", "")
	Until WinGetHandle("Lista kolumn") = 0
EndFunc   ;==>fokusNaGXWND

;---------------------------------------------------------------------------------------------------------------------
; zamykanie komunikatu o aktualizacji
Func komunikatAktualizacji()
	logs('@@ (807) :(' & @MIN & ':' & @SEC & ') komunikatAktualizacji()' & @CR) ;### Function Trace
	$hWnd = WinGetHandle('[REGEXPTITLE:(?i)(Informacja o nowej wersji: *)]')
	If $hWnd <> 0 Then
		$uchwyt = szukajUchwytu($hWnd, "&Zamknij", $CLS_BUTTON)
		ControlClick($uchwyt, "", "")
	EndIf
EndFunc   ;==>komunikatAktualizacji

;---------------------------------------------------------------------------------------------------------------------
; trzeba ustawic kolumny w subiekcie recznie

Func ustawKolumnyListy($lista)

	wybierzListe($lista)
	$kopiaListy = ""
	While $kopiaListy = ""
		$kopiaListy = kopiujListe($lista)
	WEnd
	$tmp2 = StringSplit($kopiaListy, "	", 2)
	While $tmp2[0] <> "S" Or $tmp2[1] <> "Data" Or StringReplace($tmp2[2], @CRLF, "") <> "Numer"
		$hWnd = 0
		While $hWnd = 0
			ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{CTRLDOWN}{SHIFTDOWN}k{CTRLUP}{SHIFTUP}")
			$hWnd = WinGetHandle("Lista kolumn")
		WEnd
		GUISetState(@SW_HIDE, $GuiOczekiwanie)
		MsgBox($MB_TOPMOST, "Wybór kolumn - " & $lista, "Proszę wybrać kolumny do wyświetlenia w kolejności:" & @LF & "- Status dokumentu" & @LF & "- Data wystawienia" & @LF & "- Numer")
		While $hWnd <> 0
			$hWnd = WinGetHandle("Lista kolumn")
		WEnd
		GUISetState(@SW_SHOW, $GuiOczekiwanie)
		$kopiaListy = kopiujListe($lista)
		$tmp2 = StringSplit($kopiaListy, "	", 2)
	WEnd

EndFunc   ;==>ustawKolumnyListy

;---------------------------------------------------------------------------------------------------------------------
;pobiera wszystkie uchwyty
Func pobierzUchwytFiltrow()
	logs('@@ (848) :(' & @MIN & ':' & @SEC & ') pobierzUchwytFiltrow()' & @CR) ;### Function Trace
	ControlFocus("", "", uchwytSubiekta())
	;wybierzListe($LST_FAKTURY_ZAKUPU)
	wybierzListe($LST_WYDANIA_MAGAZYNOWE)
	wybierzListe($LST_PRZYJECIA_MAGAZYNOWE)
	wybierzListe($LST_ZWROTY_DETALICZNE)
	wybierzListe($LST_SPRZEDAZ_DETALICZNA)

	$_UchwytyKontrolek.Add($LST_SPRZEDAZ_DETALICZNA, ObjCreate("Scripting.Dictionary"))
	$_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).add($HWND_OKRES, zapiszUchwytFiltry($LST_SPRZEDAZ_DETALICZNA, $FLT_DOKUMENTY_Z_OKRESU))
	$_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).add($HWND_TYPU, zapiszUchwytFiltry($LST_SPRZEDAZ_DETALICZNA, $FLT_TYPU))
	$_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).add($HWND_KATEGORIA, zapiszUchwytFiltry($LST_SPRZEDAZ_DETALICZNA, $FLT_O_KATEGORII_NS))
	$_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).add($HWND_FLAGA, zapiszUchwytFiltry($LST_SPRZEDAZ_DETALICZNA, $FLT_Z_FLAGA))
	$_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).Add($CLS_GXWND, zapiszUchwytGXWND($LST_SPRZEDAZ_DETALICZNA))
;~ 	$_UchwytyKontrolek.Add($LST_FAKTURY_ZAKUPU, ObjCreate("Scripting.Dictionary"))
;~ 	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_WEDLUG, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_DOKUMENTY_WG))
;~ 	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_OKRES, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_Z_OKRESU))
;~ 	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_TYPU, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_TYPU))
;~ 	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_STAN, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_O_STANIE_NP))
;~ 	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_KATEGORIA, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_O_KATEGORII))
;~ 	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_FLAGA, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_Z_FLAGA))
;~ 	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).Add($CLS_GXWND, zapiszUchwytGXWND($LST_FAKTURY_ZAKUPU))
	$_UchwytyKontrolek.Add($LST_WYDANIA_MAGAZYNOWE, ObjCreate("Scripting.Dictionary"))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).add($HWND_OKRES, zapiszUchwytFiltry($LST_WYDANIA_MAGAZYNOWE, $FLT_DOKUMENTY_Z_OKRESU))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).add($HWND_TYPU, zapiszUchwytFiltry($LST_WYDANIA_MAGAZYNOWE, $FLT_TYPU))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).add($HWND_STAN, zapiszUchwytFiltry($LST_WYDANIA_MAGAZYNOWE, $FLT_O_STANIE))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).add($HWND_KATEGORIA, zapiszUchwytFiltry($LST_WYDANIA_MAGAZYNOWE, $FLT_O_KATEGORII_NP))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).add($HWND_FLAGA, zapiszUchwytFiltry($LST_WYDANIA_MAGAZYNOWE, $FLT_Z_FLAGA))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).Add($CLS_GXWND, zapiszUchwytGXWND($LST_WYDANIA_MAGAZYNOWE))
	$_UchwytyKontrolek.Add($LST_PRZYJECIA_MAGAZYNOWE, ObjCreate("Scripting.Dictionary"))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).add($HWND_OKRES, zapiszUchwytFiltry($LST_PRZYJECIA_MAGAZYNOWE, $FLT_DOKUMENTY_Z_OKRESU))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).add($HWND_TYPU, zapiszUchwytFiltry($LST_PRZYJECIA_MAGAZYNOWE, $FLT_TYPU))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).add($HWND_STAN, zapiszUchwytFiltry($LST_PRZYJECIA_MAGAZYNOWE, $FLT_O_STANIE))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).add($HWND_KATEGORIA, zapiszUchwytFiltry($LST_PRZYJECIA_MAGAZYNOWE, $FLT_O_KATEGORII_NP))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).add($HWND_FLAGA, zapiszUchwytFiltry($LST_PRZYJECIA_MAGAZYNOWE, $FLT_Z_FLAGA))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).Add($CLS_GXWND, zapiszUchwytGXWND($LST_PRZYJECIA_MAGAZYNOWE))
	$_UchwytyKontrolek.Add($LST_ZWROTY_DETALICZNE, ObjCreate("Scripting.Dictionary"))
	$_UchwytyKontrolek.Item($LST_ZWROTY_DETALICZNE).add($HWND_OKRES, zapiszUchwytFiltry($LST_ZWROTY_DETALICZNE, $FLT_DOKUMENTY_Z_OKRESU))
	$_UchwytyKontrolek.Item($LST_ZWROTY_DETALICZNE).add($HWND_RODZAJ, zapiszUchwytFiltry($LST_ZWROTY_DETALICZNE, $FLT_RODZAJ_ZWROTU))
	$_UchwytyKontrolek.Item($LST_ZWROTY_DETALICZNE).add($HWND_KATEGORIA, zapiszUchwytFiltry($LST_ZWROTY_DETALICZNE, $FLT_O_KATEGORII_NS))
	$_UchwytyKontrolek.Item($LST_ZWROTY_DETALICZNE).add($HWND_FLAGA, zapiszUchwytFiltry($LST_ZWROTY_DETALICZNE, $FLT_Z_FLAGA_NS))
	$_UchwytyKontrolek.Item($LST_ZWROTY_DETALICZNE).Add($CLS_GXWND, zapiszUchwytGXWND($LST_ZWROTY_DETALICZNE))
;~ 	ustawKolumnyListy($LST_FAKTURY_ZAKUPU)
	ustawKolumnyListy($LST_WYDANIA_MAGAZYNOWE)
	ustawKolumnyListy($LST_PRZYJECIA_MAGAZYNOWE)
	ustawKolumnyListy($LST_ZWROTY_DETALICZNE)
	ustawKolumnyListy($LST_SPRZEDAZ_DETALICZNA)
	Sleep(5000)

	;------magazyn------
	$_UchwytyKontrolek.Add($CTR_MAGAZYN, pobierzUchwytMagazynow())
EndFunc   ;==>pobierzUchwytFiltrow

;---------------------------------------------------------------------------------------------------------------------
; restartowanie subiekta

Func restartujSubiekta()
	logs('@@ (902) :(' & @MIN & ':' & @SEC & ') restartujSubiekta()' & @CR) ;### Function Trace
	dodajStatus("Restartowanie Subiekta")
	ControlSend(uchwytSubiekta(), "", "", "!{F4}") ; zamkniecie subiekta
	Sleep(3000)
	$archiwizacja = 1
	Do
		$archiwizacja = oknoArchiwizacja()	;zamkniecie okna archiwizacji
	Until $archiwizacja = 0 And zlapSubiekta($_nazwaOknaSubiekta) = 0
	Run($_SciezkaUruchomieniaSubiekta, "") ;uruchomienie subiekta
	While czySubiektUruchomiony($_nazwaOknaSubiekta) <> 0 ;Czekamy az subiekt sie zaladuje
		Sleep(1000)
	WEnd
	;wszystkie czynnosci uchwytow trzeba zrobic od nowa
	$_UchwytyKontrolek = ObjCreate("Scripting.Dictionary")
	$_UchwytyKontrolek.Add($SUBIEKT, WinGetHandle($_nazwaOknaSubiekta))
	pobierzUchwytFiltrow()
	GUISetState(@SW_HIDE, $GuiOczekiwanie)
	$_RESTART = False

EndFunc   ;==>restartujSubiekta

;---------------------------------------------------------------------------------------------------------------------
; wykrywa i zamyka okno archiwizacji
Func oknoArchiwizacja()
	logs('@@ (929) :(' & @MIN & ':' & @SEC & ') oknoArchiwizacja()' & @CR) ;### Function Trace
	$hSub = WinGetHandle("Subiekt GT")
	If $hSub <> 0 Then
		$uchwyt = szukajUchwytu($hSub, "&Nie", "Button")
		ControlClick($uchwyt, "", "")
	EndIf
	Return $hSub
EndFunc   ;==>oknoArchiwizacja

;----------------------------------------------------------------------------------------------------------
; pauzujemy bota

Func pauzaSzybkaBot()
	logs('@@ (183) :(' & @MIN & ':' & @SEC & ') pauzaSzybka()' & @CR) ;### Function Trace
	If $_PAUZA = False Then
		$_PAUZA = True
		dodajStatus("PAUZA")
		logs("PAUZA........." & @LF)
		pauzaBot()
	Else
		$_PAUZA = False
		dodajStatus("WZNOWIENIE")
		logs("START........." & @LF)
		;		ControlFocus($t, "", "")
	EndIf
EndFunc   ;==>pauzaSzybka

Func pauzaBot()
	logs('@@ (183) :(' & @MIN & ':' & @SEC & ') pauzaSzybka()' & @CR) ;### Function Trace
	While $_PAUZA
		Sleep(100)
	WEnd
EndFunc   ;==>pauzaSzybka

