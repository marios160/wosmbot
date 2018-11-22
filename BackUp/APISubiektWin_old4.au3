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
Global $SUBIEKT = "Subiekt"
Global $_UchwytyKontrolek = ObjCreate("Scripting.Dictionary") ;lista z uchwytami do kontrolek

;---------------------------------------------------------------------------------------------------------------------
Func znajdzKomunikat()
	$hwnd = ControlGetHandle("[TITLE:Subiekt GT;CLASS:#32770]","","")
	$uchwyt = _WinAPI_GetWindow($hwnd,$GW_CHILD)
	While StringInStr(_WinAPI_GetWindowText($uchwyt),"Czy na pewno wykonać operację zmiany skutku magazynowego?") = 0 And $uchwyt <> 0
		$uchwyt = _WinAPI_GetWindow($uchwyt,$GW_HWNDNEXT)
	WEnd
	$uchwyt = _WinAPI_GetWindow($hwnd,$GW_CHILD)
	While StringInStr(_WinAPI_GetWindowText($uchwyt),"&Tak") = 0 And $uchwyt <> 0
		$uchwyt = _WinAPI_GetWindow($uchwyt,$GW_HWNDNEXT)
	WEnd
	ControlSend($uchwyt,"","","T")
	Return $hwnd
EndFunc

;---------------------------------------------------------------------------------------------------------------------
;Wywoluje skutek magazynowy aktualnie zaznaczonego rekordu
Func wywolajSkutekMagazynowy($lista)
	If $lista = $LST_PRZYJECIA_MAGAZYNOWE Or $lista = $LST_WYDANIA_MAGAZYNOWE Then
		WinMenuSelectItem(uchwytSubiekta(), "", $MNU_OPERACJE, $SUB_WYWOLAJ_SKUTEK_MAGAZYNOWY_S)
	Else
		WinMenuSelectItem(uchwytSubiekta(), "", $MNU_OPERACJE, $SUB_WYWOLAJ_SKUTEK_MAGAZYNOWY_W)
	EndIf
EndFunc   ;==>wywolajSkutekMagazynowy

;---------------------------------------------------------------------------------------------------------------------
;Odklada skutek magazynowy aktualnie zaznaczonego rekordu
Func odlozSkutekMagazynowy($lista)
	WinMenuSelectItem(uchwytSubiekta(), "", $MNU_OPERACJE, $SUB_ODLOZ_SKUTEK_MAGAZYNOWY)
	While ControlFocus(uchwytSubiekta(),"","") = 0
		znajdzKomunikat()
	WEnd
EndFunc   ;==>odlozSkutekMagazynowy


;---------------------------------------------------------------------------------------------------------------------
;Ustawia filtry listy na domyslne
Func wyczyscFiltry($lista)
	For $key In $_UchwytyKontrolek.Item($lista)
		ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($lista).Item($key)), "SetCurrentSelection", 0)
	Next
EndFunc   ;==>wyczyscFiltry

;---------------------------------------------------------------------------------------------------------------------
;Uruchamia liste dokumentow uzywajac skrotu nazwy
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne
Func wybierzListeNN($lista)
	Switch ($lista)
		Case $PA
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_SPRZEDAZ_DETALICZNA)
		Case $FA
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_ZAKUP, $SUB_FAKTURY_ZAKUPU)
		Case $MM
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			wybierzFiltr($LST_WYDANIA_MAGAZYNOWE, $HWND_TYPU, $POZ_PRZESUNIECIE_MIEDZYMAGAZYNOWE)
		Case $RW
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			wybierzFiltr($LST_WYDANIA_MAGAZYNOWE, $HWND_TYPU, $POZ_ROZCHOD_WEWNETRZNY)
		Case $PW
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_PRZYJECIA_MAGAZYNOWE)
			wybierzFiltr($LST_PRZYJECIA_MAGAZYNOWE, $HWND_TYPU, $POZ_PRZYCHOD_WEWNETRZNY)
		Case $ZW
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_ZWROTY_DETALICZNE)
	EndSwitch
EndFunc   ;==>wybierzListeNN


;---------------------------------------------------------------------------------------------------------------------
;Uruchamia liste dokumentow; ustawia filtry jesli podamy je w liscie key=nazwa filtra, item=pozycja(nazwa)
Func wybierzListe($lista, $filtry = ObjCreate("Scripting.Dictionary"))
	$listaBezFiltra = $lista
	Switch ($lista)
		Case $LST_SPRZEDAZ_DETALICZNA
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_SPRZEDAZ_DETALICZNA)
		Case $LST_FAKTURY_ZAKUPU
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_ZAKUP, $SUB_FAKTURY_ZAKUPU)
		Case $LST_WYDANIA_MAGAZYNOWE
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
		Case $LST_PRZYJECIA_MAGAZYNOWE
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_PRZYJECIA_MAGAZYNOWE)
		Case $LST_ZWROTY_DETALICZNE
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_ZWROTY_DETALICZNE)
		Case $POZ_PRZESUNIECIE_MIEDZYMAGAZYNOWE & " (" & $LST_WYDANIA_MAGAZYNOWE & ")"
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			wybierzFiltr($LST_WYDANIA_MAGAZYNOWE, $HWND_TYPU, $POZ_PRZESUNIECIE_MIEDZYMAGAZYNOWE)
			$listaBezFiltra = $LST_WYDANIA_MAGAZYNOWE
		Case $POZ_ROZCHOD_WEWNETRZNY & " (" & $LST_WYDANIA_MAGAZYNOWE & ")"
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			wybierzFiltr($LST_WYDANIA_MAGAZYNOWE, $HWND_TYPU, $POZ_ROZCHOD_WEWNETRZNY)
			$listaBezFiltra = $LST_WYDANIA_MAGAZYNOWE
		Case $POZ_PRZYCHOD_WEWNETRZNY & " (" & $LST_PRZYJECIA_MAGAZYNOWE & ")"
			WinMenuSelectItem(uchwytSubiekta(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_PRZYJECIA_MAGAZYNOWE)
			wybierzFiltr($LST_PRZYJECIA_MAGAZYNOWE, $HWND_TYPU, $POZ_PRZYCHOD_WEWNETRZNY)
			$listaBezFiltra = $LST_PRZYJECIA_MAGAZYNOWE
	EndSwitch
	If $filtry.count > 0 Then
		For $key In $filtry
			wybierzFiltr($lista, $key, $filtry.Item($key))
		Next
	EndIf
	Return $listaBezFiltra
EndFunc   ;==>wybierzListe


;---------------------------------------------------------------------------------------------------------------------
;Lapanie uchwytu na konkretna kontrolke. Jesli kontrolki sie powtarzaja to zwróci uchwyt
;na pierwsza zlapana
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
;Lapanie uchwytu glownego okna subiekta
Func uchwytSubiekta()
	Return 	HWnd($_UchwytyKontrolek.Item($SUBIEKT))
EndFunc   ;==>uchwytSubiekta


;---------------------------------------------------------------------------------------------------------------------
;Ustawia filtr wyswietlajacy dokumenty tylko z danego dnia
;Przyjmuje date która chcemy ustawic, nazwe listy na ktorej chcemy ustawic okres
Func ustawOkres($data, $lista)
	$dataTab = StringReplace($data, '/', '-')
	_GUICtrlListBox_ClickItem(HWnd($_UchwytyKontrolek.Item($lista).Item($HWND_OKRES)),_GUICtrlListBox_SelectString(HWnd($_UchwytyKontrolek.Item($lista).Item($HWND_OKRES)),$POZ_DOWOLNY_OKRES))
	;wybierzFiltr($lista, $HWND_OKRES, $POZ_DOWOLNY_OKRES)
	Do
		$hWnd = WinGetHandle('Dowolny okres')
		$uchwyt = szukajUchwytu($hWnd, "Data &początkowa:", $CLS_STATIC)
		$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
		MsgBox(0,"","")
		ControlSend($uchwyt, "", "", $dataTab)
		$uchwyt = szukajUchwytu($hWnd, "Data &końcowa:", $CLS_STATIC)
		$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
		ControlSend($uchwyt, "", "", $dataTab)
	Until $hWnd = 0
EndFunc   ;==>ustawOkres

;---------------------------------------------------------------------------------------------------------------------
;Szuka uchwytu dopoki nie znajdzie
Func pobierzUchwytKontrolkiListyPetla($hWnd, $broNM, $broCL, $childNM, $childCL)
	$wynik = 0
	While $wynik = 0
		$wynik = pobierzUchwytKontrolkiListy($hWnd, $broNM, $broCL, $childNM, $childCL)
		;ConsoleWrite($wynik & " " & $hWnd & " " & $broNM & " " & $broCL & " " & $childNM & " " & $childCL & @LF)
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
	kopiujListe($lista)
	$tmp = StringSplit(ClipGet(), @LF)
	If $tmp[0] < 2 Then Return False
	If StringLen($tmp[2]) = 0 Then
		Return True
	Else
		Return False
	EndIf
	Return False
EndFunc   ;==>czyPusta


;-----------------------------------------------------------------------------------------------------------------------------
;Sprawdza na ktorym etapie uruchamiania jest subiekt
Func czySubiektUruchomiony($_nazwaOknaSubiekta)
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
	If $pozycja = "" Then
		ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($lista).Item($filtr)), "SetCurrentSelection", $pozycja)
	Else
		ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($lista).Item($filtr)), "SelectString", $pozycja)
	EndIf
EndFunc   ;==>wybierzFiltr

;---------------------------------------------------------------------------------------------------------------------
;ustawia podany magazyn
Func wybierzMagazyn($magazyn)
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
	$uchwyt = szukajUchwytu(uchwytSubiekta(), $CTR_MAGAZYN, $CLS_INS_COMBOHYPERLINK)
	ControlFocus($uchwyt, "", "")
	ControlClick($uchwyt, "", "")
	$uchwyt = WinGetHandle("[REGEXPCLASS:(Afx:.*.:800)]")
	$uchwyt2 = _WinAPI_GetWindow($uchwyt, $GW_CHILD)
	If _WinAPI_GetClassName($uchwyt2) = "ListBox" And _WinAPI_GetDlgCtrlID($uchwyt2) = 2001 Then
		Send("{ESC}")
		Return $uchwyt2
	EndIf
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
	otworzFiltr($lista, $filtr)
	$uchwyt = WinGetHandle("[REGEXPCLASS:(Afx:.*.:800)]")
	$uchwyt2 = _WinAPI_GetWindow($uchwyt, $GW_CHILD)
	If _WinAPI_GetClassName($uchwyt2) = "ListBox" And _WinAPI_GetDlgCtrlID($uchwyt2) = 2001 Then
		usunSeparatory($uchwyt2)
		ControlCommand("", "", $uchwyt2, "SetCurrentSelection", 0)
		Send("{ESC}")
		Return $uchwyt2
	EndIf
	Return 0
EndFunc   ;==>zapiszUchwytFiltry

;---------------------------------------------------------------------------------------------------------------------
;pobiera uchwyt do listy dokumentow
Func zapiszUchwytGXWND($lista)
	$uchwyt = szukajUchwytu(uchwytSubiekta(), $lista, $CLS_INS_HYPERLINK)
	$uchwyt = _WinAPI_GetAncestor($uchwyt, $GA_PARENT)
	While _WinAPI_GetClassName($uchwyt) <> $CLS_INS_GRIDCNT And $uchwyt <> 0
		$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	WEnd
	$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_CHILD)
	If _WinAPI_GetClassName($uchwyt) = $CLS_GXWND Then
		Return $uchwyt
	EndIf
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
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{DOWN}")
EndFunc   ;==>sendGXWNDDown
Func sendGXWNDUp($lista)
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{UP}")
EndFunc   ;==>sendGXWNDUp
Func sendGXWNDHomeHard($lista)
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{CTRLDOWN}{HOME}{CTRLUP}")
EndFunc   ;==>sendGXWNDHomeHard
Func sendGXWNDEndHard($lista)
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{CTRLDOWN}{END}{CTRLUP}")
EndFunc   ;==>sendGXWNDEndHard
Func sendGXWNDHome($lista)
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{HOME}")
EndFunc   ;==>sendGXWNDHome
Func sendGXWNDEnd($lista)
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{END}")
EndFunc   ;==>sendGXWNDEnd
Func sendGXWNDEnter($lista)
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{ENTER}")
EndFunc   ;==>sendGXWNDEnter
Func sendGXWNDEsc($lista)
	ControlSend(HWnd($_UchwytyKontrolek.Item($lista).Item($CLS_GXWND)), "", "", "{ESC}")
EndFunc   ;==>sendGXWNDEsc

;---------------------------------------------------------------------------------------------------------------------
;kopiuje tresc rekordu
Func kopiujWiersz($lista)
	ClipPut("XEDEX_W")
	While ClipGet() = "XEDEX_W"
		ControlSend(HWnd($_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).Item($CLS_GXWND)), "", "", "{CTRLDOWN}{SHIFTDOWN}Y{CTRLUP}{SHIFTUP}")
	WEnd
	Return ClipGet()
EndFunc   ;==>kopiujWiersz

;---------------------------------------------------------------------------------------------------------------------
;kopiuje cala liste
Func kopiujListe($lista)
	ClipPut("XEDEX_L")
	While ClipGet() = "XEDEX_L"
		ControlSend(HWnd($_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).Item($CLS_GXWND)), "", "", "{CTRLDOWN}{SHIFTDOWN}C{CTRLUP}{SHIFTUP}")
	WEnd
	Return ClipGet()
EndFunc   ;==>kopiujListe

;---------------------------------------------------------------------------------------------------------------------
;pobiera wszystkie uchwyty
Func pobierzUchwytFiltrow()
	ControlFocus("", "", uchwytSubiekta())
	wybierzListe($LST_FAKTURY_ZAKUPU)
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
	$_UchwytyKontrolek.Add($LST_FAKTURY_ZAKUPU, ObjCreate("Scripting.Dictionary"))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_WEDLUG, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_DOKUMENTY_WG))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_OKRES, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_Z_OKRESU))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_TYPU, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_TYPU))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_STAN, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_O_STANIE_NP))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_KATEGORIA, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_O_KATEGORII))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add($HWND_FLAGA, zapiszUchwytFiltry($LST_FAKTURY_ZAKUPU, $FLT_Z_FLAGA))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).Add($CLS_GXWND, zapiszUchwytGXWND($LST_FAKTURY_ZAKUPU))
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

	;------magazyn------
	$_UchwytyKontrolek.Add($CTR_MAGAZYN, pobierzUchwytMagazynow())
EndFunc   ;==>pobierzUchwytFiltrow
