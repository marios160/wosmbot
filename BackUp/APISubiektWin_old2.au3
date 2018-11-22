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
Global $_nazwaOknaSubiekta = '[REGEXPTITLE:(?i)(.* - Subiekt GT)]'
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
;--------------
Global $_UchwytyKontrolek = ObjCreate("Scripting.Dictionary")

;---------------------------------------------------------------------------------------------------------------------
;Funkcja uruchamia listy dokumentów
;Przyjmuje nazwe listy która chcemy wyswietlic
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe

;ZW - zwroty detaliczne

;---------------------------------------------------------------------------------------------------------------------
;Funkcja uruchamia listy dokumentów
;Przyjmuje nazwe listy która chcemy wyswietlic
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe

;ZW - zwroty detaliczne
Func wybierzListeDeprecated($lista)
	ControlFocus($_nazwaOknaSubiekta, "", "")
	$uchwyt = 0
	While $uchwyt = 0
		Switch $lista
			Case $LST_SPRZEDAZ_DETALICZNA
					Send("{ALTDOWN}2{ALTUP}")
			Case $LST_FAKTURY_ZAKUPU
					Send("{ALTDOWN}3{ALTUP}")
			Case $LST_KOREKTY_ZAKUPU
					Send("{LALT}WZK")
			Case $LST_WYDANIA_MAGAZYNOWE
					Send("{ALTDOWN}4{ALTUP}")
			Case $LST_PRZYJECIA_MAGAZYNOWE
					Send("{ALTDOWN}5{ALTUP}")
			Case $LST_ZWROTY_DETALICZNE
					Send("{LALT}WSZ")
		EndSwitch
		$uchwyt = szukajUchwytu(MWinGetHandle(), $lista, $CLS_AFXWND80)
	WEnd
EndFunc   ;==>wybierzListe



Func wybierzListePrzestarzale($lista)
	ControlFocus($_nazwaOknaSubiekta, "", "")
	Switch $lista
		Case $PA
			$uchwyt = 0
			While $uchwyt = 0
				Send("{ALTDOWN}2")
				Send("{ALTUP}")
				$uchwyt = szukajUchwytu(MWinGetHandle(), $LST_SPRZEDAZ_DETALICZNA, $CLS_AFXWND80)
			WEnd
		Case $FA
			$uchwyt = 0
			While $uchwyt = 0
				Send("{ALTDOWN}3")
				Send("{ALTUP}")
				$uchwyt = szukajUchwytu(MWinGetHandle(), $LST_FAKTURY_ZAKUPU, $CLS_AFXWND80)
			WEnd
		Case $KFZ
			$uchwyt = 0
			While $uchwyt = 0
				Send("{LALT}WZK")
				$uchwyt = szukajUchwytu(MWinGetHandle(), $LST_KOREKTY_ZAKUPU, $CLS_AFXWND80)
			WEnd
		Case $MM
			$uchwyt = 0
			While $uchwyt = 0
				Send("{ALTDOWN}4")
				Send("{ALTUP}")
				$uchwyt = szukajUchwytu(MWinGetHandle(), $LST_WYDANIA_MAGAZYNOWE, $CLS_AFXWND80)
			WEnd
			$hWnd = MWinGetHandle()
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_WYDANIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_TYPU, $CLS_STATIC)
			$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
			ControlFocus($uchwyt, "", "")
			ControlClick($uchwyt, "", "")

			Send("{END}{ENTER}")
		Case $RW
			$uchwyt = 0
			While $uchwyt = 0
				Send("{ALTDOWN}4")
				Send("{ALTUP}")
				$uchwyt = szukajUchwytu(MWinGetHandle(), $LST_WYDANIA_MAGAZYNOWE, $CLS_AFXWND80)
			WEnd
			$hWnd = MWinGetHandle()
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_WYDANIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_TYPU, $CLS_STATIC)
			$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
			ControlFocus($uchwyt, "", "")
			ControlClick($uchwyt, "", "")
			Send("{END}{UP}{ENTER}")
		Case $PW
			$uchwyt = 0
			While $uchwyt = 0
				Send("{ALTDOWN}5")
				Send("{ALTUP}")
				$uchwyt = szukajUchwytu(MWinGetHandle(), $LST_PRZYJECIA_MAGAZYNOWE, $CLS_AFXWND80)
			WEnd
			$hWnd = MWinGetHandle()
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_PRZYJECIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_TYPU, $CLS_STATIC)
			$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
			ControlFocus($uchwyt, "", "")
			ControlClick($uchwyt, "", "")
			Send("{END}{UP}{ENTER}")
		Case $ZW
			$uchwyt = 0
			While $uchwyt = 0
				Send("{LALT}WSZ")
				$uchwyt = szukajUchwytu(MWinGetHandle(), $LST_ZWROTY_DETALICZNE, $CLS_AFXWND80)
			WEnd
	EndSwitch
EndFunc   ;==>wybierzListe


;---------------------------------------------------------------------------------------------------------------------
;Zlapanie fokusa na liste dokumentow
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
		Case $PA
			$lista = $LST_SPRZEDAZ_DETALICZNA
		Case $FA
			$lista = $LST_FAKTURY_ZAKUPU
		Case $KFZ
			$lista = $LST_KOREKTY_ZAKUPU
		Case $MM
			$lista = $LST_WYDANIA_MAGAZYNOWE
		Case $RW
			$lista = $LST_WYDANIA_MAGAZYNOWE
		Case $PW
			$lista = $LST_PRZYJECIA_MAGAZYNOWE
		Case $ZW
			$lista = $LST_ZWROTY_DETALICZNE
	EndSwitch
	Local $uchwyt = 0
	Local $hWnd = szukajUchwytu(MWinGetHandle(), $lista, $CLS_INS_HYPERLINK)
	$hWnd = _WinAPI_GetAncestor($hWnd)
	$hWnd = _WinAPI_GetAncestor($hWnd)
	$tmp = ControlGetPos($hWnd, "", "")
	$uchwyt = szukajUchwytu($hWnd, "", $CLS_GXWND)
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
;Zlapanie fokusa na liste dokumentow
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne
Func fokusNaListe($nazwa)
	Local $lista = ""
	Switch $nazwa
		Case $PA
			$lista = $LST_SPRZEDAZ_DETALICZNA
		Case $FA
			$lista = $LST_FAKTURY_ZAKUPU
		Case $KFZ
			$lista = $LST_KOREKTY_ZAKUPU
		Case $MM
			$lista = $LST_WYDANIA_MAGAZYNOWE
		Case $RW
			$lista = $LST_WYDANIA_MAGAZYNOWE
		Case $PW
			$lista = $LST_PRZYJECIA_MAGAZYNOWE
		Case $ZW
			$lista = $LST_ZWROTY_DETALICZNE
	EndSwitch
	Local $uchwyt = 0
	Local $hWnd = szukajUchwytu(MWinGetHandle(), $lista, $CLS_INS_HYPERLINK)
	$hWnd = _WinAPI_GetAncestor($hWnd)
	$hWnd = _WinAPI_GetAncestor($hWnd)
	$uchwyt = szukajUchwytu($hWnd, "", $CLS_GXWND)
	$tmp = ControlGetPos($uchwyt, "", "")
	ControlFocus($uchwyt, "", "")
	Return $uchwyt
EndFunc   ;==>fokusNaListe



;---------------------------------------------------------------------------------------------------------------------
Func szukajUchwytuPetla($hWnd, $text, $class)
	$uchwyt = 0
	While $uchwyt = 0
		$uchwyt = szukajUchwytu($hWnd, $text, $class)
	WEnd
	Return $uchwyt
EndFunc   ;==>szukajUchwytuPetla

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


Func MWinGetHandle()
	$hWnd = WinGetHandle('[REGEXPTITLE:(?i)(.* - Subiekt GT)]')
	If @error Then
		MsgBox(Null, "Brak uchwytu", "Brak uchwytu")
	Else
	EndIf
	Return $hWnd
EndFunc   ;==>MWinGetHandle

;---------------------------------------------------------------------------------------------------------------------
;Wybór magazyn
;Przyjmuje skrócona nazwe magazynu
Func magazynDeprecated($nazwa)
	$hWnd = MWinGetHandle()
	ControlFocus("", "", $hWnd)
	Send("{CTRLDOWN}{F3}")
	Send("{CTRLUP}")
	homeBtn()
	Local $mag[2] = ["", ""]
	While StringCompare($mag[1], $nazwa) <> 0 And StringCompare(_ArrayToString($mag, "	"), kopiujPozycje()) <> 0
		$mag = StringSplit(kopiujPozycje(), "	")
		If StringCompare($mag[1], $nazwa) = 0 Then
			$hSub = WinGetHandle("Zmiana magazynu")
			$uchwyt = szukajUchwytu($hSub, "OK", $CLS_BUTTON)
			ControlClick($uchwyt, "", "")
			ExitLoop
		EndIf
		downBtn()
	WEnd
EndFunc   ;==>magazyn

;---------------------------------------------------------------------------------------------------------------------
;Ustawia filtr wyswietlajacy dokumenty tylko z danego dnia
;Przyjmuje date która chcemy ustawic, nazwe listy na ktorej chcemy ustawic okres
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne
Func ustawOkres($data, $lista)
	$dataTab = StringReplace($data, '/', '-')
	$hWnd = MWinGetHandle()
	$uchwyt = 0
	Switch $lista
		Case $PA
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_SPRZEDAZ_DETALICZNA, $CLS_INS_HYPERLINK, $FLT_DOKUMENTY_Z_OKRESU, $CLS_STATIC)
		Case $FA
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_FAKTURY_ZAKUPU, $CLS_INS_HYPERLINK, $FLT_Z_OKRESU, $CLS_STATIC)
		Case $KFZ
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_KOREKTY_ZAKUPU, $CLS_INS_HYPERLINK, $FLT_Z_OKRESU, $CLS_STATIC)
		Case $MM
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_WYDANIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_DOKUMENTY_Z_OKRESU, $CLS_STATIC)
		Case $RW
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_WYDANIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_DOKUMENTY_Z_OKRESU, $CLS_STATIC)
		Case $PW
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_PRZYJECIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_DOKUMENTY_Z_OKRESU, $CLS_STATIC)
		Case $ZW
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_ZWROTY_DETALICZNE, $CLS_INS_HYPERLINK, $FLT_DOKUMENTY_Z_OKRESU, $CLS_STATIC)
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
Func filtryDomyslne($listy)
	For $i = 1 To $listy[0]
		ustawDomyslnyFiltr($listy[$i])
	Next
EndFunc   ;==>filtryDomyslne


;---------------------------------------------------------------------------------------------------------------------
;Ustawia filtr wyswietlajacy dokumenty tylko z danego dnia
;Przyjmuje date która chcemy ustawic, nazwe listy na ktorej chcemy ustawic okres
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne
Func ustawDomyslnyFiltr($lista)
;	wybierzListe($lista)
	$hWnd = WinGetHandle($_nazwaOknaSubiekta)
	ControlClick($hWnd, "", "")
	$uchwyt = 0
	Switch $lista
		Case $PA
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_SPRZEDAZ_DETALICZNA, $CLS_INS_HYPERLINK, $FLT_FILTR, $CLS_BUTTON)
		Case $FA
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_FAKTURY_ZAKUPU, $CLS_INS_HYPERLINK, $FLT_FILTR, $CLS_BUTTON)
		Case $KFZ
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_KOREKTY_ZAKUPU, $CLS_INS_HYPERLINK, $FLT_FILTR, $CLS_BUTTON)
		Case $MM
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_WYDANIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_FILTR, $CLS_BUTTON)
		Case $RW
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_WYDANIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_FILTR, $CLS_BUTTON)
		Case $PW
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_PRZYJECIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_FILTR, $CLS_BUTTON)
		Case $ZW
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $LST_ZWROTY_DETALICZNE, $CLS_INS_HYPERLINK, $FLT_FILTR, $CLS_BUTTON)
		Case Else
			Return
	EndSwitch
	If $uchwyt <> 0 Then
		ControlFocus($uchwyt, "", "")
		$x = ControlClick($uchwyt, "", "")
		Send("P")
		;MsgBox(null,"ASd","ASD")
		ConsoleWrite($x & @LF)
		;Sleep(5000)
	EndIf
EndFunc   ;==>ustawDomyslnyFiltr


;---------------------------------------------------------------------------------------------------------------------
;Ustawia filtr wyswietlajacy dokumenty na okres nieokreslony
;Przyjmuje nazwe listy na ktorej chcemy ustawic okres
;
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne
Func ustawOkresNieokreslony($lista)
	$hWnd = WinGetHandle($_nazwaOknaSubiekta)
	ControlClick($hWnd, "", "")
	$uchwyt = 0
	Switch $lista
		Case $PA
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, $LST_SPRZEDAZ_DETALICZNA, $CLS_INS_HYPERLINK, $FLT_DOKUMENTY_Z_OKRESU, $CLS_STATIC)
		Case $FA
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, $LST_FAKTURY_ZAKUPU, $CLS_INS_HYPERLINK, $FLT_Z_OKRESU, $CLS_STATIC)
		Case $KFZ
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, $LST_KOREKTY_ZAKUPU, $CLS_INS_HYPERLINK, $FLT_Z_OKRESU, $CLS_STATIC)
		Case $MM
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, $LST_WYDANIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_DOKUMENTY_Z_OKRESU, $CLS_STATIC)
		Case $RW
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, $LST_WYDANIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_DOKUMENTY_Z_OKRESU, $CLS_STATIC)
		Case $PW
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, $LST_PRZYJECIA_MAGAZYNOWE, $CLS_INS_HYPERLINK, $FLT_DOKUMENTY_Z_OKRESU, $CLS_STATIC)
		Case $ZW
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, $LST_ZWROTY_DETALICZNE, $CLS_INS_HYPERLINK, $FLT_DOKUMENTY_Z_OKRESU, $CLS_STATIC)
	EndSwitch
	$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	ControlFocus($uchwyt, "", "")
	ControlClick($uchwyt, "", "")
	Send("{HOME}{ENTER}")
EndFunc   ;==>ustawOkresNieokreslony


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
;Lapie kontrolke w galezi danego rodzica
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
		;logs("PUSTA LISTA")
		Return True
	Else
		Return False
	EndIf
	Return False
EndFunc   ;==>czyPusta




;-----------------------------------------------------------------------------------------------------------------------------
Func czySubiektUruchomiony()


	$hWnd = zlapSubiekta()
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
Func zlapSubiekta()
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
Func kopiujPozycje()
	ControlFocus($_nazwaOknaSubiekta, "", "")
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
Func kopiujListe()
	ControlFocus($_nazwaOknaSubiekta, "", "")
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
Func homeBtn()
	Send("{CTRLDOWN}{HOME}")
	Send("{CTRLUP}")
EndFunc   ;==>homeBtn

;---------------------------------------------------------------------------------------------------------------------
Func endBtn()
	Send("{CTRLDOWN}{END}")
	Send("{CTRLUP}")
EndFunc   ;==>endBtn

;---------------------------------------------------------------------------------------------------------------------
Func downBtn()
	Send("{DOWN}")
EndFunc   ;==>downBtn
