;#####################################################################################################################
;#####################################################################################################################
;################################													##################################
;################################			Biblioteka do obslugi okienek 			##################################
;################################				w programie Subiekt GT				##################################
;################################													##################################
;################################								Mateusz Blaszczak	##################################
;#####################################################################################################################
;#####################################################################################################################
;#####################################################################################################################
;#####################################################################################################################
;#####################################################################################################################
;#####################################################################################################################

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
Global $_nazwaOknaSubiekta = '[REGEXPTITLE:(?i)(.* - Subiekt GT)]'

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
Func wybierzListe($lista)
	ControlFocus($_nazwaOknaSubiekta, "", "")
	Switch $lista
		Case "PA"
			$uchwyt = 0
			While $uchwyt = 0
				Send("{ALTDOWN}2")
				Send("{ALTUP}")
				$uchwyt = szukajUchwytu(MWinGetHandle(),"Sprzedaż detaliczna","AfxWnd80")
			WEnd
		Case "FA"
			$uchwyt = 0
			While $uchwyt = 0
				Send("{ALTDOWN}3")
				Send("{ALTUP}")
				$uchwyt = szukajUchwytu(MWinGetHandle(),"Faktury zakupu","AfxWnd80")
			WEnd
		Case "KFZ"
			$uchwyt = 0
			While $uchwyt = 0
				Send("{LALT}WZK")
				$uchwyt = szukajUchwytu(MWinGetHandle(),"Korekty zakupu","AfxWnd80")
			WEnd
		Case "MM"
			$uchwyt = 0
			While $uchwyt = 0
				Send("{ALTDOWN}4")
				Send("{ALTUP}")
				$uchwyt = szukajUchwytu(MWinGetHandle(),"Wydania magazynowe","AfxWnd80")
			WEnd
			$hWnd = MWinGetHandle()
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Wydania magazynowe", "ins_hyperlink", ", typu:", "Static")
			$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
			ControlFocus($uchwyt,"","")
			ControlClick($uchwyt, "", "")
			Send("{END}{ENTER}")
		Case "RW"
			$uchwyt = 0
			While $uchwyt = 0
				Send("{ALTDOWN}4")
				Send("{ALTUP}")
				$uchwyt = szukajUchwytu(MWinGetHandle(),"Wydania magazynowe","AfxWnd80")
			WEnd
			$hWnd = MWinGetHandle()
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Wydania magazynowe", "ins_hyperlink", ", typu:", "Static")
			$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
			ControlFocus($uchwyt,"","")
			ControlClick($uchwyt, "", "")
			Send("{END}{UP}{ENTER}")
		Case "PW"
			$uchwyt = 0
			While $uchwyt = 0
				Send("{ALTDOWN}5")
				Send("{ALTUP}")
				$uchwyt = szukajUchwytu(MWinGetHandle(),"Przyjęcia magazynowe","AfxWnd80")
			WEnd
			$hWnd = MWinGetHandle()
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Przyjęcia magazynowe", "ins_hyperlink", ", typu:", "Static")
			$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
			ControlFocus($uchwyt,"","")
			ControlClick($uchwyt, "", "")
			Send("{END}{UP}{ENTER}")
		Case "ZW"
			$uchwyt = 0
			While $uchwyt = 0
				Send("{LALT}WSZ")
				$uchwyt = szukajUchwytu(MWinGetHandle(),"Zwroty detaliczne","AfxWnd80")
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
		Case "PA"
			$lista = "Sprzedaż detaliczna"
		Case "FA"
			$lista = "Faktury zakupu"
		Case "KFZ"
			$lista = "Korekty zakupu"
		Case "MM"
			$lista = "Wydania magazynowe"
		Case "RW"
			$lista = "Wydania magazynowe"
		Case "PW"
			$lista = "Przyjęcia magazynowe"
		Case "ZW"
			$lista = "Zwroty detaliczne"
	EndSwitch
	Local $uchwyt = 0
	Local $hWnd = szukajUchwytu(MWinGetHandle(), $lista, "ins_hyperlink")
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
		Case "PA"
			$lista = "Sprzedaż detaliczna"
		Case "FA"
			$lista = "Faktury zakupu"
		Case "KFZ"
			$lista = "Korekty zakupu"
		Case "MM"
			$lista = "Wydania magazynowe"
		Case "RW"
			$lista = "Wydania magazynowe"
		Case "PW"
			$lista = "Przyjęcia magazynowe"
		Case "ZW"
			$lista = "Zwroty detaliczne"
	EndSwitch
	Local $uchwyt = 0
	Local $hWnd = szukajUchwytu(MWinGetHandle(), $lista, "ins_hyperlink")
	$hWnd = _WinAPI_GetAncestor($hWnd)
	$hWnd = _WinAPI_GetAncestor($hWnd)
	$uchwyt = szukajUchwytu($hWnd, "", "GXWND")
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
EndFunc

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
	$hwnd = WinGetHandle('[REGEXPTITLE:(?i)(.* - Subiekt GT)]')
	If @error Then
		MsgBox(null,"Brak uchwytu","Brak uchwytu")
	Else
	EndIf
	Return $hwnd
EndFunc

;---------------------------------------------------------------------------------------------------------------------
;Wybór magazyn
;Przyjmuje skrócona nazwe magazynu
Func magazyn($nazwa)
	$hWnd = MWinGetHandle()
	ControlFocus("","",$hWnd)
	Send("{CTRLDOWN}{F3}")
	Send("{CTRLUP}")
	homeBtn()
	Local $mag[2] = ["",""]
	While  StringCompare($mag[1],$nazwa) <> 0 And StringCompare(_ArrayToString($mag,"	"),kopiujPozycje()) <> 0
		$mag = StringSplit(kopiujPozycje(),"	")
		If StringCompare($mag[1],$nazwa) = 0 Then
			$hSub = WinGetHandle("Zmiana magazynu")
			$uchwyt = szukajUchwytu($hSub, "OK", "Button")
			ControlClick($uchwyt,"","")
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
		Case "PA"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Sprzedaż detaliczna", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "FA"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Faktury zakupu", "ins_hyperlink", ", z okresu:", "Static")
		Case "KFZ"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Korekty zakupu", "ins_hyperlink", ", z okresu:", "Static")
		Case "MM"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Wydania magazynowe", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "RW"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Wydania magazynowe", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "PW"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Przyjęcia magazynowe", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "ZW"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Zwroty detaliczne", "ins_hyperlink", "Dokumenty z okresu:", "Static")
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
	For $i = 1 to $listy[0]
		ustawDomyslnyFiltr($listy[$i])
	Next
EndFunc   ;==>ustawieniaDomyslne


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
	wybierzListe($lista)
	$hWnd = WinGetHandle($_nazwaOknaSubiekta)
	ControlClick($hWnd,"","")
	$uchwyt = 0
	Switch $lista
		Case "PA"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Sprzedaż detaliczna", "ins_hyperlink", "Filtr: ", "Button")
		Case "FA"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Faktury zakupu", "ins_hyperlink", "Filtr: ", "Button")
		Case "KFZ"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Korekty zakupu", "ins_hyperlink", "Filtr: ", "Button")
		Case "MM"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Wydania magazynowe", "ins_hyperlink", "Filtr: ", "Button")
		Case "RW"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Wydania magazynowe", "ins_hyperlink", "Filtr: ", "Button")
		Case "PW"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Przyjęcia magazynowe", "ins_hyperlink", "Filtr: ", "Button")
		Case "ZW"
			$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, "Zwroty detaliczne", "ins_hyperlink", "Filtr: ", "Button")
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
EndFunc   ;==>ustawOkres


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
	ControlClick($hWnd,"","")
	$uchwyt = 0
	Switch $lista
		Case "PA"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Sprzedaż detaliczna", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "FA"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Faktury zakupu", "ins_hyperlink", ", z okresu:", "Static")
		Case "KFZ"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Korekty zakupu", "ins_hyperlink", ", z okresu:", "Static")
		Case "MM"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Wydania magazynowe", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "RW"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Wydania magazynowe", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "PW"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Przyjęcia magazynowe", "ins_hyperlink", "Dokumenty z okresu:", "Static")
		Case "ZW"
			$uchwyt = pobierzUchwytKontrolkiListy($hWnd, "Zwroty detaliczne", "ins_hyperlink", "Dokumenty z okresu:", "Static")
	EndSwitch
	$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	ControlFocus($uchwyt, "", "")
	ControlClick($uchwyt, "", "")
	Send("{HOME}{ENTER}")
EndFunc   ;==>ustawOkres


Func pobierzUchwytKontrolkiListyPetla($hWnd, $broNM, $broCL, $childNM, $childCL)
	$wynik = 0
	While $wynik = 0
		$wynik = pobierzUchwytKontrolkiListy($hWnd, $broNM, $broCL, $childNM, $childCL)
		ConsoleWrite($wynik & " " & $hWnd & " " &  $broNM & " " & $broCL & " " &  $childNM & " " &  $childCL & @LF)
	WEnd
	Return $wynik
EndFunc

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
		$hWnd = WinGetHandle('[REGEXPTITLE:(?i)(Informacja o nowej wersji: *)]')
		If $hWnd <> 0 Then
			$uchwyt = szukajUchwytu($hWnd, "&Zamknij", "Button")
			ControlClick($uchwyt,"","")
		EndIf
		Return 0 ;subiekt
	Else
		Return -1 ;brak
	EndIf
EndFunc   ;==>czySubiektUruchomiony


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