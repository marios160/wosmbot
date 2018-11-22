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
#include <APISubiektWin.au3>
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




Global $_ListaMagazynowBot
;~ ;Sleep(5000)
;~ $uchwyt = _WinAPI_GetWindow(MWinGetHandle(),$GW_CHILD)
;~ ConsoleWrite(MWinGetHandle() & @LF)
ControlFocus(MWinGetHandle(),"","")
pobierzUchwytFiltrow3()

;~ wybierzListeNN($RW)

;~ wyczyscFiltry($LST_WYDANIA_MAGAZYNOWE)

;~ $f = ObjCreate("Scripting.Dictionary")
;~ $f.add($HWND_TYPU,$POZ_ROZCHOD_WEWNETRZNY)
;~ $f.add($HWND_OKRES,$POZ_BIEZACY_ROK)


;~ Sleep(5000)
;~ wybierzListeNN($PW)
;~ Sleep(3000)
;~ wybierzListeNN($RW)
;~ Sleep(3000)
;~ wybierzListeNN($FA)
;~ Sleep(3000)
;~ wybierzListeNN($MM)
;~ Sleep(3000)
;~ wybierzListeNN($ZW)

wybierzListeNN($PW)
;~ wybierzFiltr($LST_SPRZEDAZ_DETALICZNA,$HWND_OKRES,$POZ_BIEZACY_ROK)
;~ wybierzFiltr($LST_SPRZEDAZ_DETALICZNA,$HWND_OKRES)
wywolajSkutekMagazynowy($LST_PRZYJECIA_MAGAZYNOWE)

Func wywolajSkutekMagazynowy($lista)
	If $lista = $LST_PRZYJECIA_MAGAZYNOWE Or $lista = $LST_WYDANIA_MAGAZYNOWE Then
		WinMenuSelectItem(MWinGetHandle(), "", $MNU_OPERACJE, $SUB_WYWOLAJ_SKUTEK_MAGAZYNOWY_S)
	Else
		WinMenuSelectItem(MWinGetHandle(), "", $MNU_OPERACJE, $SUB_WYWOLAJ_SKUTEK_MAGAZYNOWY_W)
	EndIf
EndFunc

Func odlozSkutekMagazynowy($lista)
		WinMenuSelectItem(MWinGetHandle(), "", $MNU_OPERACJE, $SUB_ODLOZ_SKUTEK_MAGAZYNOWY)
EndFunc


Func wyczyscFiltry($lista)
	For $key In $_UchwytyKontrolek.Item($lista)
		ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($lista).Item($key)), "SetCurrentSelection", 0)
	Next
EndFunc

Func wybierzListeNN($lista)
	Switch ($lista)
		Case $PA
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_SPRZEDAZ_DETALICZNA)
		Case $FA
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_ZAKUP, $SUB_FAKTURY_ZAKUPU)
		Case $MM
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			wybierzFiltr($LST_WYDANIA_MAGAZYNOWE,$HWND_TYPU,$POZ_PRZESUNIECIE_MIEDZYMAGAZYNOWE)
		Case $RW
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
			wybierzFiltr($LST_WYDANIA_MAGAZYNOWE,$HWND_TYPU,$POZ_ROZCHOD_WEWNETRZNY)
		Case $PW
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_PRZYJECIA_MAGAZYNOWE)
			wybierzFiltr($LST_PRZYJECIA_MAGAZYNOWE,$HWND_TYPU,$POZ_PRZYCHOD_WEWNETRZNY)
		Case $ZW
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_ZWROTY_DETALICZNE)
	EndSwitch
EndFunc


Func wybierzListe($lista, $filtry = ObjCreate("Scripting.Dictionary"))
	Switch ($lista)
		Case $LST_SPRZEDAZ_DETALICZNA
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_SPRZEDAZ_DETALICZNA)
		Case $LST_FAKTURY_ZAKUPU
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_ZAKUP, $SUB_FAKTURY_ZAKUPU)
		Case $LST_WYDANIA_MAGAZYNOWE
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_WYDANIA_MAGAZYNOWE)
		Case $LST_PRZYJECIA_MAGAZYNOWE
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_MAGAZYN, $SUB_PRZYJECIA_MAGAZYNOWE)
		Case $LST_ZWROTY_DETALICZNE
			WinMenuSelectItem(MWinGetHandle(), "", $MNU_WIDOK, $SUB_SPRZEDAZ, $SUB_ZWROTY_DETALICZNE)
	EndSwitch
	If $filtry.count > 0 Then
		For $key in $filtry
			wybierzFiltr($lista,$key,$filtry.Item($key))
		Next
	EndIf
	;WinMenuSelectItem(MWinGetHandle(), "", "&Operacje","Odłóż skutek &magazynowy")
EndFunc

Func pobierzUchwytFiltrow3()

	ControlFocus("", "", MWinGetHandle())

	wybierzListe($LST_FAKTURY_ZAKUPU)
	wybierzListe($LST_WYDANIA_MAGAZYNOWE)
	wybierzListe($LST_PRZYJECIA_MAGAZYNOWE)
	wybierzListe($LST_ZWROTY_DETALICZNE)
	wybierzListe($LST_SPRZEDAZ_DETALICZNA)

	$_UchwytyKontrolek.Add($LST_SPRZEDAZ_DETALICZNA , ObjCreate("Scripting.Dictionary"))
	$_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).add( $HWND_OKRES,zapiszUchwyt($LST_SPRZEDAZ_DETALICZNA, $FLT_DOKUMENTY_Z_OKRESU))
	$_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).add( $HWND_TYPU, zapiszUchwyt($LST_SPRZEDAZ_DETALICZNA, $FLT_TYPU))
	$_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).add( $HWND_KATEGORIA, zapiszUchwyt($LST_SPRZEDAZ_DETALICZNA, $FLT_O_KATEGORII_NS))
	$_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA).add( $HWND_FLAGA, zapiszUchwyt($LST_SPRZEDAZ_DETALICZNA, $FLT_Z_FLAGA))

	$_UchwytyKontrolek.Add($LST_FAKTURY_ZAKUPU, ObjCreate("Scripting.Dictionary"))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add( $HWND_WEDLUG, zapiszUchwyt($LST_FAKTURY_ZAKUPU, $FLT_DOKUMENTY_WG))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add( $HWND_OKRES, zapiszUchwyt($LST_FAKTURY_ZAKUPU, $FLT_Z_OKRESU))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add( $HWND_TYPU, zapiszUchwyt($LST_FAKTURY_ZAKUPU, $FLT_TYPU))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add( $HWND_STAN, zapiszUchwyt($LST_FAKTURY_ZAKUPU, $FLT_O_STANIE_NP))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add( $HWND_KATEGORIA, zapiszUchwyt($LST_FAKTURY_ZAKUPU, $FLT_O_KATEGORII))
	$_UchwytyKontrolek.Item($LST_FAKTURY_ZAKUPU).add( $HWND_FLAGA, zapiszUchwyt($LST_FAKTURY_ZAKUPU, $FLT_Z_FLAGA))

	$_UchwytyKontrolek.Add($LST_WYDANIA_MAGAZYNOWE, ObjCreate("Scripting.Dictionary"))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).add( $HWND_OKRES,zapiszUchwyt($LST_WYDANIA_MAGAZYNOWE, $FLT_DOKUMENTY_Z_OKRESU))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).add( $HWND_TYPU,zapiszUchwyt($LST_WYDANIA_MAGAZYNOWE, $FLT_TYPU))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).add( $HWND_STAN,zapiszUchwyt($LST_WYDANIA_MAGAZYNOWE, $FLT_O_STANIE))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).add( $HWND_KATEGORIA,zapiszUchwyt($LST_WYDANIA_MAGAZYNOWE, $FLT_O_KATEGORII_NP))
	$_UchwytyKontrolek.Item($LST_WYDANIA_MAGAZYNOWE).add( $HWND_FLAGA,zapiszUchwyt($LST_WYDANIA_MAGAZYNOWE, $FLT_Z_FLAGA))

	$_UchwytyKontrolek.Add($LST_PRZYJECIA_MAGAZYNOWE, ObjCreate("Scripting.Dictionary"))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).add( $HWND_OKRES,zapiszUchwyt($LST_PRZYJECIA_MAGAZYNOWE, $FLT_DOKUMENTY_Z_OKRESU))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).add( $HWND_TYPU,zapiszUchwyt($LST_PRZYJECIA_MAGAZYNOWE, $FLT_TYPU))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).add( $HWND_STAN,zapiszUchwyt($LST_PRZYJECIA_MAGAZYNOWE, $FLT_O_STANIE))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).add( $HWND_KATEGORIA,zapiszUchwyt($LST_PRZYJECIA_MAGAZYNOWE, $FLT_O_KATEGORII_NP))
	$_UchwytyKontrolek.Item($LST_PRZYJECIA_MAGAZYNOWE).add( $HWND_FLAGA,zapiszUchwyt($LST_PRZYJECIA_MAGAZYNOWE, $FLT_Z_FLAGA))

	$_UchwytyKontrolek.Add($LST_ZWROTY_DETALICZNE, ObjCreate("Scripting.Dictionary"))
	$_UchwytyKontrolek.Item($LST_ZWROTY_DETALICZNE).add( $HWND_OKRES,zapiszUchwyt($LST_ZWROTY_DETALICZNE, $FLT_DOKUMENTY_Z_OKRESU))
	$_UchwytyKontrolek.Item($LST_ZWROTY_DETALICZNE).add( $HWND_RODZAJ,zapiszUchwyt($LST_ZWROTY_DETALICZNE, $FLT_RODZAJ_ZWROTU))
	$_UchwytyKontrolek.Item($LST_ZWROTY_DETALICZNE).add( $HWND_KATEGORIA,zapiszUchwyt($LST_ZWROTY_DETALICZNE, $FLT_O_KATEGORII_NS))
	$_UchwytyKontrolek.Item($LST_ZWROTY_DETALICZNE).add( $HWND_FLAGA,zapiszUchwyt($LST_ZWROTY_DETALICZNE, $FLT_Z_FLAGA_NS))

	;------magazyn------
	$_UchwytyKontrolek.Add($CTR_MAGAZYN,pobierzUchwytMagazynow())



;~ 	$uchwyt2 = HWnd($_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA & $FLT_DOKUMENTY_Z_OKRESU))
;~ 	ControlCommand("", "", $uchwyt2, "SetCurrentSelection", 3)
;~ 	ControlCommand("", "", $uchwyt2, "GetCurrentSelection", "")

EndFunc   ;==>pobierzUchwytyFiltrow



Func wybierzFiltr($lista,$filtr,$pozycja = "")
	If $pozycja = "" Then
		ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($lista).Item($filtr)), "SetCurrentSelection", $pozycja)
	Else
		ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($lista).Item($filtr)), "SelectString", $pozycja)
	EndIf
EndFunc

Func wybierzMagazyn($magazyn)
	For $i = 0 to UBound($_ListaMagazynowBot) - 1
		If StringLeft($_ListaMagazynowBot[$i],3) = $magazyn Then
			ControlCommand("", "", HWnd($_UchwytyKontrolek.Item($CTR_MAGAZYN)), "SetCurrentSelection", $i)
			ExitLoop
		EndIf
	Next
EndFunc

Func pobierzListeMagazynow()
	$uchwyt = HWnd($_UchwytyKontrolek.Item($CTR_MAGAZYN))
	$count = _GUICtrlListBox_GetCount($uchwyt)
	Local $magazyn[$count]
	For $i = 0 To $count - 1
		$magazyn[$i] = _GUICtrlListBox_GetText($uchwyt, $i)
	Next
	Return $magazyn
EndFunc


Func pobierzUchwytMagazynow()
	$uchwyt = szukajUchwytu(MWinGetHandle(),$CTR_MAGAZYN, $CLS_INS_COMBOHYPERLINK)
	ControlFocus($uchwyt, "", "")
	ControlClick($uchwyt, "", "")
	$uchwyt = WinGetHandle("[REGEXPCLASS:(Afx:.*.:800)]")
	$uchwyt2 = _WinAPI_GetWindow($uchwyt, $GW_CHILD)
	If _WinAPI_GetClassName($uchwyt2) = "ListBox" And _WinAPI_GetDlgCtrlID($uchwyt2) = 2001 Then
		Send("{ESC}")
		Return $uchwyt2
	EndIf
	Return 0
EndFunc

Func usunSeparatory($uchwyt)
	$count = _GUICtrlListBox_GetCount($uchwyt)
	For $i = 0 To $count - 1
		If _GUICtrlListBox_GetText($uchwyt, $i) = "" Then
			ControlCommand("", "", $uchwyt, "DelString", $i)
		EndIf
	Next
EndFunc   ;==>usunSeparatory


Func zapiszUchwyt($lista, $filtr)
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
EndFunc   ;==>zapiszUchwyt


Func otworzFiltr($lista, $filtr)
	$hWnd = MWinGetHandle()
	$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $lista, $CLS_INS_HYPERLINK, $filtr, $CLS_STATIC)
	$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	ControlFocus($uchwyt, "", "")
	ControlClick($uchwyt, "", "")
	Return $uchwyt
EndFunc   ;==>otworzFiltr




