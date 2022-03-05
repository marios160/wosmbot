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
;------LISTY-----
Global $LST_SPRZEDAZ_DETALICZNA = "Sprzedaż detaliczna"
Global $LST_WYDANIA_MAGAZYNOWE = "Wydania magazynowe"
Global $LST_FAKTURY_ZAKUPU = "Faktury zakupu"
Global $LST_KOREKTY_ZAKUPU = "Korekty zakupu"
Global $LST_PRZYJECIA_MAGAZYNOWE = "Przyjęcia magazynowe"
Global $LST_ZWROTY_DETALICZNE = "Zwroty detaliczne"
;-----LISTY SKROTY-----
Global $PA = "PA"
Global $PW = "PW"
Global $MM = "MM"
Global $ZW = "ZW"
Global $FA = "FA"
Global $RW = "RW"
Global $KFZ = "KFZ"
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
Global $FLT_TYPU =  ", typu:"
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

pobierzUchwytyFiltrow()

Func pobierzUchwytyFiltrow()

ControlFocus("","",MWinGetHandle())

wybierzListe($LST_FAKTURY_ZAKUPU)
wybierzListe($LST_WYDANIA_MAGAZYNOWE)
wybierzListe($LST_PRZYJECIA_MAGAZYNOWE)
wybierzListe($LST_ZWROTY_DETALICZNE)
wybierzListe($LST_SPRZEDAZ_DETALICZNA)

$_UchwytyKontrolek.Add($LST_SPRZEDAZ_DETALICZNA&$FLT_DOKUMENTY_Z_OKRESU,zapiszUchwyt($LST_SPRZEDAZ_DETALICZNA,$FLT_DOKUMENTY_Z_OKRESU))
$_UchwytyKontrolek.Add($LST_SPRZEDAZ_DETALICZNA&$FLT_TYPU,zapiszUchwyt($LST_SPRZEDAZ_DETALICZNA,$FLT_TYPU))
$_UchwytyKontrolek.Add($LST_SPRZEDAZ_DETALICZNA&$FLT_O_KATEGORII_NS,zapiszUchwyt($LST_SPRZEDAZ_DETALICZNA,$FLT_O_KATEGORII_NS))
$_UchwytyKontrolek.Add($LST_SPRZEDAZ_DETALICZNA&$FLT_Z_FLAGA,zapiszUchwyt($LST_SPRZEDAZ_DETALICZNA,$FLT_Z_FLAGA))

$_UchwytyKontrolek.Add($LST_FAKTURY_ZAKUPU&$FLT_DOKUMENTY_WG,zapiszUchwyt($LST_FAKTURY_ZAKUPU,$FLT_DOKUMENTY_WG))
$_UchwytyKontrolek.Add($LST_FAKTURY_ZAKUPU&$FLT_Z_OKRESU,zapiszUchwyt($LST_FAKTURY_ZAKUPU,$FLT_Z_OKRESU))
$_UchwytyKontrolek.Add($LST_FAKTURY_ZAKUPU&$FLT_TYPU,zapiszUchwyt($LST_FAKTURY_ZAKUPU,$FLT_TYPU))
$_UchwytyKontrolek.Add($LST_FAKTURY_ZAKUPU&$FLT_O_STANIE_NP,zapiszUchwyt($LST_FAKTURY_ZAKUPU,$FLT_O_STANIE_NP))
$_UchwytyKontrolek.Add($LST_FAKTURY_ZAKUPU&$FLT_O_KATEGORII,zapiszUchwyt($LST_FAKTURY_ZAKUPU,$FLT_O_KATEGORII))
$_UchwytyKontrolek.Add($LST_FAKTURY_ZAKUPU&$FLT_Z_FLAGA,zapiszUchwyt($LST_FAKTURY_ZAKUPU,$FLT_Z_FLAGA))

$_UchwytyKontrolek.Add($LST_WYDANIA_MAGAZYNOWE&$FLT_DOKUMENTY_Z_OKRESU,zapiszUchwyt($LST_WYDANIA_MAGAZYNOWE,$FLT_DOKUMENTY_Z_OKRESU))
$_UchwytyKontrolek.Add($LST_WYDANIA_MAGAZYNOWE&$FLT_TYPU,zapiszUchwyt($LST_WYDANIA_MAGAZYNOWE,$FLT_TYPU))
$_UchwytyKontrolek.Add($LST_WYDANIA_MAGAZYNOWE&$FLT_O_STANIE,zapiszUchwyt($LST_WYDANIA_MAGAZYNOWE,$FLT_O_STANIE))
$_UchwytyKontrolek.Add($LST_WYDANIA_MAGAZYNOWE&$FLT_O_KATEGORII_NP,zapiszUchwyt($LST_WYDANIA_MAGAZYNOWE,$FLT_O_KATEGORII_NP))
$_UchwytyKontrolek.Add($LST_WYDANIA_MAGAZYNOWE&$FLT_Z_FLAGA,zapiszUchwyt($LST_WYDANIA_MAGAZYNOWE,$FLT_Z_FLAGA))

$_UchwytyKontrolek.Add($LST_PRZYJECIA_MAGAZYNOWE&$FLT_DOKUMENTY_Z_OKRESU,zapiszUchwyt($LST_PRZYJECIA_MAGAZYNOWE,$FLT_DOKUMENTY_Z_OKRESU))
$_UchwytyKontrolek.Add($LST_PRZYJECIA_MAGAZYNOWE&$FLT_TYPU,zapiszUchwyt($LST_PRZYJECIA_MAGAZYNOWE,$FLT_TYPU))
$_UchwytyKontrolek.Add($LST_PRZYJECIA_MAGAZYNOWE&$FLT_O_STANIE,zapiszUchwyt($LST_PRZYJECIA_MAGAZYNOWE,$FLT_O_STANIE))
$_UchwytyKontrolek.Add($LST_PRZYJECIA_MAGAZYNOWE&$FLT_O_KATEGORII_NP,zapiszUchwyt($LST_PRZYJECIA_MAGAZYNOWE,$FLT_O_KATEGORII_NP))
$_UchwytyKontrolek.Add($LST_PRZYJECIA_MAGAZYNOWE&$FLT_Z_FLAGA,zapiszUchwyt($LST_PRZYJECIA_MAGAZYNOWE,$FLT_Z_FLAGA))

$_UchwytyKontrolek.Add($LST_ZWROTY_DETALICZNE&$FLT_DOKUMENTY_Z_OKRESU,zapiszUchwyt($LST_ZWROTY_DETALICZNE,$FLT_DOKUMENTY_Z_OKRESU))
$_UchwytyKontrolek.Add($LST_ZWROTY_DETALICZNE&$FLT_RODZAJ_ZWROTU,zapiszUchwyt($LST_ZWROTY_DETALICZNE,$FLT_RODZAJ_ZWROTU))
$_UchwytyKontrolek.Add($LST_ZWROTY_DETALICZNE&$FLT_O_KATEGORII_NS,zapiszUchwyt($LST_ZWROTY_DETALICZNE,$FLT_O_KATEGORII_NS))
$_UchwytyKontrolek.Add($LST_ZWROTY_DETALICZNE&$FLT_Z_FLAGA_NS,zapiszUchwyt($LST_ZWROTY_DETALICZNE,$FLT_Z_FLAGA_NS))


$uchwyt2 = HWnd($_UchwytyKontrolek.Item($LST_SPRZEDAZ_DETALICZNA&$FLT_DOKUMENTY_Z_OKRESU))
ControlCommand("","",$uchwyt2,"SetCurrentSelection",3)
ControlCommand("","",$uchwyt2,"GetCurrentSelection","")

EndFunc


Func usunSeparatory($uchwyt)
	$count = _GUICtrlListBox_GetCount($uchwyt)
	For $i = 0 to $count - 1
		If _GUICtrlListBox_GetText($uchwyt,$i) = "" Then
			ControlCommand("","",$uchwyt,"DelString",$i)
		EndIf
	Next
EndFunc




Func zapiszUchwytXXX($lista,$filtr)
	otworzFiltr($lista,$filtr)
	$uchwyt = WinGetHandle("[REGEXPCLASS:(Afx:.*.:800)]")
	While $uchwyt <> 0
		$uchwyt2 = _WinAPI_GetWindow($uchwyt, $GW_CHILD)
		If _WinAPI_GetClassName($uchwyt2) = "ListBox" And _WinAPI_GetDlgCtrlID($uchwyt2) = 2001 Then
			ConsoleWrite("__ " & $uchwyt2 & " " & _WinAPI_GetClassName($uchwyt2) & " " & _WinAPI_GetDlgCtrlID($uchwyt2) & @LF)
			ControlCommand("","",$uchwyt2,"SetCurrentSelection",0)
			If @error Then ConsoleWrite(@error & @lf)
			ConsoleWrite(ControlCommand("","",$uchwyt2,"GetCurrentSelection","") & @LF & @LF)
		EndIf
		$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	WEnd
EndFunc   ;==>czySubiektUruchomiony2


Func zapiszUchwyt($lista,$filtr)
	otworzFiltr($lista,$filtr)
	$uchwyt = WinGetHandle("[REGEXPCLASS:(Afx:.*.:800)]")
	$uchwyt2 = _WinAPI_GetWindow($uchwyt, $GW_CHILD)
	If _WinAPI_GetClassName($uchwyt2) = "ListBox" And _WinAPI_GetDlgCtrlID($uchwyt2) = 2001 Then
		usunSeparatory($uchwyt2)
		Send("{ESC}")
		Return $uchwyt2
	EndIf
	Return 0
EndFunc   ;==>


Func otworzFiltr($lista,$filtr)
	$hWnd = MWinGetHandle()
	$uchwyt = pobierzUchwytKontrolkiListyPetla($hWnd, $lista, $CLS_INS_HYPERLINK, $filtr, $CLS_STATIC)
	$uchwyt = _WinAPI_GetWindow($uchwyt, $GW_HWNDNEXT)
	ControlFocus($uchwyt, "", "")
	ControlClick($uchwyt, "", "")
	Return $uchwyt
EndFunc

