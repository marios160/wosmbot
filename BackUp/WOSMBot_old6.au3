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
#include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiComboBoxEx.au3>
#include <GuiComboBox.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <ComboConstants.au3>
#include <APISubiektWin.au3>
#include <WOSMBot_GuiLog.au3>
#include <WOSMBot_lstKolejnosc.au3>
#include <WOSMBot_DaneWejsciowe.au3>
;#include <test.au3>
#include <BotOSM.au3>
;#include <BotWSM.au3>
Opt("GUIOnEventMode", 1) ;zezwolenie na działanie funkcji onEvent
Opt("GUICloseOnESC", 1) ;po wcisnieciu ESC gui sie zamyka




;#############################################################################################################################
;#############################################################################################################################
;##############################################		USTAWIENIE GUI		######################################################
;#############################################################################################################################
;#############################################################################################################################

#Region ### START Koda GUI section ### Form=C:\Users\Mateusz Błaszczak\Desktop\Bramka\()AUTOIT\GUI\Form1_1.kxf
$WOSMBot = GUICreate("Bot do wywoływania/odkładania skutków magazynowych", 662, 530, -1, -1)
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
$dataOd = GUICtrlCreateDate("2018/08/11 10:10:35", 48, 104, 162, 21, $DTS_SHORTDATEFORMAT)
GUICtrlSetOnEvent($dataOd, "dataOdChange")
$dataDo = GUICtrlCreateDate("2018/06/11 10:10:39", 48, 136, 162, 21, $DTS_SHORTDATEFORMAT)
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
$lstKolejnosc = GUICtrlCreateList("", 310, 32, 340, 448, BitOR($LBS_NOTIFY, $WS_HSCROLL, $WS_VSCROLL, $WS_BORDER))
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
$chckZaznaczWszystkie = GUICtrlCreateCheckbox("Zaznacz wszystkie", 24, 448, 137, 17)
GUICtrlSetFont(-1, 8, 800, 0, "MS Sans Serif")
GUICtrlSetOnEvent(-1, "chckZaznaczWszystkieClick")
GUICtrlCreateGroup("", -99, -99, 1, 1)
#EndRegion ### END Koda GUI section ###

;#############################################################################################################################
;#############################################################################################################################
;##############################################		ZMIENNE GLOBALNE	######################################################
;#############################################################################################################################
;#############################################################################################################################

;Tablica zawiera uchwyty do checkboksow z wyborem listy do wykonywania
Global $checkBoxes[8] = [$chckSprzedazDetaliczna, $chckWydMag, $chckWydMagRozchodWewnetrzny, _
		$chckWydMagPrzesuniecieMiedzymagazynowe, $chckPrzyjMag, _
		$chckPrzyjMagPrzychodWewnetrzny, $chckZwrotyDetaliczne, $chckZaznaczWszystkie]

Global $_Magazyny[0][0] ;stan checkboksow list dla kazdego magazynu
Global $aktMag = 0 ;numer aktualnie wybranego magazynu w comboboxie
Global $aktMagTxt = "" ;nazwa aktualnie wybranego magazynu w comboboxie
Global $_WybraneListyDoWykonania[0] ;Tablica zawiera wybrane listy dokumentow do przetworzenia
Global $_WybraneListyDoWykonaniaBot[0] ;Tablica zawiera wybrane listy dokumentow do przetworzenia zmienione na skrocone nazwy dokumentow
Global $_ListaMagazynow = ObjCreate("Scripting.Dictionary") ;Zawiera liste dostepnych magazynow; symbol , pelna nazwa
Global $_OdkladanieCzyWywolywanie = 0 ;Czy bot ma wykonywać odkladanie skutkow(0) czy wywolywanie(1)
Global $_DataOd = "" ;Data OD ktorego dnia ma sie wykonywac skrypt
Global $_DataDo = "" ;Data DO ktorego dnia ma sie wykonywac skrypt
Global $_Podmiot = "" ;Nazwa podmiotu
Global $_SciezkaUruchomieniaSubiekta = "" ;sciezka do programu subiekta ustawiana w funkcji sprawdzCzyJestSubiekt
Global $_nazwaOknaSubiekta = "" ;nazwa okna programu subiekta ustawiana w funkcji utworzPlikKonfiguracyjny
Global $GuiOczekiwanie = 0 ;uchwyt okienka oczekiwania

;#############################################################################################################################
;#############################################################################################################################
;##############################################	DEFINICJA ZMIENNNYCH	######################################################
;#############################################################################################################################
;#############################################################################################################################

uruchomZamknacSubiekta()
utworzPlikKonfiguracyjny() ;tworzenie pliku konfiguracyjnego
sprawdzCzyJestSubiekt() ;sprawdzanie czy program subiekt istnieje
przygotowanieGUI() ;pobieranie magazynow i ustawianie w GUI

GUISetState(@SW_HIDE, $GuiOczekiwanie)
GUISetState(@SW_SHOW, $WOSMBot) ;Wyswietlenie GUI



;#############################################################################################################################
;#############################################################################################################################
;##############################################		PETLA PROGRAMOWA	######################################################
;#############################################################################################################################
;#############################################################################################################################

;~ GUICtrlSetState($chckZaznaczWszystkie, $GUI_CHECKED)
;~ chckZaznaczWszystkieClick()
While 1
	Sleep(100)
WEnd


;#############################################################################################################################
;#############################################################################################################################
;##############################################			FUNKCJE			######################################################
;#############################################################################################################################
;#############################################################################################################################

;klikniecie przycisku start - uruchamia skrypt bota
;PA - paragony
;FA - faktury i rachunki zakupu
;MM - przesuniecie miedzymagazynowe - wydania magazynowe
;RW - rozchod wewnetrzny - wydania magazynowe
;PW - przychod wewnetrzny - przyjecia magazynowe
;ZW - zwroty detaliczne
Func btnStartClick()
	$_DataOd = GUICtrlRead($dataOd) ;pobieranie daty z kontrolek
	$_DataDo = GUICtrlRead($dataDo) ;pobieranie daty z kontrolek
	ReDim $_WybraneListyDoWykonaniaBot[UBound($_WybraneListyDoWykonania)] ;ustawianie rozmiarow tablicy
	For $i = 0 To UBound($_WybraneListyDoWykonania) - 1
		$el = StringSplit($_WybraneListyDoWykonania[$i], "[", 2) ;wyciaganie nazw list (przyklad: sprzedaz detaliczna [MAG])
		$_WybraneListyDoWykonaniaBot[$i] = StringStripWS($el[0], BitOR($STR_STRIPLEADING, $STR_STRIPTRAILING))
	Next

	While czySubiektUruchomiony($_nazwaOknaSubiekta) <> 0 ;Czekamy az subiekt sie zaladuje
		Sleep(1000)
	WEnd
	Sleep(2000)
	Local $daneBot[3] = [$_Podmiot, $_DataOd, $_DataDo] ;tablica ze zmiennymi dla bota
	If GUICtrlRead($radioOdkladanie) = $GUI_CHECKED Then ;odkladanie czy wywolywanie
		uruchomLog() ;uruchomienie loga dla bota (WOSMBot_GuiLog.au3)
		GUISetState(@SW_HIDE, $WOSMBot) ;ukrywanie okienka
		startOSMBot($daneBot, $_WybraneListyDoWykonaniaBot, $_ListaMagazynow)			;uruchomienie bota do odkladania(BotOSM.au3)
	Else
		;startWSMBot($daneBot,$_WybraneListyDoWykonaniaBot, $_ListaMagazynow)			;uruchomienie bota do wywolywania(BotWSM.au3)
	EndIf
EndFunc   ;==>btnStartClick

;-----------------------------------------------------------------------------------------------------------------------------
;zmiana magazynu
Func cmbMagazynChange()
	$aktMag = _GUICtrlComboBox_GetCurSel($cmbMagazyn) ;ustawienie numeru aktualnego magazynu
	$aktMagTxt = StringLeft(GUICtrlRead($cmbMagazyn), 3) ;ustawienie nazwy aktualnego magazynu
	For $i = 0 To 7
		GUICtrlSetState($checkBoxes[$i], $_Magazyny[$aktMag][$i]) ;ustawiamy checkboxy tak jak byly zapisane wczesniej
	Next
EndFunc   ;==>cmbMagazynChange

;-----------------------------------------------------------------------------------------------------------------------------
;Czyszczenie listy wybranych dokumentow
Func btnWyczyscClick()
	ReDim $_WybraneListyDoWykonania[0]
	GUICtrlSetData($lstKolejnosc, $_WybraneListyDoWykonania)
	For $j = 0 To UBound($_Magazyny) - 1
		For $i = 0 To 7
			$_Magazyny[$j][$i] = $GUI_UNCHECKED
			GUICtrlSetState($checkBoxes[$i], $_Magazyny[$j][$i])
		Next
	Next
EndFunc   ;==>btnWyczyscClick

;-----------------------------------------------------------------------------------------------------------------------------
;ustawianie kolejnosci wykonywania dokumentow, przesuwanie w gore
Func btnWGoreClick()
	$item = _GUICtrlListBox_GetCurSel($lstKolejnosc) ;wybrana lista
	If $item = -1 Then Return
	$selStr = $_WybraneListyDoWykonania[$item] ;pobieramy nazwe wybranej listy
	$_WybraneListyDoWykonania = moveUp($item, $_WybraneListyDoWykonania) ;przesuwamy w gore (WOSMBot_lstKolejnosc.au3)
	GUICtrlSetData($lstKolejnosc, "") ;czyscimy liste kolejnosci
	GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania)) ;dodajemy zaktualizowana liste kolejnosci
	$item = _ArraySearch($_WybraneListyDoWykonania, $selStr, 0, $item - 1, 0, 0, 0) ;szukamy wybranej wczesniej listy
	_GUICtrlListBox_SelectString($lstKolejnosc, $selStr, $item - 1) ;ustawiamy na niej fokus
EndFunc   ;==>btnWGoreClick

;-----------------------------------------------------------------------------------------------------------------------------
;ustawianie kolejnosci wykonywania dokumentow, przesuwanie w dol
Func btnWDolClick()
	$item = _GUICtrlListBox_GetCurSel($lstKolejnosc) ;dzialanie analogiczne do powyzszej funkcji (btnWGoreClick)
	If $item = -1 Then Return
	$selStr = $_WybraneListyDoWykonania[$item]
	$_WybraneListyDoWykonania = moveDown($item, $_WybraneListyDoWykonania)
	GUICtrlSetData($lstKolejnosc, "")
	GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
	$item = _ArraySearch($_WybraneListyDoWykonania, $selStr, $item)
	_GUICtrlListBox_SelectString($lstKolejnosc, $selStr, $item - 1)
EndFunc   ;==>btnWDolClick

;-----------------------------------------------------------------------------------------------------------------------------
;checkbox przyjecia magazynowe (ogolne)
Func chckPrzyjMagClick()
	If GUICtrlRead($chckPrzyjMag) = $GUI_CHECKED Then ;jesli zaznaczymy chck
		GUICtrlSetState($chckPrzyjMagPrzychodWewnetrzny, $GUI_CHECKED) ;zaznaczamy wszystkie podrzędne chck (tu jest tylko jeden)
		chckPrzyjMagPrzychodWewnetrznyClick() ;uruchamiamy ich funkcje klikniecia
	Else ;jesli odznaczymy
		GUICtrlSetState($chckPrzyjMagPrzychodWewnetrzny, $GUI_UNCHECKED) ;odznaczamy wszystkie podrzedne chck (tu jest tylko jeden)
		chckPrzyjMagPrzychodWewnetrznyClick() ;uruchamiamy ich funkcje klikniecia
	EndIf
EndFunc   ;==>chckPrzyjMagClick

;-----------------------------------------------------------------------------------------------------------------------------
;checkbox przyjecia magazynowe - przychod wewnetrzny
Func chckPrzyjMagPrzychodWewnetrznyClick()
	If GUICtrlRead($chckPrzyjMagPrzychodWewnetrzny) = $GUI_CHECKED Then ;jesli zaznaczymy chck dodajemy liste dokumentow do listy kolejnosci
		$_WybraneListyDoWykonania = addEl("     Przchód wewnętrzny (Przyjęcia magazynowe) [" & $aktMagTxt & "]", $_WybraneListyDoWykonania)
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
		GUICtrlSetState($chckPrzyjMag, $GUI_CHECKED) ;ustawiamy nadrzedny chck jako zaznaczony
		$_Magazyny[$aktMag][4] = 1 ;aktualizujemy tablice stanow chck dla obu chck
		$_Magazyny[$aktMag][5] = 1
	Else ;w przeciwnym razie to samo tylko odznaczamy i usuwamy
		$_WybraneListyDoWykonania = removeEl("     Przchód wewnętrzny (Przyjęcia magazynowe) [" & $aktMagTxt & "]", $_WybraneListyDoWykonania)
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
		GUICtrlSetState($chckPrzyjMag, $GUI_UNCHECKED)
		$_Magazyny[$aktMag][4] = 4
		$_Magazyny[$aktMag][5] = 4
	EndIf
EndFunc   ;==>chckPrzyjMagPrzychodWewnetrznyClick

;-----------------------------------------------------------------------------------------------------------------------------
;analogicznie jak powyzej tylko ta lista dokumentow nie ma podrzednych ani nadrzednych chck
Func chckSprzedazDetalicznaClick()
	If GUICtrlRead($chckSprzedazDetaliczna) = $GUI_CHECKED Then
		$_WybraneListyDoWykonania = addEl("     Sprzedaż detaliczna [" & $aktMagTxt & "]", $_WybraneListyDoWykonania)
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
		$_Magazyny[$aktMag][0] = 1
	Else
		$_WybraneListyDoWykonania = removeEl("     Sprzedaż detaliczna [" & $aktMagTxt & "]", $_WybraneListyDoWykonania)
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
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
		$_WybraneListyDoWykonania = addEl("     Przesunięcie międzymagazynowe (Wydania magazynowe) [" & $aktMagTxt & "]", $_WybraneListyDoWykonania)
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
		GUICtrlSetState($chckWydMag, $GUI_CHECKED)
		$_Magazyny[$aktMag][1] = 1
		$_Magazyny[$aktMag][3] = 1
	Else
		$_WybraneListyDoWykonania = removeEl("     Przesunięcie międzymagazynowe (Wydania magazynowe) [" & $aktMagTxt & "]", $_WybraneListyDoWykonania)
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
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
		$_WybraneListyDoWykonania = addEl("     Rozchód wewnętrzny (Wydania magazynowe) [" & $aktMagTxt & "]", $_WybraneListyDoWykonania)
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
		GUICtrlSetState($chckWydMag, $GUI_CHECKED)
		$_Magazyny[$aktMag][1] = 1
		$_Magazyny[$aktMag][2] = 1
	Else
		$_WybraneListyDoWykonania = removeEl("     Rozchód wewnętrzny (Wydania magazynowe) [" & $aktMagTxt & "]", $_WybraneListyDoWykonania)
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
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
		$_WybraneListyDoWykonania = addEl("     Zwroty detaliczne [" & $aktMagTxt & "]", $_WybraneListyDoWykonania)
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
		$_Magazyny[$aktMag][6] = 1
	Else
		$_WybraneListyDoWykonania = removeEl("     Zwroty detaliczne [" & $aktMagTxt & "]", $_WybraneListyDoWykonania)
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, toString($_WybraneListyDoWykonania))
		$_Magazyny[$aktMag][6] = 4
	EndIf
EndFunc   ;==>chckZwrotyDetaliczneClick

;-----------------------------------------------------------------------------------------------------------------------------
;zaznaczamy wszystkie chck i wykonujemy ich funkcje klikniecia
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
EndFunc   ;==>chckZaznaczWszystkieClick



;-----------------------------------------------------------------------------------------------------------------------------
;Tworzenie pliku konfiguracyjnego dla subiekta
Func utworzPlikKonfiguracyjny()
	$dane = uruchomOkienko() ;Okienko z danymi bazy danych (WOSMBot_DaneWejsciowe.au3)
	uruchomOkienkoOczekiwania()

	$ServerAddress = $dane[0] ;w wyniku zamkniecia okienka z danymi, pozyskujemy potrzebne dane
	$ServerUserName = $dane[1]
	$ServerPassword = $dane[2]
	$DatabaseName = $dane[3]
	$uzytkownik = $dane[4]
	$haslo = $dane[5]
	$_Podmiot = $dane[3]
	$_nazwaOknaSubiekta = $_Podmiot & " na serwerze " & $ServerAddress & " - Subiekt GT" ;ustawiamy nazwe okna



	;sprawdzamy czy dane zostaly wpisane
	If $ServerAddress = "" Or $ServerUserName = "" Or $DatabaseName = "" Then
		MsgBox(0, "Error", "Brak potrzebnych danych!")
		Exit
	EndIf

	;Tworzenie pliku konfiguracyjnego
	;dzieki temu otwieramy subiekta pozbawionego zakladek ktorych nie uzywamy
	$subiektxml = '<?xml version="1.0" encoding="windows-1250"?>' & @CR & @LF & _
			'<cfg>' & @CR & @LF & _
			'	<startup>'

	;wyszukiwanie aktualnej wersji subiekta w aktualnym pliku konfiguracyjnym
	$hF = FileOpen(EnvGet("ProgramData") & "\InsERT\InsERT GT\Subiekt.xml")
	$line = FileRead($hF)
	$wersja = StringMid($line, StringInStr($line, "<version>") + 9, StringInStr($line, "</version>") - StringInStr($line, "<version>") - 9)
	FileClose($hF)

	;dalsza czesc pliku konfiguracyjnego
	$subiektxml2 = '	</startup>' & @CR & @LF & _
			'	<menu_view>' & @CR & @LF & _
			'		<menu_popup image="0" label="&amp;Sprzedaż">' & @CR & @LF & _
			'			' & @CR & @LF & _
			'			<menu_item style="0" progid="Subiekt.WidokPA.1" accel_key="2" accel_modifier="1" params="" visible="1" label="Sprzedaż &amp;detaliczna"/>' & @CR & @LF & _
			'			<menu_item style="0" progid="Subiekt.WidokZW.1" params="" visible="1" label="&amp;Zwroty detaliczne"/>' & @CR & @LF & _
			'		</menu_popup>' & @CR & @LF & _
			'		<menu_popup image="1" label="&amp;Zakup">' & @CR & @LF & _
			'			<menu_item style="0" progid="Subiekt.WidokFZ.1" accel_key="3" accel_modifier="1" params="" visible="1" label="&amp;Faktury zakupu"/>' & @CR & @LF & _
			'		</menu_popup>' & @CR & @LF & _
			'		<menu_popup image="2" label="&amp;Magazyn">' & @CR & @LF & _
			'			<menu_item style="0" progid="Subiekt.WidokMagW.1" accel_key="4" accel_modifier="1" params="" visible="1" label="&amp;Wydania magazynowe"/>' & @CR & @LF & _
			'			<menu_item style="0" progid="Subiekt.WidokMagP.1" accel_key="5" accel_modifier="1" params="" visible="1" label="&amp;Przyjęcia magazynowe"/>' & @CR & @LF & _
			'		</menu_popup>' & @CR & @LF & _
			'	</menu_view>' & @CR & @LF & _
			'	<version>' & $wersja & '</version>' & @CR & @LF & _
			'</cfg>'

	;wklejanie tresci do pliku konfiguracyjnego
	$hFile = FileOpen("SubiektBot.xml", 2 + 512)
	FileWrite($hFile, $subiektxml & @CR & @LF)
	FileWrite($hFile, '		<sql_server>' & $ServerAddress & '</sql_server>' & @CR & @LF)
	FileWrite($hFile, '		<auth_mode>MIXED</auth_mode>' & @CR & @LF)
	FileWrite($hFile, '		<sql_login encrypted="0">' & $ServerUserName & '\' & $ServerPassword & '</sql_login>' & @CR & @LF)
	FileWrite($hFile, '		<database>' & $DatabaseName & '</database>' & @CR & @LF)
	FileWrite($hFile, '		<login encrypted="0">' & $uzytkownik & '\' & $haslo & '</login>' & @CR & @LF)
	FileWrite($hFile, $subiektxml2 & @CR & @LF)
	FileClose($hFile)
EndFunc   ;==>utworzPlikKonfiguracyjny

;-----------------------------------------------------------------------------------------------------------------------------
;sprawdzenie czy subiekt gt jest zainstalowany
Func sprawdzCzyJestSubiekt()
	;sprawdzamy pod dwoma sciezkami dla 32bit i 64bit
	If FileExists(EnvGet("ProgramFiles(x86)") & "\InsERT\InsERT GT\Subiekt.exe") Then
		$pathSubiekt = EnvGet("ProgramFiles(x86)") & "\InsERT\InsERT GT"
	ElseIf FileExists(EnvGet("ProgramFilesW6432") & "\InsERT\InsERT GT\Subiekt.exe") Then
		$pathSubiekt = EnvGet("ProgramFilesW6432") & "\InsERT\InsERT GT"
	Else
		Do
			$odp = MsgBox(1, "Nie znaleziono programu 'Subiekt GT'", "Wybierz folder, w którym zainstalowano program 'Subiekt.exe'")
			If $odp = 2 Then
				Exit
			EndIf
			;jesli nie znaleziono subiekta w standartowych sciezkach mozna wybrac sciezke w oknie wyboru plikow
			$pathSubiekt = FileSelectFolder("Wybierz folder gdzie zainstalowano 'Subiekt.exe'", EnvGet("ProgramFiles(x86)"))
		Until FileExists($pathSubiekt & "\subiekt.exe") = 1
	EndIf
	;ustawiamy sciezke do programu subiekt
	$_SciezkaUruchomieniaSubiekta = $pathSubiekt & '\Subiekt.exe "' & @ScriptDir & '\SubiektBot.xml"'
	Run($_SciezkaUruchomieniaSubiekta, "") ;uruchomienie subiekta
	While czySubiektUruchomiony($_nazwaOknaSubiekta) <> 0 ;Czekamy az subiekt sie zaladuje
		Sleep(1000)
	WEnd
	$_UchwytyKontrolek.Add($SUBIEKT,WinGetHandle($_nazwaOknaSubiekta))

EndFunc   ;==>sprawdzCzyJestSubiekt

;-----------------------------------------------------------------------------------------------------------------------------
;ustawianie danych startowych GUI
Func przygotowanieGUI()
	pobierzUchwytFiltrow()
	$magTmp = pobierzListeMagazynow() ;pobieramy liste dostepnych magazynow
	ReDim $_Magazyny[UBound($magTmp)][8]
	For $i = 0 To UBound($magTmp) - 1
		For $j = 0 To 7
			$_Magazyny[$i][$j] = 4
		Next
		If $magTmp[$i] <> "" Then $_ListaMagazynow.add(StringLeft($magTmp[$i], 3), $magTmp[$i])
	Next

	GUICtrlSetData($cmbMagazyn, _ArrayToString($magTmp)) ;ustawiamy liste dla wyboru magazynow
	_GUICtrlComboBox_SetCurSel($cmbMagazyn, 0)
	$aktMagTxt = StringLeft($magTmp[0], 3)
EndFunc   ;==>przygotowanieGUI


Func uruchomOkienkoOczekiwania()

	#Region ### START Koda GUI section ### Form=
	$GuiOczekiwanie = GUICreate("Przygotowanie do pracy", 346, 127, -1, -1, BitOR($WS_BORDER, $WS_POPUP), BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
	$Label1 = GUICtrlCreateLabel("Trwa przygotowywanie programu Subiekt GT do pracy.", 48, 48, 263, 17)
	$Label2 = GUICtrlCreateLabel("Proszę czekać...", 48, 64, 83, 17)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

EndFunc   ;==>uruchomOkienkoOczekiwania

Func uruchomZamknacSubiekta()
	If ProcessExists("Subiekt.exe") <> 0 Then
		#Region ### START Koda GUI section ### Form=
		$GuiZamknacSubiekta = GUICreate("Oczekiwanie na zamknięcie Subiekt GT", 346, 127, -1, -1, $WS_BORDER, BitOR($WS_EX_TOPMOST, $WS_EX_TOOLWINDOW))
		$Label1 = GUICtrlCreateLabel("Proszę zamknąć program Subiekt GT", 40, 40, 263, 40)
		GUICtrlSetFont($Label1, 12)
		GUISetState(@SW_SHOW)
		#EndRegion ### END Koda GUI section ###

		While ProcessExists("Subiekt.exe") <> 0
			Sleep(1000)
		WEnd
		GUISetState(@SW_HIDE, $GuiZamknacSubiekta)
	EndIf
EndFunc   ;==>uruchomZamknacSubiekta

;-----------------------------------------------------------------------------------------------------------------------------
;zamykanie programu po wcisnieciu krzyzyka
Func Form1Close()
	GUIDelete()
	Exit
EndFunc   ;==>Form1Close


;-----------------------------------------------------------------------------------------------------------------------------
Func edtLogiChange()
EndFunc   ;==>edtLogiChange

;-----------------------------------------------------------------------------------------------------------------------------
Func dataDoChange()
EndFunc   ;==>dataDoChange

;-----------------------------------------------------------------------------------------------------------------------------
Func dataOdChange()
EndFunc   ;==>dataOdChange

;-----------------------------------------------------------------------------------------------------------------------------
Func Form1Maximize()
EndFunc   ;==>Form1Maximize

;-----------------------------------------------------------------------------------------------------------------------------
Func Form1Minimize()
EndFunc   ;==>Form1Minimize

;-----------------------------------------------------------------------------------------------------------------------------
Func Form1Restore()
EndFunc   ;==>Form1Restore

;-----------------------------------------------------------------------------------------------------------------------------
Func Label1Click()
EndFunc   ;==>Label1Click

;-----------------------------------------------------------------------------------------------------------------------------
Func Label2Click()
EndFunc   ;==>Label2Click

;-----------------------------------------------------------------------------------------------------------------------------
Func Label3Click()
EndFunc   ;==>Label3Click

;-----------------------------------------------------------------------------------------------------------------------------
Func lstKolejnoscClick()
EndFunc   ;==>lstKolejnoscClick

;-----------------------------------------------------------------------------------------------------------------------------
Func radioOdkladanieClick()
EndFunc   ;==>radioOdkladanieClick

;-----------------------------------------------------------------------------------------------------------------------------
Func radioWywolywanieClick()
EndFunc   ;==>radioWywolywanieClick

;-----------------------------------------------------------------------------------------------------------------------------
