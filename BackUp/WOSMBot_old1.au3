#include <ButtonConstants.au3>
#include <DateTimeConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListBox.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <_sql.au3>
Opt("GUIOnEventMode", 1)
#Region ### START Koda GUI section ### Form=C:\Users\Mateusz Błaszczak\Desktop\Bramka\()AUTOIT\GUI\Form1_1.kxf
$Form1 = GUICreate("Form1", 791, 590, -1008, 125)
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
$dataOd = GUICtrlCreateDate("2018/08/01 10:10:35", 48, 104, 162, 21)
GUICtrlSetOnEvent($dataOd, "dataOdChange")
$dataDo = GUICtrlCreateDate("2018/08/01 10:10:39", 48, 136, 162, 21)
GUICtrlSetOnEvent($dataDo, "dataDoChange")
$Label1 = GUICtrlCreateLabel("Od:", 24, 104, 21, 17)
GUICtrlSetOnEvent($Label1, "Label1Click")
$Label2 = GUICtrlCreateLabel("Do:", 24, 136, 21, 17)
GUICtrlSetOnEvent($Label2, "Label2Click")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$Group3 = GUICtrlCreateGroup("Dokumenty", 8, 184, 225, 193)
$chckSprzedazDetaliczna = GUICtrlCreateCheckbox("Sprzedaż detaliczna", 24, 208, 121, 17)
GUICtrlSetOnEvent($chckSprzedazDetaliczna, "chckSprzedazDetalicznaClick")
$chckWydMag = GUICtrlCreateCheckbox("Wydania magazynowe", 24, 232, 129, 17)
GUICtrlSetOnEvent($chckWydMag, "chckWydMagClick")
$chckWydMagRozchodWewnetrzny = GUICtrlCreateCheckbox("Rozchód wewnętrzny", 40, 256, 129, 17)
GUICtrlSetOnEvent($chckWydMagRozchodWewnetrzny, "chckWydMagRozchodWewnetrznyClick")
$chckWydMagPrzesuniecieMiedzymagazynowe = GUICtrlCreateCheckbox("Przesunięcie międzymagazynowe", 40, 280, 177, 17)
GUICtrlSetOnEvent($chckWydMagPrzesuniecieMiedzymagazynowe, "chckWydMagPrzesuniecieMiedzymagazynoweClick")
$chckPrzyjMag = GUICtrlCreateCheckbox("Przyjęcia magazynowe", 24, 304, 129, 17)
GUICtrlSetOnEvent($chckPrzyjMag, "chckPrzyjMagClick")
$chckPrzyjMagPrzychodWewnetrzny = GUICtrlCreateCheckbox("Przychód wewnętrzny", 40, 328, 153, 17)
GUICtrlSetOnEvent($chckPrzyjMagPrzychodWewnetrzny, "chckPrzyjMagPrzychodWewnetrznyClick")
$chckZwrotyDetaliczne = GUICtrlCreateCheckbox("Zwroty detaliczne", 24, 352, 105, 17)
GUICtrlSetOnEvent($chckZwrotyDetaliczne, "chckZwrotyDetaliczneClick")
GUICtrlCreateGroup("", -99, -99, 1, 1)
$lstKolejnosc = GUICtrlCreateList("", 310, 32, 300, 150, $LBS_NOTIFY)
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
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

$_ListaDokumentow = ""

Local $ServerAddress = "XERO\INSERTGT"
Local $ServerUserName = "sa"
Local $ServerPassword = ""
Local $DatabaseName = "Arctus_Test_BotSM"


_SQL_RegisterErrorHandler();register the error handler to prevent hard crash on COM error
$OADODB = _SQL_Startup()
If $OADODB = $SQL_ERROR Then MsgBox(0 + 16 + 262144, "Error", _SQL_GetErrMsg())
If _sql_Connect(-1, $ServerAddress, $DatabaseName, $ServerUserName, $ServerPassword) = $SQL_ERROR Then
MsgBox(0 + 16 + 262144, "Error 1", _SQL_GetErrMsg())
_SQL_Close()
Exit
EndIf
Local $fullSQL
If _Sql_GetTableAsString(-1, "SELECT [mag_Symbol],[mag_Nazwa] FROM [Arctus_Test_BotSM].[dbo].[sl_Magazyn]", $fullSQL) = $SQL_OK Then
ConsoleWrite($fullSQL)
Else
MsgBox(0 + 16 + 262144, "SQL Error", _SQL_GetErrMsg())
EndIf



While 1
	Sleep(100)
WEnd





Exit

Func btnWGoreClick()
	$item = _GUICtrlListBox_GetCurSel($lstKolejnosc)
	$items = StringSplit($_ListaDokumentow,"|", $STR_NOCOUNT)
	If $item > 0 Then
		$_ListaDokumentow = ""
		$tmp = $items[$item]
		$items[$item] = $items[$item - 1]
		$items[$item - 1] = $tmp
		For	$i = 0 To UBound($items) - 2
			$_ListaDokumentow = $_ListaDokumentow & $items[$i] & "|"
		Next
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		_GUICtrlListBox_SelectString($lstKolejnosc,$tmp, $item - 1)
	EndIf
EndFunc

Func btnWDolClick()
	$item = _GUICtrlListBox_GetCurSel($lstKolejnosc)
	$items = StringSplit($_ListaDokumentow,"|", $STR_NOCOUNT)
	If $item < UBound($items) - 2 Then
		$_ListaDokumentow = ""
		$tmp = $items[$item]
		$items[$item] = $items[$item + 1]
		$items[$item + 1] = $tmp
		For	$i = 0 To UBound($items) - 2
			$_ListaDokumentow = $_ListaDokumentow & $items[$i] & "|"
		Next
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		_GUICtrlListBox_SelectString($lstKolejnosc,$tmp, $item - 1)
	EndIf
EndFunc

Func chckPrzyjMagClick()
	If GUICtrlRead($chckPrzyjMag) = $GUI_CHECKED Then
		GUICtrlSetState($chckPrzyjMagPrzychodWewnetrzny,$GUI_CHECKED)
		chckPrzyjMagPrzychodWewnetrznyClick()
	Else
		GUICtrlSetState($chckPrzyjMagPrzychodWewnetrzny,$GUI_UNCHECKED)
		chckPrzyjMagPrzychodWewnetrznyClick()
	EndIf
EndFunc

Func chckPrzyjMagPrzychodWewnetrznyClick()
	If GUICtrlRead($chckPrzyjMagPrzychodWewnetrzny) = $GUI_CHECKED Then
		$_ListaDokumentow = $_ListaDokumentow & "Przchód wewnętrzny (Przyjęcia magazynowe)|"
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetState($chckPrzyjMag,$GUI_CHECKED)
	Else
		$_ListaDokumentow = StringReplace($_ListaDokumentow,"Przchód wewnętrzny (Przyjęcia magazynowe)|","")
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetState($chckPrzyjMag,$GUI_UNCHECKED)
	EndIf
EndFunc

Func chckSprzedazDetalicznaClick()
	If GUICtrlRead($chckSprzedazDetaliczna) = $GUI_CHECKED Then
		$_ListaDokumentow = $_ListaDokumentow &"Sprzedaż detaliczna|"
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
	Else
		$_ListaDokumentow = StringReplace($_ListaDokumentow,"Sprzedaż detaliczna|","")
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
	EndIf
EndFunc

Func chckWydMagClick()
	If GUICtrlRead($chckWydMag) = $GUI_CHECKED Then
		GUICtrlSetState($chckWydMagRozchodWewnetrzny,$GUI_CHECKED)
		GUICtrlSetState($chckWydMagPrzesuniecieMiedzymagazynowe,$GUI_CHECKED)
		chckWydMagPrzesuniecieMiedzymagazynoweClick()
		chckWydMagRozchodWewnetrznyClick()
	Else
		GUICtrlSetState($chckWydMagRozchodWewnetrzny,$GUI_UNCHECKED)
		GUICtrlSetState($chckWydMagPrzesuniecieMiedzymagazynowe,$GUI_UNCHECKED)
		chckWydMagPrzesuniecieMiedzymagazynoweClick()
		chckWydMagRozchodWewnetrznyClick()
	EndIf
EndFunc

Func chckWydMagPrzesuniecieMiedzymagazynoweClick()
	If GUICtrlRead($chckWydMagPrzesuniecieMiedzymagazynowe) = $GUI_CHECKED Then
		$_ListaDokumentow = $_ListaDokumentow & "Przesunięcie międzymagazynowe (Wydania magazynowe)|"
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetState($chckWydMag,$GUI_CHECKED)
	Else
		$_ListaDokumentow = StringReplace($_ListaDokumentow,"Przesunięcie międzymagazynowe (Wydania magazynowe)|","")
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)

		If GUICtrlRead($chckWydMagRozchodWewnetrzny) = $GUI_UNCHECKED Then
			GUICtrlSetState($chckWydMag,$GUI_UNCHECKED)
		EndIf
	EndIf
EndFunc

Func chckWydMagRozchodWewnetrznyClick()
	If GUICtrlRead($chckWydMagRozchodWewnetrzny) = $GUI_CHECKED Then
		$_ListaDokumentow = $_ListaDokumentow & "Rozchód wewnętrzny (Wydania magazynowe)|"
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
		GUICtrlSetState($chckWydMag,$GUI_CHECKED)
	Else
		$_ListaDokumentow = StringReplace($_ListaDokumentow,"Rozchód wewnętrzny (Wydania magazynowe)|","")
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)

		If GUICtrlRead($chckWydMagPrzesuniecieMiedzymagazynowe) = $GUI_UNCHECKED Then
			GUICtrlSetState($chckWydMag,$GUI_UNCHECKED)
		EndIf

	EndIf
EndFunc

Func chckZwrotyDetaliczneClick()
	If GUICtrlRead($chckZwrotyDetaliczne) = $GUI_CHECKED Then
		$_ListaDokumentow = $_ListaDokumentow & "Zwroty detaliczne|"
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
	Else
		$_ListaDokumentow = StringReplace($_ListaDokumentow,"Zwroty detaliczne|","")
		GUICtrlSetData($lstKolejnosc, "")
		GUICtrlSetData($lstKolejnosc, $_ListaDokumentow)
	EndIf

EndFunc

Func dataDoChange()

EndFunc

Func dataOdChange()

EndFunc

Func Form1Close()
	GUIDelete()
	Exit
EndFunc

Func Form1Maximize()

EndFunc
Func Form1Minimize()

EndFunc
Func Form1Restore()

EndFunc
Func Label1Click()

EndFunc
Func Label2Click()

EndFunc
Func Label3Click()

EndFunc
Func lstKolejnoscClick()

EndFunc
Func radioOdkladanieClick()

EndFunc
Func radioWywolywanieClick()

EndFunc


