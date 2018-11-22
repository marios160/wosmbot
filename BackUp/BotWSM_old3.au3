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


Func startWSMBot($dane, $dokumenty, $_ListaMagazynow)
	logs('@@ (28) :(' & @MIN & ':' & @SEC & ') startWSMBot()' & @CR) ;### Function Trace
	HotKeySet("{F6}", "zamknijW")
	HotKeySet("{PAUSE}", "pauzaW")
	HotKeySet("{F4}", "pauzaSzybkaW")

;~ $odp = $IDNO
;~ While $odp == $IDNO
;~ 	$odp = MsgBox($MB_YESNOCANCEL, "Start", "Proszê w³¹czyæ na liscie zwrotow kolumne Dokument korygowany oraz na ka¿dej liœcie ustawiæ sortowanie od najm³odszych dokumentów do najstarszych!" & @LF & "Czy chcesz rozpocz¹æ pracê?")
;~ WEnd
;~ If $odp == $IDCANCEL Then
;~ 	Exit
;~ EndIf

;~ $odp = MsgBox($MB_YESNO, "Czyszczenie loga", "Czy chcesz wyczyscic 'logWSM.log'?")
;~ If $odp == $IDYES Then
;~ 	$hF = FileOpen("logWSM.log", 2)
;~ 	FileClose($hF)
;~ EndIf


;~ $odp = MsgBox($MB_YESNO, "Czyszczenie brakuj¹cych towarów", "Czy chcesz wyczyscic 'WSMbraktowaru.txt'?")
;~ If $odp == $IDYES Then
;~ 	$hF = FileOpen("WSMbraktowaru.txt", 2)
;~ 	FileClose($hF)
;~ EndIf
	;_______________________________________Inicjalizacja zmiennych______________________________________________
	Global $_PAUSE = False
	Global $_DATA = $dane[1]
	Global $_ENDDATA = $dane[2]
	Global $_PODMIOT = $dane[0]
	Global $listaBrakujacychTowarow = ObjCreate("Scripting.Dictionary")
	Global $iloscLiniiBrakow = 0
	Global $ileBrakow = 0
	Global $ileZnaleziono = 0

	If FileExists("WSMBrakTowaru.log") Then
		FileMove("WSMBrakTowaru.log","logi\WSMBrakTowaru_" & FileGetTime("WSMBrakTowaru.log",0,1) & ".log",8)
	EndIf
	$_ENDDATA = StringReplace(_DateAdd('d', 1, $_ENDDATA), "/", "-")
	;____________Petla programowa_______________________________________________________________________________
	logs("START BotWSM")
	While $_DATA <> $_ENDDATA
		$hData = FileOpen("_DATA.conf",$FO_OVERWRITE)
		FileWrite($hData,$_DATA & @LF)
		FileWrite($hData,$_ENDDATA & @LF)
		FileClose($hData)
		logs("**************************************************************************************")
		logs("*************************************" & $_DATA & " **************************************")
		logs("**************************************************************************************")
		$_AKTUALNY_STAN[2] = $_DATA
		For $i = 0 To UBound($dokumenty) - 1

			If StringLen($dokumenty[$i]) = 3 Then
				$_AKTUALNY_STAN[0] = $dokumenty[$i]
				wybierzMagazyn($_ListaMagazynow.Item($dokumenty[$i]))
			Else
				$_AKTUALNY_STAN[1] = $dokumenty[$i]
				Do
					If $_RESTART Then
						restartujSubiekta()
						wybierzMagazyn($_ListaMagazynow.Item($_AKTUALNY_STAN[0]))
					EndIf
					$listaBrakujacychTowarow = 0
					$listaBrakujacychTowarow = ObjCreate("Scripting.Dictionary")
					wykonujListeW($dokumenty[$i])
				Until $_RESTART = False
			EndIf
			Sleep(1000)
		Next
		$_DATA = StringReplace(_DateAdd('d', 1, $_DATA), "/", "-")
	WEnd
	logs("KONIEC BotWSM")
EndFunc   ;==>startWSMBot

;#####################################################################################################################
;#####################################################################################################################
;################################													##################################
;################################						FUNKCJE						##################################
;################################													##################################
;#####################################################################################################################
;#####################################################################################################################


;---------------------------------------------------------------------------------------------------------------------
;Funkcja wykonuje wszystkie instrukcje potrzebne do wywo³ania skutkow magazynowych wszystkich pozycji na konkretnej liscie
;Przyjmuje nazwê listy na ktorej dzia³amy
;

Func wykonujListeW($lista)
	logs('@@ (108) :(' & @MIN & ':' & @SEC & ') wykonujListeW()' & @CR) ;### Function Trace
	$listaPelna = $lista
	$lista = wybierzListe($lista)
	Sleep(500)
	ustawOkres($_DATA, $lista)
	Sleep(500)
	logs("------------------------------- START " & $lista & " ----------------------------------------")
	If $_PAUSE = True Then
		pauzaLoopW()
	EndIf
	If Not czyPusta($lista) Then
		fokusNaGXWND($lista)
		Do
			$ileBrakow = 0
			$ileZnaleziono = 0
			Sleep(500)
			sendGXWNDEndHard($lista)
			Do
				$endpos = kopiujWiersz($lista)
			Until StringLen($endpos) > 0
			;logs("ENDPOS: " & $endpos)
			$curpos = ""
			sendGXWNDHomeHard($lista)
			$rekord = ""
			While $endpos <> $curpos
				$czyDaSieOdlozyc = False
				Do
					$curpos = kopiujWiersz($lista)
				Until StringInStr($curpos, "magazynow") > 0 And $rekord <> getNumer($curpos)
				$rekord = getNumer($curpos)
				;logs($rekord)
				If StringInStr($curpos, $STR_WYWOLUJE_SKUTEK) = 0 Then
					While Not $czyDaSieOdlozyc
						Do
							$curpos2 = kopiujWiersz($lista)
						Until StringInStr($curpos2, "magazynow") > 0
						;logs("curpos2: " & $curpos2)
						$czyDaSieOdlozyc = wywolajSkutekMagazynowy($listaPelna)
						If $_RESTART Then Return
					WEnd
				EndIf
				Sleep(1)
				If $_PAUSE = True Then
					pauzaLoopW()
					ExitLoop
				EndIf
				sendGXWNDDown($lista)
			WEnd
			$listaKopia = kopiujListe($lista)
			If StringInStr($listaKopia, $STR_ODLOZONY_SKUTEK) <> 0 Then
				StringReplace($listaKopia, $STR_ODLOZONY_SKUTEK, "xxx")
				$ileZnaleziono = @extended
				If $ileZnaleziono <> $ileBrakow Then
					logs("OMINIETO PRODUKT! " & "; Brakuje: " & $ileBrakow & ", znaleziono: " & @extended)
				EndIf
			EndIf
			logs($ileZnaleziono & " " & $ileBrakow & @LF & $listaKopia)
		Until $ileZnaleziono = $ileBrakow
	EndIf
	logs("------------------------------- KONIEC " & $lista & " ----------------------------------------")
EndFunc   ;==>wykonujListeW

;---------------------------------------------------------------------------------------------------------------------
;Funkcja obsluguje okienko brakuj¹cego towaru. Zapisuje do pliku wszystkie pozycje z listy a nastepnie zamyka okienko
;przyjmuje parametr okreslajacy czy dany produkt juz zosta³ zapisany oraz nazwe przerabianej listy

Func zapiszBrakTowaru($lista)
	logs('@@ (175) :(' & @MIN & ':' & @SEC & ') zapiszBrakTowaru()' & @CR) ;### Function Trace
	$brak = WinGetHandle("[REGEXPTITLE:(?i)(.* - brak towaru w magazynie)]")
	$uchwyt = szukajUchwytu($brak, "", "GXWND")
	$paragon = ""
	$data = ""
	logs(ClipGet())
	$tmp = StringSplit(ClipGet(), "	")
	$data = $tmp[2]
	$paragon = $tmp[3]

	If Not $listaBrakujacychTowarow.Exists($paragon) Then
		$hFile = FileOpen("WSMBrakTowaru.log", 1)
		FileWrite($hFile, $paragon & ";" & $data & @LF)
		$iloscLiniiBrakow = $iloscLiniiBrakow + 1
		$towary = ""
		Do
			ControlSend($uchwyt, "", "", "^+{C}")
			$towary = ClipGet()
		Until $towary <> ""
		logs($towary)
		$tmp2 = StringSplit($towary, @LF)
		For $i = 2 To UBound($tmp2) - 1
			FileWrite($hFile, $tmp2[$i])
			$iloscLiniiBrakow = $iloscLiniiBrakow + 1
		Next
		FileClose($hFile)
		$listaBrakujacychTowarow.Add($paragon, $listaBrakujacychTowarow.Count)
		;If $lista <> "ZW" Then
		$ileBrakow = $ileBrakow + 1
		;EndIf
		logs("Zapisano brak towaru: " & $paragon)
	Else
		logs("Proba zduplikowania brakujacego towaru: " & $paragon)
		;If $lista <> "ZW" Then
		$ileBrakow = $ileBrakow + 1
		;EndIf
	EndIf
	While $brak <> 0
		$brak = WinGetHandle("[REGEXPTITLE:(?i)(.* - brak towaru w magazynie)]")
		$uchwyt = szukajUchwytu($brak, "Zamknij", "Button")
		ControlClick($uchwyt, "", "")
	WEnd

EndFunc   ;==>zapis
;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Zamyka program po wcisnieciu BREAK(ctrl+break)

Func zamknijW()
	logs('@@ (220) :(' & @MIN & ':' & @SEC & ') zamknijW()' & @CR) ;### Function Trace
	Exit
EndFunc   ;==>zamknijW

;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Wstrzymuje dzia³anie programu

Func pauzaW()
	logs('@@ (228) :(' & @MIN & ':' & @SEC & ') pauzaW()' & @CR) ;### Function Trace

	If $_PAUSE = False Then
		$_PAUSE = True
		logs("PAUZA........." & @LF)
		;pauzaLoop()
	Else
		$_PAUSE = False
		logs("START........." & @LF)
	EndIf
EndFunc   ;==>pauzaW

Func pauzaLoopW()
	logs('@@ (241) :(' & @MIN & ':' & @SEC & ') pauzaLoopW()' & @CR) ;### Function Trace
	While $_PAUSE
		Sleep(1000)
	WEnd
EndFunc   ;==>pauzaLoopW

Func pauzaSzybkaW()
	logs('@@ (248) :(' & @MIN & ':' & @SEC & ') pauzaSzybkaW()' & @CR) ;### Function Trace

	If $_PAUSE = False Then
		$_PAUSE = True
		logs("PAUZA........." & @LF)
		pauzaLoopW()
	Else
		$_PAUSE = False
		logs("START........." & @LF)
	EndIf
EndFunc   ;==>pauzaSzybkaW
