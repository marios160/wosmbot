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
	ConsoleWrite('@@ (28) :(' & @MIN & ':' & @SEC & ') startWSMBot()' & @CR) ;### Function Trace
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


	$_ENDDATA = StringReplace(_DateAdd('d', 1, $_ENDDATA), "/", "-")
	;____________Petla programowa_______________________________________________________________________________
	logsW("START BotWSM")
	While $_DATA <> $_ENDDATA
		logsW("**************************************************************************************")
		logsW("*************************************" & $_DATA & " **************************************")
		logsW("**************************************************************************************")
		For $i = 0 To UBound($dokumenty) - 1
			If StringLen($dokumenty[$i]) = 3 Then
				wybierzMagazyn($_ListaMagazynow.Item($dokumenty[$i]))
			Else
				wykonujListeW($dokumenty[$i])
			EndIf
			Sleep(1000)
		Next
		$_DATA = StringReplace(_DateAdd('d', 1, $_DATA), "/", "-")
	WEnd
	logsW("KONIEC BotWSM")
EndFunc

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
	ConsoleWrite('@@ (99) :(' & @MIN & ':' & @SEC & ') wykonujListeW()' & @CR) ;### Function Trace
	$listaPelna = $lista
	$lista = wybierzListe($lista)
	Sleep(500)
	ustawOkres($_DATA, $lista)
	Sleep(500)
	logsW("------------------------------- START " & $lista & " ----------------------------------------")
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
			;ConsoleWrite("ENDPOS: " & $endpos)
			$curpos = ""
			sendGXWNDHomeHard($lista)
			$rekord = ""
			While $endpos <> $curpos
				$czyDaSieOdlozyc = False
				Do
					$curpos = kopiujWiersz($lista)
				Until StringInStr($curpos,"magazynow") > 0 And $rekord <> getNumer($curpos)
				$rekord = getNumer($curpos)
				logsW($rekord)
				If StringInStr($curpos, $STR_WYWOLUJE_SKUTEK ) = 0 Then
					While Not $czyDaSieOdlozyc
						Do
							$curpos2 = kopiujWiersz($lista)
						Until StringInStr($curpos2,"magazynow") > 0
						;logsW("curpos2: " & $curpos2)
						$czyDaSieOdlozyc = wywolajSkutekMagazynowy($listaPelna)
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
				;MsgBox(0,"a","a")
				StringReplace($listaKopia, $STR_ODLOZONY_SKUTEK, "xxx")
				$ileZnaleziono = @extended
				If $ileZnaleziono <> $ileBrakow Then
					logsW("OMINIETO PRODUKT! " & "; Brakuje: " & $ileBrakow & ", znaleziono: " & @extended)
				EndIf
			EndIf
			logsW($ileZnaleziono & " " & $ileBrakow & @LF & $listaKopia)
			Sleep(5000)
		Until $ileZnaleziono = $ileBrakow
	EndIf
	logs("------------------------------- KONIEC " & $lista & " ----------------------------------------")
EndFunc   ;==>wykonujListe

;---------------------------------------------------------------------------------------------------------------------
;Funkcja obsluguje okienko brakuj¹cego towaru. Zapisuje do pliku wszystkie pozycje z listy a nastepnie zamyka okienko
;przyjmuje parametr okreslajacy czy dany produkt juz zosta³ zapisany oraz nazwe przerabianej listy

Func zapiszBrakTowaru($lista)
	ConsoleWrite('@@ (167) :(' & @MIN & ':' & @SEC & ') zapiszBrakTowaru()' & @CR) ;### Function Trace
	$brak = WinGetHandle("[REGEXPTITLE:(?i)(.* - brak towaru w magazynie)]")
	$uchwyt = szukajUchwytu($brak, "", "GXWND")
	$paragon = ""
	$data = ""
	$tmp = StringSplit(ClipGet(), "	")
	$data = $tmp[2]
	$paragon = $tmp[3]

	If Not $listaBrakujacychTowarow.Exists($paragon) Then
		$hFile = FileOpen("WSMbraktowaru.txt", 1)
		FileWrite($hFile, $paragon & ";" & $data & @LF)
		$iloscLiniiBrakow = $iloscLiniiBrakow + 1
		ControlSend($uchwyt, "", "", "^+{C}")
		$tmp2 = StringSplit(ClipGet(), @LF)
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
		ConsoleWrite("a")
	WEnd

EndFunc   ;==>zapiszBrakTowaru

Func logsW($text)
	ConsoleWrite('@@ (222) :(' & @MIN & ':' & @SEC & ') logsW()' & @CR) ;### Function Trace
	_FileWriteLog("logWSM.log", "[" & $_DATA & "] " & $text)
	ConsoleWrite("[" & $_DATA & "] " & $text & @LF)
EndFunc   ;==>logs

;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Zamyka program po wcisnieciu BREAK(ctrl+break)

Func zamknijW()
	ConsoleWrite('@@ (231) :(' & @MIN & ':' & @SEC & ') zamknijW()' & @CR) ;### Function Trace
	Exit
EndFunc   ;==>zamknij

;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Wstrzymuje dzia³anie programu

Func pauzaW()
	ConsoleWrite('@@ (239) :(' & @MIN & ':' & @SEC & ') pauzaW()' & @CR) ;### Function Trace

	If $_PAUSE = False Then
		$_PAUSE = True
		logsW("PAUZA........." & @LF)
		;pauzaLoop()
	Else
		$_PAUSE = False
		logsW("START........." & @LF)
	EndIf
EndFunc   ;==>pauza

Func pauzaLoopW()
	ConsoleWrite('@@ (252) :(' & @MIN & ':' & @SEC & ') pauzaLoopW()' & @CR) ;### Function Trace
	While $_PAUSE
		Sleep(1000)
	WEnd
EndFunc   ;==>pauzaLoop

Func pauzaSzybkaW()
	ConsoleWrite('@@ (259) :(' & @MIN & ':' & @SEC & ') pauzaSzybkaW()' & @CR) ;### Function Trace

	If $_PAUSE = False Then
		$_PAUSE = True
		logsW("PAUZA........." & @LF)
		pauzaLoopW()
	Else
		$_PAUSE = False
		logsW("START........." & @LF)
	EndIf
EndFunc   ;==>pauzaSzybka