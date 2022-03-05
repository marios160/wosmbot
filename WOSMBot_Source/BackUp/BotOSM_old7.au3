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
;#include <APISubiektWin.au3>

Func startOSMBot($dane, $dokumenty, $_ListaMagazynow)
	logs('@@ (28) :(' & @MIN & ':' & @SEC & ') startOSMBot()' & @CR) ;### Function Trace
	HotKeySet("{F6}", "zamknij")
	HotKeySet("{PAUSE}", "pauza")
	HotKeySet("{F4}", "pauzaSzybka")
;~ 	$odp = $IDNO
;~ 	While $odp == $IDNO
;~ 		$odp = MsgBox($MB_YESNOCANCEL, "Start", "Proszę włączyć na liscie zwrotow kolumne Dokument korygowany oraz na każdej liście ustawić sortowanie od najmłodszych dokumentów do najstarszych!" & @LF & "Czy chcesz rozpocząć pracę?")
;~ 	WEnd
;~ 	If $odp == $IDCANCEL Then
;~ 		Exit
;~ 	EndIf

;~ 	$odp = MsgBox($MB_YESNO, "Czyszczenie loga", "Czy chcesz wyczyscic 'logOSM.log'?")
;~ 	If $odp == $IDYES Then
;~ 	$hF = FileOpen("logOSM.log", 2)
;~ 	FileClose($hF)
;~ 	EndIf


	;;-----------Inicjalizacja zmiennych
	Global $_PAUSE = False
	Global $_DATA = $dane[1]
	Global $_ENDDATA = $dane[2]
	Global $_PODMIOT = $dane[0]

	;;-----------Petla programowa,,,,,,,,,,
	logs("START BotOSM")
	logs($_DATA & " " & $_ENDDATA)
	While $_DATA >= $_ENDDATA
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
					wykonujListeO($dokumenty[$i])
				Until $_RESTART = False
			EndIf
			Sleep(1000)
		Next
		$_DATA = StringReplace(_DateAdd('d', -1, $_DATA), "/", "-")

	WEnd
	logs("KONIEC BotOSM")
EndFunc   ;==>startOSMBot

;#####################################################################################################################
;#####################################################################################################################
;################################													##################################
;################################						FUNKCJE						##################################
;################################													##################################
;#####################################################################################################################
;#####################################################################################################################

;---------------------------------------------------------------------------------------------------------------------
;Funkcja wykonuje wszystkie instrukcje potrzebne do wywołania skutkow magazynowych wszystkich pozycji na konkretnej liscie
;Przyjmuje nazwę listy na ktorej działamy
Func wykonujListeO($lista)
	logs('@@ (95) :(' & @MIN & ':' & @SEC & ') wykonujListeO()' & @CR) ;### Function Trace
	$listaPelna = $lista
	$lista = wybierzListe($lista)
	Sleep(500)
	ustawOkres($_DATA, $lista)
	Sleep(500)
	logs("------------------------------- START " & $lista & " ----------------------------------------")
	If $_PAUSE = True Then
		pauzaLoop()
	EndIf
	If Not czyPusta($lista) Then
		fokusNaGXWND($lista)
		Do
			Sleep(500)
			sendGXWNDEndHard($lista)
			Do
				$endpos = kopiujWiersz($lista)
			Until StringLen($endpos) > 0
			;logs("ENDPOS: " & $endpos)
			$curpos = ""
			sendGXWNDHomeHard($lista)
			While $endpos <> $curpos
				$czyDaSieOdlozyc = False
				Do
					$curpos = kopiujWiersz($lista)
				Until StringInStr($curpos, "magazynow") > 0
				;logs("curpos: " & $curpos)
				If StringInStr($curpos, $STR_ODLOZONY_SKUTEK) = 0 Then
					While Not $czyDaSieOdlozyc
						Do
							$curpos2 = kopiujWiersz($lista)
						Until StringInStr($curpos2, "magazynow") > 0
						;logs("curpos2: " & $curpos2)
						$czyDaSieOdlozyc = odlozSkutekMagazynowy($listaPelna)
						If $_RESTART Then Return
					WEnd
				EndIf
				Sleep(1)
				If $_PAUSE = True Then
					pauzaLoop()
					ExitLoop
				EndIf
				sendGXWNDDown($lista)
			WEnd
			;logs("EEE:" & $endpos)
			;logs("CCC:" & $curpos)
			If StringInStr(kopiujListe($lista), $STR_WYWOLUJE_SKUTEK) <> 0 Then
				StringReplace(kopiujListe($lista), $STR_WYWOLUJE_SKUTEK, "xxx")
				logs("OMINIETO PRODUKT! " & "; " & " znaleziono: " & @extended)
			EndIf
		Until StringInStr(kopiujListe($lista), $STR_WYWOLUJE_SKUTEK) = 0
	EndIf
	logs("------------------------------- KONIEC " & $lista & " ----------------------------------------")
EndFunc   ;==>wykonujListeO





;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Zamyka program po wcisnieciu BREAK(ctrl+break)
Func zamknij()
	logs('@@ (158) :(' & @MIN & ':' & @SEC & ') zamknij()' & @CR) ;### Function Trace
	Exit
EndFunc   ;==>zamknij

;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Wstrzymuje działanie programu
Func pauza()
	logs('@@ (165) :(' & @MIN & ':' & @SEC & ') pauza()' & @CR) ;### Function Trace
	If $_PAUSE = False Then
		$_PAUSE = True
		logs("PAUZA........." & @LF)
		;pauzaLoop()
	Else
		$_PAUSE = False
		logs("START........." & @LF)
		;ControlFocus($t, "", "")
	EndIf
EndFunc   ;==>pauza
Func pauzaLoop()
	logs('@@ (177) :(' & @MIN & ':' & @SEC & ') pauzaLoop()' & @CR) ;### Function Trace
	While $_PAUSE
		Sleep(1000)
	WEnd
EndFunc   ;==>pauzaLoop
Func pauzaSzybka()
	logs('@@ (183) :(' & @MIN & ':' & @SEC & ') pauzaSzybka()' & @CR) ;### Function Trace
	If $_PAUSE = False Then
		$_PAUSE = True
		logs("PAUZA........." & @LF)
		pauzaLoop()
	Else
		$_PAUSE = False
		logs("START........." & @LF)
		;		ControlFocus($t, "", "")
	EndIf
EndFunc   ;==>pauzaSzybka





