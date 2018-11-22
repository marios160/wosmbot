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



	;;-----------Inicjalizacja zmiennych
	Global $_PAUSE = False
	Global $_DATA = $dane[1]
	Global $_ENDDATA = $dane[2]
	Global $_PODMIOT = $dane[0]

	;;-----------Petla programowa,,,,,,,,,,
	logs("START BotOSM")
	logs($_DATA & " - " & $_ENDDATA)
	While $_DATA >= $_ENDDATA
		$hData = FileOpen("_DATA.conf", $FO_OVERWRITE) ; zapisujemy date
		FileWrite($hData, $_DATA & @LF)
		FileWrite($hData, $_ENDDATA & @LF)
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
		dodajStatus("----- " & $_DATA & " -----")
	WEnd
	logs("KONIEC BotOSM")
	dodajStatus("Koniec BotOSM")
	MsgBox(0, "Koniec", "Koniec")
	Exit
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
	$lista = wybierzListe($lista) ;ustawiamy liste
	Sleep(500)
	ustawOkres($_DATA, $lista) ;ustawiamy okres
	Sleep(500)
	dodajStatus("-- " & $lista & " --")
	logs("------------------------------- START " & $lista & " ----------------------------------------")
	If Not czyPusta($lista) Then ;sprawdzamy czy lista pusta lub czy wszystko na niej jest juz wywolane
		fokusNaGXWND($lista) ;fokus na liste
		Do
			Sleep(500)
			sendGXWNDEndHard($lista) ;lecimy na ostatni rekord zeby wiedziec kiedy skonczyc
			Do
				$endpos = kopiujWiersz($lista) ;zapisujemy ostatni rekord
			Until StringLen($endpos) > 0
			$curpos = ""
			sendGXWNDHomeHard($lista) ;wracamy na poczatek
			While $endpos <> $curpos ;dopokie nie dojdziemy do ostatniego rekordu
				$czyDaSieOdlozyc = False
				Do
					$curpos = kopiujWiersz($lista)
				Until StringInStr($curpos, "magazynow") > 0
				If StringInStr($curpos, $STR_ODLOZONY_SKUTEK) = 0 Then ;jesli skutek nie jest odlozony
					While Not $czyDaSieOdlozyc ;
						Do
							$curpos2 = kopiujWiersz($lista)
						Until StringInStr($curpos2, "magazynow") > 0
						$czyDaSieOdlozyc = odlozSkutekMagazynowy($listaPelna)
						If $_RESTART Then Return
					WEnd
				EndIf
				Sleep(1)
				sendGXWNDDown($lista)
			WEnd
			$kopiaListy = kopiujListe($lista)
			If StringInStr($kopiaListy, $STR_WYWOLUJE_SKUTEK) <> 0 Then
				StringReplace($kopiaListy, $STR_WYWOLUJE_SKUTEK, "xxx")
				logs("OMINIETO PRODUKT! " & "; " & " znaleziono: " & @extended)

			EndIf
		Until StringInStr($kopiaListy, $STR_WYWOLUJE_SKUTEK) = 0
	EndIf
	logs("------------------------------- KONIEC " & $lista & " ----------------------------------------")
EndFunc   ;==>wykonujListeO





;---------------------------------------------------------------------------------------------------------------------
;Funkcja uzywana do skrotow klawiszowych. Zamyka program po wcisnieciu BREAK(ctrl+break)
Func zamknij()
	logs('@@ (158) :(' & @MIN & ':' & @SEC & ') zamknij()' & @CR) ;### Function Trace
	Exit
EndFunc   ;==>zamknij





