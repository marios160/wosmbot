;@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
;@#####################################################################################################################@
;@###############################@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@##################################@
;@###############################@													@##################################@
;@###############################@		Funkcje wspomagające ustawianie 			@##################################@
;@###############################@		kolejnosci wykonywania list dokumentow		@##################################@
;@###############################@													@##################################@
;@###############################@													@##################################@
;@###############################@						Mateusz Błaszczak	2018	@##################################@
;@###############################@						mateuszblaszczakb@gmail.com	@##################################@
;@###############################@													@##################################@
;@###############################@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@##################################@
;@#####################################################################################################################@
;@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

;Nalezy dodac ten plik do WOSMBot.au3

#include <Array.au3>
Global $docs = ObjCreate("Scripting.Dictionary")

;-----------------------------------------------------------------------------------------------------------------------------
;czy lista dokumentow o nazwie $nzwLstDok jest dodana do listy kolejnosci $lstKolej
;zwaraca false jesli nie ma, true jesli jest
Func exist($nzwLstDok, $lstKolej)
	If _ArraySearch($lstKolej, $nzwLstDok) < 0 Then ;wyszukiwanie nazwy listy dokumentow w na liscie kolejnosci
		Return False ;jesli nie ma
	Else
		Return True ;jesli jest
	EndIf
EndFunc   ;==>exist

;-----------------------------------------------------------------------------------------------------------------------------
;Dodawanie listy dokumentow $nzwLstDok do listy kolejnosci $lstKolej z magazynu $smbMag
;zwraca liste kolejnosci z dodaną lista dokumentow
Func addEl($nzwLstDok, $lstKolej) ;, $smbMag)
	If exist($nzwLstDok, $lstKolej) Then Return $lstKolej ;sprawdzamy czy dana lista dokumentow jest juz na liscie kolejnosci
	ReDim $lstKolej[UBound($lstKolej) + 1]
	If $docs.Exist($nzwLstDok)Then
		_ArrayAdd($lstKolej,$nzwLstDok,$docs.Item($nzwLstDok))
		$docs.Remove($nzwLstDok)
	Else
		$lstKolej[UBound($lstKolej) - 1] = $nzwLstDok
	ElseIf
	$lstKolej = ustawListe($lstKolej)
	Return $lstKolej
EndFunc   ;==>addEl



Func removeEl($nzwLstDok, $lstKolej) ;, $smbMag)
	$docs.Add($nzwLstDok,_ArraySearch($lstKolej, $nzwLstDok))
	_ArrayDelete($lstKolej, _ArraySearch($lstKolej, $nzwLstDok))
	$lstKolej = ustawListe($lstKolej)
	Return $lstKolej
EndFunc   ;==>removeEl



Func moveUp($index, $lstKolej)
	If $index > 1 Then
		If StringRegExp($lstKolej[$index], "^[A-Z][A-Z][A-Z]$") = 1 Then
			;przenosimy wszystkie dokumenty pod magazynem
			$tmp = $lstKolej[$index - 1]
			$lstKolej[$index - 1] = $lstKolej[$index]
			$i = $index + 1
			While StringRegExp($lstKolej[$i], " \[[A-Z][A-Z][A-Z]\]") = 1 And $i < UBound($lstKolej) - 1
				$lstKolej[$i - 1] = $lstKolej[$i]
				$i = $i + 1
			WEnd
			$lstKolej[$i - 1] = $lstKolej[$i]
			$lstKolej[$i] = $tmp
		Else
			If StringRegExp($lstKolej[$index - 1], "^[A-Z][A-Z][A-Z]$") = 1 Then
				$tmp = $lstKolej[$index - 2]
				$tmp2 = $lstKolej[$index - 1]
				$lstKolej[$index - 2] = $lstKolej[$index]
				$lstKolej[$index - 1] = $tmp2
				$lstKolej[$index] = $tmp
			Else
				$tmp = $lstKolej[$index - 1]
				$lstKolej[$index - 1] = $lstKolej[$index]
				$lstKolej[$index] = $tmp
			EndIf
		EndIf
	EndIf
	Return ustawListe($lstKolej)
EndFunc   ;==>moveUp

Func moveDown($index, $lstKolej)
	If $index < UBound($lstKolej) - 1 Then
		If StringRegExp($lstKolej[$index], "^[A-Z][A-Z][A-Z]$") = 1 Then
			;przenosimy wszystkie dokumenty pod magazynem
			$i = $index + 1
			While $i < UBound($lstKolej) - 1 And StringRegExp($lstKolej[$i], " \[[A-Z][A-Z][A-Z]\]") = 1
				$i = $i + 1
			WEnd
			If $i >= UBound($lstKolej) - 1 Then Return $lstKolej
			$tmp = $lstKolej[$i + 1]
			$i = $i - 1
			While StringRegExp($lstKolej[$i], " \[[A-Z][A-Z][A-Z]\]") = 1 And $i < UBound($lstKolej) - 1
				$lstKolej[$i + 2] = $lstKolej[$i]
				$i = $i - 1
			WEnd
			$lstKolej[$i + 2] = $lstKolej[$i]
			$lstKolej[$i + 1] = $tmp
		Else
			If StringRegExp($lstKolej[$index + 1], "^[A-Z][A-Z][A-Z]$") = 1 Then
				$tmp = $lstKolej[$index + 2]
				$tmp2 = $lstKolej[$index + 1]
				$lstKolej[$index + 2] = $lstKolej[$index]
				$lstKolej[$index + 1] = $tmp2
				$lstKolej[$index] = $tmp
			Else
				$tmp = $lstKolej[$index + 1]
				$lstKolej[$index + 1] = $lstKolej[$index]
				$lstKolej[$index] = $tmp
			EndIf
		EndIf
	EndIf
	Return ustawListe($lstKolej)
EndFunc   ;==>moveDown

Func ustawListe($lstKolej)
	$mag = ""
	Local $nowaTab[0]
	$licz = 0

	For $index = 0 To UBound($lstKolej) - 1
		If StringRegExp($lstKolej[$index], "^[A-Z][A-Z][A-Z]$") = 0 Then
			;jesli to nie jest nazwa magazynu
			$tmpMag = StringMid($lstKolej[$index], StringInStr($lstKolej[$index], "[") + 1, 3)
			If $mag = $tmpMag Then
				;jesli ostatnio przerabiany magazyn jest taki sam
				ReDim $nowaTab[UBound($nowaTab) + 1]
				$nowaTab[$licz] = $lstKolej[$index]
				$licz = $licz + 1
			Else
				ReDim $nowaTab[UBound($nowaTab) + 1]
				$nowaTab[$licz] = $tmpMag
				$licz = $licz + 1
				ReDim $nowaTab[UBound($nowaTab) + 1]
				$nowaTab[$licz] = $lstKolej[$index]
				$licz = $licz + 1
				$mag = $tmpMag
			EndIf
		EndIf
	Next

	Return $nowaTab
EndFunc   ;==>ustawListe




Func toString($lstKolej)
	$result = ""
	For $i = 0 To UBound($lstKolej) - 1
		$result = $result & $lstKolej[$i] & "|"
	Next
	Return $result
EndFunc   ;==>toString









