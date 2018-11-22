Local $tablica[0]
$aktMagTxt = ""
;~ $tablica = addEl("asd",$tablica)
;~ $tablica = addEl("asda",$tablica)
;~ $tablica = addEl("asdd",$tablica)
;~ $tablica = addEl("asdd1",$tablica)
;~ $tablica = addEl("asdd2",$tablica)
;~ $tablica = addEl("asdd3",$tablica)
;~ $tablica = addEl("asdd4",$tablica)
;~ $tablica = addEl("asdd56",$tablica)
;~ $tablica = removeEl("asda",$tablica)
;~ $tablica = removeEl("asdd3",$tablica)

;~ ConsoleWrite(toString($tablica))



;~ For $i = 0 to UBound($tablica) - 1
;~ 	ConsoleWrite($tablica[$i] & @LF)
;~ Next



Func addEl($text, $tablica)
	If UBound($tablica) > 0 Then
		If StringInStr($tablica[UBound($tablica) - 1 ]," [" & $aktMagTxt & "]") = 0 Then
			ReDim $tablica[UBound($tablica)+2]
			$tablica[UBound($tablica)-2] = $aktMagTxt
		Else
			ReDim $tablica[UBound($tablica)+1]
		EndIf
	Else
		ReDim $tablica[UBound($tablica)+2]
		$tablica[UBound($tablica)-2] = $aktMagTxt
	EndIf
	$tablica[UBound($tablica)-1] = $text
	$tablica = ustawListe($tablica)
	return $tablica
EndFunc



Func removeEl($text, $tablica)
	For $i = 0 To UBound($tablica) - 1
		If $tablica[$i] = $text Then
			If $tablica[$i - 1] = $aktMagTxt Then
				If $i + 1 > UBound($tablica) - 1 Then
					For $j = $i - 1 To UBound($tablica) - 3
						$tablica[$j] = $tablica[$j + 2]
					Next
					ReDim $tablica[UBound($tablica)- 2]
				ElseIf StringInStr($tablica[$i + 1 ]," [" & $aktMagTxt & "]") = 0 Then
					For $j = $i - 1 To UBound($tablica) - 3
						$tablica[$j] = $tablica[$j + 2]
					Next
					ReDim $tablica[UBound($tablica)- 2]
				Else
					For $j = $i To UBound($tablica) - 2
						$tablica[$j] = $tablica[$j + 1]
					Next
					ReDim $tablica[UBound($tablica)-1]
				EndIf
			Else
				For $j = $i To UBound($tablica) - 2
					$tablica[$j] = $tablica[$j + 1]
				Next
				ReDim $tablica[UBound($tablica)-1]
			EndIf
			ExitLoop
		EndIf
	Next
	$tablica = ustawListe($tablica)
	return $tablica
EndFunc



Func moveUp($index,$tablica)
	If $index > 1 Then
	If StringRegExp($tablica[$index],"^[A-Z][A-Z][A-Z]$") = 1 Then
		;przenosimy wszystkie dokumenty pod magazynem
		$tmp = $tablica[$index - 1]
		$tablica[$index - 1] = $tablica[$index]
		$i = $index + 1
		While StringRegExp($tablica[$i]," \[[A-Z][A-Z][A-Z]\]") = 1 And $i < UBound($tablica) - 1
			$tablica[$i - 1] = $tablica[$i]
			$i = $i + 1
			;ConsoleWrite($tablica[$i])
		WEnd
		$tablica[$i - 1] = $tablica[$i]
		$tablica[$i] = $tmp
	Else
		If StringRegExp($tablica[$index - 1],"^[A-Z][A-Z][A-Z]$") = 1 Then
			$tmp = $tablica[$index - 2]
			$tmp2 = $tablica[$index - 1]
			$tablica[$index -  2] = $tablica[$index]
			$tablica[$index -  1] = $tmp2
			$tablica[$index] = $tmp
		Else
			$tmp = $tablica[$index - 1]
			$tablica[$index -  1] = $tablica[$index]
			$tablica[$index] = $tmp
		EndIf
	EndIf
	EndIf
	return ustawListe($tablica)
EndFunc

Func moveDown($index,$tablica)
	If $index < UBound($tablica) - 1 Then
	If StringRegExp($tablica[$index],"^[A-Z][A-Z][A-Z]$") = 1 Then
		;przenosimy wszystkie dokumenty pod magazynem
		$i = $index + 1
		While $i < UBound($tablica) - 1 And StringRegExp($tablica[$i]," \[[A-Z][A-Z][A-Z]\]") = 1
			$i = $i + 1
		WEnd
		If $i >= UBound($tablica) - 1 Then Return $tablica
		$tmp = $tablica[$i + 1]
		;$tablica[$index - 1] = $tablica[$index]
		$i = $i - 1
		While StringRegExp($tablica[$i]," \[[A-Z][A-Z][A-Z]\]") = 1 And $i < UBound($tablica) - 1
			$tablica[$i + 2] = $tablica[$i]
			$i = $i - 1
			;ConsoleWrite($tablica[$i])
		WEnd
		$tablica[$i + 2] = $tablica[$i]
		$tablica[$i + 1] = $tmp
	Else
		If StringRegExp($tablica[$index + 1],"^[A-Z][A-Z][A-Z]$") = 1 Then
			$tmp = $tablica[$index + 2]
			$tmp2 = $tablica[$index + 1]
			$tablica[$index +  2] = $tablica[$index]
			$tablica[$index +  1] = $tmp2
			$tablica[$index] = $tmp
		Else
			$tmp = $tablica[$index + 1]
			$tablica[$index + 1] = $tablica[$index]
			$tablica[$index] = $tmp
		EndIf
	EndIf
	EndIf
	return ustawListe($tablica)
EndFunc

Func ustawListe($tablica)
	$mag = ""
	Local $nowaTab[0]
	$licz = 0

	For $index = 0 to UBound($tablica) - 1
;~ 		ConsoleWrite("------" & $tablica[$index] & @LF)
		If StringRegExp($tablica[$index],"^[A-Z][A-Z][A-Z]$") = 0 Then
		;jesli to nie jest nazwa magazynu
			$tmpMag = StringMid($tablica[$index],StringInStr($tablica[$index],"[") + 1, 3)
			If $mag = $tmpMag Then
			;jesli ostatnio przerabiany magazyn jest taki sam
				ReDim $nowaTab[UBound($nowaTab) + 1]
				$nowaTab[$licz] = $tablica[$index]
				$licz = $licz + 1
			Else
				ReDim $nowaTab[UBound($nowaTab) + 1]
				$nowaTab[$licz] = $tmpMag
				$licz = $licz + 1
				;ConsoleWrite ($licz & " ; " & $index & " ; ")
				;ConsoleWrite ($nowaTab[$licz]& " ; " & $tablica[$index] & " ; " &  @LF)
				ReDim $nowaTab[UBound($nowaTab) + 1]
				$nowaTab[$licz] = $tablica[$index]
				$licz = $licz + 1

				$mag = $tmpMag
			EndIf
		EndIf
		;ConsoleWrite("+++++" & $nowaTab[$index] & @LF)
	Next

	Return $nowaTab
EndFunc




Func toString($tablica)
	$result = ""
	For $i=0 to UBound($tablica) - 1
		$result = $result & $tablica[$i] & "|"
	Next
	Return $result
EndFunc









