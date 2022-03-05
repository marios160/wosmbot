#include-once

Global Const $HTTP_STATUS_OK = 200

Func HttpPost($sURL, $sData = "")
	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("POST", $sURL, False)
	If (@error) Then Return SetError(1, 0, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")

	$oHTTP.Send($sData)
	If (@error) Then Return SetError(2, 0, 0)

	If ($oHTTP.Status <> $HTTP_STATUS_OK) Then Return SetError(3, 0, 0)

	Return SetError(0, 0, $oHTTP.ResponseText)
EndFunc   ;==>HttpPost

Func HttpGet($sURL, $sData = "")
	Return
	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sURL & "?" & $sData, False)
	If (@error) Then Return SetError(1, 0, 0)

	$oHTTP.Send()
	If (@error) Then Return SetError(2, 0, 0)

	If ($oHTTP.Status <> $HTTP_STATUS_OK) Then Return SetError(3, 0, 0)

	Return SetError(0, 0, $oHTTP.ResponseText)
EndFunc   ;==>HttpGet

Func wyslijSMS($numer, $msg)
	HttpGet("http://212.182.115.194:5641/index.php/http_api/send_sms?login=backup&pass=Haslonasmseagle&to=" & $numer & "&message=" & $msg)
EndFunc   ;==>wyslijSMS
