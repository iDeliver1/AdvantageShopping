Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("next_btn").Click

If Browser("Advantage Shopping").Page("Advantage Shopping").WebRadioGroup("safepay").Exist Then
	Browser("Advantage Shopping").Page("Advantage Shopping").WebRadioGroup("safepay").Select "#0"
		Call fPass_report("Safepay Radio","Click on the Safepay")
		Call fExcelreport("Safepay Radio","Click on the Safepay","PASS",Time,Date)
Else
		Call fFail_report("Safepay Radio","Safepay Radio is not avaliable")
		Call fExcelreport("Safepay Radio","Safepay Radio is not avaliable","FAIL",Time,Date)
End If

If Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("safepay_username").Exist(2) Then
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("safepay_username").Set "iDeliver"
		Call fPass_report("Safepay_username","Enter on the Safepay_username")
		Call fExcelreport("Safepay_username","Enter on the Safepay_username","PASS",Time,Date)
Else
		Call fFail_report("Safepay_username","Safepay_username is not avaliable")
		Call fExcelreport("Safepay_username","Safepay_username is not avaliable","FAIL",Time,Date)
End If

If Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("safepay_password").Exist(2) Then
	Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("safepay_password").Set "iDeliver1"
		Call fPass_report("Safepay_password","Enter on the Safepay_password")
		Call fExcelreport("Safepay_password","Enter on the Safepay_password","PASS",Time,Date)
Else
		Call fFail_report("Safepay_password","Safepay_password is not avaliable")
		Call fExcelreport("Safepay_password","Safepay_password edit field is not avaliable","FAIL",Time,Date)
End If

If Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("pay_now_btn_SAFEPAY").Exist(2) Then
	Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("pay_now_btn_SAFEPAY").Click
		Call fPass_report("Pay_now","Click on the Pay_now")
		Call fExcelreport("Pay_now","Click on the Pay_now","PASS",Time,Date)
Else
		Call fFail_report("Pay_now","Pay_now is not avaliable")
		Call fExcelreport("Pay_now","Pay_now is not avaliable","FAIL",Time,Date)
End If

Browser("Advantage Shopping").Page("Advantage Shopping").Sync @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("Advantage Shopping").Close
