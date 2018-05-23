
If Browser("Advantage Shopping").Page("Advantage Shopping").Link("SPEAKERS").Exist(2) Then
	Browser("Advantage Shopping").Page("Advantage Shopping").Link("SPEAKERS").Click
		Call fPass_report("Product Select","Click on the Product")
		Call fExcelreport("Product Select","Click on the Product","PASS",Time,Date)
Else
		Call fFail_report("Product Select","Product is not avaliable")
		Call fExcelreport("Product Select","Product is not avaliable","FAIL",Time,Date)
End If

If Browser("Advantage Shopping").Page("Advantage Shopping_2").Image("fetchImage?image_id=4700").Exist(2) Then
	Browser("Advantage Shopping").Page("Advantage Shopping_2").Image("fetchImage?image_id=4700").Click
		Call fPass_report("Sub-Product Select","Click on the Product")
		Call fExcelreport("Sub-Product Select","Click on the Product","PASS",Time,Date)
Else
		Call fFail_report("Sub-Product Select","Sub-Product is not avaliable")
		Call fExcelreport("Sub-Product Select","Sub-Product is not avaliable","FAIL",Time,Date)
End If

If Browser("Advantage Shopping").Page("Advantage Shopping_3").WebButton("save_to_cart").Exist(2) Then
	Browser("Advantage Shopping").Page("Advantage Shopping_3").WebButton("save_to_cart").Click
		Call fPass_report("Add To Cart","Click on the cart button")
		Call fExcelreport("Add To Cart","Click on the cart button","PASS",Time,Date)
Else
		Call fFail_report("Add To Cart","Cart button is not showing")
		Call fExcelreport("Add To Cart","Cart button is not showing","FAIL",Time,Date)
End If

If Browser("Advantage Shopping").Page("Advantage Shopping_3").WebButton("check_out_btn").Exist(2) Then
	Browser("Advantage Shopping").Page("Advantage Shopping_3").WebButton("check_out_btn").Click
		Call fPass_report("Check Out","Click on the Check Out button")
		Call fExcelreport("Check Out","Click on the Check Out button","PASS",Time,Date)
Else
		Call fFail_report("Check Out","Check Out button is not showing")
		Call fExcelreport("Check Out","Check Out button is not showing","FAIL",Time,Date)
End If

