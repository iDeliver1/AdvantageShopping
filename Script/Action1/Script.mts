'Variables 
Dim Username
Dim Password
Dim URL
Dim Browser_

'
Username = DataTable.Value("Username",Global)
Password = DataTable.Value("Password",Global)
URL_ 	 = DataTable.Value("URL",Global)
Browser_  = DataTable.Value("Browser",Global)

'Launch the website
systemutil.Run Browser_,URL_

'User login
Browser("Advantage Shopping").Page("Advantage Shopping").Link("My account My Orders Sign").Click
Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("username").Set Username
Browser("Advantage Shopping").Page("Advantage Shopping").WebEdit("password").Set Password
Browser("Advantage Shopping").Page("Advantage Shopping").WebButton("sign_in_btnundefined").Click

'Verify the User Login
Username_Verify = Browser("Advantage Shopping").Page("Advantage Shopping").WebElement("Username").GetROProperty("innertext")
	If(Username = trim(Username_Verify)) then
		Call fPass_report("User Login","User loggned in successfully")
		Call fExcelreport("User Login","User loggned in successfully","PASS",Time,Date)
	else
		Call fFail_report("User Login","User loggned in unsuccessfully")
		Call fExcelreport("User Login","User loggned in unsuccessfully","FAIL",Time,Date)
	End if
