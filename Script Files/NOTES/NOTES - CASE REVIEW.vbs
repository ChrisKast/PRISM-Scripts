Option Explicit


DIM beta_agency

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF



'This is what was declared
DIM case_review_dialog, prism_case_number, case_type, date_of_last_payment, amount_of_last_payment, total_arrears, prior_contempt_checkbox, auto_warrant_checkbox, date_of_last_contempt_order, worker_signature, buttonpressed


'This is the dialog i am using
BeginDialog case_review_dialog, 0, 0, 181, 240, "Case Review"
  EditBox 110, 5, 65, 15, prism_case_number
  DropListBox 120, 30, 55, 45, "(select one)"+chr(9)+"Ongoing"+chr(9)+"Arrears Only", case_type
  EditBox 130, 55, 45, 15, date_of_last_payment
  EditBox 130, 80, 45, 15, amount_of_last_payment
  EditBox 110, 105, 65, 15, total_arrears
  CheckBox 5, 135, 60, 10, "Prior Contempt", Prior_contempt_checkbox
  EditBox 130, 155, 45, 15, date_of_last_contempt_order
  CheckBox 5, 180, 85, 10, "Auto Warrant Provision", auto_warrant_checkbox
  EditBox 145, 195, 30, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 220, 50, 15
    CancelButton 125, 220, 50, 15
  Text 5, 160, 95, 10, "Date of Last Contempt Order"
  Text 5, 60, 75, 10, "Date of Last Payment"
  Text 5, 10, 70, 10, "PRISM Case Number"
  Text 5, 85, 80, 10, "Amount of Last Payment"
  Text 5, 30, 50, 10, "Case Type"
  Text 70, 200, 60, 10, "Worker Signature"
  Text 5, 110, 45, 10, "Total Arrears"
EndDialog

EMConnect ""	'connect to bluezone

CALL check_for_PRISM(TRUE) 'checks to see if you're timed out

CALL navigate_to_PRISM_screen ("CAST")

EMWaitReady 0, 0	'causes a pause for screen to load

EMReadScreen PRISM_case_number, 13, 4, 08 'grabs prism case number

CALL navigate_to_PRISM_screen ("PALC")

EMWaitReady 0, 0

EMReadScreen date_of_last_payment, 8, 9, 59 'grabs last pmt date

EMReadScreen amount_of_last_payment, 10, 9, 70 'grabs last pmt amt

amount_of_last_payment = TRIM (amount_of_last_payment)'trims extra spaces

CALL navigate_to_PRISM_screen ("CAFS")

EMWaitReady 0, 0	'causes a pause for screen to load

EMReadScreen total_arrears, 13, 12, 65 'grabs total arrears

total_arrears = TRIM (total_arrears)'trims extra spaces

'This is the dialog that will run

DO			'this will cause a loop causing worker to add signature
	Dialog case_review_dialog
	If buttonpressed = 0 THEN stopscript
	If worker_signature = "" THEN MsgBox "Please add worker signature!"
LOOP UNTIL worker_signature <> ""

'this is when and what will dump into CAAD

CALL navigate_to_PRISM_screen ("CAAD")

PF5

EMWriteScreen "A", 3, 29

EMWriteScreen "FREE", 4, 54

EMSetCursor 16, 4

CALL write_variable_in_CAAD ("***Case Review Note***")
CALL write_bullet_and_variable_in_CAAD("date of last payment", date_of_last_payment)
CALL write_bullet_and_variable_in_CAAD("amount of last payment", amount_of_last_payment)
CALL write_bullet_and_variable_in_CAAD("total arrears", total_arrears)
IF prior_contempt_checkbox = checked then CALL write_variable_in_CAAD("* Prior contempt")
IF auto_warrant_checkbox = checked then CALL write_variable_in_CAAD("* Contempt order allows for auto warrant")
CALL write_variable_in_CAAD (worker_signature)

StopScript
