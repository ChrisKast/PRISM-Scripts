'GATHERING STATS---------------------------------------------------------------------------------------------------- 
name_of_script = "record-iw-info.vbs" 
start_time = timer 

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/13/2016", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS---------------------------------------------------------------------------

BeginDialog IW_CALC_Dialog, 0, 0, 256, 236, "IW Calculator Dialog"
  EditBox 70, 5, 100, 14, PRISM_case_number
  EditBox 100, 30, 110, 14, employer_name
  EditBox 100, 50, 60, 14, Current_Support
  CheckBox 20, 80, 60, 10, "20 Percent", Percent
  CheckBox 20, 94, 60, 10, "Nonaccrual Amt", Other_Amount
  EditBox 80, 90, 50, 14, Nonaccrual
  CheckBox 10, 130, 140, 10, "Add 30-day FREE worklist?", cawd_check
  CheckBox 10, 150, 70, 10, "Add a CAAD note? ", caad_check
  Text 10, 164, 90, 10, "Additional CAAD note text:"
  EditBox 100, 160, 120, 14, caad_text
  EditBox 110, 190, 50, 14, worker_signature
  ButtonGroup ButtonPressed
    OkButton 140, 210, 50, 14
    CancelButton 200, 210, 50, 14
  Text 10, 10, 50, 10, "Case Number:"
  Text 10, 30, 90, 10, "Employer/Payor of Funds:"
  Text 10, 54, 90, 10, "Ongoing Monthly Obligation:"
  GroupBox 10, 70, 150, 50, "Arrears Collection Rate:"
  Text 10, 190, 100, 20, "Please sign your CAAD note:"
EndDialog




'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""

'to pull up my prism 
EMFocus

'checks to make sure we are in PRISM
CALL check_for_PRISM(True)

CALL navigate_to_PRISM_screen ("CAFS")

'variable name in edit box that i want autofilled
EMReadScreen PRISM_case_number, 13, 4, 8
EMReadScreen Current_Support, 10, 9, 29
EMReadScreen Nonaccrual, 10, 10, 29

IF trim(Nonaccrual) <> "0.00" THEN
	Other_Amount = checked
ELSE
	Other_Amount = unchecked
END IF

'taking me to cast so i can read the case number to put in dialog box
CALL navigate_to_PRISM_screen ("NCID")

EMWriteScreen "B", 3, 29
Transmit

'it is reading the employer name and putting in dialog box
EMReadScreen Employer_Name, 20, 8, 51 


CALL navigate_to_PRISM_screen ("CAFS")

'adding a loop
Do
	err_msg = ""
	caad_check = checked
	cawd_check = checked
	Dialog IW_CALC_Dialog				'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "Prism case number must be completed."
		IF Current_Support = "" THEN err_msg = err_msg & vbNewline & "Current Support must be completed."
		
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & "Please resolve for the script to continue."
		END IF
LOOP UNTIL err_msg = ""

'Converting accrual amts to number from string 
Current_Support = CDbl(Current_Support)
Nonaccrual = CDbl(Nonaccrual)

IF Other_Amount = checked THEN MoTotal = Current_Support + Nonaccrual
IF Percent = checked THEN MoTotal = (Current_Support + Nonaccrual) * 1.2

IF Percent = unchecked AND Other_Amount = unchecked THEN MoTotal = Current_Support

'Calculating pay period amounts
Dim WeekPay, BiWeekPay, SemiMoPay

WeekPay = MoTotal /4.333
WeekPay = FormatNumber(WeekPay, 2)
BiWeekPay = MoTotal /2.167
BiWeekPay = FormatNumber(BiWeekPay, 2)
SemiMoPay = MoTotal/2
SemiMoPay = FormatNumber(SemiMoPay, 2)
MoTotal = FormatNumber(MoTotal, 2)

IF caad_check = checked THEN
	'brings me to caad and creates a FREE note
	CALL navigate_to_PRISM_screen ("CAAD")
	PF5
	EMWriteScreen "A", 3, 29
	EMWriteScreen "free", 4, 54
	EMSetCursor 16, 4

	'this will add information to the caad note
	CALL write_variable_in_CAAD ("* IW sent to " & Employer_Name  &  " " & FormatCurrency(MoTotal)  & " per month")
	IF Other_Amount = checked THEN CALL write_variable_in_CAAD ("* Non-accrual is " & FormatCurrency(Nonaccrual))
	IF Percent = checked THEN CALL write_variable_in_CAAD("* Plus Additional 20% of obligation")
	CALL write_variable_in_CAAD ("weekly: " & FormatCurrency(WeekPay) & "  biweekly: " & FormatCurrency(BiWeekPay) & "  semimonthly: "& FormatCurrency(SemiMoPay))
	CALL write_variable_in_CAAD (caad_text)
	CALL write_variable_in_CAAD(worker_signature)
	transmit
	PF3
END IF


IF cawd_check = checked THEN
	CALL navigate_to_PRISM_screen ("CAWD")
	PF5
	EMWriteScreen "A", 3, 30
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "30", 17, 52

	EMSetCursor 10, 4
	CALL write_variable_in_CAAD ("Payments? Monthly: " & FormatCurrency(MoTotal))
	CALL write_variable_in_CAAD ("Bi-Weekly: " & FormatCurrency(BiWeekPay))
	CALL write_variable_in_CAAD ("Semi-Monthly: " & FormatCurrency(SemiMoPay))
	CALL write_variable_in_CAAD ("Weekly: " & FormatCurrency(WeekPay))
END IF

IF caad_check = unchecked AND cawd_check = unchecked THEN	
 
	'msgbox needed to show calculations, weekly, biweekly, semi monthly, and monthly
	MsgBox ("Monthly: " & FormatCurrency(MoTotal) & VbNewline & VbNewline & _
	"Weekly: " & FormatCurrency(WeekPay) & VbNewline & VbNewline & _
	"Bi-Weekly: " & FormatCurrency(BiWeekPay) & VbNewline & VbNewline & _
	"Semi-Monthly: " & FormatCurrency(SemiMoPay))
END IF
script_end_procedure("")
