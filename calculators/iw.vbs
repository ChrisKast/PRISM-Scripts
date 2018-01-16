'STATS GATHERING ===========================================================================================================
name_of_script = "iw.vbs"
start_time = timer
STATS_Counter = 1
STATS_manualtime = 120
STATS_denomination = "C"
'END OF STATS BLOCK ========================================================================================================

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
DIM IW_CALC_Dialog, PRISM_case_number, Current_Support, Percent, Manual, Other_Amount, err_msg, ButtonPressed, case_number_is_valid, MoTotal, Month_NonAccrual, Month_Accrual

BeginDialog IW_CALC_Dialog, 0, 0, 176, 110, "IW CALC Dialog"
  EditBox 60, 5, 95, 15, PRISM_case_number
  EditBox 100, 25, 55, 15, Current_Support
  CheckBox 10, 65, 45, 10, "20 Percent", Percent
  ButtonGroup ButtonPressed
    OkButton 60, 85, 50, 15
    CancelButton 120, 85, 50, 15
  Text 5, 10, 50, 10, "Case Number"
  Text 5, 30, 90, 10, "Current Monthly Obligation "
  Text 5, 50, 80, 10, "Arrears Collection Rate"
EndDialog



'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

CALL navigate_to_PRISM_screen ("CAFS")

'variable name in edit box that i want autofilled
EMReadScreen PRISM_case_number, 13, 4, 8

'getting amounts to autofill
EMReadScreen Month_Accrual, 8, 9, 31
EMReadScreen Month_NonAccrual, 8, 10, 31
Month_Accrual = Trim(Month_Accrual)
Month_NonAccrual = Trim(Month_NonAccrual)

'Converting accrual amts to number from string and calculating total monthly amount
Month_Accrual = CDbl(Month_Accrual)
Month_NonAccrual = CDbl(Month_NonAccrual)


Current_Support = Month_Accrual + Month_NonAccrual
Current_Support = Trim(Current_Support)
Current_Support = FormatNumber(Current_Support)

'adding a loop
Do
	err_msg = ""
	Dialog IW_CALC_Dialog				'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF PRISM_case_number = "" THEN err_msg = err_msg & vbNewline & "Prism case number must be completed."
		IF Current_Support = "" THEN err_msg = err_msg & vbNewline & "Current Support must be completed"
		'IF CP = 0 AND NCP = 0 THEN err_msg = vbNewline & "Either CP or NCP must be selected."

		IF err_msg <> "" THEN
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & "Please resolve for the script to continue."
		END IF
LOOP UNTIL err_msg = ""


Current_Support = CDbl(Current_Support)

IF Percent = checked THEN MoTotal = Current_Support * 1.2
IF Percent = 0 THEN MoTotal = Current_Support

'Calculating pay period amounts
Dim WeekPay, BiWeekPay, SemiMoPay

WeekPay = MoTotal /4.333
WeekPay = FormatNumber(WeekPay, 2)


BiWeekPay = MoTotal /2.167
BiWeekPay = FormatNumber(BiWeekPay, 2)

SemiMoPay = MoTotal/2
SemiMoPay = FormatNumber(SemiMoPay, 2)

MoTotal = FormatNumber(Mototal)

'takes you to palc so you can see the amount that is being received on the case
CALL navigate_to_PRISM_screen ("PALC")


'msgbox needed to show calculations, weekly, biweekly, semi monthly, and monthly with 20%
IF Percent = checked THEN
	MsgBox ("Monthly: $" & MoTotal & VbNewline & VbNewline & _
		"Weekly: $" & WeekPay & VbNewline & VbNewline & _
		"Bi-Weekly: $" & BiWeekPay & VbNewline & VbNewline & _
		"Semi-Monthly: $" & SemiMoPay & VbNewline & VbNewline & _
		"20% of current support:  $" & Current_Support * .2)

END IF

'without 20%
IF Percent = 0 THEN
	MsgBox ("Monthly: $" & MoTotal & VbNewline & VbNewline & _
		"Weekly: $" & WeekPay & VbNewline & VbNewline & _
		"Bi-Weekly: $" & BiWeekPay & VbNewline & VbNewline & _
		"Semi-Monthly: $" & SemiMoPay)


END IF


script_end_procedure("")
