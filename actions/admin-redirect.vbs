'STATS GATHERING ---------------------------
name_of_script = "admin-redirect.vbs"
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

thirty_days_away = DateAdd("d", 30, date)
month_after = DateAdd("m", 1, thirty_days_away)
redirection_month = DatePart("m", month_after)
redirection_year = DatePart("yyyy", month_after)
If len(redirection_month) = 1 then redirection_month = "0" & redirection_month

'Dim case_number, caregiver_case_number, caregiver_name, prorate_yes, prorate_no, child_one, child_two, child_three, child_four, child_five, child_six, cch_amount, cms_amount, ccc_amount, total_amount, original_cp_name
BeginDialog redirection_dialog, 0, 0, 236, 285, "Redirection Script"
  EditBox 90, 5, 145, 15, case_number
  EditBox 90, 20, 145, 15, caregiver_case_number
  EditBox 90, 35, 145, 15, original_cp_name
  EditBox 90, 50, 145, 15, caregiver_name
  CheckBox 115, 75, 25, 10, "Yes", prorate_yes
  CheckBox 140, 75, 20, 10, "No", prorate_no
  EditBox 10, 105, 220, 15, child_one
  EditBox 10, 120, 220, 15, child_two
  EditBox 10, 135, 220, 15, child_three
  EditBox 10, 150, 220, 15, child_four
  EditBox 10, 165, 220, 15, child_five
  EditBox 10, 180, 220, 15, child_six
  EditBox 5, 210, 70, 15, cch_amount
  EditBox 80, 210, 70, 15, cms_amount
  EditBox 155, 210, 70, 15, ccc_amount
  EditBox 80, 235, 70, 15, total_amount
  ButtonGroup ButtonPressed
    OkButton 55, 260, 50, 15
    CancelButton 125, 260, 50, 15
  Text 40, 10, 50, 10, "Case Number:"
  Text 5, 25, 80, 10, "Caregiver Case Number:"
  Text 20, 55, 65, 10, "Name Of Caregiver:"
  Text 5, 75, 105, 10, "Prorate Support Per Each Child?"
  Text 5, 95, 200, 10, "Child(ren) that are having support redirected (FULL NAME)"
  Text 10, 200, 180, 10, "Amounts To Be Redirected: (must be xxx.xx format)"
  Text 30, 225, 20, 10, "CCH"
  Text 100, 225, 35, 10, "CMS"
  Text 180, 225, 35, 10, "CCC"
  Text 95, 250, 40, 10, "TOTAL"
  Text 25, 40, 60, 10, "Original CP name:"
EndDialog


'Connects to Bluezone
EMConnect ""

'Starts dialog
DO
	err_msg = ""
	Dialog redirection_dialog
    	IF ButtonPressed = 0 THEN StopScript
		IF case_number = "" 											THEN err_msg = err_msg & vbCr & "* Please provide a PRISM case number."
		IF caregiver_case_number = "" 									THEN err_msg = err_msg & vbCr & "* Please provide the caregiver's case number."
		IF (case_number = caregiver_case_number) AND case_number <> ""	THEN err_msg = err_msg & vbCr & "* The current case number and the caregiver's case number match. Please review the case numbers you are providing."
		IF original_cp_name = "" 										THEN err_msg = err_msg & vbCr & "* Please provide the CP's name."
		IF caregiver_name = "" 											THEN err_msg = err_msg & vbCr & "* Please provide the caregiver's name."
		IF child_one = "" 												THEN err_msg = err_msg & vbCr & "* You have not provided the name of any children on this case."
		IF prorate_no = 1 AND prorate_yes = 1							THEN err_msg = err_msg & vbCr & "* Please indicate if the support is prorated. You cannot select YES and NO."
		IF prorate_yes = 0 AND prorate_no = 0 							THEN err_msg = err_msg & vbCr & "* Please indicate if the support is prorated. You must select either YES or NO."
		IF cch_amount = "" AND cms_amount = "" AND ccc_amount = "" 		THEN err_msg = err_msg & vbCr & "* Please indicate the CCH amount OR the CMS amount OR the CCC amount. These fields cannot be left blank."
		IF total_amount = "" 											THEN err_msg = err_msg & vbCr & "* Please indicate the total amount. This field cannot be left blank."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

CALL check_for_PRISM(false)

'goes to correct case
CALL write_value_and_transmit("CAST", 21, 18)

'Puts case number in from Dialog box
EMWriteScreen Left (case_number, 10), 4, 8
EMWriteScreen Right (case_number, 2), 4, 19

CALL write_value_and_transmit("D", 3, 29)

'________________________________________________________________________________________________________________________________________________________________________________________ NCP NOTICE
'goes to DORD
CALL navigate_to_PRISM_screen("DORD")

'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0500", 6, 36
Transmit

'entering user labels
PF14
CALL write_value_and_transmit("U", 20, 14)

FOR i = 7 to 18
	EMWriteScreen "S", i, 5
NEXT
Transmit

EMWriteScreen redirection_month, 16, 15
EMWriteScreen "/01/", 16, 17
EMWriteScreen redirection_year, 16, 21
Transmit
EMWriteScreen caregiver_name, 16, 15
Transmit
If prorate_yes = checked then CALL write_value_and_transmit("Y", 16, 15)
If prorate_no = checked then CALL write_value_and_transmit("N", 16, 15)

CALL write_value_and_transmit(child_one, 16, 15)
CALL write_value_and_transmit(child_two, 16, 15)
CALL write_value_and_transmit(child_three, 16, 15)
CALL write_value_and_transmit(child_four, 16, 15)
CALL write_value_and_transmit(child_five, 16, 15)
CALL write_value_and_transmit(child_six, 16, 15)

CALL write_value_and_transmit(cch_amount, 16, 15)
CALL write_value_and_transmit(cms_amount, 16, 15)
CALL write_value_and_transmit(ccc_amount, 16, 15)

PF8
CALL write_value_and_transmit("S", 7, 5)
CALL write_value_and_transmit(total_amount, 16, 15)

PF3
CALL write_value_and_transmit("M", 3, 29)		'At this point, the notice to the NCP is ready to be printed
'________________________________________________________________________________________________________________________________________________________________________________________ CP NOTICE
'clears DORD screen and adds and completes notice to CP
CALL write_value_and_transmit("C", 3, 29)

EMWriteScreen "F0501", 6, 36
CALL write_value_and_transmit("A", 3, 29)

'entering user labels
PF14
CALL write_value_and_transmit("U", 20, 14)

FOR i = 7 to 18
	EMWriteScreen "S", i, 5
NEXT
Transmit

EMWriteScreen redirection_month, 16, 15
EMWriteScreen "/01/", 16, 17
EMWriteScreen redirection_year, 16, 21
Transmit
EMWriteScreen caregiver_name, 16, 15
Transmit
If prorate_yes = checked then CALL write_value_and_transmit("Y", 16, 15)
If prorate_no = checked then CALL write_value_and_transmit("N", 16, 15)

CALL write_value_and_transmit(child_one, 16, 15)
CALL write_value_and_transmit(child_two, 16, 15)
CALL write_value_and_transmit(child_three, 16, 15)
CALL write_value_and_transmit(child_four, 16, 15)
CALL write_value_and_transmit(child_five, 16, 15)
CALL write_value_and_transmit(child_six, 16, 15)

CALL write_value_and_transmit(cch_amount, 16, 15)
CALL write_value_and_transmit(cms_amount, 16, 15)
CALL write_value_and_transmit(ccc_amount, 16, 15)

PF8
CALL write_value_and_transmit("S", 7, 5)
CALL write_value_and_transmit(total_amount, 16, 15)

PF3
CALL write_value_and_transmit("M", 3, 29)	'At this point, the notice to the CP is ready to be printed

'Enters worklist explaining to start redirection effective for the following month.
CALL navigate_to_PRISM_screen("CAWT")

PF5
EMWriteScreen "A", 3, 30
EMWriteScreen "FREE", 4, 37
EMWriteScreen "The redirection should be effective the 1st of next month", 10, 4
EMWriteScreen "30", 17, 52
Transmit

'________________________________________________________________________________________________________________________________________________________________________________________ CAREGIVER NOTICE

'goes to caregiver case
CALL navigate_to_PRISM_screen("CAST")

'Puts caregiver case number in from Dialog box
EMWriteScreen Left (caregiver_case_number, 10), 4, 8
EMWriteScreen Right (caregiver_case_number, 2), 4, 19

CALL write_value_and_transmit("D", 3, 29)

'goes to DORD
CALL navigate_to_PRISM_screen("DORD")

'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0502", 6, 36
Transmit

'entering user labels
PF14
CALL write_value_and_transmit("U", 20, 14)

FOR i = 7 to 18
	EMWriteScreen "S", i, 5
NEXT
Transmit

EMWriteScreen redirection_month, 16, 15
EMWriteScreen "/01/", 16, 17
EMWriteScreen redirection_year, 16, 21
Transmit

CALL write_value_and_transmit(original_cp_name, 16, 15)

CALL write_value_and_transmit(child_one, 16, 15)
CALL write_value_and_transmit(child_two, 16, 15)
CALL write_value_and_transmit(child_three, 16, 15)
CALL write_value_and_transmit(child_four, 16, 15)
CALL write_value_and_transmit(child_five, 16, 15)
CALL write_value_and_transmit(child_six, 16, 15)

CALL write_value_and_transmit(cch_amount, 16, 15)
CALL write_value_and_transmit(cms_amount, 16, 15)
CALL write_value_and_transmit(ccc_amount, 16, 15)
CALL write_value_and_transmit(total_amount, 16, 15)

PF3
CALL write_value_and_transmit("M", 3, 29)		'At this point, the notice to the caregiver is ready to be printed

'Enters worklist explaining to start redirection effective for the following month.
CALL navigate_to_PRISM_screen("CAWT")
PF5
EMWriteScreen "A", 3, 30
EMWriteScreen "FREE", 4, 37
EMWriteScreen "The redirection should be effective the 1st of next month", 10, 4
EMWriteScreen "30", 17, 52
Transmit

script_end_procedure("DORD docs are created but incomplete.  Please modify to select appropriate legal headings.")
