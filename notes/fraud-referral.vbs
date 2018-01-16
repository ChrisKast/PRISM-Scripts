'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "fraud-referral.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 120
STATS_denomination = "C"
'END OF STATS BLOCK-------------------------------------------------------------------------------------------------

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

'THE SCRIPT-----------------------------------------------------------------------------------------------------------------------
'THE DIALOG---------------------------------------------------

BeginDialog fraud_referral_dialog, 0, 0, 301, 280, "Fraud Referral"
  EditBox 80, 5, 85, 15, prism_case_number
  EditBox 130, 30, 160, 15, ref_reason
  EditBox 115, 75, 145, 15, case_num
  EditBox 115, 95, 145, 15, cp_fw
  CheckBox 30, 130, 25, 15, "CCC", ccc_checkbox
  CheckBox 65, 130, 30, 15, "DWP", dwp_checkbox
  CheckBox 105, 130, 25, 15, "FCC", fcc_checkbox
  CheckBox 140, 130, 30, 15, "MAO", ma_checkbox
  CheckBox 175, 130, 25, 15, "MFP", mfp_checkbox
  DropListBox 70, 165, 125, 15, "Select One..."+chr(9)+"Email"+chr(9)+"Telephone Call"+chr(9)+"Fax"+chr(9)+"Mail", ref_droplistbox
  EditBox 70, 185, 125, 15, sent_to
  CheckBox 10, 215, 205, 10, "Check here to create a worklist to check on referral status in", worklist_checkbox
  EditBox 215, 210, 20, 15, days_out
  EditBox 75, 235, 135, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 165, 260, 50, 15
    CancelButton 220, 260, 50, 15
  Text 10, 10, 70, 10, "Prism Case Number:"
  Text 10, 240, 60, 10, "Worker Signature:"
  Text 35, 100, 75, 10, "CP's Financial Worker:"
  Text 10, 35, 120, 10, "Reason for sending Fraud Referral:"
  Text 10, 165, 60, 10, "Referral Sent Via:"
  Text 20, 80, 95, 10, "MAXIS/METS Case Number:"
  GroupBox 5, 55, 285, 95, "Financial Assistance Case Details"
  Text 240, 215, 20, 10, "days"
  Text 10, 190, 60, 10, "Referral Sent To:"
  Text 20, 115, 70, 10, "PA Program Open:"
EndDialog


'Connects to BLUEZONE
EMConnect ""

'MAKES SURE YOU ARE NOT PASSWORDED OUT
CALL check_for_PRISM(True)

'AUTO POPULATES PRISM CASE NUMBER INTO DIALOG
CALL PRISM_case_number_finder(PRISM_case_number)

'MAKES THINGS MANDATORY
DO
	err_msg = ""
	Dialog fraud_referral_dialog
	cancel_confirmation
	CALL Prism_case_number_validation(PRISM_case_number, case_number_valid)
	IF case_number_valid = FALSE THEN err_msg = err_msg & vbNewline & "You must enter a valid PRISM case number!"
	IF ref_reason = "" THEN err_msg = err_msg & vbNewline & "You must enter a reason for the fraud referral!"
	IF case_num = "" THEN err_msg = err_msg & vbNewline & "You must enter a MAXIS/METS case number!"
	IF ref_droplistbox = "Select One..." THEN err_msg = err_msg & vbNewline & "You must select how you sent the Referral!"
	IF sent_to = "" THEN err_msg = err_msg & vbNewline & "You must enter who the Referral was sent to!"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "You sign your CAAD note!"
	IF err_msg <> "" THEN MsgBox "***NOTICE***" & vbNewLine & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue!"
LOOP UNTIL err_msg = ""

'NAVIGATES TO CAAD
CALL navigate_to_PRISM_screen("CAAD")

'ENTERING CASE NUMBER
CALL enter_PRISM_case_number(PRISM_case_number, 20, 8)


'MAKING THE PA PROGRAM OPEN NEATER
IF ccc_checkbox = checked THEN prog_open = prog_open & "CCC,"
IF dwp_checkbox = checked THEN prog_open = prog_open & " DWP,"
IF fcc_checkbox = checked THEN prog_open = prog_open & " FCC,"
IF ma_checkbox = checked THEN prog_open = prog_open & " MAO,"
IF mfp_checkbox = checked THEN prog_open = prog_open & " MFP,"
prog_open = trim(prog_open)
IF right(prog_open, 1) = "," THEN prog_open = left(prog_open, len(prog_open) - 1)


'ADDS NEW CAAD NOTE WITH FREE CAAD CODE
PF5
EMWritescreen "M0010", 4, 54

'SETS THE CURSOR
EMSetCursor 16, 4

'WRITES THE CAAD NOTE
CALL write_bullet_and_variable_in_CAAD("Reason for sending Fraud Referral", ref_reason)
CALL write_bullet_and_variable_in_CAAD("MAXIS/METS Case Number", case_num)
CALL write_bullet_and_variable_in_CAAD("CP's financial worker", cp_fw)
CALL write_bullet_and_variable_in_CAAD("PA Program Open", prog_open)
CALL write_bullet_and_variable_in_CAAD("Referral Sent Via", ref_droplistbox)
CALL write_bullet_and_variable_in_CAAD("Referral Sent to", sent_to)
CALL write_variable_in_CAAD(worker_signature)
transmit


'ADDS A WORKLIST IF THE CHECKBOX TO ADD ONE IS CHECKED
IF worklist_checkbox = CHECKED THEN
	CALL navigate_to_PRISM_screen("CAWT")
	PF5
	EMWritescreen "FREE", 4, 37

	'SETS THE CURSOR AND STARTS THE WORKLIST
	EMSetCursor 10, 4
	EMWriteScreen "Check status of Fraud Referral made", 10, 4
	EMSetCursor 11, 4
	CALL write_bullet_and_variable_in_CAAD("Reason for Referral", ref_reason)
	'EMSetCursor 17, 52
	EMWritescreen days_out, 17, 52

END IF

script_end_procedure("")
