'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "generic-enforcement-intake.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog

BeginDialog CS_intake_dialog, 0, 0, 377, 266, "CS intake dialog"
  CheckBox 20, 30, 140, 10, "DORD F0999 - PIN Notice", NCP_PIN_Notice_Check
  CheckBox 20, 40, 150, 10, "DORD F0924 - Health Insurance Verification", NCP_health_ins_verif_check
  CheckBox 30, 80, 140, 10, "DORD F0100 - Auth to Collect Support", dord_F0100_check
  CheckBox 30, 100, 140, 10, "DORD F0109 - Notice of Parental Liability", dord_F0109_check
  CheckBox 30, 120, 140, 10, "DORD F0107 - Notice of Med Liability", dord_F0107_check
  CheckBox 20, 150, 130, 10, "Set File Location to QC 30", qc_30_file_loc_check
  CheckBox 20, 160, 120, 10, "Set File Location to SAFETY", safety_file_loc_check
  CheckBox 190, 30, 150, 10, "DORD F0999 - PIN Notice", CP_PIN_Notice_check
  CheckBox 190, 40, 160, 10, "DORD F0924 - Health Insurance Verification", CP_health_ins_verif_check
  CheckBox 190, 70, 110, 10, "10 day tickler to call NCP", t_10_day_tickler_check
  CheckBox 190, 90, 110, 10, "30 day tickler to load arrears", t_30_day_to_load_arrears_check
  CheckBox 190, 100, 110, 10, "30 day case review", t_30_day_case_review_check
  EditBox 200, 110, 140, 20, t_30_day_cawd_txt
  CheckBox 190, 130, 110, 10, "Create a FREE worklist", t_60_day_case_review_check
  EditBox 200, 140, 140, 20, t_60_day_cawd_txt
  EditBox 240, 180, 110, 20, worker_name
  EditBox 240, 200, 110, 20, worker_phone
  EditBox 90, 230, 90, 20, worker_signature
  ButtonGroup ButtonPressed
    OkButton 220, 230, 50, 20
    CancelButton 270, 230, 50, 20
  Text 10, 240, 70, 10, "Sign your CAAD note:"
  Text 20, 70, 70, 10, "NPA, DWP"
  Text 20, 90, 100, 10, "MFIP, CCA"
  Text 20, 110, 90, 10, "MA only"
  Text 180, 200, 50, 10, "Worker phone:"
  Text 180, 180, 50, 10, "Worker name:"
  GroupBox 0, 140, 170, 30, "File Location on CAST"
  GroupBox 0, 20, 170, 110, "Letters to NCP"
  GroupBox 180, 20, 180, 40, "Letters to CP"
  Text 0, 0, 170, 20, "Enforcement Intake Script "
  GroupBox 180, 60, 170, 110, "CAWD notes to add"
  EditBox 10, 210, 120, 20, add_caad_txt
  Text 0, 200, 90, 10, "Additional CAAD note text"
  GroupBox 10, 60, 160, 70, "Liability Notice to NCP"
  DropListBox 70, 180, 90, 20, "M2123"+chr(9)+"E0001", caad_type
  Text 0, 180, 70, 10, "Select CAAD type:"
EndDialog


'CUSTOM FUNCTION***************************************************************************************************************


FUNCTION send_dord_doc(recipient, dord_doc)
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen dord_doc, 6, 36
	EMWriteScreen recipient, 11, 51
	transmit
END FUNCTION

'This is a custom function to fix data that we are reading from PRISM that includes underscores.  The parameter is a string for the
'variable to be searched.  The function searches the variable and removes underscores.  Then, the fix case function is called to format
'the string in the correct case.  Finally, the data is trimmed to remove any excess spaces.
FUNCTION fix_read_data (search_string)
	search_string = replace(search_string, "_", "")
	call fix_case(search_string, 1)
	search_string = trim(search_string)
	fix_read_data = search_string 'To make this a return function, this statement must set the value of the function name
END FUNCTION

' This is a custom function to change the format of a participant name.  The parameter is a string with the
' client's name formatted like "Levesseur, Wendy K", and will change it to "Wendy K LeVesseur".
FUNCTION change_client_name_to_FML(client_name)
	client_name = trim(client_name)
	length = len(client_name)
	position = InStr(client_name, ", ")
	last_name = Left(client_name, position-1)
	first_name = Right(client_name, length-position-1)
	client_name = first_name & " " & last_name
	client_name = lcase(client_name)
	call fix_case(client_name, 1)
	change_client_name_to_FML = client_name 'To make this a return function, this statement must set the value of the function name
END FUNCTION

'This is a custom function to change the file location on the CAST screen
FUNCTION set_file_loc_on_CAST(new_file_location)
	call navigate_to_PRISM_screen("CAST")
	EMWriteScreen "M", 3, 29
	EMWriteScreen new_file_location, 14, 17
	transmit
END FUNCTION

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Finds the PRISM case number using a custom function
call PRISM_case_number_finder(PRISM_case_number)

'Shows case number dialog
Do
	Do
		Dialog PRISM_case_number_dialog
		If buttonpressed = 0 then stopscript
		call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	Loop until case_number_valid = True
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"

'Clearing case info from PRISM
call navigate_to_PRISM_screen("REGL")
transmit

'Navigating to CAPS
call navigate_to_PRISM_screen("CAPS")

'Entering case number and transmitting
EMSetCursor 4, 8
EMSendKey replace(PRISM_case_number, "-", "")									'Entering the specific case indicated
EMWriteScreen "d", 3, 29												'Setting the screen as a display action
transmit															'Transmitting into it

'Getting worker info for case note
EMSetCursor 5, 53
PF1
EMReadScreen worker_name, 27, 6, 50
EMReadScreen worker_phone, 12, 8, 35
PF3

Do
	EMReadScreen child_name_on_CAPS, 30, PRISM_row, 16	'reading name
	child_name_on_CAPS = trim(child_name_on_CAPS)		'removing spaces from beginning and end
	EMReadScreen child_DOB_on_CAPS, 10, PRISM_row, 64	'reading DOB
	If child_name_on_CAPS <> "" then CAPS_kids = CAPS_kids & child_name_on_CAPS & " (DOB: " & child_DOB_on_CAPS & ")" & chr(13) 		'If there's a name, add to the CAPS_kids variable
	PRISM_row = PRISM_row + 1					'increase the PRISM row
	If PRISM_row = 21 then						'If we're on row 21, go to the next page
		PF8
		PRISM_row = 18
	End if
Loop until child_name_on_CAPS = ""

'Cleaning up worker info
worker_name = trim(worker_name)
call fix_case(worker_name, 1)
worker_name = change_client_name_to_FML(worker_name)


'Shows intake dialog, checks to make sure we're still in PRISM (not passworded out)
Do
	Dialog CS_intake_dialog
	If buttonpressed = 0 then stopscript
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"


'If F0999 is indicated on the dialog then it navigates to DORD to send it.
If NCP_PIN_Notice_Check = checked then 'Send PIN Notice
	call send_dord_doc("NCP", "F0999")
End if

'If F0924 is indicated on the dialog then it navigates to DORD to send it.
If NCP_health_ins_verif_check = checked then
	call send_dord_doc("NCP", "F0924")
End if

'If F0100 is indicated on the dialog then it navigates to DORD to send it.
If dord_F0100_check = checked then
	call send_dord_doc("NCP", "F0100")
End if

'If F0109 is indicated on the dialog then it navigates to DORD to send it.
If dord_F0109_check = checked then
	call send_dord_doc("NCP", "F0109")
End if

'If F0107 is indicated on the dialog then it navigates to DORD to send it.
If dord_F0107_check = checked then
	call send_dord_doc("NCP", "F0107")
End if
'If F0924 is indicated on the dialog then it navigates to DORD to send it.
If CP_health_ins_verif_check = checked then
	call send_dord_doc("CPP", "F0924")
End if
'If F0999 is indicated on the dialog then it navigates to DORD to send it.
If CP_PIN_Notice_check = checked then
	call send_dord_doc("CPP", "F0999")
End if

'************************Change File Location on Cast

If qc_30_file_loc_check = checked then
	set_file_loc_on_CAST("QC 30")
End if


if safety_file_loc_check = checked then
	set_file_loc_on_CAST("Safety")
End if


'**************************Add worklists

If t_10_day_tickler_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "*** Call NCP to answer any questions NCP has about case setup.", 10, 4
	EMWriteScreen dateadd("d", date, 10), 17, 21
	transmit
End if
If t_30_day_to_load_arrears_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "Load arrears?", 10, 4
	EMWriteScreen dateadd("d", date, 30), 17, 21
	transmit
End if
If t_30_day_case_review_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "30 Day Case Review", 10, 4
	EMWriteScreen t_30_day_cawd_txt, 11, 4
	EMWriteScreen dateadd("d", date, 30), 17, 21
	transmit
End if
If t_60_day_case_review_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen t_60_day_cawd_txt, 10, 4
	EMWriteScreen dateadd("d", date, 60), 17, 21
	transmit
End if

'**********************************




'Going to CAAD, adding a new note
call navigate_to_PRISM_screen("CAAD")
PF5
EMReadScreen case_activity_detail, 20, 2, 29
If case_activity_detail <> "Case Activity Detail" then script_end_procedure("The script could not navigate to a case note. There may have been a script error. Add case note manually, and report the error to a script writer.")

'Setting the type
EMWriteScreen caad_type, 4, 54

'Setting cursor in write area and writing note details
EMSetCursor 16, 4
	call write_variable_in_CAAD(add_caad_txt)
	call write_variable_in_CAAD("* The following documents were sent:")
	If NCP_PIN_Notice_check = checked then call write_variable_in_CAAD("    * F0999 - PIN Notice to NCP")
	If NCP_health_ins_verif_check = checked then call write_variable_in_CAAD("    * F0924 - Health Insurance Verification to NCP")
	If dord_F0100_check = checked then call write_variable_in_CAAD("    * F0100 sent to NCP")
	If dord_F0109_check = checked then call write_variable_in_CAAD("    * F0109 sent to NCP")
	If dord_F0107_check = checked then call write_variable_in_CAAD("    * F0107 sent to NCP")
	If CP_PIN_Notice_check = checked then call write_variable_in_CAAD("    * F0999 - PIN Notice to CP")
	If CP_health_ins_verif_check = checked then call write_variable_in_CAAD("    * F0924 - Health Insurance Verification to CP")
	call write_variable_in_CAAD("---")
	call write_variable_in_CAAD("* The following worklists created:")
	If t_10_day_tickler_check = checked then call write_variable_in_CAAD("    * 10 day tickler to call NCP")
	If t_30_day_to_load_arrears_check = checked then call write_variable_in_CAAD("    * 30 day tickler to load arrears")
	If t_30_day_case_review_check = checked then call write_variable_in_CAAD("    * 30 day case review")
	If t_60_day_case_review_check = checked then call write_variable_in_CAAD("    * FREE worklist")
	call write_variable_in_CAAD("---")
	call write_variable_in_CAAD(worker_signature)
'	transmit

script_end_procedure("")
