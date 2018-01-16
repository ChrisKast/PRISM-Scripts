'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "adjustments.vbs"
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

'VARIABLES THAT NEED DECLARING----------------------------------------------------------------------------------------------------
CAAD_note_check = checked

'Dialogs -------------------------------------------------------------------------------------------------------------------------

'case number dialog-
BeginDialog case_number_dialog, 0, 0, 176, 85, "Case number dialog"
  EditBox 60, 5, 75, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 70, 65, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 50, 10, "Case number:"
EndDialog

'Adjustment Dialog-
BeginDialog Adjustment_Dialog, 0, 0, 216, 235, "Adjustment(s)"
  DropListBox 75, 5, 110, 10, "Please Select One:"+chr(9)+"Arrears Management"+chr(9)+"Direct Support"+chr(9)+"Error"+chr(9)+"Forgiveness"+chr(9)+"Interest Adjustment"+chr(9)+"Order"+chr(9)+"Other"+chr(9)+"Overpayment", Reason_List
  CheckBox 70, 40, 30, 10, "CCC", CCC_Obli_checkbox
  CheckBox 70, 50, 30, 10, "CCH", CCH_Obli_checkbox
  CheckBox 70, 60, 30, 10, "CMI", CMI_Obli_checkbox
  CheckBox 70, 70, 30, 10, "CMS", CMS_Obli_checkbox
  CheckBox 70, 80, 30, 10, "CSP", CSP_Obli_checkbox
  CheckBox 70, 90, 25, 10, "CUF", CUF_Obli_checkbox
  CheckBox 110, 40, 30, 10, "JCC", JCC_Obli_checkbox
  CheckBox 110, 50, 30, 10, "JCH", JCH_Obli_checkbox
  CheckBox 110, 60, 30, 10, "JME", JME_Obli_checkbox
  CheckBox 110, 70, 30, 10, "JMI", JMI_Obli_checkbox
  CheckBox 110, 80, 30, 10, "JMS", JMS_Obli_checkbox
  CheckBox 145, 40, 30, 10, "Other", Other_Obli_checkbox
  EditBox 85, 115, 50, 15, start_date
  EditBox 155, 115, 50, 15, end_date
  EditBox 85, 145, 50, 15, Amount_Adjusted
  EditBox 50, 170, 150, 15, Descrip_Box
  EditBox 70, 190, 130, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 100, 215, 50, 15
    CancelButton 155, 215, 50, 15
  Text 5, 30, 65, 10, "Affected Obligations"
  Text 5, 120, 75, 10, "Date Range (optional):"
  Text 140, 120, 10, 10, "TO"
  Text 5, 150, 75, 10, "Total Amount Adjusted:"
  Text 5, 175, 40, 10, "Description:"
  Text 5, 195, 60, 10, "Worker Signature:"
  Text 5, 10, 65, 10, "Adjustment Reason"
EndDialog



'Connecting to BlueZone
EMConnect ""

'Brings Bluezone to the Front
EMFocus

'Makes sure you are not passworded out
CALL check_for_PRISM(True)

call PRISM_case_number_finder(PRISM_case_number)

'Case number display dialog
Do
	Dialog case_number_dialog
	If buttonpressed = 0 then stopscript
	call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
	If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
Loop until case_number_valid = True


'Displays dialog for adjustments and checks for information
Do
	err_msg = ""
	'Shows dialog, validates that PRISM is up and not timed out, with transmit
	Dialog Adjustment_Dialog
	If buttonpressed = 0 then stopscript
	If Reason_List = "Please Select One:" THEN err_msg = err_msg & vbNewline & "Adjustment REASON must be completed."
	If CCC_Obli_checkbox = 0 AND CCH_Obli_checkbox = 0 AND CMI_Obli_checkbox = 0 AND CMS_Obli_checkbox = 0 AND CSP_Obli_checkbox = 0 AND CUF_Obli_checkbox = 0 AND JCC_Obli_checkbox = 0 AND JCH_Obli_checkbox = 0 AND JME_Obli_checkbox = 0 AND JMI_Obli_checkbox =0 AND JMS_Obli_checkbox = 0 AND Other_Obli_checkbox = 0 THEN err_msg = err_msg & vbNewline & "You must check at least ONE obligation."
	If Amount_Adjusted = "" THEN err_msg = err_msg & vbNewline & "Adjustment AMOUNT must be completed."
	If worker_signature = "" THEN err_msg = err_msg & vbNewline & "Sign your CAAD note."
	If err_Msg <> "" THEN
				Msgbox "***NOTICE***" & vbcr & err_msg & vbNewline & vbNewline & "Please resolve for this script to continue."
	END IF
LOOP UNTIL err_msg = ""


'Cleaning up the case note for check boxes
If CCH_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CCH, ")
If CMS_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CMS, ")
If CMI_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CMI, ")
If CCC_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CCC, ")
If CSP_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CSP, ")
If CUF_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("CUF, ")
If JCH_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("JCH, ")
If JMS_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("JMS, ")
If JME_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("JME, ")
If JCC_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("JCC, ")
If JMI_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("JMI, ")
If Other_Obli_checkbox = checked then line_for_CAAD_note = line_for_CAAD_note & ("Other")
If right(line_for_CAAD_note, 2) = ", " then line_for_CAAD_note = left(line_for_CAAD_note, len(line_for_CAAD_note) - 2)


'Going to CAAD note
call navigate_to_PRISM_screen("CAAD")

'Entering case number
call enter_PRISM_case_number(PRISM_case_number, 20, 8)

PF5					'Did this because you have to add a new note
EMWriteScreen "FREE", 4, 54   'adds correct caad code
EMSetCursor 16, 4			'Because the cursor does not default to this location

''information to be added to CAAD note
CALL write_variable_in_CAAD (">>>Adjustments<<<")
CALL write_bullet_and_variable_in_CAAD ("Adjustment Reason", Reason_List)
CALL write_bullet_and_variable_in_CAAD ("Total Amount Adjusted", "$" & Amount_Adjusted)
CALL write_bullet_and_variable_in_CAAD ("Affected Obligations", line_for_CAAD_note)
IF start_date <> "" and end_date <> "" THEN CALL write_bullet_and_variable_in_CAAD ("Date Range", start_date & "  to  " & end_date)
CALL write_bullet_and_variable_in_CAAD ("Description", Descrip_Box)
CALL write_variable_in_CAAD(worker_signature)

script_end_procedure("")

script_end_procedure("")
