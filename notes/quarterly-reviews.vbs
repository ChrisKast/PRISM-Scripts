'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "quarterly-reviews.vbs"
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

'VARIABLES TO DECLARE----------------------------------------------------------------------------------------------------
days_to_push_out_worklist = "90"	'This is the default

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog quarterly_reviews_dialog, 0, 0, 176, 85, "Quarterly Reviews Dialog"
  EditBox 60, 5, 75, 15, PRISM_case_number
  EditBox 140, 25, 35, 15, days_to_push_out_worklist
  EditBox 70, 45, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 65, 50, 15
    CancelButton 125, 65, 50, 15
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 130, 10, "Days to push out worklist (default is 90):"
  Text 5, 50, 60, 10, "Worker signature:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connecting to BlueZone
EMConnect ""

'Sends a transmit to check for password issues
transmit

'Checking to make sure we're on USWT or USWD. If not the script will stop.
EMReadScreen worklist_check, 3, 21, 75
If worklist_check <> "USW" and worklist_check <> "CAW" then script_end_procedure("Worklist screen not found. Please start this script from the worklist you are trying to copy over.")

'Searches for the case number.
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if

Do
	Do
		Do
			dialog quarterly_reviews_dialog
			If buttonpressed = 0 then stopscript
			call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
			If case_number_valid = False then MsgBox("Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''")
		Loop until case_number_valid = True
		If isnumeric(days_to_push_out_worklist) = False then MsgBox ("You must put a number in for the days to push out worklist.")
	Loop until isnumeric(days_to_push_out_worklist) = True


	EMReadScreen worklist_line_01, 72, 10, 4			'Reads worklist info, line by line
	EMReadScreen worklist_line_02, 72, 11, 4
	EMReadScreen worklist_line_03, 72, 12, 4
	EMReadScreen worklist_line_04, 72, 13, 4
	EMWriteScreen "__________", 17, 21				'clearing out worklist date
	EMWriteScreen days_to_push_out_worklist, 17, 52		'Adding the number of days to push out worklist
	EMWriteScreen "m", 3, 30					'Must modify the panel
	transmit
	call navigate_to_PRISM_screen("CAAD")
	pf5
	EMReadScreen case_activity_detail, 20, 2, 29
	If case_activity_detail <> "Case Activity Detail" then MsgBox "The script could not navigate to a case note. You might be locked out of your case. Navigate to a blank case note and try again."
Loop until case_activity_detail = "Case Activity Detail"

EMWriteScreen worklist_line_01, 16, 4
EMWriteScreen worklist_line_02, 17, 4
EMWriteScreen worklist_line_03, 18, 4
EMWriteScreen worklist_line_04, 19, 4
EMWriteScreen "------" & worker_signature, 20, 4
EMWriteScreen "E0002", 4, 54

script_end_procedure("")
