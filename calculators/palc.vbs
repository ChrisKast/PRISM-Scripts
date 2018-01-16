'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "palc.vbs"
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
BeginDialog start_end_date_dialog, 0, 0, 171, 65, "Start and End Date Dialog"
  ButtonGroup ButtonPressed
    OkButton 120, 25, 50, 15
    CancelButton 120, 45, 50, 15
  EditBox 60, 5, 110, 15, PRISM_case_number
  EditBox 45, 25, 70, 15, start_date
  EditBox 45, 45, 70, 15, end_date
  Text 5, 10, 50, 10, "Case number:"
  Text 5, 30, 40, 10, "Start date:"
  Text 5, 50, 40, 10, "End date:"
EndDialog


'THE SCRIPT-----------------------------------

'Connecting to BlueZone
EMConnect ""

call PRISM_case_number_finder(PRISM_case_number)

Do
	Do
		Dialog start_end_date_dialog				'Shows dialog
		If buttonpressed = 0 then stopscript		'Cancel
		call PRISM_case_number_validation(PRISM_case_number, case_number_is_valid)
		If case_number_is_valid = False then MsgBox "Your case number isn't valid. Try again."
	Loop until case_number_is_valid = True
	If isdate(start_date) = False or isdate(end_date) = False then MsgBox "You must enter valid dates for both the start and end dates."		'Because a date for both fields is required
Loop until isdate(start_date) = True and isdate(end_date) = True



'Checks to make sure PRISM isn't locked out
CALL check_for_PRISM(false)

'Clearing case info from PRISM
call navigate_to_PRISM_screen("REGL")
transmit

'Navigating to PALC
call navigate_to_PRISM_screen("PALC")

'Entering case number and transmitting
EMSetCursor 20, 9
EMSendKey replace(PRISM_case_number, "-", "")	 	'Entering the specific case indicated

EMWriteScreen cstr(start_date), 20, 35
EMWriteScreen cstr(end_date), 20, 49
transmit								'Transmitting into it


row = 9		'Setting variable for the do...loop

Do
	EMReadScreen end_of_data_check, 19, row, 28									'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do							'Exits do if we have

	'Reading payment date, which for some crazy reason is YYMMDD, without slashes. This converts.
	EMReadScreen pmt_ID_YY, 2, row, 7
	EMReadScreen pmt_ID_MM, 2, row, 9
	EMReadScreen pmt_ID_DD, 2, row, 11
	pmt_ID_date = pmt_ID_MM & "/" & pmt_ID_DD & "/" & pmt_ID_YY

		EMReadScreen proc_type, 3, row, 25														'Reading the proc type
		EMReadScreen case_alloc_amt, 10, row, 70													'Reading the amt allocated
		IF case_alloc_amt = "          " THEN case_alloc_amt = 0
		If proc_type = "FTS" or proc_type = "MCE" or proc_type = "NOC" or proc_type = "IFC" or proc_type = "OST" or _
		proc_type = "PCA" or proc_type = "PIF" or proc_type = "STJ" or proc_type = "STS" or proc_type = "FTJ" or proc_type = "FIN"  then 		'If proc type is one of these, it's involuntary. Else, it's voluntary.
			total_involuntary_alloc = total_involuntary_alloc + abs(case_alloc_amt)							'Adds the alloc amt for involuntary
		Else
			total_voluntary_alloc = total_voluntary_alloc + abs(case_alloc_amt)							'Adds the alloc amt for voluntary
		End if

	row = row + 1														'Increases the row variable by one, to check the next row
	EMReadScreen end_of_data_check, 19, row, 28									'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do							'Exits do if we have
	If row = 19 then														'Resets row and PF8s
		PF8
		row = 9
	End if
Loop until end_of_data_check = "*** End of Data ***"

If total_involuntary_alloc = "" then total_involuntary_alloc = "0"
If total_voluntary_alloc = "" then total_voluntary_alloc = "0"

string_for_msgbox = "---PAYMENT BREAKDOWN FOR " & start_date & " THROUGH " & end_date & "---" & chr(10) & chr(10) & "Involuntary: $" & total_involuntary_alloc & chr(10) & "Voluntary: $" & total_voluntary_alloc

MsgBox string_for_msgbox

script_end_procedure("")
