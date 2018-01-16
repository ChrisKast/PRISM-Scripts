'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ddpl.vbs"
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

'Declared
DIM ddpl_calculator, PRISM_MCI_number, PRISM_begin_date, PRISM_end_date, buttonpressed, row, direct_deposit_issued_date, end_of_data_check, direct_deposit_amount, end_date, total_amount_issued, string_for_msgbox

'DDPL Dialog Box
BeginDialog ddpl_calculator, 0, 0, 191, 105, "DDPL Calculator"
  ButtonGroup ButtonPressed
    OkButton 80, 80, 50, 15
    CancelButton 135, 80, 50, 15
  Text 15, 10, 65, 10, "PRISM MCI Number"
  EditBox 95, 5, 60, 15, PRISM_MCI_number
  Text 35, 30, 50, 10, "Start Date"
  EditBox 95, 25, 50, 15, PRISM_begin_date
  Text 35, 50, 50, 10, "End Date"
  EditBox 95, 45, 50, 15, PRISM_end_date
EndDialog

Dialog ddpl_calculator

IF ButtonPressed = cancel THEN StopScript

EMConnect ""

CALL check_for_prism(TRUE)

CALL navigate_to_prism_screen ("DDPL")

EMWriteScreen PRISM_MCI_number, 20, 007

EMSendKey "<enter>"

EMWaitReady 0,0

EMWriteScreen PRISM_begin_date, 20, 038

EMWriteScreen PRISM_end_date, 20, 067

transmit

row = 8

total_amount_issued = 0

Do

EMReadScreen end_of_data_check, 19, row, 28 					'Checks to see if we've reached the end of the list
	If end_of_data_check = "*** End of Data ***" then exit do 		'Exits do if we have
EMReadScreen direct_deposit_issued_date, 9, row, 11 				'Reading the issue date
EMReadScreen direct_deposit_amount, 10, row, 33 				'Reading amount issued

total_amount_issued = total_amount_issued + abs(direct_deposit_amount) 	'Totals amount issued

row = row + 1 										'Increases the row variable by one, to check the next row

EMReadScreen end_of_data_check, 19, row, 28 					'Checks to see if we've reached the end of the list
    If end_of_data_check = "*** End of Data ***" then exit do 		'Exits do if we have

    If row = 19 then 									'Resets row and PF8s
        PF8
        row = 8
    End if
Loop until end_of_data_check = "*** End of Data ***"

string_for_msgbox = " Total payments issued for the period of " & PRISM_begin_date & " through " & PRISM_end_date & " is $" & total_amount_issued

MsgBox string_for_msgbox
script_end_procedure("")
