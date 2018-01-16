''GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "~utilities-menu.vbs"
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

'Loading all scripts
If run_locally <> true then
	If use_master_branch = TRUE then
		CALL run_from_GitHub("https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/~complete-list-of-scripts.vbs")
	Else
		CALL run_from_GitHub("https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/release/~complete-list-of-scripts.vbs")
	End if
Else
	CALL run_from_GitHub("C:\DHS-PRISM-Scripts\~complete-list-of-scripts.vbs")
End if

DIM ButtonPressed, button_placeholder
DIM SIR_instructions_button
DIM Dialog1

Function declare_main_menu(menu_type, script_array)

	'Figures out how tall to make the dialog
	FOR current_script = 0 to ubound(script_array)
		IF InStr(script_array(current_script).category, menu_type) <> 0 THEN scripts_to_display = scripts_to_display + 1
	Next

	'If not declared elsewhere in the script...		...set variable default
	If vert_button_position = "" then 				vert_button_position = 10			'Where the buttons start
	If button_height = "" then  					button_height = 10					'Height of each button
	If button_spacing = "" then 					button_spacing = 5					'Spacing between each button
	If cancel_button_height = "" then 				cancel_button_height = 15			'Height of the cancel button
	If button_width = "" then						button_width = 120					'Width of the buttons
	If description_text_width = "" then				description_text_width = 360		'Width of description text
	
	'Doing height/width calculations
	pixels_needed_for_height = (scripts_to_display * (button_height + button_spacing)) + vert_button_position + (button_spacing + cancel_button_height)
	pixels_needed_for_width = button_width + description_text_width + 20
	
	'Displays the dialog
	BeginDialog Dialog1, 0, 0, pixels_needed_for_width, pixels_needed_for_height, ucase(menu_type) & " SCRIPTS MENU"
	  ButtonGroup ButtonPressed
	    CancelButton pixels_needed_for_width - 55, pixels_needed_for_height - 20, 50, cancel_button_height			'Puts this first to avoid accidental tabbing
	 	
		button_placeholder = 100
		FOR current_script = 0 to ubound(script_array)
			IF InStr(script_array(current_script).category, menu_type) <> 0 THEN
				'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
				'FUNCTION		HORIZ. ITEM POSITION			VERT. ITEM POSITION		ITEM WIDTH					ITEM HEIGHT				ITEM TEXT/LABEL										BUTTON VARIABLE
				PushButton 		5, 								vert_button_position, 	button_width,				button_height, 			script_array(current_script).script_name, 			button_placeholder
				Text 			button_width + 10,				vert_button_position, 	description_text_width,		button_height, 			"--- " & script_array(current_script).description
				'----------
				vert_button_position = vert_button_position + button_height + button_spacing 'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
				'----------
				script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
			END IF
			button_placeholder = button_placeholder + 1
		NEXT
	EndDialog
	
End function

DO
	CALL declare_main_menu("utilities", cs_scripts_array)
	Dialog
	IF ButtonPressed = 0 THEN script_end_procedure("")

LOOP UNTIL ButtonPressed <> SIR_instructions_button

'Determining the script selected from the value of ButtonPressed
'Since we start at 100 and then go up, we will simply subtract 100 when determining the position in the array
script_picked = ButtonPressed - 100

'Running the selected script
CALL run_from_GitHub(script_repository & cs_scripts_array(script_picked).category & "/" & cs_scripts_array(script_picked).file_name)
