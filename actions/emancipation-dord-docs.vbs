'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "emancipation-dord-docs.vbs"
start_time = timer
STATS_Counter = 1
STATS_manualtime = 120
STATS_denomination = "C"
'End of stats block-------------------------------------------------------------------------------------------------

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

'DIALOG--------------------------------------------------------
DIM EMC, CH_F, CPdord, CPfollowup, NCPdord, PRISM_case_number, ButtonPressed

BeginDialog EMC, 0, 0, 191, 160, "Emancipation Doc's"
  EditBox 55, 0, 70, 15, PRISM_case_number
  EditBox 55, 25, 75, 15, CH_F
  CheckBox 25, 65, 120, 10, "DORD Emancipation Notice to CP", CPdord
  CheckBox 25, 80, 160, 10, "DORD Emancipation Notice to CP - Follow-up", CPfollowup
  CheckBox 25, 115, 125, 10, "DORD Emancipation Notice to NCP", NCPdord
  ButtonGroup ButtonPressed
    OkButton 75, 140, 50, 15
    CancelButton 130, 140, 50, 15
  Text 5, 5, 50, 10, "Case Number"
  Text 5, 30, 45, 10, "Child's Name"
  GroupBox 10, 55, 175, 40, "CP DOCUMENTS"
  GroupBox 10, 105, 175, 30, "NCP DOCUMENTS"
EndDialog

'END DIALOG----------------------------------------------------

'connecting to bluezone
EMConnect ""

'checks o make sure we are in PRISM
CALL check_for_PRISM(True)

'brings me to the CAPS screen to auto fill prims case number in dialog
CALL navigate_to_PRISM_screen ("CAPS")
EMReadScreen PRISM_case_number, 13, 4, 8

'''find child that is going to be 18
DIM Row, Child_Active, Child_Age, CH_MCI, CH_M, CH_L, CH_S, Child_DOB, childs_name

	call navigate_to_PRISM_screen("CHDE")
	EMWriteScreen "B", 3, 29
	transmit


	'BEGINNING LOOP TO FIND CHILD
	Row = 8
	Do
		EMReadScreen Child_Active, 1, Row, 35
		If Child_Active = " " Then
			MsgBox "Unable to find child within 3 months of being 18 years old and up to age 19! Please process worklist manually! Script Ended.", VBExclamation
			StopScript
		ElseIf Child_Active = "Y" Then
			EMReadScreen Child_DOB, 8, Row, 57
			'CONFIRMING CHILD'S 18TH BIRTHDAY WILL BE IN THE NEXT 3 MOS but not over age 19
			'BY CALCULATING CHILD'S DOB FROM TODAY'S DATE (MUST BE BETWEEN 213 AND 229 MONTHS)
			Child_Age = DateDiff("m", Child_DOB, Date)
			If (Child_Age >= 213) And (Child_Age <= 229) Then
				EMReadScreen CH_MCI, 10, Row, 67
				Exit Do
			End If
		End If
	Row = Row + 1
	Loop


'Get  child's name to add to dialog boxes and word docs
call navigate_to_PRISM_screen("CHDE")
EMWriteScreen "D", 3, 29
EMWriteScreen CH_MCI, 4, 7
transmit
EMReadScreen CH_F, 12, 9, 34
EMReadScreen CH_M, 12, 9, 56
EMReadScreen CH_L, 17, 9, 8
EMReadScreen CH_S, 3, 9, 74
childs_name = fix_read_data(CH_F) & " " & fix_read_data(CH_M) & " " & fix_read_data(CH_L)
childs_name = trim(childs_name)


'end child info-------------------------------------------------------------------------------------

'this where arrears info was------------------------------------------------------------------------------------------------------------

CALL navigate_to_PRISM_screen ("CAPS")

'THE LOOP--------------------------------------------------------------------------
DIM err_msg

Do
	err_msg = ""
	Dialog EMC				'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF CH_F = "" THEN err_msg = err_msg & vbNewline & "Please enterd child's name."
		IF err_msg <> "" THEN
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

Loop until err_msg = ""

'END LOOP-------------------------------------------------------------------------


'SENDING DORD DOCS--------------------------------------------------------------
DIM Child_Row, Child_Col

'send emc letter to CP Dord F0300
IF CPdord = 1 THEN
	CALL navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0300", 6, 36
	transmit
	Child_Row = 1
	Child_Col = 1
	EMSearch CH_MCI, Child_Row, Child_Col
	EMSetCursor Child_Row, Child_Col
	transmit
END IF

'send emc follow upletter to CP Dord F0306
IF CPfollowup = 1 THEN
	CALL navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	EMSendKey "<enter>"
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0306", 6, 36
	transmit
	Child_Row = 1
	Child_Col = 1
	EMSearch CH_MCI, Child_Row, Child_Col
	EMSetCursor Child_Row, Child_Col
	transmit
END IF

'send emc letter to NCP Dord F0302
IF NCPdord = 1 THEN
	CALL navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	EMSendKey "<enter>"
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0302", 6, 36
	transmit
	Child_Row = 1
	Child_Col = 1
	EMSearch CH_MCI, Child_Row, Child_Col
	EMSetCursor Child_Row, Child_Col
	transmit
END IF
'---------------------------------END DORD DOCS------------------------------------

script_end_procedure("")
