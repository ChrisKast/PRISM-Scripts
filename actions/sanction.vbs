'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "sanction.vbs"
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

BeginDialog sanction_dialog, 0, 0, 277, 336, "Paternity Sanction"
  DropListBox 10, 20, 220, 20, "Select One"+chr(9)+"Return Financial Statement"+chr(9)+"Return General Testimony and Petition"+chr(9)+"Return General Testimony, Petition, and Aff"+chr(9)+"Return Locate Request or Provide Info"+chr(9)+"Return PIF, Request Sheet, Assessment, and Fin Stmt"+chr(9)+"Attend Genetic Test Appointment"+chr(9)+"Attend CAO Appointment"+chr(9)+"Provide Requested Info", sanction_reason
  DropListBox 130, 40, 100, 20, "Select One"+chr(9)+"Traci Melberg"+chr(9)+"Carrie Freeland"+chr(9)+"Andrea Hesse", CAO_contact


  CheckBox 20, 80, 140, 10, "Enter CAAD note", caad_noncoop_check
  CheckBox 20, 90, 140, 10, "Update GCSC for non-cooperation", gcsc_noncoop_check
  CheckBox 20, 100, 140, 10, "Update CAST file location to SANC", cast_noncoop_check
  CheckBox 20, 110, 170, 10, "Add FREE worklist to send reminder in 28 days", CAWD_noncoop_check
  CheckBox 20, 120, 230, 10, "Create WORD memo to FAS and/or CCA worker re: noncooperation", FAS_or_CCA_Memo_noncoop_check
  CheckBox 20, 130, 170, 10, "Send DORD Memo to CP", sanction_memo_check
  CheckBox 20, 180, 140, 10, "Enter CAAD note", caad_coop_check
  CheckBox 20, 190, 140, 10, "Update GCSC for cooperation", gcsc_coop_check
  CheckBox 20, 200, 140, 10, "Update CAST file location ", cast_coop_check
  CheckBox 20, 210, 170, 10, "Remove reminder worklist", CAWD_coop_check
  CheckBox 20, 220, 210, 10, "Create WORD memo to FAS and/or CCA worker re: cooperation", FAS_or_CCA_Memo_coop_check
  EditBox 60, 250, 190, 20, Maxis_number
  EditBox 60, 270, 190, 20, CAAD_note
  EditBox 60, 290, 190, 20, worker_signature
  ButtonGroup ButtonPressed
    PushButton 10, 60, 130, 20, "Toggle All Sanction Applied Checkboxes", check_sanction_applied_check
    PushButton 10, 160, 130, 20, "Toggle All Sanction Cured Checkboxes", check_sanction_cured_check
    OkButton 70, 310, 40, 20
    CancelButton 140, 310, 40, 20
  GroupBox 0, 0, 250, 140, "Sanction Applied"
  Text 10, 10, 150, 10, "Sanction Reason:"
  GroupBox 0, 150, 250, 90, "Sanction Cured"
  Text 0, 250, 50, 10, "MAXIS number:"
  Text 0, 270, 50, 10, "CAAD notes:"
  Text 0, 290, 50, 10, "Worker initials:"
  Text 10, 40, 110, 10, "CAO contact person, if applicable:"
EndDialog


'CUSTOM FUNCTION----------------------------------------------------------------------------------------------
'This is a custom function to change the file location on the CAST screen
FUNCTION set_file_loc_on_CAST(new_file_location)
	call navigate_to_PRISM_screen("CAST")
	EMWriteScreen "M", 3, 29
	EMWriteScreen new_file_location, 14, 17
	transmit
END FUNCTION

'This is a custom function to update cooperation on the GCSC screen
FUNCTION create_gcsc_update(cooperation_code, comments)
	call navigate_to_PRISM_screen("GCSC")
	EMWriteScreen "M", 3, 29
	EMWriteScreen date, 9, 18
	EMWriteScreen cooperation_code, 15, 25
	EMWriteScreen comments, 19, 3
	transmit
END FUNCTION

FUNCTION write_value_and_transmit(text, row, col)
	EMWriteScreen text, row, col
	transmit
END FUNCTION

FUNCTION send_text_to_DORD(string_to_write, recipient)
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0104", 6, 36
	EMWriteScreen recipient, 11, 51
	transmit

	'This function will add a string to DORD docs.
	IF len(string_to_write) > 1080 THEN
		MsgBox "*** NOTICE!!! ***" & vbCr & vbCr & _
				"The text below is longer than the script can handle in one DORD document. The script will not add the text to the document." & vbCr & vbCr & _
				string_to_write
		EXIT FUNCTION
	END IF

	dord_rows_of_text = Int(len(string_to_write) / 60) + 1

	ReDim write_array(dord_rows_of_text)
	'Splitting the text
	string_to_write = split(string_to_write)
	array_position = 1
	FOR EACH word IN string_to_write
		IF len(write_array(array_position)) + len(word) <= 60 THEN
			write_array(array_position) = write_array(array_position) & word & " "
		ELSE
			array_position = array_position + 1
			write_array(array_position) = write_array(array_position) & word & " "
		END IF
	NEXT

	PF14

	'Selecting the "U" label type
	CALL write_value_and_transmit("U", 20, 14)

	'Writing the values
	dord_row = 7
	FOR i = 1 TO dord_rows_of_text
		CALL write_value_and_transmit("S", dord_row, 5)
		CALL write_value_and_transmit(write_array(i), 16, 15)

		dord_row = dord_row + 1
		IF i = 12 THEN
			PF8
			dord_row = 7
		END IF
	NEXT
	PF3
	EMWriteScreen "M", 3, 29
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
transmit
														'Transmitting into it
EMReadScreen CP_name, 30, 6, 12


'Get first child's name
call navigate_to_PRISM_screen("CHDE")
EMWriteScreen CH_MCI, 4, 7
transmit
EMReadScreen CH_F, 12, 9, 34
EMReadScreen CH_M, 12, 9, 56
EMReadScreen CH_L, 17, 9, 8
EMReadScreen CH_S, 3, 9, 74
childs_name = fix_read_data(CH_F) & " " & fix_read_data(CH_M) & " " & fix_read_data(CH_L)
If trim(CH_S) <> "" then childs_name = childs_Name & " " & ucase(fix_read_data(CH_S))

'Go back to CAPS for all the kids' info
call navigate_to_PRISM_screen("CAPS")
EMSetCursor 4, 8
EMSendKey replace(PRISM_case_number, "-", "")									'Entering the specific case indicated
EMWriteScreen "d", 3, 29												'Setting the screen as a display action
transmit
'Getting all child/DOB info
PRISM_row = 18
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

all_checked = false
'Shows dialog, checks to make sure we're still in PRISM (not passworded out)
DO
	error_msg = ""
	Dialog sanction_dialog
	If buttonpressed = 0 then stopscript


	If buttonpressed = check_sanction_applied_check then
		If all_checked = false then
			caad_noncoop_check = checked
			gcsc_noncoop_check = checked
			cast_noncoop_check = checked
			cawd_noncoop_check = checked
			FAS_or_CCA_Memo_noncoop_check = checked
			sanction_memo_check = checked

			If sanction_reason = "Select One" then
				error_msg = error_msg & "Please select a sanction reason.  "
			End If
			all_checked = true
		elseif all_checked = true then
			caad_noncoop_check = unchecked
			gcsc_noncoop_check = unchecked
			cast_noncoop_check = unchecked
			cawd_noncoop_check = unchecked
			FAS_or_CCA_Memo_noncoop_check = unchecked
			sanction_memo_check = unchecked
			error_msg = ""
			all_checked = false
		End if
	End if
	If buttonpressed = check_sanction_cured_check then
		If all_checked = false then
			caad_coop_check = checked
			gcsc_coop_check = checked
			cast_coop_check = checked
			cawd_coop_check = checked
			FAS_or_CCA_Memo_coop_check = checked
			error_msg = ""
			all_checked = true
		elseif all_checked = true then
			caad_coop_check = unchecked
			gcsc_coop_check = unchecked
			cast_coop_check = unchecked
			cawd_coop_check = unchecked
			FAS_or_CCA_Memo_coop_check = unchecked
			all_checked = true
		End if
	End if
	If sanction_reason = "Attend CAO Appointment" and CAO_contact = "Select One" then
		error_msg = error_msg & "Please specify the CAO contact the client needs to follow up with.  "
	End if
	if error_msg <> "" then
		msgbox error_msg & "Please resolve to continue."
	end if
Loop until buttonpressed <> check_sanction_cured_check and buttonpressed <> check_sanction_applied_check	and error_msg = ""


check_for_PRISM(false)

'Enter GCSC record
If gcsc_noncoop_check = checked then
	CALL create_gcsc_update("N", "Noncooperation reason: " & sanction_reason)
End If

If gcsc_coop_check = checked then
	CALL create_gcsc_update("Y", "CP is cooperating with child support services.")
End If



'Resetting file locations
If cast_noncoop_check = checked then
	set_file_loc_on_CAST("SANC")
End if

If cast_coop_check = checked then
	set_file_loc_on_CAST("     ")
End if

'Creating the Word application object (if any of the Word options are selected), and making it visible
If _
	FAS_or_CCA_Memo_noncoop_check = checked or _
	FAS_or_CCA_Memo_coop_check = checked then
		Set objWord = CreateObject("Word.Application")
		objWord.Visible = True
End if
If FAS_or_CCA_Memo_noncoop_check = checked then
	set objDoc = objWord.Documents.Add("L:\Child Support\Paternity\Notice of Noncooperation.dotm")
	With objDoc
		.FormFields("CP_name").Result = CP_name
		.FormFields("MAXIS_case_number").Result = MAXIS_number
		.FormFields("PRISM_case_number").Result = PRISM_case_number
		.FormFields("children").Result = CAPS_kids
	End With
End if
If FAS_or_CCA_Memo_coop_check = checked then
	set objDoc = objWord.Documents.Add("L:\Child Support\Paternity\Notice of Cooperation.dotm")
	With objDoc
		.FormFields("CP_name").Result = CP_name
		.FormFields("MAXIS_case_number").Result = MAXIS_number
		.FormFields("PRISM_case_number").Result = PRISM_case_number
		.FormFields("children").Result = CAPS_kids
	End With
End if

If sanction_memo_check = checked then
	memo_text = "*** NOTICE OF SANCTION *** Anoka County Child Support has requested that your public assistance benefits be sanctioned because you failed to "

	If sanction_reason = "Return Financial Statement" then
		memo_text = memo_text & "return the Financial Statement which was previously mailed to you.  The sanction will be removed when the completed Financial Statement is received."
	ELSEIF sanction_reason = "Return General Testimony and Petition" then
		memo_text = memo_text & "return the General Testimony and Uniform Support Petition forms which were previously mailed to you.  The sanction will be removed when the missing documents are received."
	ELSEIF sanction_reason = "Return General Testimony, Petition, and Aff" then
		memo_text = memo_text & "return the General Testimony, Uniform Support Petition, and Affidavit in Support of Establishing Paternity.  The saction will be removed when the missing documents are received."
	ELSEIF sanction_reason = "Return Locate Request or Provide Info" then
		memo_text = memo_text & "return the Locate Request Response that was previously mailed to you or contact me to provide information regarding the other parent's address, employment, or other "_
		& "information to assist in locating his/her whereabouts and/or identity.  The sanction will be removed after you return the Locate Request Response or contact me with information."
	ELSEIF sanction_reason = "Return PIF, Request Sheet, Assessment, and Fin Stmt" then
		memo_text = memo_text & "return the Paternity Information Form, Custodial Parent's Paternity Request Sheet, Special Services Assessment, and Financial Affidavit which were previously mailed "_
		& "to you.  The sanction will be removed when the missing documents are received."
	ELSEIF sanction_reason = "Attend Genetic Test Appointment" then
		memo_text = memo_text & "attend a genetic test appointment.  Contact me to schedule the appointment.  The sanction will be removed after you have cooperated with genetic testing."
	ELSEIF sanction_reason = "Attend CAO Appointment" then
		If CAO_contact = "Traci Melberg" then CAO_contact_and_phone = "Traci Melberg, legal assistant, at 763-323-5601"
		If CAO_contact = "Carrie Freeland" then CAO_contact_and_phone = "Carrie Freeland, legal assistant, at 763-323-5650"
		If CAO_contact = "Andrea Hesse" then CAO_contact_and_phone = "Andrea Hesse, legal assistant, at 763-323-5628"

		memo_text = memo_text & "attend an appointment at the Anoka County Attorney's Office to sign legal documents regarding the paternity of your child(ren).  Contact " & CAO_contact_and_phone _
		& " to schedule an appointment.  The sanction will be removed after you have signed the required legal documents."
	ELSEIF sanction_reason = "Provide Requested Info" then
		memo_text = memo_text & "provide requested information.  The sanction will be removed after you ________________."
	END IF
	memo_text = memo_text & "  Contact me if you have questions regarding this matter."

	CALL send_text_to_DORD (memo_text, "CPP")
End If

'Creating CAAD note
If caad_coop_check = checked OR caad_noncoop_check = checked OR caad_note <> "" then

	'Going to CAAD, adding a new note
	call navigate_to_PRISM_screen("CAAD")
	EMWriteScreen "A", 8, 5
	transmit
	EMReadScreen case_activity_detail, 20, 2, 29
	If case_activity_detail <> "Case Activity Detail" then script_end_procedure("The script could not navigate to a case note. There may have been a script error. Add case note manually, and report the error to a script writer.")


	'Setting the type
	EMWriteScreen "FREE", 4, 54

	'Setting cursor in write area and writing note details
	EMSetCursor 16, 4
	If FAS_or_CCA_Memo_noncoop_check = checked then call write_variable_in_CAAD("*** Noncooperation/Sanction request emailed to FAS and/or CCA workers")
	If FAS_or_CCA_Memo_coop_check = checked then call write_variable_in_CAAD("*** Noncooperation cured - Lift sanction request emailed to FAS and/or CCA workers")
	call write_variable_in_CAAD(CAAD_note)
	call write_variable_in_CAAD("---")
	call write_variable_in_CAAD(worker_signature)

	transmit
End if

'Creating worklist
If CAWD_noncoop_check = checked then

	call navigate_to_PRISM_screen("CAWD")
	PF5
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "*** SANCTIONED " & date & "!  Send reminder!", 10, 4
	EMWriteScreen dateadd("d", date, 28), 17, 21
	transmit
End if

IF CAWD_coop_check = checked then
	call navigate_to_PRISM_screen("CAWT")
	EMWriteScreen "FREE", 20, 29
	transmit

	CAWT_row = 8
	DO

		EMReadScreen end_of_data, 11, CAWT_row, 32
		IF end_of_data <> "End of Data" THEN

			EMReadScreen worklist_text, 14, CAWT_row, 15
			IF worklist_text = "*** SANCTIONED" THEN
				EMWriteScreen "P", CAWT_row, 4
				transmit
				transmit
			END IF
		CAWT_row = CAWT_row + 1
		END IF
	LOOP UNTIL end_of_data = end_of_data

END IF
call navigate_to_PRISM_screen("CAAD")
script_end_procedure("")
