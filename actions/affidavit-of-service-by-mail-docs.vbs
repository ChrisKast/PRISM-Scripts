'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "affidavit-of-service-by-mail.vbs"
start_time = timer
'MANUAL TIME TO COMPLETE THIS SCRIPT IS NEEDED

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

Dim ncp_button, cp_button, ncp_attorney_button, cp_attorney_button, summons_and_complaint, Amended_Summons_and_Complaint, Findings_Conclusion_Order, des_information, Amended_Findings_Conclusion_Order, Amended_Motion, motion, supporting_affidavit, financial_statement, Amended_Supporting_Affidavit, Notice_of_Hearing, Genetic_Blood_Test_Order, Notice_of_Intervention, Genetic_Blood_Test_results, Notice_of_Registration, Notice_of_Settlement_Conference, Aff_of_Default_and_ID, Your_Privacy_Rights, Case_Financial_Summary, Case_Information_Sheet, Case_Payment_History, Confidential_Info_Form, Important_Statement_of_Rights, sealed_financial_doc, Request_for_Hearing, guidelines_worksheet, Notice_of_Judgment_Renewal, confidential_yes, confidential_no, date_box, certified_mail_yes, certified_mail_no, other_line_1, other_line_2
BeginDialog AffOfServDialog, 0, 0, 301, 380, "Affidavit of Service By Mail"
  ButtonGroup ButtonPressed
    OkButton 65, 345, 65, 15
    CancelButton 155, 345, 65, 15
  Text 60, 10, 50, 10, "Case Number:"
  Text 100, 45, 85, 10, " (Check all that apply)"
  Text 10, 60, 170, 10, "What documents were served? (Check all that apply)"
  Text 75, 25, 150, 10, "Who do you want to send the Affidavit to?"
  EditBox 110, 5, 110, 15, prism_case_number
  CheckBox 10, 75, 110, 10, "Summons and Complaint", summons_and_complaint
  CheckBox 10, 90, 130, 10, "Amended Summons and Complaint", Amended_Summons_and_Complaint
  CheckBox 10, 105, 100, 10, "Findings/Conclusion/Order", Findings_Conclusion_Order
  CheckBox 10, 120, 130, 10, "Amended Findings/Conclusion/Order", Amended_Findings_Conclusion_Order
  CheckBox 10, 135, 115, 10, "Motion", motion
  CheckBox 10, 150, 125, 10, "Amended Motion", Amended_Motion
  CheckBox 10, 165, 95, 10, "Supporting Affidavit", supporting_affidavit
  CheckBox 10, 180, 110, 10, "Amended Supporting Affidavit", Amended_Supporting_Affidavit
  CheckBox 10, 195, 95, 10, "Financial Statement", financial_statement
  CheckBox 10, 210, 90, 10, "DES Information", des_information
  CheckBox 10, 225, 100, 10, "Genetic/Blood Test Order", Genetic_Blood_Test_Order
  CheckBox 10, 240, 110, 10, "Genetic/Blood Test Results", Genetic_Blood_Test_results
  CheckBox 10, 255, 95, 10, "Notice of Intervention", Notice_of_Intervention
  CheckBox 10, 270, 80, 10, "Notice of Hearing", Notice_of_Hearing
  CheckBox 180, 75, 90, 10, "Notice of Registration", Notice_of_Registration
  CheckBox 180, 90, 100, 10, "Notice of Settlement Conf.", Notice_of_Settlement_Conference
  CheckBox 180, 105, 95, 10, "Aff of Default and ID", Aff_of_Default_and_ID
  CheckBox 180, 120, 115, 10, "Case Financial Summary - CAFS", Case_Financial_Summary
  CheckBox 180, 150, 115, 10, "Case Payment History", Case_Payment_History
  CheckBox 180, 135, 95, 10, "Case Information Sheet", Case_Information_Sheet
  CheckBox 180, 165, 95, 10, "Confidential Info Form", Confidential_Info_Form
  CheckBox 180, 180, 105, 10, "Sealed Financial Document", sealed_financial_doc
  CheckBox 180, 195, 110, 10, "Important Statement of Rights", Important_Statement_of_Rights
  CheckBox 180, 210, 105, 10, "Your Privacy Rights", Your_Privacy_Rights
  CheckBox 180, 225, 95, 10, "Request for Hearing", Request_for_Hearing
  CheckBox 180, 240, 110, 10, "Notice of Judgment Renewal", Notice_of_Judgment_Renewal
  CheckBox 180, 255, 100, 10, "Guidelines Worksheet", guidelines_worksheet
  CheckBox 60, 35, 25, 10, "NCP", ncp_button
  CheckBox 90, 35, 20, 10, "CP", cp_button
  CheckBox 115, 35, 55, 10, "NCP Attorney", NCP_Attorney_button
  CheckBox 175, 35, 55, 10, "CP Attorney", Cp_attorney_button
  OptionGroup RadioGroup1
    RadioButton 225, 325, 25, 10, "Yes", certified_mail_yes
    RadioButton 255, 325, 25, 10, "No", certified_mail_no
  OptionGroup RadioGroup2
    RadioButton 95, 325, 25, 10, "No", confidential_no
    RadioButton 65, 325, 25, 10, "Yes", confidential_yes
  EditBox 10, 305, 80, 15, other_line_1
  EditBox 100, 305, 80, 15, other_line_2
  Text 10, 295, 65, 10, "Other (Line 1)"
  Text 100, 295, 65, 10, "Other (Line 2)"
  Text 15, 325, 45, 10, "Confidential?"
  Text 190, 295, 40, 10, "Date Served"
  EditBox 190, 305, 85, 15, date_box
  Text 130, 325, 85, 10, "Served By Certified Mail?"
EndDialog


'Connects to Bluezone
EMConnect ""

'Searches for the case number
row = 1
col = 1
EMSearch "Case: ", row, col
If row <> 0 then
	EMReadScreen PRISM_case_number, 13, row, col + 6
	PRISM_case_number = replace(PRISM_case_number, " ", "-")
	If isnumeric(left(PRISM_case_number, 10)) = False or isnumeric(right(PRISM_case_number, 2)) = False then PRISM_case_number = ""
End if


'goes to correct case
EMWriteScreen "CAST", 21,18
Transmit
EMWriteScreen "D", 3, 29
'Puts case number in from Dialog box
	EMWriteScreen Left (prism_case_number, 10), 4, 8
	EMWriteScreen Right (prism_case_number, 2), 4, 19
Transmit

Do
	err_msg = ""
	Dialog AffOfServDialog 'Shows name of dialog
		IF buttonpressed = 0 then stopscript		'Cancel
		IF ncp_button = 0 AND cp_button = 0 AND NCP_Attorney_button = 0 AND CP_Attorney_button = 0 THEN err_msg = err_msg & vbNewline & "Please select the receipiant for your Affidavit."
		IF date_box = "" THEN err_msg = err_msg & vbNewline & "The date served must be completed." 
		IF summons_and_complaint = 0 AND Amended_Summons_and_Complaint = 0 AND Findings_Conclusion_Order = 0 AND Amended_Findings_Conclusion_Order = 0 AND motion = 0 AND Amended_Motion = 0 AND supporting_affidavit = 0 AND Amended_Supporting_Affidavit = 0 AND financial_statement = 0 AND des_information = 0 AND Genetic_Blood_Test_Order = 0 AND Genetic_Blood_Test_results = 0 AND Notice_of_Intervention = 0 AND Notice_of_Hearing = 0 AND Notice_of_Registration = 0 AND Notice_of_Settlement_Conference = 0 AND Aff_of_Default_and_ID = 0 AND Case_Financial_Summary = 0 AND Case_Payment_History = 0 AND Case_Information_Sheet = 0 AND Confidential_Info_Form = 0 AND sealed_financial_doc = 0 AND Important_Statement_of_Rights = 0 AND Your_Privacy_Rights = 0 AND Request_for_Hearing = 0 AND Notice_of_Judgment_Renewal = 0 AND guidelines_worksheet = 0 AND other_line_1 = "" AND other_line_2 = "" THEN err_msg = err_msg & vbNewline & "At least one document must be selected."
		IF err_msg <> "" THEN 
			MsgBox "***NOTICE!!!***" & vbNewline & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue."
		END IF

LOOP UNTIL err_msg = ""

'---------------------------------------------------------------------------------------------------------Creates DORD doc if CP checked
IF cp_button = checked then
'goes to DORD
EMWriteScreen "DORD", 21,18
Transmit
EMWriteScreen "C", 3, 29
Transmit
'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0016", 6, 36
EMWriteScreen "CPP", 11, 51
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit

If summons_and_complaint = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Summons_and_Complaint = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Findings_Conclusion_Order = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Findings_Conclusion_Order = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Motion = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If motion = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If supporting_affidavit = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Supporting_Affidavit = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If financial_statement = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If des_information = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF


EMSendKey (PF8)
If Genetic_Blood_Test_Order = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Genetic_Blood_Test_results = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Intervention = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Hearing = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Registration = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Settlement_Conference = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Aff_of_Default_and_ID = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Financial_Summary = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Payment_History = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Information_Sheet = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Confidential_Info_Form = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If sealed_financial_doc = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
Transmit

If Important_Statement_of_Rights = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Your_Privacy_Rights = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Request_for_Hearing = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Judgment_Renewal = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If certified_mail_yes = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen other_line_1, 16, 15
Transmit
EMWriteScreen other_line_2, 16, 15
Transmit
EMWriteScreen date_box, 16, 15
Transmit


EMSendKey (PF8)
If guidelines_worksheet = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If confidential_yes = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "Y", 16, 15
Transmit
End IF
If confidential_no = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "N", 16, 15
Transmit
End IF

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit

'''need to select legal heading
BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  ButtonGroup ButtonPressed
    OkButton 60, 75, 50, 15
    CancelButton 115, 75, 50, 15
  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  Text 5, 40, 55, 10, "2. Press ENTER"
  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
EndDialog

			Dialog LH_dialog  'name of dialog
			IF buttonpressed = 0 then stopscript		'Cancel

'EMSendKey (PF3)
End If
'--------------------------------------------------------------------------------------------------------------------------------------Creates DORD doc if NCP checked

If ncp_button = checked then
EMWriteScreen "DORD", 21,18
Transmit
EMWriteScreen "C", 3, 29
Transmit
'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0016", 6, 36
EMWriteScreen "NCP", 11, 51
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit

If summons_and_complaint = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Summons_and_Complaint = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Findings_Conclusion_Order = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Findings_Conclusion_Order = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Motion = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If motion = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If supporting_affidavit = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Supporting_Affidavit = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If financial_statement = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If des_information = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF


EMSendKey (PF8)
If Genetic_Blood_Test_Order = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Genetic_Blood_Test_results = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Intervention = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Hearing = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Registration = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Settlement_Conference = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Aff_of_Default_and_ID = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Financial_Summary = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Payment_History = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Information_Sheet = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Confidential_Info_Form = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If sealed_financial_doc = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
Transmit

If Important_Statement_of_Rights = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Your_Privacy_Rights = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Request_for_Hearing = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Judgment_Renewal = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If certified_mail_yes = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen other_line_1, 16, 15
Transmit
EMWriteScreen other_line_2, 16, 15
Transmit
EMWriteScreen date_box, 16, 15
Transmit


EMSendKey (PF8)
If guidelines_worksheet = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If confidential_yes = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "Y", 16, 15
Transmit
End IF
If confidential_no = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "N", 16, 15
Transmit
End IF

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit

'''need to select legal heading
BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  ButtonGroup ButtonPressed
    OkButton 60, 75, 50, 15
    CancelButton 115, 75, 50, 15
  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  Text 5, 40, 55, 10, "2. Press ENTER"
  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
EndDialog

			Dialog LH_dialog  'name of dialog
			IF buttonpressed = 0 then stopscript		'Cancel

'EMSendKey (PF3)

End If


'--------------------------------------------------------------------------------------------------------------------------------------Creates DORD doc if CP attorney checked

If Cp_attorney_button = checked then
EMWriteScreen "DORD", 21,18
Transmit
EMWriteScreen "C", 3, 29
Transmit
'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0016", 6, 36
EMWriteScreen "CPA", 11, 51
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit

If summons_and_complaint = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Summons_and_Complaint = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Findings_Conclusion_Order = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Findings_Conclusion_Order = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Motion = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If motion = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If supporting_affidavit = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Supporting_Affidavit = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If financial_statement = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If des_information = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF


EMSendKey (PF8)
If Genetic_Blood_Test_Order = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Genetic_Blood_Test_results = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Intervention = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Hearing = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Registration = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Settlement_Conference = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Aff_of_Default_and_ID = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Financial_Summary = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Payment_History = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Information_Sheet = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Confidential_Info_Form = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If sealed_financial_doc = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
Transmit

If Important_Statement_of_Rights = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Your_Privacy_Rights = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Request_for_Hearing = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Judgment_Renewal = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If certified_mail_yes = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen other_line_1, 16, 15
Transmit
EMWriteScreen other_line_2, 16, 15
Transmit
EMWriteScreen date_box, 16, 15
Transmit


EMSendKey (PF8)
If guidelines_worksheet = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If confidential_yes = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "Y", 16, 15
Transmit
End IF
If confidential_no = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "N", 16, 15
Transmit
End IF

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit

'''need to select legal heading
BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  ButtonGroup ButtonPressed
    OkButton 60, 75, 50, 15
    CancelButton 115, 75, 50, 15
  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  Text 5, 40, 55, 10, "2. Press ENTER"
  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
EndDialog

			Dialog LH_dialog  'name of dialog
			IF buttonpressed = 0 then stopscript		'Cancel

'EMSendKey (PF3)


End If

'--------------------------------------------------------------------------------------------------------------------------------------Creates DORD doc if NCP attorney checked

If Cp_attorney_button = checked then
EMWriteScreen "DORD", 21,18
Transmit
EMWriteScreen "C", 3, 29
Transmit
'Adds dord doc
EMWriteScreen "A", 3, 29
'blanks out any DOC ID number that may be entered
EMWriteScreen "        ", 4, 50
EMWriteScreen "       ", 4, 59
EMWriteScreen "F0016", 6, 36
EMWriteScreen "NCA", 11, 51
Transmit

'entering user labels
EMSendKey (PF14)
EMWriteScreen "U", 20, 14
Transmit

If summons_and_complaint = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Summons_and_Complaint = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Findings_Conclusion_Order = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Findings_Conclusion_Order = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Motion = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If motion = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If supporting_affidavit = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Amended_Supporting_Affidavit = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If financial_statement = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If des_information = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF


EMSendKey (PF8)
If Genetic_Blood_Test_Order = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Genetic_Blood_Test_results = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Intervention = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Hearing = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Registration = checked then
EMWriteScreen "S", 11, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Settlement_Conference = checked then
EMWriteScreen "S", 12, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Aff_of_Default_and_ID = checked then
EMWriteScreen "S", 14, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Financial_Summary = checked then
EMWriteScreen "S", 13, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Payment_History = checked then
EMWriteScreen "S", 15, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Case_Information_Sheet = checked then
EMWriteScreen "S", 16, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Confidential_Info_Form = checked then
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If sealed_financial_doc = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
Transmit

If Important_Statement_of_Rights = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Your_Privacy_Rights = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Request_for_Hearing = checked then
EMWriteScreen "S", 9, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If Notice_of_Judgment_Renewal = checked then
EMWriteScreen "S", 10, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If certified_mail_yes = checked then
EMWriteScreen "S", 18, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
EMWriteScreen "S", 15, 5
EMWriteScreen "S", 16, 5
EMWriteScreen "S", 17, 5
Transmit
EMWriteScreen other_line_1, 16, 15
Transmit
EMWriteScreen other_line_2, 16, 15
Transmit
EMWriteScreen date_box, 16, 15
Transmit


EMSendKey (PF8)
If guidelines_worksheet = checked then
EMWriteScreen "S", 7, 5
Transmit
EMWriteScreen "x", 16, 15
Transmit
End IF
If confidential_yes = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "Y", 16, 15
Transmit
End IF
If confidential_no = checked then
EMWriteScreen "S", 8, 5
Transmit
EMWriteScreen "N", 16, 15
Transmit
End IF

EMSendKey (PF3)
EMWriteScreen "M", 3, 29
Transmit

'''need to select legal heading
BeginDialog LH_dialog, 0, 0, 171, 95, "Select Legal Heading"
  ButtonGroup ButtonPressed
    OkButton 60, 75, 50, 15
    CancelButton 115, 75, 50, 15
  Text 35, 10, 100, 10, "IMPORTANT! IMPORTANT!"
  Text 5, 25, 130, 10, "1. Select the correct LEGAL HEADING"
  Text 5, 40, 55, 10, "2. Press ENTER"
  Text 5, 55, 140, 10, "3.  THEN click OK for the script to continue"
EndDialog

			Dialog LH_dialog  'name of dialog
			IF buttonpressed = 0 then stopscript		'Cancel

'EMSendKey (PF3)

End If

script_end_procedure("")
