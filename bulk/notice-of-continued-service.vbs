'STATS GATHERING ---------------------------
name_of_script = "notice-of-continued-service.vbs"
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
call changelog_update("11/14/2016", "Sped up process of this script. Now matches Ramsey County's script.", "Heather Allen, Scott County")
call changelog_update("11/14/2016", "Initial version.", "Heather Allen, Scott County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'CONNECTING TO BLUEZONE
EMConnect ""

Dim TimeOutStr
EMWriteScreen "CAST", 21, 18
Transmit

EMReadScreen TimeOutStr, 1, 12, 53
If (TimeOutStr = ">") then
	MsgBox "Please log in first!", vbExclamation
	StopScript
End If

'DIALOG BOX==================================================
BeginDialog NOCS_Enter_Date, 0, 0, 216, 75, "NOCS Enter Date"
  Text 15, 10, 205, 10, "Enter the final day of the previous month when PA ended:"
  EditBox 15, 25, 190, 15, EnterDate
  ButtonGroup ButtonPressed
    OkButton 45, 50, 50, 15
    CancelButton 130, 50, 50, 15
EndDialog


'FILTERING FOR D0800 ON USWT
EMWriteScreen "USWT", 21, 18
Transmit
EMWriteScreen "D0800", 20, 30
Transmit


'INPUT DIALOG BOX FOR WORKER TO ADD DATE THAT IS APPLIED TO EACH NOCS DOCUMENT

Do
	dialog NOCS_Enter_Date
	IF buttonpressed = 0 THEN StopScript
	IF EnterDate <> "" THEN Exit Do
	IF EnterDate = "" THEN 
		MsgBox "Script Ended"
	End If
	
Loop

'CREATING EXCEL SPREADSHEET

	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	Set objWorkbook = objExcel.Workbooks.Add()
	objExcel.DisplayAlerts = True

		objExcel.Cells(1, 1).Value = "CASE NUMBER"
		objExcel.Cells(1, 1).Font.Bold = True
		objExcel.Cells(1, 2).Value = "CUSTODIAL PARENT"
		objExcel.Cells(1, 2).Font.Bold = True
		objExcel.Cells(1, 3).Value = "NON-CUSTODIAL PARENT"
		objExcel.Cells(1, 3).Font.Bold = True
		objExcel.Cells(1, 4).Value = "GOOD CAUSE COOP?"
		objExcel.Cells(1, 4).Font.Bold = True
		objExcel.Cells(1, 5).Value = "DORD DOC SENT"
		objExcel.Cells(1, 5).Font.Bold = True
		

'BEGINNING WORKLIST LOOP FOR PROCESSING D0800
Dim Worklist_D0800, PrgCode, Confirm_FullServ, DORDcode, SanctionStr

'DEFINING VARIABLE FOR DO...LOOP
excel_row = 2

Do
	'CONFIRMS FIRST WORKLIST ON CAWT LIST IS D0800
	EMReadScreen Worklist_D0800, 5, 7, 45
	IF Worklist_D0800 <> "D0800" Then
		MsgBox ("All D0800 Worklists have been processed on this caseload!")
		Exit Do
	End If

	'DISPLAYING D0800 WORKLIST TO CONFIRM IF THE CASE IS NPA OR NOT
	EMWriteScreen "D", 7, 4
	Transmit
	EMReadScreen PrgCode, 3, 6, 68
	EMReadScreen PRISM_case_number, 13, 4, 08
	objExcel.Cells(excel_row, 1).Value = PRISM_case_number
	EMReadScreen CP_name, 20, 6, 12
	objExcel.Cells(excel_row, 2).Value = CP_name
	EMReadScreen NCP_name, 20, 7, 12
	objExcel.Cells(excel_row, 3).Value = NCP_name
	EMWriteScreen "P", 3, 30 
	Transmit
	Transmit
	PF3

	'AUTOFITTING COLUMNS OF EXCEL SPREADSHEET
	For excel_column = 1 to 5
		objExcel.Columns(excel_column).Autofit()
	Next	

	If PrgCode = "NPA" Then

		'SENDING APPROPRIATE NOCS DOC DEPENDING IF CASE IS FULL SERVICE OR NOT ON CAST SCREEN
		EMWriteScreen "CAST", 21, 18
		Transmit
		EMReadScreen Confirm_FullServ, 1, 9, 60
		If Confirm_FullServ = "Y" then DORDcode = "F0111"
		If Confirm_FullServ = "N" then DORDcode = "F0115"

		'PURGING WORKLIST M0061 CASE PROGRAMED CHANGED
		EMWriteScreen "CAWT", 21, 18
		Transmit
		EMWriteScreen "M0061", 20, 29
		Transmit
		EMWriteScreen "D", 8, 4
		Transmit
		EMWriteScreen "P", 3, 30
		Transmit
		Transmit
		PF3

		'GENERATING NOCS ON DORD
		EMWriteScreen "DORD", 21, 18
		Transmit
		EMWriteScreen "C", 3, 29
		Transmit
		EMWriteScreen "A", 3, 29
		EMSetCursor 6, 36
		EMSendKey DORDcode
		EMReadScreen DORDcode, 5, 6, 36
		objExcel.Cells(excel_row, 5).Value = DORDcode
		Transmit
		EMWriteScreen "M", 3, 29
		PF14
		EMWriteScreen "U", 20, 14
		Transmit
		EMWriteScreen "S", 7, 5
		Transmit
		EMWriteScreen EnterDate, 16, 15
		Transmit
		PF3

		'CONFIRMING IF CASE IS IN SANCTION OR NOT ON GCSC
		EMWriteScreen "GCSC", 21, 18
		Transmit
		EMReadScreen Goodcause_coop, 1, 15, 25
		EMReadScreen Goodcause_coop, 1, 15, 25
		objExcel.Cells(excel_row, 4).Value = Goodcause_coop
		If Goodcause_coop = "N" Then
			EMWriteScreen "CAWT", 21, 18
			Transmit
			PF5
			EMWriteScreen "FREE", 4, 37
			EMWriteScreen "7", 17, 52
			EMWriteScreen "**REVIEW FOR POSSIBLE CLOSURE**", 10, 4
			EMSetCursor 11, 4
			EMSendKey "CASE IS NON-COOP AND NOCS WAS SENT ON" & _
					Date & "BY D0800 WORKLIST MACRO."
			Transmit

		End If
	End If

	'RETURNING TO USWT FOR ADDITIONAL D0800 WORKLISTS
	EMWriteScreen "USWT", 21, 18
	Transmit
	EMWriteScreen "D0800", 20, 30
	Transmit
	EMReadScreen Worklist_D0800, 5, 7, 45

	'INCREMENTING EXCEL BY ONE SO IT KEEPS ADDING CASE INFO TO THE SPREADSHEET
	excel_row = excel_row + 1

Loop


