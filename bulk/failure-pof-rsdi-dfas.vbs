'GATHERING STATS---------------------------------------------------------------------------------------------------- 
name_of_script = "failure-pof-rsdi-dfas.vbs" 
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

'This is the dialog to select the CSO. The script will run off the 8 digit worker ID code entered here.
FUNCTION select_cso(ButtonPressed, cso_id, cso_name)
	DO
		DO
			CALL navigate_to_PRISM_screen("USWT")
			err_msg = ""
			'Grabbing the CSO name for the intro dialog.
			CALL find_variable("Worker Id: ", cso_id, 8)
			EMSetCursor 20, 13
			PF1
			CALL write_value_and_transmit(cso_id, 20, 35)
			EMReadScreen cso_name, 24, 13, 55
			cso_name = trim(cso_name)
			PF3
			
			BeginDialog select_cso_dlg, 0, 0, 286, 145, "E0014- Failure IW Notice to Payor of Funds - Select CSO"
			EditBox 70, 55, 65, 15, cso_id
			Text 70, 80, 155, 10, cso_name
			ButtonGroup ButtonPressed
				OkButton 130, 125, 50, 15
				PushButton 180, 125, 50, 15, "UPDATE CSO", update_cso_button
				PushButton 230, 125, 50, 15, "STOP SCRIPT", stop_script_button
			Text 10, 15, 265, 30, "This script will check for worklist items coded E0014 for the following Worker ID. If you wish to change the Worker ID, enter the desired Worker ID in the box and press UPDATE CSO. When you are ready to continue, press OK."
			Text 10, 60, 50, 10, "Worker ID:"
			Text 10, 80, 55, 10, "Worker Name:"
		
			EndDialog
		
			DIALOG select_cso_dlg
				IF ButtonPressed = stop_script_button THEN script_end_procedure("The script has stopped.")
				IF ButtonPressed = update_cso_button THEN 
					CALL navigate_to_PRISM_screen("USWT")
					CALL write_value_and_transmit(cso_id, 20, 13)
					EMReadScreen cso_name, 24, 13, 55
					cso_name = trim(cso_name)
				END IF
				IF cso_id = "" THEN err_msg = err_msg & vbCr & "* You must enter a Worker ID."
				IF len(cso_id) <> 8 THEN err_msg = err_msg & vbCr & "* You must enter a valid, 8-digit Worker ID."
																																				'The additional of IF ButtonPressed = -1 to the conditional statement is needed 
																																		'to allow the worker to update the CSO's worker ID without getting a warning message.
				IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL ButtonPressed = -1 
	LOOP UNTIL err_msg = ""
END FUNCTION

'=====VARIABLES TO DECLARE=====
checked = 1
unchecked = 0

'=====ADDITIONAL CUSTOM FUNCTIONS=====

'This first custom function is used to build the case review dialog. The values pulled through are "i" and "nocs_array."
'Changes in "i" change the case number and case information being worked on.
'nocs_array holds the information for that case.
FUNCTION build_NOCS_dlg(i, nocs_array)


BeginDialog NOCS_dlg, 0, 0, 257, 226, "E0014 Failure Notice to POF worklist"
  Text 70, 10, 110, 10, nocs_array(i, 0)
  Text 10, 70, 50, 10, "CP Name:"
  Text 60, 70, 180, 10, nocs_array(i, 1)
  Text 10, 50, 50, 10, "NCP Name:"
  Text 60, 50, 180, 10, nocs_array(i, 2)
  Text 10, 110, 70, 10, "Employer on NCID:"
  Text 100, 110, 160, 14, nocs_array(i, 4)
  Text 10, 90, 110, 10, "Employer on PAPL:"
  Text 100, 90, 150, 14, nocs_array(i, 5)
  CheckBox 20, 130, 270, 10, "Check HERE to PURGE the E0014 Failure Notice to POF worklist.", nocs_array(i, 3)
  Text 10, 10, 50, 10, "Case Number:"
  ButtonGroup ButtonPressed
    OkButton 20, 170, 50, 20
    PushButton 90, 170, 50, 20, "SKIP CASE", skip_case_button
    PushButton 150, 170, 60, 20, "STOP SCRIPT", stopscript_button
EndDialog


	'We load the dialog within the function for each case.
	'The do/loop allows for error checking.
	DO
		'loading the dialog
		DIALOG NOCS_dlg
			IF ButtonPressed = stopscript_button THEN script_end_procedure("The script has stopped.")		'If the user presses the "STOP SCRIPT" button, the script stops
						
			IF ButtonPressed = skip_case_button AND initial_run_through = True THEN 		'IF the user elects to skip this case and this is the first time through...
				nocs_array(i, 3) = unchecked			'Do not purge
			END IF											'...the value initial_run_through is important. We do not want the user to work on several cases, run decide to review the cases again, and have to duplicate their work.
															'... If the user is double checking their work (meaning this is NOT the first time they are running through the dialogs), pressing "SKIP CASE" will not reset the values in the dialog.
	LOOP UNTIL ButtonPressed = skip_case_button OR (ButtonPressed = -1)		'In order to get out of the do/loop, the user must press either the "SKIP CASE" button, OR press "OK" and have a null err_msg

	
	
END FUNCTION 

BeginDialog confirm_dlg, 0, 0, 281, 135, "Confirmation"
  ButtonGroup ButtonPressed
    PushButton 220, 65, 50, 15, "RETRY", retry_button
    OkButton 220, 110, 50, 15
  Text 10, 40, 260, 20, "If you feel any of the cases need revision, press RETRY. The script will then run through all the cases one more time."
  Text 10, 10, 260, 20, "You have reviewed all the cases. You can double check script's work using the existing Excel file."
  Text 10, 90, 265, 10, "Otherwise, if you think the cases are ready for processing, press OK to continue."
EndDialog


'=====THE SCRIPT=====
EMConnect ""
CALL check_for_PRISM(False)

'Loading the dialog to select the CSO
CALL select_cso(ButtonPressed, cso_id, cso_name)

'And away we go...
CALL write_value_and_transmit("E0014", 20, 30)

uswt_row = 7
DO
	EMReadScreen uswt_type_id, 5, uswt_row, 45
	EMReadScreen prism_case_number, 13, uswt_row, 8
	prism_case_number = replace(prism_case_number, " ", "-")
	IF uswt_type_id = "E0014" THEN cases_array = cases_array & prism_case_number & " "
	uswt_row = uswt_row + 1
	IF uswt_row = 19 THEN 
		PF8
		uswt_row = 7
	END IF
LOOP UNTIL uswt_type_id <> "E0014"

cases_array = trim(cases_array)
cases_array = split(cases_array, " ")

number_of_cases = ubound(cases_array)
DIM nocs_array()
ReDim nocs_array(number_of_cases, 6)

'>>>> HERE ARE THE 6 POSITIONS WITHIN THE ARRAY <<<<
'nocs_array(i, 0) >> PRISM_case_number
'nocs_array(i, 1) >> CP name
'nocs_array(i, 2) >> NCP name
'nocs_array(i, 3) >> Purge? (1 for Yes, 0 for No)
'nocs_array(i, 4) >> Employer on NCID
'nocs_array(i, 5) >> Employer on PAPL

position_number = 0
FOR EACH prism_case_number IN cases_array
'	nocs_array(i, 0) >> PRISM_case_number
	IF prism_case_number <> "" THEN 
		nocs_array(position_number, 0) = prism_case_number
		position_number = position_number + 1
	END IF
NEXT



FOR i = 0 to number_of_cases
'	nocs_array(i, 0) >> PRISM_case_number
'	nocs_array(i, 1) >> CP name
'	nocs_array(i, 2) >> NCP name
'	nocs_array(i, 4) >> Employer on NCID
'	nocs_array(i, 5) >> Employer on PAPL
	CALL navigate_to_PRISM_screen("CAST")
	EMWriteScreen nocs_array(i, 0), 4, 8
	EMWriteScreen right(nocs_array(i, 0), 2), 4, 19
	CALL write_value_and_transmit("D", 3, 29)
	EMReadScreen full_service, 1, 9, 60
	EMReadScreen cp_name, 35, 6, 12
	EMReadScreen ncp_name, 35, 7, 12
	cp_name = trim(cp_name)
	ncp_name = trim(ncp_name)
	nocs_array(i, 1) = cp_name
	nocs_array(i, 2) = ncp_name



		CALL navigate_to_PRISM_screen("NCDD")
		CALL navigate_to_PRISM_screen("NCSU")
			EMReadScreen NCID_emp, 30, 13, 49
			nocs_array(i, 4) = NCID_emp
		CALL navigate_to_PRISM_screen("PAPL")

		' >>>>> MAKING SURE THAT THERE IS INFORMATION ON PAPL <<<<
		EMReadScreen end_of_data, 11, USWT_row, 32
		IF end_of_data <> "End of Data" THEN

			' >>>>> READING THE MOST RECENT PAY DATE AND CONVERTING IT TO A USABLE DATE <<<<<
			EMReadScreen PAPL_most_recent_pay_date, 6, 7, 7
			Call date_converter_PALC_PAPL(PAPL_most_recent_pay_date)
			pmt_year = Right(PAPL_most_recent_pay_date, 2) 'string variables added to track the payment month and 2-digit year.
			pmt_month = Left(PAPL_most_recent_pay_date, 2)	
			
						
			' >>>> CHECKING THAT THE DATE IN THE PAYMENT ID IS FROM THE CURRENT MONTH MINUS 1 <<<<<
			current_month_minus1 = DateAdd("m", -1, date) 'variable for the current date minus one - this returns a date format
			c_month = datepart("m", current_month_minus1)
			IF len(c_month) = 1 THEN c_month = "0" & c_month
			
			
			c_year = Right(CStr(current_month_minus1), 2) 'string variables added to track the current month minus 1 month and year. 
			'c_month = Left(CStr(current_month_minus1), 2)
			
			IF pmt_year >= c_year THEN
				If  pmt_month >= c_month THEN  
 				' >>>>> IF THE PAYMENT IS FROM LAST MONTH OR CURRENT MONTH, THE SCRIPT GRABS THE EMPLOYER/SOURCE ID <<<<<
				'We want this to occur if the payment occurred last month or in the current month.				
					PF11
					EMReadScreen PAPL_name, 30, 7, 38
					nocs_array(i, 5) = PAPL_name
					' >>>>> LISTING OUT THE CONDITIONS THAT CAN BE PURGED AUTOMATICALLY <<<<<
					IF InStr(PAPL_name, "DFAS") <> 0 OR _
					   InStr(PAPL_name, "U S SOCIAL") <> 0 OR _ 
					   InStr(PAPL_name, "U S DEPT OF TREASURY") <> 0 THEN 
					   nocs_array(i, 3) = checked
					Else
					   nocs_array(i, 3) = unchecked
						
					End If
				End If
			END IF
		End If

NEXT

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
objExcel.Cells(1, 4).Value = "PURGE?"
objExcel.Cells(1, 4).Font.Bold = True
objExcel.Cells(1, 5).Value = "NCID Employer"
objExcel.Cells(1, 5).Font.Bold = True
objExcel.Cells(1, 6).Value = "PAPL Employer"
objExcel.Cells(1, 6).Font.Bold = True

excel_row = 2

'Updating the Excel spreadsheet with initial information
FOR i = 0 to number_of_cases 
	FOR k = 0 to 6
		objExcel.Cells(excel_row, k + 1).Value = nocs_array(i, k)
		IF k = 3 THEN 
			IF nocs_array(i, k) = checked THEN 
				objExcel.Cells(excel_row, k + 1).Value = "Y"		
			END IF
			IF nocs_array(i, k) = unchecked THEN 
				objExcel.Cells(excel_row, k + 1).Value = "N"
			END IF
		END IF
	NEXT
	excel_row = excel_row + 1
NEXT

'Autofitting each column.
FOR x_col = 1 to 6
	objExcel.Columns(x_col).AutoFit()
NEXT

initial_run_through = True
DO
	'Running the dialog for each case.
	excel_row = 2
	FOR i = 0 to number_of_cases
		'This changes the back fill of the selected row to draw the worker's attention.
		FOR select_column = 1 to 6
			objExcel.Cells(excel_row, select_column).Interior.ColorIndex = 6		'Setting the background fill to yellow
			objExcel.Cells(excel_row - 1, select_column).Interior.ColorIndex = 2	'Setting the previous background fill to white
		NEXT
		'Building the dialog
		CALL build_NOCS_dlg(i, nocs_array)
		
		'Updating the Excel spreadsheet in real time.
		FOR j = 0 to 5
			objExcel.Cells(excel_row, j + 1).Value = nocs_array(i, j)
			IF j = 3 THEN 
				IF nocs_array(i, 3) = checked THEN objExcel.Cells(excel_row, j + 1).Value = "Y"
				IF nocs_array(i, 3) = unchecked THEN objExcel.Cells(excel_row, j + 1).Value = "N"
			END IF
		NEXT
		excel_row = excel_row + 1
	NEXT
	
	FOR select_column = 1 to 6
		objExcel.Cells(excel_row - 1, select_column).Interior.ColorIndex = 2
	NEXT
	
	DIALOG confirm_dlg
		IF ButtonPressed <> -1 THEN initial_run_through = False
LOOP UNTIL ButtonPressed = -1

excel_row = 2
'Refreshing spreadsheet values
FOR ii = 0 to number_of_cases 
	FOR j = 0 to 5
		objExcel.Cells(excel_row, j + 1).Value = nocs_array(ii, j)
			IF j = 3 THEN 
				IF nocs_array(ii, 3) = checked THEN objExcel.Cells(excel_row, j + 1).Value = "Y"
				IF nocs_array(ii, 3) = unchecked THEN objExcel.Cells(excel_row, j + 1).Value = "N"
			END IF

	NEXT
	excel_row = excel_row + 1
NEXT
'Redoing the autofit for the columns.
FOR x_col = 1 to 6
	objExcel.Columns(x_col).AutoFit()
NEXT

'One more check for PRISM
CALL check_for_PRISM(False)


'Now the script needs to PURGE for all (i, 3) = True

number_of_cases_purged = 0
FOR i = 0 to number_of_cases
'	nocs_array(i, 0) >> PRISM_case_number
'	nocs_array(i, 3) >> Purge? (1 for Yes, 0 for No)
	IF nocs_array(i, 3) = checked THEN 
		CALL navigate_to_PRISM_screen("CAWT")
		CALL write_value_and_transmit("E0014", 20, 29)
		EMWriteScreen left(nocs_array(i, 0), 10), 20, 8	
		EMWritescreen right(nocs_array(i, 0), 2), 20, 19
		transmit
		
	
		DO
			EMReadscreen cawd_type, 5, 8, 8
			IF cawd_type = "E0014" THEN
				EMWriteScreen "P", 8, 4
				transmit
				transmit
				number_of_cases_purged = number_of_cases_purged + 1
			END IF
		LOOP UNTIL cawd_type <> "E0014"
	END IF
NEXT

script_end_procedure("Success!! " &  number_of_cases_purged  & " items have been purged.")
