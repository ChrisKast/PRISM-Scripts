Set oLst_vbs = Nothing													'Creates a blank object/variable
Set goFS = CreateObject("Scripting.FileSystemObject")					'Creates a scripting FSO
script_directory = "C:\DHS-PRISM-Scripts\actions\"						'Defining the script directory

'This part searches the entire folder for PRISM script files
For Each oFile In goFS.GetFolder(script_directory).Files				'For each file in the folder...
	If "vbs" = LCase(goFS.GetExtensionName(oFile.Name)) Then			'If it's a .vbs file we need to do stuff...
		If oLst_vbs Is Nothing Then 									'...but first if the variable is undefined...
			Set oLst_vbs = oFile ' the first could be the last			'...then add the first file to the variable...
		Else															'...otherwise...
			If oLst_vbs.DateLastModified < oFile.DateLastModified Then	'...check to see if the next file is older than the last file...
				Set oLst_vbs = oFile									'...if it is, add that file to the variable overwriting the previous!
			End If
		End If
	End If
Next																	'Loop until we've searched all files!

'This part yells at you if it found nothing, then stops.
If oLst_vbs Is Nothing Then
	MsgBox "no .vbs found"
	StopScript
End if

'LOADING GLOBAL VARIABLES--------------------------------------------------------------------
GlobVar_path = "C:\DHS-PRISM-Scripts\locally-installed-files\~globvar.vbs"													'Setting a default path, which is modified by the installer
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")														'Creating an FSO for the work
If run_another_script_fso.FileExists(GlobVar_path) then																		'If a Global Variables file is found in above directory...
	Set fso_command = run_another_script_fso.OpenTextFile(GlobVar_path)														'...run it!
Else																														'If a Global Variables file is not found in the above directory...
	Set fso_command = run_another_script_fso.OpenTextFile("Scripts\~globvar-local.vbs")										'...use the default BlueZone Scripts directory, and insert a custom "local flavor" Global Variables file, which can override the default selections.
End if
text_from_the_other_script = fso_command.ReadAll																			'Once we have the text from the other script, read it all!
fso_command.Close																											'Close the other script file, and...
Execute text_from_the_other_script	

'Runs the new script
script_to_run = script_directory & oLst_vbs.Name						'Determines the script to run
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")	'Makes an FSO for running another script
Set fso_command = run_another_script_fso.OpenTextFile(script_to_run)	'Opens the script to run
text_from_the_other_script = fso_command.ReadAll						'Reads all the contents from this script
fso_command.Close														'Closes the file
Execute text_from_the_other_script										'Runs the script
