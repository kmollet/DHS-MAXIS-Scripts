name_of_script = "NOTES - Employment verif received.vbs"
start_timer = timer
'basic script for when someone either receives an EVF, or some sort of employment verification and they need to input the information.  This has mandatory fields and will also be ready to update to serve as a function and note script to update the jobs panel in itself

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

EMConnect ""
call check_for_maxis(false)
Emfocus

BeginDialog DialogEVF_received, 0, 0, 251, 225, "Employment Verification Received "
  EditBox 65, 15, 45, 15, case_number
  EditBox 200, 15, 45, 15, docs_received
  EditBox 65, 35, 45, 15, start_date
  EditBox 200, 35, 45, 15, first_check
  EditBox 125, 60, 90, 15, employer_name
  EditBox 125, 80, 90, 15, rate_of_pay
  EditBox 125, 100, 90, 15, anticipated_hours
  DropListBox 125, 120, 90, 15, "Select one"+chr(9)+"monthly"+chr(9)+"semi-monthly"+chr(9)+"bi-weekly"+chr(9)+"weekly"+chr(9)+"other", pay_frequency
  EditBox 125, 140, 90, 15, frequency_explain
  EditBox 70, 160, 170, 15, income_notes
  EditBox 110, 180, 55, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 60, 200, 50, 15
    CancelButton 125, 200, 50, 15
  Text 15, 15, 50, 10, "Case Number"
  Text 130, 15, 65, 10, "Date docs received"
  Text 5, 35, 60, 10, "Income Start date"
  Text 110, 35, 90, 10, "Anticipated first check date"
  Text 20, 60, 55, 10, "Employer Name"
  Text 20, 80, 80, 10, "Anticipated rate of pay" 
  Text 20, 100, 85, 10, "Anticipated hours per week"
  Text 20, 120, 50, 10, "Pay frequency"
  Text 20, 140, 105, 10, "If other pay frequency, explain"
  Text 5, 165, 60, 10, "Notes on income"
  Text 45, 185, 60, 10, "Worker Signature"
EndDialog



Do
	Error_message = ""
'error message so that the dialog box is fool proof and this box can then be upgraded to edit the jobs panel for the worker.  Made mandatory all fields which are necessary to fill out the jobs panel.  Basically the minimum information nneeded for jobs panel
'This also makes sure that when someone needs to enter a date that it is an actual date, or it is an actual number, etc
	Dialog DialogEVF_received
	If buttonpressed = 0 then cancel_confirmation
	if case_number = "" or isnumeric(case_number) = false then error_message = error_message & vbcr & "You must enter a valid case number"
	if docs_received = "" or isdate(docs_received) = false then error_message = error_message & vbcr & "You must enter a valid received date"
	if first_check = "" or isdate(first_check) = false then error_message = error_message & vbcr & "You must enter a valid date of first check"
	if anticipated_hours = "" or isnumeric(anticipated_hours) = false then error_message = error_message & vbcr & "You must enter valid anticipated hours"
	if employer_name = "" then error_message = error_message & vbcr & "You must enter an employer name"
	if rate_of_pay = "" or isnumeric(rate_of_pay) = false then error_message = error_message & vbcr & "You must enter a valid rate of pay"
	if pay_frequency = "Select one" then error_message = error_message & vbcr & "You must enter a frequency"
	if worker_signature = "" then error_message = error_message & vbcr & "You must sign your case note"
	if error_message <> "" then msgbox error_message
Cancel_confirmation
'allows the user to press "cancel" and exit the script
Loop until error_message = ""
'when there are no more error messages then the dialog box can finish and go to case note
call hh_member_custom_dialog(hh_member_array)
'to let the user elect the working member of household



call check_for_maxis(false)
'to make sure someone did not log out of maxis or that they are stuck in background

call start_a_blank_case_note
call write_variable_in_case_note ("***Employment verification received for Memb " & hh_member_array(0) & "***")
call write_bullet_and_variable_in_case_note("Date docs received", docs_received)
call write_bullet_and_variable_in_case_note("Income start date", start_date)
call write_bullet_and_variable_in_case_note("Anticipated rate of pay ", "$" & rate_of_pay) 
call write_bullet_and_variable_in_case_note("Anticipated hours per week ", anticipated_hours)
call write_bullet_and_variable_in_case_note("Employer name", employer_name)
call write_bullet_and_variable_in_case_note("Anticapted pay frequency ", pay_frequency)
call write_bullet_and_variable_in_case_note("Anticipated first check date ", first_check)
call write_bullet_and_variable_in_case_note("Explanation of other pay frequency ", frequency_explain)
call write_bullet_and_variable_in_case_note("Other notes on income and budget", income_notes)		
call write_variable_in_case_note (worker_signature)

		
script_end_procedure ("")
