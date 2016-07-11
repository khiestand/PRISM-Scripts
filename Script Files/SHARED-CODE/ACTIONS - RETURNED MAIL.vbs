'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - RETURNED MAIL.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 60
STATS_denomination = "C"
'END OF STATS BLOCK-------------------------------------------------------------------------------------------


'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Robert Kalb, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Robert Kalb and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Robert will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF


'Calling the Returned Mail Dialog--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog returned_mail_dialog, 0, 0, 356, 195, "Returned Mail Received"
  EditBox 85, 5, 95, 15, prism_number
  CheckBox 100, 30, 25, 10, "CP", rm_cp_checkbox
  CheckBox 140, 30, 40, 10, "NCP/ALF", rm_ncp_checkbox
  EditBox 225, 25, 80, 15, rm_other
  DropListBox 70, 50, 150, 15, "Select one..."+chr(9)+"Update to Unknown"+chr(9)+"Update to New Forwarding Address", updated_ADDR
  EditBox 75, 80, 125, 15, new_ADDR
  EditBox 75, 100, 120, 15, new_CITY
  EditBox 75, 120, 25, 15, new_STATE
  EditBox 160, 120, 45, 15, new_ZIP
  EditBox 230, 60, 75, 15, date_received
  EditBox 260, 80, 45, 15, source
  EditBox 295, 105, 50, 15, postal_resp
  EditBox 100, 150, 240, 15, misc_notes
  EditBox 75, 170, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 235, 170, 50, 15
    CancelButton 290, 170, 50, 15
  Text 10, 10, 70, 10, "PRISM Case Number:"
  Text 10, 30, 85, 10, "Returned Mail Rec'd for:"
  Text 200, 30, 20, 10, "Other:"
  Text 10, 50, 60, 10, "Update Address?"
  Text 230, 50, 80, 10, "Effective/Verified Date:"
  GroupBox 5, 65, 215, 80, "New Address Info (If given by Post Office)"
  Text 20, 85, 50, 10, "Street Address:"
  Text 55, 105, 20, 10, "City:"
  Text 50, 125, 20, 10, "State:"
  Text 125, 125, 35, 10, "Zip Code:"
  Text 230, 85, 30, 10, "Source:"
  Text 230, 110, 60, 10, "Postal Response:"
  Text 10, 155, 90, 10, "Misc notes/Actions Taken:"
  Text 10, 175, 65, 10, "Worker Signature:"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Connecting to Bluezone
EMConnect ""			

CALL check_for_prism(TRUE)
CALL PRISM_case_number_finder(prism_number)



'The script will not run unless the mandatory fields are completed
DO
	err_msg = ""
	Dialog returned_mail_dialog
	IF ButtonPressed = 0 THEN StopScript		                                       
	CALL PRISM_case_number_validation(prism_number, case_number_valid)
	IF case_number_valid = FALSE THEN err_msg = err_msg & vbNewline & "You must enter a valid PRISM case number!"
	IF forwarding_addr = "Select One..." THEN err_msg = err_msg & vbNewline & "You must answer if there was a forwarding address given!"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewline & "You sign your CAAD note!"       
	IF date_received = "" THEN err_msg = err_msg & vbNewline & "You must enter a effective/verified date."            
	IF err_msg <> "" THEN MsgBox "***NOTICE***" & vbNewLine & err_msg & vbNewline & vbNewline & "Please resolve for the script to continue!"
LOOP UNTIL err_msg = ""

'Checks to make sure PRISM is open and you are logged in
CALL check_for_PRISM(True)

CALL navigate_to_PRISM_screen("REGL")
CALL navigate_to_PRISM_screen("CAST")
EMWritescreen "D", 3, 29
EMWritescreen prism_number, 4, 8
transmit

'Navigates to CPDD, NCDD or caad note for other address
IF rm_cp_checkbox = CHECKED THEN 
'Do we need to add a new address or set address to unknown
	
	CALL navigate_to_PRISM_screen("CAAD")
	PF5
	EMWritescreen date_received, 4, 37
	EMWritescreen "R0011", 4, 54
	EMSetCursor 16, 4
	CALL write_new_line_in_PRISM_case_note(misc_notes)
	CALL write_new_line_in_PRISM_case_note("---" & worker_signature)
	transmit



	IF updated_ADDR = "Update to New Forwarding Address" THEN
		CALL navigate_to_PRISM_screen("CPDD")
		EMWritescreen "M", 3, 29
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit

		CALL navigate_to_PRISM_screen("CPDD")
		EMwritescreen "M", 3, 29
		EMWritescreen date_received, 10, 18
		EMWritescreen "Y", 10, 46
		EMwritescreen new_addr, 15, 11
		EMWritescreen new_CITY, 17, 11
		EMWritescreen new_STATE, 17, 39
		EMWritescreen new_ZIP, 17, 50
		EMWritescreen date_received, 19, 7
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit

		EMReadScreen standardization_msg, 6, 4, 35
		IF standardization_msg = "Code-1" THEN
			EMReadscreen error_msg, 29, 12, 25
			IF trim(error_msg) <> "" OR trim(error_msg) <> "Address has been standardized" THEN
				PF6
				PF3	
				Msgbox "PRISM reports this message: " & trim(error_msg) & ". Please verify and/or update the address if applicable."
				stopscript
			END IF	
			PF6
		END IF

		MsgBox "** Review case to see if maintaining county request needs to be made **"

	ELSEIF updated_ADDR = "Update to Unknown" THEN
		CALL navigate_to_PRISM_screen("CPDD")
		EMWritescreen "M", 3, 29
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit

		EMWritescreen "M", 3, 29	
		EMWritescreen date_received, 10, 18
		EMWritescreen "N", 10, 46	
		EMSetCursor 14, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 15, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 16, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 39	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 50	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 56	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 69	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 7	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 38	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 62	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		transmit
	END IF
END IF

IF rm_ncp_checkbox = CHECKED THEN
	CALL navigate_to_PRISM_screen("CAAD")
	PF5
	EMWritescreen date_received, 4, 37
	EMWritescreen "R0010", 4, 54
	EMSetCursor 16, 4
	CALL write_new_line_in_PRISM_case_note(misc_notes)
	CALL write_new_line_in_PRISM_case_note("---" & worker_signature)
	transmit

	IF updated_ADDR = "Update to New Forwarding Address" THEN
		CALL navigate_to_PRISM_screen("NCDD")
		EMWritescreen "M", 3, 29
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit

		EMwritescreen "M", 3, 29
		EMWritescreen date_received, 10, 18
		EMWritescreen "Y", 10, 46
		EMwritescreen new_addr, 15, 11
		EMWritescreen new_CITY, 17, 11
		EMWritescreen new_STATE, 17, 39
		EMWritescreen new_ZIP, 17, 50
		EMWritescreen date_received, 19, 7
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit
		
		EMReadScreen standardization_msg, 6, 4, 35
		IF standardization_msg = "Code-1" THEN
			EMReadscreen error_msg, 29, 12, 25
			IF trim(error_msg) <> "" OR trim(error_msg) <> "Address has been standardized" THEN
				PF6
				PF3	
				Msgbox "PRISM reports this message: " & trim(error_msg) & ". Please verify and/or update the address if applicable."
				stopscript
			END IF	
			PF6
		END IF
		
		Msgbox "** Review case to see if maintaining county request needs to be done.**"
	ELSEIF updated_ADDR = "Update to Unknown" THEN
		CALL navigate_to_PRISM_screen("NCDD")
		EMWritescreen "M", 3, 29
		EMWritescreen source, 19, 38
		EMWritescreen postal_resp, 19, 62
		transmit

		EMWritescreen "M", 3, 29	
		EMWritescreen date_received, 10, 18
		EMWritescreen "N", 10, 46	
		EMSetCursor 14, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 15, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 16, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 11	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 39	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 50	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 56	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 17, 69	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 7	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 38	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		EMSetCursor 19, 62	
		EMSendkey  "<EraseEof>"
		EMWaitReady 0, 0
		transmit
	END IF
END IF


IF rm_other <> "" THEN
	CALL navigate_to_PRISM_screen("CAAD")
	PF5
	EMWritescreen "R0012", 4, 54
	EMSetCursor 16, 4
	CALL write_new_line_in_PRISM_case_note(rm_other)
	CALL write_new_line_in_PRISM_case_note(misc_notes)
	CALL write_new_line_in_PRISM_case_note("---" & worker_signature)
	transmit
END IF


script_end_procedure("")












