BeginDialog initial_dialog, 0, 0, 206, 85, "Initial Dialog"
  Text 5, 5, 50, 10, "Intake Script"
  Text 15, 20, 45, 10, "Case Number"
  EditBox 80, 15, 115, 15, case_number_edit
  Text 15, 40, 70, 10, "Type of Intake Action"
  DropListBox 100, 40, 95, 15, "Establishment"+chr(9)+"Enforcement"+chr(9)+"Motion to Set"+chr(9)+"Paternity", type_intake_drpdwn
  ButtonGroup ButtonPressed
    OkButton 95, 65, 50, 15
    CancelButton 150, 65, 50, 15
EndDialog

BeginDialog intake_enforcement_dialog, 0, 0, 401, 355, "Enforcement Intake Dialog"
  CheckBox 20, 35, 145, 10, "Case Opening - Welcome Letter (*.docx)", CP_welcome_ltr_checkbox
  CheckBox 20, 50, 130, 10, "Child Care Verification (*.docx)", child_care_verif_checkbox
  CheckBox 20, 65, 140, 10, "Court Order Summary Letter (*.docx)", CP_new_order_summary_checkbox
  CheckBox 20, 80, 155, 10, "Health Insurance Verification (F0924)", CP_healthverif_checkbox
  CheckBox 20, 95, 100, 10, "Pin Notice (F0999)", CP_pinnotice_checkbox
  CheckBox 215, 35, 110, 10, "Arrears Amount Letter (*.docx)", NCP_arrearsletter_checkbox
  CheckBox 215, 50, 150, 10, "Case Opening - Introduction Letter (*.docx)", NCP_welcomeltr_checkbox
  CheckBox 215, 65, 130, 10, "Court Order Summary Letter (*.docx)", NCP_courtordersummary_checkbox
  CheckBox 215, 80, 155, 10, "Health Insurance Verfication (F0924)", NCP_healthverif_checkbox
  CheckBox 215, 95, 100, 10, "Pin Notice (F0999)", NCP_pinnotice_checkbox
  CheckBox 225, 150, 140, 10, "Authorization to Collect Support (F0100)", authtocollect_F0100_checkbox
  CheckBox 225, 165, 170, 10, "Notice of Child Support/Spousal Liability (F0108)", notice_csandspousal_dord_F0108_checkbox
  CheckBox 225, 190, 165, 10, "Notification of Parental Liability Notice (F0109)", notification_of_parental_liability_dord_F0109_checkbox
  CheckBox 225, 220, 165, 10, "Notification of Medical Support Liability (F0107)", notification_of_medical_liability_dord_F0107_checkbox
  EditBox 35, 150, 135, 15, worklist_description_01
  EditBox 115, 170, 30, 15, calendar_days_01
  EditBox 35, 205, 135, 15, worklist_description_02
  EditBox 115, 225, 30, 15, calendar_days_02
  EditBox 35, 260, 135, 15, worklist_description_03
  GroupBox 210, 130, 185, 110, "Send Liability Notice to NCP:"
  Text 5, 5, 95, 10, "Enforcement Intake"
  GroupBox 210, 20, 185, 90, "Documents Sending to NCP:"
  Text 35, 285, 85, 10, "Calendar days until due"
  EditBox 115, 280, 30, 15, calendar_days_03
  EditBox 220, 250, 95, 15, filelocation_editbox
  EditBox 220, 290, 135, 15, caad_note_editbox
  EditBox 295, 310, 65, 15, signature_editbox
  ButtonGroup ButtonPressed
    OkButton 245, 335, 50, 15
    CancelButton 300, 335, 50, 15
  Text 35, 175, 80, 10, "Calendar days until due"
  Text 25, 195, 80, 10, "Worklist Description 2"
  Text 220, 210, 50, 10, "MA Only Cases"
  Text 35, 230, 85, 15, "Calendar days utnil due"
  Text 220, 140, 50, 10, "NPA Cases"
  Text 25, 250, 85, 10, "Worklist Description 3"
  Text 220, 180, 50, 10, "PA Cases"
  Text 220, 240, 75, 10, "File Location on CAST"
  GroupBox 10, 20, 180, 90, "Documents Sending to CP:"
  Text 220, 270, 135, 20, "Additional text to CAAD note (Docs sent will automatically list in CAAD Note):"
  GroupBox 15, 125, 175, 175, "CAWD Notes to Add:"
  Text 220, 315, 70, 10, "Sign your CAAD note:"
  Text 25, 140, 80, 15, "Worklist Description 1"
EndDialog
