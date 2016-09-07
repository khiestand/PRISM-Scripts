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

