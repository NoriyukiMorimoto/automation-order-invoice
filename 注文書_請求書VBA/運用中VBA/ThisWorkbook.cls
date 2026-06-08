Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    mod_FillManagerName.CancelScheduledOfficeComboBoxPrompt
    mod_VendorMaster.CancelScheduledVendorListDropdown
    mod_BasicInfoCalendar.ClearProjectSelectionState
    On Error GoTo 0
End Sub
