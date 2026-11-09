Option Explicit

' Install in Microsoft Excel Objects > Dashboard (worksheet module).
' This is not a standard module. Do not create two Worksheet_Change handlers.
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Failed
    JGB_HandleDashboardSourceChange Me, Target
    Exit Sub
Failed:
    Application.EnableEvents = True
    MsgBox "Series dropdown update failed: " & Err.Description, _
        vbExclamation, "JGB Chart Controller"
End Sub

