VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmbBoxCustName_Change()
    If cmbBoxCustName.MatchFound Then
        Dim custIdRange As Range
        Set custIdRange = ThisWorkbook.Sheets("CUSTOMER_DATA").Range("customer_name_list").Offset(columnOffset:=-1)
    
        ThisWorkbook.Sheets("PIVOT").Range("C101").Value = custIdRange.Cells(cmbBoxCustName.ListIndex + 1).Value
    End If
End Sub

Private Sub cmbBoxCustName_LostFocus()
    ' Checks if typed value matches the list
    If Not cmbBoxCustName.MatchFound Then
        MsgBox "Typed name is not in the customer database." & vbNewLine & "Please re-enter other names!", vbCritical + vbOKOnly, "AR Dashboard"
        cmbBoxCustName.Activate
        Exit Sub
    End If
End Sub

Private Sub Worksheet_Activate()
    Me.ScrollArea = "A1:W40"
    cmbBoxCustName.List = ThisWorkbook.Names("customer_name_list").RefersToRange.Value
End Sub
