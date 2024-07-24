Attribute VB_Name = "MainModule"
Option Explicit

Sub asOfDate_Click()
    CalendarModule.Launch
    
    If Not frmETRcalendar Is Nothing Then
        If frmETRcalendar.UserSelectedDateStr <> "" Then
            ' Change string to date
            Dim strDate As String
            strDate = frmETRcalendar.UserSelectedDateStr
            
            Dim formatDate As Date
            formatDate = DateSerial(CInt(Right(strDate, 4)), CInt(Left(strDate, 2)), CInt(Mid(strDate, 4, 2)))
            
            ThisWorkbook.Sheets("PIVOT").Range("C2").Value = formatDate
        End If
        Unload frmETRcalendar
    End If
End Sub

Sub dashboardOpenPage2()
    Dim ws As Worksheet
    Dim page1Shapes As Shape
    Dim page2Shapes As Shape
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("DASHBOARD")

    ' Attempt to reference the page 1 shapes by name
    On Error Resume Next
    Set page1Shapes = ws.Shapes("Grp_Pg1")
    On Error GoTo 0
    
    ' Hide page 1
    ' Check if the shape was found
    If Not page1Shapes Is Nothing Then
        If page1Shapes.Visible Then
            page1Shapes.Visible = msoFalse
            Debug.Print "Grouped shape hidden: " & page1Shapes.Name
        Else
            Debug.Print "Grouped shape " & page1Shapes.Name & " already hidden."
        End If
    Else
        MsgBox "Grouped shape " & page1Shapes.Name & " not found."
    End If
    
    ' Attempt to reference the page 2 shapes by name
    On Error Resume Next
    Set page2Shapes = ws.Shapes("Grp_Pg2")
    On Error GoTo 0
    
    ' Show page 2
    ' Check if the shape was found
    If Not page2Shapes Is Nothing Then
        If Not page2Shapes.Visible Then
            page2Shapes.Visible = msoTrue
            Debug.Print "Grouped shape shown: " & page2Shapes.Name
        Else
            Debug.Print "Grouped shape " & page2Shapes.Name & " already shows."
        End If
    Else
        MsgBox "Grouped shape " & page2Shapes.Name & " not found."
    End If
End Sub

Sub dashboardOpenPage1()
    Dim ws As Worksheet
    Dim page1Shapes As Shape
    Dim page2Shapes As Shape
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("DASHBOARD")

    ' Attempt to reference the page 2 shapes by name
    On Error Resume Next
    Set page2Shapes = ws.Shapes("Grp_Pg2")
    On Error GoTo 0
    
    ' Hide page 1
    ' Check if the shape was found
    If Not page2Shapes Is Nothing Then
        If page2Shapes.Visible Then
            page2Shapes.Visible = msoFalse
            Debug.Print "Grouped shape hidden: " & page2Shapes.Name
        Else
            Debug.Print "Grouped shape " & page2Shapes.Name & " already hidden."
        End If
    Else
        MsgBox "Grouped shape " & page2Shapes.Name & " not found."
    End If
    
    ' Attempt to reference the page 1 shapes by name
    On Error Resume Next
    Set page1Shapes = ws.Shapes("Grp_Pg1")
    On Error GoTo 0
    
    ' Show page 2
    ' Check if the shape was found
    If Not page1Shapes Is Nothing Then
        If Not page1Shapes.Visible Then
            page1Shapes.Visible = msoTrue
            Debug.Print "Grouped shape shown: " & page1Shapes.Name
        Else
            Debug.Print "Grouped shape " & page1Shapes.Name & " already shows."
        End If
    Else
        MsgBox "Grouped shape " & page1Shapes.Name & " not found."
    End If
End Sub

Sub UIHide()
    With Application
        .WindowState = xlMaximized
        .ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",False)"
        .DisplayStatusBar = False
        .DisplayScrollBars = False
        .DisplayFormulaBar = False
    End With
    With ActiveWindow
        .DisplayWorkbookTabs = False
        .DisplayHeadings = False
        .DisplayRuler = False
        .DisplayFormulas = False
        .DisplayGridlines = False
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = True
    End With
End Sub

Sub UIShow()
    With Application
        .ExecuteExcel4Macro "Show.Toolbar(""Ribbon"",True)"
        .DisplayStatusBar = True
        .DisplayScrollBars = True
        .DisplayFormulaBar = True
        .WindowState = xlMaximized
    End With
    With ActiveWindow
        .DisplayWorkbookTabs = True
        .DisplayHeadings = True
        .DisplayRuler = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
End Sub

Sub HideOrShowUI()
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        If .CommandBars("Ribbon").Visible Or ActiveWindow.DisplayWorkbookTabs Then
            UIHide
        Else
            UIShow
        End If
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
