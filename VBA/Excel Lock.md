# Put on Mudole


```
Option Explicit

Sub WotksheetToShow(ByVal nameOfWs As String)

    Dim wsToShow As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set wsToShow = ThisWorkbook.Worksheets(nameOfWs)
    On Error GoTo 0
    
    If Not wsToShow Is Nothing Then
        
        wsToShow.Visible = xlSheetVisible
        
        For Each ws In ThisWorkbook.Worksheets
        
            If ws.Name <> wsToShow.Name Then
            
                ws.Visible = xlSheetVeryHidden
                
            End If
            
        Next ws
        
        HideWorksheetElements
    
    End If
    
End Sub


Sub HideWorksheetElements()
    
    ActiveWindow.DisplayHeadings = False
    
End Sub

Sub RestoreElements()
    
    Application.DisplayFormulaBar = True
    
End Sub

Sub HideAll()
' Author: Topan
'
' Hide the Excel Interface

        ' Hide the Horizontal Scroll Bar
        ActiveWindow.DisplayHorizontalScrollBar = True
        
        ' Hide the Vertical Scroll Bar
        ActiveWindow.DisplayVerticalScrollBar = True
        
        ' Hide the Row/Column Headings
        ActiveWindow.DisplayHeadings = True
        
        ' Hide the Worksheet Tabs
        ActiveWindow.DisplayWorkbookTabs = False
        
        ' Hide the Status Bar
        Application.DisplayStatusBar = False
        
        ' Hide the Formula Bar
        Application.DisplayFormulaBar = False
        
        ' Hide the Ribbon Menu and Quick Access Toolbar
        Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
        
        ' Display Full Screen
        'Application.DisplayFullScreen = True
        
End Sub

Sub ShowAll()
' Author: Topan
' Show the Excel Interface
        ' Show the Horizontal Scroll Bar
        ActiveWindow.DisplayHorizontalScrollBar = True
        
        ' Show the Vertical Scroll Bar
        ActiveWindow.DisplayVerticalScrollBar = True
        
        ' Show the Row/Column Headings
        ActiveWindow.DisplayHeadings = True
        
        ' Show the Worksheet Tabs
        ActiveWindow.DisplayWorkbookTabs = True
        
        ' Show the Status Bar
        Application.DisplayStatusBar = True
        
        ' Show the Formula Bar
        Application.DisplayFormulaBar = True
        
        ' Show the Ribbon Menu and Quick Access Toolbar
        Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
        
        ' Exit Full Screen
        'Application.DisplayFullScreen = True
        
End Sub
```

# Put on Thisworkbook
```
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RestoreElements
End Sub

Private Sub Workbook_Open()
    HideAll
End Sub
```
