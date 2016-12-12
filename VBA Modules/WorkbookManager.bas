Attribute VB_Name = "WorkbookManager"
'---------------------------------------------------------------------------------------------------------------------------------
'   Module controls the behaviour of the Excel Workbook and Worksheet Objects. It also provided some abstraction for
'   Excel basic procedures
'
'   Written By Krzysztof Grzeslak 10/08/2015
'
'---------------------------------------------------------------------------------------------------------------------------------

Option Explicit

    Private m_wasInitialized As Boolean
    Private Const PASSWORD = "" ' Please enter Your desired password for securing the Worksheet and Workbook

Private Sub InitMaterialList()
    Const SETUP_SHEET_NAME As String = "Settings"
    
    Dim setupWs As Worksheet
    Set setupWs = getWorkbook.Worksheets(SETUP_SHEET_NAME)
    Call f_MaterialList.Initialize(setupWs)

End Sub

' Implementation hide for retrieving Workbook object
Public Function getWorkbook() As Workbook
    
    Set getWorkbook = ThisWorkbook
    
End Function

' Implementation hide for retrieving main Worksheet object
Public Function getMainWorksheet() As Worksheet
    Const WORKSHEET_NAME As String = "DesignTable"
    
    Dim wkb As Workbook
    Set wkb = getWorkbook
    Set getMainWorksheet = wkb.Worksheets(WORKSHEET_NAME)

End Function

' Perform initials procedures required for proper application performance
Public Function InitializeApp()
    
    If m_wasInitialized Then Exit Function
    
    Dim wkb As Workbook
    Set wkb = getWorkbook
    
    Dim ws As Worksheet
    Set ws = getMainWorksheet
    
    ' password need to be set at each application start, so the code can alter the worksheet contents
    Call SetPassword(wkb, ws)

    ' Hide all the the named ranges from the user
    Dim nm As name
    For Each nm In wkb.Names
        nm.Visible = False
    Next

    m_wasInitialized = True

End Function

' Password management procedures
Private Sub SetPassword(wkb As Workbook, ws As Worksheet)
    
    Call lockWorksheet(ws)
    Call lockWorkbook(wkb)

End Sub

Public Sub unlockWorksheet(ws As Worksheet)
    
    ws.Unprotect PASSWORD
    
End Sub

Public Sub lockWorksheet(ws As Worksheet)
    
    ws.Protect PASSWORD, UserInterfaceOnly:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    
End Sub

Private Sub unlockWorkbook(wkb As Workbook)
    
    wkb.Unprotect PASSWORD
    
End Sub

Private Sub lockWorkbook(wkb As Workbook)
    
    wkb.Protect PASSWORD, True, False
    
End Sub

' Ensure user won't see the flicker, app will be faster and any event won't stop the app run
Public Sub RangeModificationInitiate()
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
End Sub

' Return to standard values
Public Sub RangeModificationTerminate()
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

' If underlay is used, then user can set Workbook Global Variable, saying how many mm should the underlay be biger than:
' * the pallet itself, if the load size is smaller than the pallet
' * the load, if the load size is bigger than the pallet

Public Function changeOversizeModifier()
    Const UNDERLAY_MODIFIER_NAME As String = "UnderlayOversizeModifier"
    
    If Not m_wasInitialized Then Call InitializeApp
    Dim wkb As Workbook
    Set wkb = getWorkbook
    
    Dim name As name
    Set name = wkb.Names(UNDERLAY_MODIFIER_NAME)
    
    ' Determine current value of the modifier
    Dim oldValue As Integer
    oldValue = Replace(name.RefersTo, "=", vbNullString)
    
    ' Prepare the msg for the input box
    Dim msg As String
    msg = "Current Underlay Oversize Modifier value: " & oldValue & vbNewLine
    msg = msg & "Please specify new value:"
    
    Dim newValue As Variant
    
    ' Ask user for new modifier value
    Do
        newValue = InputBox(prompt:=msg)
        If newValue = vbNullString Then Exit Function   ' if nothing is given or cancel pressed, then the value is not changed
    Loop While Not IsNumeric(newValue)                  ' Ensure user will give numeric value
    
    name.RefersTo = newValue
    
End Function

' To ensure, that user won't compromise the application, worksheet should be secure
' Adding and removing rows are done through macros

Public Sub AddNewRow()
    
    ' Ask user how many rows need to be added
    Dim newRowsQtty As Long
    newRowsQtty = Application.InputBox("How many rows add?", "Add new rows", Type:=1)
        
    ' If less than one, then exit procedure
    If newRowsQtty < 1 Then Exit Sub
    
    If Not m_wasInitialized Then Call InitializeApp
    
    Dim ws As Worksheet
    Set ws = getMainWorksheet
    
    ' Determine index of the last row with data
    Dim lastRowIndex As Long
    lastRowIndex = ws.Rows(ws.Rows.Count).End(xlUp).row
    
    ' Determine index of the first new row
    Dim newRowIndex As Long
    newRowIndex = lastRowIndex + 1
    
    Call RangeModificationInitiate
        Dim i As Integer
        For i = 0 To newRowsQtty - 1
            ' Populate new row with default values and new name
            Call PalletCalculation.DefaultValues(ws.Rows(newRowIndex + i))
            getRng(ws, newRowIndex + i, "ModelName").value = "New model" & (i + 1)
            
            ' Ensure the cell formating is correct for automatic mode
            Call PrepareCollections(ws.Rows(newRowIndex + i))
            Call SetToAutomatic(ws.Rows(newRowIndex + i))
        Next i
    Call RangeModificationTerminate

End Sub

Public Function DeleteRow()
    If Not m_wasInitialized Then Call InitializeApp
    
    Dim ws As Worksheet
    Set ws = getMainWorksheet
    
    Dim firstRow As Integer
    firstRow = ws.Range(HEADER_NAME).row
    
    ' Header rows cannot be deleted
    If ActiveCell.row <= firstRow Then Exit Function
    
    ' Determine the range, that will be deleted
    Dim rowsToDelete As Range
    Set rowsToDelete = Intersect(Selection, ws.UsedRange)
    If rowsToDelete Is Nothing Then Exit Function
    
    ' Count how many rows are in the range to be deleted
    Dim rowsCount As Integer
    rowsCount = rowsToDelete.Rows.Count

    ' Delete all selected rows
    Dim i As Integer
    Dim lastRowIndex As Long
    For i = rowsCount To 1 Step -1
        lastRowIndex = ws.Rows(ws.Rows.Count).End(xlUp).row
        
        ' at least one row with data must remain, so the function won't be deleted
        If lastRowIndex <= 4 Then
            MsgBox "Cannot delete last data row!", vbCritical
            Exit Function
        End If
                
        ' delete each row one by one
        Call RangeModificationInitiate
            rowsToDelete.Rows(i).EntireRow.delete
        Call RangeModificationTerminate
    Next i

End Function
