Attribute VB_Name = "ButtonManager"
'----------------------------------------------------------------------------------------------------------------
'   Provides abstraction for Excel button object
'   Enables to easily add new buttons to the Excel contextMenu
'
'   Written By Krzysztof Grzeslak 05/11/2015
'
'   Preconditions:
'   *   Excel macro file must include all three cooperating modules: c_Button, c_ButtonCounter and ButtonManager
'   *   Excel macro file This_Workbook Object should include following statements:
'       Workbook_Activate: Call ButtonManager.RemoveButtons, Call ButtonManager.AddButtons
'       Workbook_Deactivate: Call ButtonManager.RemoveButtons
'
'   Usage:
'   *   Functions RowButtonDataArray and CellButtonDataArray should be modified to hardcode the buttons properties
'   *   Module includes only the Cell and Row buttons. Columns can be added by modified parts of the code.
'----------------------------------------------------------------------------------------------------------------

Option Explicit

Private m_wasError As Boolean

Public Sub AddButtons(Optional internalProcedure As Boolean = True)
    Dim buttonCnt As C_ButtonCounter
    Set buttonCnt = New C_ButtonCounter 'Start new button counter
    
    Dim Buttons As Variant
    Buttons = ButtonDataArray 'retrieve all row buttons data and create array
    Call AddButtonGroup(Buttons, e_contextMenu_Row, buttonCnt)
    
    ' Checking for SolidWorks Design Table. If user opened the file in the SW window, then no button can be added
    ' User must open the table in new window
    If m_wasError Then
        Dim msg As String
        msg = "To be able to add new rows or delete existing ones" + vbNewLine
        msg = msg + "please edit the DesignTable in the separate window"
        MsgBox msg, vbInformation
        Exit Sub
    End If
    
    Call AddButtonGroup(Buttons, e_contextMenu_Cell, buttonCnt)
    Call AddButtonGroup(Buttons, e_contextMenu_ListRange, buttonCnt)

    Set buttonCnt = Nothing
End Sub

Public Sub RemoveButtons(Optional internalProcedure As Boolean = True)
    Dim button As C_Button
    Set button = New C_Button
    
    Call button.DeleteAllButtons
    Set button = Nothing
    
End Sub

Private Sub AddButtonGroup(ByVal buttonArray As Variant, ByVal menu As e_contextMenu, ByRef counter As C_ButtonCounter)
    'Extracts buttons information from array and adds them to chosen contextMenu
    Dim button As C_Button
    Set button = New C_Button
    
    Dim macroName As String
    Dim buttonCaption As String
    Dim buttonTag As e_buttonTag
    Dim buttonFace As e_buttonFace
    Dim buttonIndex As Integer
    
    On Error Resume Next
        For buttonIndex = LBound(buttonArray) To UBound(buttonArray)
            macroName = buttonArray(buttonIndex)(0)
            buttonCaption = buttonArray(buttonIndex)(1)
            buttonTag = buttonArray(buttonIndex)(2)
            buttonFace = buttonArray(buttonIndex)(3)
    
            With button
                Call .Initialize(macroName, buttonCaption, buttonTag, buttonFace)
                Call .AddToMenu(menu, counter)
            End With
        Next buttonIndex
        
        Call button.CreateGroup(menu, counter)
        If Err <> 0 Then m_wasError = True
    On Error GoTo 0
    
    Set button = Nothing
End Sub

'Function is a list of hardcoded row menu button data. Array is created from the list and provided as output.
Private Function ButtonDataArray() As Variant

    Dim tempButtonArray() As Variant
    tempButtonArray = Array( _
    Array("AddNewRow", "Insert row(s)", e_buttonTag_addRow, e_buttonFace_Menu), _
    Array("DeleteRow", "Delete Row(s)", e_buttonTag_deleteRow, e_buttonFace_Menu), _
    Array("changeOversizeModifier", "Modify oversize modifier", e_buttonTag_oversize, e_buttonFace_Menu), _
    Array("InitMaterialList", "Modify Pallet Sizes Database", e_buttonTag_oversize, e_buttonFace_Menu) _
    )
    
    ButtonDataArray = tempButtonArray

End Function
