VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'   *   Module includes only the Cell and Row buttons. Columns and other contexts
'       can be added by modifying parts of the code.
'----------------------------------------------------------------------------------------------------------------

Option Explicit

' Customizable button tag. Adjust to meet Your purpose
Public Enum e_buttonTag
    e_buttonTag_First
    e_buttonTag_addRow = e_buttonTag_First
    e_buttonTag_deleteRow
    e_buttonTag_modeChange
    e_buttonTag_tpdMode
    e_buttonTag_cellChange
    e_buttonTag_oversize
    e_buttonTag_Last = e_buttonTag_oversize
End Enum

' Customizable button faces. Number should correspond with the correct FaceID
' List can be found here: http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
Public Enum e_buttonFace
    e_buttonFace_Change = 1020
    e_buttonFace_Menu = 1750
    e_buttonFace_Approve = 1087
    e_buttonFace_Discard = 1088
    e_buttonFace_Help = 1089
End Enum

' Abstration for contextMenu enumerated values
Public Enum e_contextMenu
    e_contextMenu_First = 0
    e_contextMenu_Cell = 0
    e_contextMenu_Column = 1
    e_contextMenu_Row = 2
    e_contextMenu_ListRange = 3
    e_contextMenu_Last = 3
End Enum
    
' Class fields to store button property values
Private m_name As String
Private m_desc As String
Private m_face As e_buttonFace
Private m_tag As e_buttonTag

Private wasInitialized As Boolean
Private Const msoControlButton = 1

' c_Button Properties Getters / Setters
Public Property Get GetMacroName() As String
    GetMacroName = m_name
End Property

Private Property Let setMacroName(value As String)
    m_name = "'" & ThisWorkbook.name & "'!" & value
End Property

Private Property Get GetDescription() As String
    GetDescription = m_desc
End Property

Private Property Let setDescription(value As String)
    m_desc = value
End Property

Public Property Get GetTag() As e_buttonTag
    GetTag = m_tag
End Property

Private Property Let setTag(value As e_buttonTag)
    m_tag = value
End Property

Public Property Get GetFaceID() As e_buttonFace
    GetFaceID = m_face
End Property

Private Property Let setFaceID(value As e_buttonFace)
    m_face = value
End Property

' Button initializer with provided property values
Public Sub Initialize(macroName As String, description As String, tag As e_buttonTag, faceID As e_buttonFace)
   
    setMacroName = macroName
    setDescription = description
    setTag = tag
    setFaceID = faceID
    wasInitialized = True

End Sub
  
' Adds new button to the e_contextMenu object and counts the instances
Public Sub AddToMenu(ByVal m_menu As e_contextMenu, ByRef m_counter As C_ButtonCounter)
   
    If Not wasInitialized Then
        Debug.Print "Button was not initialized!"
        Exit Sub
    End If
    
    Dim contextMenu As Object
    Set contextMenu = getContextMenu(m_menu)

    Dim iBefore As Integer
    iBefore = m_counter.InstanceCount(m_menu)

    With contextMenu.Controls.Add(Type:=msoControlButton, before:=iBefore)
        .OnAction = GetMacroName
        .faceID = GetFaceID
        .caption = GetDescription
        .tag = GetTag
    End With
    
    Call m_counter.IncrementCounter(m_menu)

End Sub

'Retrieve specific e_contextMenu object
Private Function getContextMenu(ByVal m_menu As e_contextMenu)

    Select Case m_menu
        Case e_contextMenu_Cell
            Set getContextMenu = Application.CommandBars("Cell")
        Case e_contextMenu_Column
            Set getContextMenu = Application.CommandBars("Column")
        Case e_contextMenu_Row
            Set getContextMenu = Application.CommandBars("Row")
        Case e_contextMenu_ListRange
            Set getContextMenu = Application.CommandBars("List Range Popup")
    End Select
    
End Function

' Deletes all the buttons in all e_contextMenus with defined e_buttonTags
Public Sub DeleteAllButtons()

    Dim menu As Object
    Dim button As Object
    Dim menuIndex As Integer
    Dim tagIndex As Integer
    
    For menuIndex = e_contextMenu_First To e_contextMenu_Last
        Set menu = getContextMenu(menuIndex)
        For Each button In menu.Controls
            For tagIndex = e_buttonTag_First To e_buttonTag_Last
                If button.tag = tagIndex Then
                    button.delete
                    Exit For
                End If
            Next tagIndex
            
        Next button
    Next menuIndex
End Sub

' Insert the interline after the last added button and resets the counter.
Public Sub CreateGroup(ByVal m_menu As e_contextMenu, ByRef buttonCounter As C_ButtonCounter)
    
    Dim menu As Object
    Dim iButton As Integer
    
    Set menu = getContextMenu(m_menu)
    iButton = buttonCounter.InstanceCount(m_menu)
    
    menu.Controls(iButton).BeginGroup = True
    Call buttonCounter.ResetCounter(m_menu)
End Sub
