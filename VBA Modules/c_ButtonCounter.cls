VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_ButtonCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------
'   Provides counting mechanism of the c_Button Class Module
'
'   Written By Krzysztof Grzeslak 05/11/2015
'
'   Preconditions:
'   *   Excel macro file must include all three cooperating modules: c_Button, c_ButtonCounter and ButtonManager
'
'   Usage:
'   *   Button counting is automated from ButtonManager
'   *   Counter counts all new instances of the button, until the reset.
'
'----------------------------------------------------------------------------------------------------------------

Private m_isInitialized As Boolean
Private m_rowCount As Integer
Private m_cellCount As Integer
Private m_columnCount As Integer
Private m_listRangeCount As Integer
Private Const DEFAULT_VALUE As Integer = 1

Public Property Get InstanceCount(i_menu As e_contextMenu) As Integer
    
    If Not m_isInitialized Then
        Call InitializeCounter
    End If
    
    Select Case i_menu
        Case e_contextMenu_Cell
            InstanceCount = m_cellCount
        Case e_contextMenu_Column
            InstanceCount = m_columnCount
        Case e_contextMenu_Row
            InstanceCount = m_rowCount
        Case e_contextMenu_ListRange
            InstanceCount = m_listRangeCount
    End Select
End Property

Public Sub IncrementCounter(menu As e_contextMenu)
    
    If Not m_isInitialized Then
        Call InitializeCounter
    End If
    
    Select Case menu
        Case e_contextMenu_Cell
            m_cellCount = m_cellCount + 1
        Case e_contextMenu_Column
            m_columnCount = m_columnCount + 1
        Case e_contextMenu_Row
            m_rowCount = m_rowCount + 1
        Case e_contextMenu_ListRange
            m_listRangeCount = m_listRangeCount + 1
    End Select

End Sub

Public Sub ResetCounter(menu As e_contextMenu)

    Select Case menu
        Case e_contextMenu_Cell
            m_cellCount = DEFAULT_VALUE
        Case e_contextMenu_Column
            m_columnCount = DEFAULT_VALUE
        Case e_contextMenu_Row
            m_rowCount = DEFAULT_VALUE
        Case e_contextMenu_ListRange
            m_listRangeCount = DEFAULT_VALUE
    End Select

End Sub

Private Sub InitializeCounter()

    Call ResetCounter(e_contextMenu_Cell)
    Call ResetCounter(e_contextMenu_Column)
    Call ResetCounter(e_contextMenu_Row)
    Call ResetCounter(e_contextMenu_ListRange)
    
    m_isInitialized = True

End Sub
