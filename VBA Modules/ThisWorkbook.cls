VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------------------------------
'   Workbook Event Module, that Initializes the Application and maintains custom added buttons
'
'   Written By Krzysztof Grzeslak 05/11/2015
'----------------------------------------------------------------------------------------------------------------

Option Explicit

Private Sub Workbook_Activate()
    Call InitializeApp
    Call RemoveButtons
    Call AddButtons
End Sub

Private Sub Workbook_Deactivate()

    Call RemoveButtons
    
End Sub
