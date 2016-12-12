VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'----------------------------------------------------------------------------------------------------------------
'   Worksheet Event Module, that pre-checks user inputs and maintains the mode changes
'
'   Written By Krzysztof Grzeslak 05/11/2015
'----------------------------------------------------------------------------------------------------------------

Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Ensure that App was initialized
    Call InitializeApp
    
    ' Check each row, where data was changed
    Dim row As Range
    For Each row In Target.Rows
        Call checkRange(row)
    Next row
    
End Sub

' Checks single row for
Private Sub checkRange(Target As Range)
    Const CALC_MANUAL As String = "Manual"
    Const CALC_AUTO As String = "Automatic"
    
    Dim rowIndex As Long
    rowIndex = Target.row
    If rowIndex > Range(HEADER_NAME).row Then
        
        Dim PalletDimCell As Range
        Set PalletDimCell = getRng(Me, rowIndex, PALLET_DIMENSIONS_NAME)
        
        ' Checks if pallet nominal dimensions is in correct format
        If InStr(PalletDimCell.value, PALLET_DIM_SEPARATOR) = 0 And PalletDimCell.value <> vbNullString Then
            Call RangeModificationInitiate
                PalletDimCell.value = vbNullString
            Call RangeModificationTerminate
            Exit Sub
        End If
        
        Dim InputCell As Range
        Set InputCell = getRng(Me, rowIndex, DATA_INPUT_NAME)
        
        ' Checks if the mode was changed and calls the corresponding procedure
        If Not Intersect(Target, InputCell) Is Nothing Then
            Call RangeModificationInitiate
                Select Case InputCell.value
                    Case CALC_MANUAL
                        Call SetToManual(Target)
                    Case CALC_AUTO
                        Call SetToAutomatic(Target)
                    ' In case of any other value, set it back to automatic
                    Case Else
                        InputCell.value = CALC_AUTO
                        Call SetToAutomatic(Target)
                End Select
            Call RangeModificationTerminate
        End If
            
            ' Ensure user input all necessary parameters
            If Not isInputCorrect(Me, rowIndex) Then
                    Call ClearCells(Target) ' Delete all previous values from the calculated cells (to avoid possible errors)
                    Exit Sub
            End If
            
            ' If row is set to the automatic mode
            If Cells(Target.row, Range(DATA_INPUT_NAME).Column).value = CALC_AUTO Then
                    
                Call CalculateSingle(Target)
                    
            End If
    End If

End Sub

' Ensure user input all necessary parameters
Private Function isInputCorrect(ws As Worksheet, rowIndex As Long) As Boolean
    isInputCorrect = False
    
    Dim inputArray As Variant
    inputArray = Array(BOX_WIDTH_NAME, BOX_DEPTH_NAME, BOX_HEIGHT_NAME, MAX_HEIGHT_NAME)
    
    Dim i As Integer
    For i = LBound(inputArray) To UBound(inputArray)
        If getRngVal(ws, rowIndex, inputArray(i)) = vbNullString Then
            Exit Function
        End If
    Next i
    isInputCorrect = True
    
End Function
