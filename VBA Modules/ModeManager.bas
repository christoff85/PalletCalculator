Attribute VB_Name = "ModeManager"
'----------------------------------------------------------------------------------------------------------------
'   Module implements palletizing mode change between Automatic and Manual for Selected Row in the Worksheet
'
'   Written By Krzysztof Grzeslak 05/11/2015
'
'   Preconditions:
'   *   Excel file must include module PalletCalculation, that includes constant range names
'   *   Worksheet Event module should call apropriate procedures from this module based on cell change
'----------------------------------------------------------------------------------------------------------------

Option Explicit
Private OutputParameters As Collection
Private InputParameters As Collection

Private Enum e_parType
    e_parType_Input
    e_parType_Output
End Enum

' Creates new Validate List and adds True and False values, if checked cell intersect with reference cells range
Private Sub createBoolList(cell As Range, refCell As Range)
    
    If Not Intersect(cell, refCell) Is Nothing Then
        Call RangeModificationInitiate
            With cell.Validation
                .delete
                .Add Type:=xlValidateList, Formula1:="False,True"
            End With
        Call RangeModificationTerminate
    End If
    
End Sub

' Switches the selected row to manual mode, where user can input all parameters. No calculation is made for this row
Public Sub SetToManual(Target As Range)
    Call PrepareCollections(Target)
    
    Dim ws As Worksheet
    Set ws = Target.Parent
    
    Dim cell As Range
    Call unlockWorksheet(ws)
    For Each cell In OutputParameters
        With cell
            .locked = False
            .FormulaHidden = False
            With .Font
                .Bold = True
                .Color = RGB(Red:=10, Green:=150, Blue:=100)
                .Size = "11"
            End With
        End With
        
        Call createBoolList(cell, ws.Range(ROUND_NAME))

    Next
    
    For Each cell In InputParameters
        With cell
            .locked = True
            .FormulaHidden = True
            .Font.Color = Target.EntireRow.Interior.ColorIndex
        End With
    Next
    Call lockWorksheet(ws)
    Call terminateCollections
    
End Sub

' Switches the selected row to automatic mode, where pallet parameter are based on calculation made
Public Sub SetToAutomatic(Target As Range)
    Call PrepareCollections(Target)
    
    Dim cell As Range
    For Each cell In OutputParameters
        With cell
            .locked = True
            .FormulaHidden = True
            With .Font
                .Bold = False
                .Color = RGB(Red:=0, Green:=0, Blue:=0)
                .Size = "10"
            End With
        End With
    Next
    
    For Each cell In InputParameters
        With cell
            .locked = False
            .FormulaHidden = False
            .Font.Color = RGB(Red:=0, Green:=0, Blue:=0)
        End With
    Next
    
    Call terminateCollections

End Sub

' Clears the the given cell
Public Sub ClearCells(Target As Range)
    Call PrepareCollections(Target)
    
    Call RangeModificationInitiate
        Dim cell As Range
        For Each cell In OutputParameters
            cell.value = vbNullString
        Next
    Call RangeModificationTerminate
    
    Call terminateCollections
End Sub

' Prepares the collection of named ranges for both input and output parameters
Public Sub PrepareCollections(Target As Range)
    
    ' Prepare output parametes
    Set OutputParameters = New Collection
    
    Dim rangeNames As Variant
    rangeNames = Array(LENGTH_QTTY_NAME, WIDTH_QTTY_NAME, LAYERS_QTTY_NAME, PACK_DIM_ON_LENGTH, PACK_DIM_ON_WIDTH, PACK_DIM_ON_HEIGHT, _
                    ADD_LENGTH_QTTY_NAME, ADD_WIDTH_QTTY_NAME, ROUND_NAME)
    
    Call populateCollection(OutputParameters, rangeNames, Target)
    rangeNames = vbNullString
    
    ' Prepare input parameters
    Set InputParameters = New Collection
    
    rangeNames = Array(BOX_POSITION_NAME, MAX_HEIGHT_NAME, MAX_OVERLAY_NAME)
    
    Call populateCollection(InputParameters, rangeNames, Target)
    rangeNames = vbNullString

End Sub

' Adds to collection Range Object with name in the names array
Private Sub populateCollection(ByRef collect As Collection, rangeNames As Variant, Target As Range)
    Dim i As Integer
    For i = LBound(rangeNames) To UBound(rangeNames)
        collect.Add Cells(Target.row, Range(rangeNames(i)).Column)
    Next i
End Sub

' Dispose of the collections
Private Sub terminateCollections()
    Set OutputParameters = Nothing
    Set InputParameters = Nothing
End Sub
