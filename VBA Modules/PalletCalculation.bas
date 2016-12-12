Attribute VB_Name = "PalletCalculation"
'---------------------------------------------------------------------------------------------------------------------------------
'   Provides main body for the palletizing calculation. Whole Workbook with this application is designed to work as SolidWorks
'   DesignTable, that works with parametrized Pallet 3D model.
'
'   Written By Krzysztof Grzeslak 10/09/2015
'
'   Preconditions:
'   *   Requires modules s_Pallet, c_Pallet and WorkbookManager
'
'   Usage:
'   Based on the user inputs from Excel Worksheet: package dimensions,
'   pallet requirements and package vertical orientation, application calculates all possible variants of the pallet alligment.
'   Pallets are next sorted into groups:
'   Group   I: load size do not exceed pallet nominal size
'   Group  II: load size exceeds pallet nominal size, but is within the allowed maximum oversize
'   Group III: load size exceeds the pallet allowed maximum oversize, but there is only one package on the pallet (reserve, if
'              previous groups will be empty)
'
'   Second sort mechanism find the best pallet in Group I, by mainly comparing the pallet quantity. If the quantity is the same,
'   then it checks another minor parameters.
'   Best pallet from Group I is then compared with each pallet from Group II. If any pallet have bigger package quantity, then it
'   is chosen as better one.
'   If Group I and II were empty, then pallets from Group III are compared between each other to find, the one with the lowest
'   possible oversize
'
'   The chosen pallet variant parameters are given back to the Worksheet.
'---------------------------------------------------------------------------------------------------------------------------------

Option Explicit

' Constants are named ranges of the cells with input values
Public Const BOX_WIDTH_NAME As String = "PackagingWidth"
Public Const BOX_DEPTH_NAME As String = "PackagingDepth"
Public Const BOX_HEIGHT_NAME As String = "PackagingHeight"
Public Const BOX_POSITION_NAME As String = "boxPosition"
Public Const MAX_HEIGHT_NAME As String = "MaxPalletHeight"
Public Const MAX_LAYERS_NAME As String = "maxLayersQtty"
Public Const MAX_OVERLAY_NAME As String = "MaxPalletOverlay"
Public Const MODEL_NAME As String = "ModelName"
Public Const PALLET_DIMENSIONS_NAME As String = "PalletDimensions"
Public Const PALLET_DIM_SEPARATOR As String = "x"

Public Const DATA_INPUT_NAME As String = "DataInput"
Public Const HEADER_NAME As String = "HeaderRow"

' Constants are named ranges of the cells for output values
Public Const LENGTH_QTTY_NAME As String = "swLengthQtty"
Public Const WIDTH_QTTY_NAME As String = "swWidthQtty"
Public Const LAYERS_QTTY_NAME As String = "swLayersQtty"
Public Const PACK_DIM_ON_LENGTH As String = "swPackageLength"
Public Const PACK_DIM_ON_WIDTH As String = "swPackageWidth"
Public Const PACK_DIM_ON_HEIGHT As String = "swPackageHeight"
Public Const ADD_LENGTH_QTTY_NAME As String = "swAddPackLengthQtty"
Public Const ADD_WIDTH_QTTY_NAME As String = "swAddPackWidthQtty"
Public Const ROUND_NAME As String = "roundAlligment"
Public Const UNDERLAY_USE_NAME As String = "UnderlayUse"
Public Const UNDERLAY_THICK_NAME As String = "UnderlayThickness"

' Calculates single product palletization based on the data in single row
Sub CalculateSingle(Target As Range)
    
    Call DefaultValues(Target)

    Dim ws As Worksheet
    Set ws = Target.Parent
    
    Dim row As Long
    row = Target.row
    
    Dim bestPallet As C_Pallet
    Set bestPallet = getBestPallet(ws, row)
    
    Call DataReturn(Target, bestPallet)

End Sub

' Sets default values for crucial cells in the worksheet
Public Sub DefaultValues(Target As Range)
    
    Dim rangeNames() As Variant
    rangeNames = Array(PALLET_DIMENSIONS_NAME, BOX_POSITION_NAME, DATA_INPUT_NAME)
    
    Dim defValues() As Variant
    defValues = Array("1200x800", "Up Up", "Automatic")
    
    Call setDefaultValues(Target, rangeNames, defValues)
       
End Sub

Private Sub setDefaultValues(Target As Range, rangeNames As Variant, defValues As Variant)
    
    Dim ws As Worksheet
    Set ws = Target.Parent
    Dim row As Long
    row = Target.row
    
    Dim i As Integer
    For i = LBound(rangeNames) To UBound(rangeNames)
        If getRngVal(ws, row, rangeNames(i)) = vbNullString Then
            getRng(ws, row, rangeNames(i)).value = defValues(i)
        End If
    Next i
End Sub

Private Function getBestPallet(ws As Worksheet, row As Long) As C_Pallet

    ' Gather all the requirements
    Dim pack As Package
    pack = getPackageDimensions(ws, row)
    
    Dim palreq As palletReq
    palreq = getRequirements(ws, row)
    
    Dim orientReq As orientation
    orientReq = getOrientationReq(ws, row)

    ' Retrieve all possible pallet variants for given orientation
    Dim pallets As Collection
    Set pallets = newPalletsCollection(orientReq, pack, palreq)
    
    'First sort of the pallets
    
    Dim firstSort As Collection
    Set firstSort = New Collection
    Dim firstSortOversize As Collection
    Set firstSortOversize = New Collection
    Dim firstSortReserve As Collection
    Set firstSortReserve = New Collection
    
    Dim pallet As C_Pallet
    For Each pallet In pallets
        With pallet
            
            ' Load size is smaller than pallet dimensions
            If .LengthOversize <= 0 And .WidthOversize <= 0 Then
                firstSort.Add pallet
            
            ' Load size is bigger than pallet dimensions but within allowed oversize
            ElseIf .LengthOversize <= palreq.maxOversize And .WidthOversize <= palreq.maxOversize Then
                firstSortOversize.Add pallet
                
            Else
                ' Special case for oversized pallet with minimum boxes, if any other pallet won't be good
                If .LengthQtty = 1 And .AddLengthQtty = 0 And (.WidthOversize <= palreq.maxOversize _
                    Or (.WidthQtty = 1 And .AddWidthQtty = 0)) Then
                    
                    firstSortReserve.Add pallet
                Else
                    'Do nothing
                End If
            End If
        End With
    Next pallet
    
    'Second sort of the pallets
    
    Dim bestPallet As C_Pallet
    ' If there is a least one pallet wihout oversize
    If 0 < firstSort.Count Then
        Set bestPallet = SecondSort(firstSort)                      ' get best pallet without oversize
        Set bestPallet = SecondSort(firstSortOversize, bestPallet)  ' check, if any oversized pallet will be better
    
    ' if all pallet were oversized, but at least some in the allowed tolerance
    ElseIf firstSort.Count <= 0 And 0 < firstSortOversize.Count Then
        Set bestPallet = SecondSort(firstSortOversize)
    
    ' if all pallets were oversized over the allowed tolerance
    Else
        Set bestPallet = SecondSortReserve(firstSortReserve)
    End If
    
    ' return the best pallet object
    Set getBestPallet = bestPallet
End Function

Private Function SecondSort(firstSort As Collection, Optional bestPallet As C_Pallet = Nothing) As C_Pallet
    
    Dim i As Integer
    Dim currentPallet As C_Pallet
    
    ' Compare each pallet to current best pallet
    For i = 1 To firstSort.Count
        Set currentPallet = firstSort(i)
        
        If bestPallet Is Nothing Then
            Set bestPallet = currentPallet
        Else
            ' The pallet with bigger package quantity is better
            If bestPallet.TotalQtty < currentPallet.TotalQtty Then
                Set bestPallet = currentPallet
            
            ' If the quantity is the same
            ElseIf bestPallet.TotalQtty = currentPallet.TotalQtty Then
                
                ' Comparison by: added pieces, the length-width ratio and the total height
                If (bestPallet.areAddedPieces And Not currentPallet.areAddedPieces) _
                    Or ((currentPallet.LengthToWidthRatio < bestPallet.LengthToWidthRatio) _
                    And (currentPallet.TotalHeight <= bestPallet.TotalHeight)) Then
                        Set bestPallet = currentPallet
                End If
            Else
                'Do nothing
            End If
        End If
    Next i
    
    ' return the best pallet object
    Set SecondSort = bestPallet
    
End Function

Private Function SecondSortReserve(firstSortReserve As Collection) As C_Pallet
    
    Dim currentPallet As C_Pallet
    Dim bestPallet As C_Pallet
    Dim i As Integer
    
    ' Compare each pallet to current best pallet
    For i = 1 To firstSortReserve.Count
        If i = 1 Then
            Set bestPallet = firstSortReserve(i)
        Else
            Set currentPallet = firstSortReserve(i)
            
            ' Best pallet will have the least width oversize
            If (currentPallet.WidthOversize < bestPallet.WidthOversize) Then
                    Set bestPallet = firstSortReserve(i)
            End If
        End If
    Next i
    
    Set SecondSortReserve = bestPallet
End Function


' Retrieve the package dimensions from the worksheet
Private Function getPackageDimensions(ws As Worksheet, row As Long) As Package
  
    With getPackageDimensions
        .Width = getRngVal(ws, row, BOX_WIDTH_NAME)
        .depth = getRngVal(ws, row, BOX_DEPTH_NAME)
        .height = getRngVal(ws, row, BOX_HEIGHT_NAME)
    End With
    
End Function

' Retrieve pallet size requirement from the worksheet
Private Function getRequirements(ws As Worksheet, row As Long) As palletReq
    Dim palletDim As String
    palletDim = getRngVal(ws, row, PALLET_DIMENSIONS_NAME)
    With getRequirements
        ' Retrieve data from cells
        .Length = val(Split(palletDim, PALLET_DIM_SEPARATOR)(0))
        .Width = val(Split(palletDim, PALLET_DIM_SEPARATOR)(1))
        .maxHeight = getRngVal(ws, row, MAX_HEIGHT_NAME)
        .maxOversize = getRngVal(ws, row, MAX_OVERLAY_NAME)
                
        ' If no value was given for the max layers quantity, then there is no limit (negative number)
        If getRngVal(ws, row, MAX_LAYERS_NAME) <> vbNullString Then
            .maxLayers = getRngVal(ws, row, MAX_LAYERS_NAME)
        Else
            .maxLayers = -1
        End If
        
        ' If underlay should be used, then get its thickness. Otherwhise set it to 0
        Select Case getRngVal(ws, row, UNDERLAY_USE_NAME)
            Case "True"
                .underlayThick = getRngVal(ws, row, UNDERLAY_THICK_NAME)
            Case Else
                .underlayThick = 0
        End Select
    End With
    
End Function

' Retrieve the given orientation requirements from the worksheet
Private Function getOrientationReq(ws As Worksheet, row As Long) As orientation
    
    Select Case getRngVal(ws, row, BOX_POSITION_NAME)
        Case "Up Up"
            getOrientationReq = upUp
        
        Case "Front Up"
            getOrientationReq = frontUp
        
        Case "Side Up"
            getOrientationReq = sideUp
        
        ' Special case, where every pallet variant is taken into consideration
        Case "Anyway Up"
            getOrientationReq = anywayUp
    End Select

End Function

' Prepares the collection of the pallet objects, based on given requirements
Private Function newPalletsCollection(orientReq As orientation, pack As Package, palreq As palletReq) As Collection
    
    Dim tempCol As Collection
    Set tempCol = New Collection
    
    If orientReq = anywayUp Then
        Dim i As orientation
        For i = orientation.First To orientation.Last - 1
            Call addNewPalletSet(tempCol, i, pack, palreq)
        Next i
    Else
        Call addNewPalletSet(tempCol, orientReq, pack, palreq)
    End If

    Set newPalletsCollection = tempCol
End Function

' Returns the set of the pallets for given orientation
Private Sub addNewPalletSet(ByRef collect As Collection, orient As orientation, pack As Package, palreq As palletReq)
    Dim i As orientationVar
    For i = orientationVar.First To orientationVar.Last
        collect.Add newPallet(pack, palreq, orient, i)
    Next i
    
End Sub

' Factory method for creating new pallet object
Public Function newPallet(pack As Package, req As palletReq, orient As orientation, orientVar As orientationVar) As C_Pallet
    
    Dim tempPallet As New C_Pallet

    Call tempPallet.newPallet(pack, req, orient, orientVar)
    Set newPallet = tempPallet
    
End Function

' Return the parameters of the best pallet to the worksheet
Private Sub DataReturn(Target As Range, bestPallet As C_Pallet)
    
    ' Prepare array of range names, where values will be stored
    Dim rangeNames() As Variant
    rangeNames = Array(PACK_DIM_ON_LENGTH, PACK_DIM_ON_WIDTH, PACK_DIM_ON_HEIGHT, _
                        LENGTH_QTTY_NAME, WIDTH_QTTY_NAME, LAYERS_QTTY_NAME, _
                        ADD_LENGTH_QTTY_NAME, ADD_WIDTH_QTTY_NAME, ROUND_NAME)
    
    ' Prepare array of pallet parameters - they must correspond to proper array place in rangeNames (!)
    Dim palletParameters() As Variant
    With bestPallet
        palletParameters = Array(.DimOnLength, .DimOnWidth, .DimOnHeight, _
                                    .LengthQtty, .WidthQtty, .LayersQtty, _
                                    .AddLengthQtty, .AddWidthQtty, .RoundStatus)
    End With

    Call updateSheet(Target, rangeNames, palletParameters)

End Sub

Private Sub updateSheet(Target As Range, rangeNames As Variant, palletParameters As Variant)

        Dim ws As Worksheet
        Set ws = Target.Parent
        Dim rowIndex As Long
        rowIndex = Target.row
        
    Call RangeModificationInitiate
        Dim i As Integer
        For i = LBound(rangeNames) To UBound(rangeNames)
            getRng(ws, rowIndex, rangeNames(i)).value = palletParameters(i)
        Next i
    Call RangeModificationTerminate
    
End Sub

' Retrieves value from the named range
Public Function getRngVal(ws As Worksheet, row As Long, rangeName As Variant) As Variant
    
    getRngVal = ws.Cells(row, ws.Range(rangeName).Column).value
    
End Function

' Get the named range object
Public Function getRng(ws As Worksheet, row As Long, rangeName As Variant) As Range
    
    Set getRng = ws.Cells(row, ws.Range(rangeName).Column)
    
End Function
