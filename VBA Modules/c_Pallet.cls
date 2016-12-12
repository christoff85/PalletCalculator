VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Pallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------------------------------------------------
'   Class module, that calculates the most efficient single pallet configuration
'
'   Written By Krzysztof Grzeslak 10/06/2015
'
'   Preconditions:
'   *   Requires module s_Pallet, that consists User Defined Types
'
'   Usage:
'   *   New pallet is calculated by creating new object and calling subprocedure newPallet with following parameters:
'       -   Package consisting the package dimensions
'       -   palletReq consisting all the requirements for the pallet dimensions
'       -   orientation, that specifies the package vertical direction in context of the pallet
'       -   orientationVar, that specifies which horizontal package dimensions will correspond to the pallet horizontal dimensions
'
'   *   Returned Object provides pallet parameters through Properties
'   *   Initialized button can be added or removed from the contextMenu
'   *   After any button, the button group can be created. Button counting is automated for the group creation.
'   *   Enum button tags and faces should be adjusted to the user needs.
'---------------------------------------------------------------------------------------------------------------------------------

Option Explicit

' Class private fields / properties
Private m_LengthQtty As Integer
Private m_AddLengthQtty As Integer
Private m_WidthQtty As Integer
Private m_AddWidthQtty As Integer
Private m_LayersQtty As Integer
Private m_TotalQtty As Integer
Private m_LengthOversize As Integer
Private m_WidthOversize As Integer
Private m_DimOnLength As Integer
Private m_DimOnWidth As Integer
Private m_DimOnHeight As Integer
Private m_RoundStatus As Boolean
Private m_loadLength As Integer
Private m_loadWidth As Integer
Private m_totalHeight As Integer

Private Const NOMINAL_PALLET_HEIGHT = 144               ' Standard pallet nominal height
Private Const ROUND_PALLET_MULTIPLIER As Integer = 2    ' Round pallet procedure calculates half of the pallet.

' Public property accessors

' Return the length to width / ratio. If two cases are equivalent in all other comparisons, it is better to have longer
' package dimension per pallet length
Public Property Get LengthToWidthRatio() As Double
    Dim Length As Double
    Length = m_loadLength
    
    Dim Width As Double
    Width = m_loadWidth
    
    LengthToWidthRatio = Length / Width
End Property

' returns whether pallet has additional packages on either dimension
Public Property Get areAddedPieces() As Boolean
    areAddedPieces = False
    
    If m_AddLengthQtty > 0 Or m_AddWidthQtty > 0 Then
        areAddedPieces = True
    End If
    
End Property

' Returns the pallet alligment type
Public Property Get RoundStatus() As String
    Const RND_STATUS_TRUE As String = "True"
    Const RND_STATUS_FALSE As String = "False"
    
    Select Case m_RoundStatus
        Case True
            RoundStatus = RND_STATUS_TRUE
        Case False
            RoundStatus = RND_STATUS_FALSE
    End Select
    
End Property

Public Property Get TotalHeight() As Integer
    TotalHeight = m_totalHeight
End Property

Public Property Get LengthQtty() As Integer
    LengthQtty = m_LengthQtty
End Property

Public Property Get AddLengthQtty() As Integer
    AddLengthQtty = m_AddLengthQtty
End Property

Public Property Get WidthQtty() As Integer
    WidthQtty = m_WidthQtty
End Property

Public Property Get AddWidthQtty() As Integer
    AddWidthQtty = m_AddWidthQtty
End Property

Public Property Get LayersQtty() As Integer
    LayersQtty = m_LayersQtty
End Property

Public Property Get TotalQtty() As Integer
    TotalQtty = m_TotalQtty
End Property

Public Property Get LengthOversize() As Integer
    LengthOversize = m_LengthOversize
End Property

Public Property Get WidthOversize() As Integer
    WidthOversize = m_WidthOversize
End Property

Public Property Get DimOnLength() As Integer
    DimOnLength = m_DimOnLength
End Property

Public Property Get DimOnWidth() As Integer
    DimOnWidth = m_DimOnWidth
End Property

Public Property Get DimOnHeight() As Integer
    DimOnHeight = m_DimOnHeight
End Property

' Calculate, how many packages fits per size into the round alligment
Private Function countRoundPieces(palletDim As Integer, packageDim As Integer, secondaryPackageDim As Integer, Optional maxOversize As Integer = 0, _
    Optional lengthCalculation As Boolean = False, Optional minQtty As Integer = 1) As Integer

    countRoundPieces = Application.RoundUp(((palletDim - secondaryPackageDim) / packageDim), 0)
        
    Do While countRoundPieces > minQtty And (countRoundPieces * packageDim) > palletDim + maxOversize - secondaryPackageDim
        countRoundPieces = countRoundPieces - 1
    Loop
    
    If countRoundPieces < 0 Then countRoundPieces = 1
    
End Function

' Routing subprocedure for calculating new pallet
Public Sub newPallet(pack As Package, req As palletReq, orient As orientation, orientVar As orientationVar)
    If orientVar = var1 Or orientVar = var2 Then
        Call newStdPallet(pack, req, orient, orientVar)
    Else
        Call newRndPallet(pack, req, orient, orientVar)
    End If
End Sub

' Calculates the best possible round pallet alligment with given package dimensions and box orientation
Private Sub newRndPallet(pack As Package, req As palletReq, orient As orientation, orientVar As orientationVar)
    
    Dim packSize As PalletPackage
    packSize = determineSizes(pack, orient, orientVar)
    
    ' Determine package count per pallet length
    m_LengthQtty = countRoundPieces(req.Length, packSize.perLength, packSize.perWidth, req.maxOversize)
    ' determine additional 90deg rotated package count at the end of the pallet length
    m_AddLengthQtty = countRoundPieces(req.Width, packSize.perLength, packSize.perWidth, req.maxOversize)
    ' Determine number of layers
    m_LayersQtty = countPieces(req.maxHeight, packSize.perHeight, , NOMINAL_PALLET_HEIGHT + req.underlayThick, req.maxLayers)
    
    ' Ensure Load Size is not biger than pallet + maximum allowed oversize
    Do While (m_LengthQtty * packSize.perLength) > (req.Length + req.maxOversize) / 2 And packSize.perWidth > (req.Width + req.maxOversize) / 2
        m_LengthQtty = m_LengthQtty - 1
    Loop
    
    ' Fields not relevant to round alligment
    m_WidthQtty = 0
    m_AddWidthQtty = 0
    
    ' Calculate total quantity of packages on the pallet
    m_TotalQtty = (m_LengthQtty + m_AddLengthQtty) * m_LayersQtty * ROUND_PALLET_MULTIPLIER
    
    ' Calculate size of the load
    m_loadLength = countLoadSize(packSize.perLength, m_LengthQtty, packSize.perWidth)
    m_loadWidth = countLoadSize(packSize.perLength, m_AddLengthQtty, packSize.perWidth)

    ' Calculate total pallet height
    m_totalHeight = NOMINAL_PALLET_HEIGHT + req.underlayThick + m_LayersQtty * packSize.perHeight

    ' Calculate pallet oversizes
    m_LengthOversize = m_loadLength - req.Length
    m_WidthOversize = m_loadWidth - req.Width

    ' assign package size in the context of the pallet
    m_DimOnLength = packSize.perLength
    m_DimOnWidth = packSize.perWidth
    m_DimOnHeight = packSize.perHeight
    
    ' set the round pallet alligment parameter
    m_RoundStatus = True
    
End Sub

Private Sub newStdPallet(pack As Package, req As palletReq, orient As orientation, orientVar As orientationVar)
    
    Dim packSize As PalletPackage
    packSize = determineSizes(pack, orient, orientVar)
    
    ' Determine package count per pallet length, width and numbers of layers
    m_LengthQtty = countPieces(req.Length, packSize.perLength, req.maxOversize)
    m_WidthQtty = countPieces(req.Width, packSize.perWidth, req.maxOversize)
    m_LayersQtty = countPieces(req.maxHeight, packSize.perHeight, , NOMINAL_PALLET_HEIGHT + req.underlayThick, req.maxLayers)
    
    ' Add 90deg rotated package at the end of the pallet length, if there is still space
    If (req.Length - m_LengthQtty * packSize.perLength) + req.maxOversize >= packSize.perWidth Then
        m_AddLengthQtty = countPieces(req.Width, packSize.perLength, req.maxOversize, , , 0)
    End If
    
    ' Add 90deg rotated package at the end of the pallet width, if there is still space
    If (req.Width - m_WidthQtty * packSize.perWidth) + req.maxOversize >= packSize.perLength Then
        m_AddWidthQtty = countPieces(req.Length, packSize.perWidth, req.maxOversize, , , 0)
    End If
    
    ' Calculate total quantity of packages on the pallet
    m_TotalQtty = countQtty(m_LengthQtty, m_WidthQtty, m_AddLengthQtty, m_AddWidthQtty, m_LayersQtty)
    
    ' Calculate size of the load:
    ' if there are additional packages at the end of pallet length
    If m_AddLengthQtty >= 1 And m_AddWidthQtty = 0 Then
        m_loadLength = countLoadSize(packSize.perLength, m_LengthQtty, packSize.perWidth)
        m_loadWidth = countLoadSize2(packSize.perWidth, m_WidthQtty, packSize.perLength, m_AddLengthQtty)
        
    ' if there are additional packages at the end of pallet width
    ElseIf m_AddLengthQtty = 0 And m_AddWidthQtty >= 1 Then
        m_loadLength = countLoadSize2(packSize.perLength, m_LengthQtty, packSize.perWidth, m_AddWidthQtty)
        m_loadWidth = countLoadSize(packSize.perWidth, m_WidthQtty, packSize.perLength)
        
    ' if there are additional packages at the end of pallet length and also pallet width
    ElseIf m_AddLengthQtty >= 1 And m_AddWidthQtty >= 1 Then
        m_loadLength = countLoadSize2(packSize.perLength, m_LengthQtty, packSize.perWidth, m_AddWidthQtty, True)
        m_loadWidth = countLoadSize2(packSize.perWidth, m_WidthQtty, packSize.perLength, m_AddLengthQtty, True)
        
    ' if there are no addtional packages on the pallet
    Else
        m_loadLength = countLoadSize(packSize.perLength, m_LengthQtty)
        m_loadWidth = countLoadSize(packSize.perWidth, m_WidthQtty)
    End If
    
    ' Calculate total pallet height
    m_totalHeight = NOMINAL_PALLET_HEIGHT + req.underlayThick + m_LayersQtty * packSize.perHeight
    
    ' Calculate pallet oversizes
    m_LengthOversize = m_loadLength - req.Length
    m_WidthOversize = m_loadWidth - req.Width
    
    ' assign package size in the context of the pallet
    m_DimOnLength = packSize.perLength
    m_DimOnWidth = packSize.perWidth
    m_DimOnHeight = packSize.perHeight
    
    ' set the round pallet alligment parameter
    m_RoundStatus = False
End Sub

' Based on the given package orientation, determine which package size will corresond to each pallet dimension
Private Function determineSizes(pack As Package, orient As orientation, orientVar As orientationVar) As PalletPackage

    With determineSizes
        Select Case orient
            
            Case upUp
                If orientVar = var1 Then
                    .perLength = pack.Width
                    .perWidth = pack.depth
                Else
                    .perLength = pack.depth
                    .perWidth = pack.Width
                End If
                
                .perHeight = pack.height
                
            Case frontUp
                If orientVar = var1 Then
                    .perLength = pack.Width
                    .perWidth = pack.height
                Else
                    .perLength = pack.height
                    .perWidth = pack.Width
                End If
                
                .perHeight = pack.depth
                
            Case sideUp
                If orientVar = var1 Then
                    .perLength = pack.height
                    .perWidth = pack.depth
                Else
                    .perLength = pack.depth
                    .perWidth = pack.height
                End If
                
                .perHeight = pack.Width
        End Select
    End With

End Function

' Calculate how many box pieces will fit per pallet size
Private Function countPieces(palletDim As Integer, packageDim As Integer, Optional maxOversize As Integer = 0, _
    Optional heightModifier As Integer = 0, Optional maxLayers As Integer = -1, Optional minQtty As Integer = 1) As Integer

    countPieces = Application.RoundUp(((palletDim - heightModifier) / packageDim), 0)
    
    If maxLayers > 0 And countPieces > maxLayers Then
        countPieces = maxLayers
    End If
    
    Do While countPieces > minQtty And (countPieces * packageDim) + heightModifier > palletDim + maxOversize
        countPieces = countPieces - 1
    Loop

End Function

' Calcuate total boxes quantities on the pallet (standard alligment)
Private Function countQtty(piecesLength As Integer, piecesWidth As Integer, piecesLengthAdd As Integer, piecesWidthAdd As Integer, m_LayersQtty As Integer) As Integer

    countQtty = ((piecesLength * piecesWidth) + piecesLengthAdd + piecesWidthAdd) * m_LayersQtty

End Function

' Calculate load size per given pallet dimension (without additional packages)
Private Function countLoadSize(packageDim As Integer, countedPieces As Integer, Optional packageDimAdd As Integer = 0) As Integer

    countLoadSize = countedPieces * packageDim + packageDimAdd

End Function

' Calculate load size per given pallet dimension (with additional packages)
Private Function countLoadSize2(packageDim As Integer, countedPieces As Integer, packageDimAdd As Integer, countPiecesAdd As Integer, _
    Optional extendedAddcheck As Boolean = False)
    
    Dim additionalDim As Integer
    
    If extendedAddcheck Then
        additionalDim = packageDimAdd
    Else
        additionalDim = 0
    End If
        
    If countedPieces * packageDim + additionalDim >= countPiecesAdd * packageDimAdd Then
        countLoadSize2 = countLoadSize(packageDim, countedPieces, additionalDim)
    Else
        countLoadSize2 = countLoadSize(packageDimAdd, countPiecesAdd)
    End If

End Function
