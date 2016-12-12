Attribute VB_Name = "s_Pallet"
'---------------------------------------------------------------------------------------------------------------------------------
'   Consists User Defined Types and Enumerations required by c_Pallet class module
'
'   Written By Krzysztof Grzeslak 10/06/2015
'
'---------------------------------------------------------------------------------------------------------------------------------

' Pallet requirements input
Public Type palletReq
    Length As Integer
    Width As Integer
    maxHeight As Integer
    maxLayers As Integer
    maxOversize As Integer
    underlayThick As Integer
End Type

' Single package dimensions
Public Type Package
    Width As Integer
    depth As Integer
    height As Integer
End Type

' Stores assigment between package and pallet dimensions
Public Type PalletPackage
    perLength As Integer
    perWidth As Integer
    perHeight As Integer
End Type

' Stores single pallet paramaters
Public Type palletParameters
    pcsLength As Integer
    pcsLengthAdd As Integer
    pcsWidth As Integer
    pcsWidthAdd As Integer
    layers As Integer
    quantity As Integer
    loadSizeLength As Integer
    loadSizeWidth As Integer
    loadOversizeLength As Integer
    loadOversizeWidth As Integer
    orient As orientation
End Type

' Correspond to user input, how packages should be stored on the pallet
Public Enum orientation
    First
    upUp = First
    frontUp
    sideUp
    anywayUp ' means all of the above are allowed
    Last = anywayUp
End Enum

' Each single orientation is calculated in four variants
Public Enum orientationVar
    First
    var1 = First
    var2
    round1
    round2
    Last = round2
End Enum
