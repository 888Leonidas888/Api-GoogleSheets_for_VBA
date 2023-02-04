VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GoogleSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**************************************************************************
'>>>Dev : Jhony Escriba Bautista
'>>>Email : jhonny14_1@hotmail.com
'>>>Release 2023
'>>>Proyect : Client Api GoogleSheets for VBA
'>>>Class : GoogleSheet
'>>>Ver : 1.0.0.0
'***************************************************************************

Private Type tGs
    SheetId As Long
    Title As String
    Index As Long
    SheetType As String
    RowCount As Integer
    ColumnCount As Integer
End Type

Private This As tGs

Public Property Get SheetId() As Long
    SheetId = This.SheetId
End Property

Public Property Let SheetId(ByVal lSheetId As Long)
    This.SheetId = lSheetId
End Property

Public Property Get Title() As String
    Title = This.Title
End Property

Public Property Let Title(ByVal stitle As String)
    This.Title = stitle
End Property

Public Property Get Index() As Long
    Index = This.Index
End Property

Public Property Let Index(ByVal lIndex As Long)
    This.Index = lIndex
End Property

Public Property Get SheetType() As String
    SheetType = This.SheetType
End Property

Public Property Let SheetType(ByVal sSheetType As String)
    This.SheetType = sSheetType
End Property


Rem let and get,row and column
Public Function gridProperties() As gridProperties
    
    Dim gp As New gridProperties
    
    gp.ColumnCount = This.ColumnCount
    gp.RowCount = This.RowCount
    
    Set gridProperties = gp
    
    Set gp = Nothing
End Function

Public Property Get RowCount() As Long
    RowCount = This.RowCount
End Property

Public Property Let RowCount(ByVal row As Long)
    This.RowCount = row
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = This.ColumnCount
End Property

Public Property Let ColumnCount(ByVal column As Long)
    This.ColumnCount = column
End Property
