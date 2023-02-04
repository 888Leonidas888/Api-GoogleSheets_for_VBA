VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gridProperties"
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
'>>>Class : GridProperties
'>>>Ver : 1.0.0.0
'***************************************************************************

Private Type tGp
    RowCount As Long
    ColumnCount As Long
End Type

Private This As tGp

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
