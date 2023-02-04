VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GoogleSheets"
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
'>>>Class : GoogleSheets
'>>>Ver : 1.0.0.0
'***************************************************************************

Private colSheets As Collection

Public Function Count() As Long
    Count = colSheets.Count
End Function

Public Sub Add(ByRef oGoogleSheet As GoogleSheet)
    colSheets.Add oGoogleSheet
End Sub

Public Function Item(ByVal Index As Variant) As GoogleSheet
    Set Item = colSheets.Item(Index)
End Function

Private Sub Class_Initialize()
        Set colSheets = New Collection
End Sub
