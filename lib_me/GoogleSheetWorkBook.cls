VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GoogleSheetWorkBook"
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
'>>>Class : GoogleSheetWorkBook
'>>>Ver : 1.0.0.0
'***************************************************************************

Private Type tGsw
    SpreadsheetId As String
    SpreadSheetUrl As String
    respSheetsAfter As String
End Type

Private This As tGsw

Public Sub Create(ByVal responseJSON As String)
    
    Dim dicSheets As Dictionary
    Set dicSheets = JsonConverter.ParseJson(responseJSON)
    
    This.SpreadSheetUrl = dicSheets("spreadsheetUrl")
    This.SpreadsheetId = dicSheets("spreadsheetId")
    
    This.respSheetsAfter = responseJSON
    
    Set dicSheets = Nothing
    
End Sub

Public Property Get SpreadsheetId() As String
    SpreadsheetId = This.SpreadsheetId
End Property

Public Function Properties(ByVal propertie As String) As String
    
    Dim dicSheets As Dictionary
    
    Set dicSheets = JsonConverter.ParseJson(This.respSheetsAfter)
    
    Properties = dicSheets("properties")(propertie)
    
    Set dicSheets = Nothing
End Function

Public Function Sheets(ByVal Properties As String, Optional Index = 1) As String

    Dim dicSheets As Dictionary

    Set dicSheets = JsonConverter.ParseJson(This.respSheetsAfter)

    Sheets = dicSheets("sheets")(Index)("properties")(Properties)

    Dim oSheet As New GoogleSheet
    Set dicSheets = Nothing

End Function

Public Property Get SpreadSheetUrl() As String
    SpreadSheetUrl = This.SpreadSheetUrl
End Property
Public Function GoogleSheets() As GoogleSheets
    
    Dim gss As New GoogleSheets
    Dim gs As GoogleSheet
    Dim dicJson As Dictionary
    Dim dicSheets As Dictionary
    Dim dicPorperties As Dictionary
    Dim dicGrid As Dictionary
    Dim i As Integer
    
    Set dicJson = JsonConverter.ParseJson(This.respSheetsAfter)
    
    For i = 1 To dicJson("sheets").Count
    
         Set dicSheets = dicJson("sheets")(i)
         Set dicPorperties = dicSheets("properties")
         Set gs = New GoogleSheet
         
            With gs
                .SheetId = dicPorperties("sheetId")
                .Title = dicPorperties("title")
                .SheetType = dicPorperties("sheetType")
                .Index = dicPorperties("index")
                
                Set dicGrid = dicPorperties("gridProperties")
                
                .RowCount = dicGrid("rowCount")
                .ColumnCount = dicGrid("columnCount")
       
            End With
    
            gss.Add gs
    Next i
     
     Set GoogleSheets = gss
     
     Set gss = Nothing
     Set gs = Nothing
     Set dicJson = Nothing
     Set dicSheets = Nothing
     Set dicPorperties = Nothing
     Set dicGrid = Nothing
     
    
End Function



