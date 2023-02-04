Attribute VB_Name = "ProcessResponse"
Option Explicit

Public Function AppendValue(ByVal responseJSON As String) As String
   
    Dim dicJson As Dictionary
    Dim SpreadsheetId$, tableRange$
    Dim spreadsheetId_Id$, updatedRange$, updatedRows$
    Dim updatedColumns$, updatedCells
    Dim message$
    
    
    Set dicJson = JsonConverter.ParseJson(responseJSON)
    
    SpreadsheetId = dicJson("spreadsheetId")
    tableRange = dicJson("tableRange")
    
    spreadsheetId_Id = dicJson("updates")("spreadsheetId")
    updatedRange = dicJson("updates")("updatedRange")
    updatedRows = dicJson("updates")("updatedRows")
    updatedColumns = dicJson("updates")("updatedColumns")
    updatedCells = dicJson("updates")("updatedCells")
    
    message = "ID                    : " & SpreadsheetId & vbCrLf & _
            "Rango tabla           : " & tableRange & vbCrLf & vbCrLf & _
            "Actualizaciones " & vbCrLf & vbCrLf & _
            "Id hoja de cálculo    : " & spreadsheetId_Id & vbCrLf & _
            "Rango actualizado     : " & updatedRange & vbCrLf & _
            "Filas actualizadas    : " & updatedRows & vbCrLf & _
            "Columnas actualizadas : " & updatedColumns & vbCrLf & _
            "Celdas actualizadas   : " & updatedCells
    
    AppendValue = message
    
End Function
Rem Este módulo le permite leer las respuestas del Api de Google Sheets, que son devueltas en formato json
Public Function UpdateValue(ByVal responseJSON As String) As String
    Rem esta función de formato a json a texto plano  lo retornará
    Rem Use esta función solo si el argumento de la función SpreadSheetValue.Update
    Dim dicJson As Dictionary
    Dim SpreadsheetId$, updatedRange$
    Dim updatedRows$, updatedColumns$, updatedCells$
    Dim message As String
    
    
    Set dicJson = JsonConverter.ParseJson(responseJSON)
    
    SpreadsheetId = dicJson("spreadsheetId")
    updatedRange = dicJson("updatedRange")
    updatedRows = dicJson("updatedRows")
    updatedColumns = dicJson("updatedColumns")
    updatedCells = dicJson("updatedCells")
    
    message = "ID                    : " & SpreadsheetId & vbCrLf & _
            "Rango actualizado     : " & updatedRange & vbCrLf & _
            "Filas actualizadas    : " & updatedRows & vbCrLf & _
            "Columnas actualizadas : " & updatedColumns & vbCrLf & _
            "Celdas actualizadas   : " & updatedCells
    
    UpdateValue = message
    
End Function

Public Function GetValue(ByVal responseJSON As String) As Variant()


    Dim dicJson As Dictionary
    Dim rowSheet As Long
    Dim columnHeaderSheet As Long
    Dim columnIndexSheet As Long
    Dim arrSheet() As Variant
    Dim i As Integer
    Dim o As Integer
    
'    On Error Resume Next
    Set dicJson = JsonConverter.ParseJson(responseJSON)
    
    rowSheet = dicJson("values").Count
    columnHeaderSheet = dicJson("values")(1).Count
    
    ReDim Preserve arrSheet(0 To rowSheet - 1, 0 To columnHeaderSheet - 1)

    For i = 1 To rowSheet
        columnIndexSheet = dicJson("values")(i).Count
        For o = 1 To columnIndexSheet
            arrSheet(i - 1, o - 1) = dicJson("values")(i)(o)
        Next o
    Next i
    
    GetValue = arrSheet
    
End Function
Public Function ClearValues(ByVal responseJSON As String) As String
    
    Dim dicJson As Dictionary
    Dim SpreadsheetId As String
    Dim clearedRange As String
    Dim message As String
    
    Set dicJson = JsonConverter.ParseJson(responseJSON)
    
    SpreadsheetId = dicJson("spreadsheetId")
    clearedRange = dicJson("clearedRange")
    
    message = "ID              : " + SpreadsheetId + vbCrLf + _
              "Rango eliminado : " + clearedRange
 
    ClearValues = message
End Function
Public Function batchUpdate(ByVal responseJSON As String) As String

    Dim dicJson As Dictionary
    Dim SpreadsheetId As String
    
    Set dicJson = JsonConverter.ParseJson(responseJSON)
    SpreadsheetId = dicJson("spreadsheetId")
    
    batchUpdate = SpreadsheetId
    
End Function

Rem Ejemplos
'Function ReadValuesSheets(ByVal text As String) As String()
'
'    Dim jsonCollection As Dictionary
'    Dim arrValueSheets() As String
'    Dim endArr As Integer
'
'    Set jsonCollection = JsonConverter.ParseJson(text)
'    endArr = jsonCollection("values").Count - 1
'
'    ReDim arrValueSheets(endArr)
'
''    For Each Item In jsonCollection("values")(1)
''        Debug.Print Item
''    Next Item
''
''    Debug.Print vbCrLf
''
''    For Each Item In jsonCollection("values")
''        Debug.Print Item(2)
''    Next Item
'
'    For i = 1 To endArr + 1
'        For Each Item In jsonCollection("values")(i)
'
'        Next Item
'    Next i
'
'End Function

