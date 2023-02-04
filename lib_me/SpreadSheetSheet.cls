VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpreadSheetSheet"
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
'>>>Class : SpreadSheetSheet
'>>>Ver : 1.0.0.0
'***************************************************************************

Private WithEvents HTTP As HTTPRequestService
Attribute HTTP.VB_VarHelpID = -1
Private Const SERVICE_END_POINT As String = "https://sheets.googleapis.com"

Private Type tSss
   yourApiKey As String
   yourAccessToken As String
   responseText As String
   status As String
End Type

Private This As tSss

Public Sub ConnectionService(ByVal oFlowOauth As FlowOauth)
    
    With oFlowOauth
        This.yourAccessToken = .GetTokenAccess()
        This.yourApiKey = .GetApiKey()
    End With
    
End Sub

Public Property Get Operation() As Integer

    Operation = This.status
    
End Property

Public Function batchUpdate(ByVal spreadsheetsId As String, ByRef request As Collection) As String
    
    Dim body As String
    Dim requets_requets As String
    Dim url As String
    Dim parametersPath As String
    Dim i As Integer
    
    Rem recuperamos cada uno de los REQUESTS para GoogleSheet
    For i = 1 To request.Count
        If i = 1 Then
            requets_requets = request.Item(i)
        Else
            requets_requets = requets_requets & "," & request.Item(i)
        End If
    Next i
    
    body = formatString( _
            Replace( _
            "{'includeSpreadsheetInResponse':'false','requests':[{0}]}", _
            "'", """"), _
            requets_requets)
            
    parametersPath = formatString("/v4/spreadsheets/{0}:batchUpdate?key={1}", spreadsheetsId, This.yourApiKey)
    url = SERVICE_END_POINT & parametersPath
    
    This.responseText = HTTPGoogleSheets(HT_POST, url, body)
    batchUpdate = This.responseText
    
End Function

Public Function CreateWorkBook(Optional nameWorkBook, Optional nameSheet, Optional bodyJson) As String
    
    Dim url As String
    Dim parametersPath As String
    Dim body As String
    
    If IsMissing(nameWorkBook) Then nameWorkBook = "WorkBook_for_VBA"
    If IsMissing(nameSheet) Then nameSheet = "Google_sheet"
        
    parametersPath = formatString("/v4/spreadsheets?key={0}", This.yourApiKey)
    url = SERVICE_END_POINT + parametersPath
    
    If IsMissing(bodyJson) Then
        body = formatString( _
                Replace( _
                "{'properties':{'title':'{0}'},'sheets':[{'properties':{'title':'{1}'}}]}", "'", """"), _
                nameWorkBook, nameSheet)
    Else
        body = bodyJson
    End If
    
    This.responseText = HTTPGoogleSheets(HT_POST, url, body)
    
    CreateWorkBook = This.responseText
    
End Function

Public Function FindReplace(ByVal find As String, _
                            ByVal replacement As String, _
                            ByRef tUnionFieldScope As unionFieldScope, _
                            Optional matchCase = False, _
                            Optional matchEntireCell = False, _
                            Optional searchByRegex = False, _
                            Optional includeFormulas = False) As String
                        
    
    Dim request As String
    Dim scope_scope As String
    
    scope_scope = convertTypeToStr(tUnionFieldScope)
    
    request = formatString( _
              Replace( _
              "{'findReplace':{'find':'{0}','replacement':'{1}','matchCase':'{2}','matchEntireCell':'{3}','searchByRegex':'{4}','includeFormulas':'{5}',{6}}}", _
              "'", """"), _
                find, _
                replacement, _
                boolToStr(matchCase), _
                boolToStr(matchEntireCell), _
                boolToStr(searchByRegex), _
                boolToStr(includeFormulas), _
                scope_scope)
    
    FindReplace = request
    
End Function

Public Function DeleteRange(ByRef tUnionFieldScope As unionFieldScope, ByVal shiftDimension As String) As String
    
    Dim requests As String
    Dim scope_scope As String
    
    scope_scope = convertTypeToStr(tUnionFieldScope)
    requests = formatString( _
                Replace("{'deleteRange':{0},'shiftDimension':'{1}'}", "'", """") _
                , scope_scope, shiftDimension)

    DeleteRange = requests

End Function

Public Function RecoveryById(ByVal spreadsheetsId As String) As String
    
    Dim parametersPath As String
    Dim url As String

    parametersPath = formatString("/v4/spreadsheets/{0}?key={1}", spreadsheetsId, This.yourApiKey)
    url = SERVICE_END_POINT + parametersPath
    
    This.responseText = HTTPGoogleSheets(HT_GET, url)
    
    RecoveryById = This.responseText
    
End Function

Public Function CopyTo(ByVal spreadsheetsId As String, _
                       ByVal SheetId As Integer, _
                       ByVal destinationSpreadsheetId As String) As String
    
    Dim url As String
    Dim parametersPath As String
    Dim body As String
    
    parametersPath = formatString("/v4/spreadsheets/{0}/sheets/{1}:copyTo?&key={2}", _
                                    spreadsheetsId, _
                                    SheetId, _
                                    This.yourApiKey)
                                    
    url = SERVICE_END_POINT + parametersPath
    body = formatString(Replace("{'destinationSpreadsheetId':'{0}'}", "'", """"), destinationSpreadsheetId)
    This.responseText = HTTPGoogleSheets(HT_POST, url, body)
    
    CopyTo = This.responseText
        
End Function

Public Function DetailsError() As String
    Dim dicError As Dictionary
    Dim messageError As String
    Dim k As Variant
    
    Set dicError = JsonConverter.ParseJson(This.responseText)
    
    If dicError.Exists("error") Then

        For Each k In dicError("error").Keys
            messageError = messageError & k & " : " & dicError("error")(k) & vbCrLf
        Next k
        
        DetailsError = "ERROR" & vbCrLf & messageError

    End If
    
    Set dicError = Nothing
    Set k = Nothing
End Function

'@Function private
'------------------
Private Function HTTPGoogleSheets(ByVal methods As String, ByVal url As String, Optional body = Empty, Optional dicHeaders As Dictionary) As String

    Dim dicHeaders1 As New Dictionary
    Dim k As Dictionary
    Dim v As String
    
    Set HTTP = New HTTPRequestService
    
    With dicHeaders1
        .Add "Authorization", "Bearer " & This.yourAccessToken
        .Add "Accept", "application/json"
        .Add "Content-Type", "application/json"
    End With
    
    If Not dicHeaders Is Nothing Then
        For Each k In dicHeaders.Keys
            If Not dicHeaders1.Exists(k) Then
                v = dicHeaders(k)
                dicHeaders1.Add k, v
            End If
        Next k
    End If
    
    This.responseText = HTTP.HTTPrequest(methods, url, body, dicHeaders1)
    
    HTTPGoogleSheets = This.responseText
    
    Set k = Nothing
    Set dicHeaders1 = Nothing
    Set HTTP = Nothing
    
End Function

Private Sub HTTP_forRequestHTTP(ByVal status As Integer)
    This.status = status
End Sub

Private Function convertTypeToStr(ByRef tUnionFieldScope As unionFieldScope) As String
    
    Dim fieldScope As String
    
    With tUnionFieldScope
        If Not IsEmpty(.rng.SheetId) Then
            fieldScope = formatString( _
                         Replace( _
                         "'range':{'sheetId':{0},'startRowIndex':{1},'endRowIndex':{2},'startColumnIndex':{3},'endColumnIndex':{4}}", "'", """"), _
                         .rng.SheetId, _
                         .rng.startRowIndex, _
                         .rng.endRowIndex, _
                         .rng.startColumnIndex, _
                         .rng.endColumnIndex)
                                       
        ElseIf Not IsEmpty(.SheetId) Then
            fieldScope = formatString(Replace("'sheetId':{0}", "'", """"), .SheetId)
        ElseIf Not IsEmpty(.allSheets) Then
            fieldScope = formatString(Replace("'allSheets':'{0}'", "'", """"), boolToStr(.allSheets))
        End If
    End With
    
    convertTypeToStr = fieldScope
    
End Function

