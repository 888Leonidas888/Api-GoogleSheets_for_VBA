VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SpreadSheetValue"
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
'>>>Class : SpreadSheetValue
'>>>Ver : 1.0.0.0
'***************************************************************************

Private WithEvents HTTP As HTTPRequestService
Attribute HTTP.VB_VarHelpID = -1
Private Const SERVICE_END_POINT As String = "https://sheets.googleapis.com"

Private Type tSsv
    status As Integer
    responseTxt As String
    yourApiKey As String
    yourAccessToken As String
End Type

Private This As tSsv

Public Sub ConnectionService(ByRef oFlowOauth As FlowOauth)

    With This
        .yourAccessToken = oFlowOauth.GetTokenAccess()
        .yourApiKey = oFlowOauth.GetApiKey()
    End With
    
End Sub

Public Function GetValue(ByVal spreadsheetsId As String, _
                        ByVal rng As String, _
                        Optional majordimension = "ROWS", _
                        Optional valueRenderOption = "FORMATTED_VALUE", _
                        Optional dateTimeRenderOption = "SERIAL_NUMBER") As String
            
    Dim url As String
    Dim parametersPath As String
    Dim parametersQuery As String
    
    rng = codificationUrl(rng)
    parametersPath = formatString("/v4/spreadsheets/{0}/values/{1}", spreadsheetsId, rng)
    parametersQuery = formatString("?dateTimeRenderOption={0}&majorDimension={1}&valueRenderOption={2}&key={3}", _
                                    dateTimeRenderOption, _
                                    majordimension, _
                                    valueRenderOption, _
                                    This.yourApiKey)
    
    url = SERVICE_END_POINT + parametersPath + parametersQuery

    This.responseTxt = HTTPGoogleSheets(HT_GET, url)

    GetValue = This.responseTxt
    
End Function
Public Function Update(ByVal spreadsheetsId As String, _
                      ByVal rng As String, _
                      ByRef updateRangeValue As Collection, _
                      Optional majordimension = "rows", _
                      Optional valueInputOption = "user_entered", _
                      Optional includeValuesInResponse = "false", _
                      Optional responseValueRenderOption = "formatted_value", _
                      Optional responseDateTimeRenderOption = "serial_number") As String
    
    Dim body As String
    Dim url As String
    Dim parametersPath As String
    Dim parametersQuery As String

    Rem convertir updateRangeValue a string para enviarlo en el body
    body = createBodyJson(majordimension, rng, updateRangeValue)

    rng = codificationUrl(rng)
    parametersPath = formatString("/v4/spreadsheets/{0}/values/{1}", spreadsheetsId, rng)
    parametersQuery = formatString( _
                    "?valueInputOption={0}&includeValuesInResponse={1}&responseValueRenderOption={2}&responseDateTimeRenderOption={3}&key={4}", _
                    valueInputOption, _
                    includeValuesInResponse, _
                    responseValueRenderOption, _
                    responseDateTimeRenderOption, _
                    This.yourApiKey)


    url = SERVICE_END_POINT + parametersPath + parametersQuery
    
    This.responseTxt = HTTPGoogleSheets(HT_PUT, url, body)
    
    Update = This.responseTxt
    
End Function
Public Function Append(ByVal spreadsheetsId As String, _
                      ByVal rng As String, _
                      ByRef appendRangeValue As Collection, _
                      Optional majordimension = "rows", _
                      Optional valueInputOption = "user_entered", _
                      Optional insertDataOption = "insert_rows", _
                      Optional includeValuesInResponse = "false", _
                      Optional responseValueRenderOption = "formatted_value", _
                      Optional responseDateTimeRenderOption = "serial_number") As String

    Dim body As String
    Dim url As String
    Dim parametersPath As String
    Dim parametersQuery As String
    
    Rem convertir appendRangeValue a string para enviarlo en el body
    body = createBodyJson(majordimension, rng, appendRangeValue)

    rng = codificationUrl(rng)
    parametersPath = formatString("/v4/spreadsheets/{0}/values/{1}:append", spreadsheetsId, rng)
    parametersQuery = formatString( _
    "?valueInputOption={0}&includeValuesInResponse={1}&responseValueRenderOption={2}&insertDataOption={3}&responseDateTimeRenderOption={4}&key={5}", _
            valueInputOption, _
            includeValuesInResponse, _
            responseValueRenderOption, _
            insertDataOption, _
            responseDateTimeRenderOption, _
            This.yourApiKey)

    url = SERVICE_END_POINT + parametersPath + parametersQuery
     
    This.responseTxt = HTTPGoogleSheets(HT_POST, url, body)
    
    Append = This.responseTxt
    
End Function

Public Function Clear(ByVal spreadsheetsId As String, ByVal rng As String) As String
    
    Dim url As String
    Dim parametersPath As String
    
    rng = codificationUrl(rng)
    parametersPath = formatString("/v4/spreadsheets/{0}/values/{1}:clear?key={2}", spreadsheetsId, rng, This.yourApiKey)
    url = SERVICE_END_POINT + parametersPath
        
    This.responseTxt = HTTPGoogleSheets(HT_POST, url)
    
    Clear = This.responseTxt

End Function

Public Property Get Operation() As Integer

    Operation = This.status
    
End Property

Public Function DetailsError() As String
    Dim dicError As Dictionary
    Dim messageError As String
    Dim k As Variant
    
    Set dicError = JsonConverter.ParseJson(This.responseTxt)
    
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
Private Function HTTPGoogleSheets(ByVal methods As String, _
                                  ByVal url As String, _
                                  Optional body = Empty, _
                                  Optional dicHeaders As Dictionary) As String

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
    
    This.responseTxt = HTTP.HTTPrequest(methods, url, body, dicHeaders1)
    
    HTTPGoogleSheets = This.responseTxt
    
    Set k = Nothing
    Set dicHeaders1 = Nothing
    Set HTTP = Nothing
    
End Function

Private Sub HTTP_forRequestHTTP(ByVal status As Integer)

    This.status = status
    
End Sub

Private Function converterCollectionToList(ByRef rangeCollection As Collection) As String
    
    Dim converterString As String
    Dim valuesList As String
    Dim i As Integer
    
    For i = 1 To rangeCollection.Count
        converterString = rangeCollection(i)
'        converterString = "[""" + Replace(rangeCollection.Item(i), "|", """," + vbCrLf + """") + """]"
        converterString = "[""" + Replace(rangeCollection.Item(i), "|", """,""") + """]"
        
        If i = 1 Then
            valuesList = converterString
        Else
            valuesList = valuesList + "," + converterString
        End If
        
    Next i
    
    converterCollectionToList = valuesList
    
End Function

Private Function createBodyJson(ByVal majordimension As String, _
                                ByVal rng As String, _
                                ByRef rangeCollection As Collection) As String
  Rem PROBAR!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Dim body As String
    Dim rangeValue As String
    
    rangeValue = converterCollectionToList(rangeCollection)
    
    body = formatString(Replace("{'majorDimension':'{0}','range':'{1}','values':[{2}]}", "'", """"), majordimension, rng, rangeValue)
    createBodyJson = body
    
End Function

