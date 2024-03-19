Attribute VB_Name = "Main"
Sub init_FlowOauth()
    
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApiKey As String
    Dim oFlowOauth As New FlowOauth
        
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
        
    oFlowOauth.InitializeFlow _
                            credentialsClient, _
                            credentialsToken, _
                            credentialsApiKey, _
                            OU_SCOPE_SPREADSHEETS
       
End Sub
Sub copyTo_SpreadSheetSheet()
    
    Dim credentialsClient As String, credentialsToken As String, credentialsApiKey As String
    Dim oFlowOauth As New FlowOauth
    Dim oSpreadSheet As New SpreadSheetSheet
    Dim json As String
    Dim SheetId As Long
    Dim spreadsheetsId As String
    Dim destinationSpreadsheetId As String
        
        
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
        
        
    spreadsheetsId = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
    SheetId = 0
    destinationSpreadsheetId = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
    
    oFlowOauth.InitializeFlow _
                            credentialsClient, _
                            credentialsToken, _
                            credentialsApiKey, _
                            OU_SCOPE_SPREADSHEETS
    
    With oSpreadSheet
        .ConnectionService oFlowOauth
         json = .CopyTo(spreadsheetsId, SheetId, destinationSpreadsheetId)
    End With
    
    Debug.Print json
    
End Sub

Sub get_GoogleSheetWorkBook()
    
    Dim credentialsClient As String, credentialsToken As String, credentialsApiKey As String
    Dim oFlowOauth As New FlowOauth
    Dim oSpreadSheet As New SpreadSheetSheet
    Dim oGoogleSheet As New GoogleSheetWorkBook
    Dim json As String
    Dim range As String
    Dim spreadsheetsId As String
    
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
    
    spreadsheetsId = "18Ady06ugiz971soeIpbuuBtrtwAugv9HXyMJs1Dun4I"
    
   oFlowOauth.InitializeFlow _
                            credentialsClient, _
                            credentialsToken, _
                            credentialsApiKey, _
                            OU_SCOPE_SPREADSHEETS
    
    With oSpreadSheet
        .ConnectionService oFlowOauth
         json = .RecoveryById(spreadsheetsId)
    End With
    
    With oGoogleSheet
        .Create json
        Debug.Print .Properties("title")
        Debug.Print .Sheets("title", 2)
        Debug.Print .Sheets("sheetId", 2)
    End With

End Sub

Sub get__SpreadSheetValue()
        
    Dim credentialsClient As String, credentialsToken As String, credentialsApiKey As String
    Dim oFlowOauth As New FlowOauth
    Dim SpreadSG As New SpreadSheetValue
    Dim responseJSON As String
    Dim arrValue() As Variant
    Dim strValue As String
    Dim id As String
    Dim range As String
    
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
    
    id = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
    range = "bbdd_libros"
    
    Rem comienza de el flujo de Oauth (autenticación y autorización)
    oFlowOauth.InitializeFlow _
                               credentialsClient, _
                               credentialsToken, _
                               credentialsApiKey, _
                               OU_SCOPE_SPREADSHEETS
    
    Rem realizamos la consulta con GoogleSheets
    With SpreadSG
        .ConnectionService oFlowOauth
         responseJSON = .GetValue(id, range)
    End With
    
    If SpreadSG.Operation = GO_SUCCESSFUL Then
        Rem leemos la respuesta
        arrValue = ProcessResponse.GetValue(responseJSON)
         
        For i = LBound(arrValue, 1) To UBound(arrValue, 1)
            strValue = Empty
            For o = LBound(arrValue, 2) To UBound(arrValue, 2)
        Rem la estructura condicional solo es para obtener una
        Rem mejor vista de los datos puede obviarse junto con la función ConsoleShow()
                If o = 0 Then
                    strValue = strValue & ConsoleShow(arrValue(i, o), 4)
                Else
                    strValue = strValue & ConsoleShow(arrValue(i, o), 25)
                End If
            Next o
            Debug.Print strValue
        Next i
    Else
        Debug.Print SpreadSG.DetailsError
    End If

End Sub

Sub update__SpreadSheetValue()
    
    Dim credentialsClient As String, credentialsToken As String, credentialsApiKey As String
    Dim oFlowOauth As New FlowOauth
    Dim SpreadSG As New SpreadSheetValue
    Dim json As String
    Dim updateRangeValue As New Collection
    Dim id As String
    Dim range As String
    
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
    
    id = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
    range = "BBDD_LIBROS!a16:d18"

    updateRangeValue.Add "45|Rimas y leyendas|9789583003103|Gustavo Adolfo Bécquer"
    updateRangeValue.Add "85|Estudio Escarlata|9786075562261|Arhur Conan Doyle"
    updateRangeValue.Add "CANTIDAD DE DATOS|=COUNTA(A1:A15)"
'    updateRangeValue.Add "=QUERY(BBDD_LIBROS!A1:E;\""SELECT *\"";1)"
 
    oFlowOauth.InitializeFlow _
                                 credentialsClient, _
                                 credentialsToken, _
                                 credentialsApiKey, _
                                 OU_SCOPE_SPREADSHEETS
    
    With SpreadSG
        .ConnectionService oFlowOauth
         json = .Update(id, range, updateRangeValue)
    End With
    
    If SpreadSG.Operation = GO_SUCCESSFUL Then
        Debug.Print ProcessResponse.UpdateValue(json)
    Else
        Debug.Print SpreadSG.DetailsError()
    End If
    
End Sub
Sub append__SpreadSheetvalue()
    
    Dim credentialsClient As String, credentialsToken As String, credentialsApiKey As String
    Dim oFlowOauth As New FlowOauth
    Dim SpreadSG As New SpreadSheetValue
    Dim json As String
    Dim appendRangeValue As New Collection
    Dim id As String
    Dim range As String
    
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
    
    id = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
    range = "bbdd_libros"

'    appendRangeValue.Add "ID|DESCRIPCION|UME"
    appendRangeValue.Add "14|aplicaciones VBA con Excel|978612302653|Manuel Torres Remon"
    
    oFlowOauth.InitializeFlow _
                                 credentialsClient, _
                                 credentialsToken, _
                                 credentialsApiKey, _
                                 OU_SCOPE_SPREADSHEETS
    
    With SpreadSG
        .ConnectionService oFlowOauth
         json = .Append(id, range, appendRangeValue)
         
        If .Operation = GO_SUCCESSFUL Then
            Debug.Print ProcessResponse.AppendValue(json)
        Else
            Debug.Print .DetailsError()
        End If
    End With
    
End Sub

Sub clear__SpreadSheetValue()
    
    Dim credentialsClient As String, credentialsToken As String, credentialsApiKey As String
    Dim oFlowOauth As New FlowOauth
    Dim SpreadSG As New SpreadSheetValue
    Dim json As String
    
    Dim id As String
    Dim rng As String
    
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
    
    id = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
    rng = "BBDD_LIBROS!a14:d14"

    oFlowOauth.InitializeFlow _
                                 credentialsClient, _
                                 credentialsToken, _
                                 credentialsApiKey, _
                                 OU_SCOPE_SPREADSHEETS
    
    With SpreadSG
        .ConnectionService oFlowOauth
         json = .Clear(id, rng)
         
         If .Operation = GO_SUCCESSFUL Then
            Debug.Print ProcessResponse.ClearValues(json)
         Else
            Debug.Print .DetailsError()
         End If
    End With
    
End Sub

Sub create_SpreadSheetSheet()

    Dim credentialsClient As String, credentialsToken As String, credentialsApiKey As String
    Dim oFlowOauth As New FlowOauth
    Dim SpreadSG As New SpreadSheetSheet
    Dim GoogleSheet As New GoogleSheetWorkBook
    Dim json As String
    
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
    
    oFlowOauth.InitializeFlow _
                                 credentialsClient, _
                                 credentialsToken, _
                                 credentialsApiKey, _
                                 OU_SCOPE_SPREADSHEETS
    
    With SpreadSG
        .ConnectionService oFlowOauth
        json = .CreateWorkBook()
    End With
    
    
    With GoogleSheet
        .Create json
        Debug.Print "datos generales"
        Debug.Print "url : "; .SpreadSheetUrl
        Debug.Print "ID : "; .SpreadsheetId
        Debug.Print "nombre libro--> "; .Properties("title")
        Debug.Print "Local--> "; .Properties("locale")
        Debug.Print "Zona Horaria--> "; .Properties("timeZone")
        
        Debug.Print "Datos de la hoja : "
        Debug.Print "Nombre de la hoja--> "; .Sheets("title", 1)
        Debug.Print "Id de la hoja--> "; .Sheets("sheetId")
        Debug.Print "Indice de la hoja--> "; .Sheets("index")
        Debug.Print "Tipo de la hoja--> "; .Sheets("sheetType")
        
        Shell "cmd.exe /k start chrome.exe " & .SpreadSheetUrl, vbHide
        
    End With
    
End Sub

Sub create_SpreadSheetSheet_with_collections()
    
    Dim credentialsClient As String, credentialsToken As String, credentialsApiKey As String
    Dim oFlowOauth As New FlowOauth
    Dim SpreadSG As New SpreadSheetSheet
    Dim GoogleSheet As New GoogleSheetWorkBook
    Dim json As String
    
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
    
    oFlowOauth.InitializeFlow _
                                 credentialsClient, _
                                 credentialsToken, _
                                 credentialsApiKey, _
                                 OU_SCOPE_SPREADSHEETS
    
    With SpreadSG
        .ConnectionService oFlowOauth
        json = .CreateWorkBook()
    End With
    
    
    With GoogleSheet
        .Create json
        
        For i = 1 To .GoogleSheets.Count
            With .GoogleSheets
                Debug.Print .Count
                Debug.Print .Item(i).Title
                Debug.Print .Item(i).SheetId
                
                Debug.Print .Item(i).gridProperties.ColumnCount
                Debug.Print .Item(i).gridProperties.RowCount
            End With
            Debug.Print " ----- Fin " & .GoogleSheets.Item(i).Title & " -------------" & vbCrLf
        Next i
        
        Debug.Print .Properties("title")
        Debug.Print .SpreadsheetId
        Debug.Print .SpreadSheetUrl
    End With
    
End Sub
Sub recovery_SpreadSheetSheet_with_collections()
    
    Dim credentialsClient As String, credentialsToken As String, credentialsApiKey As String
    Dim oFlowOauth As New FlowOauth
    Dim SpreadSG As New SpreadSheetSheet
    Dim GoogleSheet As New GoogleSheetWorkBook
    Dim json As String
    Dim id As String
    
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
    
    oFlowOauth.InitializeFlow _
                                 credentialsClient, _
                                 credentialsToken, _
                                 credentialsApiKey, _
                                 OU_SCOPE_SPREADSHEETS
    
    id = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
    
    With SpreadSG
        .ConnectionService oFlowOauth
        json = .RecoveryById(id)
    End With
    
    
    With GoogleSheet
        .Create json
        
        For i = 1 To .GoogleSheets.Count
            With .GoogleSheets
                Debug.Print "Este libro contiene "; .Count; " hojas"
                Debug.Print "Nombre de la hoja "; .Item(i).Title
                Debug.Print "SheetId "; .Item(i).SheetId
                
                Debug.Print "Columnas "; .Item(i).gridProperties.ColumnCount
                Debug.Print "Filas "; .Item(i).gridProperties.RowCount
            End With
            Debug.Print " ----- Fin " & .GoogleSheets.Item(i).Title & " -------------" & vbCrLf
        Next i
        
        Debug.Print "Nombre del libro "; .Properties("title")
        Debug.Print "Id del libro "; .SpreadsheetId
        Debug.Print "Url del libro "; .SpreadSheetUrl
    End With
    
End Sub


Sub batchUpdate_SpreadSheetSheet()
    
    Dim credentialsClient As String, credentialsToken As String, credentialsApiKey As String
    Dim SpreadSG As New SpreadSheetSheet
    Dim requets As New Collection
    Dim gf_FindReplace As unionFieldScope
    Dim gf_DeleteRange As unionFieldScope
    Dim json As String
    Dim text As String
    Dim id As String
    
    Dim oFlowOauth As New FlowOauth
    
    credentialsClient = CurrentProject.Path + "\credentials\client_secret.json"
    credentialsToken = CurrentProject.Path + "\credentials\credentials_token.json"
    credentialsApiKey = CurrentProject.Path + "\credentials\api_key.json"
    
    id = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
    
    oFlowOauth.InitializeFlow _
                                credentialsClient, _
                                credentialsToken, _
                                credentialsApiKey, _
                                OU_SCOPE_SPREADSHEETS
    
    With gf_FindReplace.rng
        .SheetId = 0
        .startRowIndex = 1
        .endRowIndex = 25
        .startColumnIndex = 0
        .endColumnIndex = 5
    End With
    
    With SpreadSG
        .ConnectionService oFlowOauth
        find_replace = .FindReplace("852963741", "2099", gf_FindReplace)
'        delete_replace = .DeleteRange(gf_FindReplace, "ROWS")
        
        requets.Add find_replace
'        requets.Add delete_replace

        json = .batchUpdate(id, requets)
        
        If .Operation = GO_SUCCESSFUL Then
            Debug.Print "Actualización aplicada a : "; ProcessResponse.batchUpdate(json)
        Else
            Debug.Print .DetailsError
        End If
    End With

End Sub

