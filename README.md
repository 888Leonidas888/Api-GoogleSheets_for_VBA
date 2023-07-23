<h1 style='text-algn:center'>Client VBA for Api GoogleSheets</h1>

Use **Client VBA for Api Googlesheets** para crear libros, crear hojas, leer, actualizar, agregar, eliminar, actualizar por lotes el contenido de sus GoogleSheets.

Puede descargar el libro **Googlesheets_for_VBA.accdb** el cual viene equipado con todas las librerias tanto locales como de terceros. Encontrará un módulo ***Main*** el cual desarrollará las funciones habilitadas para este cliente.

Antes de comenzar con las pruebas deberá obtener las credenciales necesarias para poder acceder a su GoogleSheets y deberá asegurarse de tener todas las librerias instaladas.

## Referencias a habilitar

- **Microsoft Excel 16.0 Object Library**
- **Microsoft Scripting Runtime**
- **Microsoft XML,v6.0**

## También necesitas FLOW OAUTH FOR VBA

**Flow Oauth for VBA** hará el trabajo de solicitar el token de acceso para consumir la **Api de GoogleSheets**

- El archivo **Googlesheets_for_VBA.accdb** ya viene con **FLOW OAUTH FOR VBA** pero te dejo el enlace al repositorio por si hace falta [https://github.com/888Leonidas888/Flow-Oauth2.0-for-VBA](https://github.com/888Leonidas888/Flow-Oauth2.0-for-VBA "https://github.com/888Leonidas888/Flow-Oauth2.0-for-VBA")

## Biblioteca de terceros
>La siguiente biblioteca es fundamental para el desarrollo de este proyecto.

Descarga e instala el siguiente módulo del respositorio [https://github.com/VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON "https://github.com/VBA-tools/VBA-JSON")

- **JsonConverter.bas v2.3.1**

## Crea un proyecto en Google Api Console

Para poder consumir las Apis de Google es necesario crear un proyecto en ***Google Cloud Platform*** vea [Cómo usar OAuth 2.0 para acceder a las API de Google](https://developers.google.com/identity/protocols/oauth2?hl=es-419 "Cómo usar OAuth 2.0 para acceder a las API de Google") . 

Deberas crear las siguientes credenciales:

- **ID de cliente de Oauth**
- **Clave de Api**

Guarda estas credenciales en un lugar seguro.

Vea el siguiente video para mayor detalle de como crear las credenciales pulsando [Crear proyecto parte 1](https://www.youtube.com/watch?v=8GG7LnaMtuE&list=PLebWFysFNi3AuZOqFzKNzqHc6mPkkz1AX&index=10 "Crear proyecto parte 1")

## Hagamos código

### Leer
```vb
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
```

### Actualizar

Esta función sobreescribirá donde se le haya indicado como rango.

```vb
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
    range = "BBDD_LIBROS!a12:d13"

    updateRangeValue.Add "12|Ataque a los titanes 34|9788467948158|Hajime Isayama"
    updateRangeValue.Add "13|Ataque a los titanes 33|9788467948158|Hajime Isayama"
    
	rem   updateRangeValue.Add "=QUERY(BBDD_LIBROS!A1:E;\""SELECT *\"";1)"
 
    oFlowOauth.InitializeFlow _
                                 credentialsClient, _
                                 credentialsToken, _
                                 credentialsApiKey, _
                                 OU_SCOPE_SPREADSHEETS
    
    With SpreadSG
        .ConnectionService oFlowOauth
         json = .Update(id, range, updateRangeValue, valueInputOption:="raw")
    End With
    
    If SpreadSG.Operation = GO_SUCCESSFUL Then
        Debug.Print ProcessResponse.UpdateValue(json)
    Else
        Debug.Print SpreadSG.DetailsError()
    End If
    
End Sub
```
### Agregar

La función de agregar es muy similiar a la función de actualizar, lo que hará función de agregar será escribir una fila al final de nuestros registros.

```vb
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
```

### Borrar

```vb
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
```



