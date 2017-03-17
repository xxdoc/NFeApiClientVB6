Attribute VB_Name = "ComunicaoAPI"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'activate Microsoft XML, v6.0 in references
Public responseText As String
Function enviaSolicitacaoJSON(sUrl As String, sContentType As String, sContent As String, Token As String) As String
On Error GoTo SAI
    Dim obj As MSXML2.ServerXMLHTTP
    Set obj = New MSXML2.ServerXMLHTTP
    obj.Open "POST", sUrl
    obj.setRequestHeader "Content-Type", sContentType
    If Trim(Token) <> "" Then
        obj.setRequestHeader "X-AUTH-TOKEN", Token
    End If
    obj.send sContent
    Dim resposta As String
    resposta = "Status: " & obj.Status & vbNewLine & "Motivo: " & obj.statusText
    enviaSolicitacaoJSON = resposta & vbNewLine
    enviaSolicitacaoJSON = enviaSolicitacaoJSON & "ResponseText: " & obj.responseText
    responseText = obj.responseText
    Set obj = Nothing
    Exit Function
SAI:
  MsgBox (Err.Number & " " & Err.Description)
End Function


'activate microsoft script control 1.0 in references
Public Function LerDadosJSON(sJsonString As String, Key1 As String, Key2 As String, key3 As String) As String
On Error GoTo err_handler
    Dim oScriptEngine As ScriptControl
    Set oScriptEngine = New ScriptControl
    oScriptEngine.Language = "JScript"
    Dim objJSON As Object
    Set objJSON = oScriptEngine.Eval("(" + sJsonString + ")")
    If Key1 <> "" And Key2 <> "" And key3 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(VBA.CallByName(objJSON, Key1, VbGet), Key2, VbGet), key3, VbGet)
    ElseIf Key1 <> "" And Key2 <> "" Then
        LerDadosJSON = VBA.CallByName(VBA.CallByName(objJSON, Key1, VbGet), Key2, VbGet)
    ElseIf Key1 <> "" Then
        LerDadosJSON = VBA.CallByName(objJSON, Key1, VbGet)
    End If
Err_Exit:
    Exit Function
err_handler:
    LerDadosJSON = "Error: " & Err.Description
    Resume Err_Exit
End Function


Public Function savePDF(base64PDF As String, fileName As String) As Boolean
On Error GoTo SAI
    Dim fnum
    fnum = FreeFile
    Open fileName For Binary As #fnum
    Put #fnum, 1, Base64Decode(base64PDF)
    Close fnum
    Exit Function
SAI:
    MsgBox (Err.Number & " - " & Err.Description), vbCritical
End Function

