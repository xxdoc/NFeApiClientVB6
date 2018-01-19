Attribute VB_Name = "NFeAPI"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'activate Microsoft XML, v6.0 in references
Public responseText As String
Function enviaConteudoParaAPI(token As String, conteudo As String, url As String, tpConteudo As String) As String
On Error GoTo SAI
    Dim contentType As String
    
    If (tpConteudo = "txt") Then
        contentType = "text/plain"
    ElseIf (tpConteudo = "xml") Then
        contentType = "application/xml"
    Else
        contentType = "application/json"
    End If
    
    Dim obj As MSXML2.ServerXMLHTTP
    Set obj = New MSXML2.ServerXMLHTTP
    obj.Open "POST", url
    obj.setRequestHeader "Content-Type", contentType
    If Trim(token) <> "" Then
        obj.setRequestHeader "X-AUTH-TOKEN", token
    End If
    obj.send conteudo
    Dim resposta As String
    resposta = obj.responseText
    enviaConteudoParaAPI = resposta
    Exit Function
SAI:
  enviaConteudoParaAPI = Err.Number & " " & Err.Description
End Function

Public Function emitirNFe(token As String, conteudo As String, tpConteudo As String) As String
    Dim url As String
    url = "https://nfe.ns.eti.br/nfe/issue"
    emitirNFe = enviaConteudoParaAPI(token, conteudo, url, tpConteudo)
End Function

Public Function consultaStatusProcessamento(token As String, CNPJ As String, nsNRec As String, tpAmb As String) As String
    Dim conteudo As String
    conteudo = "{"
    conteudo = conteudo & """X-AUTH-TOKEN"":""" & token & ""","
    conteudo = conteudo & """CNPJ"":""" & CNPJ & ""","
    conteudo = conteudo & """nsNRec"":""" & nsNRec & ""","
    conteudo = conteudo & """tpAmb"":""" & tpAmb & """"
    conteudo = conteudo & "}"
    
    Dim url As String
    url = "https://nfe.ns.eti.br/nfe/issue/status"
    
    consultaStatusProcessamento = enviaConteudoParaAPI(token, conteudo, url, "json")
End Function

Public Function downloadNFe(token As String, chNFe As String, tpDown As String, tpAmb As String) As String
    Dim conteudo As String
    conteudo = "{"
    conteudo = conteudo & """X-AUTH-TOKEN"":""" & token & ""","
    conteudo = conteudo & """chNFe"":""" & chNFe & ""","
    conteudo = conteudo & """tpDown"":""" & tpDown & ""","
    conteudo = conteudo & """tpAmb"":""" & tpAmb & """"
    conteudo = conteudo & "}"
    
    Dim url As String
    url = "https://nfe.ns.eti.br/nfe/get"
    
    downloadNFe = enviaConteudoParaAPI(token, conteudo, url, "json")
End Function

Public Function downloadNFeAndSave(token As String, chNFe As String, tpDown As String, tpAmb As String, caminho As String, isShow As Boolean)
    Dim retornoAPI As String
    retornoAPI = downloadNFe(token, chNFe, tpDown, tpAmb)
    
    downloadNFeAndSave = retornoAPI
    
    Dim baixarJson, baixarXML, baixarPDF As Integer
    
    baixarJson = InStr(1, tpDown, "J")
    baixarXML = InStr(1, tpDown, "X")
    baixarPDF = InStr(1, tpDown, "P")
    
    If (Len(caminho) > 0) Then
        If (Not (Right(caminho, 1) = "\")) Then
            caminho = caminho & "\"
        End If
        
        If Len(Dir(caminho, vbDirectory) & "") = 0 Then
            Call criarPastas(caminho)
        End If
    End If
    Dim caminhoNomeArquivo As String
    
    If baixarXML <> 0 Then
        'salvando xml no diretorio
        caminhoNomeArquivo = caminho & chNFe & "-procNfe.xml"
        If Len(Dir(caminhoNomeArquivo, vbDirectory) & "") <> 0 Then
            Kill (caminhoNomeArquivo)
        End If
        Call Salvar_Arquivo(caminhoNomeArquivo, LerDadosJSON(retornoAPI, "xml", "", ""))
    Else
        If baixarJson <> 0 Then
            Dim inicioJson As Integer
            Dim fimJson As Integer
            Dim jsonSalvar As String
            'separando json do restante do retorno
            inicioJson = InStr(1, retornoAPI, """NFe""", 1) - 1
            
            If InStr(1, retornoAPI, """pdf""", 1) <> 0 Then
                fimJson = InStr(1, retornoAPI, """pdf""", 1) - 1
            Else
                fimJson = Len(retornoAPI)
            End If
            
            jsonSalvar = Mid(retornoAPI, inicioJson, fimJson - inicioJson)
            
            'salvando json no diretorio
            caminhoNomeArquivo = caminho & chNFe & "-procNfe.json"
            If Len(Dir(caminhoNomeArquivo, vbDirectory) & "") <> 0 Then
                Kill (caminhoNomeArquivo)
            End If
            Call Salvar_Arquivo(caminhoNomeArquivo, jsonSalvar)
        End If
    End If
    
    If baixarPDF <> 0 Then
        Dim resultPDF As String
        'Lendo somente o base64 do PDF recebido
        resultPDF = LerDadosJSON(retornoAPI, "pdf", "", "")
    
        'gerando o pdf a partir do base64 lido no JSON acima citado
        caminhoNomeArquivo = caminho & chNFe & "-procNfe.pdf"
        If Len(Dir(caminhoNomeArquivo, vbDirectory) & "") <> 0 Then
            Kill (caminhoNomeArquivo)
        End If
        
        Call savePDF(resultPDF, caminhoNomeArquivo)
        If isShow Then
            'Abrindo o PDF gerado acima
            ShellExecute 0, "open", caminho & chNFe & "-procNfe.pdf", "", "", vbNormalFocus
        End If
    End If
    
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

Public Function LerDadosXML(sXml As String, Key1 As String, Key2 As String) As String
    On Error Resume Next
    LerDadosXML = ""
    
    Set xml = New DOMDocument
    xml.async = False
    
    If xml.loadXML(sXml) Then
        'Tentar pegar o strCampoXML
        Set objNodeList = xml.getElementsByTagName(Key1 & "//" & Key2)
        Set objNode = objNodeList.nextNode
        
        Dim valor As String
        valor = objNode.Text
        
        If Len(Trim(valor)) > 0 Then 'CONSEGUI LER O XML NODE
            LerDadosXML = valor
        End If
        Else
        MsgBox "Não foi possível ler o conteúdo do XML da NFe especificado para leitura.", vbCritical, "ERRO"
    End If
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

Public Function Salvar_Arquivo(fileName As String, conteudo As String) As Boolean
    Dim fsT As Object
    Set fsT = CreateObject("ADODB.Stream")
    fsT.Type = 2
    fsT.Charset = "utf-8"
    fsT.Open
    fsT.WriteText conteudo
    fsT.SaveToFile fileName
    
End Function

Public Function criarPastas(caminho As String) As Boolean
    Dim diretorio() As String
    diretorio = Split(caminho, "\")
    Dim caminhoAtual As String
    caminhoAtual = diretorio(0)
    
    For i = 1 To UBound(diretorio)
        If Len(Dir(caminhoAtual, vbDirectory) & "") = 0 Then
            MkDir caminhoAtual
        End If
        caminhoAtual = caminhoAtual & "\" & diretorio(i)
    Next i
End Function
