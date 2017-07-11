Attribute VB_Name = "NFeAPI"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'activate Microsoft XML, v6.0 in references

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

Public Function consultaStatusProcessamento(token As String, CNPJ As String, nsNRec As String) As String
    Dim conteudo As String
    conteudo = "{"
    conteudo = conteudo & """X-AUTH-TOKEN"":""" & token & ""","
    conteudo = conteudo & """CNPJ"":""" & CNPJ & ""","
    conteudo = conteudo & """nsNRec"":""" & nsNRec & """"
    conteudo = conteudo & "}"
    
    Dim url As String
    url = "https://nfe.ns.eti.br/nfe/issue/status"
    
    consultaStatusProcessamento = enviaConteudoParaAPI(token, conteudo, url, "json")
End Function

Public Function downloadNFe(token As String, chNFe As String, tpDown As String) As String
    Dim conteudo As String
    conteudo = "{"
    conteudo = conteudo & """X-AUTH-TOKEN"":""" & token & ""","
    conteudo = conteudo & """chNFe"":""" & chNFe & ""","
    conteudo = conteudo & """tpDown"":""" & tpDown & """"
    conteudo = conteudo & "}"
    
    Dim url As String
    url = "https://nfe.ns.eti.br/nfe/get"
    
    downloadNFe = enviaConteudoParaAPI(token, conteudo, url, "json")
End Function

Public Function downloadNFeAndSave(token As String, chNFe As String, tpDown As String, caminho As String, isShow As Boolean)
    Dim retornoAPI As String
    retornoAPI = downloadNFe(token, chNFe, tpDown)
    
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


'
'ABAIXO SEGUE UM CÓDIGO PARA CONVERSÃO DE BASE64 PARA PDF, JUNTO COM INFORMAÇÕES DO AUTOR
'




' A Base64 Encoder/Decoder.
'
' This module is used to encode and decode data in Base64 format as described in RFC 1521.
'
' Home page: www.source-code.biz.
' License: GNU/LGPL (www.gnu.org/licenses/lgpl.html).
' Copyright 2007: Christian d'Heureuse, Inventec Informatik AG, Switzerland.
' This module is provided "as is" without warranty of any kind.

Option Explicit

Private InitDone  As Boolean
Private Map1(0 To 63)  As Byte
Private Map2(0 To 127) As Byte

' Encodes a string into Base64 format.
' No blanks or line breaks are inserted.
' Parameters:
'   S         a String to be encoded.
' Returns:    a String with the Base64 encoded data.
Public Function Base64EncodeString(ByVal s As String) As String
   Base64EncodeString = Base64Encode(ConvertStringToBytes(s))
   End Function

' Encodes a byte array into Base64 format.
' No blanks or line breaks are inserted.
' Parameters:
'   InData    an array containing the data bytes to be encoded.
' Returns:    a string with the Base64 encoded data.
Public Function Base64Encode(InData() As Byte)
   Base64Encode = Base64Encode2(InData, UBound(InData) - LBound(InData) + 1)
   End Function

' Encodes a byte array into Base64 format.
' No blanks or line breaks are inserted.
' Parameters:
'   InData    an array containing the data bytes to be encoded.
'   InLen     number of bytes to process in InData.
' Returns:    a string with the Base64 encoded data.
Public Function Base64Encode2(InData() As Byte, ByVal InLen As Long) As String
   If Not InitDone Then Init
   If InLen = 0 Then Base64Encode2 = "": Exit Function
   Dim ODataLen As Long: ODataLen = (InLen * 4 + 2) \ 3     ' output length without padding
   Dim OLen As Long: OLen = ((InLen + 2) \ 3) * 4           ' output length including padding
   Dim Out() As Byte
   ReDim Out(0 To OLen - 1) As Byte
   Dim ip0 As Long: ip0 = LBound(InData)
   Dim ip As Long
   Dim op As Long
   Do While ip < InLen
      Dim i0 As Byte: i0 = InData(ip0 + ip): ip = ip + 1
      Dim i1 As Byte: If ip < InLen Then i1 = InData(ip0 + ip): ip = ip + 1 Else i1 = 0
      Dim i2 As Byte: If ip < InLen Then i2 = InData(ip0 + ip): ip = ip + 1 Else i2 = 0
      Dim o0 As Byte: o0 = i0 \ 4
      Dim o1 As Byte: o1 = ((i0 And 3) * &H10) Or (i1 \ &H10)
      Dim o2 As Byte: o2 = ((i1 And &HF) * 4) Or (i2 \ &H40)
      Dim o3 As Byte: o3 = i2 And &H3F
      Out(op) = Map1(o0): op = op + 1
      Out(op) = Map1(o1): op = op + 1
      Out(op) = IIf(op < ODataLen, Map1(o2), Asc("=")): op = op + 1
      Out(op) = IIf(op < ODataLen, Map1(o3), Asc("=")): op = op + 1
      Loop
   Base64Encode2 = ConvertBytesToString(Out)
   End Function

' Decodes a string from Base64 format.
' Parameters:
'    s        a Base64 String to be decoded.
' Returns     a String containing the decoded data.
Public Function Base64DecodeString(ByVal s As String) As String
   If s = "" Then Base64DecodeString = "": Exit Function
   Base64DecodeString = ConvertBytesToString(Base64Decode(s))
   End Function

' Decodes a byte array from Base64 format.
' Parameters
'   s         a Base64 String to be decoded.
' Returns:    an array containing the decoded data bytes.
Public Function Base64Decode(ByVal s As String) As Byte()
   If Not InitDone Then Init
   Dim IBuf() As Byte: IBuf = ConvertStringToBytes(s)
   Dim ILen As Long: ILen = UBound(IBuf) + 1
   If ILen Mod 4 <> 0 Then Err.Raise vbObjectError, , "Length of Base64 encoded input string is not a multiple of 4."
   Do While ILen > 0
      If IBuf(ILen - 1) <> Asc("=") Then Exit Do
      ILen = ILen - 1
      Loop
   Dim OLen As Long: OLen = (ILen * 3) \ 4
   Dim Out() As Byte
   ReDim Out(0 To OLen - 1) As Byte
   Dim ip As Long
   Dim op As Long
   Do While ip < ILen
      Dim i0 As Byte: i0 = IBuf(ip): ip = ip + 1
      Dim i1 As Byte: i1 = IBuf(ip): ip = ip + 1
      Dim i2 As Byte: If ip < ILen Then i2 = IBuf(ip): ip = ip + 1 Else i2 = Asc("A")
      Dim i3 As Byte: If ip < ILen Then i3 = IBuf(ip): ip = ip + 1 Else i3 = Asc("A")
      If i0 > 127 Or i1 > 127 Or i2 > 127 Or i3 > 127 Then _
         Err.Raise vbObjectError, , "Illegal charaNFer in Base64 encoded data."
      Dim b0 As Byte: b0 = Map2(i0)
      Dim b1 As Byte: b1 = Map2(i1)
      Dim b2 As Byte: b2 = Map2(i2)
      Dim b3 As Byte: b3 = Map2(i3)
      If b0 > 63 Or b1 > 63 Or b2 > 63 Or b3 > 63 Then _
         Err.Raise vbObjectError, , "Illegal charaNFer in Base64 encoded data."
      Dim o0 As Byte: o0 = (b0 * 4) Or (b1 \ &H10)
      Dim o1 As Byte: o1 = ((b1 And &HF) * &H10) Or (b2 \ 4)
      Dim o2 As Byte: o2 = ((b2 And 3) * &H40) Or b3
      Out(op) = o0: op = op + 1
      If op < OLen Then Out(op) = o1: op = op + 1
      If op < OLen Then Out(op) = o2: op = op + 1
      Loop
   Base64Decode = Out
   End Function

Private Sub Init()
   Dim c As Integer, i As Integer
   ' set Map1
   i = 0
   For c = Asc("A") To Asc("Z"): Map1(i) = c: i = i + 1: Next
   For c = Asc("a") To Asc("z"): Map1(i) = c: i = i + 1: Next
   For c = Asc("0") To Asc("9"): Map1(i) = c: i = i + 1: Next
   Map1(i) = Asc("+"): i = i + 1
   Map1(i) = Asc("/"): i = i + 1
   ' set Map2
   For i = 0 To 127: Map2(i) = 255: Next
   For i = 0 To 63: Map2(Map1(i)) = i: Next
   InitDone = True
   End Sub

Private Function ConvertStringToBytes(ByVal s As String) As Byte()
   Dim b1() As Byte: b1 = s
   Dim l As Long: l = (UBound(b1) + 1) \ 2
   If l = 0 Then ConvertStringToBytes = b1: Exit Function
   Dim b2() As Byte
   ReDim b2(0 To l - 1) As Byte
   Dim p As Long
   For p = 0 To l - 1
      Dim c As Long: c = b1(2 * p) + 256 * CLng(b1(2 * p + 1))
      If c >= 256 Then c = Asc("?")
      b2(p) = c
      Next
   ConvertStringToBytes = b2
   End Function

Private Function ConvertBytesToString(b() As Byte) As String
   Dim l As Long: l = UBound(b) - LBound(b) + 1
   Dim b2() As Byte
   ReDim b2(0 To (2 * l) - 1) As Byte
   Dim p0 As Long: p0 = LBound(b)
   Dim p As Long
   For p = 0 To l - 1: b2(2 * p) = b(p0 + p): Next
   Dim s As String: s = b2
   ConvertBytesToString = s
End Function
