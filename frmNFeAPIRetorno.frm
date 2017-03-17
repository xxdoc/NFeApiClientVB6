VERSION 5.00
Begin VB.Form frmNFeAPIRetorno 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Busca Retorno de Processamento Documento"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   27198.64
   ScaleMode       =   0  'User
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox comboTpDown 
      Height          =   315
      ItemData        =   "frmNFeAPIRetorno.frx":0000
      Left            =   2160
      List            =   "frmNFeAPIRetorno.frx":0013
      TabIndex        =   28
      Text            =   "X"
      Top             =   8880
      Width           =   975
   End
   Begin VB.CheckBox checkExibir 
      Caption         =   "Exibir PDF"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   26
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimirDoc 
      Caption         =   "Imprimir Documento Autorizado"
      Enabled         =   0   'False
      Height          =   615
      Left            =   5400
      TabIndex        =   25
      Top             =   8760
      Width           =   3735
   End
   Begin VB.TextBox txtdhRecbto 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   22
      Top             =   8160
      Width           =   4695
   End
   Begin VB.TextBox txtnProt 
      Enabled         =   0   'False
      Height          =   315
      Left            =   4920
      TabIndex        =   21
      Top             =   8160
      Width           =   5415
   End
   Begin VB.TextBox txtStatusSefaz 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox txtMotivoSefaz 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      TabIndex        =   17
      Top             =   7440
      Width           =   8415
   End
   Begin VB.TextBox txtStatus 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox txtMotivo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1920
      TabIndex        =   12
      Top             =   6000
      Width           =   8415
   End
   Begin VB.TextBox txtChaveRetorno 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   6720
      Width           =   10215
   End
   Begin VB.TextBox txtnsNRec 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5280
      TabIndex        =   9
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtCNPJ 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtToken 
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   10215
   End
   Begin VB.CommandButton cmdTestar 
      Caption         =   "Verificar Retorno de Processamento do Documento"
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox txtJSON 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1680
      Width           =   10215
   End
   Begin VB.TextBox txtResult 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4080
      Width           =   10215
   End
   Begin VB.Label Label13 
      Caption         =   "Tipo de Download:"
      Height          =   255
      Left            =   720
      TabIndex        =   27
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Protocolo de Autorização"
      Height          =   195
      Left            =   4920
      TabIndex        =   24
      Top             =   7920
      Width           =   1785
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Data e Hora de Recebimento"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   7920
      Width           =   2085
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Motivo Sefaz"
      Height          =   195
      Left            =   1920
      TabIndex        =   20
      Top             =   7200
      Width           =   930
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Status Sefaz"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Processamento Motivo"
      Height          =   195
      Left            =   1920
      TabIndex        =   16
      Top             =   5760
      Width           =   1620
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Processamento Status"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   5760
      Width           =   1590
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Chave de Acesso Documento"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   2130
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "JSON a Ser Enviado a API"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   1905
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "nRec do Envio para Sefaz"
      Height          =   195
      Left            =   5280
      TabIndex        =   8
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informe o Token Liberado para Sua Software House"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ Emitente"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resposta do Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1530
   End
End
Attribute VB_Name = "frmNFeAPIRetorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimirDoc_Click()
On Error GoTo SAI
    Dim baixarPDF As Integer
    Dim baixarJSON As Integer
    Dim baixarXML As Integer
    Dim conteudoEnviar As String
    'montando JSON para requisição de Download
    conteudoEnviar = "{"
    conteudoEnviar = conteudoEnviar & """X-AUTH-TOKEN"":""" & txtToken.Text & ""","
    conteudoEnviar = conteudoEnviar & """chNFe"":""" & txtChaveRetorno.Text & ""","
    conteudoEnviar = conteudoEnviar & """tpDown"":""" & comboTpDown.Text & """"
    'Tipos de Download possiveis
    'X  = XML
    'J = JSON
    'P = PDF
    'XP = XML E PDF
    'JP = JSON E PDF
    conteudoEnviar = conteudoEnviar & "}"
    
    'Requisitando download para a API
    Call enviaSolicitacaoJSON("https://nfe.ns.eti.br/nfe/get", "application/json", conteudoEnviar, txtToken.Text)
    Dim result As String
    'lendo o responsetext, que é onde está ou estarão o xml, pdf, JSON conforme o tipo informado
    result = responseText
    txtResult = responseText
    
    baixarJSON = InStr(1, comboTpDown.Text, "J")
    baixarXML = InStr(1, comboTpDown.Text, "X")
    baixarPDF = InStr(1, comboTpDown.Text, "P")
    
    If baixarXML <> 0 Then
        'salvando xml no diretorio
        Call Salvar_Arquivo(App.Path & "\XML\" & txtChaveRetorno.Text & "-procNfe.xml", LerDadosJSON(result, "xml", "", ""))
    Else
        If baixarJSON <> 0 Then
            Dim inicioJson As Integer
            Dim fimJson As Integer
            Dim jsonSalvar As String
            
                        'separando json do restante do retorno
            inicioJson = InStr(1, result, """NFe""", 1) - 1
            
            If InStr(1, result, """pdf""", 1) <> 0 Then
                fimJson = InStr(1, result, """pdf""", 1) - 1
            Else
                fimJson = Len(result)
            End If
            
            jsonSalvar = Mid(result, inicioJson, fimJson - inicioJson)
            
                        'salvando json no diretorio
            Call Salvar_Arquivo(App.Path & "\JSON\" & txtChaveRetorno.Text & "-procNfe.json", jsonSalvar)
        End If
    End If
    
    If baixarPDF <> 0 Then
        Dim resultPDF As String
        'Lendo somente o base64 do PDF recebido
        resultPDF = LerDadosJSON(result, "pdf", "", "")
    
        'gerando o pdf a partir do base64 lido no JSON acima citado
        Call savePDF(resultPDF, App.Path & "\PDF\" & txtChaveRetorno.Text & "-procNfe.pdf")
        If checkExibir.Value = 1 Then
            'Abrindo o PDF gerado acima
            ShellExecute 0, "open", App.Path & "\PDF\" & txtChaveRetorno.Text & "-procNfe.pdf", "", "", vbNormalFocus
        End If
    End If
        
        
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI
End Sub

Private Sub cmdTestar_Click()
On Error GoTo SAI
    Dim conteudoEnviar As String
    conteudoEnviar = txtJSON.Text
    'requisitando status de processamento a API REST
    txtResult.Text = enviaSolicitacaoJSON("https://nfe.ns.eti.br/nfe/issue/status", "application/json", conteudoEnviar, txtToken.Text)
    Dim result As String
    Dim protocolo As String
    'lendo o responseText, onde estão os retornos de processamento
    result = responseText
    
    'lendo status do JSON recebido da API
    txtStatus.Text = LerDadosJSON(result, "status", "", "")
    'lendo motivo do JSON recebido da API
    txtMotivo.Text = LerDadosJSON(result, "motivo", "", "")
    'lendo chave de acesso do JSON recebido da API
    txtChaveRetorno.Text = LerDadosJSON(result, "chNFe", "", "")
    'lendo Data e Hora de Recebimento na Sefaz, retornado no JSON recebido da API
    txtdhRecbto.Text = LerDadosJSON(result, "dhRecbto", "", "")
    'lendo cSat da Sefaz retornado no JSON recebido da API
    txtStatusSefaz.Text = LerDadosJSON(result, "cStat", "", "")
    'lendo xMotivo da Sefaz retornado no JSON recebido da API
    txtMotivoSefaz.Text = LerDadosJSON(result, "xMotivo", "", "")
    'lendo nProt(Protocolo de Autorização) retornado no JSON recebido da API
    protocolo = LerDadosJSON(result, "nProt", "", "")
    If txtStatusSefaz <> "100" Then
        txtnProt.Text = "Não Possui"
    Else
        txtnProt.Text = protocolo
    End If
    
    'verificação se o status da sefaz é 100(Autorizado) libera o botão de Download e Impressão
    If LerDadosJSON(result, "cStat", "", "") = "100" Then
        cmdImprimirDoc.Enabled = True
        comboTpDown.Enabled = True
        checkExibir.Enabled = True
    End If
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI
End Sub
