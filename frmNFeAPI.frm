VERSION 5.00
Begin VB.Form frmNFeAPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NF-e API"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkTXT 
      Caption         =   "Enviar TXT para Processamento"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txtResult 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   6120
      Width           =   10215
   End
   Begin VB.TextBox txtJSON 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1080
      Width           =   10215
   End
   Begin VB.CommandButton cmdTestar 
      Caption         =   "Enviar Documento para Processamento >>>>>>"
      Height          =   615
      Left            =   6600
      TabIndex        =   1
      Top             =   4800
      Width           =   3735
   End
   Begin VB.TextBox txtToken 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resposta do Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Informe o Documento em Formato JSON para Transmitir ao Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informe o Token Liberado para Sua Software House"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3705
   End
End
Attribute VB_Name = "frmNFeAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdTestar_Click()
On Error GoTo SAI
    'VARIAVEL QUE VAI ARMAZENAR O CONTEUDO A SER ENVIADO A NS API
    Dim conteudoEnviar As String
    Dim CNPJ As String
    'verificação se o que será enviado é uma mensagem JSON ou TXT
    
        
    txtJSON.Text = Trim(txtJSON.Text)
    
    
    If chkTXT.Value = 0 Then
        'processamento de uma mensagem JSON
        
                'Checa se o arquivo json começa com chaves
                If (Mid(txtJSON.Text, 1, 1) = "{" And Mid(txtJSON.Text, Len(txtJSON.Text), 1) = "}") Then
                        txtJSON.Text = Mid(txtJSON.Text, 2, Len(txtJSON.Text) - 2)
                End If
                
        'INICIALIZAÇÃO DA VARIAVEL
        conteudoEnviar = "{ "
        'MONTANDO A PARTE DE AUTENTICAÇÃO NA API, OU SEJA O TOKEN DA SOFTWARE HOUSE
        conteudoEnviar = conteudoEnviar & """X-AUTH-TOKEN"": """ & txtToken.Text & ""","
        'COMPLEMENTANDO A VARIAVEL COM O CONTEUDO DA NFE
        conteudoEnviar = conteudoEnviar & txtJSON.Text
        'FECHANDO STRING JSON PARA ENVIO AO SERVER
        conteudoEnviar = conteudoEnviar & " }"
        'CHAMANDO FUNÇÃO QUE CONSOME A API passando como padrão de mensagem "application/json" que siginifica que estarei enviando um json para processamento
        txtResult.Text = enviaSolicitacaoJSON("https://nfe.ns.eti.br/nfe/issue", "application/json", conteudoEnviar, txtToken.Text)

        'PEGANDO CNPJ
        Dim localCNPJ As Integer
        localCNPJ = InStr(1, txtJSON.Text, "CNPJ", 1) + 6
        localCNPJ = InStr(localCNPJ, txtJSON.Text, """", 1) + 1
        CNPJ = Mid(txtJSON.Text, localCNPJ, 14)
    Else
        'processamento de um TXT
        
        conteudoEnviar = txtJSON.Text
        'CHAMANDO FUNÇÃO QUE CONSOME A API passando como padrão de mensagem "text/plain" que siginifica que estarei enviando um txt para processamento
        txtResult.Text = enviaSolicitacaoJSON("https://nfe.ns.eti.br/nfe/issue", "text/plain", conteudoEnviar, txtToken.Text)
        
                'PEGANDO CNPJ
        CNPJ = Mid(txtJSON.Text, InStr(1, txtJSON.Text, "C02|", 1) + 4, 14)
        
    End If
    Dim result As String
    result = responseText
    With frmNFeAPIRetorno
        .txtToken.Text = txtToken.Text
        .txtCNPJ.Text = CNPJ
        'lendo o nnNRec do JSON recebido
        .txtnsNRec.Text = LerDadosJSON(result, "nsNRec", "", "")
        'montando o JSON para buscar o status do processamento do documento
        .txtJSON.Text = "{"
        .txtJSON.Text = .txtJSON.Text & """X-AUTH-TOKEN"":""" & .txtToken.Text & ""","
        .txtJSON.Text = .txtJSON.Text & """CNPJ"":""" & .txtCNPJ.Text & ""","
        .txtJSON.Text = .txtJSON.Text & """nsNRec"":""" & .txtnsNRec.Text & """"
        .txtJSON.Text = .txtJSON.Text & "}"
     
        'abrindo formulario para buscar retorno
        .Show 1
    End With
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI
End Sub


