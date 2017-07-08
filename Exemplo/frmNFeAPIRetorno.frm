VERSION 5.00
Begin VB.Form frmNFeAPIRetorno 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Busca Retorno de Processamento Documento"
   ClientHeight    =   8925
   ClientLeft      =   5370
   ClientTop       =   1545
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   25405.32
   ScaleMode       =   0  'User
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox comboTpDown 
      Height          =   315
      ItemData        =   "frmNFeAPIRetorno.frx":0000
      Left            =   2280
      List            =   "frmNFeAPIRetorno.frx":0013
      TabIndex        =   26
      Text            =   "X"
      Top             =   8160
      Width           =   975
   End
   Begin VB.CheckBox checkExibir 
      Caption         =   "Exibir PDF"
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimirDoc 
      Caption         =   "Imprimir Documento Autorizado"
      Height          =   615
      Left            =   5520
      TabIndex        =   23
      Top             =   8040
      Width           =   3735
   End
   Begin VB.TextBox txtdhRecbto 
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   7200
      Width           =   4695
   End
   Begin VB.TextBox txtnProt 
      Height          =   315
      Left            =   4920
      TabIndex        =   19
      Top             =   7200
      Width           =   5415
   End
   Begin VB.TextBox txtStatusSefaz 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   6480
      Width           =   1695
   End
   Begin VB.TextBox txtMotivoSefaz 
      Height          =   315
      Left            =   1920
      TabIndex        =   15
      Top             =   6480
      Width           =   8415
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtMotivo 
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   5040
      Width           =   8415
   End
   Begin VB.TextBox txtChaveRetorno 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   5760
      Width           =   10215
   End
   Begin VB.TextBox txtnsNRec 
      Height          =   315
      Left            =   5280
      TabIndex        =   8
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtCNPJ 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   5055
   End
   Begin VB.TextBox txtToken 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   10215
   End
   Begin VB.CommandButton cmdTestar 
      Caption         =   "Verificar Retorno de Processamento do Documento"
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox txtResult 
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2520
      Width           =   10215
   End
   Begin VB.Label Label13 
      Caption         =   "Tipo de Download:"
      Height          =   255
      Left            =   840
      TabIndex        =   25
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Protocolo de Autorização"
      Height          =   195
      Left            =   4920
      TabIndex        =   22
      Top             =   6960
      Width           =   1785
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Data e Hora de Recebimento"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Width           =   2085
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Motivo Sefaz"
      Height          =   195
      Left            =   1920
      TabIndex        =   18
      Top             =   6240
      Width           =   930
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Status Sefaz"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   6240
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Processamento Motivo"
      Height          =   195
      Left            =   1920
      TabIndex        =   14
      Top             =   4800
      Width           =   1620
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Processamento Status"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   1590
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Chave de Acesso Documento"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   2130
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "nRec do Envio para Sefaz"
      Height          =   195
      Left            =   5280
      TabIndex        =   7
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informe o Token Liberado para Sua Software House"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CNPJ Emitente"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resposta do Servidor"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2280
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
    'Requisitando download para a API
    Dim result As String
    Dim isShow As Boolean
    isShow = checkExibir.Value
    
    'lendo o responsetext, que é onde está ou estarão o xml, pdf, JSON conforme o tipo informado
    result = downloadNFeAndSave(txtToken.Text, txtChaveRetorno.Text, comboTpDown.Text, "C:\Documentos", isShow)
    'result = downloadNFe(txtToken.Text, txtChaveRetorno.Text, comboTpDown.Text)
    txtResult.Text = result
    
    Exit Sub
SAI:
    txtResult.Text = ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description)
End Sub

Private Sub cmdTestar_Click()
On Error GoTo SAI
    Dim result As String
    result = consultaStatusProcessamento(txtToken.Text, txtCNPJ.Text, txtnsNRec.Text)
    txtResult.Text = result
    Dim protocolo As String
    
    
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
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI
End Sub
