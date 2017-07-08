VERSION 5.00
Begin VB.Form frmNFeAPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NF-e API"
   ClientHeight    =   9300
   ClientLeft      =   6810
   ClientTop       =   990
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   10500
   Begin VB.CommandButton cmdConsultaDownload 
      Caption         =   "Consulta de Status / Download de NFe"
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   4800
      Width           =   3735
   End
   Begin VB.ComboBox cbTpConteudo 
      Height          =   315
      ItemData        =   "frmNFeAPI.frx":0000
      Left            =   9000
      List            =   "frmNFeAPI.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   1215
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
   Begin VB.TextBox txtConteudo 
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
      Caption         =   "Conteudo"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   690
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

Private Sub cmdConsultaDownload_Click()
    With frmNFeAPIRetorno
        .txtToken.Text = txtToken.Text
        .Show 1
    End With
    
    Exit Sub

End Sub

Private Sub cmdTestar_Click()
    On Error GoTo SAI
    'VARIAVEL QUE VAI ARMAZENAR O CONTEUDO A SER ENVIADO A NS API
    Dim retorno As String
    Dim CNPJ As String
    'verificação se o que será enviado é uma mensagem JSON ou TXT
    
        
    txtConteudo.Text = Trim(txtConteudo.Text)
    If (txtConteudo.Text <> "") Then
        retorno = emitirNFe(txtToken.Text, txtConteudo.Text, cbTpConteudo.Text)
    End If
    txtResult.Text = retorno
    
    Dim result As String
    result = responseText
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleCTeAPI

End Sub
