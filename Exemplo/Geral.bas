Attribute VB_Name = "Geral"
Public NFe As String
Public Const titleNFeAPI = "NS API - NFe"

Function lerJson(txt As Boolean)
On Error GoTo SAI
    NFe = ""
    Select Case txt
    Case False
        Open App.Path & "\json\NFeAPIExample.json" For Input As #1
            
            NFe = input(FileLen(App.Path & "\json\exemploNFe.json"), #1)
        Close #1
    Case True
        Open App.Path & "\txt\NFe.txt" For Input As #1
            NFe = input(FileLen(App.Path & "\txt\NFe.txt"), #1)
        Close #1
    End Select
    Exit Function
SAI:
    MsgBox ("Problemas ao Ler o Arquivo JSON" & vbNewLine & Err.Description), vbInformation, titleNFeAPI
End Function


Function lerArquivo(fileName As String) As String
On Error GoTo SAI
    Open fileName For Input As #1
        lerArquivo = input(FileLen(fileName), #1)
    Close #1
    Exit Function
SAI:
    MsgBox ("Problemas ao Ler o Arquivo JSON" & vbNewLine & Err.Description), vbInformation, titleNFeAPI
End Function
