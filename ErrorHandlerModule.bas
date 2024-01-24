Attribute VB_Name = "ErrorHandlerModule"
Public Sub ReportError(Modulo, Procedimento, Numero, Descricao, Fonte, Linha)
Dim r As Integer
Dim errStr As String
Dim msg as String
Dim resposta as Integer
msg ="Oops...surgiu um erro inesperado!" & Chr(13) & "Está a ser criado um ficheiro de despistagem contendo informações que podem ajudar a compreender e resolver o problema. O ficheiro de logo está a ser criado em:" & App.Path & chr(13) & chr(13)
msg = msg & "Pretende continuar?"
resposta = MsgBox(msg, vbYesNo + vbQuestion, "Error Detection")
r = FreeFile
errStr = Now & vbnewline & "Module: " & Modulo & "; Procedure: " & Procedimento & "; Erro: " & Numero & " - " & Descricao & "; Line: " & Linha & "; Source: " & Fonte & vbnewline
Open App.Path & "\errlog.txt" For Append As r
Print #r, errStr
Close #r
if resposta <> 6 then end
End Sub
