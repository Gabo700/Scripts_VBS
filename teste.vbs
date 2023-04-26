Option Explicit

Dim objFSO, strPasta, strArquivo, datDataArquivo, intMesArquivo, intMesAnterior

Set objFSO = CreateObject("Scripting.FileSystemObject")
strPasta = "C:\Users\gabriel.ribeiro\Downloads"
intMesAnterior = Month(Date) - 1 ' mês atual menos 1 
If intMesAnterior = 0 Then
    intMesAnterior = 12 ' caso seja dezembro
End If

For Each strArquivo In objFSO.GetFolder(strPasta).Files
    datDataArquivo = objFSO.GetFile(strArquivo).DateLastModified ' verificando ultima modificação
    intMesArquivo = Month(datDataArquivo)
    If intMesArquivo = intMesAnterior Then
        objFSO.DeleteFile(strArquivo), True
    End If
Next

'  Em resumo apagando o mês anterios mesmo assim sera nescessario agendar tarefa. 