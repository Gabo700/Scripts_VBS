
Option Explicit

Dim objFSO, strPasta, strArquivo, datDataArquivo, datLimite

Set objFSO = CreateObject("Scripting.FileSystemObject")
strPasta = "C:\Users\gabriel.ribeiro\Documents"
datLimite = DateAdd("d", -14, Date) ' subtrai 14 dias da data atual.

For Each strArquivo In objFSO.GetFolder(strPasta).Files
    datDataArquivo = objFSO.GetFile(strArquivo).DateLastModified
    If datDataArquivo < datLimite Then ' verifica se o arquivo foi modificado há mais de 14 dias.
        objFSO.DeleteFile(strArquivo), True
    End If
Next

'Wesley em qualquer um dos casos seja o de 14 dias ou de 1 mês será nescessário agendar tarefa no windows server. O espaço é liberado para sobrescrição sem passar por nenhum gerenciador (lixeira).
