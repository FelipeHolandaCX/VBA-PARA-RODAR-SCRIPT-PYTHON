#Formulário

Private Sub Form_Open(Cancel As Integer)
    ' Executar o script Python
    ExecutarScriptPython

    ' Minimizar o Access
    DoCmd.RunCommand acCmdAppMinimize

    ' Fechar o Access após a execução do script
    Application.Quit
End Sub

#Módulo

Public Sub ExecutarScriptPython()
    Dim strPythonPath As String
    Dim strScriptPath As String
    Dim strCommand As String
    Dim shellResult As Double
    
    ' Caminho para o executável do Python
    strPythonPath = "C:\Users\PC GAMER\AppData\Local\Programs\Python\Python312\python.exe"
    
    ' Caminho para o script Python na mesma pasta do arquivo Access
    strScriptPath = CurrentProject.Path & "\main.py"
    
    ' Verifique se o executável do Python existe
    If Dir(strPythonPath) = "" Then
        MsgBox "O executável do Python não foi encontrado no caminho especificado: " & strPythonPath, vbCritical
        Exit Sub
    End If
    
    ' Verifique se o script Python existe
    If Dir(strScriptPath) = "" Then
        MsgBox "O script Python não foi encontrado no caminho especificado: " & strScriptPath, vbCritical
        Exit Sub
    End If
    
    ' Comando para executar o script Python
    strCommand = """" & strPythonPath & """ """ & strScriptPath & """"
    
    ' Executar o script Python
    shellResult = Shell(strCommand, vbNormalFocus)
    
End Sub
