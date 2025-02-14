Attribute VB_Name = "Module1"
Sub ConvertToTextFileUTF8()
    Dim tempFileName As String
    Dim textFilePath As String
    Dim htmlContent As String
    Dim currentDirectory As String
    Dim fileNum As Integer
    Dim fso As Object
    Dim textFile As Object

    ' ���� ���ј���� ���� ���� ���
    currentDirectory = ActiveDocument.Path

    ' ǐ� ���� ���� ��ϡ �� ���� temp ����� ������
    If currentDirectory = "" Then
        currentDirectory = Environ("TEMP")
    End If

    ' ���� ���� ���� �� ����� �흘���
    textFilePath = currentDirectory & "\outputfile.txt"

    ' ����� ���� �� ���� HTML
    Application.DisplayAlerts = wdAlertsNone
    ActiveDocument.SaveAs2 FileName:=currentDirectory & "\tempfile.html", FileFormat:=wdFormatFilteredHTML
    Application.DisplayAlerts = wdAlertsAll

    ' ������ ����� �� ���� HTML �� ���� ����
    On Error GoTo CleanupError
    fileNum = FreeFile
    Open currentDirectory & "\tempfile.html" For Input As fileNum
    htmlContent = Input$(LOF(fileNum), fileNum)
    Close fileNum

    ' ����� ��� FileSystemObject ���� ����� �� ���� �� ��Ґ���� UTF-8
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set textFile = fso.CreateTextFile(textFilePath, True, True)  ' True ���� UTF-8

    ' ����� ����� �� ���� ����
    textFile.Write htmlContent
    textFile.Close

    ' ��� ���� ���� HTML
    On Error Resume Next
    Kill currentDirectory & "\tempfile.html"
    On Error GoTo 0

    MsgBox "����� �� ���� ���� ����� ��: " & textFilePath, vbInformation
    Exit Sub

CleanupError:
    MsgBox "��� �� ������ �� ����� ����: " & Err.Description, vbExclamation
    On Error Resume Next
    Close fileNum
End Sub

