Attribute VB_Name = "Module1"
Sub ConvertToTextFileUTF8()
    Dim tempFileName As String
    Dim textFilePath As String
    Dim htmlContent As String
    Dim currentDirectory As String
    Dim fileNum As Integer
    Dim fso As Object
    Dim textFile As Object

    ' „”Ì— œ«Ì—ò Ê—Ì ›⁄·Ì ›«Ì· Ê—œ
    currentDirectory = ActiveDocument.Path

    ' «ê— „”Ì— Œ«·Ì »Êœ° »Â ÅÊ‘Â temp  €ÌÌ— „ÌùœÂÌ„
    If currentDirectory = "" Then
        currentDirectory = Environ("TEMP")
    End If

    ' „”Ì— ›«Ì· „ ‰Ì —«  ‰ŸÌ„ „Ìùò‰Ì„
    textFilePath = currentDirectory & "\outputfile.txt"

    ' –ŒÌ—Â ›«Ì· »Â ›—„  HTML
    Application.DisplayAlerts = wdAlertsNone
    ActiveDocument.SaveAs2 FileName:=currentDirectory & "\tempfile.html", FileFormat:=wdFormatFilteredHTML
    Application.DisplayAlerts = wdAlertsAll

    ' ŒÊ«‰œ‰ „Õ Ê« «“ ›«Ì· HTML »Â ’Ê—  —‘ Â
    On Error GoTo CleanupError
    fileNum = FreeFile
    Open currentDirectory & "\tempfile.html" For Input As fileNum
    htmlContent = Input$(LOF(fileNum), fileNum)
    Close fileNum

    ' «ÌÃ«œ ‘Ì¡ FileSystemObject »—«Ì ‰Ê‘ ‰ »Â ›«Ì· »« —„“ê–«—Ì UTF-8
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set textFile = fso.CreateTextFile(textFilePath, True, True)  ' True »—«Ì UTF-8

    ' ‰Ê‘ ‰ „Õ Ê« œ— ›«Ì· „ ‰Ì
    textFile.Write htmlContent
    textFile.Close

    ' Õ–› ›«Ì· „Êﬁ  HTML
    On Error Resume Next
    Kill currentDirectory & "\tempfile.html"
    On Error GoTo 0

    MsgBox "„Õ Ê« œ— ›«Ì· „ ‰Ì –ŒÌ—Â ‘œ: " & textFilePath, vbInformation
    Exit Sub

CleanupError:
    MsgBox "Œÿ« œ— ŒÊ«‰œ‰ Ì« ‰Ê‘ ‰ ›«Ì·: " & Err.Description, vbExclamation
    On Error Resume Next
    Close fileNum
End Sub

