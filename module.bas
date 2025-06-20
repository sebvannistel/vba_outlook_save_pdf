' --------------------------------------------------
'
' Outlook macro to save a selected item(s) as pdf
' files on your hard-disk. You can select as many mails
' you want and hop hop hop each mails will be saved on
' your disk.
'
' Note : requires Winword (referenced by late-bindings)
'
' @see https://github.com/cavo789/vba_outlook_save_pdf
'
' --------------------------------------------------

Option Explicit

Private Const cFolder As String = "C:\Mails\"

Private objWord As Object

' --------------------------------------------------
'
' Ask the user for the folder where to store emails
'
' --------------------------------------------------
Private Function AskForTargetFolder(ByVal sTargetFolder As String) As String

    Dim dlgSaveAs As FileDialog

    sTargetFolder = Trim(sTargetFolder)

    ' Be sure that sTargetFolder is well ending by a slash
    If Not (Right(sTargetFolder, 1) = "\") Then
        sTargetFolder = sTargetFolder & "\"
    End If

    ' Already initialized before, so it's safe to just get the object
    Set dlgSaveAs = objWord.FileDialog(msoFileDialogFolderPicker)

    With dlgSaveAs
        .Title = "Select a Folder where to save emails"
        .AllowMultiSelect = False
        .InitialFileName = sTargetFolder
        .Show

        On Error Resume Next

        sTargetFolder = .SelectedItems(1)

        If Err.Number <> 0 Then
            sTargetFolder = ""
            Err.Clear
        End If

        On Error GoTo 0

    End With

    ' Be sure that sTargetFolder is well ending by a slash
    If Not (Right(sTargetFolder, 1) = "\") Then
        sTargetFolder = sTargetFolder & "\"
    End If

    AskForTargetFolder = sTargetFolder

End Function

' --------------------------------------------------
'
' Ask the user for a filename
'
' --------------------------------------------------
Private Function AskForFileName(ByVal sFileName As String) As String

    Dim dlgSaveAs As FileDialog
    Dim wResponse As VBA.VbMsgBoxResult
    Dim wPos As Integer

    Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)

    ' Set the initial location and file name for SaveAs dialog
    dlgSaveAs.InitialFileName = sFileName

    ' Show the SaveAs dialog and save the message as pdf
    If dlgSaveAs.Show = -1 Then

        sFileName = dlgSaveAs.SelectedItems(1)

        ' Verify if pdf is selected
        If Right(sFileName, 4) <> ".pdf" Then

            wResponse = MsgBox("Sorry, only saving in the pdf-format " & _
                "is supported." & vbNewLine & vbNewLine & _
                "Save as pdf instead?", vbInformation + vbOKCancel)

            If wResponse = vbCancel Then
                sFileName = ""
            ElseIf wResponse = vbOK Then
                wPos = InStrRev(sFileName, ".")
                If wPos > 0 Then
                    sFileName = Left(sFileName, wPos - 1)
                End If
                sFileName = sFileName & ".pdf"
            End If

        End If
    End If

    ' Return the filename
    AskForFileName = sFileName

End Function

' --------------------------------------------------
'
' Do the job, process every selected emails and
' export them as .pdf files.
'
'
' --------------------------------------------------
Sub SaveAsPDFfile()

    Const wdExportFormatPDF = 17
    Const wdExportOptimizeForPrint = 0
    Const wdExportAllDocument = 0
    Const wdExportDocumentContent = 0
    Const wdExportCreateNoBookmarks = 0

    Dim oSelection As Outlook.Selection
    Dim oMail As Object

    ' Use late-bindings
    Dim objDoc As Object
    Dim objFSO As Object

    Dim dlgSaveAs As FileDialog
    Dim objFDFS As FileDialogFilters
    Dim fdf As FileDialogFilter
    Dim I As Integer, wSelectedeMails As Integer
    Dim sFileName As String
    Dim sTargetFolder As String
    Dim iCount As Long

    Dim bContinue As Boolean
    Dim bAskForFileName As Boolean

    ' Get all selected items
    Set oSelection = Application.ActiveExplorer.Selection

    ' Get the number of selected emails
    wSelectedeMails = oSelection.Count

    ' Make sure at least one item is selected
    If wSelectedeMails < 1 Then
        Call MsgBox("Please select at least one email", _
            vbExclamation, "Save as PDF")
        Exit Sub
    End If

    ' --------------------------------------------------
    bContinue = MsgBox("You're about to export " & wSelectedeMails & " " & _
        "emails as PDF files, do you want to continue? If you Yes, you'll " & _
        "first need to specify the name of the folder where to store the files", _
        vbQuestion + vbYesNo + vbDefaultButton1) = vbYes

    If Not bContinue Then
        Exit Sub
    End If

    ' --------------------------------------------------
    ' Start Word and make initializations
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False

    ' --------------------------------------------------
    ' Define the target folder, where to save emails
    sTargetFolder = AskForTargetFolder(cFolder)

    If sTargetFolder = "" Then
        objWord.Quit
        Set objWord = Nothing
        Exit Sub
    End If

    ' --------------------------------------------------
    ' When more than one email has been selected, just ask the
    ' user if we need to ask for filenames each time (can be
    ' annoying)
    bAskForFileName = True

    If (wSelectedeMails > 1) Then
        bAskForFileName = MsgBox("You're about to save " & wSelectedeMails & " " & _
            "emails as PDF files. Do you want to see " & wSelectedeMails & " " & _
            "prompts so you can update the filename or just use the automated " & _
            "one (so no prompt)." & vbCrLf & vbCrLf & _
            "Press Yes to see prompts, Press No to use automated name", _
            vbQuestion + vbYesNo + vbDefaultButton2) = vbYes

        MsgBox "BE CAREFULL: You'll not see a progression on the screen (unfortunately, " & _
            "Outlook doesn't allow this)." & vbCrLf & vbCrLf & _
            "If you're exporting a lot of mails, the process can take a while. " & _
            "Perhaps the best way to see that things are working is to open a " & _
            "explorer window and see how files are added to the folder." & vbCrLf & vbCrLf & _
            "Once finished, you'll see a feedback message.", _
            vbInformation + vbOKOnly
    End If

    ' --------------------------------------------------
    ' Define the SaveAs dialog
    If bAskForFileName Then

        Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)

        ' --------------------------------------------------
        ' Determine the FilterIndex for saving as a pdf-file
        ' Get all the filters and make sure we've "pdf"
        Set objFDFS = dlgSaveAs.Filters

        I = 0

        For Each fdf In objFDFS
            I = I + 1

            If InStr(1, fdf.Extensions, "pdf", vbTextCompare) > 0 Then
                Exit For
            End If
        Next fdf

        Set objFDFS = Nothing

        ' Set the FilterIndex to pdf-files
        dlgSaveAs.FilterIndex = I

    End If

    ' ----------------------------------------------------
    ' Initialize file system for folder and file checks
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    If Not objFSO.FolderExists(sTargetFolder) Then
        objFSO.CreateFolder sTargetFolder
    End If

    ' ----------------------------------------------------
    ' We are ready to start
    For Each oMail In oSelection

            Const MAX_PATH As Long = 260              'official value

            '1. Trim quoted thread, render, and grab live Word.Document
            oMail.HTMLBody = StripQuotedBody(oMail)    'update body in place
            oMail.Display False                        'forces Word to build editor
            Set objDoc = oMail.GetInspector.WordEditor

            '2.  Build the unique PDF name
            Dim base$, try$, dup&, room&
            base = sTargetFolder & Format(ItemDate(oMail), "yyyymmdd-hhnnss") _
                   & " – " & CleanSubject(oMail.Subject)

            room = MAX_PATH - Len(sTargetFolder) - 5
            If Len(base) > room Then base = Left$(base, room)

            try = base & ".pdf": dup = 1
            Do While objFSO.FileExists(try)
                try = base & "_" & dup & ".pdf": dup = dup + 1
            Loop

            If bAskForFileName Then sFileName = AskForFileName(try) Else sFileName = try

            '3.  Export and close
            If Len(Trim$(sFileName)) > 0 Then
                objDoc.ExportAsFixedFormat _
                    OutputFileName:=sFileName, _
                    ExportFormat:=wdExportFormatPDF, _
                    OptimizeFor:=wdExportOptimizeForPrint, _
                    Range:=wdExportAllDocument, _
                    Item:=wdExportDocumentContent, _
                    CreateBookmarks:=wdExportCreateNoBookmarks
            End If

            iCount = iCount + 1
            If iCount Mod 50 = 0 Then DoEvents

    Next oMail

    Set dlgSaveAs = Nothing

    On Error GoTo 0

    ' Close the document and Word

    On Error Resume Next
    objWord.Quit
    On Error GoTo 0

    ' Cleanup

    Set oSelection = Nothing
    Set oMail = Nothing
    Set objDoc = Nothing
    Set objWord = Nothing
    Set objFSO = Nothing

    MsgBox "Done, mails have been exported to " & sTargetFolder, vbInformation

End Sub

Private Function CleanSubject(raw As String) As String
    Static rePfx As Object, reBad As Object

    If rePfx Is Nothing Then
        '--- 1) prefix stripper ------------------------------------
        Set rePfx = CreateObject("VBScript.RegExp")
        With rePfx
            .Global = True            'remove every prefix found
            .IgnoreCase = True        'case-insensitive match
            .Pattern = "^\s*((re|fw|fwd)\s*:)+\s*"
        End With

        '--- 2) illegal Windows-filename chars ---------------------
        Set reBad = CreateObject("VBScript.RegExp")
        With reBad
            .Global = True
            .Pattern = "[\\/:*?""<>|]"
        End With
    End If

    'Guard against NULL subjects
    CleanSubject = Trim$(reBad.Replace(rePfx.Replace(CStr(raw), ""), ""))
End Function

'========== 1) universal time getter ==========
Private Function ItemDate(itm As Object) As Date
    Select Case True
        Case TypeOf itm Is Outlook.MailItem
            If itm.ReceivedTime = #1/1/4501# Then
                ItemDate = itm.SentOn
            Else
                ItemDate = itm.ReceivedTime
            End If
        Case TypeOf itm Is Outlook.ReportItem
            ItemDate = itm.CreationTime
        Case Else
            ItemDate = Now
    End Select
End Function

Function StripQuotedBody(mi As Outlook.MailItem) As String
    Dim html$, pos&
    html = mi.HTMLBody
    If Len(html) = 0 Then html = Replace(mi.Body, vbCrLf, "<br>")

    'Common Outlook marker for the reply header
    pos = InStr(html, "class=""OutlookMessageHeader""")
    If pos = 0 Then pos = InStr(html, "-----Original Message-----")
    If pos = 0 Then pos = InStr(html, "<hr")              'fallback

    If pos > 0 Then html = Left$(html, pos - 1) & "</body></html>"
    StripQuotedBody = html
End Function
