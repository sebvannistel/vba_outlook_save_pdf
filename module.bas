#If VBA7 Then
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As LongPtr) As LongPtr
#Else
Private Declare Function SetForegroundWindow Lib "user32" _
        (ByVal hWnd As Long) As Long
#End If

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

Private Const wdExportFormatPDF As Long = 17     'moved to module level
Private Const olMHTML As Long = 10               'Added for late-binding
Private Const olMSG As Long = 3                  'Added for late-binding
Private Const olUnrestricted As Long = 0         'NEW: For version-independent IRM/RMS check

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
        
        ' *** UPDATE: Make the folder dialog stay on top ***
        SetForegroundWindow objWord.Hwnd
        
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
'Sub SaveAsPDFfile()
'
'    Const wdExportOptimizeForPrint = 0
'    Const wdExportAllDocument = 0
'    Const wdExportDocumentContent = 0
'    Const wdExportCreateNoBookmarks = 0
'
'    Dim oSelection As Outlook.Selection
'    Dim oMail As Object
'
'    ' Use late-bindings
'    Dim objDoc As Object
'    Dim objFSO As Object
'
'    Dim dlgSaveAs As FileDialog
'    Dim objFDFS As FileDialogFilters
'    Dim fdf As FileDialogFilter
'    Dim I As Integer, wSelectedeMails As Integer
'    Dim sFileName As String
'    Dim sTargetFolder As String
'    Dim iCount As Long
'
'    Dim bContinue As Boolean
'    Dim bAskForFileName As Boolean
'
'    ' Get all selected items
'    Set oSelection = Application.ActiveExplorer.Selection
'
'    ' Get the number of selected emails
'    wSelectedeMails = oSelection.Count
'
'    ' Make sure at least one item is selected
'    If wSelectedeMails < 1 Then
'        Call MsgBox("Please select at least one email", _
'            vbExclamation, "Save as PDF")
'        Exit Sub
'    End If
'
'    ' --------------------------------------------------
'    bContinue = MsgBox("You're about to export " & wSelectedeMails & " " & _
'        "emails as PDF files, do you want to continue? If you Yes, you'll " & _
'        "first need to specify the name of the folder where to store the files", _
'        vbQuestion + vbYesNo + vbDefaultButton1) = vbYes
'
'    If Not bContinue Then
'        Exit Sub
'    End If
'
'    ' --------------------------------------------------
'    ' Start Word and make initializations
'    Set objWord = CreateObject("Word.Application")
'    objWord.Visible = False
'
'    ' --------------------------------------------------
'    ' Define the target folder, where to save emails
'    sTargetFolder = AskForTargetFolder(cFolder)
'
'    If sTargetFolder = "" Then
'        objWord.Quit
'        Set objWord = Nothing
'        Exit Sub
'    End If
'
'    ' --------------------------------------------------
'    ' When more than one email has been selected, just ask the
'    ' user if we need to ask for filenames each time (can be
'    ' annoying)
'    bAskForFileName = True
'
'    If (wSelectedeMails > 1) Then
'        bAskForFileName = MsgBox("You're about to save " & wSelectedeMails & " " & _
'            "emails as PDF files. Do you want to see " & wSelectedeMails & " " & _
'            "prompts so you can update the filename or just use the automated " & _
'            "one (so no prompt)." & vbCrLf & vbCrLf & _
'            "Press Yes to see prompts, Press No to use automated name", _
'            vbQuestion + vbYesNo + vbDefaultButton2) = vbYes
'
'        MsgBox "BE CAREFULL: You'll not see a progression on the screen (unfortunately, " & _
'            "Outlook doesn't allow this)." & vbCrLf & vbCrLf & _
'            "If you're exporting a lot of mails, the process can take a while. " & _
'            "Perhaps the best way to see that things are working is to open a " & _
'            "explorer window and see how files are added to the folder." & vbCrLf & vbCrLf & _
'            "Once finished, you'll see a feedback message.", _
'            vbInformation + vbOKOnly
'    End If
'
'    ' --------------------------------------------------
'    ' Define the SaveAs dialog
'    If bAskForFileName Then
'
'        Set dlgSaveAs = objWord.FileDialog(msoFileDialogSaveAs)
'
'        ' --------------------------------------------------
'        ' Determine the FilterIndex for saving as a pdf-file
'        ' Get all the filters and make sure we've "pdf"
'        Set objFDFS = dlgSaveAs.Filters
'
'        I = 0
'
'        For Each fdf In objFDFS
'            I = I + 1
'
'            If InStr(1, fdf.Extensions, "pdf", vbTextCompare) > 0 Then
'                Exit For
'            End If
'        Next fdf
'
'        Set objFDFS = Nothing
'
'        ' Set the FilterIndex to pdf-files
'        dlgSaveAs.FilterIndex = I
'
'    End If
'
'    ' ----------------------------------------------------
'    ' Initialize file system for folder and file checks
'    Set objFSO = CreateObject("Scripting.FileSystemObject")
'
'    If Not objFSO.FolderExists(sTargetFolder) Then
'        objFSO.CreateFolder sTargetFolder
'    End If
'
'    ' ----------------------------------------------------
'    ' We are ready to start
'    For Each oMail In oSelection
'
'            Const MAX_PATH As Long = 260              'official value
'
'            '1. Render and grab live Word.Document
'            oMail.Display False                        'forces Word to build editor
'            Set objDoc = oMail.GetInspector.WordEditor
'
'            '2.  Build the unique PDF name
'            Dim base$, try$, dup&, room&
'            base = sTargetFolder & Format(ItemDate(oMail), "yyyymmdd-hhnnss") _
'                   & " – " & CleanSubject(oMail.Subject)
'
'            room = MAX_PATH - Len(sTargetFolder) - 5
'            If Len(base) > room Then base = Left$(base, room)
'
'            try = base & ".pdf": dup = 1
'            Do While objFSO.FileExists(try)
'                try = base & "_" & dup & ".pdf": dup = dup + 1
'            Loop
'
'            If bAskForFileName Then sFileName = AskForFileName(try) Else sFileName = try
'
'            '3.  Export and close
'            If Len(Trim$(sFileName)) > 0 Then
'                objDoc.ExportAsFixedFormat _
'                    OutputFileName:=sFileName, _
'                    ExportFormat:=wdExportFormatPDF, _
'                    OptimizeFor:=wdExportOptimizeForPrint, _
'                    Range:=wdExportAllDocument, _
'                    Item:=wdExportDocumentContent, _
'                    CreateBookmarks:=wdExportCreateNoBookmarks
'            End If
'
'            iCount = iCount + 1
'            If iCount Mod 50 = 0 Then DoEvents
'
'    Next oMail
'
'    Set dlgSaveAs = Nothing
'
'    On Error GoTo 0
'
'    ' Close the document and Word
'
'    On Error Resume Next
'    objWord.Quit
'    On Error GoTo 0
'
'    ' Cleanup
'
'    Set oSelection = Nothing
'    Set oMail = Nothing
'    Set objDoc = Nothing
'    Set objWord = Nothing
'    Set objFSO = Nothing
'
'    MsgBox "Done, mails have been exported to " & sTargetFolder, vbInformation
'
'End Sub

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


' --------------------------------------------------
' Export selected mails as PDFs entirely in the background
' No Inspector windows will open and the original message
' remains untouched.
'Sub SaveSelectedMails_AsPDF_NoPopups()
'
'    Const tempExtMHT = ".mht"
'
'    Dim sel As Outlook.Selection
'    Dim mi As Outlook.MailItem
'    Dim objWord As Object, doc As Object
'    Dim tmpFile As String, pdfFile As String
'    Dim fso As Object
'
'    Set sel = Application.ActiveExplorer.Selection
'    If sel.Count = 0 Then
'        MsgBox "Nothing selected"
'        Exit Sub
'    End If
'
'    Set objWord = CreateObject("Word.Application")
'    objWord.Visible = False
'    Set fso = CreateObject("Scripting.FileSystemObject")
'
'    For Each mi In sel
'
'        '1. build filenames
'        tmpFile = Environ$("TEMP") & "\" & _
'                  Format(mi.ReceivedTime, "yyyymmdd-hhnnss") & "_" & _
'                  CleanFile(mi.Subject) & tempExtMHT
'
'        pdfFile = cFolder & _
'                  Format(mi.ReceivedTime, "yyyymmdd-hhnnss") & " – " & _
'                  CleanFile(mi.Subject) & ".pdf"
'
'        '2. save e-mail as an MHT without opening an Inspector
'        mi.SaveAs tmpFile, olMHTML
'
'        '3. let Word convert that file straight to PDF
'        Set doc = objWord.Documents.Open(tmpFile, ReadOnly:=True, Visible:=False)
'        doc.ExportAsFixedFormat pdfFile, wdExportFormatPDF
'        doc.Close False
'
'        fso.DeleteFile tmpFile
'
'    Next mi
'
'    objWord.Quit
'    MsgBox sel.Count & " mail(s) exported."
'End Sub

Private Function CleanFile(s As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In badChars
        s = Replace$(s, ch, " ")
    Next ch
    CleanFile = Trim$(s)
End Function

' *** NEW FUNCTION: Logs skipped items to a text file for review. ***
Private Sub LogSkippedItem(ByVal logPath As String, ByVal itemSubject As String, ByVal reason As String)
    If Len(logPath) > 0 Then
        Dim fileNum As Integer
        fileNum = FreeFile
        On Error Resume Next ' In case of file system errors (e.g., folder not accessible)
        Open logPath For Append As #fileNum
        If Err.Number = 0 Then
            Print #fileNum, CStr(Now) & " | SKIPPED: """ & itemSubject & """ | Reason: " & reason
            Close #fileNum
        End If
        On Error GoTo 0
    End If
End Sub


'Save selected Outlook messages as PDFs – quiet & with headers
Sub SaveMails_ToPDF_Background()

    Const tmpExt As String = ".mht"
    Const olMail As Long = 43 ' Added for late-binding check

    Dim sel As Outlook.Selection
    Dim mi  As Object ' Use generic object for the loop to handle non-mail items
    Dim mailItem As Outlook.MailItem ' Use a specific variable after type check
    Dim wrd As Object, doc As Object
    Dim tmpFile As String, pdfFile As String, tgtFolder As String
    Dim fso As Object, hdr As String, logFilePath As String ' Added logFilePath
    Dim attCnt As Long ' *** ADDED: For safe attachment count checking ***

    'Pick a target folder (no hard-coded C:\Mails)
    'NOTE: This will call your existing 'AskForTargetFolder' function
    Set wrd = CreateObject("Word.Application")
    Set objWord = wrd      ' sync the global variable for AskForTargetFolder
    wrd.Visible = False
    tgtFolder = AskForTargetFolder("")

    Set sel = Application.ActiveExplorer.Selection
    If sel.Count = 0 Or Len(tgtFolder) = 0 Then
        wrd.Quit
        Set wrd = Nothing      ' FIX: Ensure Word object is released on early exit
        Set objWord = Nothing  ' FIX: Ensure global Word object is also released
        Exit Sub
    End If

    ' *** UPDATE: Define the path for the log file ***
    logFilePath = tgtFolder & "_SkippedItems.log"

    wrd.DisplayAlerts = 0             'wdAlertsNone
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim total As Long
    Dim done  As Long
    Dim showProgress As Boolean
    total = sel.Count
    showProgress = (total > 100)

    For Each mi In sel

        ' FIX: Check if the selected item is a genuine MailItem before processing.
        ' This prevents errors on items like meeting requests or reports.
        If TypeOf mi Is Outlook.MailItem Then
            Set mailItem = mi

            ' --- UPDATE 2.1 & 2.3: Additional guards for item validity ---
            If mailItem.Class <> olMail Then
                LogSkippedItem logFilePath, mailItem.Subject, "Not a true MailItem (e.g., Report, Meeting Invite)"
                GoTo NextItemInLoop
            End If
            If mailItem.Size = 0 Then
                LogSkippedItem logFilePath, mailItem.Subject, "Item size is 0 (header only, not downloaded)"
                GoTo NextItemInLoop
            End If

            ' --- START: UPDATES BASED ON YOUR INSTRUCTIONS ---
            
            ' UPDATE (Instruction 2): A version-independent way to detect protected mail.
            ' This replaces the newer .IsRestricted property with the backward-compatible .Permission property.
            If mailItem.Permission <> olUnrestricted Then
                LogSkippedItem logFilePath, mailItem.Subject, "Item is protected by Information Rights Management (RMS/IRM)"
                GoTo NextItemInLoop
            End If

            ' *** FIX: Harden S/MIME check to prevent "Array index out-of-bounds" error. ***
            ' The 'And' operator in VBA is not short-circuited, so both conditions were
            ' always evaluated, causing a crash if there were no attachments.
            ' This nested check is safe.
            attCnt = mailItem.Attachments.Count
            If attCnt = 1 Then
                If LCase$(mailItem.Attachments(1).FileName) Like "*.p7m" Then
                    LogSkippedItem logFilePath, mailItem.Subject, _
                        "Item is an S/MIME encrypted message (.p7m attachment)"
                    GoTo NextItemInLoop
                End If
            End If

            ' --- END: UPDATES BASED ON YOUR INSTRUCTIONS ---


            '--- 1  FIX A: Harden the filename builder to prevent MAX_PATH errors ---
            Const MAX_PATH As Long = 260
            Dim safeSubj As String, datePrefix As String, room As Long
            Dim roomForTmp As Long, roomForPdf As Long

            ' Sanitize subject line to remove illegal filename characters
            safeSubj = CleanFile(mailItem.Subject)
            ' --- UPDATE 2.4: Use ItemDate helper to handle drafts ---
            datePrefix = Format(ItemDate(mailItem), "yyyymmdd-hhnnss")

            ' --- UPDATE 2.2: Adjust for null terminator (MAX_PATH - 1) ---
            ' Calculate the maximum allowed length for the subject part to avoid
            ' exceeding MAX_PATH for both the temporary and the final PDF file.
            roomForTmp = (MAX_PATH - 1) - (Len(Environ$("TEMP") & "\") + Len(datePrefix) + Len("_") + Len(tmpExt))
            roomForPdf = (MAX_PATH - 1) - (Len(tgtFolder) + Len(datePrefix) + Len(" – ") + Len(".pdf"))

            ' Use the more restrictive of the two calculated lengths
            If roomForTmp < roomForPdf Then room = roomForTmp Else room = roomForPdf
            If room < 10 Then room = 10 ' Ensure a minimum filename length

            ' Truncate the sanitized subject if it's too long
            If Len(safeSubj) > room Then safeSubj = Left$(safeSubj, room)

            ' Build the final, safe filenames
            tmpFile = Environ$("TEMP") & "\" & datePrefix & "_" & safeSubj & tmpExt
            pdfFile = tgtFolder & datePrefix & " – " & safeSubj & ".pdf"


            '--- 2  FIX B: Wrap SaveAs with a fallback to handle errors gracefully ---
            On Error Resume Next
            mailItem.SaveAs tmpFile, olMHTML
            If Err.Number <> 0 Then
                ' MHTML save failed. This can be due to retention policies, sync issues, etc.
                ' Clear the error and attempt the fallback action.
                Err.Clear

                ' *** UPDATE: Log the failure before attempting fallback ***
                LogSkippedItem logFilePath, mailItem.Subject, "Failed to save as MHTML, attempting MSG fallback"

                ' *** UPDATE: Stop read-only errors on fallback save ***
                Dim msgFallbackFile As String
                msgFallbackFile = Replace(pdfFile, ".pdf", ".msg")
                
                If fso.FileExists(msgFallbackFile) Then
                    On Error Resume Next
                    SetAttr msgFallbackFile, vbNormal    'clear RO flag
                    fso.DeleteFile msgFallbackFile, True 'hard delete
                    On Error GoTo 0
                End If
                
                mailItem.SaveAs msgFallbackFile, 65      '65 = olMSGUnicode, safer for accents (as per update request)

                ' Since MHTML creation failed, we cannot create a PDF. Skip to the next item.
                GoTo NextItemInLoop
            End If
            ' If we got here, SaveAs MHT succeeded. Restore normal error handling.
            On Error GoTo 0


            '--- 3  open in Word and prepend header -------------------
            Set doc = wrd.Documents.Open(tmpFile, ReadOnly:=True, Visible:=False)

            ' *** UPDATE: Avoid inserting a double header ***
            If Not Left$(doc.Range(0, 6).Text, 5) = "From:" Then
                hdr = "From:    " & mailItem.SenderName & vbCrLf & _
                      "Sent:    " & mailItem.SentOn & vbCrLf & _
                      "To:      " & mailItem.To & vbCrLf & _
                      "CC:      " & mailItem.CC & vbCrLf & _
                      "BCC:     " & mailItem.BCC & vbCrLf & _
                      "Subject: " & mailItem.Subject & vbCrLf & _
                      String(60, "—") & vbCrLf & vbCrLf
                doc.Range.InsertBefore hdr
            End If

            '--- 4  export to PDF & clean up --------------------------

            ' *** FIX: Ensure the target file is writable before exporting. ***
            ' This prevents "file is read-only" errors by clearing the attribute
            ' and deleting any existing locked file left from a previous failed run.
            If fso.FileExists(pdfFile) Then
                On Error Resume Next
                SetAttr pdfFile, vbNormal      ' Clear read-only attribute
                fso.DeleteFile pdfFile, True   ' Delete the file to ensure a clean save
                On Error GoTo 0
            End If

            ' *** UPDATE: Use explicit OptimizeFor:=0 parameter ***
            doc.ExportAsFixedFormat OutputFileName:=pdfFile, ExportFormat:=wdExportFormatPDF, OptimizeFor:=0
            doc.Close False
            
            ' *** FIX: This is the key change. By setting doc to Nothing immediately after closing, ***
            ' *** we prevent the cleanup block at 'NextItemInLoop' from trying to close it again. ***
            Set doc = Nothing
            
            fso.DeleteFile tmpFile

            done = done + 1
            If showProgress Then
                Application.StatusBar = "Saving mail " & done & " of " & total & "..."
                If done Mod 25 = 0 Then DoEvents
            End If
        Else ' Item in selection is not a MailItem
             ' *** UPDATE: Log non-mail items that are skipped ***
             Dim itemType As String
             itemType = TypeName(mi)
             LogSkippedItem logFilePath, "Unknown Subject (Type: " & itemType & ")", "Item is not an email (e.g., a " & itemType & ")"
        End If ' End of "If TypeOf mi Is Outlook.MailItem" check

NextItemInLoop:
        ' This label is the target for the GoTo statement when an error occurs.
        ' It ensures the loop continues with the next item.
        ' --- UPDATE 3.1: Clean up Word document object if it exists before next loop ---
        ' This block now safely handles cleanup for items that failed mid-process,
        ' while the 'Set doc = Nothing' change above prevents it from causing an
        ' error on successfully processed items.
        If Not doc Is Nothing Then doc.Close False
        Set doc = Nothing
    Next mi

    If showProgress Then Application.StatusBar = False
    wrd.Quit
    
    ' *** FIX: Added full cleanup block to explicitly release all COM objects at the end, as per best practices. ***
    Set doc = Nothing
    Set wrd = Nothing
    Set objWord = Nothing
    Set fso = Nothing
    Set sel = Nothing
    Set mi = Nothing
    Set mailItem = Nothing

    MsgBox done & " mail(s) saved as PDF to " & tgtFolder & vbCrLf & vbCrLf & _
           "A log of any skipped items has been saved to _SkippedItems.log", vbInformation

End Sub
