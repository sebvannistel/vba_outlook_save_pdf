'#If VBA7 Then
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
'#Else
'    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
'    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
'    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'#End If
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
Private Const wdOpenFormatWebPages As Long = 7

Private objWord As Object

' =========================================================================================
' === STABILITY FIX: Use Outlook's native File Dialog instead of Word's.             ===
' === This avoids cross-application UI issues and the need for API calls.             ===
' =========================================================================================
Private Function AskForTargetFolder(ByVal sInitialFolder As String) As String
    Dim dlg As FileDialog
    Dim selectedFolder As String

    ' Use Outlook's Application.FileDialog - it's 100% stable
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)

    With dlg
        .Title = "Select a Folder where to save emails"
        .AllowMultiSelect = False
        If Len(sInitialFolder) > 0 Then
            .InitialFileName = sInitialFolder
        End If

        ' Show the dialog. If the user clicks OK, .Show returns -1
        If .Show = -1 Then
            selectedFolder = .SelectedItems(1)
            ' Ensure the path ends with a backslash
            If Right(selectedFolder, 1) <> "\" Then
                selectedFolder = selectedFolder & "\"
            End If
        Else
            ' User cancelled
            selectedFolder = ""
        End If
    End With

    AskForTargetFolder = selectedFolder
    Set dlg = Nothing
End Function


' --------------------------------------------------
'
' Ask the user for a filename
' (CRITICAL RISK FIX: Re-written to be stable)
'
' --------------------------------------------------
Private Function AskForFileName(ByVal sFileName As String) As String
    ' STABILITY FIX: This function has been re-written to use Outlook's native
    ' Application.FileDialog. This is 100% stable and avoids using Word's UI
    ' and the fragile SetForegroundWindow API call that caused crashes.

    Dim dlg As FileDialog
    Dim sSelectedFile As String

    ' Use Outlook's Application.FileDialog for stability
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)

    With dlg
        .Title = "Save As PDF"
        .InitialFileName = sFileName ' Use the suggested filename

        ' Clear existing filters and add a specific PDF filter
        .Filters.Clear
        .Filters.Add "PDF Files", "*.pdf"
        .FilterIndex = 1 ' Make PDF the default and only option

        ' Show the dialog. If the user clicks Save, .Show returns -1
        If .Show = -1 Then
            sSelectedFile = .SelectedItems(1)

            ' The SaveAs dialog usually appends the extension automatically if the
            ' filter is set, but we can double-check to be certain.
            If LCase(Right(sSelectedFile, 4)) <> ".pdf" Then
                sSelectedFile = sSelectedFile & ".pdf"
            End If
        Else
            ' User cancelled the dialog
            sSelectedFile = ""
        End If
    End With

    ' Return the selected filename or an empty string if cancelled
    AskForFileName = sSelectedFile

    Set dlg = Nothing
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
'            Const cMAX_PATH As Long = 260              'official value
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
'            room = cMAX_PATH - Len(sTargetFolder) - 5
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
    On Error Resume Next
    Select Case True
        Case TypeOf itm Is Outlook.MailItem
            If itm.ReceivedTime < #1/2/1900# Then 'Check for very old or invalid dates
                ItemDate = itm.SentOn
            Else
                ItemDate = itm.ReceivedTime
            End If
        Case TypeOf itm Is Outlook.ReportItem
            ItemDate = itm.CreationTime
        Case Else
            ItemDate = Now
    End Select
    If Err.Number <> 0 Then ItemDate = Now ' Fallback for any error
    On Error GoTo 0
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
'    objWord. Quit
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

'--- HELPER: Injects a simple header that looks like Outlook's print style ---
Private Sub InjectSimpleHeader(doc As Object, m As Outlook.MailItem)
    On Error Resume Next ' In case a property is not available
    Dim hdr As String
    hdr = "From: " & m.SenderName & vbCrLf & _
          "Sent: " & m.SentOn & vbCrLf & _
          "To: " & m.To & vbCrLf & _
          IIf(Len(m.CC) > 0, "Cc: " & m.CC & vbCrLf, "") & _
          "Subject: " & m.Subject & vbCrLf & _
          String(60, "—") & vbCrLf & vbCrLf
    doc.Range.InsertBefore hdr
    On Error GoTo 0
End Sub

'--- NEW HELPER (AS PER FIX): Injects a full header with a duplicate guard ---
Private Sub InjectFullHeader(doc As Object, m As Outlook.MailItem)
    On Error Resume Next ' In case a property is not available
    Dim hdr As String
    hdr = "From: " & m.SenderName & vbCrLf & _
          "Sent: " & m.SentOn & vbCrLf & _
          "To:   " & m.To & vbCrLf & _
          IIf(Len(m.CC) > 0, "Cc:   " & m.CC & vbCrLf, "") & _
          IIf(Len(m.BCC) > 0, "Bcc:  " & m.BCC & vbCrLf, "") & _
          "Subject: " & m.Subject & vbCrLf & _
          String(60, "—") & vbCrLf & vbCrLf
          
    ' Per fix, only insert header if one doesn't already exist
    If InStr(1, doc.Range(0, 60).Text, "From:") = 0 Then
        doc.Range.InsertBefore hdr
    End If
    On Error GoTo 0
End Sub

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

'--- helper: always create a unique temp MHT name
Private Function GetUniqueTempMHT(mi As Outlook.MailItem, ext As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base$, try$
    base = Environ$("TEMP") & "\" & Format(ItemDate(mi), "yyyymmdd-hhnnss") _
           & "_" & CleanFile(mi.Subject)
    Do
        ' Timer can be 0 at midnight; add Rnd for more robustness
        try = Left$(base, 200) & "_" & Hex((Timer * 1000) + (Rnd * 1000)) & ext
    Loop While fso.FileExists(try)
    GetUniqueTempMHT = try
End Function

'---------------------------------------------------------------------------------------
' Procedure : TrimQuotedContent (Version 6 - Robust and Corrected)
' Author    : sebvannistel / 2025-06-21
' Purpose   : Finds the EARLIEST reply separator in the document across multiple
'             patterns and deletes all content from that point forward.
'---------------------------------------------------------------------------------------
Private Sub TrimQuotedContent(ByVal doc As Object)
    On Error Resume Next
    
    Const wdFindContinue As Long = 1
    Dim findRange As Object ' Word.Range
    Dim patterns As Variant
    Dim pat As Variant
    Dim firstSeparatorPos As Long
    
    ' Define all patterns to search for, from most to least specific
    patterns = Array( _
        "[-]{5,}Original Message[-]{5,}", _
        "From:*Sent:*To:*Subject:*", _
        "[<]hr[!>]*[>]", _
        "<blockquote*>" _
    )
    
    firstSeparatorPos = -1 ' Initialize to a "not found" state
    
    ' Loop through each pattern to find the one that appears EARLIEST
    For Each pat In patterns
        Set findRange = doc.Content
        With findRange.Find
            .ClearFormatting
            .Text = pat
            .Forward = True
            .Wrap = wdFindContinue ' Use wdFindStop to prevent wrapping around
            .Format = False
            .MatchCase = False
            .MatchWildcards = True
            
            If .Execute = True Then
                ' Safety check: don't trim if the separator is at the very top
                ' (e.g., the main "From:" line of the original email).
                If findRange.Start > 200 Then
                    ' If this is the first separator found, or if it's earlier
                    ' than the previous best, record its position.
                    If firstSeparatorPos = -1 Or findRange.Start < firstSeparatorPos Then
                        firstSeparatorPos = findRange.Start
                    End If
                End If
            End If
        End With
    Next pat
    
    ' After checking all patterns, if we found a separator, delete from that point.
    If firstSeparatorPos > -1 Then
        doc.Range(Start:=firstSeparatorPos, End:=doc.Content.End).Delete
    End If
    
    Set findRange = Nothing
    On Error GoTo 0
End Sub


' =========================================================================================
' === REVISED, STABLE MAIN PROCEDURE                                                    ===
' =========================================================================================
Sub SaveAsPDFfile()
    ' --- SETUP ---
    Const olDiscard As Long = 1
    Const wdExportFormatPDF As Long = 17
    Const MAX_PATH As Long = 259 ' Windows API limit for full file paths

    ' --- OBJECTS & VARIABLES ---
    Dim sel As Outlook.Selection
    Dim wrd As Object, doc As Object, fso As Object
    Dim mailItem As Outlook.MailItem
    Dim mailInspector As Outlook.Inspector ' Kept for reference, but not used in MHT method

    Dim tgtFolder As String, logFilePath As String
    Dim done As Long, skipped As Long, total As Long

    ' --- INITIALIZE ---
    On Error GoTo ErrorHandler

    ' Step 1: Get target folder using the STABLE native dialog
    ' Note: Using "C:\Temp" as a default starting point.
    tgtFolder = AskForTargetFolder("C:\Temp")
    If Len(tgtFolder) = 0 Then Exit Sub ' User cancelled

    ' Step 2: Get selections and create worker objects
    Set sel = Application.ActiveExplorer.Selection
    total = sel.Count
    If total = 0 Then GoTo Cleanup

    Set fso = CreateObject("Scripting.FileSystemObject")
    logFilePath = tgtFolder & "_SkippedItems_" & Format(Now, "yyyymmdd_hhnnss") & ".log"

    Set wrd = CreateObject("Word.Application")
    wrd.Visible = False
    wrd.DisplayAlerts = 0 ' Suppress Word's own popups
    Set objWord = wrd     ' Make Word object available to older helper functions if needed

    '================================================================================
    '--- MAIN EXPORT LOOP (STABILIZED) ---
    '================================================================================
    Dim item As Object
    Dim progressCounter As Long
    
    Application.StatusBar = "Preparing to save " & total & " selected email(s)..."

    For Each item In sel
        progressCounter = progressCounter + 1
        ' STABILITY FIX: Add DoEvents to prevent freezing and update status
        If progressCounter Mod 5 = 0 Then DoEvents
        Application.StatusBar = "Processing " & progressCounter & " of " & total & "..."
        
        If TypeOf item Is Outlook.MailItem Then
            Set mailItem = item
        Else
            ' Log and skip non-mail items
            skipped = skipped + 1
            LogSkippedItem logFilePath, "Unknown Item Type", "Item in selection was not a mail item."
            GoTo NextItem
        End If

        ' Use the MHT method as it is more stable than GetInspector
        Dim tmpMht As String, pdfFile As String, baseName As String
        On Error Resume Next
        
        ' 1. Build filenames
        tmpMht = GetUniqueTempMHT(mailItem, ".mht")
        baseName = Format(ItemDate(mailItem), "yyyymmdd-hhnnss") & " – " & CleanFile(mailItem.Subject)
        
        ' ... add path length check here ...
        ' Truncate baseName to ensure the full path with a potential suffix does not exceed MAX_PATH
        If Len(tgtFolder & baseName & ".pdf") >= MAX_PATH Then
            Dim room As Long
            ' Reserve space for a suffix like "_99.pdf" (7 chars) to avoid path-too-long errors during de-duplication
            room = MAX_PATH - Len(tgtFolder) - 7
            If room > 0 Then
                baseName = Left$(baseName, room)
            Else
                 ' Handle case where target folder path is already too long to create any file
                 LogSkippedItem logFilePath, mailItem.Subject, "Target folder path is too long to create a valid filename."
                 skipped = skipped + 1
                 GoTo NextItem
            End If
        End If
        
        ' DATA LOSS RISK FIX: De-duplicate filename to prevent silently overwriting existing files.
        Dim dupCounter As Long
        pdfFile = tgtFolder & baseName & ".pdf"
        dupCounter = 1
        ' The fso object is created at the start of the sub
        Do While fso.FileExists(pdfFile)
            ' If file exists, append a suffix and check again
            pdfFile = tgtFolder & baseName & "_" & dupCounter & ".pdf"
            dupCounter = dupCounter + 1
        Loop

        ' 2. Save to MHT
        mailItem.SaveAs tmpMht, olMHTML
        If Err.Number <> 0 Then
            Err.Clear
            LogSkippedItem logFilePath, mailItem.Subject, "Failed to save as MHT (IRM protected or locked)."
            skipped = skipped + 1
            GoTo NextItem
        End If
        
        ' 3. Open in Word, process, and export
        Set doc = wrd.Documents.Open(tmpMht, ReadOnly:=True, Visible:=False)
        If Err.Number <> 0 Then
             Err.Clear
             LogSkippedItem logFilePath, mailItem.Subject, "Word failed to open the MHT file."
             skipped = skipped + 1
             GoTo NextItem
        End If
        
        ' The helper functions are available from the original code
        Call InjectFullHeader(doc, mailItem)
        Call TrimQuotedContent(doc)
        
        doc.ExportAsFixedFormat pdfFile, wdExportFormatPDF
        
        ' ... update counters (done/skipped) ...
        If Err.Number <> 0 Then
            LogSkippedItem logFilePath, mailItem.Subject, "Word failed to export MHT to PDF. Error: " & Err.Description
            skipped = skipped + 1
            Err.Clear
        Else
            done = done + 1
        End If

' ================================================================================
' === CRITICAL FIX: The cleanup block for EACH loop iteration                  ===
' ================================================================================
NextItem:
        ' The CORRECT cleanup order: child objects first, then parents.
        If Not doc Is Nothing Then
            doc.Close False ' Close the Word doc without saving changes
            Set doc = Nothing
        End If
        
        ' The MHT file is no longer tied to an inspector, so we just delete it
        If Len(tmpMht) > 0 And fso.FileExists(tmpMht) Then
            fso.DeleteFile tmpMht, True
        End If

        ' Release the mail item object for this loop
        Set mailItem = Nothing
        tmpMht = "" ' Clear the temp file path
    Next item

    ' --- FINAL CLEANUP ---
    Dim msg As String
    msg = done & " mail(s) successfully saved as PDF to " & vbCrLf & tgtFolder
    If skipped > 0 Then
        msg = msg & vbCrLf & vbCrLf & skipped & " item(s) were skipped. See the log file for details:" & vbCrLf & logFilePath
    End If
    MsgBox msg, vbInformation, "Export Complete"

Cleanup:
    Application.StatusBar = False
    On Error Resume Next
    If Not wrd Is Nothing Then wrd.Quit
    ' ... release all other objects ...
    Set wrd = Nothing: Set fso = Nothing: Set sel = Nothing
    Set doc = Nothing: Set mailItem = Nothing: Set objWord = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "A critical error occurred: " & Err.Description, vbCritical, "Macro Error"
    GoTo Cleanup
End Sub
