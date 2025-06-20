#If VBA7 Then
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
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
Private Const wdOpenFormatWebPages As Long = 7

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
        
        '--- START: Version-safe method to bring Word to the foreground ---
        ' The Word.Application object does not have an .Hwnd property.
        ' We must get the handle from the ActiveWindow or use FindWindow API.
        Dim wHwnd As LongPtr
        On Error Resume Next              ' Suppress error 438 on older Word versions or if no window is active
        wHwnd = objWord.ActiveWindow.hWnd ' Word’s real property
        If Err.Number <> 0 Then           ' If that fails, fall back to API
            Err.Clear
            ' Use Win32 FindWindow on the Word application's class name "OpusApp"
            wHwnd = FindWindowA("OpusApp", vbNullString)
        End If
        On Error GoTo 0                   ' Restore normal error trapping

        ' If we successfully got a handle, bring the window to the front.
        If wHwnd <> 0 Then
            Call SetForegroundWindow(wHwnd)
        End If
        '--- END: Version-safe method ---
        
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

    ' UPDATE 1: Make the Save-As picker topmost as well
    '--- START: Version-safe method to bring Word to the foreground ---
    ' The Word.Application object does not have an .Hwnd property.
    ' We must get the handle from the ActiveWindow or use FindWindow API.
    Dim wHwnd As LongPtr
    On Error Resume Next              ' Suppress error 438 on older Word versions or if no window is active
    wHwnd = objWord.ActiveWindow.hWnd ' Word’s real property
    If Err.Number <> 0 Then           ' If that fails, fall back to API
        Err.Clear
        ' Use Win32 FindWindow on the Word application's class name "OpusApp"
        wHwnd = FindWindowA("OpusApp", vbNullString)
    End If
    On Error GoTo 0                   ' Restore normal error trapping

    ' If we successfully got a handle, bring the window to the front.
    If wHwnd <> 0 Then
        Call SetForegroundWindow(wHwnd)
    End If
    '--- END: Version-safe method ---
    
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
    Select Case True
        Case TypeOf itm Is Outlook.mailItem
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
Private Sub InjectSimpleHeader(doc As Object, m As Outlook.mailItem)
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
Private Sub InjectFullHeader(doc As Object, m As Outlook.mailItem)
    On Error Resume Next ' In case a property is not available
    Dim hdr As String
    hdr = "From: " & m.SenderName & vbCrLf & _
          "Sent: " & m.SentOn & vbCrLf & _
          "To: " & m.To & vbCrLf & _
          IIf(Len(m.CC) > 0, "Cc: " & m.CC & vbCrLf, "") & _
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
Private Function GetUniqueTempMHT(mi As Outlook.mailItem, ext As String) As String
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim base$, try$
    base = Environ$("TEMP") & "\" & Format(ItemDate(mi), "yyyymmdd-hhnnss") _
           & "_" & CleanFile(mi.Subject)
    Do
        try = Left$(base, 200) & "_" & Hex(Timer * 1000) & ext   ' <= always unique
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
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWildcards = True
            
            If .Execute = True Then
                ' Safety check: don't trim if the separator is at the very top
                ' (e.g., the main "From:" line of the original email).
                If .Parent.Start > 200 Then
                    ' If this is the first separator found, or if it's earlier
                    ' than the previous best, record its position.
                    If firstSeparatorPos = -1 Or .Parent.Start < firstSeparatorPos Then
                        firstSeparatorPos = .Parent.Start
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


'Save selected Outlook messages as PDFs – quiet, de-duplicated, & with headers
Sub SaveMails_ToPDF_Background()
    ' --- SETUP ---
    Const olMail As Long = 43
    Const PATH_WARN As Long = 220
    Const wdExportOptimizeForPrint As Long = 0
    Const olDiscard As Long = 1 ' Constant for MailItem.Close
    Const wdExportCreateNoBookmarks = 0 ' New constant for bookmarks

    Dim sel As Outlook.Selection
    Dim wrd As Object, doc As Object, fso As Object
    Dim pdfFile As String, tgtFolder As String, logFilePath As String
    Dim total As Long, done As Long, skipped As Long, showProgress As Boolean

    ' --- INITIALIZE ---
    On Error GoTo ErrorHandler
    Set wrd = CreateObject("Word.Application")
    Set objWord = wrd
    ' wrd.Visible = True ' <-- For debugging only
    
    tgtFolder = AskForTargetFolder("")
    
    If Len(tgtFolder) > PATH_WARN Then
        MsgBox "The selected folder path is too long. Please choose a shorter path to avoid errors.", vbExclamation, "Path Too Long"
        GoTo Cleanup
    End If
    
    Set sel = Application.ActiveExplorer.Selection
    If sel.Count = 0 Or Len(tgtFolder) = 0 Then
        GoTo Cleanup
    End If
    
    logFilePath = tgtFolder & "_SkippedItems.log"
    wrd.DisplayAlerts = 0
    Set fso = CreateObject("Scripting.FileSystemObject")

    '================================================================================
    '--- NO PRE-FILTERING: Iterate through all selected items ---
    '================================================================================
    'FIX: The de-duplication dictionary (convDict) has been removed to ensure
    'every selected item is processed, not just the latest in a conversation.
    
    total = sel.Count ' Process all selected items
    showProgress = (total > 1)
    If showProgress Then wrd.StatusBar = "Preparing to save " & total & " selected mail(s)..."

    '================================================================================
    '--- MAIN EXPORT LOOP ---
    '================================================================================
    Dim mailItemVariant As Variant
    Dim mailItem As Outlook.MailItem

    'FIX: Iterate directly over the selection (sel) instead of the de-duped dictionary
    For Each mailItemVariant In sel

        If TypeOf mailItemVariant Is Outlook.MailItem Then
            Set mailItem = mailItemVariant
        Else
            skipped = skipped + 1
            LogSkippedItem logFilePath, "Unknown Item", "Item in collection was not a valid Mail object."
            GoTo NextSelectedItem
        End If
        
        On Error Resume Next
        mailItem.Display False
        Set doc = mailItem.GetInspector.WordEditor
        If doc Is Nothing Or Err.Number <> 0 Then
            Err.Clear
            mailItem.Close olDiscard
            skipped = skipped + 1
            LogSkippedItem logFilePath, mailItem.Subject, "Could not access content (possibly protected)."
            GoTo NextSelectedItem
        End If
        On Error GoTo ErrorHandler

        'FIX: Call header routine unconditionally to ensure it's always added.
        'The routine itself now contains a guard against duplicate headers.
        Call InjectFullHeader(doc, mailItem)

        'FIX: Call the content stripper to remove quoted replies from the body.
        Call TrimQuotedContent(doc)

        ' *** START: HARDENED FILENAME AND EXPORT BLOCK ***
        Dim safeSubj As String, datePrefix As String, baseName As String
        Const MAX_PATH As Long = 259 ' Windows API limit
        safeSubj = CleanFile(mailItem.Subject)
        datePrefix = Format(ItemDate(mailItem), "yyyymmdd-hhnnss")
        baseName = datePrefix & " – " & safeSubj
        
        ' 1. Stricter path length check to prevent the export error
        If Len(tgtFolder & baseName & ".pdf") >= MAX_PATH Then
            baseName = Left$(baseName, MAX_PATH - Len(tgtFolder) - 5)
        End If
        pdfFile = tgtFolder & baseName & ".pdf"

        ' 2. Ensure we're not overwriting a locked file
        If fso.FileExists(pdfFile) Then fso.DeleteFile pdfFile, True

        On Error Resume Next ' Catch any final, unexpected export errors
        
        ' 3. Robust PDF export with parameters to prevent common failures
        doc.ExportAsFixedFormat _
            OutputFileName:=pdfFile, _
            ExportFormat:=wdExportFormatPDF, _
            OptimizeFor:=wdExportOptimizeForPrint, _
            KeepIRM:=False, _
            BitmapMissingFonts:=True, _
            CreateBookmarks:=wdExportCreateNoBookmarks
            
        If Err.Number <> 0 Then
            skipped = skipped + 1
            LogSkippedItem logFilePath, mailItem.Subject, "Export failed (Error " & Err.Number & ": " & Err.Description & ")"
            Err.Clear
        Else
            done = done + 1
        End If
        On Error GoTo ErrorHandler
        ' *** END: HARDENED FILENAME AND EXPORT BLOCK ***

        mailItem.Close olDiscard
        Set doc = Nothing
        
        If showProgress Then wrd.StatusBar = "Processing mail " & (done + skipped) & " of " & total & "..."

NextSelectedItem:
    Next mailItemVariant

    ' --- FINAL CLEANUP ---
    If showProgress Then wrd.StatusBar = False
    Dim msg As String
    msg = done & " mail(s) saved as PDF to " & tgtFolder
    If skipped > 0 Then
        msg = msg & vbCrLf & vbCrLf & skipped & " item(s) were skipped. See the log file for details."
    End If
    MsgBox msg, vbInformation

Cleanup:
    On Error Resume Next
    If Not wrd Is Nothing Then wrd.Quit
    Set doc = Nothing: Set wrd = Nothing: Set objWord = Nothing
    Set fso = Nothing: Set sel = Nothing
    Set mailItem = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "A critical error occurred." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Macro Error"
    GoTo Cleanup
End Sub
