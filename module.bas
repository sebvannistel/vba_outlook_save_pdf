#If VBA7 Then
Private Declare PtrSafe Function SetForegroundWindow _
        Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
Private Declare PtrSafe Function FindWindowA Lib "user32" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
Private Declare Function SetForegroundWindow _
        Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare Function FindWindowA Lib "user32" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
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
        wHwnd = objWord.ActiveWindow.Hwnd ' Word’s real property
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
    wHwnd = objWord.ActiveWindow.Hwnd ' Word’s real property
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
    Const tmpExt As String = ".mht"
    Const olMail As Long = 43
    Const PATH_WARN As Long = 220                   ' <-- PATCH 2: Use a named constant for path length
    Const wdExportOptimizeForPrint As Long = 0      ' <-- PATCH 2: Use a named constant for readability

    Dim sel As Outlook.Selection
    Dim wrd As Object, doc As Object, fso As Object
    Dim tmpFile As String, pdfFile As String, tgtFolder As String, logFilePath As String
    Dim total As Long, done As Long, showProgress As Boolean

    ' --- INITIALIZE ---
    Set wrd = CreateObject("Word.Application")
    Set objWord = wrd
    ' wrd.Visible = True ' <-- For debugging only
    
    tgtFolder = AskForTargetFolder("")
    
    If Len(tgtFolder) > PATH_WARN Then ' <-- PATCH 2: APPLIED
        MsgBox "The selected folder path is too long. Please choose a shorter path to avoid errors.", vbExclamation, "Path Too Long"
        wrd.Quit
        Set wrd = Nothing: Set objWord = Nothing
        Exit Sub
    End If
    
    Set sel = Application.ActiveExplorer.Selection
    If sel.Count = 0 Or Len(tgtFolder) = 0 Then
        wrd.Quit
        Set wrd = Nothing: Set objWord = Nothing
        Exit Sub
    End If
    
    logFilePath = tgtFolder & "_SkippedItems.log"
    wrd.DisplayAlerts = 0
    Set fso = CreateObject("Scripting.FileSystemObject")

    '================================================================================
    '--- LAYER 1: PRE-FILTER SELECTION (FIX #1: Verify Dictionary Creation) ---
    '================================================================================
    Dim convDict As Object
    Dim itm As Object, key As String

    On Error Resume Next
    Set convDict = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    
    If convDict Is Nothing Then
        MsgBox "Critical Error: Could not create the 'Scripting.Dictionary' object." & vbCrLf & _
               "This may be due to a missing or corrupted system library (scrrun.dll).", vbCritical
        wrd.Quit
        Set wrd = Nothing: Set objWord = Nothing
        Exit Sub
    End If

    On Error Resume Next
    For Each itm In sel
        If TypeOf itm Is Outlook.MailItem Then
            key = itm.ConversationID
            If Err.Number = 0 And Len(key) > 0 Then
                If Not convDict.Exists(key) Then
                    Set convDict(key) = itm
                ElseIf convDict(key).ReceivedTime > 0 And itm.ReceivedTime > convDict(key).ReceivedTime Then ' <-- PATCH 3: APPLIED
                    Set convDict(key) = itm
                End If
            Else
                Err.Clear
                Set convDict(itm.EntryID) = itm ' Corrected: Added 'Set'
            End If
        End If
    Next itm
    On Error GoTo 0
    
    total = convDict.Count
    showProgress = (total > 1)
    If showProgress Then wrd.StatusBar = "Preparing to save " & total & " top-level mail(s)..."

    '================================================================================
    '--- MAIN EXPORT LOOP (FIX #2: Use a safe Variant loop) ---
    '================================================================================
    Dim miAsVariant As Variant
    Dim mailItem As Outlook.MailItem
    
    For Each miAsVariant In convDict.Items
        ' Check if the item from the dictionary is a valid MailItem object
        If TypeOf miAsVariant Is Outlook.MailItem Then
            Set mailItem = miAsVariant ' Safe to assign now
            
            ' --- Safety checks ---
            If mailItem.Class <> olMail Then LogSkippedItem logFilePath, mailItem.Subject, "Not a true MailItem": GoTo NextItemInLoop
            If mailItem.Size = 0 Then LogSkippedItem logFilePath, mailItem.Subject, "Item size is 0": GoTo NextItemInLoop
            If mailItem.Permission <> olUnrestricted Then LogSkippedItem logFilePath, mailItem.Subject, "Item is IRM Protected": GoTo NextItemInLoop

            ' PATCH: Use nested If to prevent "Array index out of bounds" error
            If mailItem.Attachments.Count = 1 Then
                If LCase$(mailItem.Attachments(1).FileName) Like "*.p7m" Then
                    LogSkippedItem logFilePath, mailItem.Subject, "S/MIME Encrypted"
                    GoTo NextItemInLoop
                End If
            End If

            ' --- 1. FILENAME BUILDER ---
            Dim safeSubj As String, datePrefix As String, room As Long, roomForTmp As Long, roomForPdf As Long
            Const cMAX_PATH As Long = 260
            safeSubj = CleanFile(mailItem.Subject)
            datePrefix = Format(ItemDate(mailItem), "yyyymmdd-hhnnss")
            roomForTmp = (cMAX_PATH - 1) - (Len(Environ$("TEMP") & "\") + Len(datePrefix) + Len("_") + Len(tmpExt))
            roomForPdf = (cMAX_PATH - 1) - (Len(tgtFolder) + Len(datePrefix) + Len(" – ") + Len(".pdf"))
            If roomForTmp < roomForPdf Then room = roomForTmp Else room = roomForPdf
            If room < 10 Then room = 10
            If Len(safeSubj) > room Then safeSubj = Left$(safeSubj, room)
            tmpFile = GetUniqueTempMHT(mailItem, tmpExt)
            pdfFile = tgtFolder & datePrefix & " – " & safeSubj & ".pdf"

            ' --- 2. SAVE AS MHT (with MSG fallback) ---
            On Error Resume Next
            mailItem.SaveAs tmpFile, olMHTML
            If Err.Number <> 0 Then
                Err.Clear
                LogSkippedItem logFilePath, mailItem.Subject, "MHTML save failed, falling back to .MSG"
                Dim msgFallbackFile As String
                msgFallbackFile = tgtFolder & datePrefix & " – " & safeSubj & ".msg"
                If fso.FileExists(msgFallbackFile) Then fso.DeleteFile msgFallbackFile, True
                mailItem.SaveAs msgFallbackFile, 9 ' olMSGUnicode
                GoTo NextItemInLoop
            End If
            On Error GoTo 0
            
            ' --- 3. OPEN IN WORD & PREPARE ---
            Set doc = Nothing
            On Error Resume Next
            Set doc = wrd.Documents.Open(FileName:=tmpFile, ConfirmConversions:=False, ReadOnly:=True, Visible:=False, Format:=wdOpenFormatWebPages)
            On Error GoTo 0
            If doc Is Nothing Then LogSkippedItem logFilePath, mailItem.Subject, "Word could not open MHT": GoTo NextItemInLoop

            ' --- Inject Header ---
            Dim hdr As String
            hdr = "From:    " & mailItem.SenderName & vbCrLf & "Sent:    " & mailItem.SentOn & vbCrLf & "To:      " & mailItem.To & vbCrLf & "CC:      " & mailItem.CC & vbCrLf & "BCC:     " & mailItem.BCC & vbCrLf & "Subject: " & mailItem.Subject & vbCrLf & String(60, "—") & vbCrLf & vbCrLf
            If InStr(1, doc.Range(0, 50).Text, "From:") = 0 Then doc.Range.InsertBefore hdr

            '================================================================================
            '--- LAYER 2: QUOTED TEXT REMOVAL (Using the new, corrected function) ---
            '================================================================================
            Call TrimQuotedContent(doc)
            
            ' --- 4. EXPORT TO PDF & CLEANUP ---
            If fso.FileExists(pdfFile) Then fso.DeleteFile pdfFile, True
            doc.ExportAsFixedFormat OutputFileName:=pdfFile, ExportFormat:=wdExportFormatPDF, OptimizeFor:=wdExportOptimizeForPrint ' <-- PATCH 2: APPLIED
            doc.Close False
            Set doc = Nothing
            fso.DeleteFile tmpFile
            
            done = done + 1
            If showProgress Then wrd.StatusBar = "Saving mail " & done & " of " & total & "..."

        End If ' End of "If TypeOf miAsVariant Is Outlook.MailItem"

NextItemInLoop:
        If Not doc Is Nothing Then doc.Close False
    Next miAsVariant

    ' --- FINAL CLEANUP ---
    If showProgress Then wrd.StatusBar = False
    wrd.Quit
    Set doc = Nothing: Set wrd = Nothing: Set objWord = Nothing
    Set fso = Nothing: Set sel = Nothing
    Set itm = Nothing
    Set mailItem = Nothing: Set convDict = Nothing

    MsgBox done & " mail(s) saved as PDF to " & tgtFolder & vbCrLf & vbCrLf & "A log of any skipped items has been saved to _SkippedItems.log", vbInformation

End Sub
