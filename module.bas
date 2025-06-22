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

Private Const wdExportFormatPDF As Long = 17       'moved to module level
Private Const olMHTML As Long = 10                 'Added for late-binding
Private Const wdStatisticPages As Long = 2         'For readability, avoids magic number
Private Const wdExportOptimizeForPrint As Long = 0 'For late binding, fixes compile error

Private objWord As Object

' =========================================================================================
' === UNIVERSAL FOLDER PICKER (Fallback for when FileDialog fails)                      ===
' =========================================================================================
Private Function GetTargetFolder_Universal() As String
    Dim shellApp As Object
    Dim folder As Object
    Dim folderPath As String
    
    On Error Resume Next
    Set shellApp = CreateObject("Shell.Application")
    If shellApp Is Nothing Then
        GetTargetFolder_Universal = InputBox("Could not create the Shell object. Please enter the full folder path:", "Enter Folder Path")
        Exit Function
    End If
    
    ' BIF_RETURNONLYFSDIRS (1) + BIF_NEWDIALOGSTYLE (64)
    Set folder = shellApp.BrowseForFolder(0, "Please select a folder to save the PDFs", 1 + 64)
    
    If Not folder Is Nothing Then
        folderPath = folder.Self.Path
        ' Ensure the path ends with a backslash
        If Right(folderPath, 1) <> "\" Then
            folderPath = folderPath & "\"
        End If
    End If
    
    GetTargetFolder_Universal = folderPath
    
    Set folder = Nothing
    Set shellApp = Nothing
    On Error GoTo 0
End Function

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

'--- HELPER for dictionary building. Items that are not standard mail
'--- (e.g., reports) don't have a ConversationID and are treated uniquely.
Private Function IsSpecial(itm As Object) As Boolean
    IsSpecial = Not (TypeOf itm Is Outlook.MailItem)
End Function

Private Function CleanFile(s As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In badChars
        s = Replace$(s, ch, " ")
    Next ch
    CleanFile = Trim$(s)
End Function

'--- NEW ATTACHMENT HELPER FUNCTION ---
Private Function AttachmentList(m As Outlook.MailItem) As String
    Dim tmp As String, i As Long
    For i = 1 To m.Attachments.Count
        tmp = tmp & m.Attachments(i).FileName & _
              IIf(i < m.Attachments.Count, "; ", "")
    Next i
    AttachmentList = tmp
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
          IIf(m.Attachments.Count > 0, "Attachments: " & AttachmentList(m) & vbCrLf, "") & _
          String(60, "—") & vbCrLf & vbCrLf
    doc.Range.InsertBefore hdr
    On Error GoTo 0
End Sub

'------ Helper: Tidy final Word doc and trim quoted text/footers -----------------
Private Sub TidyAndTrimDocument(wdDoc As Word.Document)

    '------ 2 a  ▸ unify font & size  ---------------------------------------
    'Attach corporate template *and* wipe direct formatting
    wdDoc.ApplyDocumentTemplate "C:\Path\Brand.dotx"

    With wdDoc.Content.Font                 ' ← restores uniformity
        .Name = "Calibri"
        .Size = 11
    End With
    With wdDoc.Styles(wdStyleNormal).Font   ' ← keeps tables/lists consistent
        .Name = "Calibri"
        .Size = 11
    End With

    wdDoc.Styles(wdStyleHeading1).Font.Name = "Calibri Light"

    '------ 2 b  ▸ remove quoted replies / footers  --------------------------
    Dim rng As Word.Range
    Set rng = wdDoc.Content

    'Typical separators you said slip through ("From:", huge disclaimers, etc.)
    Const CUT_MARKS As String = _
          "-----Original Message-----|From: |Brucher Thieltgen & Partners"
    Const FOOTER_MARK As String = "Please consider the impact on the environment"
    Dim parts() As String: parts = Split(CUT_MARKS, "|")

    Dim i As Long, pos As Long

    'first remove known footer blocks without affecting replies
    pos = InStr(1, rng.Text, FOOTER_MARK, vbTextCompare)
    If pos > 0 Then
        wdDoc.Range(rng.Start + pos - 1, wdDoc.Content.End).Delete
        Set rng = wdDoc.Content
    End If

    For i = LBound(parts) To UBound(parts)
        pos = InStr(1, rng.Text, parts(i), vbTextCompare)
        If pos > 0 Then
            rng.End = rng.Start + pos - 2        '-2 keeps the line break neat
            Exit For
        End If
    Next i
    rng.Delete                                       'delete everything after cut mark

    '------ 2 c  ▸ compact extra blank lines left behind  --------------------
    wdDoc.Range.ParagraphFormat.SpaceAfter = 0
    wdDoc.Range.Find.Execute FindText:=vbLf & vbLf, _
                              ReplaceWith:=vbLf, _
                              Replace:=wdReplaceAll
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
          IIf(m.Attachments.Count > 0, "Attachments: " & AttachmentList(m) & vbCrLf, "") & _
          String(60, "—") & vbCrLf & vbCrLf
          
    ' FIX #3: Guard against double-inserting your own header
    ' Increased look-ahead from 60 to 120 characters to be safer.
    If InStr(1, doc.Range(0, 120).Text, "From:") = 0 Then
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

' *** NEW HELPER SUB: Robustly clears the status bar, handling all errors silently. ***
Private Sub SafeClearStatusBar()
    On Error Resume Next ' swallow 438 or other errors silently
    Dim exp As Outlook.Explorer
    Set exp = Application.ActiveExplorer
    If Not exp Is Nothing Then
        exp.StatusBar = "" ' clear custom text
    End If
    On Error GoTo 0
End Sub

'--- helper: always create a unique temp MHT name
Private Function GetUniqueTempMHT(mi As Outlook.MailItem, ext As String) As String
    ' OPTIONAL POLISH (from checklist item #3): Seed the random number generator.
    Randomize
    
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

' NEW FUNCTION - Strips quoted replies from the HTMLBody BEFORE saving to MHT.
' This is the primary fix for preventing conversation history in PDFs.
Private Function StripQuotedBody(ByVal mailItem As Outlook.MailItem) As String
    Dim body As String
    Dim patterns As Variant
    Dim pat As Variant
    Dim pos As Long
    Dim firstSeparatorPos As Long
    
    body = mailItem.HTMLBody
    If Len(body) = 0 Then
        StripQuotedBody = "" ' Return empty if body is empty
        Exit Function
    End If
    
    ' These patterns identify the start of a replied/forwarded message.
    ' The function finds the EARLIEST match and trims from that point onwards.
    ' This array can be extended with other language-specific separators.
    ' ========================================================================
    ' === FIX 3.1: PATTERN ORDER CHANGED & GUARD ADDED (as per instructions) ===
    ' The header-div is now last to prevent premature matches.
    patterns = Array( _
        "-----Original Message-----", _
        "<div class=3D""gmail_quote"">", _
        "<blockquote>", _
        "<hr", _
        "<div class=""OutlookMessageHeader"">" _
    )
    
    ' Using 0 as "not found"
    firstSeparatorPos = 0

    ' Loop through each pattern to find the one that appears EARLIEST
    For Each pat In patterns
        pos = InStr(1, body, CStr(pat), vbTextCompare)
        ' The length guard prevents early matches from wiping the entire email.
        If pos > 250 Then   ' <-- UPDATED: skip early matches that would wipe the whole mail
            If firstSeparatorPos = 0 Or pos < firstSeparatorPos Then
                firstSeparatorPos = pos
            End If
        End If
    Next pat
    ' === END OF FIX 3.1 ===
    ' ======================
    
    ' If a separator was found, truncate the string. Otherwise, return the original.
    If firstSeparatorPos > 0 Then
        StripQuotedBody = Left(body, firstSeparatorPos - 1)
    Else
        StripQuotedBody = body
    End If
End Function

' === NEW Helper ===
Private Sub SaveHtmlToMht(ByVal html As String, ByVal mhtPath As String, wrd As Object)
    Dim tmpHtml As String, stm As Object, docTmp As Object
    tmpHtml = Replace(mhtPath, ".mht", ".html")

    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Charset = "utf-8"
        .Type = 2          'text
        .Open
        .WriteText html
        .SaveToFile tmpHtml, 2  'adSaveCreateOverWrite
        .Close
    End With

    Set docTmp = wrd.Documents.Open(tmpHtml, ReadOnly:=True, Visible:=False, ConfirmConversions:=False)
    
    ' REQUIRED FIX (from checklist item #1): Use 9 for MHT, not 10.
    docTmp.SaveAs2 mhtPath, 9   '9 = wdFormatWebArchive
    
    docTmp.Close False
    
    ' OPTIONAL POLISH (from checklist item #4): Release object before file operation.
    Set docTmp = Nothing
    
    Kill tmpHtml
End Sub

'---------------------------------------------------------------------------------------
' Procedure : TrimQuotedContent (Version 7 - Safer patterns)
' Author    : sebvannistel / 2025-06-21
' Purpose   : Finds the EARLIEST reply separator in the document across multiple
'             patterns and deletes all content from that point forward.
'---------------------------------------------------------------------------------------
Private Sub TrimQuotedContent(ByVal doc As Object)
    On Error Resume Next
    
    Const wdFindStop As Long = 0
    Dim findRange As Object ' Word.Range
    Dim patterns As Variant
    Dim pat As Variant
    Dim firstSeparatorPos As Long
    
    ' UPDATED patterns as per new guidance for better safety
    patterns = Array( _
       "[-]{5,}Original Message[-]{5,}", _
       "<div class=""OutlookMessageHeader"">", _
       "<div class=3D""gmail_quote"">", _
       "(^|\r)(>+ )?On *[0-9]{4}.*wrote:?", _
       "<hr[^>]*>", _
       "<blockquote[^>]*>" _
    )
    
    firstSeparatorPos = -1 ' Initialize to a "not found" state
    
    ' Loop through each pattern to find the one that appears EARLIEST
    For Each pat In patterns
        Set findRange = doc.Content
        With findRange.Find
            .ClearFormatting
            .Text = pat
            .Forward = True
            .Wrap = wdFindStop ' Use wdFindStop to prevent finding a later separator
            .Format = False
            .MatchCase = False
            .MatchWildcards = True
            
            If .Execute = True Then
                ' Safety check "If findRange.Start > 200" removed as requested.
                
                ' If this is the first separator found, or if it's earlier
                ' than the previous best, record its position.
                If firstSeparatorPos = -1 Or findRange.Start < firstSeparatorPos Then
                    firstSeparatorPos = findRange.Start
                End If
            End If
        End With
    Next pat
    
    ' After checking all patterns, if we found a separator, delete from that point.
    '— TrimQuotedContent (restore safety) —
    If firstSeparatorPos > 249 Then
        doc.Range(Start:=firstSeparatorPos, End:=doc.Content.End).Delete
    End If
    
    Set findRange = Nothing
    On Error GoTo 0
End Sub

' =========================================================================================
' === FINAL, ROBUST MAIN PROCEDURE (WITH ALL SAFETY CHECKS)                             ===
' =========================================================================================
Private SaveAsPDFfile()
    ' --- SETUP ---
    Const MAX_PATH As Long = 259

    ' --- OBJECTS & VARIABLES ---
    Dim sel As Outlook.Selection
    Dim wrd As Object, doc As Object, fso As Object
    Dim mailItem As Outlook.MailItem
    Dim tgtFolder As String, logFilePath As String
    Dim done As Long, skipped As Long, total As Long

    On Error GoTo ErrorHandler

    ' Step 1: Get target folder
    tgtFolder = GetTargetFolder_Universal()
    If Len(tgtFolder) = 0 Then Exit Sub

    ' Step 2: Get selections
    ' *** NEW: First, check if an Explorer window is even active ***
    If Application.ActiveExplorer Is Nothing Then
        MsgBox "Cannot run the macro." & vbCrLf & vbCrLf & _
               "Please go to your main Outlook window, select the emails you want to save, " & _
               "and then run the macro again.", vbExclamation, "No Active Window"
        GoTo Cleanup
    End If

    ' Now that we know the Explorer exists, we can safely get the selection
    Set sel = Application.ActiveExplorer.Selection

    If sel.Count = 0 Then
        MsgBox "Please select one or more emails to save.", vbInformation, "No Items Selected"
        GoTo Cleanup
    End If

    ' FIX #1: This section now filters by EntryID to ensure each unique mail is
    ' processed, solving the "one-per-conversation" problem.
    '-----------------------------------------------------------------
    Dim latest As Object
    Set latest = CreateObject("Scripting.Dictionary")
    
    ' *** UPDATE #1: Add a creation guard immediately after you build the dictionary ***
    If latest Is Nothing Then                     ' -- COM failed (scrrun.dll not registered)
        MsgBox "Could not create Scripting.Dictionary – check scrrun.dll registration.", vbCritical
        Exit Sub
    End If
    
    Dim k As String, it As Object

    ' This loop ensures every unique email you selected is included for export.
    For Each it In sel
        k = it.EntryID                 ' <- unique key per message
        If Not latest.Exists(k) Then
            latest.Add k, it
        End If
    Next it
    '-----------------------------------------------------------------

    total = latest.Count ' Use the count of the filtered list

    ' Step 3: Initialize worker objects
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFilePath = tgtFolder & "_SkippedItems_" & Format(Now, "yyyymmdd_hhnnss") & ".log"

    ' --- DETAILED DIAGNOSTIC FOR WORD OBJECT CREATION ---
    Dim failedStep As String
    Const msoAutomationSecurityForceDisable As Long = 3 ' For late binding

    On Error Resume Next ' Temporarily disable the main error handler

    failedStep = "CreateObject(""Word.Application"")"
    Set wrd = CreateObject("Word.Application")
    If Err.Number <> 0 Then GoTo WordCreationFailed

    '--- IMPLEMENTING BLUEPRINT STEP 1: The "Fort Knox" Word Setup ---
    ' Make it invisible (.Visible = False).
    failedStep = "wrd.Visible = False"
    wrd.Visible = False
    If Err.Number <> 0 Then GoTo WordCreationFailed

    ' Tell it to suppress all alerts (.DisplayAlerts = wdAlertsNone).
    failedStep = "wrd.DisplayAlerts = 0 (wdAlertsNone)"
    wrd.DisplayAlerts = 0 ' wdAlertsNone
    If Err.Number <> 0 Then GoTo WordCreationFailed

    ' Disable the "Update Links" prompt (.Options.UpdateLinksAtOpen = False).
    failedStep = "wrd.Options.UpdateLinksAtOpen = False"
    wrd.Options.UpdateLinksAtOpen = False
    If Err.Number <> 0 Then GoTo WordCreationFailed

    ' Force it to disable all security prompts (.AutomationSecurity = msoAutomationSecurityForceDisable).
    failedStep = "wrd.AutomationSecurity = msoAutomationSecurityForceDisable"
    wrd.AutomationSecurity = msoAutomationSecurityForceDisable ' Value is 3
    If Err.Number <> 0 Then GoTo WordCreationFailed
    '--- END OF BLUEPRINT STEP 1 IMPLEMENTATION ---

    On Error GoTo ErrorHandler ' Restore the main error handler

    Set objWord = wrd
    ' --- END OF DIAGNOSTIC BLOCK ---

    '================================================================================
    '--- MAIN EXPORT LOOP ---
    '================================================================================
    Dim item As Variant
    Dim progressCounter As Long
    
    ' *** UPDATED: Set status bar only if Explorer is active and property exists ***
    If Not Application.ActiveExplorer Is Nothing Then
        On Error Resume Next ' In case .StatusBar is deprecated or causes an error
        Application.ActiveExplorer.StatusBar = "Preparing to save " & total & " selected email(s)..."
        On Error GoTo 0
    End If

    ' *** UPDATE #2: Add a type-safety guard *right before* the For Each loop ***
    '— sanity check: has <latest> been clobbered in the meantime? —
    If Not (TypeName(latest) = "Dictionary") Then
        MsgBox "Internal error – variable <latest> is no longer a Dictionary.", vbCritical
        Exit Sub
    End If

    ' --- FIX: Prevent Run-time error 424 because latest.Items is Empty ---
    ' This guard ensures the loop is only attempted if the dictionary
    ' contains items. If the count is 0, we jump to the cleanup section.
    If latest.Count = 0 Then
        MsgBox "Nothing to export – the filter removed every item.", vbInformation
        GoTo Cleanup
    End If
    
    ' Iterate over the filtered dictionary items, not the original selection
    For Each item In latest.Items
        progressCounter = progressCounter + 1
        If progressCounter Mod 5 = 0 Then DoEvents
        
        ' *** UPDATED: Set status bar only if Explorer is active and property exists ***
        If Not Application.ActiveExplorer Is Nothing Then
            On Error Resume Next ' In case .StatusBar is deprecated or causes an error
            Application.ActiveExplorer.StatusBar = "Processing " & progressCounter & " of " & total & "..."
            On Error GoTo 0
        End If
        
        If TypeOf item Is Outlook.MailItem Then
            Set mailItem = item
        Else
            skipped = skipped + 1
            LogSkippedItem logFilePath, "Unknown Item Type", "Item in selection was not a mail item."
            GoTo NextItem
        End If

        Dim tmpMht As String, pdfFile As String, baseName As String, cleanHtml As String
        On Error Resume Next
        
        ' Build filenames and de-duplicate
        tmpMht = GetUniqueTempMHT(mailItem, ".mht")
        baseName = Format(ItemDate(mailItem), "yyyymmdd-hhnnss") & " – " & CleanFile(mailItem.Subject)
        
        If Len(tgtFolder & baseName & "_99.pdf") >= MAX_PATH Then
            Dim room As Long
            room = MAX_PATH - Len(tgtFolder) - 7
            If room > 0 Then
                baseName = Left$(baseName, room)
            Else
                 LogSkippedItem logFilePath, mailItem.Subject, "Target folder path is too long to create a valid filename."
                 skipped = skipped + 1
                 GoTo NextItem
            End If
        End If
        
        Dim dupCounter As Long
        pdfFile = tgtFolder & baseName & ".pdf"
        dupCounter = 1
        Do While fso.FileExists(pdfFile)
            pdfFile = tgtFolder & baseName & "_" & dupCounter & ".pdf"
            dupCounter = dupCounter + 1
        Loop

        ' *** NEW LOGIC: Trim the HTML body BEFORE saving to MHT ***
        ' This ensures reply history is removed without breaking attachments.
        cleanHtml = StripQuotedBody(mailItem)

        ' ========================================================================
        ' === FIX STEP 2 & 3: DIAGNOSTIC & FAIL-SAFE (as per instructions)     ===
        ' This proves the body is being wiped and prevents Word from receiving
        ' an empty string, which causes the dead-lock.
        ' Debug.Print mailItem.Subject, Len(cleanHtml)
        
        If Len(cleanHtml) < 100 Then
            LogSkippedItem logFilePath, mailItem.Subject, "Body trimmed to <100 chars; skipped."
            skipped = skipped + 1 ' Increment skipped counter
            GoTo NextItem
        End If
        ' === END OF FIX ===
        ' ========================================================================
        
        If Len(cleanHtml) > 0 Then
            '— make sure the HTML Word gets is well-formed —
            cleanHtml = "<html><body>" & cleanHtml & "</body></html>"
            Call SaveHtmlToMht(cleanHtml, tmpMht, wrd)
        Else
            ' Fallback for empty or problematic HTML bodies
            mailItem.SaveAs tmpMht, olMHTML
        End If
        
        If Err.Number <> 0 Then
            Err.Clear
            LogSkippedItem logFilePath, mailItem.Subject, "Failed to save as MHT (IRM protected or locked)."
            skipped = skipped + 1
            GoTo NextItem
        End If
        
        ' *** UPDATE #3: Add colons to numeric labels ***
1000:   Set doc = wrd.Documents.Open(tmpMht, ReadOnly:=True, Visible:=False, ConfirmConversions:=False)
        If Err.Number <> 0 Then
             Err.Clear
             LogSkippedItem logFilePath, mailItem.Subject, "Word failed to open the MHT file."
             skipped = skipped + 1
             GoTo NextItem
        End If

1010:   Call InjectFullHeader(doc, mailItem)
1020:   Call TrimQuotedContent(doc) ' This now serves as a secondary safety net

        '–-–- NEW: force and wait for pagination (as per instructions) –-–-
        doc.Repaginate                          ' Forces layout
        DoEvents
        Do While doc.ActiveWindow.Panes(1).Pages.Count = 0: DoEvents: Loop
        '–-–- END NEW –-–-

        '— export without carrying IRM over —
        Call TidyAndTrimDocument(doc)
        doc.ExportAsFixedFormat pdfFile, wdExportFormatPDF, _
                OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, _
                KeepIRM:=False
        
        If Err.Number <> 0 Then
            LogSkippedItem logFilePath, mailItem.Subject, "Word failed to export MHT to PDF. Error: " & Err.Description
            skipped = skipped + 1
            Err.Clear
        Else
            done = done + 1
        End If

NextItem:
        ' Per-item cleanup
        If Not doc Is Nothing Then
            doc.Close False
            Set doc = Nothing
        End If
        If Len(tmpMht) > 0 And fso.FileExists(tmpMht) Then
            fso.DeleteFile tmpMht, True
        End If
        Set mailItem = Nothing
        tmpMht = ""
    Next item

    ' --- FINAL MESSAGE ---
    Dim msg As String
    msg = done & " mail(s) successfully saved as PDF to " & vbCrLf & tgtFolder
    If skipped > 0 Then
        msg = msg & vbCrLf & vbCrLf & skipped & " item(s) were skipped. See the log file for details:" & vbCrLf & logFilePath
    End If
    MsgBox msg, vbInformation, "Export Complete"

Cleanup:
    ' -- This block is now fully robust --
    On Error Resume Next
    
    ' Safely clear Word's status bar
    If Not wrd Is Nothing Then
        wrd.StatusBar = False
    End If

    ' *** UPDATED: Safely clear Outlook's status bar using the new helper function ***
    Call SafeClearStatusBar
    
    ' Safely quit Word and release all objects
    If Not wrd Is Nothing Then wrd.Quit
    
    Set wrd = Nothing: Set fso = Nothing: Set sel = Nothing
    Set doc = Nothing: Set mailItem = Nothing: Set objWord = Nothing
    Exit Sub

WordCreationFailed:
    MsgBox "The macro failed to initialize Microsoft Word." & vbCrLf & vbCrLf & _
           "Failing step: " & failedStep & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "This usually indicates a problem with the Office installation or registry.", vbCritical, "Word Initialization Failed"
    GoTo Cleanup

ErrorHandler:
    MsgBox "A critical error occurred on line " & Erl & "." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Macro Error"
GoTo Cleanup
End Sub
'--- Convenience wrapper to match original examples ---------------------------
Public Sub SaveAsPDF()
    Call SaveAsPDFfile
End Sub
