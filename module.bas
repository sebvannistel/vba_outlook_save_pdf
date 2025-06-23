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

'------ Helper: Tidy final Word doc and trim quoted text/footers (UNIVERSAL LATE BINDING VERSION) --
Private Sub TidyAndTrimDocument(wdDoc As Object)
    ' --- Define Word constants for late binding ---
    Const wdReplaceAll As Long = 2
    Const wdStyleNormal As Long = -1
    Const wdStyleHeading1 As Long = -2
    Const wdFindStop As Long = 0

    On Error Resume Next ' In case of errors during styling

    '------ 1. Apply Basic Formatting (Optional) --------------------------------
    ' Comment this out if you don't have a corporate template
    ' wdDoc.ApplyDocumentTemplate "C:\Path\To\Your\Brand.dotx"

    With wdDoc.Content.Font
        .Name = "Calibri"
        .Size = 11
    End With
    With wdDoc.Styles(wdStyleNormal).Font
        .Name = "Calibri"
        .Size = 11
    End With

    '------ 2. Universally Find and Trim Quoted Replies --------------------------
    Dim findRange As Object ' Word.Range
    Dim patterns As Variant
    Dim pat As Variant
    Dim firstCutPos As Long
    
    ' This array contains universal WILDCARD patterns.
    ' They are processed in order, but the code finds the EARLIEST match overall.
    patterns = Array( _
        "[-_]{5,}Original Message[-_]{5,}", _
        "From:?*Sent:?*To:?*Subject:?*", _
        "<blockquote*>", _
        "<hr*>" _
    )
    
    firstCutPos = -1 ' Initialize to a "not found" state

    ' Loop through each pattern to find the one that appears EARLIEST in the document
    For Each pat In patterns
        Set findRange = wdDoc.Content
        With findRange.Find
            .ClearFormatting
            .Text = pat
            .Forward = True
            .Wrap = wdFindStop ' IMPORTANT: Do not loop around the document
            .Format = False
            .MatchCase = False
            .MatchWildcards = True ' THE MAGIC SWITCH!
            
            If .Execute = True Then
                ' *** CRITICAL SAFETY CHECK ***
                ' Do not trim if the match is in the first 400 characters.
                ' This prevents the macro from matching the *main email's own header*
                ' and deleting the entire message body.
                If findRange.Start > 400 Then
                    ' If this is the first separator found, or if it's earlier
                    ' than the previous best, record its position.
                    If firstCutPos = -1 Or findRange.Start < firstCutPos Then
                        firstCutPos = findRange.Start
                    End If
                End If
            End If
        End With
    Next pat

    ' After checking all patterns, if we found a valid separator, delete from that point.
    If firstCutPos > -1 Then
        wdDoc.Range(Start:=firstCutPos, End:=wdDoc.Content.End).Delete
    End If
    
    '------ 3. Compact Extra Blank Lines Left Behind --------------------
    wdDoc.Range.ParagraphFormat.SpaceAfter = 0
    With wdDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = vbCr & vbCr
        .Replacement.Text = vbCr
        .MatchWildcards = False ' Turn off wildcards for this simple find/replace
        .Wrap = 1 ' wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With

    Set findRange = Nothing
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
' === FINAL, ROBUST MAIN PROCEDURE (COMBINING BEST STRATEGIES)                          ===
' =========================================================================================
Public Sub SaveAsPDFfile()
    ' --- SETUP ---
    Const MAX_PATH As Long = 259

    ' --- OBJECTS & VARIABLES ---
    Dim sel As Outlook.Selection
    Dim wrd As Object, doc As Object, fso As Object
    Dim mailItem As Outlook.MailItem
    Dim tgtFolder As String, logFilePath As String
    Dim done As Long, skipped As Long, total As Long

    'On Error GoTo ErrorHandler

    ' Step 1: Get target folder
    tgtFolder = GetTargetFolder_Universal()
    If Len(tgtFolder) = 0 Then Exit Sub

    ' Step 2: Get and de-duplicate selections
    If Application.ActiveExplorer Is Nothing Then
        MsgBox "Cannot run macro. Please select emails in the main Outlook window.", vbExclamation, "No Active Window"
        GoTo Cleanup
    End If

    Set sel = Application.ActiveExplorer.Selection
    If sel.Count = 0 Then
        MsgBox "Please select one or more emails to save.", vbInformation, "No Items Selected"
        GoTo Cleanup
    End If

    Dim latest As Object: Set latest = CreateObject("Scripting.Dictionary")
    If latest Is Nothing Then
        MsgBox "Could not create Scripting.Dictionary.", vbCritical
        Exit Sub
    End If
    Dim k As String, it As Object
    For Each it In sel
        k = it.EntryID
        If Not latest.Exists(k) Then latest.Add k, it
    Next it
    total = latest.Count
    If total = 0 Then
        MsgBox "No unique items to process.", vbInformation
        GoTo Cleanup
    End If

    ' Step 3: Initialize worker objects
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFilePath = tgtFolder & "_SkippedItems_" & Format(Now, "yyyymmdd_hhnnss") & ".log"

    ' --- "FORT KNOX" WORD SETUP (from module3) ---
    Dim failedStep As String
    Const msoAutomationSecurityForceDisable As Long = 3
    On Error Resume Next
    failedStep = "CreateObject(""Word.Application"")"
    Set wrd = CreateObject("Word.Application")
    If Err.Number <> 0 Then GoTo WordCreationFailed
    
    failedStep = "wrd.Visible = False"
    wrd.Visible = False
    
    failedStep = "wrd.DisplayAlerts = 0"
    wrd.DisplayAlerts = 0
    
    failedStep = "wrd.Options.UpdateLinksAtOpen = False"
    wrd.Options.UpdateLinksAtOpen = False
    
    failedStep = "wrd.AutomationSecurity = msoAutomationSecurityForceDisable"
    wrd.AutomationSecurity = msoAutomationSecurityForceDisable
    If Err.Number <> 0 Then GoTo WordCreationFailed
    On Error GoTo ErrorHandler
    ' --- END OF WORD SETUP ---

    '================================================================================
    '--- MAIN EXPORT LOOP ---
    '================================================================================
    Dim item As Variant
    Dim progressCounter As Long
    
    If Not Application.ActiveExplorer Is Nothing Then
        Application.ActiveExplorer.StatusBar = "Preparing to save " & total & " selected email(s)..."
    End If

    For Each item In latest.Items
        progressCounter = progressCounter + 1
        DoEvents
        
        If Not Application.ActiveExplorer Is Nothing Then
            Application.ActiveExplorer.StatusBar = "Processing " & progressCounter & " of " & total & "..."
        End If
        
        If Not TypeOf item Is Outlook.MailItem Then
            skipped = skipped + 1
            LogSkippedItem logFilePath, "Unknown Item Type", "Item was not a mail item."
            GoTo NextItem
        End If
        Set mailItem = item

        Dim tmpMht As String, pdfFile As String, baseName As String
        On Error Resume Next
        
        ' --- FILENAME LOGIC ---
        tmpMht = GetUniqueTempMHT(mailItem, ".mht")
        baseName = Format(ItemDate(mailItem), "yyyymmdd-hhnnss") & " – " & CleanFile(mailItem.Subject)
        
        If Len(tgtFolder & baseName & "_99.pdf") >= MAX_PATH Then
             baseName = Left$(baseName, MAX_PATH - Len(tgtFolder) - 7)
        End If
        
        Dim dupCounter As Long
        pdfFile = tgtFolder & baseName & ".pdf"
        dupCounter = 1
        Do While fso.FileExists(pdfFile)
            pdfFile = tgtFolder & baseName & "_" & dupCounter & ".pdf"
            dupCounter = dupCounter + 1
        Loop

        ' --- CORE LOGIC: SAVE FULL MHT FIRST (module2 strategy) ---
        mailItem.SaveAs tmpMht, olMHTML
        If Err.Number <> 0 Then
            Err.Clear
            LogSkippedItem logFilePath, mailItem.Subject, "Failed to save as MHT (IRM protected or locked)."
            skipped = skipped + 1
            GoTo NextItem
        End If
        
        Set doc = wrd.Documents.Open(tmpMht, ReadOnly:=True, Visible:=False, ConfirmConversions:=False)
        If Err.Number <> 0 Then
             Err.Clear
             LogSkippedItem logFilePath, mailItem.Subject, "Word failed to open the MHT file."
             skipped = skipped + 1
             GoTo NextItem
        End If

        ' --- PROCESSING IN WORD ---
        Call InjectFullHeader(doc, mailItem) ' Inject the Outlook header
        
        Call TidyAndTrimDocument(doc) ' <<<<<<<< CALLING YOUR NEW, FIXED FUNCTION
        
        ' --- EXPORT TO PDF ---
        doc.ExportAsFixedFormat pdfFile, wdExportFormatPDF, OpenAfterExport:=False, KeepIRM:=False
        
        If Err.Number <> 0 Then
            LogSkippedItem logFilePath, mailItem.Subject, "Word failed to export to PDF. Error: " & Err.Description
            skipped = skipped + 1
            Err.Clear
        Else
            done = done + 1
        End If

NextItem:
        ' --- PER-ITEM CLEANUP ---
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
    msg = done & " mail(s) successfully saved as PDF." & vbCrLf & "Folder: " & tgtFolder
    If skipped > 0 Then
        msg = msg & vbCrLf & vbCrLf & skipped & " item(s) were skipped. See log for details:" & vbCrLf & logFilePath
    End If
    MsgBox msg, vbInformation, "Export Complete"

Cleanup:
    On Error Resume Next
    Call SafeClearStatusBar
    If Not wrd Is Nothing Then
        wrd.StatusBar = False
        wrd.Quit
    End If
    Set wrd = Nothing: Set fso = Nothing: Set sel = Nothing
    Set doc = Nothing: Set mailItem = Nothing
    Exit Sub

WordCreationFailed:
    MsgBox "Failed to initialize Microsoft Word." & vbCrLf & vbCrLf & _
           "Failing step: " & failedStep & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Word Initialization Failed"
    GoTo Cleanup

ErrorHandler:
    MsgBox "A critical error occurred." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Macro Error"
    GoTo Cleanup
End Sub

'--- Convenience wrapper to match original examples ---------------------------
Public Sub SaveAsPDF()
    Call SaveAsPDFfile
End Sub
