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
    patterns = Array( _
        "<div class=""OutlookMessageHeader"">", _
        "<hr", _
        "-----Original Message-----", _
        "<div class=3D""gmail_quote"">", _
        "<blockquote>" _
    )
    
    ' Using 0 as "not found"
    firstSeparatorPos = 0

    ' Loop through each pattern to find the one that appears EARLIEST
    For Each pat In patterns
        pos = InStr(1, body, CStr(pat), vbTextCompare)
        If pos > 0 Then
            If firstSeparatorPos = 0 Or pos < firstSeparatorPos Then
                firstSeparatorPos = pos
            End If
        End If
    Next pat
    
    ' If a separator was found, truncate the string. Otherwise, return the original.
    If firstSeparatorPos > 0 Then
        StripQuotedBody = Left(body, firstSeparatorPos - 1)
    Else
        StripQuotedBody = body
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : TrimQuotedContent (Version 6 - Robust and Corrected)
' Author    : sebvannistel / 2025-06-21
' Purpose   : Finds the EARLIEST reply separator in the document across multiple
'             patterns and deletes all content from that point forward.
'---------------------------------------------------------------------------------------
Private Sub TrimQuotedContent(ByVal doc As Object)
    On Error Resume Next
    
    ' FIX #2: Add wdFindStop constant
    Const wdFindStop As Long = 0
    Dim findRange As Object ' Word.Range
    Dim patterns As Variant
    Dim pat As Variant
    Dim firstSeparatorPos As Long
    
    ' FIX #2: Add new pattern for OWA/Gmail quotes
    patterns = Array( _
        "[-]{5,}Original Message[-]{5,}", _
        "From:*Sent:*To:*Subject:*", _
        "<div class=3D""gmail_quote"">", _
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
            ' FIX #2: Change Wrap to wdFindStop to prevent finding the last separator
            .Wrap = wdFindStop
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
' === FINAL, ROBUST MAIN PROCEDURE (WITH ALL SAFETY CHECKS)                             ===
' =========================================================================================
Sub SaveAsPDFfile()
    ' --- SETUP ---
    Const wdExportFormatPDF As Long = 17
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

    On Error Resume Next ' Temporarily disable the main error handler

    failedStep = "CreateObject(""Word.Application"")"
    Set wrd = CreateObject("Word.Application")
    If Err.Number <> 0 Then GoTo WordCreationFailed

    failedStep = "wrd.Visible = False"
    wrd.Visible = False
    If Err.Number <> 0 Then GoTo WordCreationFailed

    failedStep = "wrd.DisplayAlerts = 0"
    wrd.DisplayAlerts = 0
    If Err.Number <> 0 Then GoTo WordCreationFailed

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

        '🟢 NEW – cut the quoted thread BEFORE saving the .mht
        mailItem.HTMLBody = StripQuotedBody(mailItem)
        'Fallback to plain-text bodies
        If Len(mailItem.HTMLBody) = 0 Then
            mailItem.Body = Split(mailItem.Body, vbCrLf & vbCrLf & "-----")(0)
        End If

        Dim tmpMht As String, pdfFile As String, baseName As String
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

        ' Save to MHT and export
        mailItem.SaveAs tmpMht, olMHTML
        If Err.Number <> 0 Then
            Err.Clear
            LogSkippedItem logFilePath, mailItem.Subject, "Failed to save as MHT (IRM protected or locked)."
            skipped = skipped + 1
            GoTo NextItem
        End If
        
        ' *** UPDATE #3: Add colons to numeric labels ***
1000:   Set doc = wrd.Documents.Open(tmpMht, ReadOnly:=True, Visible:=False)
        If Err.Number <> 0 Then
             Err.Clear
             LogSkippedItem logFilePath, mailItem.Subject, "Word failed to open the MHT file."
             skipped = skipped + 1
             GoTo NextItem
        End If

1010:   Call InjectFullHeader(doc, mailItem)
1020:   Call TrimQuotedContent(doc)
1030:   doc.ExportAsFixedFormat pdfFile, wdExportFormatPDF
        
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
