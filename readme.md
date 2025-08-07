Note: This project was forked from cavo789/vba_outlook_save_pdf. It just has some updated instructions and resolves an issue I was having with the VBA.

# Enhanced Guide: Saving Selected Emails as PDFs in Outlook

![Enhanced Banner](./banner.svg)

> Elevate your Outlook experience: seamlessly convert and save emails as PDFs directly to your disk with our macro.

## Overview

Imagine effortlessly saving one or multiple emails from your Outlook inbox directly as PDFs with just a few clicks. This guide introduces a custom-built macro that integrates with your Outlook client, enabling you to save selected emails into a designated folder on your hard drive. Whether you're archiving important correspondence or organizing project-related emails, this tool simplifies the process, allowing for up to 250 emails to be saved as PDFs in a swift operation.

## Contents

- [Installation Guide](#installation-guide)
- [Enable Macros in Outlook](#enable-macros-in-outlook)
- [How to Use](#usage)
- [License Information](#license)
- [Access the Code Directly!](./module.bas)

## Installation Guide

Transform your Outlook into a powerful email-to-PDF converter in a few simple steps:

1. **Integrating the VBA Code**:
   - Activate the Visual Basic Editor in Outlook by pressing `ALT-F11`.
   - Navigate to Menu Bar -> Insert, then choose "Module" to create a new module.
   - Obtain the `module.bas` script from [here](./module.bas) and paste it into the newly created module.
   - Exit the Visual Basic Editor.

2. **Adding a Custom Button to the Ribbon**:
   - Customize your Outlook Ribbon by right-clicking it and selecting `Customize The Ribbon`.
   - In the "Choose commands from:" dropdown, select Macro.
   - Add a `New Tab` (with a New Group within it) on the right side.
   - Locate `Project1.SaveAsPDFfile` (or the background-only `Project1.SaveSelectedMails_AsPDF_NoPopups`) in the command list, select it, and hit the `Add` button.
   - Confirm your changes by clicking OK.

3. **Customizing Your Ribbon** (Optional):
   - Personalize the name of the new ribbon tab for easy access.

**Prerequisite**: Ensure Microsoft Word is installed on your computer for the macro to function.

## Enable Macros in Outlook

Make sure Outlook actually loads the macros. By default Outlook disables them for security reasons.

1. Navigate to `File ▶ Options ▶ Trust Center ▶ Trust Center Settings ▶ Macro Settings`.
   - For testing you can choose **Enable all macros** (tighten this later). See [support.microsoft.com](https://support.microsoft.com) for more details.
2. For long-term use create a self-signed certificate with `SelfCert.exe`, add the signature in `VBE ▶ Tools ▶ Digital Signature…`, then switch security to **Disable all except digitally-signed** so Outlook trusts it automatically. See [learn.microsoft.com](https://learn.microsoft.com) for the procedure.
3. If you're running the "New Outlook" preview, note that it doesn't support VBA macros at all. Switch back to "Classic Outlook" if you want the macros to persist. See [learn.microsoft.com](https://learn.microsoft.com) for more information.

## How to Use

1. **Select the Emails**: Choose one or multiple emails you wish to save as PDFs.
2. **Activate the Macro**: Click on the `SaveAsPDFfile` button (or use `SaveSelectedMails_AsPDF_NoPopups` for silent exports) on your custom ribbon.
3. **Specify Save Location**: Follow the prompts to select a destination folder for the PDFs and any attachments.
4. **Completion**: Sit back and watch as your selected emails are transformed and saved as PDFs on your disk, with their attachments saved alongside them.

![Demonstration](images/demo.gif)

## Special message types

Certain Outlook messages are worth keeping in their own PDF even if they share a
conversation topic.  The macro uses their `MessageClass` (and sometimes a
subject prefix) to detect them:

1. **Message recall** – `IPM.Outlook.Recall` or `IPM.Recall.Report`
2. **Delivery-status notices** – `REPORT.IPM.NOTE.*` or `REPORT.IPM.SCHEDULE.*`
3. **Read / Non-read receipts** – `REPORT.IPM.NOTE.IPNRN` / `REPORT.IPM.NOTE.IPNNRN`
4. **Out-of-office / automatic replies** – subject includes `Automatic reply:` or class `IPM.Note.Rules.OofTemplate*`
5. **Meeting-related mail** – `IPM.Schedule.Meeting.*`
6. **Voting-button responses** – the reply subject begins with the chosen option
7. **Task delegation & updates** – `IPM.TaskRequest.*`
8. **Encrypted S/MIME reports** – `REPORT.IPM.NOTE.SMIME.*`


## ⚠️ Troubleshooting: Resolving Blank PDFs and Freezes

In some environments, especially with corporate security policies, Adobe Acrobat, or custom add-ins, the macro may appear to run but produce blank PDF files or cause Outlook/Word to freeze. This is typically due to Microsoft Word's security features silently blocking the automation process.

If you encounter this issue, follow these steps to configure Word's Trust Center. **These changes make Word more permissive for automation but should be understood before applying.**

1. **Open Microsoft Word** (not Outlook).
2. Go to `File > Options > Trust Center > Trust Center Settings...`.
3. Make the following adjustments in the left-hand menu:

    * #### **1. Disable Conflicting Add-ins (Most Common Fix)**
        * Go to **Add-ins**.
        * At the bottom, next to `Manage:`, select **COM Add-ins** and click **Go...**.
        * **Uncheck** any add-ins related to PDF creation (e.g., **"Acrobat PDFMaker Office COM Addin"**) or document management. These often conflict with Word's native PDF export.
        * Click **OK**.

    * #### **2. Disable Protected View (Crucial for Automation)**
        * Go to **Protected View**.
        * **Uncheck all three boxes**:
            * `[ ] Enable Protected View for files originating from the Internet`
            * `[ ] Enable Protected View for files located in potentially unsafe locations`
            * `[ ] Enable Protected View for Outlook attachments`
        * *Reason: The macro creates temporary files that Word may open in a restricted "sandbox" mode, preventing the PDF conversion. Disabling this allows the automation to run unimpeded.*

    * #### **3. Adjust ActiveX Settings**
        * Go to **ActiveX Settings**.
        * Select **"Disable all controls without notification"**.
        * *Reason: This prevents Word from hanging on a hidden prompt if an email contains an embedded ActiveX control.*

4. Click **OK** to close the Trust Center and Word Options.
5. **Restart both Word and Outlook** to ensure all changes take effect.

After applying these settings, the macro should now be able to run without interference, correctly generating non-blank PDFs. If issues persist, it may indicate a deeper conflict with enterprise security software (EDR, Antivirus) that requires an IT department to whitelist the automation process.

## License

This tool is freely distributed under the [MIT License](LICENSE), promoting open and unrestricted use while encouraging contributions and modifications.

## Testing

Run the test suite before committing any changes to ensure everything works as expected:

```bash
pytest -q
```

Executing the tests prior to committing helps maintain the project's stability.
