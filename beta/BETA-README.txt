Spec Formatter — Beta install
=============================

Requirements
------------
- Microsoft Word (desktop) on Windows, signed in with your work account.
- IMPORTANT: this folder must live on a NETWORK SHARE (a \\server\share path).
  Office rejects a local folder (C:\...) as an add-in catalog. Run the installer
  from the shared-drive copy, not from your Desktop/Downloads.

Install (one time)
------------------
1. From the shared-drive copy, double-click  Install-Spec-Formatter-Beta.cmd
     (Windows may warn about an unknown script — choose "More info" > "Run anyway".
      It only writes a per-user registry entry; no admin rights are used. This
      registers the catalog for you — you do NOT need to touch the Trust Center.)
2. Fully close Word (every window), then reopen it and open any document.
3. Home tab  >  Add-ins  >  Advanced...  >  SHARED FOLDER
     >  "Spec Formatter (Beta)"  >  Add.
     (The new Add-ins panel only lists Store/Developer add-ins on its front page;
      the "Advanced..." link opens the dialog that has the Shared Folder section.
      On older Word the Add-ins button is on the Insert tab.)
4. The "Spec Formatter" button appears on the Home ribbon. Click it to open the pane.

If the SHARED FOLDER section is empty
-------------------------------------
File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs.
The entry should be listed; tick "Show in Menu", click OK, restart Word, retry step 3.

If an OLD "Spec Formatter" button is stuck on the ribbon
-------------------------------------------------------
(Only affects the developer/preview build.) File > Options > Customize Ribbon,
find the old button, and Remove it. The beta shows as "Spec Formatter (Beta)".

Uninstall
---------
Double-click  Uninstall-Spec-Formatter-Beta.cmd  (then close Word).

Feedback
--------
Note the Word version (File > Account) and steps to reproduce any issue, and send them
to Evie. This is a beta build — please don't share it outside the test group.
