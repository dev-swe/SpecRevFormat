Spec Formatter — Beta install
=============================

Requirements
------------
- Microsoft Word (desktop) on Windows, signed in with your work account.
- You must be on the company network (the add-in loads from a hosted URL and this
  folder lives on a shared drive).

Install (one time)
------------------
1. Double-click  Install-Spec-Formatter-Beta.cmd
     (Windows may warn about an unknown script — choose "More info" > "Run anyway".
      It only writes a per-user registry entry; no admin rights are used.)
2. Fully close Word (every window), then reopen it and open any document.
3. Insert tab  >  Add-ins  >  My Add-ins  >  SHARED FOLDER tab
     >  "Spec Formatter (Beta)"  >  Add.
4. The "Spec Formatter" button appears on the Home ribbon. Click it to open the pane.

If the SHARED FOLDER tab is empty
---------------------------------
File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs.
The entry should be listed; tick "Show in Menu", click OK, restart Word, retry step 3.

Uninstall
---------
Double-click  Uninstall-Spec-Formatter-Beta.cmd  (then close Word).

Feedback
--------
Note the Word version (File > Account) and steps to reproduce any issue, and send them
to Evie. This is a beta build — please don't share it outside the test group.
