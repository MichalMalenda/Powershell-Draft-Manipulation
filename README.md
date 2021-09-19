# Powershell-Draft-Manipulation
## This is an example module you can use with your automation solutions
## PLEASE KEEP IN MIND THAT THIS CAN ONLY BE USED VIA POWERRSHELL OF THE SAME VERSION AS YOUR OFFICE SUIT VERSION (for 64 bit outlook version normal PowerShell is alright but for 32bit version, you need to use PowerShellx86)
Functions contained in the module are as follows: <br>
1. **New-Draft** This is for creation of a new draft, with a mandatory variable being Subject, it can be set as an empty string if you want to.
    * **-Subject** sets the subject for draft 
    * **-Body** sets the body for the draft
    * **-Recipients** sets the recipients for the draft
2. **Set-Draft** This is for changing an existing draft information, selecting it by the current subject.
    * **bySubject** lets you pick the draft by subject 
    * **AddSubject** lets you add string to the subject
    * **ChangeSubject** lets you replace the old subject with the new one
    * **AddBody** lets you add string to the body(new line adds it after new line which is convenient) 
    * **ChangeBody** lets you replace the old body with the new one
    * **AddRecipients** lets you add other recipients to the message
    * (No change recipients yet)
3. **Get-Draft** This is for checking really if the draft exists
    * **Subject** gives you back infromation about the message with the selected Subject
4. **Send-Draft** This sends the selected message.
    * **bySubject** lets you chose the message that you want to send by the subject.