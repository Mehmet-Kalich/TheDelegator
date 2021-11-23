Within this repository is the Delegator, an example of a script that I wrote for my job as an IT 
Support Analyst at my last job. 

#The Delegator (An O365 Exchange Delegation Tool)

During a particularly busy period at work, I realised that a great deal of our tickets
were to do with changing and setting different delegation and group delegation permissions 
between members of Business support department for the Exec' they worked for. 

With the IT team extremely busy with other pressing matters, I decided to write a script that would not only 
check the status of delegates of a partner, but also add + remove single and group delegation. 

On top of this, I decided to use the open source Windows Form graphical interface and integrate it with
the Powershell code so that members of my team could use a simple GUI to sort out all of their Delegation tickets
without having to log onto the server, and after inputting their admin credentials, use the Delegator. 

On start-up, the Delegator connects to MsolService, which initiates connecting to Azure Active Directory and prompts
for credentials to change the delegation rights of a user. 

From then on, a list of if statements are used to refer to the Windows Form radio buttons + execute button, and 
to execute various script options (Get-Mailbox Permissions, Add/Set/Remove Mailbox Folder permissions).
After this, the Delegator will make sure to give output to the text box so you're clear exactly with what action
has been taken, but also there is an option to check the Delegation using "Check Delegation" to make sure what 
you have done is correct. 

As a final note, for PS + best security practices, I made sure to add an onexit function which removes the PSSession
from Powershell and sets the execution policy back to Remote Signed.  
