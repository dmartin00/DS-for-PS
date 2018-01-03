# DS-for-PS
DocuSign for PowerShell 

# What is this?
A personal project that uses DocuSign's API to do a few things. No SDK involved, instead uses Invoke-RestMethod.

# What's available?
## Built out functions
*Show-APIForm* - Used to grab API logs for a user without having to deal with NDSE's backwash. By design, this overwrites the same file every time. If you need to store multiple sets of logs, be sure to move or rename logs.zip from your %userprofile%\ directory. Reset button performs API calls to reset logs but does not touch local .zip file.

*Show-ContactForm* - Lists and removes contacts in an address book. Built when MAR-21171 prevented contact management for high-volume accounts.

## Additional functions

*Show-LoginForm* - Performs a GetLoginInfo call and records username and correct baseurl in header variables. Intended to be used with the two above commands, somewhat useless on its own.

*Show-TemplateForm* - Gets a list of templates available to a user, allows searching and sorting.
Built before Show-LoginForm, and uses a slightly different method for storing header info. Allows selecting between multiple linked acocunts.
