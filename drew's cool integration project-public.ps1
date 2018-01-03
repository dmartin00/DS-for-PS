Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework


# # # # 
# form commands are: Show-LoginForm, Show-APIForm, Show-TemplateForm, Show-ContactForm
# # # # 


$ikey = "[[ Redacted ]]" # don't forget to use a live key if you need to hit prod
                                               # if you need to go live, use the Contact Form to delete some contacts

$global:Login = $null
[bool]$global:loginerror = $false


function New-DSLogin{ #does a getLoginInfo call
    param(
    [Parameter(Mandatory=$true)]
    [bool]$prod,
    [Parameter(Mandatory=$true)]
    [string]$username,
    [Parameter(Mandatory=$true)]
    [string]$password
        )
        $global:loginerror = $false
        $global:headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
        $global:headers.add("X-DocuSign-Authentication", "<DocuSignCredentials><Username>$username</Username><Password>$password</Password><IntegratorKey>$ikey</IntegratorKey></DocuSignCredentials>")
        $global:headers.add("Content-Type", "application/json")
        $global:headers.add("Cache-Control", "no-cache")
        $global:headers.add("Accept", "application/json")

        if($prod){ $global:uri = "https://www.docusign.net/restapi/v2/login_information/?include_account_id_guid=true" }
        else{ $global:uri = "https://demo.docusign.net/restapi/v2/login_information/?include_account_id_guid=true" }
    $global:prod = $prod        
  try{$global:loginResponse = Invoke-RestMethod -uri $uri -headers $headers -method get}
  catch{ $global:loginerror = $true 
  Write-Error "Failed. Please try again" }
       
    return $global:loginResponse
} #fxn new-dslogin

function Select-Account{ #Gets acct base url
    param(
    [Parameter(Mandatory=$false)]
    [int]$acct
    )#param
    if(!$global:loginResponse){$response = New-DSLogin -prod $false} #assumes demo if not sure, potentially an issue
    if($global:loginResponse){$response = $global:loginResponse}
    if (!$acct){$acct = 0}
$Global:uri = $response.loginAccounts[$acct].baseUrl + "/"
} #fxn sel-acct

function Get-AcctTemplates{ #for use with show-templateform
    param(
   [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$True)]
    [string]$searchString, 
    [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$True)]
    [string]$sortBy
    )#param
    $local:uri = $global:uri + "templates"
        if([bool]$searchString -and [bool]$sortBy){ $local:uri += "?search_text=" + $searchString + "&order_by=" + $sortBy }
        elseif([bool]$searchString -and ![bool]$sortBy){ $local:uri += "?search_text=" + "$searchString" }
        elseif(![bool]$searchString -and [bool]$sortBy){ $local:uri += "?order_by=" + $sortBy }
    $response = Invoke-RestMethod -uri $local:uri -headers $global:headers -method get
    return $response
    }#fxn acctemplates

function Get-TemplateBody{ #currently unused
    param(
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
    [string]$templateId
    )#param
    $local:uri = $global:uri + "templates/" + $templateId + "?include=recipients"
    return Invoke-RestMethod -uri $local:uri -headers $global:headers -method get
}#fxn


# # # #
#    Login Form
# # # #


function Show-LoginForm{
$LForm = New-Object system.Windows.Forms.Form
$LForm.Text = "DS4PS Login"
$LForm.TopMost = $true
$LForm.Width = 200
$LForm.Height = 200
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$LForm.Icon = $Icon

$prodCheckBox = New-Object system.windows.Forms.CheckBox
$prodCheckBox.Text = "prod?"
$prodCheckBox.AutoSize = $true
$prodCheckBox.location = new-object system.drawing.point(80,5)
$LForm.controls.add($prodcheckbox)

$UsernameLabel = New-Object system.windows.Forms.Label
$UsernameLabel.Text = "Username"
$UsernameLabel.AutoSize = $true
$UsernameLabel.location = new-object system.drawing.point(10,35)
$LForm.controls.Add($UsernameLabel)

$PasswordLabel = New-Object system.windows.Forms.Label
$PasswordLabel.Text = "Password"
$PasswordLabel.AutoSize = $true
$PasswordLabel.location = new-object system.drawing.point(10,75)
$LForm.controls.Add($PasswordLabel)

$UsernameTextBox = New-Object system.windows.Forms.TextBox
$UsernameTextBox.location = new-object system.drawing.point(80,35)
$LForm.controls.Add($UsernameTextBox)

$PasswordTextBox = New-Object system.windows.Forms.TextBox
$PasswordTextBox.UseSystemPasswordChar = $true
$PasswordTextBox.location = new-object system.drawing.point(80,75)
$LForm.controls.Add($PasswordTextBox)

$LoginButton = New-Object system.windows.Forms.Button
$LoginButton.Text = "Login"
$LoginButton.Add_Click({
    New-DSLogin -username $UsernameTextBox.Text -password $passwordTextBox.Text -prod $prodcheckbox.Checked
    if(!$global:loginerror){
    Select-Account #sets base url
    
    #set Environment variable
    if($prodcheckbox.Checked){
        [string]$s = $global:uri.trimStart("https://")
        [string]$global:environment = $s.Substring(0, $s.IndexOf("."))
        }#ifprod
    else{[string]$global:environment = "demo"}
    $LForm.Dispose()
    }
}) #loginclick
$LoginButton.location = new-object system.drawing.point(60,100)
$LForm.controls.Add($LoginButton)

[void]$LForm.ShowDialog()
$LForm.Dispose()

}#fxn LoginForm

function Get-ApiLogs{
    param()
    $local:uri = "https://" + $global:environment.tostring() + ".docusign.net/restapi/v2/diagnostics/request_logs"
    $local:headers = $global:headers
    $local:headers.Accept = "application/zip"
    $response = Invoke-RestMethod -uri $local:uri -headers $local:headers -method get -OutFile $env:HOMEPATH\logs.zip
    Invoke-Item $env:HOMEPATH\logs.zip
    return $response
}#fxn get-log

function Reset-ApiLogs{
param(    )#param
    $local:uri = "https://" + $global:environment.tostring() + ".docusign.net/"
    
    $deluri = $local:uri + "restapi/v2/diagnostics/request_logs"
    Invoke-RestMethod -uri $deluri -headers $global:headers -method delete
    $seturi = $local:uri + "restapi/v2/diagnostics/settings"
    $body = @"
{"apiRequestLogging":"True"}    
"@
    Invoke-RestMethod -uri $seturi -headers $global:headers -method put -Body $body
}#fxn Reset-ApiLogs


#####
# API Form
#####


function Show-APIForm{
$AForm = New-Object system.Windows.Forms.Form
$AForm.Text = "DS4PS - API Log Manager"
$AForm.TopMost = $true
$AForm.Width = 200
$AForm.Height = 200
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$AForm.Icon = $Icon

if(!$global:headers){Show-LoginForm} #if no headers, do a login 

$getLogs = New-Object system.windows.Forms.Button
$getLogs.Text = "getLogs"
$getLogs.Width = 60
$getLogs.Height = 30
$getLogs.Add_Click({
Get-ApiLogs
})
$getLogs.location = new-object system.drawing.point(25,50)

$AForm.controls.Add($getLogs)

$userInfo = New-Object system.windows.Forms.Label
$userInfo.Text = ""
$userInfo.Width = 25
$userInfo.Height = 10
$userInfo.location = new-object system.drawing.point(25,10)
$AForm.controls.Add($userInfo)

$clearLogs = New-Object system.windows.Forms.Button
$clearLogs.Text = "reset logs"
$clearLogs.Width = 60
$clearLogs.Height = 30
$clearLogs.Add_Click({
Reset-ApiLogs
})
$clearLogs.location = new-object system.drawing.point(100,50)

$AForm.controls.Add($clearLogs)

[void]$AForm.ShowDialog()
$AForm.Dispose()
}#fxn show form


function New-Account {  #Creates an account 'object' when given all the parameters of an account
    Param(              #used in the template form, mostly made redundant

    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
    [string]$isDefault,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
    [string]$name,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
    [string]$accountId,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
    [string]$baseUrl,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
    [string]$accountIdGuid,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
    [string]$userId,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
    [string]$userName
    )#param
    Process{
        $accountObject = new-object psobject
        $accountObject | add-member -type NoteProperty -name Name -Value $name
        $accountObject | add-member -type NoteProperty -name accountId -Value $accountId
        $accountObject | add-member -type NoteProperty -name accountIdGuid -Value $accountIdGuid
        $accountObject | add-member -type NoteProperty -name isDefault -Value $isDefault
        $accountObject | add-member -type NoteProperty -name userId -Value $userId 
        $accountObject | add-member -type NoteProperty -name userName -Value $userName
        $accountObject | add-member -type NoteProperty -name baseurl -Value $baseurl
        $accountObject | add-member scriptmethod tostring {$this.accountId + " - " + $this.name} -force
     return $accountObject
    }#proc
    
 }#function


# # # #
# Create the Template Form
# # # # 

function Show-TemplateForm{
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "DocuSign for PowerShell"
$Form.AutoSize = $true
#$Form.AutoSizeMode = "GrowAndShrink"
#$Form.SizeGripStyle = "Hide"
$Form.StartPosition = "CenterScreen"
$Form.Width = 500
$Form.Height = 300

[int]$buttonX = "255"
[int]$columnX = "25"

$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$Form.Icon = $Icon

$unameLabel = New-Object System.Windows.Forms.Label
$unameLabel.Text = "Username"
$unameLabel.Location = new-object system.drawing.point($columnX,10)
$Form.Controls.Add($unameLabel)

$unameText = New-Object System.Windows.Forms.TextBox
$unameText.location = new-object system.drawing.point($columnX,35)
$Form.Controls.Add($unameText)

$pwLabel = New-Object System.Windows.Forms.Label
$pwLabel.Text = "Password"
$pwLabel.location = New-Object system.drawing.point(150,10)
$Form.Controls.Add($pwLabel)

$pwText = New-Object System.Windows.Forms.TextBox
$pwText.UseSystemPasswordChar = $true
$pwText.Location = new-object system.drawing.point(150,35)
$form.Controls.add($pwText)

$demoCheckBox = New-Object System.Windows.Forms.CheckBox
$demoCheckBox.Text = "Prod"
$demoCheckBox.Location = new-object System.Drawing.Point($buttonX,5)
#$demoCheckBox.enabled = $false #todo: go live ## we live now
$Form.Controls.Add($demoCheckBox)

<#
$responseText = new-object System.Windows.Forms.TextBox
$responseText.Multiline = $true
$responseText.location = new-object system.drawing.point(25,55)
$responseText.AutoSize = $true
$responseText.Width = "250"
$responseText.Height = "250"
$Form.Controls.Add($responseText)
#>


# # # # # 
# Login for TemplateForm  #  Note this doesn't use the API login form because I wrote this first and it sucks
# # # # #  But it does give you an account picker!

$loginButton = new-object System.Windows.Forms.Button
$loginButton.text = "Login"
$loginButton.Location = New-Object system.drawing.point ($buttonX,33)
$loginButton.Add_click({
    $global:Login = New-DSLogin -username $unameText.Text -prod $demoCheckBox.Checked -password $pwText.Text
    [int]$inc = "0"
    foreach($item in $Login.loginaccounts)
        {
        $entry = New-Account -name $item.name -accountId $item.accountid -accountIdGuid $item.accountidguid -isDefault $item.isdefault -baseurl $item.baseurl -username $item.username -userid $item.userid
        $acctDropDown.items.add($inc.ToString() + " - " + $entry.toString())
        $inc += 1
        } #foreach
    $Form.Controls.Add($acctDropDown)
    $Form.Controls.Add($selectAcctButton)
    })#login click
$Form.Controls.Add($loginButton)

# # # # 
# Select Account
# # # # 

$acctDropDown = New-Object System.Windows.Forms.ComboBox
$acctDropDown.location = new-object system.drawing.point($columnX,55)
$acctDropDown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList

$selectAcctButton = new-object System.Windows.Forms.Button
$selectAcctButton.text = "Select Acct"
$selectAcctButton.Location = New-Object system.drawing.point ($buttonX,55)
$selectAcctButton.add_click({
#write-host $global:Login.loginaccounts[$acctDropDown.SelectedIndex].name
#[System.Windows.MessageBox]::Show("you have selected account " + $global:ActiveAccount.name + "`nWhich has AccountId: " + $global:activeAccount.accountId)
#$selectedAcctLabel.text = $activeAccount.ToString()

    $global:ActiveAccount = New-Account -name $global:Login.loginaccounts[$acctDropDown.SelectedIndex].name -accountId $global:Login.loginaccounts[$acctDropDown.SelectedIndex].accountid -accountIdGuid $global:Login.loginaccounts[$acctDropDown.SelectedIndex].accountidguid -isDefault $global:Login.loginaccounts[$acctDropDown.SelectedIndex].isdefault -baseurl $global:Login.loginaccounts[$acctDropDown.SelectedIndex].baseurl -username $global:Login.loginaccounts[$acctDropDown.SelectedIndex].username -userid $global:Login.loginaccounts[$acctDropDown.SelectedIndex].userid
    $Form.Text += " - " + $activeAccount.ToString()
    $global:uri = $activeAccount.baseurl + "/"
    $Form.Controls.Add($sortCombo)
    $Form.Controls.Add($sortLabel)
    $Form.Controls.Add($searchLabel)
    $Form.Controls.Add($searchText)
    $Form.Controls.Add($templateButton)

})#select button click

# # # # 
# Show template options
# # # # 

$sortLabel = New-Object System.Windows.Forms.Label
$sortLabel.text = "Sort by:"
$sortLabel.Size = New-Object System.Drawing.Size (100,20)
$sortLabel.Location = New-Object system.drawing.point ($columnX,75)

$sortCombo = New-Object System.Windows.Forms.ComboBox
$sortCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$sortCombo.location = New-Object system.drawing.point ($columnX,95)
$sortCombo.items.add("Name") > $null   # prevent echoing 0 1 2 on startup
$sortCombo.items.add("Modified")  > $null
$sortCombo.items.add("Used") > $null

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Optional - Search:"
$searchLabel.Location = new-object System.Drawing.Point (150,75)
$searchLabel.Size = New-Object System.Drawing.Size (100,20)

$searchText = New-Object System.Windows.Forms.TextBox
$searchText.Location  = new-object System.Drawing.Point (150,95)

$templateButton = New-Object System.Windows.Forms.Button
$templateButton.Text = "Get Templates"
$templateButton.Location = New-Object system.drawing.point ($buttonX, 95)
$templateButton.Add_click({
$Form.Controls.Add($templateLabel)
$global:templateList = Get-AcctTemplates -sortBy $sortCombo.SelectedText -searchString $searchText.text
    foreach($item in $templateList.envelopeTemplates){
        $templateCombo.items.add($item.name)
    } #forEach Template
    $form.Controls.add($templateCombo)
})#template click

# # # #
# Process the templates
# # # #

$templateLabel = New-Object System.Windows.Forms.Label
$templateLabel.Text = "Select a template:"
$templateLabel.Location  = new-object System.Drawing.Point ($columnX,115)
$templateLabel.Size = New-Object System.Drawing.Size (100,10)

$templateCombo = New-Object System.Windows.Forms.ComboBox
$templateCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$templateCombo.location = New-Object system.drawing.point ($columnX,130)

# # # #
# Display the form
# # # #

$Form.ShowDialog()
$Form.Dispose()
}


# # # # 
# Contact Manager form
# # # #

Function Search-Contacts{
Param(
[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
    [string]$keyword)

    $local:uri = $global:uri + "contacts/?keyword=" + $keyword
    return Invoke-RestMethod -uri $local:uri -headers $global:headers -method get
}#fxn Search 

Function Delete-Contact{
param(
[Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$True)]
    [string]$id
    )
        $local:uri = $global:uri + "contacts/" + $id
    return Invoke-RestMethod -uri $local:uri -headers $global:headers -method delete
}#fxn Delcon


Function Show-ContactForm{

$ContactForm = New-Object system.Windows.Forms.Form
$ContactForm.Text = "DS4PS - Contact Form"
$ContactForm.MinimumSize = new-object system.drawing.size(775,450)
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$ContactForm.Icon = $Icon

if(!$global:headers){Show-LoginForm} #if no headers, do a login 

$searchText = New-Object system.windows.Forms.TextBox
$searchText.Width = 100
$searchText.Height = 20
$searchText.location = new-object system.drawing.point(25,20)
$ContactForm.controls.Add($searchText)

$searchButton = New-Object system.windows.Forms.Button
$searchButton.Text = "Search"
$searchButton.Width = 75
$searchButton.Height = 20
$searchButton.Add_Click({
    $response = Search-Contacts -keyword $searchText.text
    foreach($item in $response.contacts){
    $searchOut.text += $item.contactId + " - " + $item.name + " - " + $item.emails + " `r`n"
    }# contacts loop
    
})#searchclick
$searchButton.location = new-object system.drawing.point(150,20)
$ContactForm.controls.Add($searchButton)

$clearButton = New-Object system.windows.Forms.Button
$clearButton.Text = "Clear"
$clearButton.Width = 75
$clearButton.Height = 20
$clearButton.Add_Click({
    $searchout.text = ""
})#clearclick
$clearButton.location = new-object system.drawing.point(225,20)
$ContactForm.controls.Add($clearButton)

$deleteText = New-Object system.windows.Forms.TextBox
$deleteText.Width = 199
$deleteText.Height = 20
$deleteText.location = new-object system.drawing.point(25,50)
$ContactForm.controls.Add($deleteText)

$deleteButton = New-Object system.windows.Forms.Button
$deleteButton.Text = "Delete"
$deleteButton.Width = 75
$deleteButton.Height = 20
$deleteButton.Add_Click({
    $deleteLabel.text = "Attempting deletion of ID " + $deletetext.Text
    $response = delete-contact -id $deletetext.text
    $deleteLabel.text = "Complete - ID: " + $response.contacts.contactId
})
$deleteButton.location = new-object system.drawing.point(225,50)
$ContactForm.controls.Add($deleteButton)

$deleteLabel = New-Object system.windows.Forms.Label
$deleteLabel.location = new-object system.drawing.point(325,50)
$deleteLabel.Height = 50
$deleteLabel.Width = 375
#$deleteLabel.dock = [System.Windows.Forms.DockStyle]::Right
$deleteLabel.Text = "Ready!"
$ContactForm.controls.Add($deleteLabel)

$searchOut = New-Object system.windows.Forms.TextBox
$searchOut.Multiline = $true
$searchOut.Text = ""
$searchOut.Width = 700
$searchOut.Height = 300
#$searchOut.Anchor = [System.Windows.Forms.AnchorStyles]::Top

$searchOut.Dock = [System.Windows.Forms.DockStyle]::Bottom
$searchOut.Font = "Lucida Console,8"
$searchOut.location = new-object system.drawing.point(25,90)
$ContactForm.controls.Add($searchOut)

[void]$ContactForm.ShowDialog()
$ContactForm.Dispose()
}


# Show-APIForm #added for exe compiling
