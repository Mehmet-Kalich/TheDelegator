<# The Delegator - An O365 Exchange Delegation Tool - Mehmet Kalich 15/02/2021
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$TheDelegator                   = New-Object system.Windows.Forms.Form
$TheDelegator.ClientSize        = New-Object System.Drawing.Point(450,500)
$TheDelegator.text              = "The Delegator"
$TheDelegator.TopMost           = $false

$Partner                         = New-Object system.Windows.Forms.TextBox
$Partner.multiline               = $false
$Partner.width                   = 206
$Partner.height                  = 20
$Partner.location                = New-Object System.Drawing.Point(14,35)
$Partner.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Delegate                        = New-Object system.Windows.Forms.TextBox
$Delegate.multiline              = $false
$Delegate.width                  = 208
$Delegate.height                 = 20
$Delegate.location               = New-Object System.Drawing.Point(13,91)
$Delegate.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Execute!"
$Button1.width                   = 351
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(40,145)
$Button1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$OutputBox                       = New-Object System.Windows.Forms.TextBox
$OutputBox.Location              = New-Object System.Drawing.Size (14,180) 
$OutputBox.Size                  = New-Object System.Drawing.Size (425,300) 
$OutputBox.Multiline             = $True
$OutputBox.ScrollBars            = "Vertical" 
$TheDelegator.Controls.Add($OutputBox) 

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Partner"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(13,14)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Delegate"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(14,69)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$CD                              = New-Object system.Windows.Forms.RadioButton
$CD.text                         = "Check Delegates"
$CD.AutoSize                     = $true
$CD.width                        = 89
$CD.height                       = 20
$CD.location                     = New-Object System.Drawing.Point(230,15)
$CD.Font                         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$AG                              = New-Object system.Windows.Forms.RadioButton
$AG.text                         = "Give Group Access"
$AG.AutoSize                     = $true
$AG.width                        = 89
$AG.height                       = 20
$AG.location                     = New-Object System.Drawing.Point(230,40)
$AG.Font                         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$RG                              = New-Object system.Windows.Forms.RadioButton
$RG.text                         = "Remove Group Access"
$RG.AutoSize                     = $true
$RG.width                        = 89
$RG.height                       = 20
$RG.location                     = New-Object System.Drawing.Point(230,65)
$RG.Font                         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$PA                              = New-Object system.Windows.Forms.RadioButton
$PA.text                         = "Give Delegate Access"
$PA.AutoSize                     = $true
$PA.width                        = 89
$PA.height                       = 20
$PA.location                     = New-Object System.Drawing.Point(230,90)
$PA.Font                         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$RA                             = New-Object system.Windows.Forms.RadioButton
$RA.text                        = "Remove Delegate Access"
$RA.AutoSize                    = $true
$RA.width                       = 89
$RA.height                      = 20
$RA.location                    = New-Object System.Drawing.Point(230,115)
$RA.Font                        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TheDelegator.controls.AddRange(@($Partner,$Delegate,$Button1,$Label1,$Label2,$OutputBox))
$TheDelegator.controls.AddRange(@($CD,$AG,$RG,$PA,$RA))
$UserDomain = '@baringa.com'
$Calendar = ':\Calendar'

$Button1.Add_Click({ Execute })
$TheDelegator.Add_Load({ onload })
$TheDelegator.Add_FormClosing({ onexit })
 
function onexit { 
    Remove-PSSession $Session
    Set-ExecutionPolicy RemoteSigned
}

function onload { 

    Connect-MsolService 
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection 
    Import-PSSession $Session -DisableNameChecking 
    Set-ExecutionPolicy RemoteSigned
    clear
}
function Execute {

    if($CD.Checked -eq $true){
       #Check Delegates    
       $DelegateResult = Get-MailboxFolderPermission -Identity ($Partner.Text + ($UserDomain) + ($Calendar)) | Out-String; 
       $OutputBox.text=$DelegateResult
    
    }
    if($AG.Checked -eq $true){
        #Give Group Access 
        Add-MailboxFolderPermission -identity ($Partner.Text + ($UserDomain) + ($Calendar)) -User ("MESG_PartnersCalendarView" + ($UserDomain)) -AccessRights LimitedDetails
        Add-MailboxFolderPermission -identity ($Partner.Text + ($UserDomain) + ($Calendar)) -User ("MESG_Exec-Assistants" + ($UserDomain)) -AccessRights LimitedDetails
        Add-MailboxFolderPermission -identity ($Partner.Text + ($UserDomain) + ($Calendar)) -User ("Recruitment.Administrators" + ($UserDomain)) -AccessRights LimitedDetails | Out-String;
        $OutputBox.Text="Delegation given to MESG_PartnersCalendarView, MESG-Exec-Assistants and Recruitment.Administrators Security Groups" 

    }
    if($RG.Checked -eq $true){
        #Give Group Access 
        Remove-MailboxFolderPermission -identity ($Partner.Text + ($UserDomain) + ($Calendar)) -User ("MESG_PartnersCalendarView" + ($UserDomain)) -Confirm:$false
        Remove-MailboxFolderPermission -identity ($Partner.Text + ($UserDomain) + ($Calendar)) -User ("MESG_Exec-Assistants" + ($UserDomain)) -Confirm:$false
        Remove-MailboxFolderPermission -identity ($Partner.Text + ($UserDomain) + ($Calendar)) -User ("Recruitment.Administrators" + ($UserDomain)) -Confirm:$false | Out-String;
        $OutputBox.Text="Delegate Access Removed for MESG_PartnersCalendarView, MESG-Exec-Assistants and Recruitment.Administrators Security Groups"
            
    }
    if($PA.Checked -eq $true){
        #Set Email Copy access + Publishing Editor Rights
        Add-MailboxFolderPermission -Identity ($Partner.Text + ($UserDomain) + ($Calendar)) -User ($Delegate.text + ($UserDomain)) -AccessRights Editor -SharingPermissionFlags Delegate 
        Set-MailboxFolderPermission -Identity ($Partner.Text + ($UserDomain) + ($Calendar)) -User ($Delegate.text + ($UserDomain)) -AccessRights PublishingEditor | Out-String;
        $OutputBox.text="Delegate Access Added" 
        
    }
    if($RA.Checked -eq $true){
        #Remove Access
        Remove-MailboxFolderPermission -Identity ($Partner.Text + ($UserDomain) + ($Calendar)) -User ($Delegate.text + ($UserDomain)) -Confirm:$false | Out-String;
        $OutputBox.text="Delegate Access Removed" 
        
    }
    


}

[void]$TheDelegator.ShowDialog()