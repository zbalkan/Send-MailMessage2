#Requires -Version 3.0
<#
.Synopsis
Sends an email message using existing Exchange Server with user's credential.
.DESCRIPTION
The Send-MailMessage2 cmdlet sends an email message from within Windows PowerShell. It utilizes Autodiscovery feature of Exchange Server.
.EXAMPLE
Send an email from one user to another:

PS C:\>Send-MailMessage2 -To "User01 <user01@example.com>"  -Subject "Test mail"

This command sends an email message fom the user who runs the script to User01.

The mail message has a subject, which is required, but it does not have a body, which is optional. Also, because the SmtpServer parameter is not specified, Send-MailMessage2 uses the value of the Autodiscover DNS record for the Exchange Server.

.INPUTS
System.String
You can pipe the path and file names of attachments to Send-MailMessage2 .
.OUTPUTS
None
This cmdlet does not generate any output.
.NOTES
Note that you cannot define From, SmtpServer, Port and Credential for this Cmdlet.
.LINK
https://gist.github.com/zbalkan/47b8916cf2823e8e383500c7ae513641
#>
function Send-MailMessage2
{
    [CmdletBinding(SupportsShouldProcess=$true, 
    PositionalBinding=$false,
    HelpUri = 'https://gist.github.com/zbalkan/47b8916cf2823e8e383500c7ae513641',
    ConfirmImpact='Medium')]
    Param
    (
    # Specifies the addresses to which the mail is sent. Enter names (optional) and the email address, such as Name <someone@example.com>. This parameter is required.
    [Parameter(Mandatory=$true, 
    ValueFromPipeline=$false,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false, 
    Position=0)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $To,
    
    # Specifies the subject of the email message. This parameter is required.
    [Parameter(Mandatory=$true, 
    ValueFromPipeline=$false,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false, 
    Position=1)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [string]
    $Subject,
    
    # Specifies the body of the email message.
    [Parameter(Mandatory=$false, 
    ValueFromPipeline=$false,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false, 
    Position=2)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [string]
    $Body,
    
    # Specifies the path and file names of files to be attached to the email
    # message. You can use this parameter or pipe the paths and file names to
    # Send-MailMessage2.
    [Parameter(Mandatory=$false, 
    ValueFromPipeline=$true,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $Attachments,
    
    # Specifies the email addresses that receive a copy of the mail but are not
    # listed as recipients of the message. Enter names (optional) and the email
    # address, such as Name <someone@example.com>.
    [Parameter(Mandatory=$false, 
    ValueFromPipeline=$false,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $Bcc,
    
    # Indicates that the value of the Body parameter contains HTML.
    [Parameter(Mandatory=$false, 
    ValueFromPipeline=$false,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false)]
    [switch]
    $BodyAsHtml,
    
    # Specifies the encoding used for the body and subject. The acceptable
    # values for this parameter are: ASCII, BigEndianUnicode, Default, OEM,
    # Unicode, UTF7, UTF8, UTF32
    [Parameter(Mandatory=$false, 
    ValueFromPipeline=$false,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("ASCII","BigEndianUnicode","Default","OEM","Unicode","UTF7","UTF8","UTF32")]
    [string]
    $Encoding,
    
    # Specifies the email addresses to which a carbon copy (CC) of the email
    # message is sent. Enter names (optional) and the email address, such as
    # Name <someone@example.com>.
    [Parameter(Mandatory=$false, 
    ValueFromPipeline=$false,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $Cc,
    
    # Specifies the delivery notification options for the email message. You can
    # specify multiple values. None is the default value. The alias for this
    # parameter is dno.
    [Parameter(Mandatory=$false, 
    ValueFromPipeline=$false,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("None","OnSuccess","OnFailure","Delay","Never")]
    [Alias("dno")]
    [string[]]
    $DeliveryNotificationOption,
    
    # Specifies the priority of the email message. The acceptable values for this parameter are: Low, Normal, High.
    [Parameter(Mandatory=$false, 
    ValueFromPipeline=$false,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false)]
    [ValidateNotNull()]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("Low","Normal","High")]
    [string]
    $Priority,
    
    # Indicates that the cmdlet uses the Secure Sockets Layer (SSL) protocol to establish a connection to the remote computer to send mail. By default, SSL is not used.
    [Parameter(Mandatory=$false, 
    ValueFromPipeline=$false,
    ValueFromPipelineByPropertyName=$false, 
    ValueFromRemainingArguments=$false)]
    [switch]
    $UseSsl
    )
    
    Begin
    {
    }
    Process
    {
        if ($pscmdlet.ShouldProcess("specified addresses", "Send an email using default credentials and Exchange server of the network"))
        {
            # Get Exchange Server host name
            $ExchServer = [Net.DNS]::GetHostByAddress([Net.DNS]::GetHostEntry("Autodiscover").AddressList[0]).Hostname
            Write-Verbose "Exchange Server Address: $ExchServer"

            # Get user email address from LDAP attributes
            $UserEmail = ((New-Object System.DirectoryServices.DirectorySearcher -ArgumentList "(&(objectCategory=User)(samAccountName=$env:USERNAME))").FindOne().GetDirectoryEntry()).mail.Value     
            Write-Verbose "User Email Address: $UserEmail"   
            
            # Splat parameters
            $parameters = @{
                SmtpServer = $ExchServer
                From = $UserEmail
                To = $To
                Subject = $Subject
                Body = $Body
                Attachments = $Attachments
                Cc = $Cc
                Bcc = $Bcc
                BodyAsHtml = $BodyAsHtml
                DeliveryNotificationOption = $DeliveryNotificationOption
                Priority = $Priority
                Encoding = $Encoding
                UseSsl = $UseSsl
            }
            
            # Remove undefined parameters. 
            ($parameters.GetEnumerator() |
            Where-Object { -not $_.Value }) |
            ForEach-Object { $parameters.Remove($_.Name) }
            
            Write-Verbose "Email parameters: $($parameters | Out-String)"

            # Send mail message
            try {
                 Send-MailMessage @parameters   
            }
            catch {
                Write-Error -Exception $_.Exception
            }
            finally {
                Write-Verbose "Email sent successfully."
            }
            
        }
    }
    End
    {
    }
}
