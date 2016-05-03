
  
#region Gmail Functions

#region var inits
$privkey = '.\clientsecret.p12'
$keypassword = 'notasecret'
$adminuser =

$AdminId = 


$dll = gci -Path $DllPath -Filter *.dll | sort FullName -Descending
foreach ($d in $dll) {Add-Type -Path $d.FullName}

#endregion vars

function Get-GmailAllUsers {
  Param ($DirectoryService)
  $ListRequest = $DirectoryService.Users.List()
  $ListRequest.Customer = 'my_customer'
  $ListRequest.MaxResults = '100'
  $ListRequest.Projection = 'Full'
  $ListRequest.OrderBy = '0' #primary email
  $ListRequest.ShowDeleted = $false # $true only returns deleted


  Do {
      $rtn = $ListRequest.ExecuteAsync()  
      [array]$list += $rtn.Result.UsersValue
      $ListRequest.PageToken = $rtn.Result.NextPageToken
    }
  While ($rtn.Result.NextPageToken)

  $list | select @{n='FirstName';e={$_.Name.GivenName}}, @{n='LastName';e={$_.Name.FamilyName}},
    PrimaryEmail, Aliases,  CreationTime, LastLoginTime, IsMailboxSetup, Suspended
   
}

function Get-GmailAllGroups {
  Param ($DirectoryService)
   $ListGroups = $DirectoryService.Groups.List()
  $ListGroups.MaxResults = '50'
  $ListGroups.Customer = 'my_customer'


  Do {
      $groups = $ListGroups.ExecuteAsync()  
      [array]$glist += $groups.Result.GroupsValue |
        select Name, Description, Email, Aliases, DirectMembersCount
      $ListGroups.PageToken = $groups.Result.NextPageToken
  }
  While ($groups.Result.NextPageToken)

  return $glist
}

function Get-GmailGroup {
  Param($DirectoryService, $GroupEmail)

     Do {
      $members = $DirectoryService.Members.List($GroupEmail).ExecuteAsync()  
      [array]$mlist += $members.Result.MembersValue |
        select Email, Role, Type
      $ListMembers.PageToken = $members.Result.NextPageToken
  }
  While ($members.Result.NextPageToken)

  $Group = $DirectoryService.Groups.Get($GroupEmail).ExecuteAsync()
  $ManagedBy = $mlist | where {$_.Role -eq 'OWNER'}
  $GroupInfo = $Group.Result | select *, @{n='ManagedBy';e={$ManagedBy}},
    @{n='Members';e={$mlist}}

  Return $GroupInfo
}

function Get-GmailUserLabelsAll {
  Param ($GmailService)

  $GmailService.Users.Labels.List($Impersonate).ExecuteAsync().Result.Labels
}

function Get-GmailLabel {
  Param ($GmailService, $Impersonate, $LabelId)
  $lbl = $GmailService.Users.Labels.Get($Impersonate, $LabelId).ExecuteAsync().Result

  Return $lbl
}

function Get-GmailUserMessageIds {
  Param ($GmailService, $Impersonate, $SearchQuery, $MaxResults = '500')
  $ListMsgs = $GmailService.Users.Messages.List($Impersonate)
  if ($SearchQuery) {
    $ListMsgs.Q = $SearchQuery
  }
  $ListMsgs.MaxResults = $MaxResults
  $ListMsgs.IncludeSpamTrash = $false #default
 #$ListMsgs.Fields = 'Messages(Id)'

  $MessageIds = New-Object System.Collections.ArrayList

    Do {
        $msgs = $ListMsgs.ExecuteAsync()
      $MessageIds.AddRange($msgs.Result.Messages.Id)
      $ListMsgs.PageToken = $msgs.Result.NextPageToken
  }
  While ($msgs.Result.NextPageToken)

  return $MessageIds
}

function Get-GmailMessageRaw {
  Param ($GmailService, $Impersonate, $MsgId)

  $GetMsg = $GmailService.Users.Messages.Get($Impersonate, $MsgId).ExecuteAsync()
  $GetMsg.Format = 'RAW'
 # $GetMsg.Fields = 'raw,LabelIds'

  $GetMsg.ExecuteAsync().Result
}

function New-GmailDirectoryService {
  Param ($privkey, $AdminId, $Impersonate)

$scopes = @(
  [Google.Apis.Admin.Directory.directory_v1.DirectoryService+Scope]::AdminDirectoryUser,
  [Google.Apis.Admin.Directory.directory_v1.DirectoryService+Scope]::AdminDirectoryGroup,
  [Google.Apis.Admin.Directory.directory_v1.DirectoryService+Scope]::AdminDirectoryGroupMember
    )

    #Build auth scopes object
    $auth_scopes = New-Object "System.Collections.Generic.List``1[system.string]"
    $scopes | foreach { $auth_scopes.Add($_) }


    $X509KeyStorageFlags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
    $certificate = New-Object System.Security.Cryptography.x509Certificates.X509Certificate2(
      $privkey, $keypassword, $X509KeyStorageFlags)

    #Construct service account credential initializer
    $Initializer = new-object Google.Apis.Auth.OAuth2.ServiceAccountCredential+Initializer($AdminId)
    $Initializer.FromCertificate($certificate) | Out-Null
    $Initializer.Scopes = $auth_scopes
    $Initializer.User = $Impersonate
    $credential = New-Object Google.Apis.Auth.OAuth2.ServiceAccountCredential($Initializer)

    $BaseClientSvc = New-Object Google.Apis.Services.BaseClientService+Initializer
    $BaseClientSvc.HttpClientInitializer = $credential
    $BaseClientSvc.ApplicationName = 'GmailLabels'
    
    New-Object Google.Apis.Admin.Directory.directory_v1.DirectoryService($BaseClientSvc)
}


function New-GmailService {
  Param ($privkey, $AdminId, $Impersonate)

  $scopes = @(
      [Google.Apis.Gmail.v1.GmailService+Scope]::GmailLabels,
      [Google.Apis.Gmail.v1.GmailService+Scope]::MailGoogleCom
      [Google.Apis.Calendar.v3.CalendarService+Scope]::Calendar
        )

    #Build auth scopes object
    $auth_scopes = New-Object "System.Collections.Generic.List``1[system.string]"
    $scopes | foreach { $auth_scopes.Add($_) }


    $X509KeyStorageFlags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable
    $certificate = New-Object System.Security.Cryptography.x509Certificates.X509Certificate2(
      $privkey, $keypassword, $X509KeyStorageFlags)

    #Construct service account credential initializer
    $Initializer = new-object Google.Apis.Auth.OAuth2.ServiceAccountCredential+Initializer($AdminId)
    $Initializer.FromCertificate($certificate) | Out-Null
    $Initializer.Scopes = $auth_scopes
    $Initializer.User = $Impersonate
    $credential = New-Object Google.Apis.Auth.OAuth2.ServiceAccountCredential($Initializer)

    $BaseClientSvc = New-Object Google.Apis.Services.BaseClientService+Initializer
    $BaseClientSvc.HttpClientInitializer = $credential
    $BaseClientSvc.ApplicationName = 'GmailLabels'
    
    New-Object Google.Apis.Gmail.v1.GmailService($BaseClientSvc)
}

function Build-Filename {
  Param ($MsgId,[array]$Labels)
  $lblString =  ''
  $Labels = $Labels | where {$_.id -notlike 'Category*'} | 
    where {$_.id -notin ('Unread', 'Draft', 'Spam', 'Important',
      'Chat', 'Trash')}

  foreach ($l in $Labels) {
    $lblString += ('-' + $l)
  }
  $MsgId + $lblString + '.eml'
}

function Export-MimeKitMessageToEml {
  Param ($MimeMessage, $Path, $FileName)

  #$FilePath = [System.IO.Path]::Combine($Path, $MimeMessage.Date.Year, $FileName)
  $FilePath = [System.IO.Path]::Combine($Path, $FileName)
  $MimeMessage.WriteTo($FilePath)

}

function Convert-GmailRawToMimeMessage {
  Param ($string)
  
  #Converting to base64url string
  $string = $string.Replace('-', '+').Replace('_', '/')

  switch ($string.Length % 4) {
      '0' {break} #no need to pad
      '2' {$string += '=='; break}
      '3' {$string += '='; break}
      Default {Write-Host 'Baseurl string not valid.'; break}
  }

  #send to memorystream
  $buffer = [convert]::fromBase64String($string)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($buffer)
  $stream = [System.IO.MemoryStream]::new($buffer)

  #Parse as MimeMessage
  $MimeFormat = [MimeKit.MimeFormat]::Entity
  $Parser = New-Object MimeKit.MimeParser($stream,$MimeFormat)
  $Parser.ParseMessage()
  
  #Returns MimeKit.MimeMessage obj
  Return $Parser
}

  function Get-MD5HashFromString  {
    Param ([string]$String)
    $hasher = new-object System.Security.Cryptography.MD5Cng
    $toHash = [System.Text.Encoding]::UTF8.GetBytes($String)
    $hashByteArray = $hasher.ComputeHash($toHash)
    foreach($byte in $hashByteArray)
     {
      $res += $byte.ToString("x2")
     }
    return $res
  }

function Get-rfc822DateString {
  Param ([datetimeoffset]$Date)
  [MimeKit.Utils.DateUtils]::FormatDate($Date)
}

function Get-MimeFields {
  #FieldOrder: Date,From,To,CC,Subject,Body.Text
  Param ($MimeMsg)

  $strDate = Get-rfc822DateString -Date $MimeMsg.Date
  $strFrom = $MimeMsg.From.ToString()
  $strTo = $MimeMsg.To.ToString()
  $strCC = $MimeMsg.Cc.ToString()
  $strSubject = $MimeMsg.Subject.ToString()
  $strBodyText = $MimeMsg.Body.TextBody.ToString()

  ($strDate + $strFrom + $strTo + $strCC + $strSubject + $strBodyText)
}



function Out-CombinedHash {
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param
    (
        # MimeKit.MimeMessage obj
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $MimeMsg)
    Process {

      ($strDate + $strFrom + $strTo + $strCC + $strSubject + $strBodyText)

      $strMimeFields = Get-MimeFields -MimeMsg $MimeMsg
      Get-MD5HashFromString -String $strMimeFields
    }
}
