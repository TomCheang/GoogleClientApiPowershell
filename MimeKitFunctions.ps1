
function Get-MimeKitMimeMessage {
  Param ($FilePath)
  $Stream = [System.IO.File]::OpenRead($FilePath)
  $MimeFormat = [MimeKit.MimeFormat]::Entity
  $Parser = New-Object MimeKit.MimeParser($Stream, $MimeFormat)
  $Parser.ParseMessage()
}

function Get-MD5HashFromString  {
  Param ([string]$String)
  $hasher = new-object System.Security.Cryptography.MD5Cng
  $toHash = [System.Text.Encoding]::UTF8.GetBytes($String)
  $hashByteArray = $hasher.ComputeHash($toHash)
  foreach($byte in $hashByteArray) {
    $res += $byte.ToString("x2")
   }
  return $res
}

#rfc822 date string
function Get-rfc822DateString {
  Param ([datetimeoffset]$Date)
  [MimeKit.Utils.DateUtils]::FormatDate($Date)
}

function Get-MimeFieldsToString {
  #FieldOrder: Date,From,To,CC,Subject,Body.Text
  Param ($MimeMsg)

  $strDate = Get-rfc822DateString -Date $MimeMsg.Date
  if ($MimeMsg.From) {
    $strFrom = $MimeMsg.From.ToString()
  }

  if ($MimeMsg.To) {
    $strTo = $MimeMsg.To.ToString()
  }

  if ($MimeMsg.CC) {
    $strCC = $MimeMsg.CC.ToString()
  }
  $strSubject = $MimeMsg.Subject
  $strBodyText = $MimeMsg.Body.TextBody

  ($strDate + $strFrom + $strTo + $strCC + $strSubject + $strBodyText)
}

function Get-CombinedHash {
    [CmdletBinding()]
    [Alias()]
    [OutputType([string])]
    Param (
    # MimeKit.MimeMessage obj
    [Parameter(Mandatory=$true,
      ValueFromPipelineByPropertyName=$true,
      Position=0)]
      $MimeMsg
    )
    Process {
      $strMimeFields = Get-MimeFieldsToString -MimeMsg $MimeMsg
      Get-MD5HashFromString -String $strMimeFields
    }
  }
