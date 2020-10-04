function Encode-VBS {
  param (
    [Parameter(Mandatory = $true)][string]$File,
    [string]$OutFile = $File.Replace(".vbs", ".vbe"),
    [Switch]$Force
  )

  if (!(Test-Path $File -PathType Leaf)) {
    throw (New-Object System.IO.FileNotFoundException $File)
  }

  if ((Test-Path $OutFile -PathType Leaf) -and $Force) {
    throw (New-Object System.IO.IOException "OutFile '$OutFile' already exists (use -force)")
  }

  $encoder = New-Object -ComObject Scripting.Encoder

  $Content = (Get-Content $File -Raw)

  $encoder.EncodeScriptFile(".vbs", $Content, 0, "VBScript") | Out-File $OutFile -Encoding ascii
}

$TAG_BEGIN1 = "#@~^"
$TAG_BEGIN2 = "=="
$TAG_BEGIN2_OFFSET = 10
$TAG_BEGIN_LEN = 12
$TAG_END = "==^#~@"
$TAG_END_LEN = 6

function Decode-VBS {
  param (
    [Parameter(Mandatory = $true)][string]$File,
    [string]$OutFile = $File.Replace(".vbe", ".vbs"),
    [Switch]$Force
  )

  if (!(Test-Path $File -PathType Leaf)) {
    throw (New-Object System.IO.FileNotFoundException $File)
  }

  if ((Test-Path $OutFile -PathType Leaf) -and $Force) {
    throw (New-Object System.IO.IOException "OutFile '$OutFile' already exists (use -force)")
  }

  $content = Get-Content $file -Raw

  $iTagBeginPos = $content.IndexOf($TAG_BEGIN1)

  if ($iTagBeginPos -eq 0) {
    throw "$file does not appear to be encoded.  Missing Beginning Tag.  Skipping file."
  }
  elseif ($iTagBeginPos -eq 1) {
    if (!$content.IndexOf($TAG_BEGIN2, $iTagBeginPos) -eq $TAG_BEGIN2_OFFSET) {
      throw "$file does not appear to be encoded.  Incomplete Beginning Tag.  Skipping file."
    }
    else {
      $iTagEndPos = $content.IndexOf($TAG_END, $iTagBeginPos)
      if ($iTagEndPos -gt 0) {
        if ($content.Substring($iTagEndPos + $TAG_END_LEN) -in '', 0) {
          Decode($content.Substring($iTagBeginPos + $TAG_BEGIN_LEN, $iTagEndPos - $iTagBeginPos - $TAG_BEGIN_LEN - $TAG_END_LEN)) | Out-File $OutFile -NoNewline
        }
        else {
          throw "$file does not appear to be encoded.  Found xxx characters AFTER Ending Tag.  Skipping file."
        }
      }
      else {
          throw "$file does not appear to be encoded.  Missing ending Tag.  Skipping file."
      }
    }
  }
}

function Decode {
  param (
    $Chaine
  )
  $tDecode = new-object String[] 127
  $Combinaison="1231232332321323132311233213233211323231311231321323112331123132"

  $encoder = New-Object -ComObject "Scripting.Encoder"
  
  for ($i = 9; $i -lt 127; $i++) {
    $tDecode[$i] = "JLA"
  }
  for ($i = 9; $i -lt 127; $i++) {
    $ChaineTemp = $encoder.EncodeScriptFile(".vbs",(new-object string ($i, 3)),0,"").SubString(13,3)
    for ($j = 1; $j -lt 3; $j++) {
      $c = $ChaineTemp[$j]
      $tDecode[$c] = $tDecode[$c].Substring(0,$j-1) + $i + $tDecode[$c].Substring($j-1)
    }
  }
}