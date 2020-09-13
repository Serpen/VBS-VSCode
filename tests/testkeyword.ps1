param (
    [string]$FileWithStrings = 'C:\Users\serpe\Documents\VSCode-VBS\VBS-VSCode\tests\vbscript-dll-strings.txt' # 'c:\temp\sccrun.txt'
)
$sc = New-Object -ComObject MSScriptControl.ScriptControl
$sc.Language = "VBS"

add-type 'namespace Serpen {[System.Runtime.InteropServices.ComVisible(true)]public class RetObject {private string val = ""; public string Val {get {return val;} set {val = value;}}}}' -Language CSharp

$retObj = New-Object Serpen.RetObject
$sc.AddObject("retObject", $retObj, $true)

$content = Get-Content $FileWithStrings
$toCheck = [System.Collections.Generic.List[string]]::new()

foreach ($line in $content | Select-Object -First 39999) {
    $sc.ExecuteStatement('retObject.Val = ""')
    try {
        $sc.ExecuteStatement("Option Explicit : retObject.Val = $line")
        "$line = '" + $sc.Eval("retObject.val") + "' works!"
        $toCheck.Add($line)
        
    } catch {
        $errStr = "$line : $($_.Exception.Message) [$($_.Exception.ErrorCode), $($_.Exception.HResult)]"
        if ($_.Exception.ErrorCode -eq -2146827256) {
            Write-Verbose $errStr
            # ungültiges Zeichen
        } elseif ($_.Exception.ErrorCode -eq -2146827286) {
            $errStr
            #wrong syntax
            $toCheck.Add($line)
        } elseif ($_.Exception.ErrorCode -eq -2146827788) {
            # not def
            Write-Verbose $errStr
        } elseif ($_.Exception.HResult -eq -2146233067) {
            # obj hasn't property
            Write-Verbose $errStr
        } elseif ($_.Exception.ErrorCode -eq -2146827788) {
            # not def
            Write-Verbose $errStr
        } else {
            $errStr
            $toCheck.Add($line)
        }
    }
}

#$toCheck | % 