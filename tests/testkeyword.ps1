$sc = New-Object -ComObject MSScriptControl.ScriptControl
$sc.Language = "VBS"

add-type 'namespace Serpen {[System.Runtime.InteropServices.ComVisible(true)]public class RetObject {private string val = ""; public string Val {get {return val;} set {val = value;}}}}' -Language CSharp

$retObj = New-Object Serpen.RetObject
$sc.AddObject("retObject", $retObj, $true)

$content = cat D:\Austausch\vbscript-dll\vbscript-10-x86-Strings-2-sorted.txt
$toCheck = [System.Collections.Generic.List[string]]::new()

foreach ($line in $content | select -First 39999) {
    $sc.ExecuteStatement('retObject.Val = ""')
    try {
        $sc.ExecuteStatement("Option Explicit : retObject.Val = $line")
        "$line : works " + $sc.Eval("retObject.val")
        $toCheck.Add($line)
    } catch {
        if ($_.Exception.ErrorCode -eq -2146827256) {
            # ungültiges Zeichen
        } elseif ($_.Exception.ErrorCode -eq -2146827286) {
            "$line : wrong syntax $($_.Exception.ErrorCode)"
            $toCheck.Add($line)
        } elseif ($_.Exception.ErrorCode -eq -2146827788) {
            # not def
        } elseif ($_.Exception.ErrorCode -eq -2146827788) {
            # not def
        } else {
            "$line : $($_.Exception.Message) $($_.Exception.ErrorCode)"
            $toCheck.Add($line)
        }
    } finally {
        $sc.ExecuteStatement('retObject.Val = ""')
    }
}

#$toCheck | % 