dir *10*.dll | % {
	Write-Progress -Activity "String Extraction" -CurrentOperation $_.Name
	E:\SysinternalsSuite\strings.exe $_.FullName | out-file $_.FullName.Replace('.dll', '-Strings.txt')
	$stringsfile = cat $_.FullName.Replace('.dll', '-Strings.txt')
	$stringsFile | Select-String -Pattern '^[a-z]+$' | out-file $_.FullName.Replace('.dll', '-Strings-2.txt')
	$stringsFile | Select-String -Pattern '^[a-z]+$' | sort -unique | out-file $_.FullName.Replace('.dll', '-Strings-2-sorted.txt')
	$stringsFile | Select-String -Pattern '^vb[a-z0-9_]+[^W]$' | sort -unique | out-file $_.FullName.Replace('.dll', '-Strings-vb.txt')
	$REGEX = '^(([a-z]+[A-Z]+[a-zA-Z]+[a-z])|([A-Z]+[a-z]+[a-zA-Z]+))[^A-Z]$'
	$stringsFile | Select-String -Pattern $REGEX -CaseSensitive | out-file $_.FullName.Replace('.dll', '-Strings-Case.txt')
	$stringsFile | Select-String -Pattern $REGEX -CaseSensitive | sort -unique | out-file $_.FullName.Replace('.dll', '-Strings-Case-sorted.txt')
}