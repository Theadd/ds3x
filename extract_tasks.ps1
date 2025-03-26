$vbaCode = Get-Content vba_code.txt;

$taskRegex = [regex]'AddTask\s*\"([^\"]+)\",\s*\"((?:[^\"]|\"\")+)\",\s*(?:Description\s*:=\s*StringLines\(\(([\s\S]*?)\s*\)\))?,?\s*(?:Usage\s*:=\s*((?:[^\"]|\"\")+))?';

$taskData = $taskRegex.Matches($vbaCode) | ForEach-Object {
    $taskName = $_.Groups[1].Value.Trim();
    $parameters = $_.Groups[2].Value -replace '""', ''''.Trim();
    $description = if($_.Groups[3].Success) { $_.Groups[3].Value -replace '_', '' -replace '@param\s*', ''  -replace '<font[^>]+>([^<]+)<\/font>', '$1'-replace '\s*_\s*', '' -replace 'StringLines\(', '' -replace '\)', '' } else { '' };
    $usage = if($_.Groups[4].Success) { $_.Groups[4].Value -replace '""', ''''.Trim() -replace '_', '' } else { '' };
    [pscustomobject]@{ TaskName = $taskName; Parameters = $parameters; Description = $description; Usage = $usage }
};

$taskData | ConvertTo-Json -Depth 10 | Out-File -FilePath task_data.json;
