# Script to list relevant patches using BigFix QnA tool
# Run as administrator. Adjust paths if needed.

$bigfixPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client"
$qnaExe = "$bigfixPath\qna.exe"
$tempQueryFile = "$env:TEMP\bigfix_query.txt"
$relevantPatches = @()

# Relevance query: List names, IDs, and sites of relevant patch Fixlets
# Filters out non-Fixlet types like Tasks/Analyses/Baselines
$query = 'concatenation "; " of (name of it & " (ID: " & id of it as string & ", Site: " & name of site of it & ")") of relevant fixlets whose (fixlet flag of it and (not exists headers "X-Fixlet-Type" of it or value of header "X-Fixlet-Type" of it != "Task" and value of header "X-Fixlet-Type" of it != "Analysis" and value of header "X-Fixlet-Type" of it != "Baseline")) of sites whose (name of it contains "Patch" or name of it contains "Update" or name of it contains "Security")'

# Write query to temp file (QnA expects "Q: " prefix)
"Q: $query" | Out-File -FilePath $tempQueryFile -Encoding ASCII

# Run qna.exe and capture output
if (Test-Path $qnaExe) {
    $output = & $qnaExe $tempQueryFile
    Write-Host "QnA Output:"
    $output | ForEach-Object { Write-Host $_ }

    # Parse the result (look for lines starting with "A: " - the answer)
    $resultLine = $output | Where-Object { $_ -match "^A: " } | Select-Object -First 1
    if ($resultLine) {
        $patchesString = $resultLine -replace "^A: ", ""
        if ($patchesString -ne "<nothing>") {
            $relevantPatches = $patchesString -split "; " | Sort-Object
        }
    }

    # Clean up temp file
    Remove-Item $tempQueryFile -Force -ErrorAction SilentlyContinue

    # Display results
    if ($relevantPatches.Count -gt 0) {
        Write-Host "`nRelevant Patches:"
        $relevantPatches | ForEach-Object { Write-Host "- $_" }
    } else {
        Write-Host "`nNo relevant patches found, or query returned nothing."
    }
} else {
    Write-Host "qna.exe not found at $qnaExe. Ensure it's in the BigFix client folder or copy from console installation."
}
