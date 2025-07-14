# Script to list relevant patches using BigFix QnA tool (pure client relevance)
# Run as administrator. Adjust paths if needed.

$bigfixPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client"
$qnaExe = "$bigfixPath\qna.exe"
$tempQueryFile = "$env:TEMP\bigfix_query.txt"
$relevantPatches = @()

# Client relevance query: List names, IDs, and sites of relevant patch Fixlets
# Filters by MIME field "X-Fixlet-Type" == "Fixlet" (or absent, for older content)
$query = 'concatenation "; " of (name of it & " (ID: " & id of it as string & ", Site: " & name of site of it & ")") of relevant fixlets whose (not exists mime field "X-Fixlet-Type" of it or value of mime field "X-Fixlet-Type" of it = "Fixlet") of sites whose (name of it contains "Patch" or name of it contains "Update" or name of it contains "Security")'

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
        if ($patchesString -ne "E: Singular expression refers to nonexistent object." -and $patchesString -ne "<nothing>") {
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
        Write-Host "`nNo relevant patches found, or no matching sites subscribed. Check site subscriptions in BigFix or tweak the site filter in the query."
    }
} else {
    Write-Host "qna.exe not found at $qnaExe. Ensure it's in the BigFix client folder or copy from the console installation."
}
