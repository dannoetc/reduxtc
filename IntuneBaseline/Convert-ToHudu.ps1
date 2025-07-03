# Path to your original and output files
$input  = 'README.md'
$output = 'README_Hudu.md'


# Read all lines into a string[] array
$lines = Get-Content $input

# Find the first line matching a top-level heading
$firstHeading = $lines | Where-Object { $_ -match '^#\s+' } | Select-Object -First 1
if ($firstHeading) {
    # Get its zero-based index
    $insertIndex = [Array]::IndexOf($lines, $firstHeading)
} else {
    Write-Warning "No top-level heading found; appending [TOC] at top."
    $insertIndex = 0
}

# Build new content
$new  = @()
$new += $lines[0..($insertIndex - 1)]      # everything before the first H1
$new += '[TOC]'                             # Hudu Table of Contents macro
$new += ''                                  # blank line
$new += $lines[$insertIndex..($lines.Length - 1)]  # the rest of the file

# Normalize headings (ensure a single space after each ‘#’)
$new = $new -replace '^######\s+', '###### '
$new = $new -replace '^#####\s+',  '##### '
$new = $new -replace '^####\s+',   '#### '
$new = $new -replace '^###\s+',    '### '
$new = $new -replace '^##\s+',     '## '
$new = $new -replace '^#\s+',      '# '        

# Write out the result
$new | Set-Content $output -Encoding UTF8

Write-Host "Hudu-ready Markdown generated at $output"