# language: powershell

# Description: This script compares two files
# and returns the differences between them.

$file1 = Get-Content -Path ".\testfile1.txt"
$file2 = Get-Content -Path ".\testfile2.txt"

# The comparator can be either "<=" or "=>"
# "<=" means that the line is in file1 but not in file2
# "=>" means that the line is in file2 but not in file1
$comparator = "<="

# Compare the two files and return the differences
$diff = Compare-Object -ReferenceObject $file1 -DifferenceObject $file2 -PassThru | Where-Object SideIndicator -eq $comparator

# Print the differences
$diff
