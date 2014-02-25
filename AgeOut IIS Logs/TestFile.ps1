Param
(
    [string]$TestArg
)

$msg = "Argument Passed: " + $TestArg

Write-EventLog -LogName "Application" -Source "TestSource" -EntryType Information -EventID 666 -Message $msg