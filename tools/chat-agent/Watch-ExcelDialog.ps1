<#
.SYNOPSIS
Polls for a Microsoft Visual Basic for Applications error dialog,
reads its error text, clicks OK to dismiss it, and writes the result to stdout.
Run this in parallel alongside any Excel COM automation.
#>
param(
    [int]$TimeoutMs  = 60000,
    [int]$PollMs     = 300
)

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$root     = [System.Windows.Automation.AutomationElement]::RootElement
$nameProp = [System.Windows.Automation.AutomationElement]::NameProperty
$deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)

while ([DateTime]::UtcNow -lt $deadline) {
    try {
        $cond   = [System.Windows.Automation.PropertyCondition]::new($nameProp, 'Microsoft Visual Basic for Applications')
        $dialog = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)

        if ($dialog) {
            # Walk immediate children to collect text (skips OK / Help button labels)
            $parts  = @()
            $walker = [System.Windows.Automation.TreeWalker]::ContentViewWalker
            $child  = $walker.GetFirstChild($dialog)
            while ($child) {
                $n = $child.Current.Name
                if ($n -and $n.Trim() -ne '' -and $n -ne 'OK' -and $n -ne 'Help') {
                    $parts += $n.Trim()
                }
                $child = $walker.GetNextSibling($child)
            }
            $errorText = if ($parts.Count -gt 0) { $parts -join ' — ' } else { '(no text captured)' }

            # Click OK
            $okCond = [System.Windows.Automation.PropertyCondition]::new($nameProp, 'OK')
            $okBtn  = $dialog.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $okCond)
            if ($okBtn) {
                $invoke = $okBtn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
                $invoke.Invoke()
            }

            Write-Output "DIALOG_DISMISSED:$errorText"
            exit 0
        }
    } catch {
        # UIAutomation can throw on race conditions — keep polling
    }

    Start-Sleep -Milliseconds $PollMs
}

Write-Output 'DIALOG_TIMEOUT'
exit 1
