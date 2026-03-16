<#
.SYNOPSIS
Polls for Microsoft Visual Basic for Applications error dialogs using Win32 API,
reads their error text, clicks OK to dismiss them, and writes results to stdout.
Handles MULTIPLE sequential dialogs.

Uses Win32 FindWindowEx / SendMessage instead of UIAutomation for reliability
across different Windows/Office configurations where UIAutomation may be blocked
by UIPI or accessibility restrictions.
#>
param(
    [int]$TimeoutMs  = 60000,
    [int]$PollMs     = 300
)

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class VbaDialogHelper {
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern IntPtr FindWindowEx(
        IntPtr hwndParent, IntPtr hwndChildAfter,
        string lpClassName, string lpWindowName);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(
        IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(
        IntPtr hWnd, StringBuilder sb, int maxCount);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    public const uint BM_CLICK = 0x00F5;
}
"@

$deadline    = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)
$dismissed   = 0
$dialogTitle = 'Microsoft Visual Basic for Applications'

while ([DateTime]::UtcNow -lt $deadline) {
    try {
        # Enumerate ALL top-level windows matching the VBA dialog title.
        # FindWindowEx(NULL, prev, NULL, title) iterates desktop children by title.
        $prev          = [IntPtr]::Zero
        $foundThisPass = $false

        while ($true) {
            $hwnd = [VbaDialogHelper]::FindWindowEx(
                [IntPtr]::Zero, $prev, $null, $dialogTitle)
            if ($hwnd -eq [IntPtr]::Zero) { break }
            $prev = $hwnd

            if (-not [VbaDialogHelper]::IsWindowVisible($hwnd)) { continue }

            # Only act on windows that have an OK button child.
            # The VBA IDE window also matches the title but has no OK button.
            $okBtn = [VbaDialogHelper]::FindWindowEx(
                $hwnd, [IntPtr]::Zero, 'Button', 'OK')
            if ($okBtn -eq [IntPtr]::Zero) { continue }

            # Read error text from Static label controls inside the dialog
            $texts     = @()
            $childPrev = [IntPtr]::Zero
            while ($true) {
                $child = [VbaDialogHelper]::FindWindowEx(
                    $hwnd, $childPrev, 'Static', $null)
                if ($child -eq [IntPtr]::Zero) { break }
                $childPrev = $child

                $len = [VbaDialogHelper]::GetWindowTextLength($child)
                if ($len -gt 0) {
                    $sb = New-Object System.Text.StringBuilder ($len + 1)
                    [VbaDialogHelper]::GetWindowText($child, $sb, $sb.Capacity) | Out-Null
                    $t = $sb.ToString().Trim()
                    if ($t -ne '') { $texts += $t }
                }
            }
            $errorText = if ($texts.Count -gt 0) { $texts -join ' — ' } else { '(no text captured)' }

            # Click OK via SendMessage (no mouse coordinates needed)
            [VbaDialogHelper]::SendMessage(
                $okBtn, [VbaDialogHelper]::BM_CLICK,
                [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null

            $dismissed++
            Write-Output "DIALOG_DISMISSED:$errorText"
            $foundThisPass = $true

            # Brief pause to let Excel process the click, then re-check for more dialogs
            Start-Sleep -Milliseconds 500
            break
        }

        if ($foundThisPass) { continue }
    } catch {
        # Keep polling on any error
    }

    Start-Sleep -Milliseconds $PollMs
}

if ($dismissed -gt 0) {
    Write-Output "DIALOG_WATCH_DONE:$dismissed dialog(s) dismissed"
    exit 0
} else {
    Write-Output 'DIALOG_TIMEOUT'
    exit 1
}
