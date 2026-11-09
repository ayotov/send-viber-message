<#
.SYNOPSIS
    Finds (or starts) the Viber desktop app and sends a message to a specific contact.

.DESCRIPTION
    Viber desktop exposes no public automation API, so this script drives it via
    UI Automation (UIA) plus mouse clicks:

      1. Locate (or launch) Viber.exe.
      2. Show the Viber window. If it is parked in the system tray, click its tray
         icon (the reliable way to un-hide a tray Qt app on Windows 11).
      3. Focus the Search box and set the contact name via UIA ValuePattern
         (this sets exact text regardless of the active keyboard layout).
      4. Click the top search result to open the chat.
      5. Set the message text in the message box via UIA ValuePattern.
      6. Invoke the Send button.

    Note: this sets text through UI Automation rather than simulated keystrokes,
    so it works even when the active keyboard layout is non-Latin (e.g. Cyrillic).

.PARAMETER Contact
    Name (or phone number) of the contact, exactly as it appears in Viber search.
    Use a value unique enough to match the right person.

.PARAMETER Message
    The message text to send (may contain any Unicode characters).

.PARAMETER NoSend
    Run the full flow but do NOT press Send, so nothing is actually sent. Use this
    to verify visually before sending for real.

.PARAMETER DryRun
    Only find and show the Viber window (does not search or type anything).

.PARAMETER StartIfNotRunning
    Launch Viber.exe if it is not already running.

.PARAMETER DelayAfterActionMs
    Milliseconds to wait after each step (search, open chat) for the UI to react.

.EXAMPLE
    .\Send-ViberMessage.ps1 -Contact "John Smith" -Message "Hello John!"
    .\Send-ViberMessage.ps1 -Contact "My Notes" -Message "Note" -NoSend
    .\Send-ViberMessage.ps1 -Contact "John Smith" -Message "Hi" -DryRun
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Contact,
    [Parameter(Mandatory = $true)][string]$Message,
    [switch]$NoSend,
    [switch]$DryRun,
    [switch]$StartIfNotRunning,
    [int]$DelayAfterActionMs = 1500,
    [int]$StartupWaitSeconds = 30
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms

if (-not ('ViberApi' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class ViberApi {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)]
    public struct NOTIFYICONIDENTIFIER { public uint cbSize; public IntPtr hWnd; public uint uID; public Guid guidItem; }

    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
    [DllImport("Shell32.dll")] public static extern int Shell_NotifyIconGetRect(ref NOTIFYICONIDENTIFIER identifier, out RECT iconLocation);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int X, int Y);
    [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT pt);
    [DllImport("user32.dll")] public static extern void mouse_event(uint f, uint dx, uint dy, uint d, UIntPtr e);
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();

    public const uint MOUSEEVENTF_LEFTDOWN = 0x02;
    public const uint MOUSEEVENTF_LEFTUP = 0x04;
}
"@
}

# Make sure mouse coordinates (SetCursorPos) line up with UIA bounding rectangles.
[void][ViberApi]::SetProcessDPIAware()

# ---------------------------------------------------------------------------
# UIA helpers
# ---------------------------------------------------------------------------
function Get-ViberWindow([int]$ProcessId) {
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $ProcessId)
    return $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
}

function Get-ViberWindows([int]$ProcessId) {
    $windows = New-Object System.Collections.Generic.List[object]
    [ViberApi]::EnumWindows({
        param($hWnd, $lParam)
        $procId = 0
        [void][ViberApi]::GetWindowThreadProcessId($hWnd, [ref]$procId)
        if ($procId -eq $ProcessId) {
            $cls = New-Object System.Text.StringBuilder 256
            $ttl = New-Object System.Text.StringBuilder 256
            [void][ViberApi]::GetClassName($hWnd, $cls, 256)
            [void][ViberApi]::GetWindowText($hWnd, $ttl, 256)
            $rect = New-Object ViberApi+RECT
            [void][ViberApi]::GetWindowRect($hWnd, [ref]$rect)
            $w = $rect.Right - $rect.Left
            $h = $rect.Bottom - $rect.Top
            $windows.Add([pscustomobject]@{
                Handle  = $hWnd
                Class   = $cls.ToString()
                Title   = $ttl.ToString()
                Visible = [ViberApi]::IsWindowVisible($hWnd)
                Width   = $w
                Height  = $h
                Area    = $w * $h
            })
        }
        return $true
    }, [IntPtr]::Zero) | Out-Null
    return $windows
}

function Get-ViberMainWindow([int]$ProcessId) {
    $windows = Get-ViberWindows -ProcessId $ProcessId
    $utility = 'PowerMessageWindow|TrayIconMessageWindow|ObserverWindow|_q_titlebar|MSCTFIME|^IME$'
    $candidates = $windows | Where-Object { $_.Class -notmatch $utility }

    $main = $candidates | Where-Object { $_.Class -match 'QWindowIcon' -and $_.Title -match 'Viber' } | Select-Object -First 1
    if (-not $main) { $main = $candidates | Where-Object { $_.Title -match 'Viber' } | Select-Object -First 1 }
    if (-not $main) { $main = $candidates | Where-Object { $_.Class -match 'QWindowIcon|Chrome_WidgetWin' } | Select-Object -First 1 }
    if (-not $main) { $main = $candidates | Sort-Object Area -Descending | Select-Object -First 1 }
    return $main
}

function Find-EditByClass($win, [string]$pattern) {
    $editCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
    $edits = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $editCond)
    foreach ($e in $edits) {
        if ($e.Current.ClassName -match $pattern) { return $e }
    }
    return $null
}

function Find-ButtonByClass($win, [string]$pattern) {
    $btnCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)
    $btns = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $btnCond)
    foreach ($b in $btns) {
        if ($b.Current.ClassName -match $pattern) { return $b }
    }
    return $null
}

function Set-ValuePatternValue($el, [string]$value) {
    $vp = $el.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
    $vp.SetValue($value)
}

function Click-Point([int]$X, [int]$Y) {
    $old = New-Object ViberApi+POINT
    [void][ViberApi]::GetCursorPos([ref]$old)
    [void][ViberApi]::SetCursorPos($X, $Y)
    Start-Sleep -Milliseconds 150
    [ViberApi]::mouse_event([ViberApi]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 60
    [ViberApi]::mouse_event([ViberApi]::MOUSEEVENTF_LEFTUP, 0, 0, 0, [UIntPtr]::Zero)
    Start-Sleep -Milliseconds 150
    [void][ViberApi]::SetCursorPos($old.X, $old.Y)
}

function Click-ViberTrayIcon([int]$ProcessId) {
    # Locate Viber's tray icon message window via the (working) window enumerator.
    $windows = Get-ViberWindows -ProcessId $ProcessId
    $trayWin = $windows | Where-Object { $_.Class -match 'TrayIconMessageWindow' } | Select-Object -First 1
    if (-not $trayWin) { return $false }

    # Qt registers the tray icon with uID 0; try a few values to be safe.
    foreach ($uid in 0..5) {
        $id = New-Object ViberApi+NOTIFYICONIDENTIFIER
        $id.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf([type][ViberApi+NOTIFYICONIDENTIFIER])
        $id.hWnd = $trayWin.Handle
        $id.uID = [uint32]$uid
        $id.guidItem = [Guid]::Empty
        $rect = New-Object ViberApi+RECT
        $hr = [ViberApi]::Shell_NotifyIconGetRect([ref]$id, [ref]$rect)
        if ($hr -eq 0 -and ($rect.Right -gt $rect.Left) -and ($rect.Bottom -gt $rect.Top)) {
            $cx = [int](($rect.Left + $rect.Right) / 2)
            $cy = [int](($rect.Top + $rect.Bottom) / 2)
            Write-Host ("Clicking Viber tray icon at ({0},{1})" -f $cx, $cy) -ForegroundColor Cyan
            Click-Point -X $cx -Y $cy
            return $true
        }
    }
    return $false
}

function Show-ViberWindow([int]$ProcessId) {
    # 1) Already visible via UIA -> bring it to the foreground.
    $win = Get-ViberWindow -ProcessId $ProcessId
    if ($win) {
        try { $win.SetFocus() } catch { }
        return $win
    }

    # 2) Hidden in the tray -> click the tray icon.
    if (Click-ViberTrayIcon -ProcessId $ProcessId) {
        Start-Sleep -Seconds 2
        $win = Get-ViberWindow -ProcessId $ProcessId
        if ($win) { try { $win.SetFocus() } catch { }; return $win }
    }

    # 3) Fallback: restore the hidden/minimized main window via Win32.
    $mainWin = Get-ViberMainWindow -ProcessId $ProcessId
    if ($mainWin) {
        [void][ViberApi]::ShowWindow($mainWin.Handle, 9)   # SW_RESTORE
        [void][ViberApi]::ShowWindow($mainWin.Handle, 5)   # SW_SHOW
        [void][ViberApi]::SetForegroundWindow($mainWin.Handle)
        Start-Sleep -Seconds 1
        $win = Get-ViberWindow -ProcessId $ProcessId
    }

    return $win
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
$proc = Get-Process -Name 'Viber' -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $proc) {
    if ($StartIfNotRunning) {
        $viberPath = Join-Path $env:LOCALAPPDATA 'Viber\Viber.exe'
        if (-not (Test-Path $viberPath)) { $viberPath = Join-Path $env:LOCALAPPDATA 'Programs\Viber\Viber.exe' }
        if (-not (Test-Path $viberPath)) { Write-Error "Viber not running and Viber.exe not found under $env:LOCALAPPDATA." }
        Write-Host ("Starting Viber: {0}" -f $viberPath) -ForegroundColor Cyan
        Start-Process -FilePath $viberPath
        $deadline = (Get-Date).AddSeconds($StartupWaitSeconds)
        do { Start-Sleep -Milliseconds 500; $proc = Get-Process -Name 'Viber' -ErrorAction SilentlyContinue | Select-Object -First 1 }
        while (-not $proc -and (Get-Date) -lt $deadline)
        if (-not $proc) { Write-Error "Viber did not start within $StartupWaitSeconds seconds." }
        Start-Sleep -Seconds 5
    }
    else {
        Write-Error "Viber is not running. Start and sign in to Viber desktop, then retry (or add -StartIfNotRunning)."
    }
}
Write-Host ("Viber PID: {0}" -f $proc.Id) -ForegroundColor Cyan

# Show the window (tray click if hidden).
$win = Show-ViberWindow -ProcessId $proc.Id
if (-not $win) {
    Write-Error "Could not show the Viber window for PID $($proc.Id)."
}
Write-Host ("Viber window: '{0}'" -f $win.Current.Name) -ForegroundColor Cyan

if ($DryRun) {
    Write-Host "DRY RUN - window shown, nothing sent." -ForegroundColor Yellow
    Write-Host ("  Contact : {0}" -f $Contact)
    Write-Host ("  Message : {0}" -f $Message)
    exit 0
}

# 1) Focus the search box and set the contact text.
$search = Find-EditByClass -win $win -pattern 'TextFieldItem'
if (-not $search) { Write-Error "Could not find Viber's search box." }
try { $search.SetFocus() } catch { Write-Warning "SetFocus failed: $_" }
Start-Sleep -Milliseconds 400
Set-ValuePatternValue -el $search -value $Contact
Start-Sleep -Milliseconds $DelayAfterActionMs

# 2) Open the target contact's chat by clicking the top search result.
#    The search results are the VISIBLE list items (those with a non-empty
#    bounding rectangle). Note: Qt reports each item's height as the full list
#    height, so click near the item's TOP edge, not its (wrong) center.
$resultItem = $null
$deadline = (Get-Date).AddMilliseconds(3000)
do {
    $all = $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
    $visible = @($all | Where-Object {
        $_.Current.ClassName -match 'ListViewDelegateLoader' -and -not $_.Current.BoundingRectangle.IsEmpty
    } | Sort-Object { $_.Current.BoundingRectangle.Y })
    if ($visible.Count -gt 0) { $resultItem = $visible[0]; break }
    Start-Sleep -Milliseconds 300
} while ((Get-Date) -lt $deadline)

if ($resultItem) {
    $b = $resultItem.Current.BoundingRectangle
    $cx = [int]($b.X + $b.Width / 2)
    $cy = [int]($b.Y + 25)   # top edge + small offset into the item
    Write-Host ("Opening top result at ({0},{1})" -f $cx, $cy) -ForegroundColor Cyan
    Click-Point -X $cx -Y $cy
    Start-Sleep -Milliseconds $DelayAfterActionMs
}
else {
    Write-Warning "No visible search results - the chat may already be open."
}

# 3) Set the message text in the message box.
$msgBox = Find-EditByClass -win $win -pattern 'QQuickTextEdit'
if (-not $msgBox) {
    # Fallback: an edit near the bottom that is not the search box and not the note/feed editor.
    $editCond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
    $edits = @($win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $editCond) |
        Where-Object { $_.Current.ClassName -notmatch 'TextFieldItem|TextEditItem' } |
        Sort-Object { $_.Current.BoundingRectangle.Y } -Descending)
    if ($edits.Count -gt 0) { $msgBox = $edits[0] }
}
if (-not $msgBox) { Write-Error "Could not find Viber's message box." }
try { $msgBox.SetFocus() } catch { Write-Warning "SetFocus failed: $_" }
Start-Sleep -Milliseconds 300
Set-ValuePatternValue -el $msgBox -value $Message
Start-Sleep -Milliseconds 500

# 4) Send.
if ($NoSend) {
    Write-Host "NO-SEND mode: message set in the chat but NOT sent." -ForegroundColor Yellow
}
else {
    # Re-find the Send button AFTER setting the text (the earlier element goes stale).
    $sendBtn = Find-ButtonByClass -win $win -pattern 'SendToolbarButton'
    if ($sendBtn) {
        Write-Host "Clicking Send button." -ForegroundColor Cyan
        $inv = $sendBtn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
        $inv.Invoke()
        Start-Sleep -Milliseconds 800

        # Verify the message box cleared (a sign the message was sent).
        $check = Find-EditByClass -win $win -pattern 'QQuickTextEdit'
        if ($check) {
            $vpCheck = $check.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern)
            if ([string]::IsNullOrEmpty($vpCheck.Current.Value)) {
                Write-Host "Message sent (input cleared)." -ForegroundColor Green
            }
            else {
                Write-Warning "Message box still has text after Send - the message may not have been sent."
            }
        }
    }
    else {
        Write-Warning "Send button not found - trying Enter key instead."
        [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
    }
}

Write-Host ("Done. Contact='{0}' Message='{1}' (NoSend={2})" -f $Contact, $Message, $NoSend) -ForegroundColor Green
