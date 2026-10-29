# Validation-only Office ownership. Never loaded by the shipped library.
function Get-OfficeWindowProcess($Handle) {
    if (!$IsWindows -or !$Handle) { return $null }
    if (!('OoxPdfValidation.OfficeWindow' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace OoxPdfValidation {
    public static class OfficeWindow {
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    }
}
'@
    }
    $windowProcessId = 0u
    [void][OoxPdfValidation.OfficeWindow]::GetWindowThreadProcessId([IntPtr]$Handle, [ref]$windowProcessId)
    if (!$windowProcessId) { return $null }
    Get-Process -Id $windowProcessId -ErrorAction SilentlyContinue
}

function Test-OfficeProcessOwned($Process, [int[]] $ExistingIds, [datetime] $StartedAt, [string] $Name) {
    $null -ne $Process -and $Process.Id -notin $ExistingIds -and
        $Process.ProcessName -ieq $Name -and $Process.StartTime.ToUniversalTime() -ge $StartedAt.ToUniversalTime()
}

function Test-OfficeDocumentOwned($Document, [string] $InputCopy) {
    if ($null -eq $Document -or !$Document.FullName) { return $false }
    [IO.Path]::GetFullPath([string]$Document.FullName).Equals(
        [IO.Path]::GetFullPath($InputCopy),
        $(if ($IsWindows) {[StringComparison]::OrdinalIgnoreCase} else {[StringComparison]::Ordinal}))
}
