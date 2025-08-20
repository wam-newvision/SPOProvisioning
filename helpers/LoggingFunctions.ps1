# --------------------------------------------------------------------

function Log {
    param()
    if ($env:DEBUG -eq 'true') { 
        $caller = Get-PSCallStack | Select-Object -Skip 1 -First 1
        $fn = if ($caller -and $caller.FunctionName -ne '<ScriptBlock>') {
            $caller.FunctionName
        } elseif ($MyInvocation.ScriptName) {
            Split-Path $MyInvocation.ScriptName -Leaf
        } else {
            '<Script>'
        }
        Write-Host "[$fn] $($args -join ' ')"
    }
}

# ------------------------------------------------------------------------
# Fehlerbehandlung Helper
# ------------------------------------------------------------------------

function ErrorExit {
    param([Parameter(ValueFromRemainingArguments=$true)]$Text)
    Write-Host "❌ $($Text -join ' ')" -ForegroundColor Red
    exit 1
}
# --------------------------------------------------------------------

function Send-Resp([int]$code, [object]$body) {
    Push-OutputBinding -Name Response -Value @{ StatusCode = $code; Body = $body }
}
# --------------------------------------------------------------------