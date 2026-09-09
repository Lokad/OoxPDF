# Shared comparison math (T04): CenterX/CenterY, Delta, RoundedKey, and
# Group-Count used by chart/text comparison and summary tools. RoundedKey
# takes an optional digit count defaulting to 6. Dot-source this file
# instead of copying the functions into comparison and summary tools.

function CenterX($op) { return ([double]$op.MinX + [double]$op.MaxX) / 2d }
function CenterY($op) { return ([double]$op.MinY + [double]$op.MaxY) / 2d }
function Delta([double] $left, [double] $right) { return [Math]::Round($right - $left, 6) }
function RoundedKey($Value, [int] $Digits = 6) {
    if ($null -eq $Value -or [string]$Value -eq "") {
        return "(missing)"
    }

    return ([Math]::Round([double]$Value, $Digits)).ToString("0.######", [Globalization.CultureInfo]::InvariantCulture)
}
function Group-Count($items, [scriptblock] $keySelector) {
    $groups = @{}
    foreach ($item in $items) {
        $key = & $keySelector $item
        if ($null -eq $key -or [string]$key -eq "") {
            $key = "(missing)"
        }

        $key = [string]$key
        if ($groups.ContainsKey($key)) {
            $groups[$key]++
        }
        else {
            $groups[$key] = 1
        }
    }

    return @(
        foreach ($key in ($groups.Keys | Sort-Object)) {
            [pscustomobject]@{
                Key = $key
                Count = $groups[$key]
            }
        }
    )
}
