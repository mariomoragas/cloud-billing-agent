$source = 'C:\Users\maricost\OneDrive - Oracle Corporation\#Projetos\FY27\SkyOne\Projeto Alterdata\Infos do Cliente\inventario-hml-consolidado-2026-09-01.html'
$outputDir = Join-Path $PSScriptRoot 'output'
$outputFile = Join-Path $outputDir 'inventario-hml-consolidado-2026-09-01.csv'
$html = Get-Content -LiteralPath $source -Raw

$sections = @(
    'Resumo do inventario',
    'Origem dos dados (por ambiente)',
    '1. Recursos por tipo',
    '2. Distribuicao por regiao',
    '3. Virtual Machines (tamanho e SO)',
    '4. Discos (tipo e faixa de tamanho)',
    '5. App Service Plans (por SKU)',
    '6. App Services (plataforma/runtime)',
    '7. Bancos de dados gerenciados'
)

$tables = [regex]::Matches($html, '(?is)<table[^>]*>(.*?)</table>')
$allColumns = [System.Collections.Generic.List[string]]::new()
$allColumns.Add('Secao')
$records = [System.Collections.Generic.List[object]]::new()

for ($i = 0; $i -lt $tables.Count; $i++) {
    $table = $tables[$i].Groups[1].Value
    $headers = @(
        [regex]::Matches($table, '(?is)<th[^>]*>(.*?)</th>') |
            ForEach-Object {
                [System.Net.WebUtility]::HtmlDecode(
                    ($_.Groups[1].Value -replace '<[^>]+>', '').Trim()
                )
            }
    )

    foreach ($header in $headers) {
        if (-not $allColumns.Contains($header)) {
            $allColumns.Add($header)
        }
    }

    $bodyMatch = [regex]::Match($table, '(?is)<tbody[^>]*>(.*?)</tbody>')
    $tableRows = [regex]::Matches($bodyMatch.Groups[1].Value, '(?is)<tr[^>]*>(.*?)</tr>')
    foreach ($tr in $tableRows) {
        $cells = @(
            [regex]::Matches($tr.Groups[1].Value, '(?is)<td[^>]*>(.*?)</td>') |
                ForEach-Object {
                    [System.Net.WebUtility]::HtmlDecode(
                        ($_.Groups[1].Value -replace '<[^>]+>', '').Trim()
                    )
                }
        )

        if ($cells.Count -gt 0) {
            $row = [ordered]@{ Secao = $sections[$i] }
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $value = if ($c -lt $cells.Count) { $cells[$c] } else { $null }
                $row[$headers[$c]] = $value
            }
            $records.Add([pscustomobject]$row)
        }
    }
}

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
$records |
    Select-Object -Property $allColumns |
    Export-Csv -LiteralPath $outputFile -NoTypeInformation -Encoding UTF8 -Delimiter ';'

$csvRows = Import-Csv -LiteralPath $outputFile -Delimiter ';'
Get-Item -LiteralPath $outputFile | Select-Object FullName, Length, LastWriteTime
"Registros=$($csvRows.Count)"
"Colunas=$($allColumns -join ' | ')"
$csvRows | Group-Object Secao | Select-Object Name, Count | Format-Table -AutoSize
