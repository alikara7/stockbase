# Convert STOK.xlsx sheets to Supabase SQL
$excelPath = "C:\Users\LENOVO\Desktop\STOK.xlsx"
$outputSQL  = "C:\Users\LENOVO\Desktop\stok_import.sql"
$tmpDir     = "C:\Users\LENOVO\Desktop\stok_csv_tmp"

# Actual sheet order:
# 1: HAZIR MAMUL MUSTERİ  -> products yurtici (musteri stogu, firma adlariyla)
# 2: YOM HAZIR MAMUL      -> products fabrika (fabrikanin kendi stogu)
# 3: HAZIR MAMUL IHRACAT  -> products yurtdisi
# 4: BOBİN                -> materials
# 5: KOLİ                 -> othermaterials koli
# 6: POSET                -> othermaterials poset

New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null

Write-Host "Opening Excel..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($excelPath)

$sheetMap = @{1="yurtici_musteriler"; 2="fabrika"; 3="yurtdisi"; 4="bobin"; 5="koli"; 6="poset"}
$xlCSVUTF8 = 62   # xlCSVUTF8 - saves as UTF-8 with BOM (Excel 2016+)
$xlCellTypeVisible = 12
foreach ($idx in ($sheetMap.Keys | Sort-Object)) {
    $name = $sheetMap[$idx]
    $csv  = "$tmpDir\$name.csv"
    $ws   = $wb.Sheets.Item($idx)
    if ($idx -eq 1) {
        # Sheet 1 only: export visible rows to exclude hidden/filtered rows
        $ws.Cells.SpecialCells($xlCellTypeVisible).Copy() | Out-Null
        $newWb = $excel.Workbooks.Add()
        $newWb.Sheets.Item(1).Paste()
        $excel.CutCopyMode = [int]0
        $newWb.SaveAs($csv, $xlCSVUTF8)
        $newWb.Close($false)
    } else {
        # Other sheets: copy whole sheet (no hidden rows issue)
        $ws.Copy()
        $newWb = $excel.ActiveWorkbook
        $newWb.SaveAs($csv, $xlCSVUTF8)
        $newWb.Close($false)
    }
    Write-Host "  Sheet $idx -> $name.csv"
}

$wb.Close($false)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Excel closed."

function EscSQL($str) { return ([string]$str).Trim() -replace "'", "''" }
function NewUUID() { return [guid]::NewGuid().ToString() }
function SafeNum($str) {
    $s = ([string]$str).Trim()
    if ($s -eq "" -or $s -eq $null) { return "0" }
    $n = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return [math]::Round($n,4) }
    if ([double]::TryParse($s, [ref]$n)) { return [math]::Round($n,4) }
    return "0"
}
function IsHeader($val) { return ([string]$val) -match "ÜRÜN|URUN|FİRMA|FIRMA|NO\s*$|^NO$|BİRİM ADET|TOPLAM|MİKTAR" }

$sql = [System.Collections.Generic.List[string]]::new()
$sql.Add("-- STOK.xlsx import $(Get-Date -Format 'yyyy-MM-dd HH:mm')")
$sql.Add("")

# Clean up existing products so we can reimport with correct market values
$sql.Add("-- Clear existing products (reimport with correct market values)")
$sql.Add("DELETE FROM products WHERE market IN ('yurtici','fabrika');")
$sql.Add("")

# YURTİÇİ (Sheet 1: HAZIR MAMUL MUSTERİ - musteri stogu)
Write-Host "Processing yurtici (Sheet 1)..."
$sql.Add("-- Sheet 1: HAZIR MAMUL MUSTERİ -> products (yurtici)")
$rows = Import-Csv "$tmpDir\yurtici_musteriler.csv" -Encoding UTF8 -Header A,B,C,D,E,F,G,H,I | Select-Object -Skip 2
foreach ($row in $rows) {
    $urun = ([string]$row.D).Trim()   # Sheet1: D=urun adi, B=firma, G=miktar
    $firma = ([string]$row.B).Trim()
    if ([string]::IsNullOrWhiteSpace($urun) -or [string]::IsNullOrWhiteSpace($firma)) { continue }
    if ((IsHeader $urun) -or (IsHeader $firma)) { continue }
    $id = NewUUID
    $sql.Add("INSERT INTO products (id,name,company,unit,qty,min,market,category) VALUES ('$id','$(EscSQL $urun)','$(EscSQL $firma)','ADET',$(SafeNum $row.G),0,'yurtici','') ON CONFLICT (id) DO UPDATE SET market=EXCLUDED.market,company=EXCLUDED.company;")
}

# FABRİKA STOĞU (Sheet 2: YOM HAZIR MAMUL - fabrikanin kendi stogu)
Write-Host "Processing fabrika (Sheet 2)..."
$sql.Add("")
$sql.Add("-- Sheet 2: YOM HAZIR MAMUL -> products (fabrika)")
$rows = Import-Csv "$tmpDir\fabrika.csv" -Encoding UTF8 -Header A,B,C,D,E,F,G,H,I | Select-Object -Skip 2
foreach ($row in $rows) {
    $urun = ([string]$row.C).Trim()
    if ([string]::IsNullOrWhiteSpace($urun) -or (IsHeader $urun)) { continue }
    $id = NewUUID
    $sql.Add("INSERT INTO products (id,name,company,unit,qty,min,market,category) VALUES ('$id','$(EscSQL $row.C)',NULL,'$(EscSQL $row.D)',$(SafeNum $row.G),0,'fabrika','') ON CONFLICT (id) DO UPDATE SET market=EXCLUDED.market;")
}

# YURTDIŞI
Write-Host "Processing yurtdisi..."
$sql.Add("")
$sql.Add("-- Sheet 3: HAZIR MAMUL IHRACAT -> products (yurtdisi)")
$rows = Import-Csv "$tmpDir\yurtdisi.csv" -Encoding UTF8 -Header A,B,C,D,E,F,G,H,I | Select-Object -Skip 2
foreach ($row in $rows) {
    $urun = ([string]$row.C).Trim()
    if ([string]::IsNullOrWhiteSpace($urun) -or (IsHeader $urun)) { continue }
    $id = NewUUID
    $sql.Add("INSERT INTO products (id,name,company,unit,qty,min,market,category) VALUES ('$id','$(EscSQL $row.C)','$(EscSQL $row.B)','$(EscSQL $row.D)',$(SafeNum $row.G),0,'yurtdisi','') ON CONFLICT DO NOTHING;")
}

# BOBİN
Write-Host "Processing bobin..."
$sql.Add("")
$sql.Add("-- Sheet 4: BOBİN -> materials")
$rows = Import-Csv "$tmpDir\bobin.csv" -Encoding UTF8 -Header A,B,C,D,E,F,G,H,I | Select-Object -Skip 2
foreach ($row in $rows) {
    $cinsi = ([string]$row.C).Trim()
    $firma = ([string]$row.B).Trim()
    if ([string]::IsNullOrWhiteSpace($cinsi) -or [string]::IsNullOrWhiteSpace($firma)) { continue }
    if ((IsHeader $cinsi) -or (IsHeader $firma)) { continue }
    $parts = @($cinsi)
    if (([string]$row.D).Trim() -ne "") { $parts += ([string]$row.D).Trim() }
    if (([string]$row.E).Trim() -ne "") { $parts += "$(([string]$row.E).Trim()) mm" }
    if (([string]$row.F).Trim() -ne "") { $parts += "$(([string]$row.F).Trim()) µ" }
    $name = ($parts -join " - ").Trim("- ")
    $id = NewUUID
    $sql.Add("INSERT INTO materials (id,name,category,supplier,unit,qty,min) VALUES ('$id','$(EscSQL $name)','$(EscSQL $cinsi)','$(EscSQL $firma)','KG',$(SafeNum $row.G),0) ON CONFLICT DO NOTHING;")
}

# KOLİ
Write-Host "Processing koli..."
$sql.Add("")
$sql.Add("-- Sheet 5: KOLİ -> othermaterials (koli)")
$rows = Import-Csv "$tmpDir\koli.csv" -Encoding UTF8 -Header A,B,C,D,E,F,G,H,I | Select-Object -Skip 2
foreach ($row in $rows) {
    $cinsi   = ([string]$row.B).Trim()
    $tip     = ([string]$row.C).Trim()
    $kullanim= ([string]$row.D).Trim()
    $ebat    = ([string]$row.E).Trim()
    $birim   = ([string]$row.F).Trim()
    if ([string]::IsNullOrWhiteSpace($cinsi) -and [string]::IsNullOrWhiteSpace($kullanim)) { continue }
    if ((IsHeader $cinsi)) { continue }
    $name = if ($kullanim -ne "") { $kullanim } else { $cinsi }
    $cond = if ($tip -match "YENİ|YENI|Yeni") { "yeni" } else { "eski" }
    $id = NewUUID
    $sql.Add("INSERT INTO othermaterials (id,name,category,unit,qty,min,section,condition,measurements) VALUES ('$id','$(EscSQL $name)','$(EscSQL $cinsi)','$(EscSQL $birim)',$(SafeNum $row.G),0,'koli','$cond','$(EscSQL $ebat)') ON CONFLICT DO NOTHING;")
}

# POŞET
Write-Host "Processing poset..."
$sql.Add("")
$sql.Add("-- Sheet 6: POŞET -> othermaterials (poset)")
$rows = Import-Csv "$tmpDir\poset.csv" -Encoding UTF8 -Header A,B,C,D,E,F,G,H,I,J | Select-Object -Skip 2
foreach ($row in $rows) {
    $cinsi   = ([string]$row.B).Trim()
    $kullanim= ([string]$row.D).Trim()
    $ebat    = ([string]$row.E).Trim()
    $birim   = ([string]$row.F).Trim()
    if ([string]::IsNullOrWhiteSpace($cinsi) -and [string]::IsNullOrWhiteSpace($kullanim)) { continue }
    if ((IsHeader $cinsi)) { continue }
    $parts = @()
    if ($cinsi    -ne "") { $parts += $cinsi }
    if ($kullanim -ne "") { $parts += $kullanim }
    $name = ($parts -join " - ").Trim("- ")
    $id = NewUUID
    $sql.Add("INSERT INTO othermaterials (id,name,category,unit,qty,min,section,condition,measurements) VALUES ('$id','$(EscSQL $name)','$(EscSQL $cinsi)','$(EscSQL $birim)',$(SafeNum $row.I),0,'poset','yeni','$(EscSQL $ebat)') ON CONFLICT DO NOTHING;")
}

$sql | Out-File -FilePath $outputSQL -Encoding UTF8
$count = ($sql | Where-Object { $_ -match "^INSERT" }).Count
Write-Host ""
Write-Host "Done! -> $outputSQL"
Write-Host "Total INSERT statements: $count"
Remove-Item $tmpDir -Recurse -Force
