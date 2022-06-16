$excel = New-Object -ComObject Excel.Application
$excel.visible = $true
$path = 'C:\Evelyn Documents\RTM Missing Letters\2019.xlsx'
$excel.Workbooks.Open($path)
$file = $excel.Workbooks.Item("2019")
$all = $file.Worksheets.Item("All")
$contracts = $file.Worksheets.Item("Contracts & Suppliers")
$contractsDesc = $file.Worksheets.Item("Contracts & Suppliers Desc")
$nameRange = $all.Range("A1").EntireColumn
for ($i = 2; $i -le 15625; $i++) {
    $currName = $contracts.Cells.Item("$i,1").Value()
    $currVersion = $contracts.Cells.Item("$i,2").Value()
    $searchName = $nameRange.find($currName)
    if ($null -ne $searchName) {
        $row = $searchName.Row
        Write-Host $row
        if ($currVersion -eq $all.Cells.Item($row, 3).Value()) {
            $contractsDesc.Cells.Item($i, 1) = $currName
            $contractsDesc.Cells.Item($i, 3) = $currVersion
            $contractsDesc.Cells.Item($i, 2) = $all.Cells.Item($row, 2).Value()
        }
        else {
            $contractsDesc.Cells.Item($i, 1) = $currName
            $contractsDesc.Cells.Item($i, 3) = $currVersion
            $contractsDesc.Cells.Item($i, 2) = "No Description"
            $contractsDesc.Cells.Item($i, 2).Font.ColorIndex = "38"
        }
    }
    else {
        $contractsDesc.Cells.Item($i, 1) = $currName
        $contractsDesc.Cells.Item($i, 3) = $currVersion
        $contractsDesc.Cells.Item($i, 2) = "N/A"
        $contractsDesc.Cells.Item($i, 2).Font.ColorIndex = "3"
    }
}
Write-Host"done"