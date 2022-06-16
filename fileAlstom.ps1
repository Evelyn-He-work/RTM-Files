$excel = New-Object -ComObject Excel.Application
$excel.visible = $true
$path = 'C:\Evelyn Documents\RTM Missing Letters\2019.xlsx'
$excel.Workbooks.Open($path)
$file = $excel.Workbooks.Item("2019")
$all = $file.Worksheets.Item("All")
$alstom = $file.Worksheets.Item("Alstom")
$alstomDesc = $file.Worksheets.Item("AlstomDesc")
$nameRange = $all.Range("A2").EntireColumn
#$verRange=$all.Range("C2").EntireColumn
for ($i = 2; $i -le 1494; $i++) {
    $currName = $alstom.Cells.Item($i, 1).Value()
    $currVersion = $alstom.Cells.Item($i, 2).Value()
    $searchName = $nameRange.find($currName)
    if ($null -ne $searchName) {
        $row = $searchName.Row
        Write-Host $row
        if ($currVersion -eq $all.Cells.Item($row,3).Value()) {
            $alstomDesc.Cells.Item($i, 1) = $currName
            $alstomDesc.Cells.Item($i, 3) = $currVersion
            $alstomDesc.Cells.Item($i, 2) = $all.Cells.Item($row, 2).Value()
        }
        else {
            $alstomDesc.Cells.Item($i, 1) = $currName
            $alstomDesc.Cells.Item($i, 3) = $currVersion
            $alstomDesc.Cells.Item($i, 2) = "No Description"
            $alstomDesc.Cells.Item($i,2).Font.ColorIndex="38"
        }
    } else{
        $alstomDesc.Cells.Item($i,1)=$currName
        $alstomDesc.Cells.Item($i,3)=$currVersion
        $alstomDesc.Cells.Item($i,2)="N/A"
        $alstomDesc.Cells.Item($i,2).Font.ColorIndex="3"
    }
}
Write-Host "done"
