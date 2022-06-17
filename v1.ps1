# ---- DESCRIPTION ---- #
# Puts file name, description, and corresponding version in seperate sheet

#setting up + variable declaration
$excel = New-Object -ComObject Excel.Application
$excel.visible = $true
$path = 'C:\Evelyn Documents\RTM Missing Letters\2019.xlsx'
$excel.Workbooks.Open($path)
$file = $excel.Workbooks.Item("2019")
$all = $file.Worksheets.Item("All")
$og = $file.Worksheets.Item("MSF Expansion")
$new = $file.Worksheets.Item("MSF Expansion Desc")
$nameRange = $all.Range("A2").EntireColumn

#Loop through required documents rows
for ($i = 2; $i -le 1259; $i++) {
    $currName = $og.Cells.Item($i, 1).Value()
    $currVersion = $og.Cells.Item($i, 2).Value()

    #Finds name in document with all letter rows
    $searchName = $nameRange.find($currName)
    if ($null -ne $searchName) {
        $row = $searchName.Row
        Write-Host $row

        #make sure versions match & write information on new sheet
        if ($currVersion -eq $all.Cells.Item($row,3).Value()) {
            $new.Cells.Item($i, 1) = $currName
            $new.Cells.Item($i, 3) = $currVersion
            $new.Cells.Item($i, 2) = $all.Cells.Item($row, 2).Value()
        }
        else {
            $new.Cells.Item($i, 1) = $currName
            $new.Cells.Item($i, 3) = $currVersion
            $new.Cells.Item($i, 2) = "No Description"
            $new.Cells.Item($i,2).Font.ColorIndex="38"
        }
    } else{
        $new.Cells.Item($i,1)=$currName
        $new.Cells.Item($i,3)=$currVersion
        $new.Cells.Item($i,2)="N/A"
        $new.Cells.Item($i,2).Font.ColorIndex="3"
    }
}
Write-Host "done"
