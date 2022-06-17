# ---- DESCRIPTION ---- #
# In addition to v1, organizing letter information into appropriate columns, keeps everything on the same page
function Remove-UdfFiletype($fileType){
    $ogRange1 = $og.Range("A2").EntireColumn
    $searchType = $ogRange1.find($fileType)

    if ($null -ne $searchType) {
        $firstAddress = $searchType.Address
        do {
            $searchType.Value() = $searchType.Value() -Replace $fileType, ''
            $searchType = $ogRange1.FindNext($searchType)
        } while ($null -ne $searchType -and $searchType.Address -ne $firstAddress)
    }
}

#setting up + variable declaration
$excel = New-Object -ComObject Excel.Application
$excel.visible = $true
$path = 'C:\Evelyn Documents\RTM Missing Letters\2019.xlsx'
$excel.Workbooks.Open($path)
$file = $excel.Workbooks.Item("2019")
$all = $file.Worksheets.Item("All")
$og = $file.Worksheets.Item("Guideway")
$numRows=260
#$new = $file.Worksheets.Item("MSF Expansion Desc")
$allRange = $all.Range("A2").EntireColumn

#Eliminates file types from name
Remove-UdfFiletype ".pdf"
Remove-UdfFiletype ".xlsx"
Remove-UdfFiletype ".docx"
Remove-UdfFiletype ".dwg"
Remove-UdfFiletype ".xml"
Remove-UdfFiletype ".tif"
Remove-UdfFiletype ".zip"
Remove-UdfFiletype ".pptx"

#Seperates name and version by last delimiter '_'
for($i=2; $i -le $numRows; $i++){
    $string=$og.Cells.Item($i,1).Value()
    $delimIndex=$string.LastIndexOf('_')

    if($delimIndex -gt 0){
        write-host $string.length
        write-host $string
        write-host $delimIndex
        $og.Cells.Item($i,2)=$string.substring($delimIndex+1, $string.length-$delimIndex-1)
        $og.Cells.Item($i,1)=$string.substring(0,$delimIndex)
    }
}

#Loop through required documents rows
for ($i = 2; $i -le $numRows; $i++) {
    $currName = $og.Cells.Item($i, 1).Value()

    #Finds name in document with all letter rows
    $searchName = $allRange.find($currName)
    if ($null -ne $searchName) {
        $row = $searchName.Row
        Write-Host $row

        $og.Cells.Item($i, 3) = $all.Cells.Item($row, 2).Value()

    }
    else {
        $og.Cells.Item($i, 3) = "N/A"
        $og.Cells.Item($i, 3).Font.ColorIndex = "3"
    }

}
Write-Host "done"
