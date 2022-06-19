# ---- DESCRIPTION ---- #
# In addition to v2, automating sorting of all sheets + optimizations in name/description organization

function Remove-UdfFiletype($fileType) {
    $newRange1 = $newSheet.Range("A2").EntireColumn
    $searchType = $newRange1.find($fileType)

    if ($null -ne $searchType) {
        $firstAddress = $searchType.Address
        do {
            $searchType.Value() = $searchType.Value() -Replace $fileType, ''
            $searchType = $newRange1.FindNext($searchType)
        } while ($null -ne $searchType -and $searchType.Address -ne $firstAddress)
    }
}

#setting up + variable declaration
$excel = New-Object -ComObject Excel.Application
$excel.visible = $true

$destPath = 'C:\Evelyn Documents\RTM Missing Letters\2019\2019.xlsx'
$sourcePath = 'C:\Evelyn Documents\RTM Missing Letters\2019\List of Project Documents_2019 (RTM Master).xlsx'
$excel.Workbooks.Open($destPath)
$excel.Workbooks.Open($sourcePath)

$destFile = $excel.Workbooks.Item("2019")
$sourceFile = $excel.Workbooks.Item("List of Project Documents_2019 (RTM Master)")
$all = $destFile.Worksheets.Item("All")
$allRange = $all.Range("A2").EntireColumn

$sourceSheet = $sourceFile.Worksheets(1)

#Loops through sheets ----##### REMEMBER TO EXCLUDE 'ALL' FILE FROM LOOP ###
foreach ($sourceSheet in $sourceFile.Worksheets) {
    $sheetName = $sourceSheet.name
    $newSheet = $destFile.Worksheets.Add()
    $newSheet.name = $sheetName

    $newSheet.Cells.Item(1, 1) = "Name"
    $newSheet.Cells.Item(1, 2) = "Version"
    $newSheet.Cells.Item(1, 3) = "Description"

    #Copies relevant sheet information over
    $numRows = $sourceSheet.UsedRange.rows.count
    $sourceSheet.Range("A2:A$numRows").Copy()
    $newSheet.Range("A2").PasteSpecial(-4163) #-4163 to paste values
    
    #Accounts for header
    $numRows = $numRows + 1
    
    #Eliminates file types from name
    for ($i = 2; $i -le $numRows; $i++) {
        $string = $newSheet.Cells.Item($i, 1).Value()

        if ($string -like '*.*') {
            $delimIndex = $string.LastIndexOf('.')
            $newSheet.Cells.Item($i, 1) = $string.substring(0, $delimIndex)
        }
    }

    Remove-UdfFiletype ".pdf"
    Remove-UdfFiletype ".xlsx"
    Remove-UdfFiletype ".docx"
    Remove-UdfFiletype ".dwg"
    Remove-UdfFiletype ".xml"
    Remove-UdfFiletype ".tif"
    Remove-UdfFiletype ".zip"
    Remove-UdfFiletype ".pptx"
    Remove-UdfFiletype ".txt"
    Remove-UdfFiletype ".doc"

    #Seperates name and version by last delimiter '_'
    for ($i = 2; $i -le $numRows; $i++) {
        $string = $newSheet.Cells.Item($i, 1).Value()

        if ($string -like '*_*') {
            $delimIndex = $string.LastIndexOf('_')
            $newSheet.Cells.Item($i, 2) = $string.substring($delimIndex + 1, $string.length - $delimIndex - 1)
            $newSheet.Cells.Item($i, 1) = $string.substring(0, $delimIndex)
        }
    }

    #Matches description to name
    for ($i = 2; $i -le $numRows; $i++) {
        $currName = $newSheet.Cells.Item($i, 1).Value()

        #disregards version when searching for description
        if ($currName -like '*_*') {
            $delimIndex = $currName.LastIndexOf('_')
            if (($currName.length - $delimIndex) -le 4) {
                $currName = $currName.substring(0, $delimIndex)
                write-host $currName
            }
        }

        #Finds name in document with all letter rows
        $searchName = $allRange.find($currName)
        if ($null -ne $searchName) {
            $row = $searchName.Row
            $newSheet.Cells.Item($i, 3) = $all.Cells.Item($row, 2).Value()
        }
        else {
            $newSheet.Cells.Item($i, 3) = "N/A"
            $newSheet.Cells.Item($i, 3).Font.ColorIndex = "3"
        }
    }
    Write-Host "done $sheetName"
}
Write-Host "done all"