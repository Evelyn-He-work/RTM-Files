# ---- DESCRIPTION ---- #
# In addition to v3, automated copying of master files and formatting.
# Everything automated except for making workbook & closing workbook 

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

function Copy-UdfMasterFile($start) {
    $master.Range($start).EntireColumn.Copy() | out-null
    $newSheet.Range($start).PasteSpecial(-4163) | out-null
}

#setting up + variable declaration
$excel = New-Object -ComObject Excel.Application
$excel.visible = $true

$destPath = 'C:\Evelyn Documents\00_Projects\RTM Missing Letters\2021 Handover\2021 Handover.xlsx'
$sourcePath = 'C:\Evelyn Documents\00_Projects\RTM Missing Letters\2021 Handover\List of Handover Documents_2021.xlsx'
$masterPath = 'C:\Evelyn Documents\00_Projects\RTM Missing Letters\2021 Handover\Stage 1 - List of All Documents 2021.xlsx'
$excel.Workbooks.Open($destPath)
$excel.Workbooks.Open($sourcePath)
$excel.Workbooks.Open($masterPath)

$destFile = $excel.Workbooks.Item("2021 Handover")
$sourceFile = $excel.Workbooks.Item("List of Handover Documents_2021")
$masterFile = $excel.Workbooks.Item("Stage 1 - List of All Documents 2021")

#Copy OLRTC Master to new worksheet
$master=$masterFile.Worksheets.Item("Design Drawings")
$newSheet=$destFile.Worksheets.Add()
$newSheet.name = "All"
Copy-UdfMasterFile "A1"
Copy-UdfMasterFile "C1"

$all = $destFile.Worksheets.Item("All")
$allRange = $all.Range("A2").EntireColumn

#Loops through sheets
$sourceSheet = $sourceFile.Worksheets(1)

foreach ($sourceSheet in $sourceFile.Worksheets) {
    $sheetName = $sourceSheet.name
    $newSheet = $destFile.Worksheets.Add()
    $newSheet.name = $sheetName

    $newSheet.Cells.Item(1, 1) = "Name"
    $newSheet.Cells.Item(1, 2) = "Version"
    $newSheet.Cells.Item(1, 3) = "Description"

    #Copies relevant sheet information over
    $numRows = $sourceSheet.UsedRange.rows.count
    $sourceSheet.Range("A2:A$numRows").Copy() | out-null
    $newSheet.Range("A2").PasteSpecial(-4163) | out-null #-4163 to paste values
    
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
Write-Host "done descriptions"

$destSheet=$destFile.Worksheets.Item(1)

#formatting
foreach ($destSheet in $destFile.Worksheets){
    $destSheet.UsedRange.Columns.Autofit() | out-null
}

#$destFile.close($false)
#$sourceFile.close($true)
#$masterFile.close($false)
#$excel.quit()

Write-Host "done all"