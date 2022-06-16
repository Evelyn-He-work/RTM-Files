$excel=New-Object -ComObject Excel.Application
$excel.visible=$true
$path='C:\Evelyn Documents\RTM Missing Letters\2019.xlsx'
$excel.Workbooks.Open($path)
$file=$excel.Workbooks.Item("2019")
$all=$file.Worksheets.Item("All")
$allSorted=$file.Worksheets.Item("AllSorted")
$totalRows=80695
$rowA=2
$rowB=2
$rowC=2
$row1=2
$row2=2
$row3=2
$rowOther=2
for($row=1;$row -le $totalRows; $row++){
    if($all.Cells.Item($row,3).Value() -eq "A"){
        $allSorted.Cells.Item($rowA,1)=$all.Cells.Item($row,1).Value()
        $allSorted.Cells.Item($rowA,2)=$all.Cells.Item($row,2).Value()
        $allSorted.Cells.Item($rowA,3)=$all.Cells.Item($row,3).Value()
        $rowA++
    }elseif($all.Cells.Item($row,3).Value() -eq "B"){
        $allSorted.Cells.Item($rowB,4)=$all.Cells.Item($row,1).Value()
        $allSorted.Cells.Item($rowB,5)=$all.Cells.Item($row,2).Value()
        $allSorted.Cells.Item($rowB,6)=$all.Cells.Item($row,3).Value()
        $rowB++
    }elseif($all.Cells.Item($row,3).Value() -eq "C"){
        $allSorted.Cells.Item($rowC,7)=$all.Cells.Item($row,1).Value()
        $allSorted.Cells.Item($rowC,8)=$all.Cells.Item($row,2).Value()
        $allSorted.Cells.Item($rowC,9)=$all.Cells.Item($row,3).Value()
        $rowC++
    }elseif($all.Cells.Item($row,3).Value() -eq "1"){
        $allSorted.Cells.Item($row1,10)=$all.Cells.Item($row,1).Value()
        $allSorted.Cells.Item($row1,11)=$all.Cells.Item($row,2).Value()
        $allSorted.Cells.Item($row1,12)=$all.Cells.Item($row,3).Value()
        $row1++
    }elseif($all.Cells.Item($row,3).Value() -eq "2"){
        $allSorted.Cells.Item($row2,13)=$all.Cells.Item($row,1).Value()
        $allSorted.Cells.Item($row2,14)=$all.Cells.Item($row,2).Value()
        $allSorted.Cells.Item($row2,15)=$all.Cells.Item($row,3).Value()
        $row2++
    }elseif($all.Cells.Item($row,3).Value() -eq "3"){
        $allSorted.Cells.Item($row3,16)=$all.Cells.Item($row,1).Value()
        $allSorted.Cells.Item($row3,17)=$all.Cells.Item($row,2).Value()
        $allSorted.Cells.Item($row3,18)=$all.Cells.Item($row,3).Value()
        $row3++
    }else{
        $allSorted.Cells.Item($rowOther,19)=$all.Cells.Item($row,1).Value()
        $allSorted.Cells.Item($rowOther,20)=$all.Cells.Item($row,2).Value()
        $allSorted.Cells.Item($rowOther,21)=$all.Cells.Item($row,3).Value()
        $rowOther++
    }
}