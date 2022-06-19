$excel = New-Object -ComObject Excel.Application
$excel.visible = $true
$path = 'C:\Evelyn Documents\RTM Missing Letters\Testing\Pull.xlsx'
$excel.Workbooks.Open($path)
$file = $excel.Workbooks.Item("Pull")
$all = $file.Worksheets.Item("All")
$og = $file.Worksheets.Item("MSF Expansion")
$new = $file.Worksheets.Item("MSF Expansion Desc")
$nameRange = $all.Range("A2").EntireColumn