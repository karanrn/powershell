#path to excel file
$Path = 'D:\test.xlsx'

# Open the Excel document and pull in the 'Play' worksheet
$Excel = New-Object -Com Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $False
$Workbook = $Excel.Workbooks.Open($Path) 
$page = 'Sheet1' #name of sheet to work on
$ws = $Workbook.worksheets | Where-Object {$_.Name -eq $page}

$rows = $ws.usedRange.Rows.Count
$columns = $ws.usedRange.Columns.Count

#traverse directory for txt files
#Caution: best suited for files of same type
$files = Get-ChildItem "D:\Cpu-results"
for ($f=0; $f -lt $files.count;$f++){
    #gets name of the file without extension
    $server_name=[io.path]::GetFileNameWithoutExtension($files[$f].FullName)
   
    #open text file skipping first line
    $txtpath = Get-Content($files[$f].FullName) | Select-Object -Skip 1

    $no_of_processors = $txtpath.count
    $no_of_logicalcores=0
    foreach($x in $txtpath){
        $cores=$x
        $no_of_logicalcores+=[int]$x
    }

    for ($i = 1; $i -lt $rows +1 ; $i++) {
        for( $j = 1; $j -lt $columns+1; $j++){

        $result=$ws.Cells.Item($i,$j).Value()
        If( $result -eq $server_name){
            $match_row=$i
            $match_col=$j
            break
            }
        }
    }
    $ws.Cells.Item($match_row,$match_col+1).Value()=$no_of_processors
    $ws.Cells.Item($match_row,$match_col+2).Value()=$cores
    $ws.Cells.Item($match_row,$match_col+3).Value()=$no_of_logicalcores
    $Workbook.SaveAs("D:\test.xlsx")
}
$Workbook.Close()