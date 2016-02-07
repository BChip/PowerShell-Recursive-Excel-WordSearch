# PowerShell -Recurse parameter
Clear-Host
Set-StrictMode -Version latest
$path = "C:\Users\Brad\"
$files = Get-Childitem $path -Include *.xls,*.xlsx –Force –Recurse –ErrorAction SilentlyContinue –ErrorVariable AccessDenied | Where-Object { !($_.psiscontainer) }
$word = "test"
$Excel = New-Object -comobject excel.application
$Excel.visible = $False

foreach($file In $files){
    $ExcelWorkBook = $Excel.Workbooks.Open($file)
    $Worksheets = $ExcelWorkBook.worksheets
    foreach($worksheet In $Worksheets){
        $Range = $Worksheet.Range("A1:Z1").EntireColumn
        $found = $false
        $found = $Range.find($word)
        if($found){
            $file | Add-Content -path "test.csv"
        }
    }
    $ExcelWorkBook.close();
}
$Excel.Quit()