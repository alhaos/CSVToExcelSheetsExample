$application = New-Object -ComObject Excel.Application
$DebugPreference = 'Continue'
Set-StrictMode -Version 'latest'
function MakeGood {
    param (
        [string]$ExcelFile,
        [string[]]$CSVFiles
    )
    
    $w = $application.Workbooks.Add()

    foreach ($file in $CSVFiles) {
        $s = $w.Sheets.Add()
        $s.Name = [System.IO.Path]::GetFileName($file)
        
        $csv = Import-Csv $file
        $csvEn = $csv.GetEnumerator()
                
        $fields = ($csv | Get-Member -MemberType NoteProperty).Name
        $FieldsEn = $fields.GetEnumerator()
        
        $rowBill = 0

        while ($csvEn.MoveNext()) {
            $fieldsBill = 0
            $rowBill++
            $FieldsEn.Reset()
            while ($FieldsEn.MoveNext()) {
                $s.Cells.Item($rowBill, (++$fieldsBill)).Value = $csvEn.Current.($FieldsEn.Current)
                Write-Debug ("{0}, {1}" -f $rowBill, $fieldsBill)
            }
        }
    }
    $w.SaveAs($ExcelFile)
}

function Exit-Appication {
    $application.Quit()    
}