Function Create-DataBase($Db){
 $application = New-Object -ComObject Access.Application
 $application.NewCurrentDataBase($Db)
 $application.CloseCurrentDataBase()
 $application.Quit()
}

$acExport = 1
$acSpreadsheetTypeExcel9 = 8
$dbName="C:\Users\ignazio\Desktop\Nuova cartella\Test20151129T06060.accdb"
Create-DataBase $dbName

$db = New-Object -Comobject Access.Application
$db.OpenCurrentDatabase("C:\Users\ignazio\Desktop\Nuova cartella\Test20151129T06060.accdb")


$FileXls = get-childitem "C:\Users\ignazio\Desktop\Nuova cartella" -recurse | where {$_.extension -eq ".xls"} 

for ($i=0; $i -lt $FileXls.Count; $i++) {
   $fileNameImported=$FileXls[$i].BaseName;
   Write-Progress -Activity "importing xls " -status "now importing $fileNameImported" -percentComplete ($i / $FileXls.count*100)
   $db.DoCmd.TransferSpreadsheet($acImport, $acSpreadsheetTypeExcel9,  $FileXls[$i].BaseName, `
$FileXls[$i].FullName, $True)

}
$db.Quit()
