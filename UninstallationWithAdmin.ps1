

Write-Host "Remove AsposeCellsExcelAddIn service"
Stop-Service -Name AsposeCellsExcelAddInServices  2> error.log
 
sc.exe delete "AsposeCellsExcelAddInServices" 2> error.log


Write-Host "Remove Share folder : AsposeCellsWebAddIn"
Remove-SmbShare -Name "AsposeCellsWebAddIn" -Force 2> error.log

Write-Host "Remove AsposeCellsWebAddIn folder"
Remove-Item -Path C:\Aspose.Cells.WebAddin -Recurse -Force 2> error.log

Write-Host "Remove TrustedCatalogs"
Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs' -Recurse | ForEach-Object {
    $item = $_ -replace 'HKEY_CURRENT_USER' , 'HKCU:'
    
    Get-ItemProperty $item  | ForEach-Object {         
        if (  $_.Url.Contains("AsposeCellsWebAddIn") ){
            Write-Host $item
            Remove-Item $item -Force 
        }         
    }    
}  

Write-Host "Start to check un-installation result:"
Start-Sleep 4
.\CheckEnv.ps1
