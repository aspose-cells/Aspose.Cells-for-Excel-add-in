
Write-Host "Check AsposeCellsExcelAddIn Services : " -NoNewline
$serviceNumber = $(Get-Service  -Name AsposeCellsExcelAddInServices 2> error.log  | Measure-Object).Count
$serviceStatus = $(Get-Service  -Name AsposeCellsExcelAddInServices 2> error.log ).Status

if( $serviceNumber -eq 0 )
{
    Write-Host "AsposeCellsExcelAddInServices not found."  -ForegroundColor:Red 
}
else
{
    if($serviceStatus -eq "Running")
    {
        Write-Host "AsposeCellsExcelAddInServices is running." -ForegroundColor:Green 
    }
    else
    {
        Write-Host "AsposeCellsExcelAddInServices stoped." -ForegroundColor:Yellow 
    }
}

Write-Host "Check AsposeCellsWebAddIn share folder : " -NoNewline
$shareNumber = $(Get-SmbShare -Name AsposeCellsWebAddIn 2>error.log  | Measure-Object).Count
if( $shareNumber -eq 0 )
{
    Write-Host "AsposeCellsWebAddIn share folder not found."  -ForegroundColor:Red 
}
else
{
    Write-Host "AsposeCellsWebAddIn share folder found." -ForegroundColor:Green 
}

Write-Host "Check TrustedCatalogs : " -NoNewline
Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs' -Recurse | ForEach-Object {
    $item = $_ -replace 'HKEY_CURRENT_USER' , 'HKCU:'
    Get-ItemProperty $item  | ForEach-Object {         
        if (  $_.Url.Contains("AsposeCellsWebAddIn") ){
            Write-Host $_.Url -ForegroundColor:Green 
        } 
    }    
}