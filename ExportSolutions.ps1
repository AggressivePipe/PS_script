$error.clear()
$solutionName = "WavinCustomizations"
$SolutionFilePath = "C:\ExportedSolutions"
$ManagedSolution = $true
$ErrorOccured = $false

Set-StrictMode -Version latest

$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

function InstallRequiredModule {
    try {
        Get-PSRepository -WarningVariable wv
        if ($wv.ToString() -eq "Unable to find module repositories") {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Register-PSRepository -Default
        }
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
        $moduleName = "Microsoft.Xrm.Data.Powershell"
        $moduleVersion = "2.7.2"
        if (!(Get-Module -ListAvailable -Name $moduleName)) {
            Write-host "Module Not found, installing now"
            $moduleVersion
            Install-Module -Name $moduleName -MinimumVersion $moduleVersion -Force
        } else {
            Write-host "Module Found"
        }
    } catch {
        "Error occurred"
        $ErrorOccured = $true
    }
    if (!$ErrorOccured) { "No Error Occurred" }
}

function EstablishCRMConnection {
    try {
        Write-Host "Establishing CRM connection"
        $crm = Get-CrmConnection -InteractiveMode
        Write-Host "CRM connection established"
    } catch {
        Write-Host $_.Exception
    }
    return $crm
}

InstallRequiredModule

Write-Host "Creating source connection"
$CrmSourceConnectionString = EstablishCRMConnection
Write-Host "Source connection created"
Set-CrmConnectionTimeout -conn $CrmSourceConnectionString -TimeoutInSeconds 1000

Write-Host "Publishing Customizations in source environment"
Publish-CrmAllCustomization -conn $CrmSourceConnectionString
Write-Host "Publishing Completed in source environment."

if($ManagedSolution -eq $true){
    $solutionFileName = $solutionName + "_managed";
}
else{
    $solutionFileName = $solutionName + "_unmanaged";
}

Write-Host "Exporting Solution"
Export-CrmSolution -conn $CrmSourceConnectionString -SolutionName $solutionName -SolutionFilePath $SolutionFilePath -SolutionZipFileName "$solutionFileName.zip" -Managed:$ManagedSolution
Write-host "Solution Exported."

$stopwatch.Stop()
$elapsedTime = $stopwatch.Elapsed.ToString()

Write-Host "Script execution time: $elapsedTime"