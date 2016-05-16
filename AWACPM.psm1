function Get-FarmXml
{
    $farm = Import-Clixml "$($ConfigurationFile)\wacFarm.xml"
    return $farm
}

function Get-MachineXml
{
    $machine = Import-CliXml "$($ConfigurationFile)\$($env:COMPUTERNAME)-machine.xml"
    return $machine
}

function Set-MachineRole
{
    Set-OfficeWebAppsMachine -Roles $machine.Roles
}
function ConvertTo-Boolean
{
  param
  (
    [Parameter(Mandatory=$false)][string] $value
  )
  switch ($value)
  {
    'true' { return $true; }
    'false' { return $false; }
  }
}

function Invoke-Patch
{
    Remove-OfficeWebAppsMachine -Confirm:$false
    Stop-Service WACSM
    iisreset /stop
    $p = Start-Process $PatchFile -ArgumentList '/passive' -Wait -PassThru -NoNewWindow
    if(!($p.ExitCode -eq 0) -and !($p.ExitCode -eq 3010) -and !($p.ExitCode -eq 17022)){
        throw [System.Configuration.Install.InstallException] "The Office Web Apps patch failed to install. ExitCode: $($p.ExitCode)" 
    }

    iisreset.exe /start

    if(($p.ExitCode -eq 3010) -or ($p.ExitCode -eq 17022))
    {
        New-Item -Path HKLM:\SOFTWARE\Microsoft\WACPatch -Force
        New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WACPatch -Name RebootRequired -Value 1 -PropertyType DWORD
        Write-Host -ForegroundColor Yellow 'A reboot is required to complete setup.'
        break
    }
}

function Get-ResumeInstallation
{
    try{
        if((Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WACPatch -Name RebootRequired -EA Stop).RebootRequired -eq 1)
        {
            Write-Host -ForegroundColor Yellow '`tResuming from reboot.'
            Remove-Item -Path HKLM:\SOFTWARE\Microsoft\WACPatch -Force
            $skipInstall = $true
            return $skipInstall
        }
    }
    catch{
        Write-Host -ForegroundColor Yellow '`t Not returning from a reboot.'
        $skipInstall = $false
        return $skipInstall
    }
}

function Start-OfficeWebAppsPatch
{
    param
    (
        [string]
        [Parameter(Mandatory=$true)]
        $ConfigurationFile,
        [string]
        [Parameter(Mandatory=$true)]
        $PatchFile,
        [bool]
        [Parameter(Mandatory=$true)]
        $SetAsLeadHost,
        [string]
        [Parameter(Mandatory=$false)]
        $MachineToJoin
    )

    if(!(Test-Path -IsValid $ConfigurationFile))
    {
        Write-Host -ForegroundColor Red 'Configuration file path is invalid.'
        break
    }

    if(!(Test-Path -IsValid $PatchFile))
    {
        Write-Host -ForegroundColor Red 'Unable to find patch.'
        break
    }

    if ($SetAsLeadHost -eq $false -and [string]::IsNullOrEmpty($MachineToJoin))
    {
        Write-Host -ForegroundColor Red 'Specify the MachineToJoin parameter.'
        break
    }

    $skipInstall = Get-ResumeInstallation

    if ($skipInstall -eq $false)
    {
        $ErrorActionPreference = 'Stop'
    if((Get-OfficeWebAppsFarm).Machines.Count -gt 1 -and (Get-OfficeWebAppsMachine).MasterMachineName -eq $env:COMPUTERNAME)
    {
        Write-host -ForegroundColor Red 'This is the master machine. Patch other farm members first.'
        break
    }

    $ErrorActionPreference = 'Continue'
    }

    if($SetAsLeadHost)
    {
        if($skipInstall -eq $false)
        {
            Write-Host -ForegroundColor Green "`tExporting Office Web Apps Confguration."
            Get-OfficeWebAppsFarm | Export-Clixml "$($ConfigurationFile)\wacFarm.xml"
            Get-OfficeWebAppsMachine | Export-Clixml "$($ConfigurationFile)\$($env:COMPUTERNAME)-machine.xml"
            Write-Host -ForegroundColor Green "`tStarting patch process."
            Invoke-Patch
        }

        $farm = Get-FarmXml

        if([string]::IsNullOrEmpty($farm))
        {
            Write-Host -ForegroundColor Red 'Something went wrong getting the XML data.'
            break
        }

        $parms = @{'FarmOU'=$farm.FarmOU;'AllowHttp'=(ConvertTo-Boolean($farm.AllowHTTP));'SSLOffloaded'=(ConvertTo-Boolean($farm.SSLOffloaded));`
            'CertificateName'=$farm.CertificateName;'EditingEnabled'=(ConvertTo-Boolean($farm.EditingEnabled));'Proxy'=$farm.Proxy;`
            'LogLocation'=$farm.LogLocation; 'LogRetentionInDays'=$farm.LogRetentionInDays;'LogVerbosity'=$farm.LogVerbosity;`
            'CacheLocation'=$farm.CacheLocation;'MaxMemoryCacheSizeInMB'=$farm.MaxMemoryCacheSizeInMB;'DocumentInfoCacheSize'=$farm.DocumentInfoCacheSize;`
            'CacheSizeInGB'=$farm.CacheSizeInGB;'ClipartEnabled'=(ConvertTo-Boolean([bool]$farm.ClipartEnabled));`
            'IgnoreDeserializationFilter'=(ConvertTo-Boolean($farm.IgnoreDeserializationFilter));'TranslationEnabled'=(ConvertTo-Boolean($farm.TranslationEnabled));`
            'MaxTranslationCharacterCount'=$farm.MaxTranslationCharacterCount;'TranslationServiceAppId'=$farm.TranslationServiceAppId;`
            'TranslationServiceAddress'=$farm.TranslationServiceAddress;'RenderingLocalCacheLocation'=$farm.RenderingLocalCacheLocation;`
            'RecycleActiveProcessCount'=$farm.RecycleActiveProcessCount;'AllowCEIP'=(ConvertTo-Boolean([bool]$farm.AllowCEIP));`
            'ExcelRequestDurationMax'=$farm.ExcelRequestDurationMax;'ExcelSessionTimeout'=$farm.ExcelSessionTimeout;`
            'ExcelWorkbookSizeMax'=$farm.ExcelWorkbookSizeMax;'ExcelPrivateBytesMax'=$farm.ExcelPrivateBytesMax;`
            'ExcelConnectionLifetime'=$farm.ExcelConnectionLifetime;'ExcelExternalDataCacheLifetime'=$farm.ExcelExternalDataCacheLifetime;`
            'ExcelAllowExternalData'=(ConvertTo-Boolean($farm.ExcelAllowExternalData));'ExcelWarnOnDataRefresh'=(ConvertTo-Boolean($farm.ExcelWarnOnDataRefresh));`
            'OpenFromUrlEnabled'=(ConvertTo-Boolean([bool]$farm.OpenFromUncEnabled));'OpenFromUncEnabled'=(ConvertTo-Boolean($farm.OpenFromUncEnabled));`
            'OpenFromUrlThrottlingEnabled'=(ConvertTo-Boolean([bool]$farm.OpenFromUrlThrottlingEnabled));'PicturePasteDisabled'=(ConvertTo-Boolean([bool]$farm.PicturePasteDisabled));`
            'RemovePersonalInformationFromLogs'=(ConvertTo-Boolean($farm.RemovePersonalInformationFromLogs));`
            'AllowHttpSecureStoreConnections'=(ConvertTo-Boolean($farm.AllowHttpSecureStoreConnections));'Confirm'=$false}


        if($farm.InternalURL -ne $null)
        {
            $parms.Add('InternalUrl', $farm.InternalURL)
        }
        elseif($farm.ExternalURL -ne $null)
        {
            $parms.Add('ExternalUrl', $farm.ExternalURL)
        }
        else
        {
            Write-Host -ForegroundColor Red 'No Internal or External Url defined.'
            break
        }

        if($parms.Contains('EditingEnabled') -and $parms['EditingEnabled'] -eq $true)
        {
            Write-Host -ForegroundColor Yellow 'Editing is enabled. You will be prompted to confirm enabling this setting.'
        }

        Write-Host -ForegroundColor Green 'Creating Office Web Apps Farm.'
        New-OfficeWebAppsFarm @parms
        Write-host -ForegroundColor Green 'Completed creating farm.'
        Write-Host -ForegroundColor Green 'Setting Office Web Apps Machine Roles.'
        Get-MachineXml
        Set-OfficeWebAppsMachine -Roles $machine.Roles
        Write-Host -ForegroundColor Green 'Completed setting machine roles.'
    }
    else
    {
        if($skipInstall -eq $false)
        {
            Write-Host -ForegroundColor Green "`tExporting Office Web Apps Machine Confguration."
            Get-OfficeWebAppsMachine | Export-Clixml "$($ConfigurationFile)\$($env:COMPUTERNAME)-machine.xml"
            Write-Host -ForegroundColor Green "`tStarting patch process."
            Invoke-Patch
        }
        $machine = Get-MachineXml
        Write-Host -ForegroundColor Green 'Adding machine to existing Office Web Apps farm.'
        New-OfficeWebAppsMachine -Roles $machine.Roles -MachineToJoin $MachineToJoin
        Write-Host -ForegroundColor Green 'Completed adding server to farm.'
    }
}