$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "PVS" {

    BeforeAll {

        #Declaritive json file location
        $path = $PSScriptRoot

        #get data into object
        $jsonLocation = Join-Path $path pvs.json
        $pvsData = Get-Content $jsonLocation | ConvertFrom-Json

        #region Get External data
        $local = $true

        if ($local) {
            $pvsVersion = Import-CliXml (Join-Path $path xml\pvsVersion.xml)
            $pvsFarm = Import-CliXml (Join-Path $path xml\pvsFarm.xml)
            $pvsAuthGroupFarm = Import-CliXml (Join-Path $path xml\pvsAuthGroupFarm.xml)
            $pvsAuthGroup = Import-CliXml (Join-Path $path xml\pvsAuthGroup.xml)
            $pvsCeipData = Import-CliXml (Join-Path $path xml\pvsCeipData.xml)
            $pvsSite = Import-CliXml (Join-Path $path xml\pvsSite.xml)
            $pvsServer = Import-CliXml (Join-Path $path xml\pvsServer.xml)
            $pvsStore = Import-CliXml (Join-Path $path xml\pvsStore.xml)
            $pvsCisData = Import-CliXml (Join-Path $path xml\pvsCisData.xml)
            $pvsDiskInfo = Import-CliXml (Join-Path $path xml\pvsDiskInfo.xml)
            $pvsDiskUpdateDevice = Import-CliXml (Join-Path $path xml\pvsDiskUpdateDevice.xml)
        }
        else {
            #Get Data from PVS Farm

            Import-Module "C:\Program Files\Citrix\Provisioning Services Console\Citrix.PVS.SnapIn.dll"

            $pvsVersion = Get-PVSVersion
            $pvsFarm = Get-PVSFarm
            $pvsAuthGroupFarm = Get-PVSAuthGroup -Farm
            $pvsAuthGroup = Get-PVSAuthGroup
            $pvsCeipData = Get-PVSCeipData
            $pvsSite = Get-PVSSite
            $pvsServer = Get-PVSServer
            $pvsStore = Get-PVSStore
            $pvsCisData = Get-PVSCisData
            $pvsDiskInfo = Get-PVSDiskInfo
            $pvsDiskUpdateDevice = Get-PvsdiskUpdateDevice
            <#
            $pvsServerBootstrapName = Get-PvsServerBootstrapName -ServerName $server.serverName
            $serverPvsServerBootstrap = Get-PvsServerBootstrap -Name $Bootstrapname.name -ServerName $server.serverName
            $Stores = Get-PVSStore -ServerName $Server.serverName
            $Bootstraps = Get-PvsDeviceBootstrap -deviceName $Device.deviceName
            $DiskVersions = Get-PvsDiskVersion -diskLocatorName $Disk.diskLocatorName -storeName $disk.storeName -siteName $disk.siteName
            $Tasks = Get-PvsUpdateTask -siteName $PVSSite.SiteName
            $ManagedvDisks = Get-PvsdiskUpdateDevice -siteName $PVSSite.SiteName
            $PersonalityStrings = Get-PvsDevicePersonality -Object $ManagedvDisk
            $Device = Get-PvsDeviceInfo -deviceId $ManagedvDisk.deviceId
            $vDisks = Get-PvsDiskUpdateDevice -deviceId $ManagedvDisk.deviceId
            $Collections = Get-PvsCollection -SiteName $PVSSite.SiteName
            $AuthGroups = Get-PvsAuthGroup -CollectionId $Collection.collectionId
            $ServerStore = Get-PVSServerStore -ServerName $Server.serverName
            #>
        }

        #endregion
    } #Before All Get External Data

    Context 'PVS Farm Configuration'{
        It 'Minimum PVS Version' {
            [version]$pvsVersion.SdkVersion -ge [version]$pvsData.PVSFarmInformation.Version | Should Be $true
        }

        It 'Groups with Farm Administrator access' {
            $pvsAuthGroupFarm.AuthGroupName | Should Be $pvsData.PVSFarmInformation.Security.AuthGroupName
        }

        It 'All the Security Groups that can be assigned access rights' {
            $pvsAuthGroup.AuthGroupName | Should Be $pvsData.PVSFarmInformation.Groups.AuthGroupName
        }

        It 'License Server Name'{
            $pvsFarm.LicenseServer | Should Be $pvsData.PVSFarmInformation.Licensing.LicenseServer
        }

        It 'License Server Port'{
            $pvsFarm.LicenseServerPort | Should Be $pvsData.PVSFarmInformation.Licensing.LicenseServerPort
        }

        It 'Use Datacenter Licenses for desktops if no Desktop License is available'{
            $pvsFarm.LicenseTradeUp | Should Be $pvsData.PVSFarmInformation.Licensing.LicenseTradeUp
        }

        It 'Enable auto-add'{
            $pvsFarm.AutoAddEnabled | Should Be $pvsData.PVSFarmInformation.Options.AutoAddEnabled
        }

        if($pvsFarm.AutoAddEnabled -and $pvsData.PVSFarmInformation.Options.AutoAddEnabled){
            It 'Add new devices to this site' {
                $pvsFarm.DefaultSiteName | Should Be $pvsdata.PVSFarmInformation.Options.DefaultSiteName
            }
        }

        It 'Enable auditing'{
            $pvsFarm.AuditingEnabled | Should Be $pvsData.PVSFarmInformation.Options.AuditingEnabled
        }

        It 'Offline database support'{
            $pvsFarm.OfflineDatabaseSupportEnabled | Should Be $pvsData.PVSFarmInformation.Options.OfflineDatabaseSupportEnabled
        }

        It 'Customer Experience Improvement Program'{
            $pvsCEIPData.Enabled | Should Be $pvsdata.PVSFarmInformation.Options.CeipEnabled
        }

        It 'Alert if number of versions from base image exceeds'{
            $pvsFarm.MaxVersions | Should Be $pvsdata.PVSFarmInformation.vDiskVersion.MaxVersions
        }

        It 'Merge after automated vDisk update, if over alert threshold'{
            $pvsFarm.AutomaticMergeEnabled | Should Be $pvsdata.PVSFarmInformation.vDiskVersion.AutomaticMergeEnabled
        }

        It 'Default access mode for new merge versions'{
            $pvsFarm.MergeMode | Should Be $pvsdata.PVSFarmInformation.vDiskVersion.MergeMode
        }

        It 'Database server'{
            $pvsFarm.DatabaseServerName | Should Be $pvsdata.PVSFarmInformation.status.DatabaseServerName
        }

        It 'Database instance'{
            $pvsFarm.DatabaseInstanceName | Should Be $pvsdata.PVSFarmInformation.status.DatabaseInstanceName
        }

        It 'Database'{
            $pvsFarm.DatabaseName | Should Be $pvsdata.PVSFarmInformation.status.DatabaseName
        }

        It 'Failover Partner Server'{
            $pvsFarm.FailoverPartnerServerName | Should Be $pvsdata.PVSFarmInformation.status.FailoverPartnerServerName
        }

        It 'Failover Partner Instance'{
            $pvsFarm.FailoverPartnerInstanceName | Should Be $pvsdata.PVSFarmInformation.status.FailoverPartnerInstanceName
        }

        It 'Active Directory groups are used for access rights'{
            $pvsFarm.AdGroupsEnabled | Should Be $pvsdata.PVSFarmInformation.Status.AdGroupsEnabled
        }

        it 'My Citrix Username'{
            $pvsCisData.UserName | Should Be $pvsdata.PVSFarmInformation.ProblemReport.UserName
        }
    }
    $pvsSite | ForEach-Object {

        Context "$($_.SiteName) Site Configuration"{

            BeforeAll {
                $currentSite = $_
                $siteData = $pvsData.sites | Where-Object {$_.SiteName -eq $currentSite.SiteName}
            }

            It 'Administrator credentials used for Multiple Activation Key enabled devices'{
                $currentSite.MakUser | Should Be $siteData.Properties.MAK.MakUser
            }

            If($pvsFarm.AutoAddEnabled -and $pvsData.PVSFarmInformation.Options.AutoAddEnabled){
                It 'Auto-Add - Add new devices to this collection'{
                    $currentSite.DefaultCollectionName | Should Be $siteData.Properties.Options.DefaultCollectionName
                }
            }

            It 'Seconds between vDisk Inventory Scans'{
                $currentSite.InventoryFilePollingInterval | Should Be $siteData.Properties.Options.InventoryFilePollingInterval
            }

            It 'Enable automatic vDisk updates on this site'{
                $currentSite.EnableDiskUpdate | Should Be $siteData.Properties.vDiskUpdate.EnableDiskUpdate
            }

            If($currentSite.EnableDiskUpdate -and $siteData.Properties.vDiskUpdate.EnableDiskUpdate){
                It 'Server to run vDisk Updates for this site'{
                    $currentSite.DiskUpdateServerName | Should Be $siteData.Properties.vDiskUpdate.DiskUpdateServerName
                }
            }

            $pvsServer | Where-Object {$_.SiteName -eq $currentSite.SiteName} | ForEach-Object {
                $currentServer = $_
                $serverData = $SiteData.Servers | Where-Object {$_.Name -eq $currentServer.Name}

                If($local){
                    $pvsBootstrapOptions = Import-Clixml (Join-Path $path xml\BootstrapOptions.xml)
                }
                else{
                    $pvsBootstrapName = $currentServer | Get-PvsServerBootstrapName | Select-Object -First 1
                    $pvsBootstrapOptions = $currentServer | Get-PvsServerBootstrap -Name $pvsBootstrapName.Name
                }
                It "$($currentServer.Name) Power Rating"{
                    $currentServer.PowerRating | Should Be $serverData.General.PowerRating
                }

                It "$($currentServer.Name) Log Events to the servers Windows Event Log"{
                    $currentServer.EventLoggingEnabled | Should Be $serverData.General.EventLoggingEnabled
                }

                It "$($currentServer.Name) First Port to use for Server Communication"{
                    $currentServer.FirstPort | Should Be $serverData.Network.FirstPort
                }

                It "$($currentServer.Name) Last Port to use for Server Communication"{
                    $currentServer.LastPort | Should Be $serverData.Network.LastPort
                }

                It "$($currentServer.Name) Automate computer account password updates"{
                    $currentServer.AdMaxPasswordAgeEnabled | Should Be $serverData.Options.AdMaxPasswordAgeEnabled
                }

                If($currentServer.AdMaxPasswordAgeEnabled -and $serverData.Options.AdMaxPasswordAgeEnabled){
                    It "$($currentServer.Name) Days between password updates"{
                       $currentServer.AdMaxPasswordAge | Should Be $serverData.Options.AdMaxPasswordAge
                    }
                }
                It "$($currentServer.Name) Threads per port"{
                    $currentServer.ThreadsPerPort | Should Be $serverData.Advanced.Server.ThreadsPerPort
                }
                It "$($currentServer.Name) Buffers per thread"{
                    $currentServer.BuffersPerThread | Should Be $serverData.Advanced.Server.BuffersPerThread
                }
                It "$($currentServer.Name) Server cache timeout"{
                    $currentServer.ServerCacheTimeout | Should Be $serverData.Advanced.Server.ServerCacheTimeout
                }
                It "$($currentServer.Name) Local concurrent I`/O limit"{
                    $currentServer.LocalConcurrentIoLimit | Should Be $serverData.Advanced.Server.LocalConcurrentIoLimit
                }
                It "$($currentServer.Name) Remote concurrent I`/O limit"{
                    $currentServer.RemoteConcurrentIoLimit | Should Be $serverData.Advanced.Server.RemoteConcurrentIoLimit
                }
                It "$($currentServer.Name) Ethernet maximum transmission unit `(MTU`)"{
                    $currentServer.MaxTransmissionUnits | Should Be $serverData.Advanced.Network.MaxTransmissionUnits
                }
                It "$($currentServer.Name) I`/O burst size"{
                    $currentServer.IoBurstSize | Should Be $serverData.Advanced.Network.IoBurstSize
                }
                It "$($currentServer.Name) Enable non-blocking I`/O for network communications"{
                    $currentServer.NonBlockingIoEnabled | Should Be $serverData.Advanced.Network.NonBlockingIoEnabled
                }
                It "$($currentServer.Name) Boot pause"{
                    $currentServer.BootPauseSeconds | Should Be $serverData.Advanced.Pacing.BootPauseSeconds
                }
                It "$($currentServer.Name) Maximum boot time"{
                    $currentServer.MaxBootSeconds | Should Be $serverData.Advanced.Pacing.MaxBootSeconds
                }
                It "$($currentServer.Name) Maximum devices booting"{
                    $currentServer.MaxBootDevicesAllowed | Should Be $serverData.Advanced.Pacing.MaxBootDevicesAllowed
                }
                It "$($currentServer.Name) vDisk Creation pacing"{
                    $currentServer.VDiskCreatePacing | Should Be $serverData.Advanced.Pacing.VDiskCreatePacing
                }
                It "$($currentServer.Name) License timeout"{
                    $currentServer.LicenseTimeout | Should Be $serverData.Advanced.Device.LicenseTimeout
                }
                It "$($currentServer.Name) Bootstrap Verbose Mode"{
                    $pvsBootstrapOptions.VerboseMode | Should be $serverData.ConfigureBootstrap.Options.VerboseMode
                }
                It "$($currentServer.Name) Interrupt Safe Mode"{
                    $pvsBootstrapOptions.InterruptSafeMode | Should be $serverData.ConfigureBootstrap.Options.InterruptSafeMode
                }
                It "$($currentServer.Name) Advanced Memory Support"{
                    $pvsBootstrapOptions.PaeMode | Should be $serverData.ConfigureBootstrap.Options.PaeMode
                }
                It "$($currentServer.Name) Network Recovery Method"{
                    $pvsBootstrapOptions.BootFromHdOnFail | Should be $serverData.ConfigureBootstrap.Options.BootFromHdOnFail
                }
                If($pvsBootstrapOptions.BootFromHdOnFail -and $serverData.ConfigureBootstrap.Options.BootFromHdOnFail){
                    It "$($currentServer.Name) Reboot to Hard Drive after"{
                    $pvsBootstrapOptions.RecoveryTime | Should be $serverData.ConfigureBootstrap.Options.RecoveryTime
                    }
                }
                It "$($currentServer.Name) Login Polling Timeout"{
                    $pvsBootstrapOptions.PollingTimeout | Should be $serverData.ConfigureBootstrap.Options.PollingTimeout
                }
                It "$($currentServer.Name) Login General Timeout"{
                    $pvsBootstrapOptions.GeneralTimeout | Should be $serverData.ConfigureBootstrap.Options.GeneralTimeout
                }
            } #ForEach-Object

            #region Site vDisk Pool
            $pvsDiskInfo | Where-Object {$_.SiteName -eq $currentSite.SiteName} | ForEach-Object {
                $currentDisk = $_
                $diskData = $SiteData.vDiskPool | Where-Object {$_.Name -eq $currentDisk.Name}

                It "$($currentDisk.Name) Access Mode and Cache Type" {
                    $currentDisk.WriteCacheType | Should Be $diskData.Properties.General.WriteCacheType
                }
                It "$($currentDisk.Name) RAM Cache Size" {
                    $currentDisk.WriteCacheSize | Should Be $diskData.Properties.General.WriteCacheSize
                }
                It "$($currentDisk.Name) Enable AD machine account password management" {
                    $currentDisk.AdPasswordEnabled | Should Be $diskData.Properties.General.AdPasswordEnabled
                }
                It "$($currentDisk.Name) Enable printer management" {
                    $currentDisk.PrinterManagementEnabled | Should Be $diskData.Properties.General.PrinterManagementEnabled
                }
                It "$($currentDisk.Name) Enable streaming of this vDisk" {
                    $currentDisk.Enabled | Should Be $diskData.Properties.General.Enabled
                }
                It "$($currentDisk.Name) Cached secrets cleanup disabled" {
                    $currentDisk.ClearCacheDisabled | Should Be $diskData.Properties.General.ClearCacheDisabled
                }
                It "$($currentDisk.Name) Microsoft license type" {
                    $currentDisk.LicenseMode | Should Be $diskData.Properties.MicrosoftVolumeLicensing.LicenseMode
                }
                It "$($currentDisk.Name) Enable automatic updates for the vDisk" {
                    $currentDisk.AutoUpdateEnabled | Should Be $diskData.Properties.AutoUpdate.AutoUpdateEnabled
                }
                It "$($currentDisk.Name) Enable vDisk Load Balancing" {
                    $currentDisk.RebalanceEnabled | Should Be $diskData.LoadBalancing.RebalanceEnabled
                }
            } #Foreach-Object vDisk
            #endregion

            #region vDisk Update Management
            $pvsDiskUpdateDevice | Where-Object {$_.SiteName -eq $currentSite.SiteName} | ForEach-Object {
                $currentDiskUpdateDevice = $_
                #$updateDeviceData =
            }

            #endregion
        } #Context Site Configuration
    } #Site ForEach-Object
} #Describe PVS
