<#
.SYNOPSIS
 Increments the GPO version number causing the policy to refresh.

.DESCRIPTION
 Increments the GPO version number causing the policy to refresh. You can
 increment either User, Computer, or Both version numbers. Handles updating
 AD and the GPT.INI file. If the file or AD update fails the function should
 revert the version number. The function should take pipeline input from
 Get-GPO.

.PARAMETER Guid
 The GPO guid to increment

.PARAMETER Name
 The GPO name to increment

.PARAMETER Increment
 Choose to increment User, Computer, or Both

.PARAMETER Domain
 See Get-GPO

.PARAMETER Server
 See Get-GPO

.EXAMPLE
 Get-GPO -Name 'Default Domain Policy' | Invoke-GPOIncrementVersion -Increment Computer
 
 Increments the computer policy version for the Default Domain Policy
#>
function Invoke-GPOIncrementVersion {

    [CmdletBinding(
        DefaultParameterSetName = 'Guid',
        SupportsShouldProcess,
        ConfirmImpact='Low'
    )]
    param(

        [Parameter(
            ParameterSetName = 'Guid',
            Mandatory,
            ValueFromPipelineByPropertyName,
            Position = 1
        )]
        [Alias( 'Id' )]
        [guid[]]
        $Guid,

        [Parameter(
            ParameterSetName = 'Name',
            Mandatory
        )]
        [string[]]
        $Name,

        [Parameter(
            Mandatory
        )]
        [ValidateSet( 'Both', 'Computer', 'User' )]
        [string]
        $Increment,

        [string]
        $Domain,

        [string]
        $Server

    )

    process {

        $PSBoundParameters.Remove( $PSCmdlet.ParameterSetName ) > $null
        
        if ( $PSBoundParameters.ContainsKey( 'Increment' ) ) { $PSBoundParameters.Remove( 'Increment' ) > $null }

        $GPOs = switch ( $PSCmdlet.ParameterSetName ) {

            'Guid' { $Guid | ForEach-Object { Get-GPO -Guid $_ @PSBoundParameters } }

            'Name' { $Name | ForEach-Object { Get-GPO -Name $_ @PSBoundParameters } }

        }

        $GPOs | ForEach-Object {

            # find the GPO in AD
            $GetGpoObjectSplat = @{
                SearchBase = "CN=Policies,CN=System,$( $_.DomainName.Split('.').ForEach({"DC=$_"}) -join ',' )"
                Filter     = "Name -eq '$( $_.Id.ToString('B') )' -and objectClass -eq 'groupPolicyContainer'"
                Server     = $_.DomainName
                Properties = 'DisplayName', 'versionNumber', 'gPCFileSysPath'
            }
            $GpoObject = Get-ADObject @GetGpoObjectSplat

            # first calculate the existing version numbers
            # User Policy version is the left 16 bits
            # Computer Policy version is the right 16 bits
            $UserPolicyVersion = $GpoObject.versionNumber -shr 16
            $ComputerPolicyVersion = $GpoObject.versionNumber -band 0xffff

            # increment only the requested version
            if ( $Increment -eq 'Both' -or $Increment -eq 'User'     ) { $UserPolicyVersion += 1     }
            if ( $Increment -eq 'Both' -or $Increment -eq 'Computer' ) { $ComputerPolicyVersion += 1 }

            # move the user policy back to the left, then add the computer policy
            $NewVersion = ( $UserPolicyVersion -shl 16 ) + $ComputerPolicyVersion

            # actually apply the change
            if ( $PSCmdlet.ShouldProcess( "$($GpoObject.DisplayName) $($GpoObject.Name)", "increment the version number for $($Increment.ToLower())" ) ) {

                $ConfirmPreference = 'None'

                $GptIniPath = Join-Path $GpoObject.gPCFileSysPath 'GPT.INI'
                $BkpIniPath = Join-Path $GpoObject.gPCFileSysPath 'GPT.INI-BAK'

                if ( -not( Test-Path -Path $GptIniPath -PathType Leaf ) ) {
                    
                    Write-Error "Could not find GPO $($GpoObject.DisplayName) $($GpoObject.Name) in file system!"
                    return
                    
                }

                try {
                    
                    # make a backup
                    Copy-Item -Path $GptIniPath -Destination $BkpIniPath -Force -ErrorAction Stop

                    # update the version
                    $IniContent = Get-Content -Path $GptIniPath
                    $IniContent = $IniContent -replace 'Version=\d+', "Version=$NewVersion"
                    $IniContent | Set-Content -Path $GptIniPath -Encoding Ascii -ErrorAction Stop

                } catch {
                
                    Remove-Item -Path $BkpIniPath -ErrorAction SilentlyContinue
                    
                    Write-Error "Could not modify the GPO version for $($GpoObject.DisplayName) $($GpoObject.Name) in file system!"
                    return
                    
                }

                try {

                    $GpoObject | Set-ADObject -Replace @{ versionNumber = $NewVersion } -ErrorAction Stop

                } catch {

                    Copy-Item -Path $BkpIniPath -Destination $GptIniPath -Force | Out-Null

                    Write-Error "Could not modify the GPO version for $($GpoObject.DisplayName) $($GpoObject.Name) in AD!"
                    return

                } finally {

                    Remove-Item -Path $BkpIniPath -ErrorAction SilentlyContinue
                        
                }

            }

        }

    }

}


<#
.SYNOPSIS
 Get GPO links in a domain

.DESCRIPTION
 Bastardizes Set-GPLink to return Microsoft.GroupPolicy.GPLink objects for GPLinks in a domain.

.EXAMPLE
 Get-GPLink -Name 'Default Domain Policy'

.EXAMPLE
 Get-GPLink -Guid 31b2f340-016d-11d2-945f-00c04fb984f9

.EXAMPLE
 Get-GPO -Name 'Default Domain Policy' | Get-GPLink -Scope EntireForest

#>
function Get-GPLink {

    [CmdletBinding(
        PositionalBinding = $false
    )]
    param(
    
        [Parameter(
            ParameterSetName = 'Name',
            Mandatory
        )]
        [string]
        $Name,

        [Parameter(
            ParameterSetName = 'Guid',
            ValueFromPipelineByPropertyName,
            Mandatory
        )]
        [Alias( 'Id' )]
        [guid]
        $Guid,

        [string]
        $Target,

        [Parameter(
            ValueFromPipelineByPropertyName
        )]
        [Alias( 'DomainName' )]
        [string]
        $Domain,

        [string]
        $Server,

        [ValidateSet( 'Domain', 'AllSites', 'EntireForest' )]
        [string]
        $Scope = 'Domain'
    
    )

    $RealScope = $Scope

    # force single scope if target is specified
    if ( $Target ) { $RealScope = 'Single' }

    $PSBoundParameters.Remove( 'Target' ) > $null
    $PSBoundParameters.Remove( 'Scope' ) > $null
    $PSBoundParameters.Remove( 'Verbose' ) > $null

    $GPO = Get-GPO @PSBoundParameters

    if ( -not $GPO ) { return }

    Write-Verbose "Found GPO: $($GPO.DisplayName)"

    $PSBoundParameters.Remove( 'Name' ) > $null
    $PSBoundParameters.Guid = $GPO.Id
    $PSBoundParameters.Domain = $GPO.DomainName

    if ( $Server ) {

        $DomainDC = Get-ADDomainController -Identity $Server -Server $Server

    } else {

        Write-Verbose "Searching for domain controller for domain $Domain..."

        $DomainDC = Get-ADDomainController -DomainName $PSBoundParameters.Domain -Discover

        $Server = $DomainDC | Select-Object -ExpandProperty HostName -First 1

    }

    if ( -not $DomainDC ) { return }

    if ( $Scope -eq 'Domain' -or $DomainDC.IsGlobalCatalog ) {
    
        Write-Verbose "Using domain controller $Server."

    } else {
    
        $Server = Get-ADDomainController -DomainName $DomainDC.Forest -Discover -Service GlobalCatalog |
            Select-Object -ExpandProperty HostName -First 1

        Write-Warning "Selected server is not a global catalog, using $Server instead."

    }

    $SearchSplat = @{
        Filter      = "gplink -like '*{0}*'" -f ([guid]$GPO.Id).ToString('B')
        Properties  = 'DistinguishedName'
        Server      = $Server + ':3268'
    }

    switch ( $RealScope ) {

        'Single' {

            $SearchSplat.SearchBase  = $Target
            $SearchSplat.SearchScope = 'Base'
            
        }
        
        'Domain' {

            $SearchSplat.SearchBase  = $Domain.Split('.').ForEach({ "DC=$_" }) -join ','
            $SearchSplat.Server      = $Server

        }

        'AllSites' {

            $SearchSplat.SearchBase  = "CN=Sites," + (Get-ADRootDSE -Server $Server).configurationNamingContext
            $SearchSplat.SearchScope = 'OneLevel'
    
        }
        
        'EntireForest' {

            $SearchSplat.SearchBase  = (Get-ADRootDSE -Server $Server).rootDomainNamingContext

        }

    }

    $SearchSplat.Keys | ForEach-Object { Write-Verbose ( '{0,-14} : {1}' -f $_, $SearchSplat[$_] ) }
    
    Get-ADObject @SearchSplat |
        ForEach-Object { Set-GPLink -Target $_.DistinguishedName @PSBoundParameters }

}


<#
.SYNOPSIS
 Import Group Policy backups manifest

.DESCRIPTION
 Import Group Policy backups manifest into PSCustomObjects

.PARAMETER BackupPath
 Path to a Group Policy backup repository

.EXAMPLE
 Get-Item ~\PolicyBackups | Import-GPBackupManifest
 
#>
function Import-GPBackupManifest {

    param(

        [Parameter(
            Mandatory,
            ValueFromPipelineByPropertyName,
            Position = 1
        )]
        [Alias(
            'FullName',
            'Path'
        )]
        [string]
        $BackupPath

    )

    # user provided path to backup manifest, not backup folder
    if ( [System.IO.Path]::GetFileName( $BackupPath ) -eq 'manifest.xml' ) {
        
        $BackupPath = [System.IO.Path]::GetDirectoryName( $BackupPath )

    }

    # build the manifest path
    $ManifestPath = Join-Path $BackupPath 'manifest.xml'

    # check for a manifest
    if ( -not( Test-Path $ManifestPath ) ) {

        Write-Error 'Path must point to a Group Policy backup repository.'
        return

    }

    # load the manifest xml
    [xml]$Manifest = Get-Content -Path $ManifestPath

    # build PSCustomObjects from the xml
    foreach ( $BackupInst in $Manifest.Backups.BackupInst ) {

        $Object = @{}

        ( $BackupInst | Get-Member -MemberType Property ).Name.ForEach({

            $Object[$_] = $BackupInst.($_).'#cdata-section'

        })

        # include the path to the backup in the object
        $Object['BackupPath'] = Join-Path $BackupPath $Object.ID

        [pscustomobject]$Object

    }

}
