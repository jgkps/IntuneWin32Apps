# first we install the modules we need

Install-Module -Name "IntuneWin32App" -force
Install-Module -Name "vcredist" -force

# change contoso to whatever your tenant domain is
# could be contoso.com or contoso.onmicrosoft.com
Connect-MSIntuneGraph -TenantID contoso.com

# download the vc redist installers
new-item -Path c:\vcredist\source -ItemType Directory -Force
$metadata = Save-VcRedist (Get-VcList) -Path c:\vcredist\source



foreach ($file in $metadata) {


    # Package the source files as .intunewin file
    $SourceFolder = Get-Item $file.path | Select-Object -ExpandProperty directoryname
    $SetupFile = (Get-Item $file.path) | Select-Object -ExpandProperty name
    $OutputFolder = "C:\vcredist\Output\" + $file.name + '\' + $file.architecture

    if ( (test-path -Path $OutputFolder) -eq $false ) {
        new-item -Path $OutputFolder -ItemType Directory -Force
    }

    $Win32AppPackage = New-IntuneWin32AppPackage -SourceFolder $SourceFolder -SetupFile $SetupFile -OutputFolder $OutputFolder -Verbose

    # Get MSI meta data from .intunewin file
    $IntuneWinFile = $Win32AppPackage.Path

    # Create custom display name like 'Name' and 'Version'
    $DisplayName = $file.name + ' ' + $file.architecture
    $Publisher = "Microsoft"
    $Description = $file.URL

    # not that anyone is still using 32bit windows, but its good house keeping
    if ( $file.architecture -eq "x64" ) {
        $requirementrule = New-IntuneWin32AppRequirementRule -Architecture x64 -MinimumSupportedOperatingSystem 1607
    }
    else {
        $requirementrule = New-IntuneWin32AppRequirementRule -Architecture all -MinimumSupportedOperatingSystem 1607
    }
    

    # Detection
    # split to get rid of the switches, and remove the "", and finally split-path again into two variables
    $detectionpaths = ($silentuninstall.split('/')[0]).replace('"', '') 
    $dirpath = $detectionpaths | Split-Path
    $filepath = $detectionpaths | Split-Path -Leaf

    $DetectionRule = New-IntuneWin32AppDetectionRuleFile -Path $dirpath -DetectionType exists -FileOrFolder $filepath -Existence

    # ho ho ho, now i have a commandline!
    $installcommandline = $SetupFile + ' ' + $($file.install)


    # Add new MSI Win32 app
    # Splatting is very pretty
    $IntuneAppParameters = @{
        FilePath             = $IntuneWinFile 
        DisplayName          = $DisplayName
        RequirementRule      = $requirementrule
        Description          = $Description
        AppVersion           = $file.version
        Publisher            = $Publisher
        InstallCommandLine   = $installcommandline
        UninstallCommandLine = $file.silentuninstall
        InstallExperience    = "system"
        RestartBehavior      = "suppress"
        DetectionRule        = $DetectionRule
    }
   Add-IntuneWin32App @IntuneAppParameters
   

    # needs a slight delay, otherwise the intunewinfile upload might fail
    start-sleep -Seconds 10
}