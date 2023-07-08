# Copyright 2003 - TMurgent Technologies, LLP.
# You are hereby granted the right to do with this as you will.
# No support or liability is provided - use at your own risk!
#
# Script requires installation of PassiveInstall, a free powershell module available from TMurgent Technologies
#     https://www.tmurgent.com/APPV/en/resources/tools-downloads/tools-packaging/117-tools/packaging-tools/435-passiveinstall
#
#----------------------- Modifications should only be needed between these lines ---------------------

# AppvContentStore: Folder containing the App-V packages to be processed.
$AppvContentStore = "\\YourServer\Share"                    # Location of existing App-V files (presumably under subfolders of this folder)

# ConversionOutputFolder: Folder to contain the outputs.  Two or three subfolders will be created under it,
#                         one for logs, one for converted packages, and one for fixed packages. 
$ConversionOutputFolder = "\\nuc1\Packages\Conversions"

# Signing Certificate, used only if signing of unfixed packages is requested
$PublisherName = "CN=YourCompanyName"                       # Matching Subject line of your certificate
$PublisherDisplayName = "Converted by TMurgent Technologies, LLP"  # Whatever you'd like.
$Certificate = "C:\YourPath\YourCert.pfx"                          # Path to your code signing certificate
$CertificatePassword = "YourPassword"                              # Passwword for signing
$TimestampingURL = "http://timestamp.digicert.com"
$SignToolPath = "C:\YourPath\signtool.exe"                         # Supply from Windows SDK for example


#--------------- Temporary controls when multiple runs are of interest
# SkipAllConversions: Normally false, set to true if you want only to sign or fix packages
# SkipConversionsUntil: Normally blank, the number of App-V packages to skip.  Used if you have to restart conversion process part way through.
# ProcessOnlyPackage: Normally blank, or the name (not path) of a package file (without appv extension) if you only want to reconvert one package.
$SkipAllConversions = $false #$false
$SkipConversionsUntil = "" #""
$ProcessOnlyPackage = "" #""  Fill in the name of a specific package name if you only want to process that package
# SkipAllSigning: Normally false, set to true if you want to skip signing the converted (raw) packages
$SkipAllSigning = $false#$false
# SkipAllFixing: Normally false, set to true if you want to skip fixing with TMEditX
# AppConversionParameters: string for the parameters to use when fixing
$SkipAllFixing = $false #false
$MsixFixupParameters = "/ApplyAllFixes /UseDebugPsf /AutoReplaceFrfWithMfr /UseRegLeg /AutoSaveAsMsix /AutoSaveAsFolder"


#----------------------- Modifications should only be needed above this line ---------------------


Import-Module "C:\Program Files\WindowsPowerShell\Modules\PassiveInstall\PassiveInstall.dll"
Approve-PassiveElevation -AsAdmin

$executingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

. $PSScriptRoot\batch_convert.ps1

# Creating a folder to store the template files used for the conversion
$MPTtemplateLocation = ([System.IO.Path]::Combine($ConversionOutputFolder, "MPT_Templates"))
New-Item -Force -Type Directory ($MPTtemplateLocation)

# Creating a folder to store the MSIX packages
$SaveLocation = ([System.IO.Path]::Combine($ConversionOutputFolder, "MSIX_MMPT"))
New-Item -Force -Type Directory ($SaveLocation)

# Creating a folder to store the MSIX packages
$FixedLocation = ([System.IO.Path]::Combine($ConversionOutputFolder, "MSIX_Fixed"))
New-Item -Force -Type Directory ($SaveLocation)

# Define and clear out any existing log file
$ConversionMasterLogFile = ([System.IO.Path]::Combine($ConversionOutputFolder, "Log.txt"))
if ( (test-path "$($ConversionMasterLogFile)" ) -eq $true) {
    Remove-PassiveFiles $ConversionMasterLogFile
}
Write-Host  -ForegroundColor Cyan "Logging of this session written to $($ConversionMasterLogFile)"

function Run_MainProcessing($AppVFilePath)
{
    $Installerpath = $AppVFilePath.FullName
    $filename = $AppVFilePath.BaseName
    write-Output ""
    write-Host -ForegroundColor Cyan "starting the conversion of: " $Installerpath
    
    # MSIX package name cannot contain spaces, dashes or dots, so replacing these
    $packageStrippedName = $filename -replace '\s+', ''  -replace '_', '-'
    $job = "job" + $counter
    
    # get the contents of the template XML
    [String]$newXml = Get-Content -path $executingScriptDirectory\MsixPackagingToolTemplate.xml | Out-String
    # Replace the placeholders with the correct values
    $newXml = $newXml.Replace("[Installer]", "$Installerpath")
    $newXml = $newXml.Replace("[SaveLocation]", "$SaveLocation")
    $newXml = $newXml.Replace("[PackageName]", "$packageStrippedName")
    $newXml = $newXml.Replace("[PackageDisplayName]", "$filename")
    $newXml = $newXml.Replace("[PublisherName]", "$PublisherName")
    $newXml = $newXml.Replace("[PublisherDisplayName]", "$PublisherDisplayName")
        
    # saving the newly created template
    $JobTemplate = "$($MPTtemplateLocation)\MsixPackagingToolTemplate_$($job).xml"
    if ( (test-path "$($JobTemplate)" ) -eq $true) {
        Remove-PassiveFiles "$($JobTemplate)"
    }
    $newXml | out-File "$($JobTemplate)" -Encoding Ascii -Force
     
    # Starting the conversion  NOTE: Packaging tool does not accept putting template filepath in quotes -- so no spaces allowed!!!
    MsixPackagingTool.exe create-package --template $JobTemplate
  
    $counter = $counter + 1
    write-Output (Get-Date)
}

# get all the App-V packages from the ContentStore and convert them.
$counter = 1
write-host (Get-Date)
if ($SkipAllConversions -eq $false)
{
    $GotToSkip = $false
    if ($SkipConversionsUntil -eq "")
    {
        $GotToSkip = $true
        Write-Host -ForegroundColor Cyan "Starting all conversions"
    }
    else
    {
        Write-Host -ForegroundColor Cyan "Skipping conversions until $($SkipConversionsUntil)"
    }
    get-childitem $AppvContentStore -recurse | Where-Object { $_.extension -eq ".appv" } | ForEach-Object {
        if ($GotToSkip -eq $false)
        {
            if ($_.FullName -eq $SkipConversionsUntil)
            {
                $GotToSkip = $true
                Write-Host -ForegroundColor Cyan "Start remaining conversions"
            }
        }
        if ($GotToSkip -eq $true)
        {
            $DoProcess = $true 
            Write-Host -ForegroundColor Gray $_.BaseName
            
            Write-Host -ForegroundColor Gray "$($ProcessOnlyPackage)"
            if ($ProcessOnlyPackage -ne "")
            {
                if ($_.BaseName -notlike "$($ProcessOnlyPackage)")
                {
                    $DoProcess = $false
                }
            }
            if ($DoProcess)
            {
                Write-Host -ForegroundColor Gray "Converting $($_.FullName)" 
                $err = Run_MainProcessing $_ *>&1
                Write-Output $err >> "$($ConversionMasterLogFile)"
                Write-Output $err
            }
            else
            {
                Write-Host -ForegroundColor Gray "Skipping conversion $($_.FullName)" 
            }
        }
        else
        {
            Write-Host -ForegroundColor Gray "Skipping conversion $($_.FullName)" 
        }   
    }
}
else
{
    Write-Host -ForegroundColor Cyan "Skipping all conversions"
}

# App-V packages converted to MSIX. Signing the new MSIX packages
if ($SkipAllSigning -eq $false)
{
    Write-Host "Sign packages..." -ForegroundColor Cyan
    Get-ChildItem $SaveLocation | foreach-object {
        $MSIXpackage = $_.FullName
        #$cmd = "`"" + $SignToolPath + "\signtool.exe`""
        #$args = "sign /a /v  /f " + $Certificate + " /p `"" + $CertificatePassword + "`" /fd SHA256 /tr " + $TimestampingURL + " `"" + $MSIXpackage + "`""
        #Start-Process $cmd -Wait -ArgumentList  $args 

        Write-Output "Signing $($MSIXpackage)"
        & $SignToolPath sign /f $Certificate + " /p " + $CertificatePassword + " /fd SHA256 /t " + $TimestampingURL + " `"" + $MSIXpackage + "`""
    }
}
else
{
    Write-Host -ForegroundColor Cyan "Skipping all signings"
}

# cleanup
Write-host "Cleaning up packaging tool; ignore any errors..."
MsixPackagingTool.exe cleanup

if ($SkipAllFixing -eq $false)
{
    Write-Host -ForegroundColor Cyan "AutoFix packages..." 
    
    $Toolpath = 'TMEditX.exe'
    $Processed = 0
    $Attempted = 0

    Get-ChildItem $SaveLocation | foreach-object { 
        if ($_.Extension -eq '.msix')
        {
            $msixPath = $_.FullName
            #$msixOutPath = [System.IO.Path]::Combine($FixedLocation, $_.Name)
        
            $Attempted += 1
                                
            $arglist = "$($_.FullName) $($MsixFixupParameters) $($FixedLocation)"
            Write-Host "Running: $($Toolpath) $($arglist)" -ForegroundColor Cyan
            Start-Process -Wait -FilePath $Toolpath -ArgumentList $arglist
            $Processed += 1
        }
    }
    $countPackagesFix = (get-item "$($FixedLocation)\*.msix").Count
    Write-Host "$($Processed) of $($Attempted) packages fixed; total $($countPackagesFix) fixed packages in output folder." -ForegroundColor Green
}
else
{
    Write-Host -ForegroundColor Cyan "Skipping all fixups"
}

write-Host -ForegroundColor "Green"  "Done."
Show-PassiveTimer 20000 "End of script, Ctrl-C to end now or wait for timer. Left-Click on window to pause."