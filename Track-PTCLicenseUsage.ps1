#Requires -Version 5.0

<#
    .SYNOPSIS
    This script parses the output of PTC Creo 7's PTCStatus.bat to track license usage.

    .DESCRIPTION
    This script will attempt to locate PTCStatus.bat by searching %ProgramFiles%\PTC.  If found, the batch file will
    be ran and the output captured and parsed into Name, Date, and LicensesInUse before being exported to a tab-
    delimited csv in a file named after the license being tracked.

    The location of the exported csv can be specified on the command line by using the -FilePath parameter.

    Default licenses to be tracked are:
    PROE_DesignEss = Design Essentials Domestic
    PROE_DesignEssG = Design Essentials Global
    PROE_DesignAdv = Design Advanced Domestic
    PROE_DesignAdvG = Design Advanced Global
    PROE_DesignAdvP = Design Advanced Plus Domestic
    PROE_DesignAdvPG = Design Advanced Plus Global

    .INPUTS
    This script does not accept piped inputs.

    .OUTPUTS
    This script does not output PSObjects.

    .PARAMETER FilePath
    Specifies the location where the .csv files will be exported to. Defaults to C:\Temp.

    .PARAMETER LicenseName
    Specifies a name or names of ilcenses to track.

    .PARAMETER Delimiter
    Specifies the delimiter to be used during csv creation.  Defaults to tab delimited ("`t").

    .EXAMPLE
    PS> Track-PTCLicenseUsage -FilePath C:\PTC

    Runs Track-PTCLicenseUsage and saves output files to C:\PTC as tab delimited.  If folder path doesn't exist, it
    will be created.

    .EXAMPLE
    PS> Track-PTCLicenseUsage -LicenseName "PROE_DesignEssG","PROE_DesignAdvP"

    Runs Track-PTCLicenseUsage looking for license usage of PROE_DesignEssG and PROE_DesignAdvP.  Output files will be
    saved to C:\Temp as tab delimited.

    .EXAMPLE
    PS> Track-PTCLicenseUsage

    Runs Track-PTCLicenseUsage with all defaults.  Output files will be saved to C:\Temp as tab delimited.
    

    .NOTES
    Author: Charlie Dobson
    Date: 2021/11/4
    Release: 1.1
#>

<# 
    Expected output as tab ("`t") delimited csv:
    [license] `t [date] `t [count]
#>

[CmdletBinding(SupportsShouldProcess=$true)]
Param (
    [Parameter(Mandatory=$false)]
    [alias("Path")]
    [string]$FilePath,
    [Parameter(Mandatory=$false)]
    [alias("License","Name")]
    [string[]]$LicenseName,
    [Parameter(Mandatory=$false)]
    [alias("D","Delim")]
    [ValidateLength(1,1)]
    [char]$Delimiter
)

class LicenseInfo {
    [string]$Name
    [DateTime]$Date
    [int]$LicensesInUse
    [string]$Users

    LicenseInfo() {

    }

    LicenseInfo([string]$Name, [int]$Count) {
        if ($null -ne $Name) {
            $this.Name = $Name
            $this.Date = $(Get-Date)
        }
        if ($null -ne $Count) {
            $this.LicensesInUse = $Count
        }
    }

    LicenseInfo([string]$Name, [int]$Count, [string]$Users) {
        if ($null -ne $Name) {
            $this.Name = $Name
            $this.Date = $(Get-Date)
        }
        if ($null -ne $Count) {
            $this.LicensesInUse = $Count
        }
        $this.Users = $Users
    }
}

# Specify default save location if not given on command line
if (!$FilePath) {
    [string]$FilePath = 'C:\Temp'
}

# Create folder if it doesn't exist
if (!(Test-Path $FilePath -PathType Container)) {
    try {
        New-Item -Path $FilePath -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Output $_
        return
    }
}

# Specify default licenses to track if not given on command line
if (!$LicenseName) {
    [string[]]$LicenseName = "PROE_DesignEssG", "PROE_DesignEss", "PROE_DesignAdvPG", "PROE_DesignAdvG", "PROE_DesignAdvP", "PROE_DesignAdv", `
        "Mecui_Advanced", "Meceng_Advanced", "MECBASICUI_License", "MECBASICENG_LICENSE", "MECLITEUI_License", "MECLITEENG_License", "CREOSIM_STANDARD"
}

# Specify tab as default delimiter if not given on the command line
if (!$Delimiter) {
    [char]$Delimiter = "`t"
}

# Get script location
[string[]]$Script = (Get-Childitem -Path "${Env:ProgramFiles}\PTC" -Recurse ptcstatus.bat -ErrorAction SilentlyContinue).FullName

# Check if script was found
if (!$Script) {
    Write-Output "Unable to locate PTCStatus.bat. Aborting."
    Return
}

# Capture output from PTCStatus.bat
try {
    [System.Collections.Generic.List[PSObject]]$Output = & $Script[0] '/nopause'
}
catch {
    Write-Output $_
}

# Check if anything was captured
if (!$Output) {
    Write-Output "No output from PTCStatus.bat. Aborting."
    Return
}

# Remove preceeding whitespace
$Output = $Output.TrimStart()
# Convert whitespace into csv
$Output = $Output -replace "\s+",","
# Instantiate a list of LicenseInfo objects
[System.Collections.Generic.List[LicenseInfo]]$licenseInfo = New-Object LicenseInfo

# loop through the lines of output and parse the data
for ($i = 0; $i -lt $Output.Count; $i++) {
    # Loop through each license type specified
    foreach ($license in $LicenseName) {
        # Check if the line matches the license we're looking for and that it is followed by digits indicating number of licenses in use
        if ($Output[$i] -match "$license,\d+") {
            # Create string variable to capture the users consuming the license
            [string]$Users = $null
            # Loop through the lines of output again to locate those that contain user@computer
            for ($j = 0; $j -lt $Output.Count; $j++) {
                # if the current line "j"'s license equals the previous "i" loop's license name,
                # we've found a match of users consuming the license we're looking for
                if ($Output[$j].Split(',')[1] -ieq $Output[$i].Split(',')[0]) {
                    $Users += ($Output[$j].Split(',')[0]).Trim('(',')') + " "
                }
            }
            # Trim end of users variable to remove trailing space we added to separate values
            $Users = $Users.TrimEnd()
            # Add to LicenseInfo class List
            $licenseInfo.Add([LicenseInfo]::new($Output[$i].Split(',')[0], $Output[$i].Split(',')[1], $Users))
        }
    }
}

# now let's output our data
try {
    for ($i = 0; $i -lt $licenseInfo.Count; $i++) {
        if ($null -ne $licenseInfo[$i].Name) {
            Export-Csv -Path "$FilePath\$($licenseInfo[$i].Name).csv" -Delimiter $Delimiter -InputObject $licenseInfo[$i] -NoTypeInformation -Append
        }
    }
}
catch {
    Write-Host -ForegroundColor Red $_
}