# Author::    Jean Duprat (mailto:jean.duprat@thomsonreuters.com)

$osbitness = (Get-WmiObject Win32_OperatingSystem -computername .).OSArchitecture

#Works only on .Net >= 3
function NetFrameworkVersion()
{
    $versions = @("v4\Full","v4.0\Client","v3.5","v3.0")

    $version="Unknown"

    foreach($v in $versions)
    {
        $regpath = "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP"+"\$v"
        if(Test-Path $regpath)
        {
            return (Get-ItemProperty -Path $regpath).Version
        }            
    }
            
    return $version
}

#12 14 15 16
function OfficeVersion
{
    $excelreg = "HKLM:\SOFTWARE\Classes\Excel.Application\CurVer"
    return (Get-ItemProperty -Path $excelreg).'(default)'.Split('.')[2]
}

#x86 or x64
function OfficeBitness()
{
    $version = OfficeVersion
    #No 64 bits version for 2007 and lower
    if($version -le 12)
    {
        return "x86"
    }
    else
    {
        $outlookx64 = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Nod\Microsoft\Office\$version.0\Outlook"

        if(Test-Path $outlookx64)
        {
            $outlookBitness = "HKLM:\SOFTWARE\Wow6432Nod\Microsoft\Office\$version.0\Outlook"
        }
        else
        {
            $outlookBitness = "HKLM:\SOFTWARE\Microsoft\Office\$version.0\Outlook"
        }
        return (Get-ItemProperty -Path $outlookBitness).Bitness
    }
}

#Works only on Office >= 2007
function OfficeFullVersion()
{   
    $excelreg = "HKLM:\SOFTWARE\Classes\Excel.Application\CurVer"
    $excelNb = (Get-ItemProperty -Path $excelreg).'(default)'.Split('.')[2]
        
    $msodll = "C:\Program Files\Common Files\microsoft shared\OFFICE$excelNb\mso.dll"
        
    #x64 OS
    if($osbitness -eq "64-bit")
    {      
        #x86 Office  
        if((OfficeBitness) -eq "x86")
        {
            $msodll = "C:\Program Files (x86)\Common Files\microsoft shared\OFFICE$excelNb\mso.dll"
        }
    }

    $msodllVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($msodll).FileVersion

    $splitVersion = $msodllVersion.Split('.')

    $fullVersion = ""

    #2007 ref https://support.microsoft.com/en-us/kb/928116
    #2010 ref https://support.microsoft.com/en-us/kb/2121559
    #2013 ref https://support.microsoft.com/en-us/kb/2951141
    #2016 ref -
    switch($splitVersion[0])
    {
        12 
        { 
            $fullVersion += "Office 2007"
            if($splitVersion[2] -ge "6214" -and $splitVersion[2] -lt "6425")
            {
                $fullVersion += " SP1"
            }
            if($splitVersion[2] -ge "6425" -and $splitVersion[2] -lt "6611")
            {
                $fullVersion += " SP2"
            }
            if($splitVersion[2] -ge "6611")
            {
                $fullVersion += " SP3"
            }
        }
        14
        {
            $fullVersion += "Office 2010"
            if($splitVersion[2] -ge "6024")
            {
                $fullVersion += " SP1"
            }
        }
        15
        {
            $fullVersion += "Office 2013"
             if($splitVersion[2] -ge "4569")
            {
                $fullVersion += " SP1"
            }
        }
        16
        {
            $fullVersion += "Office 2016"
            #no SP for the moment
        }
        default
        {
            $fullVersion = "Unknown"
        }
    }
    return $fullVersion + " $msodllVersion"
}

$computerName = Get-WmiObject -Class Win32_Processor -ComputerName . | Select-Object -Property PSComputerName
$os = (Get-WmiObject -class Win32_OperatingSystem).Caption
$culture = (Get-Culture |Select-Object -Property Name).Name
$netframework = NetFrameworkVersion
$officeVersion = OfficeFullVersion
$officeBitness = OfficeBitness

Write-Host "msoffice_os=$os $osbitness"
Write-Host "msoffice_culture=$culture"
Write-Host "msoffice_dotNET=$netframework"
Write-Host "msoffice_version=$officeVersion $officeBitness"