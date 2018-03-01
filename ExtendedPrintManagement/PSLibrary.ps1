function Get-PrinterName {
    Param($Computer)
    if ($Computer -eq $null) {
        Get-WmiObject win32_printer | Where-Object -Property Local -EQ "True" | Select-Object -ExpandProperty Name
    }
    else {
        Get-WmiObject win32_printer -ComputerName $Computer | Where-Object -Property Local -EQ "True" | Select-Object -ExpandProperty Name
    }
}

function Get-PrinterObject {
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        $Printer,
    
        [Parameter(Mandatory = $false, Position = 1)]
        $Computer)

    if ($Computer -eq $null) {
        Get-WmiObject -Class win32_printer -Filter "DeviceID = '$Printer'"
    }
    else {
        Get-WmiObject -Class win32_printer -Filter "DeviceID='$Printer'" -ComputerName $Computer
    }
}

function Test-Con {
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        $Computer)

    Test-Connection -ComputerName $Computer -Count 2 -Quiet
}

function Write-Message {
    Param(
        [Parameter(Mandatory = $false, Position = 0)]
        [string]
        $Printer,
    
        [Parameter(Mandatory = $false, Position = 1)]
        [string]
        $Computer)
	
    if (($Printer -ne "") -and ($Computer -eq "")) {
        Write-Output ("$Printer printer setup completed!")
    }
    if (($Printer -ne "") -and ($Computer -ne "")) {
        Write-Output ("$Printer printer setup on computer $Computer is completed!")
    }
    if (($Computer -ne "") -and ($Printer -eq "")) {
        Write-Output ("Computer $Computer is not available.")
    }
}

function Copy-Settings {
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        $SetPrinter,
        [Parameter(Mandatory = $false, Position = 1)]
        $SourcePrinter)

    $SPrinterObject = Get-PrinterObject -Printer $SourcePrinter
    if ($SPrinterObject.StartTime -eq $SPrinterObject.UntilTime) {
        $SPrinterObject.StartTime = "********000000.000000+000"
        $SPrinterObject.UntilTime = "********000000.000000+000"
    }
    $Arguments = @{
        StartTime        = $SPrinterObject.StartTime;
        UntilTime        = $SPrinterObject.UntilTime;
        DoCompleteFirst  = $SPrinterObject.DoCompleteFirst;
        Direct           = $SPrinterObject.Direct;
        Queued           = $SPrinterObject.Queued;
        KeepPrintedJobs  = $SPrinterObject.KeepPrintedJobs;
        RawOnly          = $SPrinterObject.RawOnly;
        EnableBIDI       = $SPrinterObject.EnableBIDI;
        Priority         = $SPrinterObject.Priority;
        SpoolEnabled     = $SPrinterObject.SpoolEnabled;
        PrintJobDataType = $SPrinterObject.PrintJobDataType;
        Hidden           = $SPrinterObject.Hidden
    }
    Set-WmiInstance -Path "$SetPrinter" -Arguments $Arguments
}

function Get-OSArchitecture {
    Get-WmiObject -Class Win32_OperatingSystem |Select-Object -ExpandProperty OSArchitecture
}

function Get-OSbuild {
    Get-WmiObject -Class Win32_OperatingSystem |Select-Object -ExpandProperty BuildNumber
}

function Get-PrintDriverName {
    Param(
        [Parameter(Mandatory = $false, Position = 0)]
        $ComputerName)

    if ($ComputerName -eq $null) {
        (Get-PrinterDriver).name
    }
    else {
        (Get-PrinterDriver -ComputerName $ComputerName).name
    }
}