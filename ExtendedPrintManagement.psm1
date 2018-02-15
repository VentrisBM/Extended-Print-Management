<#
.Synopsis
   Managing printer availability
.DESCRIPTION
   Set-PrinterAvailability cmdlet allows you to set time period in which users can use printer
.EXAMPLE
   Set-PrinterAvailability -PrinterName "Test Printer" -StartTime 0800 -UntilTime 2000 -ShowProgress

   This command sets printer availability on the local computer. Printer will be available from 8AM to 8PM.
   Note that time must be entered in "military" format(ISO 8601 standard).
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-PrinterAvailability -StartTime 1000 -UntilTime 1200 -ShowProgress

   This set of commands sets printer availability on the local computer for all printers whose name starts with "HP".
.EXAMPLE
    Set-PrinterAvailability -PrinterName * -StartTime 0800 -UntilTime 1630 -ComputerName "Remote_Computer" -ShowProgress

    This command sets printer availability for all printers on a remote computer.
    You can specify hostname or IP address of remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
   Time must be entered in ISP 8601 standard.
.FUNCTIONALITY
   Sets printer availability
#>
function Set-PrinterAvailability
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies availability start time value. Valid range 0000-2359.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [ValidateRange(0000,2359)]
        [ValidateLength(4,4)]
        [Alias("st")]
        [string]
        $StartTime,

        # Specifies availability end time value. Valid range 0000-2359.
        [Parameter(Mandatory=$true,
                   Position=2,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [ValidateRange(0000,2359)]
        [ValidateLength(4,4)]
        [Alias("ut")]
        [string]
        $UntilTime,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                $StartTimeInput="********00.000000+000"
                $UntilTimeInput="********00.000000+000"
                $StartTimeInput = $StartTimeInput.Insert(8,$StartTime)
                $UntilTimeInput = $UntilTimeInput.Insert(8,$UntilTime)
                $temp = Set-WmiInstance -Path "$Path" -argument @{StartTime="$StartTimeInput"}
                $temp = Set-WmiInstance -Path "$Path" -argument @{UntilTime="$UntilTimeInput"}
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        $StartTimeInput="********00.000000+000"
                        $UntilTimeInput="********00.000000+000"
                        $StartTimeInput = $StartTimeInput.Insert(8,$StartTime)
                        $UntilTimeInput = $UntilTimeInput.Insert(8,$UntilTime)
                        $temp = Set-WmiInstance -Path "$Path" -argument @{StartTime="$StartTimeInput"}
                        $temp = Set-WmiInstance -Path "$Path" -argument @{UntilTime="$UntilTimeInput"}
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                                
                        }
                    }
                    $PrinterName=$OriginalPrinterName
                }
                else
                {
                    if($ShowProgress)
                    {
                        Write-Message -Computer $Computer
                    }
                }
            }
        }
    }
    End
    {
    }
}

#===================================================================================

<#
.Synopsis
   Managing advanced option on printer queue.
.DESCRIPTION
   Set-PrintSpooledFirst cmdlet allows you to enable or disable option "Print Spooled Documents First"
.EXAMPLE
   Set-PrintSpooledFirst -PrinterName "Test Printer" -SetActive "True" -ShowProgress

   This command activates option "Print Spooled Documents First" on  "Test Printer" on the local computer.
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-PrintSpooledFirst -SetActive "True" -ShowProgress

   This set of commands activates option on the local computer for all printers whose name starts with "HP".
.EXAMPLE
    Set-PrintSpooledFirst -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress

    This command activates option on all printers on a remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
.FUNCTIONALITY
   Sets "Print Spooled Documents First" option on printer.
#>
function Set-PrintSpooledFirst
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies desired state of printer option.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("True", "False")]
        [Alias("en")]
        [string]
        $SetActive,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output.
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName            
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                $temp = Set-WmiInstance -Path "$Path" -argument @{DoCompleteFirst="$SetActive"}
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        $temp = Set-WmiInstance -Path "$Path" -argument @{DoCompleteFirst="$SetActive"}
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                        }
                    }
                    $PrinterName=$OriginalPrinterName
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Message -Computer $Computer
		            }
	            }
            }
        }
    }
    End
    {
    }
}

#===================================================================================

<#
.Synopsis
   Managing advanced option on a printer
.DESCRIPTION
   Set-PrintDirectly cmdlet allows you to enable or disable option "Print directly to the printer"
.EXAMPLE
   Set-PrintDirectly -PrinterName "Test Printer" -SetActive "True" -ShowProgress

   This command activates option "Print directly to the printer" on  "Test Printer" on the local computer.
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-PrintDirectly -SetActive "True" -ShowProgress

   This set of commands activates option on the local computer for all printers whose name starts with "HP".
.EXAMPLE
    Set-PrintDirectly -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress

    This command activates option on all printers on a remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
   This cmdlet is not compatible with "Set-SpoolPrintDocuments" cmdlet.
.FUNCTIONALITY
   Sets "Print directly to the printer" option on a printer.
#>
function Set-PrintDirectly
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies desired state of printer option.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("True", "False")]
        [Alias("en")]
        [string]
        $SetActive,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                #Queued property must be set to "False" so that Direct property can accept value "True"
                $temp = Set-WmiInstance -Path "$Path" -argument @{Queued="False"}
                $temp = Set-WmiInstance -Path "$Path" -argument @{Direct="$SetActive"}
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        #Queued property must be set to "False" so that Direct property can accept value "True"
                        $temp = Set-WmiInstance -Path "$Path" -argument @{Queued="False"}
                        $temp = Set-WmiInstance -Path "$Path" -argument @{Direct="$SetActive"}
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                        }
                    }
                    $PrinterName=$OriginalPrinterName
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Message -Computer $Computer
		            }
	            }
            }
        }
    }
    End
    {
    }
}

#===================================================================================

<#
.Synopsis
   Managing advanced option on a printer
.DESCRIPTION
   Set-KeepDocuments cmdlet allows you to enable or disable option "Keep printed documents"
.EXAMPLE
   Set-KeepDocuments -PrinterName "Test Printer" -SetActive "True" -ShowProgress

   This command activates option "Keep printed documents" on  "Test Printer" on the local computer.
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-KeepDocuments -SetActive "True" -ShowProgress

   This set of commands activates option on the local computer for all printers whose name starts with "HP".
.EXAMPLE
    Set-KeepDocuments -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress

    This command activates option on all printers on a remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
.FUNCTIONALITY
   Sets "Keep printed documents" option on a printer.
#>
function Set-KeepDocuments
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies desired state of printer option.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("True", "False")]
        [Alias("en")]
        [string]
        $SetActive,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                $temp = Set-WmiInstance -Path "$Path" -argument @{KeepPrintedJobs="$SetActive"}
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        $temp = Set-WmiInstance -Path "$Path" -argument @{KeepPrintedJobs="$SetActive"}
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                        }
                    }
                    $PrinterName=$OriginalPrinterName
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Output -Computer $Computer
		            }
	            }
             }
        }
    }
    End
    {
    }
}

#===================================================================================

<#
.Synopsis
   Managing advanced option on a printer
.DESCRIPTION
   Set-AdvancedPrintingFeatures cmdlet allows you to enable or disable option "Enable advanced printing features"
.EXAMPLE
   Set-AdvancedPrintingFeatures -PrinterName "Test Printer" -SetActive "True" -ShowProgress

   This command activates option "Enable advanced printing features" on  "Test Printer" on the local computer.
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-AdvancedPrintingFeatures -SetActive "True" -ShowProgress

   This set of commands activates option on the local computer for all printers whose name starts with "HP".
.EXAMPLE
    Set-AdvancedPrintingFeatures -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress

    This command activates option on all printers on a remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
.FUNCTIONALITY
   Sets "Enable advanced printing features" option on a printer.
#>
function Set-AdvancedPrintingFeatures
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies desired state of printer option.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("True", "False")]
        [Alias("en")]
        [string]
        $SetActive,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
        If($SetActive -eq "True")
        {
            $SetActive="False"
        }
        else
        {
            $SetActive="True"
        }
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                $temp = Set-WmiInstance -Path "$Path" -argument @{RawOnly="$SetActive"}
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        $temp = Set-WmiInstance -Path "$Path" -argument @{RawOnly="$SetActive"}
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                        }
                    }
                    $PrinterName=$OriginalPrinterName
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Output -Computer $Computer
		            }
	            }
            }
        }
    }
    End
    {
    }
}

#===================================================================================

<#
.Synopsis
   Managing advanced option on a printer
.DESCRIPTION
   Set-EnableBidiSupport cmdlet allows you to enable or disable option "Enable bidirectional support"
.EXAMPLE
   Set-EnableBidiSupport -PrinterName "Test Printer" -SetActive "True" -ShowProgress

   This command activates option "Enable bidirectional support" on  "Test Printer" on the local computer.
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-EnableBidiSupport -SetActive "True" -ShowProgress

   This set of commands activates option on the local computer for all printers whose name starts with "HP".
.EXAMPLE
    Set-EnableBidiSupport -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress

    This command activates option on all printers on a remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
.FUNCTIONALITY
   Sets "Enable bidirectional support" option on a printer.
#>
function Set-EnableBidiSupport
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies desired state of printer option.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("True", "False")]
        [Alias("en")]
        [string]
        $SetActive,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                $temp = Set-WmiInstance -Path "$Path" -argument @{EnableBIDI="$SetActive"}
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        $temp = Set-WmiInstance -Path "$Path" -argument @{EnableBIDI="$SetActive"}
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                        }
                    }
                    $PrinterName=$OriginalPrinterName
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Output -Computer $Computer
		            }
	            }
            }
        }
    }
    End
    {
    }
}

#===================================================================================

<#
.Synopsis
   Managing advanced option on a printer
.DESCRIPTION
   Set-ListInDirectory cmdlet allows you to enable or disable option "List in the directory"
   This option is available only if printer is shared!
.EXAMPLE
   Set-ListInDirectory -PrinterName "Test Printer" -SetActive "True" -ShowProgress

   This command activates option "List in the directory" on  "Test Printer" on the local computer.
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-ListInDirectory -SetActive "True" -ShowProgress

   This set of commands activates option on the local computer for all printers whose name starts with "HP".
.EXAMPLE
    Set-ListInDirectory -PrinterName * -SetActive "True" -ComputerName "Remote_Computer" -ShowProgress

    This command activates option on all printers on a remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
.FUNCTIONALITY
   Sets "List in the directory" option on a printer.
#>
function Set-ListInDirectory
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies desired state of printer option.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("True", "False")]
        [Alias("en")]
        [string]
        $SetActive,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                $temp = Set-WmiInstance -Path "$Path" -argument @{Published="$SetActive"}
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        $temp = Set-WmiInstance -Path "$Path" -argument @{Published="$SetActive"}
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                        }
                    }
                    $PrinterName=$OriginalPrinterName
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Output -Computer $Computer
		            }
	            }
            }
        }
    }
    End
    {
    }
}

#===================================================================================

<#
.Synopsis
   Managing advanced option on a printer
.DESCRIPTION
   Set-SpoolPrintDocuments cmdlet allows you to define behaviour of spooled files.
   One of two options from -SelectOption parameter must be selected.
.EXAMPLE
   Set-SpoolPrintDocuments -PrinterName "Test Printer" -SelectOption Start_Printing_After_Last_Page -ShowProgress

   This command activates option "Start printing after last page is spooled" on  "Test Printer" on the local computer.
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-SpoolPrintDocuments -SelectOption Start_Printing_Immediately -ShowProgress

   This set of commands activates option "Start printing immediately" on the local computer for all printers whose name starts with "HP".
.EXAMPLE
    Set-SpoolPrintDocuments -PrinterName * -SelectOption Start_Printing_Immediately -ComputerName "Remote_Computer" -ShowProgress

    This command activates option on all printers on a remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
   This cmdlet is not compatible with "Set-PrintDirectly" cmdlet.
.FUNCTIONALITY
   Sets behaviour of spooled documents on a printer queue.
#>
function Set-SpoolPrintDocuments
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies desired state of printer option.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("Start_Printing_After_Last_Page", "Start_Printing_Immediately")]
        [Alias("so")]
        [string]
        $SelectOption,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                If($SelectOption -eq "Start_Printing_After_Last_Page")
                {
                    $temp = Set-WmiInstance -Path "$Path" -argument @{Queued="True"}
                }
                else
                {
                    $temp = Set-WmiInstance -Path "$Path" -argument @{Queued="False"}
                }
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        If($SelectOption -eq "Start_Printing_After_Last_Page")
                        {
                            $temp = Set-WmiInstance -Path "$Path" -argument @{Queued="True"}
                        }
                        else
                        {
                            $temp = Set-WmiInstance -Path "$Path" -argument @{Queued="False"}
                        }
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                        }
                    }
                    $PrinterName=$OriginalPrinterName
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Message -Computer $Computer
		            }
	            }
            }
        }
    }
    End
    {
    }
}

#===================================================================================

<#
.Synopsis
   Managing advanced option on a printer
.DESCRIPTION
   Set-PrinterPriority cmdlet allows you to set priority on printer queue.
.EXAMPLE
   Set-PrinterPriority -PrinterName "Test Printer" -Priority 5 -ShowProgress

   This command activates sets priority on  "Test Printer" on the local computer.
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-PrinterPriority -Priority 5 -ShowProgress

   This set of commands sets priority on the local computer for all printers whose name starts with "HP".
.EXAMPLE
    Set-PrinterPriority -PrinterName * -Priority 5 -ComputerName "Remote_Computer" -ShowProgress

    This command sets priority on all printers on a remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
.FUNCTIONALITY
   Sets priority on a printer queue.
#>
function Set-PrinterPriority
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies printer priority under load. Valid range 1-99.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [ValidateRange(1,99)]
        [Alias("pr")]
        [string]
        $Priority,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                $temp = Set-WmiInstance -Path "$Path" -argument @{Priority="$Priority"}
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        $temp = Set-WmiInstance -Path "$Path" -argument @{Priority="$Priority"}
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                        }
                    }
                    $PrinterName=$OriginalPrinterName
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Message -Computer $Computer
		            }
	            }
            }
        }
    }
    End
    {
    }
}

#===================================================================================

<#
.Synopsis
   Managing location on printer queue
.DESCRIPTION
   Set-PrinterLocation cmdlet allows you to change information about printer location.
.EXAMPLE
   Set-PrinterLocation -PrinterName "Test Printer" -Location "Location 1" -ShowProgress

   This command sets location of "Test Printer" on the local computer.
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-PrinterLocation -Location "Location 2" -ShowProgress

   This set of commands sets location for all printers whose name starts with "HP" on the local computer.
.EXAMPLE
    Set-PrinterLocation -PrinterName * -Location "Location 3" -ComputerName "Remote_Computer" -ShowProgress

    This command sets location for all printers on a remote computer.
.EXAMPLE
    $PrntList = Import-Csv -Path "C:\Locations.csv"
    foreach($Prnt in $PrntList)
    {
        Set-PrinterLocation -PrinterName $Prnt.PrntName -Location $Prnt.Location -ShowProgress
    }

    This short script sets location for all printers depending on content in CSV file.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
.FUNCTIONALITY
   Sets location on a printer queue.
#>
function Set-PrinterLocation
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies location which will be set on printer queue.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [Alias("lo")]
        [string]
        $Location,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        $SwitchTag=$false
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                $temp = Set-WmiInstance -Path "$Path" -argument @{Location="$Location"}
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                        $SwitchTag=$true
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        $temp = Set-WmiInstance -Path "$Path" -argument @{Location="$Location"}
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                        }
                    }
                    if($SwitchTag)
                    {
                    $PrinterName=$OriginalPrinterName
                    }
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Message -Computer $Computer
		            }
	            }
            }
        }
    }
    End
    {
    }
}

#===================================================================================

<#
.Synopsis
   Managing comments on printer queue
.DESCRIPTION
   Set-PrinterComment cmdlet allows you to change comment on printer queue.
.EXAMPLE
   Set-PrinterComment -PrinterName "Test Printer" -Comment "This is a test printer" -ShowProgress

   This command sets comment for "Test Printer" on the local computer.
.EXAMPLE
   (Get-Printer -Name HP*).Name | Set-PrinterComment -Comment "This is a test printer" -ShowProgress

   This set of commands sets comment for all printers whose name starts with "HP" on the local computer.
.EXAMPLE
    Set-PrinterComment -PrinterName * -Comment "This is a test printer on remote computer" -ComputerName "Remote_Computer" -ShowProgress

    This command sets comment for all printers on a remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -PrinterName parameter is used to get all local printers from computer.
.FUNCTIONALITY
   Sets comment on a printer queue.
#>
function Set-PrinterComment
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer on which to set information.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies the comment that will be set on printer queue.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [Alias("cm")]
        [string]
        $Comment,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($PrinterName -eq "*")
            {
                $PrinterName = Get-PrinterName
            }
            foreach($Printer in $PrinterName)
            {
                $Path = Get-PrinterObject $Printer
                $temp = Set-WmiInstance -Path "$Path" -argument @{Comment="$Comment"}
                if($ShowProgress)
                {
                    Write-Message -Printer $Printer
                }
            }
        }
        else
        {
            $SwitchTag=$false
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($PrinterName -eq "*")
                    {
                        $OriginalPrinterName=$PrinterName
                        $PrinterName = Get-PrinterName $Computer
                        $SwitchTag=$true
                    }
                    foreach($Printer in $PrinterName)
                    {
                        $Path = Get-PrinterObject $Printer $Computer
                        $temp = Set-WmiInstance -Path "$Path" -argument @{Comment="$Comment"}
                        if($ShowProgress)
                        {
                            Write-Message -Printer $Printer -Computer $Computer
                        }
                    }
                    if($SwitchTag)
                    {
                    $PrinterName=$OriginalPrinterName
                    }
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Message -Computer $Computer
		            }
	            }
            }
        }
    }
    End
    {
    }
}

#====================================================================================

<#
.Synopsis
   Copy printer settings form one queue to others.
.DESCRIPTION
   Copy-PrinterSettings cmdlet allows you to copy settings to printer queue.
.EXAMPLE
   Copy-PrinterSettings -SourcePrinter "HP" -TargetPrinter "Lexmark"

   This command copies settings from printer queue "HP" to printer queue "Lexmark".
.EXAMPLE
   Copy-PrinterSettings -SourcePrinter "HP" -TargetPrinter "*"

   This command copies settings from printer queue "HP" to all others printer queues on the local computer.
.EXAMPLE
    Copy-PrinterSettings -SourcePrinter "HP" -TargetPrinter "*" -ComputerName "Remote_Computer"

    This command copies settings from printer queue "HP" to all printer queues on the remote computer.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
   General notes

   This cmdlet does not support any wildcards.
   "*" for -TargetPrinter parameter is used to get all local printers from computer.
.FUNCTIONALITY
   Copies settings from one printer to other.
#>
function Copy-PrinterSettings
{
    [CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the name of the printer from which to take settings.
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [Alias("sp")]
        [string]
        $SourcePrinter,

        # Specifies the name of the printer on which settings will be set.
        [Parameter(Mandatory=$true,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [ValidateNotNullOrEmpty()]
        [Alias("tp")]
        [string[]]
        $TargetPrinter,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        if([string]::IsNullOrEmpty($ComputerName))
        {
            If($TargetPrinter -eq "*")
            {
                $TargetPrinter = Get-PrinterName
            }
            foreach($Printer in $TargetPrinter)
            {
                if($Printer -ne $SourcePrinter)
                {
                    $Path = Get-PrinterObject $Printer
                    $temp = Copy-Settings -SetPrinter $Path -SourcePrinter $SourcePrinter
                    if($ShowProgress)
                    {
                        Write-Message -Printer $Printer
                    }
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    If($TargetPrinter -eq "*")
                    {
                        $OriginalTargetPrinter=$TargetPrinter
                        $TargetPrinter = Get-PrinterName $Computer
                    }
                    foreach($Printer in $TargetPrinter)
                    {
                        if($Printer -ne $SourcePrinter)
                        {
                            $Path = Get-PrinterObject $Printer $Computer
                            $temp = Copy-Settings -SetPrinter $Path -SourcePrinter $SourcePrinter
                            if($ShowProgress)
                            {
                                Write-Message -Printer $Printer -Computer $Computer
                            }
                        }
                    }
                    $TargetPrinter=$OriginalTargetPrinter
                }
	            else
	            {
		            if($ShowProgress)
		            {
			            Write-Message -Computer $Computer
		            }
	            }
            }
        }
    }
    End
    {
    }
}

#====================================================================================

<#
.Synopsis
    Creating new printers
.DESCRIPTION
    Create new network printers manually or using printer list in CSV format.
.EXAMPLE
    Add-NetworkPrinter -PrinterIPAddress 0.0.0.0 -PrinterName "Test Printer" -SharedName "TestP" -DriverUsed "Microsoft Print To PDF" -ShowProgress

    This command manually creates Network printer with name "Test Printer" on local computer
.EXAMPLE
    Add-NetworkPrinter -PrinterIPAddress 0.0.0.0 -PrinterName "Test Printer" -SharedName "TestP" -DriverUsed "Microsoft Print To PDF" -ComputerName "Remote_Computer" -ShowProgress

    This command manually creates Network printer with name "Test Printer" on remote computer
.EXAMPLE
    Add-NetworkPrinter -CsvPath "C:\PrinterList.csv" -ShowProgress

    This command creates all printers structured in CSV file on local computer.
.EXAMPLE
    Add-NetworkPrinter -CsvPath "C:\PrinterList.csv" -ComputerName "Remote_Computer" -ShowProgress

    This command creates all printers structured in CSV file on remote computer.
    You can find more information about structuring CSV file in NOTES section.
.INPUTS
   System.Management.Automation.PSObject

   You can pipe objects to this cmdlet.
.OUTPUTS
   System.Management.Automation.PSObject
   
   Returns the objects if -ShowProgress parameter is used.
.NOTES
    Please note that columns in CSV file become parameter for script and because of that they need to be named exactly like it is listed below.
    To ensure right column naming copy next line in first row of you CSV file

        PrinterIPAddress,PrinterName,DriverUsed,SharedName,Location,Comment


    Mandatory Parameters in CSV file:
    1. PrinterIPAddress
    2. PrinterName
    3. DriverUsed

    Optional Parameters in CSV file:
    4. SharedName
    5. Location
    6. Comment

    General notes

    This cmdlet does not support any wildcards.
.FUNCTIONALITY
   Creating new network printer based on provided information
#>
function Add-NetworkPrinter
{
[CmdletBinding(DefaultParameterSetName='ExtendedPrintManagement', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  ConfirmImpact='Medium')]
    Param
    (
        # Specifies the IP address for creating TCP/IP printer port.
        [Parameter(Mandatory=$False, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("ip")]
        [string[]]
        $PrinterIPAddress,

        # Specifies the name for printer which will be created.
        [Parameter(Mandatory=$False,
                   Position=1,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("pn")]
        [string[]]
        $PrinterName,

        # Specifies shared name for printer which will be created.
        [Parameter(Mandatory=$False,
                   Position=2,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("sn")]
        [string[]]
        $SharedName,

        # Specifies print driver that will be used for creating printer.
        [Parameter(Mandatory=$False,
                   Position=3,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("du")]
        [string[]]
        $DriverUsed,

        # Specifies location that will be set on created printer.
        [Parameter(Mandatory=$False,
                   Position=4,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("loc")]
        [string[]]
        $Location,

        # Specifies comment that will be set on created printer.
        [Parameter(Mandatory=$False,
                   Position=5,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("com")]
        [string[]]
        $Comment,

         # Specifies path to CSV file.
        [Parameter(Mandatory=$False,
                   Position=6,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("csv")]
        [string]
        $CsvPath,

        # Specifies the target computer for the management operation. Enter a fully qualified domain name, a NetBIOS name, or an IP address.
        [Parameter(Mandatory=$False,
                   Position=7,
                   ParameterSetName='ExtendedPrintManagement')]
        [Alias("cn")]
        [string[]]
        $ComputerName,

        # Shows progress of cmdlet to console output
        [Parameter(Mandatory=$False,
                   ParameterSetName='ExtendedPrintManagement')]
        [switch]
        $ShowProgress
    )

    Begin
    {
        . $PSScriptRoot\PSLibrary.ps1
    }
    Process
    {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $PrinterList = Import-Csv $CsvPath
        if([string]::IsNullOrEmpty($ComputerName))
        {
            $DriverList = Get-PrintDriverName
            foreach($printer in $PrinterList)
            {
                if($DriverList -contains $printer.DriverUsed)
                {
                    Add-PrinterPort -Name $printer.PrinterIPAddress -PrinterHostAddress $printer.PrinterIPAddress
                    if([string]::IsNullOrEmpty($printer.SharedName))
                    {
                    Add-Printer -Name $printer.PrinterName -PortName $printer.PrinterIPAddress -DriverName $printer.DriverUsed
                    }
                    else
                    {
                    Add-Printer -Name $printer.PrinterName -PortName $printer.PrinterIPAddress -DriverName $printer.DriverUsed -ShareName $printer.SharedName -Shared
                    }

                    if(![string]::IsNullOrEmpty($printer.Location))
                    {
                    Set-PrinterLocation -PrinterName $printer.PrinterName -Location $printer.Location
                    }

                    if(![string]::IsNullOrEmpty($printer.Comment))
                    {
                    Set-PrinterComment -PrinterName $printer.PrinterName -Comment $printer.Comment
                    }

                    if($ShowProgress)
                    {
                        Write-Message -Printer $Printer.PrinterName
                                
                    }
                }
                else
                {
                    $DriverList
                    $printer.DriverUsed
                    $Driver=$printer.DriverUsed
                    Write-Output ("$Driver driver is not installed on this computer!!")
                }
            }
        }
        else
        {
            foreach($Computer in $ComputerName)
            {
                $TestCon = Test-Con $Computer
                if($TestCon)
                {
                    $DriverList = Get-PrintDriverName -ComputerName $Computer
                    foreach($printer in $PrinterList)
                    {
                        if($DriverList -contains $printer.DriverUsed)
                        {
                            Add-PrinterPort -Name $printer.PrinterIPAddress -PrinterHostAddress $printer.PrinterIPAddress -ComputerName $Computer
                            if([string]::IsNullOrEmpty($printer.SharedName))
                            {
                            Add-Printer -Name $printer.PrinterName -PortName $printer.PrinterIPAddress -DriverName $printer.DriverUsed -ComputerName $Computer
                            }
                            else
                            {
                            Add-Printer -Name $printer.PrinterName -PortName $printer.PrinterIPAddress -DriverName $printer.DriverUsed -ShareName $printer.SharedName -Shared -ComputerName $Computer
                            }

                            if(![string]::IsNullOrEmpty($printer.Location))
                            {
                            Set-PrinterLocation -PrinterName $printer.PrinterName -Location $printer.Location -ComputerName $Computer
                            }

                            if(![string]::IsNullOrEmpty($printer.Comment))
                            {
                            Set-PrinterComment -PrinterName $printer.PrinterName -Comment $printer.Comment -ComputerName $Computer
                            }

                            if($ShowProgress)
                            {
                                Write-Message -Printer $Printer.PrinterName -Computer $Computer
                                
                            }
                        }
                        else
                        {
                            $Driver=$printer.DriverUsed
                            Write-Output ("$Driver driver is not installed on computer $Computer!!")
                        }
                    }
                }
                else
                {
                    if($ShowProgress)
                    {
                        Write-Message -Computer $Computer
                    }
                }
            }
        }
        $sw.Stop()
        $sw.Elapsed
    }
    End
    {
    }
}