Function Get-Desktop {
    [OutputType('System.PoshDesktop.Desktop')]
    [cmdletbinding()]
    Param(
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string[]]$Name
    )
    If ($PSBoundParameters.ContainsKey('Name')) {
        $Pattern = $($Name -join '|')
    } Else {
        $Pattern = '.*'
    }
    
    $DesktopHandle = [PoshDesktop]::GetProcessWindowStation()
    $script:desktops = New-Object System.Collections.ArrayList
    $Return = [PoshDesktop]::EnumDesktops(
        $DesktopHandle,
        {[void]$script:desktops.Add($args[0]); $true },
        [IntPtr]::Zero
    )
    If ($Return) {
        $script:desktops | ForEach {
            If ($_ -match $Pattern) {
                $Name = $_
                $Desktop = [poshdesktop]::OpenDesktop(
                    $Name,
                    0,
                    $True,
                    [ACCESS_MASK]::MAXIMUM_ALLOWED
                )
                $Object = New-Object System.PoshDesktop.Desktop
                $Object.Handle = $Desktop
                $Object.Name = $Name
                $Object
            }
        }
    }
}