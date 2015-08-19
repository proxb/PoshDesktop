Function Switch-Desktop {
    [cmdletbinding()]
    Param(
        [parameter(ValueFromPipelineByPropertyName=$True)]
        [string]$Name
    )
    Process {
        $Desktop = Get-Desktop $Name
        Write-Verbose "Switching to Desktop: $($Desktop.Name) <$($Desktop.handle)>"
        $Success = [PoshDesktop]::SwitchDesktop($Desktop.handle)
        If (-NOT $Success) {
            Write-Warning "Unable to switch to $($Name)!"
        }
    }
}