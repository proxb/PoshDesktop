Function New-Desktop {
    [OutputType('System.PoshDesktop.Desktop')]
    [cmdletbinding()]
    Param(
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$Name
    )

    Write-Verbose "Creating desktop: $($Name)"
    $Handle = [PoshDesktop]::CreateDesktop(
        $Name,
        [IntPtr]::Zero,
        [IntPtr]::Zero,
        0,
        [ACCESS_MASK]::DESKTOP_ALL,
        [IntPtr]::Zero
    )
    Write-Verbose "Handle: $($Handle)"
    If ($Handle -ne [intptr]::Zero) {
        Write-Verbose "Starting explorer on $($Name)"
        $STARTUPINFO = New-Object STARTUPINFO
        $STARTUPINFO.cb = [System.Runtime.InteropServices.Marshal]::SizeOf($STARTUPINFO)
        $STARTUPINFO.lpDesktop = $Name
        $PROCESS_INFORMATION = New-Object PROCESS_INFORMATION
        $CommandLine = 'C:\WINDOWS\Explorer.exe'
        $Success = [PoshDesktop]::CreateProcess(
            $CommandLine, 
            $null, 
            [IntPtr]::Zero, 
            [IntPtr]::Zero, 
            $false, 
            [IntPtr]::Zero, 
            [IntPtr]::Zero, 
            $env:windir, 
            [ref] $STARTUPINFO, 
            [ref] $PROCESS_INFORMATION
        )   
        If (-NOT $Success) {
            Write-Warning "Unable to created Explorer.exe process on desktop: $($Name)"
        } 
        $Object = New-Object System.PoshDesktop.Desktop
        $Object.Handle = $Handle
        $Object.Name = $Name
        $Object
    } Else {
        Write-Warning "Unable to create new desktop: $($Name)!"
    }
}