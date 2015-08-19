$ScriptPath = Split-Path $MyInvocation.MyCommand.Path

#region Module Builder
$Domain = [AppDomain]::CurrentDomain
$DynAssembly = New-Object System.Reflection.AssemblyName('PoshDesktop')
$AssemblyBuilder = $Domain.DefineDynamicAssembly($DynAssembly, [System.Reflection.Emit.AssemblyBuilderAccess]::Run) # Only run in memory
$ModuleBuilder = $AssemblyBuilder.DefineDynamicModule('PoshDesktop', $False)
#endregion Module Builder

#region ENUMs
#region ACCESS_MASK
$EnumBuilder = $ModuleBuilder.DefineEnum('ACCESS_MASK', 'Public', [int32])
[void]$EnumBuilder.DefineLiteral('DELETE', [int32] 0x00010000)
[void]$EnumBuilder.DefineLiteral('READ_CONTROL', [int32] 0x00020000)
[void]$EnumBuilder.DefineLiteral('WRITE_DAC', [int32] 0x00040000)
[void]$EnumBuilder.DefineLiteral('WRITE_OWNER', [int32] 0x00080000)
[void]$EnumBuilder.DefineLiteral('SYNCHRONIZE', [int32] 0x00100000)
[void]$EnumBuilder.DefineLiteral('STANDARD_RIGHTS_REQUIRED', [int32] 0x000F0000)
[void]$EnumBuilder.DefineLiteral('STANDARD_RIGHTS_READ', [int32] 0x00020000)
[void]$EnumBuilder.DefineLiteral('STANDARD_RIGHTS_WRITE', [int32] 0x00020000)
[void]$EnumBuilder.DefineLiteral('STANDARD_RIGHTS_EXECUTE', [int32] 0x00020000)
[void]$EnumBuilder.DefineLiteral('STANDARD_RIGHTS_ALL', [int32] 0x001F0000)
[void]$EnumBuilder.DefineLiteral('SPECIFIC_RIGHTS_ALL', [int32] 0x0000FFFF)
[void]$EnumBuilder.DefineLiteral('ACCESS_SYSTEM_SECURITY', [int32] 0x01000000)
[void]$EnumBuilder.DefineLiteral('MAXIMUM_ALLOWED', [int32] 0x02000000)
[void]$EnumBuilder.DefineLiteral('GENERIC_READ', [int32] 0x80000000)
[void]$EnumBuilder.DefineLiteral('GENERIC_WRITE', [int32] 0x40000000)
[void]$EnumBuilder.DefineLiteral('GENERIC_EXECUTE', [int32] 0x20000000)
[void]$EnumBuilder.DefineLiteral('GENERIC_ALL', [int32] 0x10000000)
[void]$EnumBuilder.DefineLiteral('DESKTOP_ALL', [int32] 0xf01ff)
[void]$EnumBuilder.DefineLiteral('DESKTOP_READOBJECTS', [int32] 0x00000001)
[void]$EnumBuilder.DefineLiteral('DESKTOP_CREATEWINDOW', [int32] 0x00000002)
[void]$EnumBuilder.DefineLiteral('DESKTOP_CREATEMENU', [int32] 0x00000004)
[void]$EnumBuilder.DefineLiteral('DESKTOP_HOOKCONTROL', [int32] 0x00000008)
[void]$EnumBuilder.DefineLiteral('DESKTOP_JOURNALRECORD', [int32] 0x00000010)
[void]$EnumBuilder.DefineLiteral('DESKTOP_JOURNALPLAYBACK', [int32] 0x00000020)
[void]$EnumBuilder.DefineLiteral('DESKTOP_ENUMERATE', [int32] 0x00000040)
[void]$EnumBuilder.DefineLiteral('DESKTOP_WRITEOBJECTS', [int32] 0x00000080)
[void]$EnumBuilder.DefineLiteral('DESKTOP_SWITCHDESKTOP', [int32] 0x00000100)
[void]$EnumBuilder.DefineLiteral('WINSTA_ENUMDESKTOPS', [int32] 0x00000001)
[void]$EnumBuilder.DefineLiteral('WINSTA_READATTRIBUTES', [int32] 0x00000002)
[void]$EnumBuilder.DefineLiteral('WINSTA_ACCESSCLIPBOARD', [int32] 0x00000004)
[void]$EnumBuilder.DefineLiteral('WINSTA_CREATEDESKTOP', [int32] 0x00000008)
[void]$EnumBuilder.DefineLiteral('WINSTA_WRITEATTRIBUTES', [int32] 0x00000010)
[void]$EnumBuilder.DefineLiteral('WINSTA_ACCESSGLOBALATOMS', [int32] 0x00000020)
[void]$EnumBuilder.DefineLiteral('WINSTA_EXITWINDOWS', [int32] 0x00000040)
[void]$EnumBuilder.DefineLiteral('WINSTA_ENUMERATE', [int32] 0x00000100)
[void]$EnumBuilder.DefineLiteral('WINSTA_READSCREEN', [int32] 0x00000200)
[void]$EnumBuilder.DefineLiteral('WINSTA_ALL_ACCESS', [int32] 0x0000037F)
[void]$EnumBuilder.CreateType()
#endregion ACCESS_MASK
#endregion ENUMs

#region Structs
#region System.PoshDesktop.Desktop
$Attributes = 'AutoLayout, AnsiClass, Class, Public, SequentialLayout, Sealed, BeforeFieldInit'
$STRUCT_TypeBuilder = $ModuleBuilder.DefineType('System.PoshDesktop.Desktop', $Attributes, [System.ValueType])
[void]$STRUCT_TypeBuilder.DefineField('Handle', [intptr], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('Name', [string], 'Public')
[void]$STRUCT_TypeBuilder.CreateType()
#endregion System.PoshDesktop.Desktop
#region STARTUPINFO
$Attributes = 'AutoLayout, AnsiClass, Class, Public, SequentialLayout, Sealed, BeforeFieldInit'
$STRUCT_TypeBuilder = $ModuleBuilder.DefineType('STARTUPINFO', $Attributes, [System.ValueType])
[void]$STRUCT_TypeBuilder.DefineField('cb', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('lpReserved', [string], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('lpDesktop', [string], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('lpTitle', [string], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('dwX', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('dwY', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('dwXSize', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('dwYSize', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('dwXCountChars', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('dwYCountChars', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('dwFillAttribute', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('dwFlags', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('wShowWindow', [Int16], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('cbReserved2', [Int16], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('lpReserved2', [intptr], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('hStdInput', [intptr], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('hStdOutput', [intptr], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('hStdError', [intptr], 'Public')
[void]$STRUCT_TypeBuilder.CreateType()
#endregion STARTUPINFO
#region PROCESS_INFORMATION
$Attributes = 'AutoLayout, AnsiClass, Class, Public, SequentialLayout, Sealed, BeforeFieldInit'
$STRUCT_TypeBuilder = $ModuleBuilder.DefineType('PROCESS_INFORMATION', $Attributes, [System.ValueType])
[void]$STRUCT_TypeBuilder.DefineField('hProcess', [intptr], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('hThread', [intptr], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('dwProcessId', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('dwThreadId', [Int32], 'Public')
[void]$STRUCT_TypeBuilder.CreateType()
#endregion PROCESS_INFORMATION
#region SECURITY_ATTRIBUTES
$Attributes = 'AutoLayout, AnsiClass, Class, Public, SequentialLayout, Sealed, BeforeFieldInit'
$STRUCT_TypeBuilder = $ModuleBuilder.DefineType('SECURITY_ATTRIBUTES', $Attributes, [System.ValueType])
[void]$STRUCT_TypeBuilder.DefineField('nLength', [int], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('lpSecurityDescriptor', [intptr], 'Public')
[void]$STRUCT_TypeBuilder.DefineField('bInheritHandle', [int], 'Public')
[void]$STRUCT_TypeBuilder.CreateType()
#endregion SECURITY_ATTRIBUTES
#endregion Structs

#region Delegates
#region EnumDesktops Delegate
$DelegateBuilder = $ModuleBuilder.DefineType('EnumDesktops', 'Class, Public, Sealed, AnsiClass, AutoClass', [System.MulticastDelegate])
$ConstructorBuilder = $DelegateBuilder.DefineConstructor('RTSpecialName, HideBySig, Public', [System.Reflection.CallingConventions]::Standard, @([string],[intptr]))
$ConstructorBuilder.SetImplementationFlags('Runtime, Managed')
$MethodBuilder = $DelegateBuilder.DefineMethod('Invoke', 'Public, HideBySig, NewSlot, Virtual', [bool], @([string],[intptr]))
$MethodBuilder.SetImplementationFlags('Runtime, Managed')
#endregion EnumDesktops Delegate
#region Create Delegate
[void]$DelegateBuilder.CreateType()
#endregion Create Delegate
#endregion Delegates

#region Classes
#region Initialize Class Builder
$TypeBuilder = $ModuleBuilder.DefineType('PoshDesktop', 'Public, Class')
#endregion Class Type Builder
#endregion Classes

#region Methods
#region GetCurrentThreadId 
$PInvokeMethod = $TypeBuilder.DefineMethod(
    'GetCurrentThreadId', #Method Name
    [Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl', #Method Attributes
    [IntPtr], #Method Return Type
    [Type[]] @() #Method Parameters
)
$DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
$FieldArray = [Reflection.FieldInfo[]] @(
    [Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'),
    [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
    [Runtime.InteropServices.DllImportAttribute].GetField('ExactSpelling')
    [Runtime.InteropServices.DllImportAttribute].GetField('CharSet')
)

$FieldValueArray = [Object[]] @(
    'GetCurrentThreadId', #CASE SENSITIVE!!
    $True,
    $False,
    [System.Runtime.InteropServices.CharSet]::Unicode
)

$SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder(
    $DllImportConstructor,
    @('kernel32.dll'),
    $FieldArray,
    $FieldValueArray
)

$PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
#endregion GetCurrentThreadId 
#region GetProcessWindowStation 
$PInvokeMethod = $TypeBuilder.DefineMethod(
    'GetProcessWindowStation', #Method Name
    [Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl', #Method Attributes
    [IntPtr], #Method Return Type
    [Type[]] @() #Method Parameters
)
$DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
$FieldArray = [Reflection.FieldInfo[]] @(
    [Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'),
    [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
    [Runtime.InteropServices.DllImportAttribute].GetField('ExactSpelling')
    [Runtime.InteropServices.DllImportAttribute].GetField('CharSet')
)

$FieldValueArray = [Object[]] @(
    'GetProcessWindowStation', #CASE SENSITIVE!!
    $True,
    $False,
    [System.Runtime.InteropServices.CharSet]::Unicode
)

$SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder(
    $DllImportConstructor,
    @('user32.dll'),
    $FieldArray,
    $FieldValueArray
)

$PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
#endregion GetProcessWindowStation 
#region EnumDesktops 
$PInvokeMethod = $TypeBuilder.DefineMethod(
    'EnumDesktops', #Method Name
    [Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl', #Method Attributes
    [bool], #Method Return Type
    [Type[]] @(
        [IntPtr],
        [EnumDesktops],
        [IntPtr]
    ) #Method Parameters
)
$DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
$FieldArray = [Reflection.FieldInfo[]] @(
    [Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'),
    [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
    [Runtime.InteropServices.DllImportAttribute].GetField('ExactSpelling')
)

$FieldValueArray = [Object[]] @(
    'EnumDesktops', #CASE SENSITIVE!!
    $True,
    $False
)

$SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder(
    $DllImportConstructor,
    @('user32.dll'),
    $FieldArray,
    $FieldValueArray
)

$PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
#endregion EnumDesktops 
#region OpenDesktop 
$PInvokeMethod = $TypeBuilder.DefineMethod(
    'OpenDesktop', #Method Name
    [Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl', #Method Attributes
    [IntPtr], #Method Return Type
    [Type[]] @(
        [string],
        [Int32],
        [bool],
        [long]
    ) #Method Parameters
)
$DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
$FieldArray = [Reflection.FieldInfo[]] @(
    [Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'),
    [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
    [Runtime.InteropServices.DllImportAttribute].GetField('ExactSpelling')
    [Runtime.InteropServices.DllImportAttribute].GetField('CharSet')
)

$FieldValueArray = [Object[]] @(
    'OpenDesktop', #CASE SENSITIVE!!
    $True,
    $False,
    [System.Runtime.InteropServices.CharSet]::Unicode
)

$SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder(
    $DllImportConstructor,
    @('user32.dll'),
    $FieldArray,
    $FieldValueArray
)

$PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
#endregion OpenDesktop 
#region CloseDesktop 
$PInvokeMethod = $TypeBuilder.DefineMethod(
    'CloseDesktop', #Method Name
    [Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl', #Method Attributes
    [bool], #Method Return Type
    [Type[]] @(
        [IntPtr]
    ) #Method Parameters
)
$DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
$FieldArray = [Reflection.FieldInfo[]] @(
    [Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'),
    [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
    [Runtime.InteropServices.DllImportAttribute].GetField('ExactSpelling')
    [Runtime.InteropServices.DllImportAttribute].GetField('CharSet')
)

$FieldValueArray = [Object[]] @(
    'CloseDesktop', #CASE SENSITIVE!!
    $True,
    $False,
    [System.Runtime.InteropServices.CharSet]::Unicode
)

$SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder(
    $DllImportConstructor,
    @('user32.dll'),
    $FieldArray,
    $FieldValueArray
)

$PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
#endregion CloseDesktop 
#region GetThreadDesktop 
$PInvokeMethod = $TypeBuilder.DefineMethod(
    'GetThreadDesktop', #Method Name
    [Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl', #Method Attributes
    [IntPtr], #Method Return Type
    [Type[]] @(
        [Int]
    ) #Method Parameters
)
$DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
$FieldArray = [Reflection.FieldInfo[]] @(
    [Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'),
    [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
    [Runtime.InteropServices.DllImportAttribute].GetField('ExactSpelling')
    [Runtime.InteropServices.DllImportAttribute].GetField('CharSet')
)

$FieldValueArray = [Object[]] @(
    'GetThreadDesktop', #CASE SENSITIVE!!
    $True,
    $False,
    [System.Runtime.InteropServices.CharSet]::Unicode
)

$SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder(
    $DllImportConstructor,
    @('user32.dll'),
    $FieldArray,
    $FieldValueArray
)

$PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
#endregion GetThreadDesktop 
#region CreateDesktop 
$PInvokeMethod = $TypeBuilder.DefineMethod(
    'CreateDesktop', #Method Name
    [Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl', #Method Attributes
    [IntPtr], #Method Return Type
    [Type[]] @(
        [string], #Desktop 
        [IntPtr], #Device 
        [IntPtr], #Device Mode
        [int],    #Flags
        [ACCESS_MASK],
        [IntPtr]
    ) #Method Parameters
)
$DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
$FieldArray = [Reflection.FieldInfo[]] @(
    [Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'),
    [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
    [Runtime.InteropServices.DllImportAttribute].GetField('ExactSpelling')
    [Runtime.InteropServices.DllImportAttribute].GetField('CharSet')
)

$FieldValueArray = [Object[]] @(
    'CreateDesktop', #CASE SENSITIVE!!
    $True,
    $False,
    [System.Runtime.InteropServices.CharSet]::Unicode
)

$SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder(
    $DllImportConstructor,
    @('user32.dll'),
    $FieldArray,
    $FieldValueArray
)

$PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
#endregion CreateDesktop 
#region CreateProcess 
$PInvokeMethod = $TypeBuilder.DefineMethod(
    'CreateProcess', #Method Name
    [Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl', #Method Attributes
    [bool], #Method Return Type
    [Type[]] @(
        [string], #Application Name 
        [string], #Command Line 
        [IntPtr],
        [IntPtr],
        [bool],
        [int32],
        [IntPtr],
        [string],
        [STARTUPINFO].MakeByRefType(),
        [PROCESS_INFORMATION].MakeByRefType()
    ) #Method Parameters
)
$DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
$FieldArray = [Reflection.FieldInfo[]] @(
    [Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'),
    [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
    [Runtime.InteropServices.DllImportAttribute].GetField('ExactSpelling')
)

$FieldValueArray = [Object[]] @(
    'CreateProcess', #CASE SENSITIVE!!
    $True,
    $False
)

$SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder(
    $DllImportConstructor,
    @('kernel32.dll'),
    $FieldArray,
    $FieldValueArray
)

$PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
#endregion CreateProcess 
#region SwitchDesktop 
$PInvokeMethod = $TypeBuilder.DefineMethod(
    'SwitchDesktop', #Method Name
    [Reflection.MethodAttributes] 'PrivateScope, Public, Static, HideBySig, PinvokeImpl', #Method Attributes
    [bool], #Method Return Type
    [Type[]] @(
        [IntPtr]
    ) #Method Parameters
)
$DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
$FieldArray = [Reflection.FieldInfo[]] @(
    [Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'),
    [Runtime.InteropServices.DllImportAttribute].GetField('SetLastError')
    [Runtime.InteropServices.DllImportAttribute].GetField('ExactSpelling')
)

$FieldValueArray = [Object[]] @(
    'SwitchDesktop', #CASE SENSITIVE!!
    $True,
    $True
)

$SetLastErrorCustomAttribute = New-Object Reflection.Emit.CustomAttributeBuilder(
    $DllImportConstructor,
    @('user32.dll'),
    $FieldArray,
    $FieldValueArray
)

$PInvokeMethod.SetCustomAttribute($SetLastErrorCustomAttribute)
#endregion SwitchDesktop 
#endregion Methods

#region Create Type
[void]$TypeBuilder.CreateType()
#endregion Create Type

#region Load Functions
Try {
    Get-ChildItem "$ScriptPath\Scripts" -Filter *.ps1 | Select -Expand FullName | ForEach {
        $Function = Split-Path $_ -Leaf
        . $_
    }
} Catch {
    Write-Warning ("{0}: {1}" -f $Function,$_.Exception.Message)
    Continue
}
#endregion Load Functions

#region Aliases
New-Alias -Name ndsk -Value New-Desktop
New-Alias -Name gdsk -Value Get-Desktop
New-Alias -Name swdsk -Value Switch-Desktop
#endregion Aliases

Export-ModuleMember -Alias ndsk,gdsk,swdsk -Function 'New-Desktop','Get-Desktop','Switch-Desktop'