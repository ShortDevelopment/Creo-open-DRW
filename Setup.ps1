$oIdent= [Security.Principal.WindowsIdentity]::GetCurrent()
$oPrincipal = New-Object Security.Principal.WindowsPrincipal($oIdent)
if(!$oPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator )){

    Write-Warning "Please start script with Administrator rights! Exit script"
    
}else{

    Start-Process "C:\Program Files\PTC\Creo 7.0.0.0\Parametric\bin\vb_api_register.bat"  -WindowStyle Hidden
    $commmsgexe = 'C:\Program Files\PTC\Creo 7.0.0.0\Common Files\mech\x86e_win64\bin\pro_comm_msg.exe'
    [System.IO.File]::Exists($commmsgexe)
    [System.Environment]::SetEnvironmentVariable('PRO_COMM_MSG_EXE',$commmsgexe,[System.EnvironmentVariableTarget]::Machine)
    $directorypath = 'C:\Program Files\PTC\Creo 7.0.0.0\Common Files\'
    [System.IO.Directory]::Exists($directorypath)
    [System.Environment]::SetEnvironmentVariable('PRO_DIRECTORY', $directorypath,[System.EnvironmentVariableTarget]::Machine)

}