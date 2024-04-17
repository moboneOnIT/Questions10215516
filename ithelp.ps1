
#�Х��]�w�n���|�ѼơA�ڥu�Χڹq���W���]�w�U�h�w�]���|���A�Ш̷ӧA���q�����|�ק�

$LinePath = $env:LOCALAPPDATA+"\LINE\bin\LineLauncher.exe"
$WordPath = ${env:ProgramFiles(x86)}+"\Microsoft Office\root\Office16"


#��������~�T���ɤ����Xĵ�i�T��
$ErrorActionPreference = "SilentlyContinue"

##PS�����ܦh�����ܼƥi�H�ޥΡA��֥��@��S�Ϊ��r�A�n�d�ߥثe��Ϊ������ܼƥi�H��
#dir env:
#�N�i�H��ܧA�ثe�q�����������ܼƤF

if (($PSVersionTable).PSVersion.Major -lt 5){
    Write-Host "�ЦbPS�������H�W�����ҤU�ϥ�����"
    Read-Host -Prompt "�п�J���@������"
}

Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class User32 {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }
"@
#�ϥ�WinAPI���������F��̤j�Ƹ�e�x���
#https://learn.microsoft.com/zh-tw/windows/win32/api/winuser/nf-winuser-showwindow
#https://learn.microsoft.com/zh-tw/windows/win32/api/winuser/nf-winuser-setforegroundwindow


<#

    �p�G�A�u�O�Q�n���}�{���A�Шϥ�
    Start-Process -FilePath "�ɮ׸��|"
    �M��A�f�tdo{
        Start-Sleep -Seconds 1
        if (�����{���O�_���X�ӤF){
        �p�G���Nbreak���X�j��
        }
    }while($true)    


    AppName�������N�O�u�@�޲z�������ԲӸ�ƪ��W��
    Line   => Line
    Chrome => Chrome
    Word   => WINWORD

    Mode�u���J�Ʀr�A�����p�U
    0 => ���õ���
    1 => ��ܵ���(���u���^�_��W�������A)
    2 => �Ұʵ����̤p��
    3 => �Ұʵ����̤j��
    4 => ��ܵ���(�S���Ұ�)
    5 => ��ܵ���

    Example :   

    ShowAppFromTaskBar "WINWORD" 3 �̤j�ƫe�x���
    ShowAppFromTaskBar "WINWORD" 4 �w�]�����j�p���
    ShowAppFromTray "Line" 4 �w�]�����j�p���



#>




function ShowAppFromTaskBar{
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$AppName,
        [Parameter(Mandatory=$true, Position=1)]
        [int]$Mode
    )
    try{
        $APP = (Get-Process -Name "$AppName" -WarningAction SilentlyContinue) |Where-Object {$_.MainWindowHandle -ne 0}
        if ($APP -ne $null){
            $hwnd = $APP[0].MainWindowHandle
            if ($hwnd -ne 0){
                [User32]::ShowWindow($hwnd, $Mode)
                [User32]::SetForegroundWindow($hwnd)
                Write-Host "��ܦ��\"
            }
        }
    }catch{
        Write-Host "���~�T���G $($_.Exception.Message)"
        return $null
    }
}

function ShowAppFromTray{    
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [string]$AppName,
        [Parameter(Mandatory=$true, Position=1)]
        [int]$Mode
    )

    try{
        #�Ұ�APP 
        $appPath = GetExePath $AppName
        Start-Process $appPath
        do{
            Start-Sleep -Seconds 1
            $APP = (Get-Process -Name "$AppName" -WarningAction SilentlyContinue) |Where-Object {$_.MainWindowHandle -ne 0}
            if ($APP -ne $null){
                break
            }
        }while($true)

        if ($APP -ne $null){
            $hwnd = $APP[0].MainWindowHandle
            

            if ($hwnd -ne 0){
                [User32]::ShowWindow($hwnd, $Mode)
                [User32]::SetForegroundWindow($hwnd)
                Write-Host "��ܦ��\"
            }
        }
    }catch{
        Write-Host "���~�T���G $($_.Exception.Message)"
        return $null
    }
}

function GetExePath([string]$AppName){
    try{
        if (!([regex]::IsMatch($AppName,".*(?:\.)exe"))){ #���W��F���p�G�ŦX�h�s�W.exe��Ѽƫ��
            $AppName+=".exe"
        }
        $App =  Get-WmiObject Win32_Process -Filter "Name='$AppName'" 
        if ($App -ne$null){
            if ($App.Count -gt 0){
                return  $App.Path[0]
            }
            return  $App.Path 
        }
    }catch{
        return $null    
    }
}




