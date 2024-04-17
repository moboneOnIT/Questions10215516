#ForQuestion https://ithelp.ithome.com.tw/questions/10215516

<#
#請先設定好路徑參數，我只用我電腦上的設定下去預設路徑的，請依照你的電腦路徑修改
#$LinePath = $env:LOCALAPPDATA+"\LINE\bin\LineLauncher.exe"
#$WordPath = ${env:ProgramFiles(x86)}+"\Microsoft Office\root\Office16"
如果你只是想要打開程式，請使用
    Start-Process -FilePath $程式變數名稱
    然後再搭配do{
        Start-Sleep -Seconds 1
        if (偵測程式是否有出來了){
            如果有就break跳出迴圈
        }
    }while($true)    
#>
#偵測到錯誤訊息時不跳出警告訊息
$ErrorActionPreference = "SilentlyContinue"

##PS內有很多環境變數可以引用，減少打一堆沒用的字，要查詢目前能用的環境變數可以打
#dir env:
#就可以顯示你目前電腦有的環境變數了

if (($PSVersionTable).PSVersion.Major -lt 5){
    Write-Host "請在PS版本五以上的環境下使用謝謝"
    Read-Host -Prompt "請輸入任一鍵關閉"
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
#使用WinAPI來讓視窗達到最大化跟前台顯示
#https://learn.microsoft.com/zh-tw/windows/win32/api/winuser/nf-winuser-showwindow
#https://learn.microsoft.com/zh-tw/windows/win32/api/winuser/nf-winuser-setforegroundwindow


<#
    主要程式使用教學
    AppName對應的就是工作管理員內的詳細資料的名稱
    Line   => Line
    Chrome => Chrome
    Word   => WINWORD

    Mode只能輸入數字，對應如下
    0 => 隱藏視窗
    1 => 顯示視窗(但只有回復到上次的狀態)
    2 => 啟動視窗最小化
    3 => 啟動視窗最大化
    4 => 顯示視窗(沒有啟動)
    5 => 顯示視窗

    Example :   
    ShowAppFromTaskBar "WINWORD" 3  #最大化前台顯示
    ShowAppFromTaskBar "WINWORD" 4 #預設視窗大小顯示
    ShowAppFromTray "Line" 4 #預設視窗大小顯示


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
                Write-Host "顯示成功"
            }
        }
    }catch{
        Write-Host "錯誤訊息： $($_.Exception.Message)"
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
        #啟動APP 
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
                Write-Host "顯示成功"
            }
        }
    }catch{
        Write-Host "錯誤訊息： $($_.Exception.Message)"
        return $null
    }
}

function GetExePath([string]$AppName){
    try{
        if (!([regex]::IsMatch($AppName,".*(?:\.)exe"))){ #正規表達式如果符合則新增.exe到參數後方
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




