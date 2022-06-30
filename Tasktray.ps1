<###############################################################
タスクトレイ常駐スクリプト
###############################################################>
#Requires -Version 3

using namespace System.Management.Automation
Add-Type -AssemblyName System.Windows.Forms;

$APPLICATION_NAME = "アプリケーション名"
$MUTEX_NAME = 'CE634DBE-31E2-4E65-838A-11CF22BDBBC4';

$mutex = New-Object System.Threading.Mutex($false, $MUTEX_NAME);

function displayTooltip {
    try {
        # コンテキスト作成
        $appContext = New-Object System.Windows.Forms.ApplicationContext;

        ####################################################
        # 通知アイコン作成
        ####################################################
        # PowerShellのアイコンを設定する
        $pwshPath = Get-Process -id $pid | Select-Object -ExpandProperty Path
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($pwshPath)    
        $notifyIcon = [System.Windows.Forms.NotifyIcon]@{
            Icon           = $icon;
            Text           = $APPLICATION_NAME;
            BalloonTipIcon = 'None';
        };

        ####################################################
        # アイコン左クリック時のイベントを設定
        ####################################################
        $notifyIcon.add_Click( {
                if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
                    try {
                        $notifyIcon.BalloonTipText = (Get-Date);
                        $notifyIcon.ShowBalloonTip(5000);
                    }
                    catch {
                        $notifyIcon.BalloonTipText = $_.ToString();
                        $notifyIcon.ShowBalloonTip(5000);
                    }
                }
            });

        ####################################################
        # アイコン右クリック時のコンテキストメニューの設定
        ####################################################
        $menuItem_exit = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'Exit' };
        $menuItem_function1 = [System.Windows.Forms.ToolStripMenuItem]@{ Text = '右クリックメニュー１' };
        $menuItem_function2 = [System.Windows.Forms.ToolStripMenuItem]@{ Text = '右クリックメニュー２' };
        
        $notifyIcon.ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip;
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_function1);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_function2);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_exit);

        $menuItem_exit.add_Click( {
            $appContext.ExitThread();
        });
        $menuItem_function1.add_Click( {
            try {
                $notifyIcon.BalloonTipText = "右クリックメニュー１がクリックされました。";
                $notifyIcon.ShowBalloonTip(5000);
            }
            catch {
                $notifyIcon.BalloonTipText = $_.ToString();
                $notifyIcon.ShowBalloonTip(5000);
            }
        });
        $menuItem_function2.add_Click( {
            try {
                $notifyIcon.BalloonTipText = "右クリックメニュー２がクリックされました。";
                $notifyIcon.ShowBalloonTip(5000);
            }
            catch {
                $notifyIcon.BalloonTipText = $_.ToString();
                $notifyIcon.ShowBalloonTip(5000);
            }
        });

        # 表示
        $notifyIcon.Visible = $true;
        [void][System.Windows.Forms.Application]::Run($appContext);
        $notifyIcon.Visible = $false;
    
    }
    finally {
        $notifyIcon.Dispose();
        $mutex.ReleaseMutex();
    }
}

# タスクバー非表示
function hiddenTaskber {
    $windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
    $asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
    $null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
}
try {
    # タイトルバーの書き換え
    $Host.UI.RawUI.WindowTitle = $APPLICATION_NAME
    # 多重起動チェック
    if ($mutex.WaitOne(0, $false)) {
        hiddenTaskber
        displayTooltip
        $retcode = 0;
    }
    else {
        $retcode = 255;
    }
}
finally {
    $mutex.Dispose();
}
exit $retcode;