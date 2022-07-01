<###############################################################
Tera Term ランチャー
###############################################################>
#Requires -Version 3

using namespace System.Management.Automation
Add-Type -AssemblyName System.Windows.Forms;

$APPLICATION_NAME = "Tera Term ランチャー"
$MUTEX_NAME = 'CE634DBE-31E2-4E65-838A-11CF22BDBBC4';

$TERATERM_PATH = "C:\Program Files (x86)\teraterm\ttermpro.exe"
$SERVER_FILE = "$PSScriptRoot/接続先.csv"
$EDITOR_PATH = "notepad.exe"

$mutex = New-Object System.Threading.Mutex($false, $MUTEX_NAME);

# 接続先の選択画面を開く
function displaySelect {
    try {
        $serverlist = Import-Csv -Path $SERVER_FILE -Delimiter "," -Encoding utf8
        $select = $serverlist | Select-Object User, Host, Kanji, Memo | Out-GridView -OutputMode Multiple -Title "接続先の選択"
        $select | ForEach-Object {
            $tmp = $_
            $serverData = $serverlist | Where-Object { ($_.Host -eq $tmp.Host) -and ($_.User -eq $tmp.User) } | Select-Object -First 1
            & $TERATERM_PATH `
            /ssh "$($serverData.Host)" /2 /auth=password `
            /user="$($serverData.User)" /passwd="$($serverData.Password)" /L="$HOME\&h_%Y%m%d.log" /KR="$($serverData.Kanji)"
        }
    }
    catch {
        $notifyIcon.BalloonTipText = $_.ToString();
        $notifyIcon.ShowBalloonTip(5000);
    }

}
function displayServerFile {
    try {
        & $EDITOR_PATH $SERVER_FILE
    }
    catch {
        $notifyIcon.BalloonTipText = $_.ToString();
        $notifyIcon.ShowBalloonTip(5000);
    }
}

function displayTooltip {
    try {
        # コンテキスト作成
        $appContext = New-Object System.Windows.Forms.ApplicationContext;

        ####################################################
        # 通知アイコン作成
        ####################################################
        # TeraTermのアイコンを設定する
        $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($TERATERM_PATH)
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
                    displaySelect
                }
            });

        ####################################################
        # アイコン右クリック時のコンテキストメニューの設定
        ####################################################
        $menuItem_exit = [System.Windows.Forms.ToolStripMenuItem]@{ Text = 'Exit' };
        $menuItem_editServerFile = [System.Windows.Forms.ToolStripMenuItem]@{ Text = '接続先を編集' };

        $notifyIcon.ContextMenuStrip = New-Object System.Windows.Forms.ContextMenuStrip;
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_editServerFile);
        [void]$notifyIcon.ContextMenuStrip.Items.Add($menuItem_exit);

        $menuItem_exit.add_Click( {
            $appContext.ExitThread();
        });
        $menuItem_editServerFile.add_Click( {
            displayServerFile
        });

        $notifyIcon.Visible = $true;

        # 起動時は接続先の選択画面を開く
        displaySelect

        # 表示
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
        if (Test-Path -LiteralPath $TERATERM_PATH -PathType Leaf) {
            hiddenTaskber
            displayTooltip
            $retcode = 0;
        } else {
            [System.Windows.Forms.MessageBox]::Show("Tera Termが見つかりません。`n$TERATERM_PATH", $APPLICATION_NAME, 0, 16) > $null
            $retcode = 255;
        }
        
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("すでに起動中です。", $APPLICATION_NAME, 0, 48) > $null
        $retcode = 255;
    }
}
finally {
    $mutex.Dispose();
}
exit $retcode;
