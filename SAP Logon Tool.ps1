# ＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞Funcation Group＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜
function Open-Gui{
    param(
        [string]$sid,
        [string]$client,
        [string]$username,
        [string]$password
    )
    Start-Process "cmd.exe" "/c start sapshcut -system=$($sid) -client=$($client) -user=$($username) -pw=$($password)"
}

function LogonData-Update{
    param(
        [string]$sname,
        [string]$sid,
        [string]$client,
        [string]$username,
        [string]$password,
        [string]$passwordn,
        [string]$buttonColor
        )
        $Me = ".\logondata\$env:USERNAME.ps1"
        $Myself = Get-Content $Me
        $Myself `
        | Foreach-Object -Process {$temp = $_.split(",")
                                    $rule = '^\$null[\s]\=[\s]\$logondata.Add\(@\("{0}","{1}","{2}","{3}"' -f $sname , $sid, $client, $username, '"[\S\s]*$'
                                    if($_ -match $rule){

                                    # パスワード変更しない
                                        if($password -eq $passwordn){
                                            '$null = $logondata.Add(@("{0}","{1}","{2}","{3}","{4}","{5}",{6},{7}' -f $sname, $sid, $client, $username, $passwordn, $buttonColor, '($group+=1)', $temp[7]
                                        }else{
                                        # パスワード変更する
                                            Password-Change -sname ($sname.Split("：➢：")[0]) -client $client -username $username -password $password -passwordn $passwordn
                                            '$null = $logondata.Add(@("{0}","{1}","{2}","{3}","{4}","{5}",{6},"{7}"))' -f $sname, $sid, $client, $username, $passwordn, $buttonColor, '($group+=1)', (Get-Date).ToString("yyyy/MM/dd")
                                        }
                                        $msg = (New-Object -ComObject WScript.Shell).popup("Logon Data更新完了、再起動後に反映します。",0,"Logon Data更新")
                                    }else{
                                        $_
                                    }
                                   }`
        | Set-Content $Me -Encoding utf8 
        Remove-Variable -Name Myself,Me
}

function LogonData-Remove{
    param(
        [string]$sname,
        [string]$sid,
        [string]$client,
        [string]$username
        )

        $Me = ".\logondata\$env:USERNAME.ps1"
        $Myself = Get-Content $Me
        $Myself `
        | Foreach-Object -Process {$temp = $_.split(" ")
                                    $rule = '^\$null[\s]\=[\s]\$logondata.Add\(@\("' + $sname + '","' + $sid + '","' + $client + '","' + $username + '"[\S\s]*$'
                                    if($_ -notmatch $rule ){
                                        $_
                                    }else{
                                        $msg = (New-Object -ComObject WScript.Shell).popup("Logon Data削除完了、再起動後に反映します。",0,"Logon Data削除")
                                    }
                                   }`
        | Set-Content $Me -Encoding utf8 
        Remove-Variable -Name Myself,Me
}

function LogonData-Add{
    param(
        [string]$sname,
        [string]$sid,
        [string]$client,
        [string]$username,
        [string]$password,
        [string]$buttonColor
        )
        $Me = ".\logondata\$env:USERNAME.ps1"
        $Myself = Get-Content $Me

        $flg = 0
        $Myself `
        | Foreach-Object -Process {$temp = $_.split(" ")
                                    if($sname -notmatch "：➢："){
                                        $sname = $sname + "：➢：" + $username
                                    }
                                    if($_ -match '^#End(\s)'+ $sname){
                                        #$_ = ''
                                        '$null = $logondata.Add(@("{0}","{1}","{2}","{3}","{4}","{5}",{6},"{7}"))' -f $sname, $sid, $client, $username, $password, $buttonColor, '($group+=1)', (Get-Date).ToString("yyyy/MM/dd")
                                        '#End {0}' -f $sname
                                        $flg = 1
                                        $msg = (New-Object -ComObject WScript.Shell).popup("【Client追加】`nLogon Data追加完了①`n再起動後に反映します。",0,"Logon Data追加")
                                    
                                    }elseif($flg -eq 0 -and $_ -match '^#EndLogonData$'){
                                        #$_ = ''
                                        '$group = 0'
                                        '$null = $logondata.Add(@("{0}","{1}","{2}","{3}","{4}","{5}",{6},"{7}"))' -f $sname, $sid, $client, $username, $password, $buttonColor, '($group+=1)', (Get-Date).ToString("yyyy/MM/dd")
                                        '#End {0}' -f $sname
                                        '#EndLogonData'
                                        $msg = (New-Object -ComObject WScript.Shell).popup("【System追加】`nLogon Data追加完了②`n再起動後に反映します。",0,"Logon Data追加")
                                    
                                    }else{
                                        $_
                                    }
                                   }`
        | Set-Content $Me -Encoding utf8 
        Remove-Variable -Name Myself,Me
}

Function Password-Change{
    param(
        [string]$sname,
        [string]$client,
        [string]$username,
        [string]$password,
        [string]$passwordn
    
    )
    
    $vbsCode = @'
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
   'Set connection = application.Children(0)
   Set connection = application.Openconnection("{0}",True)
End If

If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If

If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "{1}"
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "{2}"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "{3}"
session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "JA"

session.findById("wnd[0]/tbar[1]/btn[5]").press
session.findById("wnd[1]/usr/pwdRSYST-NCODE").text = "{4}"
session.findById("wnd[1]/usr/pwdRSYST-NCOD2").text = "{5}"
session.findById("wnd[1]/tbar[0]/btn[0]").press

'@ -f $sname, $client, $username, $password, $passwordn, $passwordn

    Set-Content -Path .\pass.vbs -Value $vbsCode
    Start-Process "wscript.exe" -ArgumentList ".\temp.vbs"
    Remove-Item -Path ".\temp.vbs" -Force
}

function Create-Shortcut {
    param (
        [string]$ShortcutName,
        [string]$TargetPath,
        [string]$ShortcutPath
    )
    
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $TargetPath
    $Shortcut.Save()
}

# ＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞Logon Data＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜
$filePath = ".\logondata\$env:USERNAME.ps1"
if(Test-Path -Path $filePath){
    . $filePath
}else{
    Create-Shortcut -ShortcutName 'SAP Logon Tool.lnk' -TargetPath ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) -ShortcutPath ([System.IO.Path]::Combine([System.Environment]::GetFolderPath("Desktop"),"SAP Logon Tool.lnk"))
    $null = New-Item -Path ".\logondata" -ItemType "directory" -Force 
    Set-Content -Path $filePath -Value '$logondata = New-Object System.Collections.ArrayList'
    Add-Content -Path $filePath -Value '#EndLogonData'
    #Copy-Item -Path ".\SAP Logon Tool.lnk" -Destination ([System.IO.Path]::Combine($env:USERPROFILE, "Desktop")) -Force
}

# ＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞Layout＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜

#Layout
Add-type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = 'SAP GUI Logon Tool(Designed by jiajun.xie)'
$form.Size = New-Object System.Drawing.Size(800,400)
$form.Padding = New-Object System.Windows.Forms.Padding(10)
$form.StartPosition = "CenterScreen"

#Left Panel
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$leftPanel.Dock = [System.Windows.Forms.DockStyle]::Left
$leftPanel.Width = ($form.Width / 5*2.1 - 10)

#leftTitle
$leftTitle = New-Object System.Windows.Forms.Panel
$leftTitle.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$leftTitle.Dock = [System.Windows.Forms.DockStyle]::Top
$leftTitle.Width = $leftPanel.Width
$leftTitle.Height = ($form.Height / 3.6)


$leftTitleText = New-Object System.Windows.Forms.Label
$leftTitleText.Text = "1.緑色、金色、青色、グレーを設定できます。`n" +"`n"+
　　　　　　　        "2. デフォルト色は青色です。`n" +"`n"+
　　　　　　　        "3. 赤色になると、パスワード期限切れってログオン出来なくなり、パスワードを変更してください。。`n" +"`n"+
                      "4.作成者：jiajun.xie "
$leftTitleText.Width = $leftTitle.Width
$leftTitleText.Height = $leftTitle.Height
$leftTitle.Controls.Add($leftTitleText)

#button Group
$left = $leftPanel.Width * 0.1
$high_left = $leftTitle.Height + 10

$lastlogon =@("","","","","","","","")
foreach ($logon in $logondata){

    $button = New-Object System.Windows.Forms.Button
    #$button = [ThreeDButton]::new()
    if($logon[6] -eq 1){
        $left = $leftPanel.Width * 0.02
        $high_left += $leftPanel.Height * 0.25
        $buttonTitle = New-Object System.Windows.Forms.Label
        $buttonTitle.Text = $logon[0]
        $buttonTitle.Size = New-Object System.Drawing.Size($leftPanel.Width,($leftPanel.Height * 0.15))
        $buttonTitle.Location = New-Object System.Drawing.Point($left,$high_left)
        $buttonTitle.BackColor = [System.Drawing.Color]::AliceBlue
        
        
        $leftPanel.Controls.add($buttonTitle)
        $high_left += $leftPanel.Height * 0.2
        $button.Text = $logon[2]
        $button.Size = New-Object System.Drawing.Size (($leftPanel.Width * 0.1),($leftPanel.Height * 0.2))
        $button.Location = New-Object System.Drawing.Point($left,$high_left)
    }else{
        # Group 
        if($logon[6] % 7 -ne 1){
            $left += $leftPanel.Width * 0.14
            $button.Text = $logon[2]
            $button.Size = New-Object System.Drawing.Size (($leftPanel.Width * 0.1),($leftPanel.Height * 0.2))
            $button.Location = New-Object System.Drawing.Point($left,$high_left)
        }else{
            $left = $leftPanel.Width * 0.02
            $high_left += $leftPanel.Height * 0.25
            $button.Text = $logon[2]
            $button.Size = New-Object System.Drawing.Size (($leftPanel.Width * 0.1),($leftPanel.Height * 0.2))
            $button.Location = New-Object System.Drawing.Point($left,$high_left)
            
        }
    }

    
    $lastlogon = $logon
    $leftPanel.Controls.add($button) 

    switch($logon[5]){
        緑色     {$button.BackColor = [System.Drawing.Color]::LightGreen}
        金色     {$button.BackColor = [System.Drawing.Color]::LightYellow }
        グレー   {$button.BackColor = [System.Drawing.Color]::LightGray}
        Default{
            $button.BackColor = [System.Drawing.Color]::LightBlue
        }
    }
    $date = [datetime]::Parse($logon[7])
    $days = ((Get-Date) - $date).Days
    if($days -ge 70 -and $logon[1] -eq "MES"){
        $button.BackColor = [System.Drawing.Color]::PaleVioletRed
        $button.Add_Click({$msg = (New-Object -ComObject WScript.Shell).popup("パスワード期限切れました、変更してください。",0,"ログオン出来ません。")}.GetNewClosure())
    }else{
        $button.Add_Click({Open-Gui -sid $logon[1] -client $logon[2] -username $logon[3] -password $logon[4]}.GetNewClosure())
    }
}


$leftPanel.Controls.Add($leftTitle)
$form.Controls.Add($leftPanel)

#Right Panel
$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$rightPanel.Dock = [System.Windows.Forms.DockStyle]::Right
$rightPanel.Width = ($form.Width / 5*2.9 - 30)

#rightTitle
$rightTitle = New-Object System.Windows.Forms.Panel
$rightTitle.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$rightTitle.Dock = [System.Windows.Forms.DockStyle]::Top
$rightTitle.Width = $rightPanel.Width
$rightTitle.Height = ($form.Height / 3.6)

$rightTitleText = New-Object System.Windows.Forms.Label
$rightTitleText.Text = "1.パスワード更新について、システムにパスワード更新する場合は必ず更新してください。`n" +"`n"+
                       "2.追加、更新、削除いずれか実施した後、当ツールを再起動してください。`n"+"`n"+
                       "3.作成者：jiajun.xie "
$rightTitleText.Width = $rightTitle.Width
$rightTitleText.Height = $rightTitle.Height
$rightTitle.Controls.Add($rightTitleText)
$rightPanel.Controls.Add($rightTitle)

#button Group
$left = $rightPanel.Width * 0.1
$high_right = $rightTitle.Height + 10

$high_right += $rightPanel.Height * 0.25
$dol = New-Object System.Windows.Forms.Label
$dol.Text = "操作"
$dol.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.1),$high_right)
$dol.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))
$rightPanel.Controls.add($dol)

$doA = New-Object System.Windows.Forms.RadioButton
$doA.Text = "追加"
$doA.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.3),($high_right - 3))
$doA.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))
$rightPanel.Controls.add($doA)

$doB = New-Object System.Windows.Forms.RadioButton
$doB.Text = "変更"
$doB.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.5),($high_right - 3))
$doB.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))
$doB.Checked = $true

$rightPanel.Controls.add($doB)

$doC = New-Object System.Windows.Forms.RadioButton
$doC.Text = "削除"
$doC.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.7),($high_right - 3))
$doC.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))
$rightPanel.Controls.add($doC)

$high_right += $rightPanel.Height * 0.3

$systemname_lab = New-Object System.Windows.Forms.Label
$systemname_lab.Text = "タイトル"
$systemname_lab.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.1),$high_right)
$systemname_lab.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))

$systemname_btn = New-Object System.Windows.Forms.ComboBox
$systemname_btn.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.3),($high_right - 3))
$systemname_btn.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.65),($rightTitle.Height * 0.2))

$rightPanel.Controls.add($systemname_lab)
$rightPanel.Controls.add($systemname_btn)

$high_right += $rightPanel.Height * 0.3
$systemid_lab = New-Object System.Windows.Forms.Label
$systemid_lab.Text = "システムID"
$systemid_lab.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.1),$high_right)
$systemid_lab.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))

$systemid_btn = New-Object System.Windows.Forms.ComboBox
$systemid_btn.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.3),($high_right - 3))
$systemid_btn.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.65),($rightTitle.Height * 0.2))

$rightPanel.Controls.add($systemid_lab)
$rightPanel.Controls.add($systemid_btn)

$high_right += $rightPanel.Height * 0.3
$client_lab = New-Object System.Windows.Forms.Label
$client_lab.Text = "クライアント"
$client_lab.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.1),$high_right)
$client_lab.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))

$client_btn = New-Object System.Windows.Forms.ComboBox
$client_btn.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.3),($high_right - 3))
$client_btn.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.65),($rightTitle.Height * 0.2))

$rightPanel.Controls.add($client_lab)
$rightPanel.Controls.add($client_btn)

$high_right += $rightPanel.Height * 0.3
$username_lab = New-Object System.Windows.Forms.Label
$username_lab.Text = "ユーザ名"
$username_lab.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.1),$high_right)
$username_lab.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))

$username_btn = New-Object System.Windows.Forms.ComboBox
$username_btn.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.3),($high_right - 3))
$username_btn.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.65),($rightTitle.Height * 0.2))

$rightPanel.Controls.add($username_lab)
$rightPanel.Controls.add($username_btn)

$high_right += $rightPanel.Height * 0.3
$password_lab = New-Object System.Windows.Forms.Label
$password_lab.Text = "パスワード"
$password_lab.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.1),$high_right)
$password_lab.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))

$password_btn = New-Object System.Windows.Forms.TextBox
$password_btn.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.3),($high_right - 3))
$password_btn.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.65),($rightTitle.Height * 0.2))

$rightPanel.Controls.add($password_lab)
$rightPanel.Controls.add($password_btn)

$high_right += $rightPanel.Height * 0.3
$passwordn_lab = New-Object System.Windows.Forms.Label
$passwordn_lab.Text = "新パスワード"
$passwordn_lab.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.1),$high_right)
$passwordn_lab.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))

$passwordn_btn = New-Object System.Windows.Forms.TextBox
$passwordn_btn.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.3),($high_right - 3))
$passwordn_btn.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.65),($rightTitle.Height * 0.2))
$passwordn_btn.BackColor = [System.Drawing.Color]::LightYellow

$rightPanel.Controls.add($passwordn_lab)
$rightPanel.Controls.add($passwordn_btn)

$high_right += $rightPanel.Height * 0.3
$buttonColor_lab = New-Object System.Windows.Forms.Label
$buttonColor_lab.Text = "ボタン色"
$buttonColor_lab.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.1),$high_right)
$buttonColor_lab.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.2),($rightTitle.Height * 0.2))

$buttonColor_btn = New-Object System.Windows.Forms.ComboBox
$buttonColor_btn.Location = New-Object System.Drawing.Point(($rightPanel.Width * 0.3),($high_right - 3))
$buttonColor_btn.Size = New-Object System.Drawing.Size(($rightPanel.Width * 0.65),($rightTitle.Height * 0.2))

$null=$buttonColor_btn.Items.Add("緑色")
$null=$buttonColor_btn.Items.Add("金色")
$null=$buttonColor_btn.Items.Add("グレー")
$null=$buttonColor_btn.Items.Add("青色")

$rightPanel.Controls.add($buttonColor_lab)
$rightPanel.Controls.add($buttonColor_btn)

$high_right += $rightPanel.Height * 0.3
$ok_btn = New-Object System.Windows.Forms.Button
$ok_btn.Text =　"確定"
$ok_btn.Location = New-Object System.Drawing.Point(($rightPanel.Width *0.3),$high_right)
$ok_btn.Size = New-Object System.Drawing.Size (($rightPanel.Width *0.3),($rightPanel.Height *0.3))

$restart_btn = New-Object System.Windows.Forms.Button
$restart_btn.Text =　"再起動"
$restart_btn.Location = New-Object System.Drawing.Point(($rightPanel.Width *0.65),$high_right)
$restart_btn.Size = New-Object System.Drawing.Size (($rightPanel.Width *0.3),($rightPanel.Height *0.3))

# ＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞＞Default Value＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜＜
$snm_options = @()
foreach($temp in $logondata){
    $snm_options += $temp[0]
    
}
$snm_options | Select-Object -Unique | ForEach-Object{$null=$systemname_btn.Items.Add($_)}

$doA.Add_Click({
                $passwordn_lab.Visible = $true
                $passwordn_btn.Visible = $true
                $buttoncolor_lab.Visible = $true
                $buttonColor_btn.Visible = $true
                $passwordn_lab.Visible = $false
                $passwordn_btn.Visible = $false})

$doB.Add_Click({
                $password_lab.Visible = $true
                $password_btn.Visible = $true
                $passwordn_lab.Visible = $true
                $passwordn_btn.Visible = $true
                $buttoncolor_lab.Visible = $true
                $buttonColor_btn.Visible = $true
                })

$doC.Add_Click({
                $password_lab.Visible = $false
                $password_btn.Visible = $false
                $passwordn_lab.Visible = $false
                $passwordn_btn.Visible = $false
                $buttoncolor_lab.Visible = $false
                $buttonColor_btn.Visible = $false
          })


$systemname_btn.Add_SelectedIndexChanged({
    $sid_options = @()
    $systemid_btn.Items.Clear()
    foreach($temp in $logondata){
        if($temp[0] -eq $systemname_btn.SelectedItem){
            $sid_options += $temp[1]
        }
    
    }
    $sid_options | Select-Object -Unique | ForEach-Object{$null=$systemid_btn.Items.Add($_)}

    $systemid_btn.SelectedIndex = 0

})

$systemid_btn.Add_SelectedIndexChanged({
    $clt_options = @()
    $client_btn.Items.Clear()
    foreach($temp in $logondata){
        if($temp[1] -eq $systemid_btn.SelectedItem){
            $clt_options += $temp[2]
        }
    }
    $clt_options | Select-Object -Unique | ForEach-Object{$null=$client_btn.Items.Add($_)}
})

$systemid_btn.Add_SelectedIndexChanged({
    $usr_options = @()
    $username_btn.Items.Clear()
    foreach($temp in $logondata){
        if($temp[1] -eq $systemid_btn.SelectedItem){
            $usr_options += $temp[3]
        }
    }
    $usr_options | Select-Object -Unique | ForEach-Object{$null=$username_btn.Items.Add($_)}
})

$username_btn.Add_SelectedIndexChanged({
    foreach($temp in $logondata){
        if($temp[1] -eq $systemid_btn.SelectedItem -and $temp[2] -eq $client_btn.SelectedItem -and $temp[3] -eq $username_btn.SelectedItem){
            $password_btn.Text = $temp[4]
            $buttonColor_btn.SelectedItem = $temp[5]
        }
    }
})

$client_btn.Add_SelectedIndexChanged({
    foreach($temp in $logondata){
        if($temp[1] -eq $systemid_btn.SelectedItem -and $temp[2] -eq $client_btn.SelectedItem -and $temp[3] -eq $username_btn.SelectedItem){
            $password_btn.Text = $temp[4]
            $buttonColor_btn.SelectedItem = $temp[5]
        }
    }
})

$systemname_btn.Add_SelectedIndexChanged({
    foreach($temp in $logondata){
        if($temp[1] -eq $systemid_btn.Text -and $temp[2] -eq $client_btn.Text -and $temp[3] -eq $username_btn.Text){
            $password_btn.Text = $temp[4]
            $buttonColor_btn.SelectedItem = $temp[5]
        }
    }
})

$client_btn.Add_SelectedIndexChanged({
    foreach($temp in $logondata){
        if($temp[1] -eq $systemid_btn.Text -and $temp[2] -eq $client_btn.Text -and $temp[3] -eq $username_btn.Text){
            $password_btn.Text = $temp[4]
            $buttonColor_btn.SelectedItem = $temp[5]
        }
    }
})

$username_btn.Add_SelectedIndexChanged({
    foreach($temp in $logondata){
        if($temp[1] -eq $systemid_btn.Text -and $temp[2] -eq $client_btn.Text -and $temp[3] -eq $username_btn.Text){
            $password_btn.Text = $temp[4]
            $buttonColor_btn.SelectedItem = $temp[5]
        }
    }
})

$Me = $MyInvocation.MyCommand.Source

$ok_btn.Add_Click({
                        if($doA.Checked){LogonData-Add    -sname $systemname_btn.Text         -sid $systemid_btn.Text         -client $client_btn.Text         -username $username_btn.Text         -password $password_btn.Text  -buttonColor $buttonColor_btn.Text}
                    elseif($doB.Checked){LogonData-Update -sname $systemname_btn.SelectedItem -sid $systemid_btn.SelectedItem -client $client_btn.SelectedItem -username $username_btn.SelectedItem -password $password_btn.Text -passwordn $passwordn_btn.Text -buttonColor $buttonColor_btn.Text
                                         $restart_btn.PerformClick()}
                    elseif($doC.Checked){LogonData-Remove -sname $systemname_btn.SelectedItem -sid $systemid_btn.SelectedItem -client $client_btn.SelectedItem -username $username_btn.SelectedItem}
                 })
$rightPanel.Controls.add($ok_btn)

$restart_btn.Add_Click({#Start-Process powershell.exe -ArgumentList "-WindowStyle Hidden -File `"$Me`" "
                  Start-Process -FilePath "C:\SAP Logon Tool\SAP Logon Tool.lnk"
                  $form.Close()
                  $Global:shouldExit = $true
                 })
$rightPanel.Controls.add($restart_btn)

$form.Controls.Add($rightPanel)

$form.Add_Resize({
    $leftPanel.Width = ($form.Width / 5*2.1 - 10) 
    $rightPanel.Width = ($form.Width / 5*2.9 - 30)
    #$leftTitle.Height = ($form.Height / 3.7)
    #$rightTitle.Height = ($form.Height / 3.7)
})
$form.Height = [Math]::Max($high_left , $high_right)+ 100

$null=$form.ShowDialog()
