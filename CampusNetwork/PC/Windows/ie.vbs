'Author: CHEN
'Time: 2023/2/13

Dim ActiveWIFI,x,time
ActiveWIFI=""
Dim TargetWifiName
TargetWifiName = "iHNUST"
time=0
Do Until InStr(ActiveWIFI, TargetWifiName)
	wscript.sleep 2000 '设置延迟
	ActiveWIFI=GetCmdResult("netsh wlan show interfaces | findstr " & TargetWifiName)
	time=time+1
	If time =15 Then
                    msgbox("您当前连接的网络不是 """ & TargetWifiName & "")
	End If
loop
login(x)


Function GetCmdResult(sCmd)
    Dim ws
    Dim ws_exec

    Set ws = CreateObject("Wscript.Shell")
    host = WScript.FullName
    If LCase( right(host, len(host)-InStrRev(host,"\")) ) = "wscript.exe" Then
       ws.run "cscript """ & WScript.ScriptFullName & chr(34), 0
       WScript.Quit
    End If
    Set ws_exec = ws.Exec("cmd.exe /c """ & sCmd & """")
    GetCmdResult = ws_exec.StdOut.ReadAll
    Set ws_exec = Nothing
    Set ws = Nothing

End Function

Function RunCMD(sCmd)
    Dim ws
    Dim ws_exec

    Set ws = CreateObject("Wscript.Shell")
    Set ws_exec = ws.Exec("cmd.exe /c """ & sCmd & """")
    'GetCmdResult = ws_exec.StdOut.ReadAll
    Set ws_exec = Nothing
    Set ws = Nothing

End Function


Function login(x)
	Dim username,password,IE,content
	username="2001080202" '用户名
    password="**********" '密码
	Set IE=CreateObject("InternetExplorer.Application") '调用IE浏览器
	ie.FullScreen=0 '全屏
	IE.Visible=False '可视化？
	IE.Navigate "http://login.hnust.cn/" '打开网站
	Do while IE.ReadyState<>4 or IE.busy '等待网页加载
	wscript.sleep 2000 '设置延迟
	Loop
	On   Error   Resume   Next
		IE.document.querySelector("#edit_body > div.edit_row.ui-resizable-autohide > div.edit_loginBox.normal_box.random.loginuse.loginuse_pc.ui-resizable-autohide > form > input:nth-child(3)").value=username
		IE.document.querySelector("#edit_body > div.edit_row.ui-resizable-autohide > div.edit_loginBox.normal_box.random.loginuse.loginuse_pc.ui-resizable-autohide > form > input:nth-child(4)").value=password
		IE.document.querySelector("#edit_body > div.edit_row.ui-resizable-autohide > div.edit_loginBox.normal_box.random.loginuse.loginuse_pc.ui-resizable-autohide > div.edit_lobo_cell.edit_radio > span:nth-child(3) > input").click '电信为child(3)，校园网为2，联通为4，移动为5
		IE.document.querySelector("#edit_body > div.edit_row.ui-resizable-autohide > div.edit_loginBox.normal_box.random.loginuse.loginuse_pc.ui-resizable-autohide > form > input:nth-child(1)").click '登录
	If   Err.Number   <>   0   Then
		'msgbox("您已登录！")
	End If
                wscript.sleep 3000 '延迟关闭，防止交换未结束IE就关闭了
	'关闭IE
	set wmi=getobject("winmgmts:\\.")
	set pro_s=wmi.instancesof("win32_process")
	for each p in pro_s
	      if p.name="iexplore.exe" then p.terminate()
	next
End Function