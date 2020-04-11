
' 初始化 win32exts
On Error Resume Next
Call win32exts.pop_edx()
If err Then
	Set win32exts = WScript.CreateObject("win32exts.win32atls")
End If
On Error GoTo 0

call win32exts.load_sym("*", "*")
g_pBuf = win32exts.malloc(520)
call win32exts.write_value(g_pBuf, 8, 4, 600)
call win32exts.write_value(g_pBuf, 12, 4, 600)

'创建 ActiveX 控件
hwnd = 0      '11798276
ax = win32exts.NewActiveXControl_IE(win32exts.L("工程2.Tigers5"), hwnd, g_pBuf)

'调用 ax.Ax_ShowWindow(1) 显示控件
call win32exts.AxWrapper_SimpleInvokeHelper(ax, win32exts.L("Ax_ShowWindow"), win32exts.L("d"), 1)

'调用 ax.put_CaptionLBL("修改后的数据") 修改控件属性
call win32exts.AxWrapper_InvokeHelper(ax, 4, 0, win32exts.L("CaptionLBL"), win32exts.L("a"), "修改后的数据")

'进入消息循环
call win32exts.SysMessageLoop()
