
' ��ʼ�� win32exts
On Error Resume Next
Call win32exts.pop_edx()
If err Then
	Set win32exts = WScript.CreateObject("win32exts.win32atls")
End If
On Error GoTo 0

call win32exts.load_sym("*", "*")
g_pBuf = win32exts.malloc(520)
call win32exts.write_value(g_pBuf, 8, 4, 1600)
call win32exts.write_value(g_pBuf, 12, 4, 1600)

'���� ActiveX �ؼ�
hwnd = 0  '&H000306E2
ax = win32exts.NewActiveXControl_IE(win32exts.L("����2.Tigers5"), hwnd, g_pBuf)

'���� ax.Ax_ShowWindow(1) ��ʾ�ؼ�
call win32exts.ax_call(ax, 0, win32exts.L("Ax_ShowWindow"), win32exts.L("d"), 1)

'���� ax.put_CaptionLBL("�޸ĺ������") �޸Ŀؼ�����
call win32exts.ax_put(ax, win32exts.L("CaptionLBL"), win32exts.L("a"), "�޸ĺ������")

'������Ϣѭ��
call win32exts.SysMessageLoop()
