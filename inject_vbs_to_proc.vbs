on error resume next
err.number = 0
set win32exts = CreateObject("win32exts.win32atls")
if err.number <> 0 then
  'msgbox "���Ȱ�װ Win32Exts"
  Wscript.Quit
end if
on error goto 0


sub run_in_target_proc()
  '
  '
  '
  '
  msgbox "�����Ǳ�ע��Ľ���ִ�еĴ���λ�ã���"
  '
  '
  '
  '
end sub


'pSysSharedCommonBufferText ��һ�� 4kb �Ĺ���������
pSysSharedCommonBufferText = win32exts.load_sym(".", "SysSharedCommonBufferText")
shared_flag = win32exts.read_value(pSysSharedCommonBufferText, 0, 4, false)
if shared_flag = 1 then
  call win32exts.write_value(pSysSharedCommonBufferText, 0, 4, 0)
  call run_in_target_proc
  wscript.Quit
end if


call win32exts.load_sym("*", "*")
call win32exts.MessageBoxW(0, win32exts.L("�������ǿ�ʼע����뵽ָ����Ŀ�����"), win32exts.L("start call MessageBoxW"), 1)


hwnd = 0
while hwnd = 0
  wnd_text = inputbox("������Ŀ����̵������ڱ��⣺")
  if len(wnd_text) > 0 then
    hwnd = clng(win32exts.FindWindowA(0, wnd_text))
    if hwnd = 0 then
      msgbox "�Ҳ�������Ĵ��ڣ����������룡"
    end if
  else
    wscript.Quit
  end if
wend


tid = clng(win32exts.GetWindowThreadProcessId(hwnd, 0))
if tid = 0 then
  msgbox "ָ���Ĵ����Ѿ�ʧЧ����"
  wscript.Quit
end if


'
'������������Ҫע�뵽Ŀ����̵Ĵ��룬���Լ�
'
inject_file = wscript.ScriptFullName


pSysGetMsgHookProc = clng(win32exts.load_sym(".", "SysGetMsgHookProc"))
pSysSharedExecCommandLine = clng(win32exts.load_sym(".", "SysSharedExecCommandLine"))
call win32exts.write_astring(pSysSharedExecCommandLine, 0, -1, inject_file)
call win32exts.write_value(pSysSharedCommonBufferText, 0, 4, 1)
hHook = clng(win32exts.SetWindowsHookExW(3, pSysGetMsgHookProc, clng(win32exts.current_dll()), tid))
if hHook = 0 then
  msgbox "ע��ʧ�ܣ���"
  wscript.Quit
end if
call win32exts.PostMessageW(hwnd, 0, 0, 0)


msgbox "ע��ɹ�"
