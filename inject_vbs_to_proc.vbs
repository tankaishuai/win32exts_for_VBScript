on error resume next
err.number = 0
set win32exts = CreateObject("win32exts.win32atls")
if err.number <> 0 then
  'msgbox "请先安装 Win32Exts"
  Wscript.Quit
end if
on error goto 0


sub run_in_target_proc()
  '
  '
  '
  '
  msgbox "这里是被注入的进程执行的代码位置！！"
  '
  '
  '
  '
end sub


'pSysSharedCommonBufferText 是一个 4kb 的公共缓存区
pSysSharedCommonBufferText = win32exts.load_sym(".", "SysSharedCommonBufferText")
shared_flag = win32exts.read_value(pSysSharedCommonBufferText, 0, 4, false)
if shared_flag = 1 then
  call win32exts.write_value(pSysSharedCommonBufferText, 0, 4, 0)
  call run_in_target_proc
  wscript.Quit
end if


call win32exts.load_sym("*", "*")
call win32exts.MessageBoxW(0, win32exts.L("下面我们开始注入代码到指定的目标进程"), win32exts.L("start call MessageBoxW"), 1)


hwnd = 0
while hwnd = 0
  wnd_text = inputbox("请输入目标进程的主窗口标题：")
  if len(wnd_text) > 0 then
    hwnd = clng(win32exts.FindWindowA(0, wnd_text))
    if hwnd = 0 then
      msgbox "找不到输入的窗口，请重新输入！"
    end if
  else
    wscript.Quit
  end if
wend


tid = clng(win32exts.GetWindowThreadProcessId(hwnd, 0))
if tid = 0 then
  msgbox "指定的窗口已经失效！！"
  wscript.Quit
end if


'
'下面是我们需要注入到目标进程的代码，用自己
'
inject_file = wscript.ScriptFullName


pSysGetMsgHookProc = clng(win32exts.load_sym(".", "SysGetMsgHookProc"))
pSysSharedExecCommandLine = clng(win32exts.load_sym(".", "SysSharedExecCommandLine"))
call win32exts.write_astring(pSysSharedExecCommandLine, 0, -1, inject_file)
call win32exts.write_value(pSysSharedCommonBufferText, 0, 4, 1)
hHook = clng(win32exts.SetWindowsHookExW(3, pSysGetMsgHookProc, clng(win32exts.current_dll()), tid))
if hHook = 0 then
  msgbox "注入失败！！"
  wscript.Quit
end if
call win32exts.PostMessageW(hwnd, 0, 0, 0)


msgbox "注入成功"
