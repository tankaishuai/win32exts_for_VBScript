'
' create a new win32exts wrapper, 
' use "win32exts.win32atls_exe" for separated process,
' and "win32exts.win32atls" for owned process.
'
set win32exts = CreateObject("win32exts.win32atls")

'
' load library symbols
'
Call win32exts.load_sym("kernel32", "*")
Call win32exts.load_sym("gdi32", "*")
Call win32exts.load_sym("user32", "*")

'
' allocate a buffer
'
g_buf = win32exts.malloc( 2*260 )

'
' sample: call MessageBoxW
'
xx = 0
xx = win32exts.MessageBoxW(0, win32exts.L("start call MessageBoxW11"), null, 1)

'
' sample: get host module file name
'
Call win32exts.GetModuleFileNameW(win32exts.current_dll(), g_buf, 260)
strCurDll = win32exts.read_wstring(g_buf, 0, -1)
Call win32exts.MessageBoxW(0, win32exts.L(strCurDll), null, 1)

'
' sample: call GetWindowTextW API
'
iCount = 0

function OnEnumProc(args)
	hWnd = win32exts.arg(args, 1)
	iRetVal = win32exts.GetWindowTextW(hWnd, g_buf, 256)
	strVal = win32exts.read_wstring(g_buf, 0, -1)
	
	'
	' retern value is a string like: "retval, add_esp_bytes",
	' for stdcall add_esp_bytes usually equals to 4 * arg_count,
	' and cdecl   add_esp_bytes equals to 0.
	'
	strRet = "1, 8"
	if(strVal <> "") then
		Call win32exts.MessageBoxW(0, win32exts.L(strVal), null, 1)
		iCount = iCount + 1
		if(iCount = 4) then
			strRet = "0, 8"
		end if
	end if
	OnEnumProc = strRet
end function

iRetVal = win32exts.EnumWindows(GetRef("OnEnumProc"), 0)

'
' exit
'
Call win32exts.free(g_buf)
