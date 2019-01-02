[General]
SyntaxVersion=2
MacroID=652ab43d-d143-47d0-819d-61745b09cade
[Comment]

[Script]
'======================================[需要脚本定制可以找我]======================================
'【神梦多线程命令库】
'版本：v1.0
'更新：2018.12.31
'作者：神梦无痕
'ＱＱ：1042207232
'Ｑ群：624655641
'================================================================================================
'
'PC端挂机宝VPN，单进程单IP，购买联系Q：1042207232
'
'代查QQ、手机号是否开通微信，精确度高达99%，购买联系Q：1042207232
'
'================================================================================================
'
/*-------------------------按键精灵里不支持下面的API
'原子锁递减（跨进程）
Declare Function InterlockedDecrement Lib "kernel32.dll" Alias "InterlockedDecrement" (ByVal lpAddend As Long) As Long 
'原子锁递增
Declare Function InterlockedIncrement Lib "kernel32.dll" Alias "InterlockedIncrement" (ByVal lpAddend As Long) As Long 
'原子锁赋值
Declare Function InterlockedExchange Lib "kernel32.dll" Alias "InterlockedExchange" (Target As Long, ByVal Value As Long) As Long
'原子锁运算
Declare Function InterlockedExchangeAdd Lib "kernel32.dll" Alias "InterlockedExchangeAdd" (Addend As Long, ByVal Value As Long) As Long
'原子锁三目运算
Declare Function InterlockedCompareExchange Lib "kernel32.dll" Alias "InterlockedExchangeAdd" (Destination As Long, ByVal Exchange As Long, ByVal Comperand As Long) As Long
'
'临界许可创建（跨线程）
Declare Function InitializeCriticalSection Lib "kernel32.dll" Alias "InitializeCriticalSection" (lpCriticalSection As Any) As Long 
'临界区尝试进入
Declare Function TryEnterCriticalSection Lib "kernel32.dll" Alias "TryEnterCriticalSection" (lpCriticalSection As Any) As Boolean 
'临界区进入
Declare Function EnterCriticalSection Lib "kernel32.dll" Alias "EnterCriticalSection" (lpCriticalSection As Any) As Long 
'临界区退出
Declare Function LeaveCriticalSection Lib "kernel32.dll" Alias "LeaveCriticalSection" (lpCriticalSection As Any) As Long 
'临界许可销毁
Declare Function DeleteCriticalSection Lib "kernel32.dll" Alias "DeleteCriticalSection" (lpCriticalSection As Any) As Long 
'
'信号量创建（跨进程）
Declare Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" (lpSemaphoreAttributes As Long, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As Long
'信号量打开
Declare Function OpenSemaphore Lib "kernel32" Alias "OpenSemaphoreA" (ByVal dwDesiredAccess As Long, ByVal 是否继承 As Boolean, ByVal lpName As String  ) As Long 
'信号量递增
Declare Function ReleaseSemaphore Lib "kernel32" Alias "ReleaseSemaphore" (ByVal 句柄 As Long, ByVal 增加的数量 As Long, 之前的数量 As Long) As Long
'信号量递减
Declare Function WaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
-------------------------*/
'
'互斥锁创建（跨进程）
Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (安全属性 As Long , ByVal 是否立即拥有 As Long, ByVal 互斥名称 As String) As Long
'互斥锁打开（跨进程）
Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal 权限 As Long, ByVal 是否立即拥有 As Boolean, ByVal 互斥名称 As String) As Long
'互斥锁进入
Declare Function WaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'互斥锁退出
Declare Function ReleaseMutex Lib "kernel32" Alias "ReleaseMutex" (ByVal hMutex As Long)
'
'创建一个 Event事件对象
Declare Function CreateEvent Lib  "kernel32.dll" Alias  "CreateEventW"(ByVal 安全性结构 As Any,ByVal 人工或自动事件 As Long,ByVal 是否内部触发 As Long,ByVal 事件对象名 As String) As Long
'等待并进入/拥有这个Event事件对象
Declare Function WaitForSingleObject Lib  "kernel32.dll" Alias  "WaitForSingleObject"(ByVal hObjcte As Any,ByVal Time As Long) As Long
'将Event事件对象设置为发信号状态，触发状态
Declare Function SetEvent Lib  "kernel32.dll" Alias  "SetEvent"(ByVal hObjcte As Any) As Long
'
'关闭一个内核对象。其中包括文件、文件映射、进程、线程、安全和同步对象等。
Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long

'用于获取自windows启动以来经历的时间长度（毫秒）
Declare Function GetTickCount Lib "kernel32" () As Long
'
'
'--------[互斥锁]--------
Function 互斥锁创建()
	Dim 锁标识, i
	锁标识 = "互斥锁_神梦无痕_QQ：1042207232_由【果兒】提供_"
	For i = 0 To 12 : Randomize :锁标识 = 锁标识 & Chr((24 * Rnd) + 65) :Next
    互斥锁创建 = CreateMutex(0, false, 锁标识)
End Function
Sub 互斥锁进入(锁句柄)
    Call WaitForSingleObject(锁句柄, 4294967295)
End Sub
Sub 互斥锁退出(锁句柄)
    Call ReleaseMutex(锁句柄)
End Sub
Sub 互斥锁销毁(锁句柄)
    Call CloseHandle(锁句柄)
End Sub
'
'--------[自旋锁]--------
Function 自旋锁创建()
    自旋锁创建 = GetThreadID() & Int(GetTickCount() * Rnd)
End Function
Sub 自旋锁进入(锁句柄)
    Dim 锁ID标识, 锁状态标识
    锁ID标识 = "自旋锁_神梦无痕_QQ：1042207232_由【a188652011】提供_线程ID_" & 锁句柄
    锁状态标识 = "自旋锁_神梦无痕_QQ：1042207232_由【a188652011】提供_状态_" & 锁句柄
    Do While GetEnv(锁ID标识) <> GetThreadID()
        While GetEnv(锁状态标识) = True
            '其他线程在操作了，就进入等待，自行调控检测时间
            '尽量让线程不同步,抢占情况能低一些    
            Delay Int((200 - 100 + 1) * Rnd + 100) '线程越多，延迟的范围其实要更广才合适。
        Wend
        '把其他线程拦在等待区域,不过有少部分的线程会到这一步
        SetEnv 锁状态标识, True
        '自己的线程ID 给主要的
        SetEnv 锁ID标识, GetThreadID()
        '等待其他线程也对线程变量写完
        '这里可以适当的延长一点，出错率更低
        Delay Int((20 - 10 + 1) * Rnd + 10)
    Loop
End Sub
Sub 自旋锁退出(锁句柄)
    Dim 锁ID标识, 锁状态标识
    锁ID标识 = "自旋锁_神梦无痕_QQ：1042207232_由【a188652011】提供_线程ID_" & 锁句柄
    锁状态标识 = "自旋锁_神梦无痕_QQ：1042207232_由【a188652011】提供_状态_" & 锁句柄
    SetEnv 锁ID标识, 0
    SetEnv 锁状态标识, False
End Sub
'
'--------[信号量]--------*按键精灵里不支持该命令*
//Function 信号量创建(并发上限)
//	Dim 信号标识, i, Ret
//	信号标识 = "信号量_神梦无痕_QQ：1042207232_"
//	For i = 0 To 12 : Randomize :信号标识 = 信号标识 & Chr((24 * Rnd) + 65) :Next
//	If IsNumeric(CStr(并发上限)) = False Or 并发上限 = "0" Then 并发上限 = 1
//	Ret = CreateSemaphore(0, 0, CLng(并发上限), 信号标识)
//	SetEnv Ret, 信号标识
//	信号量创建 = Ret
//End Function
//Function 信号量打开(信号句柄)
//	信号量打开 = OpenSemaphore(2031619, True, GetEnv(信号句柄))
//End Function
//Sub 信号量递减(信号句柄)
//	Call WaitForSingleObject(信号句柄, 4294967295)
//End Sub
//Sub 信号量递增(信号句柄)
//	返回递增前的值 = 0
//	Call ReleaseSemaphore(信号句柄, 1, CLng(返回递增前的值))
//End Sub
//Sub 信号量销毁(信号句柄)
//	Call CloseHandle(信号句柄)
//End Sub 
'
'--------[事件]--------
Function 事件创建()
	Dim 事件标识, i
	事件标识 = "事件_神梦无痕_QQ：1042207232_由【风__琪仙】提供_"
	For i = 0 To 12 : Randomize :事件标识 = 事件标识 & Chr((24 * Rnd) + 65) :Next
    事件创建 = CreateEvent(0, 0, 1, 事件标识)
End Function
Sub 事件进入(事件句柄)
    Call WaitForSingleObject(事件句柄, 4294967295)
End Sub
Sub 事件退出(事件句柄)
    Call SetEvent(事件句柄)
End Sub
Sub 事件销毁(事件句柄)
    Call CloseHandle(事件句柄)
End Sub
'
Sub A_______________________________________()
End Sub
Sub A【作者】：神梦无痕()
End Sub
Sub A【ＱＱ】：1042207232()
End Sub
Sub B________［需要脚本定制Q我］_____________()
End Sub
'
/*〓〓〓〓〓〓〓〓【更新历史】〓〓〓〓〓〓〓〓
神梦_多线程v1.0 2018.12.31
\
|-- 新增 互斥锁创建() 命令
|-- 新增 互斥锁进入() 命令
|-- 新增 互斥锁退出() 命令
|-- 新增 互斥锁销毁() 命令
|-- 新增 事件创建() 命令
|-- 新增 事件进入() 命令
|-- 新增 事件退出() 命令
|-- 新增 事件销毁() 命令
|-- 新增 自旋锁创建() 命令
|-- 新增 自旋锁进入() 命令
|-- 新增 自旋锁退出() 命令
|
|
〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓〓*/