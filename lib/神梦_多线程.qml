[General]
SyntaxVersion=2
MacroID=652ab43d-d143-47d0-819d-61745b09cade
[Comment]

[Script]
'======================================[需要脚本定制可以找我]======================================
'【神梦多线程命令库】
'版本：v1.3
'更新：2019.01.18
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
'
'创建一个 Event事件对象（跨进程）
Declare Function CreateEvent Lib  "kernel32.dll" Alias  "CreateEventW"(ByVal 安全性结构 As Any,ByVal 人工或自动事件 As Long,ByVal 是否内部触发 As Long,ByVal 事件对象名 As String) As Long
'等待并进入/拥有这个Event事件对象
Declare Function WaitForSingleObject Lib  "kernel32.dll" Alias  "WaitForSingleObject"(ByVal hObjcte As Any,ByVal Time As Long) As Long
'将Event事件对象设置为发信号状态，触发状态
Declare Function SetEvent Lib  "kernel32.dll" Alias  "SetEvent"(ByVal hObjcte As Any) As Long
'将Event事件对象设置为无信号状态，重置/非触发状态。
Declare Function ResetEvent Lib  "kernel32.dll" Alias  "ResetEvent"(ByVal hObjcte As Any) As Long
'
'
Declare Function CreateEventLong Lib  "kernel32.dll" Alias  "CreateEventW"(ByVal 安全性结构 As Any,ByVal 人工或自动事件 As Long,ByVal 是否内部触发 As Long,ByVal 事件对象名 As Long) As Long
'
'关闭一个内核对象。其中包括文件、文件映射、进程、线程、安全和同步对象等。
Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long

'用于获取自windows启动以来经历的时间长度（毫秒）
Declare Function GetTickCount Lib "kernel32" () As Long

'获取当前线程一个唯一的线程标识符
Declare Function GetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long

'用于打开一个现有线程对象
Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long

'挂起线程 暂停操作 可以用 “恢复” 来继续执行
Declare Function SuspendThread Lib "kernel32" Alias "SuspendThread" (ByVal hThread As Long) As Long

'恢复继续运行被挂起的线程
Declare Function ResumeThread Lib "kernel32" Alias "ResumeThread" (ByVal hThread As Long) As Long

'强制结束线程。(不推荐使用)
Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

'获取线程优先级
Declare Function GetThreadPriority Lib "kernel32" Alias "GetThreadPriority" (ByVal hThread As Long) As Long

'设置线程优先级
Declare Function SetThreadPriority Lib "kernel32" Alias "SetThreadPriority" (ByVal hThread As Long, ByVal nPriority As Long) As Long

'设置CPU亲和性/绑定CPU
Declare Function SetProcessAffinityMask Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwProcessAffinityMask As Long) As Long
'
'
'--------------------------------[定义变量]--------------------------------
DimEnv DimEnv_Thread_Init, DimEnv_Thread_Tally, DimEnv_Thread_Data, DimEnv_Thread_NewTips, DimEnv_原子句柄
'
'--------------------------------[原子锁]--------------------------------
Sub 原子_初始化()
	If DimEnv_Thread_Init Then 
		DimEnv_原子句柄 = CreateEventLong(0, 0, 1, 0)
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "出错，请先执行【_初始化()】命令，初始化命令库！" : TracePrint Dim_Tips
	End If
End Sub
Sub 原子_销毁()
	Call CloseHandle(DimEnv_原子句柄)
End Sub
Sub 原子_递增(变量名)
	Call WaitForSingleObject(DimEnv_原子句柄, 4294967295)
	SetEnv 变量名, GetEnv(变量名) + 1
	Call SetEvent(DimEnv_原子句柄)
End Sub
Sub 原子_递减(变量名)
	Call WaitForSingleObject(DimEnv_原子句柄, 4294967295)
	SetEnv 变量名, GetEnv(变量名) - 1
	Call SetEvent(DimEnv_原子句柄)
End Sub
Sub 原子_赋值(变量名, 数值)
	Call WaitForSingleObject(DimEnv_原子句柄, 4294967295)
	SetEnv 变量名, 数值
	Call SetEvent(DimEnv_原子句柄)
End Sub
Sub 原子_交换(变量名, 交换变量名)
	Dim Temp
	Call WaitForSingleObject(DimEnv_原子句柄, 4294967295)
	Temp = GetEnv(变量名)
	SetEnv 变量名, GetEnv(交换变量名)
	SetEnv 交换变量名, Temp
	Call SetEvent(DimEnv_原子句柄)
End Sub
Sub 原子_运算(变量名, 数值)
	Call WaitForSingleObject(DimEnv_原子句柄, 4294967295)
	SetEnv 变量名, GetEnv(变量名) + 数值
	Call SetEvent(DimEnv_原子句柄)
End Sub
Sub 原子_三目运算(变量名, 赋值, 对比值)
	Call WaitForSingleObject(DimEnv_原子句柄, 4294967295)
	If GetEnv(变量名) = 对比值 Then 
		SetEnv 变量名, 赋值
	End If
	Call SetEvent(DimEnv_原子句柄)
End Sub
Sub 原子_调试输出(提示)
	Call WaitForSingleObject(DimEnv_原子句柄, 4294967295)
	TracePrint 提示
	Call SetEvent(DimEnv_原子句柄)
End Sub
'
'
'
'--------------------------------[互斥锁]--------------------------------
Function 互斥锁创建()
	If DimEnv_Thread_Init Then 
		Dim 锁标识, i
		锁标识 = "互斥锁_神梦无痕_QQ：1042207232_由【果兒】提供_"
		For i = 0 To 12 : Randomize :锁标识 = 锁标识 & Chr((24 * Rnd) + 65) :Next
    	互斥锁创建 = CreateMutex(0, false, 锁标识)
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "出错，请先执行【_初始化()】命令，初始化命令库！" : TracePrint Dim_Tips
	End If
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
'--------------------------------[临界区]--------------------------------
Function 临界区创建()
	If DimEnv_Thread_Init Then 
		Dim 许可标识, i
		许可标识 = "临界许可_神梦无痕_QQ：1042207232_许可证_"
		For i = 0 To 12 : Randomize :许可标识 = 许可标识 & Chr((24 * Rnd) + 65) :Next
    	临界区创建 = CreateEvent(0, 0, 1, 许可标识)
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "出错，请先执行【_初始化()】命令，初始化命令库！" : TracePrint Dim_Tips
	End If
End Function
Sub 临界区进入(许可证)
    Call WaitForSingleObject(许可证, 4294967295)
End Sub
Sub 临界区退出(许可证)
    Call SetEvent(许可证)
End Sub
Sub 临界区销毁(许可证)
    Call CloseHandle(许可证)
End Sub
'
'--------------------------------[事件]--------------------------------
Function 事件创建()
	If DimEnv_Thread_Init Then 
		Dim 事件标识, i
		事件标识 = "事件_神梦无痕_QQ：1042207232_由【风__琪仙】提供_"
		For i = 0 To 12 : Randomize :事件标识 = 事件标识 & Chr((24 * Rnd) + 65) :Next
    	事件创建 = CreateEvent(0, 0, 1, 事件标识)
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "出错，请先执行【_初始化()】命令，初始化命令库！" : TracePrint Dim_Tips
	End If
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
'--------------------------------[自旋锁]--------------------------------
Function 自旋锁创建()
	If DimEnv_Thread_Init Then 
        自旋锁创建 = GetThreadID() + Int(GetTickCount() * Rnd)
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "出错，请先执行【_初始化()】命令，初始化命令库！" : TracePrint Dim_Tips
	End If
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
'
'--------------------------------[信号量]--------------------------------
'作者：神梦无痕
'ＱＱ：1042207232
'Ｑ群：624655641
'
/*【描述】（空闲线程 --> 空车位）
以一个停车场的运作为例。简单起见，假设停车场只有三个车位【并发上限】， 
一开始三个车位都是空的【空闲线程】。这时如果同时来了五辆车，看门人允许其中三辆直接进入【空闲线程-3】， 
然后放下车拦【信号量等待】，剩下的车则必须在入口等待，此后来的车也都不得不在入口处等待。 
这时，有一辆车离开停车场【空闲线程+1】，看门人得知后，打开车拦【信号量释放】，放入外面的一辆进去【空闲线程-1】，
如果又离开两辆【空闲线程+2】，则又可以放入两辆【空闲线程-2】，如此往复。 

在这个停车场系统中，车位是公共资源，每辆车好比一个线程，看门人起的就是信号量的作用。
*/
Function 信号量创建(并发上限)
	If DimEnv_Thread_Init Then 
		Dim 信号标识, i, Ret, 标识_空闲线程, 标识_并发上限
		信号标识 = "信号量_神梦无痕_QQ：1042207232_"
		For i = 0 To 12 : Randomize :信号标识 = 信号标识 & Chr((24 * Rnd) + 65) :Next
		If IsNumeric(CStr(并发上限)) = False Or 并发上限 = "0" Then 并发上限 = 1
		Ret = CreateEvent(0, 0, 1, 信号标识)
		标识_空闲线程 = "信号量_神梦无痕_QQ：1042207232_空闲线程_" & Ret
		标识_并发上限 = "信号量_神梦无痕_QQ：1042207232_并发上限_" & Ret
		SetEnv 标识_空闲线程, 并发上限
		SetEnv 标识_并发上限, 并发上限
		信号量创建 = Ret
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "出错，请先执行【_初始化()】命令，初始化命令库！" : TracePrint Dim_Tips
	End If
End Function
Function 信号量是否空闲(信号句柄)
	Dim 标识_空闲线程, 标识_并发上限
	标识_空闲线程 = "信号量_神梦无痕_QQ：1042207232_空闲线程_" & 信号句柄
	标识_并发上限 = "信号量_神梦无痕_QQ：1042207232_并发上限_" & 信号句柄
	信号量是否空闲 = (GetEnv(标识_空闲线程) = GetEnv(标识_并发上限))
End Function 
Sub 信号量等待(信号句柄)
	Dim 标识_空闲线程
	标识_空闲线程 = "信号量_神梦无痕_QQ：1042207232_空闲线程_" & 信号句柄
	Call WaitForSingleObject(信号句柄, 4294967295)
	While GetEnv(标识_空闲线程) = 0
		'判断是否有空闲线程
	Wend
	SetEnv 标识_空闲线程, GetEnv(标识_空闲线程) - 1
	If GetEnv(标识_空闲线程) <> 0 Then Call SetEvent(信号句柄)
	Dim_ThreadID = GetThreadID()
End Sub
Sub 信号量释放(信号句柄)
	Dim 标识_空闲线程, 标识_并发上限
	If Dim_ThreadID = GetThreadID() Then '判断是否在同一线程内
		标识_空闲线程 = "信号量_神梦无痕_QQ：1042207232_空闲线程_" & 信号句柄
		标识_并发上限 = "信号量_神梦无痕_QQ：1042207232_并发上限_" & 信号句柄
		If GetEnv(标识_并发上限) > GetEnv(标识_空闲线程) Then
			SetEnv 标识_空闲线程, GetEnv(标识_空闲线程) + 1
		End If
		Call SetEvent(信号句柄)
	End If 
End Sub
Sub 信号量销毁(信号句柄)
	Call CloseHandle(信号句柄)
End Sub 
'
'--------------------------------[读写锁]--------------------------------
'作者：神梦无痕
'ＱＱ：1042207232
'Ｑ群：624655641
'
/*【描述】
读写锁实际是一种特殊的自旋锁，它把对共享资源的访问者划分成读者和写者，读者只对共享资源进行读访问，写者则需要对共享资源进行写操作。 
读读共享，读写互斥。

【举例】
以在黑板上上写字为例，老师是写者，学生是读着，如果老师在黑板上写一个“天”字，
如果老师没有写完，就读取的话，看到是一个“一”字、或者“二”字，这不是我们想要的， 
所以必须等老师写完，学生再读取黑板上的内容才是正确的！ 
老师问学生看完了没，学生说看完了，老师就把黑板擦干净了！ 

每个学生看做一个线程，可以同时读取【读读共享】， 
老师写的时候，学生等待老师写完再读取【写读互斥】。
当学生都读完了，老师才会把黑板檫干净【读写互斥】。
*/
Function 读写锁创建()
	If DimEnv_Thread_Init Then 
		Dim 标识_读取锁, 标识_写入锁, Ret
    	Ret = CreateEventLong(0, 0, 1, 0)
		标识_读取锁 = "读写锁_神梦无痕_QQ：1042207232_读取锁_" & Ret
		标识_写入锁 = "读写锁_神梦无痕_QQ：1042207232_写入锁_" & Ret
		SetEnv 标识_写入锁, CreateEvent(0, 1, 1, 标识_写入锁)
		SetEnv 标识_读取锁, CreateMutex(0, false, 标识_读取锁)
		SetEnv Ret, 0
		读写锁创建 = Ret
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "出错，请先执行【_初始化()】命令，初始化命令库！" : TracePrint Dim_Tips
	End If
End Function
Sub 读写锁读锁定(读写句柄)
	Dim 标识_读取锁, 标识_写入锁, code, Ret
	Ret = False 
	标识_读取锁 = "读写锁_神梦无痕_QQ：1042207232_读取锁_" & 读写句柄
	标识_写入锁 = "读写锁_神梦无痕_QQ：1042207232_写入锁_" & 读写句柄
	code = WaitForSingleObject(GetEnv(标识_读取锁), 4294967295)
	If code = 258 Then 
		Goto over
	End If
	Call WaitForSingleObject(读写句柄, 4294967295) '原子锁锁定
		SetEnv 读写句柄, GetEnv(读写句柄) + 1
		If GetEnv(读写句柄) = 1 Then 
			Call ResetEvent(GetEnv(标识_写入锁))
		End If
	Call SetEvent(读写句柄) '原子锁解锁
	Call ReleaseMutex(GetEnv(标识_读取锁))
	Ret = True 
	Rem over
	读写锁读锁定 = Ret
End Sub
Sub 读写锁写锁定(读写句柄)
	Dim 标识_读取锁, 标识_写入锁, code, Ret
	Ret = False 
	标识_读取锁 = "读写锁_神梦无痕_QQ：1042207232_读取锁_" & 读写句柄
	标识_写入锁 = "读写锁_神梦无痕_QQ：1042207232_写入锁_" & 读写句柄
	code = WaitForSingleObject(GetEnv(标识_读取锁), 4294967295)
	If code = 258 Then 
		Goto over
	End If
	code = WaitForSingleObject(GetEnv(标识_写入锁), 4294967295)
	If code = 258 Then 
		Call ReleaseMutex(GetEnv(标识_写入锁))
		Goto over
	End If
	Ret = True
	Rem over
	读写锁写锁定 = Ret
End Sub
Sub 读写锁解锁(读写句柄)
	Dim 标识_读取锁, 标识_写入锁, code
	标识_读取锁 = "读写锁_神梦无痕_QQ：1042207232_读取锁_" & 读写句柄
	标识_写入锁 = "读写锁_神梦无痕_QQ：1042207232_写入锁_" & 读写句柄
	If ReleaseMutex(GetEnv(标识_读取锁)) = 0 Then 
		Call WaitForSingleObject(读写句柄, 4294967295) '原子锁锁定
			SetEnv 读写句柄, GetEnv(读写句柄) - 1
			If GetEnv(读写句柄) = 0 Then 
				Call SetEvent(GetEnv(标识_写入锁)) 
			End If
		Call SetEvent(读写句柄) '原子锁解锁
	End If
End Sub
Sub 读写锁销毁(读写句柄)
	Dim 标识_读取锁, 标识_写入锁
	标识_读取锁 = "读写锁_神梦无痕_QQ：1042207232_读取锁_" & 读写句柄
	标识_写入锁 = "读写锁_神梦无痕_QQ：1042207232_写入锁_" & 读写句柄
	Call CloseHandle(GetEnv(标识_读取锁))
	Call CloseHandle(GetEnv(标识_写入锁))
	Call CloseHandle(读写句柄)
End Sub 
'
'
'
'--------------------------------[线程]--------------------------------
'通过线程ID获取线程句柄
//Function 线程_取线程句柄(线程ID)
//    线程_取线程句柄 = OpenThread(&H1F0FFF, 0, 线程ID)
//End Function
'
'获取当前线程句柄
Function 线程_取当前ID()
    线程_取当前句柄 = GetThreadID()  //OpenThread(&H1F0FFF, 0, GetCurrentThreadId())
End Function
'
'强制结束线程
Function 线程_结束(线程ID)
    线程_结束 = StopThread(线程ID)  //TerminateThread(线程句柄, 0)
End Function

'最低=-15；低=-2；低于标准=-1；标准=0；高于标准=1；高=2；最高=15
Sub 线程_置优先级(线程ID, 级别)
    Call SetThreadPriority(OpenThread(&H1F0FFF, 0, CLng(线程ID)), CLng(级别))
End Sub

Function 线程_取优先级(线程ID)
    线程_取优先级 = GetThreadPriority(OpenThread(&H1F0FFF, 0, CLng(线程ID)))
End Function

Function 线程_恢复(线程ID)
    线程_恢复 = ContinueThread(线程ID) //ResumeThread(线程句柄)
End Function

Function 线程_挂起(线程ID)
    线程_挂起 = PauseThread(线程ID) //SuspendThread(线程句柄)
End Function

'0=线程已结束  1=线程正在运行  -1=线程句柄已失效或销毁
Function 线程_取状态(线程ID)
	Dim Ret, 线程句柄
	Ret = -1
	线程句柄 = OpenThread(&H1F0FFF, 0, CLng(线程ID))
	If IsNumeric(CStr(线程句柄)) = False Or 线程句柄 = "0" Then Goto over
	Ret = WaitForSingleObject(线程句柄, 0)
	If Ret = 258 Then 
		Ret = 1
	ElseIf Ret = -1 Then
	 	Ret = -1
	Else 
		Ret = 0
	End If
	Rem over
	线程_取状态 = Ret
End Function
'
'
Sub A_______________________________________()
End Sub
Sub A【ＱＱ】：1042207232()
End Sub
Sub A【作者】：神梦无痕()
End Sub
Sub B________［需要脚本定制Q我］_____________()
End Sub

'//初始化命令库，使用前必须调用该命令
'返回值：逻辑型，是否成功
Function _初始化()
    Dim Ret
    '-----------------------【当前版本号】-----------------------
    当前版本 = "1.3"
    '-----------------------------------------------------------
    If DimEnv_Thread_Init = "" Then
    	Import "Msg.dll"
    	Import "Sys.dll"
    	Import "Window.dll"
    	If Window.Search("按键精灵") <> "" Then
    		DimEnv_Thread_Tally = 0
    		DimEnv_Thread_NewTips = "无需更新！"
    		Execute _
    		"On Error Resume Next:" & _
    		"Set xmlHttp = CreateObject(""WinHttp.WinHttpRequest.5.1""):" & _
    		"xmlHttp.open ""GET"", ""http://www.smwh.online/Office/SoftwTally/SoftwTally.asp?SoftName=神梦_多线程&Rem="& 当前版本 &"&SN="& Sys.GetHDDSN() &""", True:" & _
    		"xmlHttp.send:" & _
    		"If xmlHttp.waitForResponse(1) Then:" & _
    		"    If xmlHttp.statusText = ""OK"" Then:" & _
    		"        SetEnv ""DimEnv_Thread_Tally"", xmlHttp.responseText:" & _
    		"    End If:" & _
    		"End If:" & _
			"NewVer=0:Set RepEx=New RegExp:RepEx.IgnoreCase=True:RepEx.Global=True:RepEx.Pattern=""\{v(.*?)\}"":" & _
			"xmlHttp.open ""GET"", ""https://www.jianshu.com/p/84cd94b647ad"", True:" & _
			"xmlHttp.send:" & _
			"If xmlHttp.waitForResponse(1) Then:" & _
			"    If xmlHttp.statusText = ""OK"" Then:" & _
			"        Text = xmlHttp.responseText:" & _
			"    End If:" & _
			"End If:" & _
			"L1 = InStr(Text, ""【公告】""):L2 = InStr(Text, ""【/公告】""):" & _
			"If L1 > 0 And L2 > 0 Then:" & _
			"   Arr = Split(Mid(Text, L1+4, L2-L1-4) & vbCrLf & UnEscape(""%u5E7F%u544A%u4F4D%u51FA%u79DF%uFF0C%u8054%u7CFBQQ%uFF1A1042207232""), vbCrLf):" & _
			"   if UBound(Arr)>-1 Then:" & _
			"       Randomize:n = CInt((UBound(Arr)+1)*Rnd+1)-1:" & _
			"	    SetEnv ""DimEnv_Thread_Data"", ""<!--"" & Arr(n) & ""-->"" & Space(1024) & ""<span style='color:FF00FF'>"" & Arr(n) & ""</span>"":" & _
			"   End If:" & _
			"End If:" & _
			"If RepEx.Test(Text) Then:" & _
			"    NewVer = RepEx.Execute(Text).Item(0).SubMatches.Item(0):" & _
			"    If Not IsNumeric(Replace(NewVer, ""."", """")) Then:" & _
			"        NewVer = 0:" & _
			"    End If:" & _
			"End If:" & _
			"If NewVer > """& 当前版本 &""" Then:" & _
			"	SetEnv ""DimEnv_Thread_NewTips"", ""发现新的版本v"" & NewVer & ""，可以进群下载！""& Space(1024) & ""<img src='#' onerror='this.parentNode.style.color=""""#ff0000""""' style='display:none'>"":" & _
			"End If:SetEnv ""DimEnv_Thread_Init"", true"
    	End If 
    End If
    If Not IsNumeric(CStr(DimEnv_Thread_Tally)) Then DimEnv_Thread_Tally = 0
    TracePrint DimEnv_Thread_Data
    TracePrint "【库名】：神梦_多线程（v"& 当前版本 &"）"
    TracePrint "【作者】：神梦无痕"
    TracePrint "【ＱＱ】：1042207232"
    TracePrint "【Ｑ群】：624655641"
    TracePrint "【网站】：www.神梦.com"
	TracePrint "【人数】：" & DimEnv_Thread_Tally & " 人"
	TracePrint "【更新】：" & DimEnv_Thread_NewTips
	Ret = True
	_初始化 = Ret
End Function

/*〓〓〓〓〓〓〓〓【更新历史】〓〓〓〓〓〓〓〓
神梦_多线程v1.3 2019.01.18
\
|-- 新增 原子_调试输出() 命令
|
|
神梦_多线程v1.2 2019.01.17
\
|-- 新增 _初始化() 命令
|
|
神梦_多线程v1.1 2018.01.07
\
|-- 新增 读写锁创建() 命令
|-- 新增 读写锁读锁定() 命令
|-- 新增 读写锁解锁() 命令
|-- 新增 读写锁销毁() 命令
|-- 新增 读写锁写锁定() 命令
|-- 新增 临界区创建() 命令
|-- 新增 临界区进入() 命令
|-- 新增 临界区退出() 命令
|-- 新增 临界区销毁() 命令
|-- 新增 线程_挂起() 命令
|-- 新增 线程_恢复() 命令
|-- 新增 线程_结束() 命令
|-- 新增 线程_取当前ID() 命令
|-- 新增 线程_取优先级() 命令
|-- 新增 线程_取状态() 命令
|-- 新增 线程_置优先级() 命令
|-- 新增 信号量创建() 命令
|-- 新增 信号量等待() 命令
|-- 新增 信号量是否空闲() 命令
|-- 新增 信号量释放() 命令
|-- 新增 信号量销毁() 命令
|-- 新增 原子_初始化() 命令
|-- 新增 原子_递减() 命令
|-- 新增 原子_递增() 命令
|-- 新增 原子_赋值() 命令
|-- 新增 原子_交换() 命令
|-- 新增 原子_三目运算() 命令
|-- 新增 原子_运算() 命令
|-- 新增 原子_销毁() 命令
|
|
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