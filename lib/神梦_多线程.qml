[General]
SyntaxVersion=2
MacroID=652ab43d-d143-47d0-819d-61745b09cade
[Comment]

[Script]
'======================================[��Ҫ�ű����ƿ�������]======================================
'�����ζ��߳�����⡿
'�汾��v1.3
'���£�2019.01.18
'���ߣ������޺�
'�ѣѣ�1042207232
'��Ⱥ��624655641
'================================================================================================
'
'PC�˹һ���VPN�������̵�IP��������ϵQ��1042207232
'
'����QQ���ֻ����Ƿ�ͨ΢�ţ���ȷ�ȸߴ�99%��������ϵQ��1042207232
'
'================================================================================================
'
/*-------------------------���������ﲻ֧�������API
'ԭ�����ݼ�������̣�
Declare Function InterlockedDecrement Lib "kernel32.dll" Alias "InterlockedDecrement" (ByVal lpAddend As Long) As Long 
'ԭ��������
Declare Function InterlockedIncrement Lib "kernel32.dll" Alias "InterlockedIncrement" (ByVal lpAddend As Long) As Long 
'ԭ������ֵ
Declare Function InterlockedExchange Lib "kernel32.dll" Alias "InterlockedExchange" (Target As Long, ByVal Value As Long) As Long
'ԭ��������
Declare Function InterlockedExchangeAdd Lib "kernel32.dll" Alias "InterlockedExchangeAdd" (Addend As Long, ByVal Value As Long) As Long
'ԭ������Ŀ����
Declare Function InterlockedCompareExchange Lib "kernel32.dll" Alias "InterlockedExchangeAdd" (Destination As Long, ByVal Exchange As Long, ByVal Comperand As Long) As Long
'
'�ٽ���ɴ��������̣߳�
Declare Function InitializeCriticalSection Lib "kernel32.dll" Alias "InitializeCriticalSection" (lpCriticalSection As Any) As Long 
'�ٽ������Խ���
Declare Function TryEnterCriticalSection Lib "kernel32.dll" Alias "TryEnterCriticalSection" (lpCriticalSection As Any) As Boolean 
'�ٽ�������
Declare Function EnterCriticalSection Lib "kernel32.dll" Alias "EnterCriticalSection" (lpCriticalSection As Any) As Long 
'�ٽ����˳�
Declare Function LeaveCriticalSection Lib "kernel32.dll" Alias "LeaveCriticalSection" (lpCriticalSection As Any) As Long 
'�ٽ��������
Declare Function DeleteCriticalSection Lib "kernel32.dll" Alias "DeleteCriticalSection" (lpCriticalSection As Any) As Long 
'
'�ź�������������̣�
Declare Function CreateSemaphore Lib "kernel32" Alias "CreateSemaphoreA" (lpSemaphoreAttributes As Long, ByVal lInitialCount As Long, ByVal lMaximumCount As Long, ByVal lpName As String) As Long
'�ź�����
Declare Function OpenSemaphore Lib "kernel32" Alias "OpenSemaphoreA" (ByVal dwDesiredAccess As Long, ByVal �Ƿ�̳� As Boolean, ByVal lpName As String  ) As Long 
'�ź�������
Declare Function ReleaseSemaphore Lib "kernel32" Alias "ReleaseSemaphore" (ByVal ��� As Long, ByVal ���ӵ����� As Long, ֮ǰ������ As Long) As Long
'�ź����ݼ�
Declare Function WaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
-------------------------*/
'
'����������������̣�
Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (��ȫ���� As Long , ByVal �Ƿ�����ӵ�� As Long, ByVal �������� As String) As Long
'�������򿪣�����̣�
Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal Ȩ�� As Long, ByVal �Ƿ�����ӵ�� As Boolean, ByVal �������� As String) As Long
'����������
Declare Function WaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'�������˳�
Declare Function ReleaseMutex Lib "kernel32" Alias "ReleaseMutex" (ByVal hMutex As Long)
'
'
'����һ�� Event�¼����󣨿���̣�
Declare Function CreateEvent Lib  "kernel32.dll" Alias  "CreateEventW"(ByVal ��ȫ�Խṹ As Any,ByVal �˹����Զ��¼� As Long,ByVal �Ƿ��ڲ����� As Long,ByVal �¼������� As String) As Long
'�ȴ�������/ӵ�����Event�¼�����
Declare Function WaitForSingleObject Lib  "kernel32.dll" Alias  "WaitForSingleObject"(ByVal hObjcte As Any,ByVal Time As Long) As Long
'��Event�¼���������Ϊ���ź�״̬������״̬
Declare Function SetEvent Lib  "kernel32.dll" Alias  "SetEvent"(ByVal hObjcte As Any) As Long
'��Event�¼���������Ϊ���ź�״̬������/�Ǵ���״̬��
Declare Function ResetEvent Lib  "kernel32.dll" Alias  "ResetEvent"(ByVal hObjcte As Any) As Long
'
'
Declare Function CreateEventLong Lib  "kernel32.dll" Alias  "CreateEventW"(ByVal ��ȫ�Խṹ As Any,ByVal �˹����Զ��¼� As Long,ByVal �Ƿ��ڲ����� As Long,ByVal �¼������� As Long) As Long
'
'�ر�һ���ں˶������а����ļ����ļ�ӳ�䡢���̡��̡߳���ȫ��ͬ������ȡ�
Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long

'���ڻ�ȡ��windows��������������ʱ�䳤�ȣ����룩
Declare Function GetTickCount Lib "kernel32" () As Long

'��ȡ��ǰ�߳�һ��Ψһ���̱߳�ʶ��
Declare Function GetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long

'���ڴ�һ�������̶߳���
Declare Function OpenThread Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long

'�����߳� ��ͣ���� ������ ���ָ��� ������ִ��
Declare Function SuspendThread Lib "kernel32" Alias "SuspendThread" (ByVal hThread As Long) As Long

'�ָ��������б�������߳�
Declare Function ResumeThread Lib "kernel32" Alias "ResumeThread" (ByVal hThread As Long) As Long

'ǿ�ƽ����̡߳�(���Ƽ�ʹ��)
Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

'��ȡ�߳����ȼ�
Declare Function GetThreadPriority Lib "kernel32" Alias "GetThreadPriority" (ByVal hThread As Long) As Long

'�����߳����ȼ�
Declare Function SetThreadPriority Lib "kernel32" Alias "SetThreadPriority" (ByVal hThread As Long, ByVal nPriority As Long) As Long

'����CPU�׺���/��CPU
Declare Function SetProcessAffinityMask Lib "kernel32.dll" (ByVal hProcess As Long, ByVal dwProcessAffinityMask As Long) As Long
'
'
'--------------------------------[�������]--------------------------------
DimEnv DimEnv_Thread_Init, DimEnv_Thread_Tally, DimEnv_Thread_Data, DimEnv_Thread_NewTips, DimEnv_ԭ�Ӿ��
'
'--------------------------------[ԭ����]--------------------------------
Sub ԭ��_��ʼ��()
	If DimEnv_Thread_Init Then 
		DimEnv_ԭ�Ӿ�� = CreateEventLong(0, 0, 1, 0)
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "��������ִ�С�_��ʼ��()�������ʼ������⣡" : TracePrint Dim_Tips
	End If
End Sub
Sub ԭ��_����()
	Call CloseHandle(DimEnv_ԭ�Ӿ��)
End Sub
Sub ԭ��_����(������)
	Call WaitForSingleObject(DimEnv_ԭ�Ӿ��, 4294967295)
	SetEnv ������, GetEnv(������) + 1
	Call SetEvent(DimEnv_ԭ�Ӿ��)
End Sub
Sub ԭ��_�ݼ�(������)
	Call WaitForSingleObject(DimEnv_ԭ�Ӿ��, 4294967295)
	SetEnv ������, GetEnv(������) - 1
	Call SetEvent(DimEnv_ԭ�Ӿ��)
End Sub
Sub ԭ��_��ֵ(������, ��ֵ)
	Call WaitForSingleObject(DimEnv_ԭ�Ӿ��, 4294967295)
	SetEnv ������, ��ֵ
	Call SetEvent(DimEnv_ԭ�Ӿ��)
End Sub
Sub ԭ��_����(������, ����������)
	Dim Temp
	Call WaitForSingleObject(DimEnv_ԭ�Ӿ��, 4294967295)
	Temp = GetEnv(������)
	SetEnv ������, GetEnv(����������)
	SetEnv ����������, Temp
	Call SetEvent(DimEnv_ԭ�Ӿ��)
End Sub
Sub ԭ��_����(������, ��ֵ)
	Call WaitForSingleObject(DimEnv_ԭ�Ӿ��, 4294967295)
	SetEnv ������, GetEnv(������) + ��ֵ
	Call SetEvent(DimEnv_ԭ�Ӿ��)
End Sub
Sub ԭ��_��Ŀ����(������, ��ֵ, �Ա�ֵ)
	Call WaitForSingleObject(DimEnv_ԭ�Ӿ��, 4294967295)
	If GetEnv(������) = �Ա�ֵ Then 
		SetEnv ������, ��ֵ
	End If
	Call SetEvent(DimEnv_ԭ�Ӿ��)
End Sub
Sub ԭ��_�������(��ʾ)
	Call WaitForSingleObject(DimEnv_ԭ�Ӿ��, 4294967295)
	TracePrint ��ʾ
	Call SetEvent(DimEnv_ԭ�Ӿ��)
End Sub
'
'
'
'--------------------------------[������]--------------------------------
Function ����������()
	If DimEnv_Thread_Init Then 
		Dim ����ʶ, i
		����ʶ = "������_�����޺�_QQ��1042207232_�ɡ��������ṩ_"
		For i = 0 To 12 : Randomize :����ʶ = ����ʶ & Chr((24 * Rnd) + 65) :Next
    	���������� = CreateMutex(0, false, ����ʶ)
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "��������ִ�С�_��ʼ��()�������ʼ������⣡" : TracePrint Dim_Tips
	End If
End Function
Sub ����������(�����)
    Call WaitForSingleObject(�����, 4294967295)
End Sub
Sub �������˳�(�����)
    Call ReleaseMutex(�����)
End Sub
Sub ����������(�����)
    Call CloseHandle(�����)
End Sub
'
'--------------------------------[�ٽ���]--------------------------------
Function �ٽ�������()
	If DimEnv_Thread_Init Then 
		Dim ��ɱ�ʶ, i
		��ɱ�ʶ = "�ٽ����_�����޺�_QQ��1042207232_���֤_"
		For i = 0 To 12 : Randomize :��ɱ�ʶ = ��ɱ�ʶ & Chr((24 * Rnd) + 65) :Next
    	�ٽ������� = CreateEvent(0, 0, 1, ��ɱ�ʶ)
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "��������ִ�С�_��ʼ��()�������ʼ������⣡" : TracePrint Dim_Tips
	End If
End Function
Sub �ٽ�������(���֤)
    Call WaitForSingleObject(���֤, 4294967295)
End Sub
Sub �ٽ����˳�(���֤)
    Call SetEvent(���֤)
End Sub
Sub �ٽ�������(���֤)
    Call CloseHandle(���֤)
End Sub
'
'--------------------------------[�¼�]--------------------------------
Function �¼�����()
	If DimEnv_Thread_Init Then 
		Dim �¼���ʶ, i
		�¼���ʶ = "�¼�_�����޺�_QQ��1042207232_�ɡ���__���ɡ��ṩ_"
		For i = 0 To 12 : Randomize :�¼���ʶ = �¼���ʶ & Chr((24 * Rnd) + 65) :Next
    	�¼����� = CreateEvent(0, 0, 1, �¼���ʶ)
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "��������ִ�С�_��ʼ��()�������ʼ������⣡" : TracePrint Dim_Tips
	End If
End Function
Sub �¼�����(�¼����)
    Call WaitForSingleObject(�¼����, 4294967295)
End Sub
Sub �¼��˳�(�¼����)
    Call SetEvent(�¼����)
End Sub
Sub �¼�����(�¼����)
    Call CloseHandle(�¼����)
End Sub
'
'--------------------------------[������]--------------------------------
Function ����������()
	If DimEnv_Thread_Init Then 
        ���������� = GetThreadID() + Int(GetTickCount() * Rnd)
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "��������ִ�С�_��ʼ��()�������ʼ������⣡" : TracePrint Dim_Tips
	End If
End Function
Sub ����������(�����)
    Dim ��ID��ʶ, ��״̬��ʶ
    ��ID��ʶ = "������_�����޺�_QQ��1042207232_�ɡ�a188652011���ṩ_�߳�ID_" & �����
    ��״̬��ʶ = "������_�����޺�_QQ��1042207232_�ɡ�a188652011���ṩ_״̬_" & �����
    Do While GetEnv(��ID��ʶ) <> GetThreadID()
        While GetEnv(��״̬��ʶ) = True
            '�����߳��ڲ����ˣ��ͽ���ȴ������е��ؼ��ʱ��
            '�������̲߳�ͬ��,��ռ����ܵ�һЩ    
            Delay Int((200 - 100 + 1) * Rnd + 100) '�߳�Խ�࣬�ӳٵķ�Χ��ʵҪ����ź��ʡ�
        Wend
        '�������߳����ڵȴ�����,�������ٲ��ֵ��̻߳ᵽ��һ��
        SetEnv ��״̬��ʶ, True
        '�Լ����߳�ID ����Ҫ��
        SetEnv ��ID��ʶ, GetThreadID()
        '�ȴ������߳�Ҳ���̱߳���д��
        '��������ʵ����ӳ�һ�㣬�����ʸ���
        Delay Int((20 - 10 + 1) * Rnd + 10)
    Loop
End Sub
Sub �������˳�(�����)
    Dim ��ID��ʶ, ��״̬��ʶ
    ��ID��ʶ = "������_�����޺�_QQ��1042207232_�ɡ�a188652011���ṩ_�߳�ID_" & �����
    ��״̬��ʶ = "������_�����޺�_QQ��1042207232_�ɡ�a188652011���ṩ_״̬_" & �����
    SetEnv ��ID��ʶ, 0
    SetEnv ��״̬��ʶ, False
End Sub
'
'
'--------------------------------[�ź���]--------------------------------
'���ߣ������޺�
'�ѣѣ�1042207232
'��Ⱥ��624655641
'
/*���������������߳� --> �ճ�λ��
��һ��ͣ����������Ϊ���������������ͣ����ֻ��������λ���������ޡ��� 
һ��ʼ������λ���ǿյġ������̡߳�����ʱ���ͬʱ������������������������������ֱ�ӽ��롾�����߳�-3���� 
Ȼ����³������ź����ȴ�����ʣ�µĳ����������ڵȴ����˺����ĳ�Ҳ�����ò�����ڴ��ȴ��� 
��ʱ����һ�����뿪ͣ�����������߳�+1���������˵�֪�󣬴򿪳������ź����ͷš������������һ����ȥ�������߳�-1����
������뿪�����������߳�+2�������ֿ��Է��������������߳�-2������������� 

�����ͣ����ϵͳ�У���λ�ǹ�����Դ��ÿ�����ñ�һ���̣߳���������ľ����ź��������á�
*/
Function �ź�������(��������)
	If DimEnv_Thread_Init Then 
		Dim �źű�ʶ, i, Ret, ��ʶ_�����߳�, ��ʶ_��������
		�źű�ʶ = "�ź���_�����޺�_QQ��1042207232_"
		For i = 0 To 12 : Randomize :�źű�ʶ = �źű�ʶ & Chr((24 * Rnd) + 65) :Next
		If IsNumeric(CStr(��������)) = False Or �������� = "0" Then �������� = 1
		Ret = CreateEvent(0, 0, 1, �źű�ʶ)
		��ʶ_�����߳� = "�ź���_�����޺�_QQ��1042207232_�����߳�_" & Ret
		��ʶ_�������� = "�ź���_�����޺�_QQ��1042207232_��������_" & Ret
		SetEnv ��ʶ_�����߳�, ��������
		SetEnv ��ʶ_��������, ��������
		�ź������� = Ret
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "��������ִ�С�_��ʼ��()�������ʼ������⣡" : TracePrint Dim_Tips
	End If
End Function
Function �ź����Ƿ����(�źž��)
	Dim ��ʶ_�����߳�, ��ʶ_��������
	��ʶ_�����߳� = "�ź���_�����޺�_QQ��1042207232_�����߳�_" & �źž��
	��ʶ_�������� = "�ź���_�����޺�_QQ��1042207232_��������_" & �źž��
	�ź����Ƿ���� = (GetEnv(��ʶ_�����߳�) = GetEnv(��ʶ_��������))
End Function 
Sub �ź����ȴ�(�źž��)
	Dim ��ʶ_�����߳�
	��ʶ_�����߳� = "�ź���_�����޺�_QQ��1042207232_�����߳�_" & �źž��
	Call WaitForSingleObject(�źž��, 4294967295)
	While GetEnv(��ʶ_�����߳�) = 0
		'�ж��Ƿ��п����߳�
	Wend
	SetEnv ��ʶ_�����߳�, GetEnv(��ʶ_�����߳�) - 1
	If GetEnv(��ʶ_�����߳�) <> 0 Then Call SetEvent(�źž��)
	Dim_ThreadID = GetThreadID()
End Sub
Sub �ź����ͷ�(�źž��)
	Dim ��ʶ_�����߳�, ��ʶ_��������
	If Dim_ThreadID = GetThreadID() Then '�ж��Ƿ���ͬһ�߳���
		��ʶ_�����߳� = "�ź���_�����޺�_QQ��1042207232_�����߳�_" & �źž��
		��ʶ_�������� = "�ź���_�����޺�_QQ��1042207232_��������_" & �źž��
		If GetEnv(��ʶ_��������) > GetEnv(��ʶ_�����߳�) Then
			SetEnv ��ʶ_�����߳�, GetEnv(��ʶ_�����߳�) + 1
		End If
		Call SetEvent(�źž��)
	End If 
End Sub
Sub �ź�������(�źž��)
	Call CloseHandle(�źž��)
End Sub 
'
'--------------------------------[��д��]--------------------------------
'���ߣ������޺�
'�ѣѣ�1042207232
'��Ⱥ��624655641
'
/*��������
��д��ʵ����һ������������������ѶԹ�����Դ�ķ����߻��ֳɶ��ߺ�д�ߣ�����ֻ�Թ�����Դ���ж����ʣ�д������Ҫ�Թ�����Դ����д������ 
����������д���⡣

��������
���ںڰ�����д��Ϊ������ʦ��д�ߣ�ѧ���Ƕ��ţ������ʦ�ںڰ���дһ�����족�֣�
�����ʦû��д�꣬�Ͷ�ȡ�Ļ���������һ����һ���֡����ߡ������֣��ⲻ��������Ҫ�ģ� 
���Ա������ʦд�꣬ѧ���ٶ�ȡ�ڰ��ϵ����ݲ�����ȷ�ģ� 
��ʦ��ѧ��������û��ѧ��˵�����ˣ���ʦ�ͰѺڰ���ɾ��ˣ� 

ÿ��ѧ������һ���̣߳�����ͬʱ��ȡ������������ 
��ʦд��ʱ��ѧ���ȴ���ʦд���ٶ�ȡ��д�����⡿��
��ѧ���������ˣ���ʦ�Ż�Ѻڰ��߸ɾ�����д���⡿��
*/
Function ��д������()
	If DimEnv_Thread_Init Then 
		Dim ��ʶ_��ȡ��, ��ʶ_д����, Ret
    	Ret = CreateEventLong(0, 0, 1, 0)
		��ʶ_��ȡ�� = "��д��_�����޺�_QQ��1042207232_��ȡ��_" & Ret
		��ʶ_д���� = "��д��_�����޺�_QQ��1042207232_д����_" & Ret
		SetEnv ��ʶ_д����, CreateEvent(0, 1, 1, ��ʶ_д����)
		SetEnv ��ʶ_��ȡ��, CreateMutex(0, false, ��ʶ_��ȡ��)
		SetEnv Ret, 0
		��д������ = Ret
	Else
		If Not IsObject(Msg) Then Set Msg = CreateObject("QMPlugin.Msg") 
		Msg.Tips "��������ִ�С�_��ʼ��()�������ʼ������⣡" : TracePrint Dim_Tips
	End If
End Function
Sub ��д��������(��д���)
	Dim ��ʶ_��ȡ��, ��ʶ_д����, code, Ret
	Ret = False 
	��ʶ_��ȡ�� = "��д��_�����޺�_QQ��1042207232_��ȡ��_" & ��д���
	��ʶ_д���� = "��д��_�����޺�_QQ��1042207232_д����_" & ��д���
	code = WaitForSingleObject(GetEnv(��ʶ_��ȡ��), 4294967295)
	If code = 258 Then 
		Goto over
	End If
	Call WaitForSingleObject(��д���, 4294967295) 'ԭ��������
		SetEnv ��д���, GetEnv(��д���) + 1
		If GetEnv(��д���) = 1 Then 
			Call ResetEvent(GetEnv(��ʶ_д����))
		End If
	Call SetEvent(��д���) 'ԭ��������
	Call ReleaseMutex(GetEnv(��ʶ_��ȡ��))
	Ret = True 
	Rem over
	��д�������� = Ret
End Sub
Sub ��д��д����(��д���)
	Dim ��ʶ_��ȡ��, ��ʶ_д����, code, Ret
	Ret = False 
	��ʶ_��ȡ�� = "��д��_�����޺�_QQ��1042207232_��ȡ��_" & ��д���
	��ʶ_д���� = "��д��_�����޺�_QQ��1042207232_д����_" & ��д���
	code = WaitForSingleObject(GetEnv(��ʶ_��ȡ��), 4294967295)
	If code = 258 Then 
		Goto over
	End If
	code = WaitForSingleObject(GetEnv(��ʶ_д����), 4294967295)
	If code = 258 Then 
		Call ReleaseMutex(GetEnv(��ʶ_д����))
		Goto over
	End If
	Ret = True
	Rem over
	��д��д���� = Ret
End Sub
Sub ��д������(��д���)
	Dim ��ʶ_��ȡ��, ��ʶ_д����, code
	��ʶ_��ȡ�� = "��д��_�����޺�_QQ��1042207232_��ȡ��_" & ��д���
	��ʶ_д���� = "��д��_�����޺�_QQ��1042207232_д����_" & ��д���
	If ReleaseMutex(GetEnv(��ʶ_��ȡ��)) = 0 Then 
		Call WaitForSingleObject(��д���, 4294967295) 'ԭ��������
			SetEnv ��д���, GetEnv(��д���) - 1
			If GetEnv(��д���) = 0 Then 
				Call SetEvent(GetEnv(��ʶ_д����)) 
			End If
		Call SetEvent(��д���) 'ԭ��������
	End If
End Sub
Sub ��д������(��д���)
	Dim ��ʶ_��ȡ��, ��ʶ_д����
	��ʶ_��ȡ�� = "��д��_�����޺�_QQ��1042207232_��ȡ��_" & ��д���
	��ʶ_д���� = "��д��_�����޺�_QQ��1042207232_д����_" & ��д���
	Call CloseHandle(GetEnv(��ʶ_��ȡ��))
	Call CloseHandle(GetEnv(��ʶ_д����))
	Call CloseHandle(��д���)
End Sub 
'
'
'
'--------------------------------[�߳�]--------------------------------
'ͨ���߳�ID��ȡ�߳̾��
//Function �߳�_ȡ�߳̾��(�߳�ID)
//    �߳�_ȡ�߳̾�� = OpenThread(&H1F0FFF, 0, �߳�ID)
//End Function
'
'��ȡ��ǰ�߳̾��
Function �߳�_ȡ��ǰID()
    �߳�_ȡ��ǰ��� = GetThreadID()  //OpenThread(&H1F0FFF, 0, GetCurrentThreadId())
End Function
'
'ǿ�ƽ����߳�
Function �߳�_����(�߳�ID)
    �߳�_���� = StopThread(�߳�ID)  //TerminateThread(�߳̾��, 0)
End Function

'���=-15����=-2�����ڱ�׼=-1����׼=0�����ڱ�׼=1����=2�����=15
Sub �߳�_�����ȼ�(�߳�ID, ����)
    Call SetThreadPriority(OpenThread(&H1F0FFF, 0, CLng(�߳�ID)), CLng(����))
End Sub

Function �߳�_ȡ���ȼ�(�߳�ID)
    �߳�_ȡ���ȼ� = GetThreadPriority(OpenThread(&H1F0FFF, 0, CLng(�߳�ID)))
End Function

Function �߳�_�ָ�(�߳�ID)
    �߳�_�ָ� = ContinueThread(�߳�ID) //ResumeThread(�߳̾��)
End Function

Function �߳�_����(�߳�ID)
    �߳�_���� = PauseThread(�߳�ID) //SuspendThread(�߳̾��)
End Function

'0=�߳��ѽ���  1=�߳���������  -1=�߳̾����ʧЧ������
Function �߳�_ȡ״̬(�߳�ID)
	Dim Ret, �߳̾��
	Ret = -1
	�߳̾�� = OpenThread(&H1F0FFF, 0, CLng(�߳�ID))
	If IsNumeric(CStr(�߳̾��)) = False Or �߳̾�� = "0" Then Goto over
	Ret = WaitForSingleObject(�߳̾��, 0)
	If Ret = 258 Then 
		Ret = 1
	ElseIf Ret = -1 Then
	 	Ret = -1
	Else 
		Ret = 0
	End If
	Rem over
	�߳�_ȡ״̬ = Ret
End Function
'
'
Sub A_______________________________________()
End Sub
Sub A���ѣѡ���1042207232()
End Sub
Sub A�����ߡ��������޺�()
End Sub
Sub B________����Ҫ�ű�����Q�ң�_____________()
End Sub

'//��ʼ������⣬ʹ��ǰ������ø�����
'����ֵ���߼��ͣ��Ƿ�ɹ�
Function _��ʼ��()
    Dim Ret
    '-----------------------����ǰ�汾�š�-----------------------
    ��ǰ�汾 = "1.3"
    '-----------------------------------------------------------
    If DimEnv_Thread_Init = "" Then
    	Import "Msg.dll"
    	Import "Sys.dll"
    	Import "Window.dll"
    	If Window.Search("��������") <> "" Then
    		DimEnv_Thread_Tally = 0
    		DimEnv_Thread_NewTips = "������£�"
    		Execute _
    		"On Error Resume Next:" & _
    		"Set xmlHttp = CreateObject(""WinHttp.WinHttpRequest.5.1""):" & _
    		"xmlHttp.open ""GET"", ""http://www.smwh.online/Office/SoftwTally/SoftwTally.asp?SoftName=����_���߳�&Rem="& ��ǰ�汾 &"&SN="& Sys.GetHDDSN() &""", True:" & _
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
			"L1 = InStr(Text, ""�����桿""):L2 = InStr(Text, ""��/���桿""):" & _
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
			"If NewVer > """& ��ǰ�汾 &""" Then:" & _
			"	SetEnv ""DimEnv_Thread_NewTips"", ""�����µİ汾v"" & NewVer & ""�����Խ�Ⱥ���أ�""& Space(1024) & ""<img src='#' onerror='this.parentNode.style.color=""""#ff0000""""' style='display:none'>"":" & _
			"End If:SetEnv ""DimEnv_Thread_Init"", true"
    	End If 
    End If
    If Not IsNumeric(CStr(DimEnv_Thread_Tally)) Then DimEnv_Thread_Tally = 0
    TracePrint DimEnv_Thread_Data
    TracePrint "��������������_���̣߳�v"& ��ǰ�汾 &"��"
    TracePrint "�����ߡ��������޺�"
    TracePrint "���ѣѡ���1042207232"
    TracePrint "����Ⱥ����624655641"
    TracePrint "����վ����www.����.com"
	TracePrint "����������" & DimEnv_Thread_Tally & " ��"
	TracePrint "�����¡���" & DimEnv_Thread_NewTips
	Ret = True
	_��ʼ�� = Ret
End Function

/*������������������������ʷ������������������
����_���߳�v1.3 2019.01.18
\
|-- ���� ԭ��_�������() ����
|
|
����_���߳�v1.2 2019.01.17
\
|-- ���� _��ʼ��() ����
|
|
����_���߳�v1.1 2018.01.07
\
|-- ���� ��д������() ����
|-- ���� ��д��������() ����
|-- ���� ��д������() ����
|-- ���� ��д������() ����
|-- ���� ��д��д����() ����
|-- ���� �ٽ�������() ����
|-- ���� �ٽ�������() ����
|-- ���� �ٽ����˳�() ����
|-- ���� �ٽ�������() ����
|-- ���� �߳�_����() ����
|-- ���� �߳�_�ָ�() ����
|-- ���� �߳�_����() ����
|-- ���� �߳�_ȡ��ǰID() ����
|-- ���� �߳�_ȡ���ȼ�() ����
|-- ���� �߳�_ȡ״̬() ����
|-- ���� �߳�_�����ȼ�() ����
|-- ���� �ź�������() ����
|-- ���� �ź����ȴ�() ����
|-- ���� �ź����Ƿ����() ����
|-- ���� �ź����ͷ�() ����
|-- ���� �ź�������() ����
|-- ���� ԭ��_��ʼ��() ����
|-- ���� ԭ��_�ݼ�() ����
|-- ���� ԭ��_����() ����
|-- ���� ԭ��_��ֵ() ����
|-- ���� ԭ��_����() ����
|-- ���� ԭ��_��Ŀ����() ����
|-- ���� ԭ��_����() ����
|-- ���� ԭ��_����() ����
|
|
����_���߳�v1.0 2018.12.31
\
|-- ���� ����������() ����
|-- ���� ����������() ����
|-- ���� �������˳�() ����
|-- ���� ����������() ����
|-- ���� �¼�����() ����
|-- ���� �¼�����() ����
|-- ���� �¼��˳�() ����
|-- ���� �¼�����() ����
|-- ���� ����������() ����
|-- ���� ����������() ����
|-- ���� �������˳�() ����
|
|
��������������������������������������������*/