[General]
SyntaxVersion=2
MacroID=652ab43d-d143-47d0-819d-61745b09cade
[Comment]

[Script]
'======================================[��Ҫ�ű����ƿ�������]======================================
'�����ζ��߳�����⡿
'�汾��v1.0
'���£�2018.12.31
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
'����һ�� Event�¼�����
Declare Function CreateEvent Lib  "kernel32.dll" Alias  "CreateEventW"(ByVal ��ȫ�Խṹ As Any,ByVal �˹����Զ��¼� As Long,ByVal �Ƿ��ڲ����� As Long,ByVal �¼������� As String) As Long
'�ȴ�������/ӵ�����Event�¼�����
Declare Function WaitForSingleObject Lib  "kernel32.dll" Alias  "WaitForSingleObject"(ByVal hObjcte As Any,ByVal Time As Long) As Long
'��Event�¼���������Ϊ���ź�״̬������״̬
Declare Function SetEvent Lib  "kernel32.dll" Alias  "SetEvent"(ByVal hObjcte As Any) As Long
'
'�ر�һ���ں˶������а����ļ����ļ�ӳ�䡢���̡��̡߳���ȫ��ͬ������ȡ�
Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long

'���ڻ�ȡ��windows��������������ʱ�䳤�ȣ����룩
Declare Function GetTickCount Lib "kernel32" () As Long
'
'
'--------[������]--------
Function ����������()
	Dim ����ʶ, i
	����ʶ = "������_�����޺�_QQ��1042207232_�ɡ��������ṩ_"
	For i = 0 To 12 : Randomize :����ʶ = ����ʶ & Chr((24 * Rnd) + 65) :Next
    ���������� = CreateMutex(0, false, ����ʶ)
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
'--------[������]--------
Function ����������()
    ���������� = GetThreadID() & Int(GetTickCount() * Rnd)
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
'--------[�ź���]--------*���������ﲻ֧�ָ�����*
//Function �ź�������(��������)
//	Dim �źű�ʶ, i, Ret
//	�źű�ʶ = "�ź���_�����޺�_QQ��1042207232_"
//	For i = 0 To 12 : Randomize :�źű�ʶ = �źű�ʶ & Chr((24 * Rnd) + 65) :Next
//	If IsNumeric(CStr(��������)) = False Or �������� = "0" Then �������� = 1
//	Ret = CreateSemaphore(0, 0, CLng(��������), �źű�ʶ)
//	SetEnv Ret, �źű�ʶ
//	�ź������� = Ret
//End Function
//Function �ź�����(�źž��)
//	�ź����� = OpenSemaphore(2031619, True, GetEnv(�źž��))
//End Function
//Sub �ź����ݼ�(�źž��)
//	Call WaitForSingleObject(�źž��, 4294967295)
//End Sub
//Sub �ź�������(�źž��)
//	���ص���ǰ��ֵ = 0
//	Call ReleaseSemaphore(�źž��, 1, CLng(���ص���ǰ��ֵ))
//End Sub
//Sub �ź�������(�źž��)
//	Call CloseHandle(�źž��)
//End Sub 
'
'--------[�¼�]--------
Function �¼�����()
	Dim �¼���ʶ, i
	�¼���ʶ = "�¼�_�����޺�_QQ��1042207232_�ɡ���__���ɡ��ṩ_"
	For i = 0 To 12 : Randomize :�¼���ʶ = �¼���ʶ & Chr((24 * Rnd) + 65) :Next
    �¼����� = CreateEvent(0, 0, 1, �¼���ʶ)
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
Sub A_______________________________________()
End Sub
Sub A�����ߡ��������޺�()
End Sub
Sub A���ѣѡ���1042207232()
End Sub
Sub B________����Ҫ�ű�����Q�ң�_____________()
End Sub
'
/*������������������������ʷ������������������
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