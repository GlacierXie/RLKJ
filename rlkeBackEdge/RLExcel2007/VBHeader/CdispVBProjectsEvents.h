// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

// CdispVBProjectsEvents ��װ��

class CdispVBProjectsEvents : public COleDispatchDriver
{
public:
	CdispVBProjectsEvents(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CdispVBProjectsEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CdispVBProjectsEvents(const CdispVBProjectsEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _dispVBProjectsEvents ����
public:
	void ItemAdded(LPDISPATCH VBProject)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBProject);
	}
	void ItemRemoved(LPDISPATCH VBProject)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBProject);
	}
	void ItemRenamed(LPDISPATCH VBProject, LPCTSTR OldName)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR ;
		InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBProject, OldName);
	}
	void ItemActivated(LPDISPATCH VBProject)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBProject);
	}

	// _dispVBProjectsEvents ����
public:

};
