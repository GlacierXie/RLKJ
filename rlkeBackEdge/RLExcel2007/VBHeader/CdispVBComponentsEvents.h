// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��


// CdispVBComponentsEvents ��װ��

class CdispVBComponentsEvents : public COleDispatchDriver
{
public:
	CdispVBComponentsEvents(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CdispVBComponentsEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CdispVBComponentsEvents(const CdispVBComponentsEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _dispVBComponentsEvents ����
public:
	void ItemAdded(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}
	void ItemRemoved(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}
	void ItemRenamed(LPDISPATCH VBComponent, LPCTSTR OldName)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR ;
		InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent, OldName);
	}
	void ItemSelected(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}
	void ItemActivated(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}
	void ItemReloaded(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}

	// _dispVBComponentsEvents ����
public:

};
