// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

// CdispReferencesEvents ��װ��

class CdispReferencesEvents : public COleDispatchDriver
{
public:
	CdispReferencesEvents(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CdispReferencesEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CdispReferencesEvents(const CdispReferencesEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _dispReferencesEvents ����
public:
	void ItemAdded(LPDISPATCH Reference)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Reference);
	}
	void ItemRemoved(LPDISPATCH Reference)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Reference);
	}

	// _dispReferencesEvents ����
public:

};
