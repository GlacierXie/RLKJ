// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

// CdispReferences_Events ��װ��

class CdispReferences_Events : public COleDispatchDriver
{
public:
	CdispReferences_Events(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CdispReferences_Events(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CdispReferences_Events(const CdispReferences_Events& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _dispReferences_Events ����
public:
	void ItemAdded(LPDISPATCH Reference)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Reference);
	}
	void ItemRemoved(LPDISPATCH Reference)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Reference);
	}

	// _dispReferences_Events ����
public:

};
