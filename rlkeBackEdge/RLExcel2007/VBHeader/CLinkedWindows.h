// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��


// CLinkedWindows ��װ��

class CLinkedWindows : public COleDispatchDriver
{
public:
	CLinkedWindows(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CLinkedWindows(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CLinkedWindows(const CLinkedWindows& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _LinkedWindows ����
public:
	LPDISPATCH get_VBE()
	{
		LPDISPATCH result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(VARIANT& index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &index);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0xc9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPUNKNOWN _NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_METHOD, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	void Remove(LPDISPATCH Window)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xca, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Window);
	}
	void Add(LPDISPATCH Window)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xcb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Window);
	}

	// _LinkedWindows ����
public:

};
