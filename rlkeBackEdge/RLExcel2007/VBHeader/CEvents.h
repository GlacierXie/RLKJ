// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

// CEvents ��װ��

class CEvents : public COleDispatchDriver
{
public:
	CEvents(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CEvents(const CEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// Events ����
public:
	LPUNKNOWN get_ReferencesEvents(LPDISPATCH VBProject)
	{
		LPUNKNOWN result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xca, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, parms, VBProject);
		return result;
	}
	LPUNKNOWN get_CommandBarEvents(LPDISPATCH CommandBarControl)
	{
		LPUNKNOWN result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xcd, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, parms, CommandBarControl);
		return result;
	}

	// Events ����
public:

};
