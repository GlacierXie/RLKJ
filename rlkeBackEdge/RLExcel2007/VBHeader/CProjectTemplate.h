// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

// CProjectTemplate ��װ��

class CProjectTemplate : public COleDispatchDriver
{
public:
	CProjectTemplate(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CProjectTemplate(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CProjectTemplate(const CProjectTemplate& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _ProjectTemplate ����
public:
	LPDISPATCH get_Application()
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

	// _ProjectTemplate ����
public:

};
