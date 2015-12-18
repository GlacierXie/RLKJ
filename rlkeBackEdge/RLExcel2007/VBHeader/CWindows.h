// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��
#pragma once
#ifndef CWindows_h
#define CWindows_h


// CWindows ��װ��

class CWindows : public COleDispatchDriver
{
public:
	CWindows(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CWindows(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWindows(const CWindows& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _Windows ����
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
	LPDISPATCH CreateToolWindow(LPDISPATCH AddInInst, LPCTSTR ProgId, LPCTSTR Caption, LPCTSTR GuidPosition, LPDISPATCH * DocObj)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR VTS_BSTR VTS_BSTR VTS_PDISPATCH ;
		InvokeHelper(0x12c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, AddInInst, ProgId, Caption, GuidPosition, DocObj);
		return result;
	}

	// _Windows ����
public:

};

#endif  //CWindows_h
