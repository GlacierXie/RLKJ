// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��
#pragma once
#ifndef CWindows_old_h
#define CWindows_old_h

// CWindows_old ��װ��

class CWindows_old : public COleDispatchDriver
{
public:
	CWindows_old(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CWindows_old(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWindows_old(const CWindows_old& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _Windows_old ����
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

	// _Windows_old ����
public:

};

#endif  //CWindows_old_h