// �����Ϳ������á������ࡱ�����ļ�������ɵ� IDispatch ��װ��
#pragma once
#ifndef CVBComponent_h
#define CVBComponent_h


// CVBComponent ��װ��

class CVBComponent : public COleDispatchDriver
{
public:
	CVBComponent(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CVBComponent(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CVBComponent(const CVBComponent& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _VBComponent ����
public:
	BOOL get_Saved()
	{
		BOOL result;
		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x30, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Name(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x30, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Designer()
	{
		LPDISPATCH result;
		InvokeHelper(0x31, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CodeModule()
	{
		LPDISPATCH result;
		InvokeHelper(0x32, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x33, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void Export(LPCTSTR FileName)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x34, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
	}
	LPDISPATCH get_VBE()
	{
		LPDISPATCH result;
		InvokeHelper(0x35, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Collection()
	{
		LPDISPATCH result;
		InvokeHelper(0x36, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_HasOpenDesigner()
	{
		BOOL result;
		InvokeHelper(0x37, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Properties()
	{
		LPDISPATCH result;
		InvokeHelper(0x38, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH DesignerWindow()
	{
		LPDISPATCH result;
		InvokeHelper(0x39, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Activate()
	{
		InvokeHelper(0x3c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	CString get_DesignerID()
	{
		CString result;
		InvokeHelper(0x40, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}

	// _VBComponent ����
public:

};
#endif  //CVBComponent_h