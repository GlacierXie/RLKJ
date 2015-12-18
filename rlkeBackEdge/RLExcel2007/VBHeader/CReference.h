// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��
#pragma once
#ifndef CReference_h
#define CReference_h

// CReference ��װ��

class CReference : public COleDispatchDriver
{
public:
	CReference(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CReference(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CReference(const CReference& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// Reference ����
public:
	LPDISPATCH get_Collection()
	{
		LPDISPATCH result;
		InvokeHelper(0x60020000, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_VBE()
	{
		LPDISPATCH result;
		InvokeHelper(0x60020001, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x60020002, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get_Guid()
	{
		CString result;
		InvokeHelper(0x60020003, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	long get_Major()
	{
		long result;
		InvokeHelper(0x60020004, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Minor()
	{
		long result;
		InvokeHelper(0x60020005, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	CString get_FullPath()
	{
		CString result;
		InvokeHelper(0x60020006, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	BOOL get_BuiltIn()
	{
		BOOL result;
		InvokeHelper(0x60020007, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_IsBroken()
	{
		BOOL result;
		InvokeHelper(0x60020008, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x60020009, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	CString get_Description()
	{
		CString result;
		InvokeHelper(0x6002000a, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}

	// Reference ����
public:

};
#endif  //CReference_h