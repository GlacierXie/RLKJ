// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

#pragma once
#ifndef  CCOMPONENT_H
#define  CCOMPONENT_H

// CComponent ��װ��

class CComponent : public COleDispatchDriver
{
public:
	CComponent(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CComponent(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CComponent(const CComponent& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _Component ����
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
	BOOL get_IsDirty()
	{
		BOOL result;
		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_IsDirty(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0xa, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
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

	// _Component ����
public:

};
#endif   //CCOMPONENT_H