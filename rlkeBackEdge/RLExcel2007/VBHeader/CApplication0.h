// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

//using namespace VBIDE;
#pragma  once
#ifndef  CAPPLICATION0_H
#define  CAPPLICATION0_H

// CApplication0 ��װ��

class CApplication0 : public COleDispatchDriver
{
public:
	CApplication0(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CApplication0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplication0(const CApplication0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// Application ����
public:
	CString get_Version()
	{
		CString result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}

	// Application ����
public:

};
#endif  //CAPPLICATION0_H