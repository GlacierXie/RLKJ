// �����Ϳ������á�����ࡱ�����ļ�������ɵ� IDispatch ��װ��

// CdispCommandBarControlEvents ��װ��
#pragma once
#ifndef  CdispCommandBarControlEvents_H
#define  CdispCommandBarControlEvents_H

class CdispCommandBarControlEvents : public COleDispatchDriver
{
public:
	CdispCommandBarControlEvents(){} // ���� COleDispatchDriver Ĭ�Ϲ��캯��
	CdispCommandBarControlEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CdispCommandBarControlEvents(const CdispCommandBarControlEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// ����
public:

	// ����
public:


	// _dispCommandBarControlEvents ����
public:
	void Click(LPDISPATCH CommandBarControl, BOOL * handled, BOOL * CancelDefault)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL VTS_PBOOL ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, CommandBarControl, handled, CancelDefault);
	}

	// _dispCommandBarControlEvents ����
public:

};
#endif //CdispCommandBarControlEvents_H