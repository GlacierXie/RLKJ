// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CdispCommandBarControlEvents 包装类
#pragma once
#ifndef  CdispCommandBarControlEvents_H
#define  CdispCommandBarControlEvents_H

class CdispCommandBarControlEvents : public COleDispatchDriver
{
public:
	CdispCommandBarControlEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CdispCommandBarControlEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CdispCommandBarControlEvents(const CdispCommandBarControlEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _dispCommandBarControlEvents 方法
public:
	void Click(LPDISPATCH CommandBarControl, BOOL * handled, BOOL * CancelDefault)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL VTS_PBOOL ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, CommandBarControl, handled, CancelDefault);
	}

	// _dispCommandBarControlEvents 属性
public:

};
#endif //CdispCommandBarControlEvents_H