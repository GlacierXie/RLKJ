// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

#pragma once
#ifndef  CCODE_PANE_H
#define  CCODE_PANE_H

// CCodePane 包装类

class CCodePane : public COleDispatchDriver
{
public:
	CCodePane(){} // 调用 COleDispatchDriver 默认构造函数
	CCodePane(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCodePane(const CCodePane& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _CodePane 方法
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
	LPDISPATCH get_Window()
	{
		LPDISPATCH result;
		InvokeHelper(0x60020002, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void GetSelection(long * StartLine, long * StartColumn, long * EndLine, long * EndColumn)
	{
		static BYTE parms[] = VTS_PI4 VTS_PI4 VTS_PI4 VTS_PI4 ;
		InvokeHelper(0x60020003, DISPATCH_METHOD, VT_EMPTY, NULL, parms, StartLine, StartColumn, EndLine, EndColumn);
	}
	void SetSelection(long StartLine, long StartColumn, long EndLine, long EndColumn)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 ;
		InvokeHelper(0x60020004, DISPATCH_METHOD, VT_EMPTY, NULL, parms, StartLine, StartColumn, EndLine, EndColumn);
	}
	long get_TopLine()
	{
		long result;
		InvokeHelper(0x60020005, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_TopLine(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x60020005, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_CountOfVisibleLines()
	{
		long result;
		InvokeHelper(0x60020007, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CodeModule()
	{
		LPDISPATCH result;
		InvokeHelper(0x60020008, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Show()
	{
		InvokeHelper(0x60020009, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long get_CodePaneView()
	{
		long result;
		InvokeHelper(0x6002000a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}

	// _CodePane 属性
public:

};
#endif  //CCODE_PANE_H