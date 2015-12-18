// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类
#pragma once
#ifndef CWindows_h
#define CWindows_h


// CWindows 包装类

class CWindows : public COleDispatchDriver
{
public:
	CWindows(){} // 调用 COleDispatchDriver 默认构造函数
	CWindows(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CWindows(const CWindows& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _Windows 方法
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

	// _Windows 属性
public:

};

#endif  //CWindows_h
