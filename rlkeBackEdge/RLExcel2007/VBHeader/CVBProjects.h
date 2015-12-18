// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类
#pragma once
#ifndef CVBProjects_H
#define CVBProjects_H
// CVBProjects 包装类

class CVBProjects : public COleDispatchDriver
{
public:
	CVBProjects(){} // 调用 COleDispatchDriver 默认构造函数
	CVBProjects(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CVBProjects(const CVBProjects& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _VBProjects 方法
public:
	LPDISPATCH Item(VARIANT& index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &index);
		return result;
	}
	LPDISPATCH get_VBE()
	{
		LPDISPATCH result;
		InvokeHelper(0x14, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPUNKNOWN _NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_METHOD, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Add(long Type)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x89, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Type);
		return result;
	}
	void Remove(LPDISPATCH lpc)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x8a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, lpc);
	}
	LPDISPATCH Open(LPCTSTR bstrPath)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x8b, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, bstrPath);
		return result;
	}

	// _VBProjects 属性
public:

};
#endif //CVBProjects_H