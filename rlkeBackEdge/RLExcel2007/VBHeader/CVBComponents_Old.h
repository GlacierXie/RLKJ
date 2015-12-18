// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类
#pragma once
#ifndef CVBComponents_Old_h
#define CVBComponents_Old_h

// CVBComponents_Old 包装类

class CVBComponents_Old : public COleDispatchDriver
{
public:
	CVBComponents_Old(){} // 调用 COleDispatchDriver 默认构造函数
	CVBComponents_Old(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CVBComponents_Old(const CVBComponents_Old& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _VBComponents_Old 方法
public:
	LPDISPATCH Item(VARIANT& index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &index);
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
	void Remove(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}
	LPDISPATCH Add(long ComponentType)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, ComponentType);
		return result;
	}
	LPDISPATCH Import(LPCTSTR FileName)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName);
		return result;
	}
	LPDISPATCH get_VBE()
	{
		LPDISPATCH result;
		InvokeHelper(0x14, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// _VBComponents_Old 属性
public:

};

#endif  //CVBComponents_Old_h