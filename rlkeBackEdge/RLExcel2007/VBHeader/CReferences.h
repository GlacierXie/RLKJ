// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

#pragma once
#ifndef CReferences_h
#define CReferences_h
// CReferences 包装类

class CReferences : public COleDispatchDriver
{
public:
	CReferences(){} // 调用 COleDispatchDriver 默认构造函数
	CReferences(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CReferences(const CReferences& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _References 方法
public:
	LPDISPATCH get_Parent()
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
		InvokeHelper(0x60020003, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPUNKNOWN _NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_METHOD, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH AddFromGuid(LPCTSTR Guid, long Major, long Minor)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4 ;
		InvokeHelper(0x60020005, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Guid, Major, Minor);
		return result;
	}
	LPDISPATCH AddFromFile(LPCTSTR FileName)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x60020006, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName);
		return result;
	}
	void Remove(LPDISPATCH Reference)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x60020007, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Reference);
	}

	// _References 属性
public:

};
#endif //