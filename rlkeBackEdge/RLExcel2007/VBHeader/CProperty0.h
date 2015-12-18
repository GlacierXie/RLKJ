// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类
#pragma  once
#ifndef CProperty0_H
#define CProperty0_H

// CProperty0 包装类

class CProperty0 : public COleDispatchDriver
{
public:
	CProperty0(){} // 调用 COleDispatchDriver 默认构造函数
	CProperty0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CProperty0(const CProperty0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// Property 方法
public:
	VARIANT get_Value()
	{
		VARIANT result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_Value(VARIANT& newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	VARIANT get_IndexedValue(VARIANT& Index1, VARIANT& Index2, VARIANT& Index3, VARIANT& Index4)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, parms, &Index1, &Index2, &Index3, &Index4);
		return result;
	}
	void put_IndexedValue(VARIANT& Index1, VARIANT& Index2, VARIANT& Index3, VARIANT& Index4, VARIANT& newValue)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x3, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &Index1, &Index2, &Index3, &Index4, &newValue);
	}
	short get_NumIndices()
	{
		short result;
		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_I2, (void*)&result, NULL);
		return result;
	}
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
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x28, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_VBE()
	{
		LPDISPATCH result;
		InvokeHelper(0x29, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Collection()
	{
		LPDISPATCH result;
		InvokeHelper(0x2a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPUNKNOWN get_Object()
	{
		LPUNKNOWN result;
		InvokeHelper(0x2d, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	void putref_Object(LPUNKNOWN newValue)
	{
		static BYTE parms[] = VTS_UNKNOWN ;
		InvokeHelper(0x2d, DISPATCH_PROPERTYPUTREF, VT_EMPTY, NULL, parms, newValue);
	}

	// Property 属性
public:

};
#endif  //CProperty0_H