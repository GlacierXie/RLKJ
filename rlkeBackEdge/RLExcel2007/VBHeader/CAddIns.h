// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类
#pragma  once
#ifndef  CADDINS_H
#define  CADDINS_H
// CAddIns 包装类

class CAddIns : public COleDispatchDriver
{
public:
	CAddIns(){} // 调用 COleDispatchDriver 默认构造函数
	CAddIns(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CAddIns(const CAddIns& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _AddIns 方法
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
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
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
		InvokeHelper(0x28, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPUNKNOWN _NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_METHOD, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	void Update()
	{
		InvokeHelper(0x29, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// _AddIns 属性
public:

};
#endif  //CADDINS_H