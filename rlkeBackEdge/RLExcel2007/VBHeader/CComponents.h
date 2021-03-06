// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类
#pragma once
#ifndef  CCOMPONENTS_H
#define  CCOMPONENTS_H

// CComponents 包装类

class CComponents : public COleDispatchDriver
{
public:
	CComponents(){} // 调用 COleDispatchDriver 默认构造函数
	CComponents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CComponents(const CComponents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _Components 方法
public:
	LPDISPATCH Item(VARIANT& index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &index);
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
	void Remove(LPDISPATCH Component)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Component);
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

	// _Components 属性
public:

};
#endif   //CCOMPONENTS_H