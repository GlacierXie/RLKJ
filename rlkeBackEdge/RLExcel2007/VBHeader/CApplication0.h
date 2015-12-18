// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

//using namespace VBIDE;
#pragma  once
#ifndef  CAPPLICATION0_H
#define  CAPPLICATION0_H

// CApplication0 包装类

class CApplication0 : public COleDispatchDriver
{
public:
	CApplication0(){} // 调用 COleDispatchDriver 默认构造函数
	CApplication0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CApplication0(const CApplication0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// Application 方法
public:
	CString get_Version()
	{
		CString result;
		InvokeHelper(0x64, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}

	// Application 属性
public:

};
#endif  //CAPPLICATION0_H