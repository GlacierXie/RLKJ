// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CEvents 包装类

class CEvents : public COleDispatchDriver
{
public:
	CEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CEvents(const CEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// Events 方法
public:
	LPUNKNOWN get_ReferencesEvents(LPDISPATCH VBProject)
	{
		LPUNKNOWN result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xca, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, parms, VBProject);
		return result;
	}
	LPUNKNOWN get_CommandBarEvents(LPDISPATCH CommandBarControl)
	{
		LPUNKNOWN result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0xcd, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, parms, CommandBarControl);
		return result;
	}

	// Events 属性
public:

};
