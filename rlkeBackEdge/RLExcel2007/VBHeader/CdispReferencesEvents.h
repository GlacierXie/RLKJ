// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CdispReferencesEvents 包装类

class CdispReferencesEvents : public COleDispatchDriver
{
public:
	CdispReferencesEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CdispReferencesEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CdispReferencesEvents(const CdispReferencesEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _dispReferencesEvents 方法
public:
	void ItemAdded(LPDISPATCH Reference)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Reference);
	}
	void ItemRemoved(LPDISPATCH Reference)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Reference);
	}

	// _dispReferencesEvents 属性
public:

};
