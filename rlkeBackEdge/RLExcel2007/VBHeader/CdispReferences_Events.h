// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CdispReferences_Events 包装类

class CdispReferences_Events : public COleDispatchDriver
{
public:
	CdispReferences_Events(){} // 调用 COleDispatchDriver 默认构造函数
	CdispReferences_Events(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CdispReferences_Events(const CdispReferences_Events& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _dispReferences_Events 方法
public:
	void ItemAdded(LPDISPATCH Reference)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Reference);
	}
	void ItemRemoved(LPDISPATCH Reference)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Reference);
	}

	// _dispReferences_Events 属性
public:

};
