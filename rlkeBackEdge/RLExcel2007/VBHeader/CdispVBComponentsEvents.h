// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类


// CdispVBComponentsEvents 包装类

class CdispVBComponentsEvents : public COleDispatchDriver
{
public:
	CdispVBComponentsEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CdispVBComponentsEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CdispVBComponentsEvents(const CdispVBComponentsEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _dispVBComponentsEvents 方法
public:
	void ItemAdded(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}
	void ItemRemoved(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}
	void ItemRenamed(LPDISPATCH VBComponent, LPCTSTR OldName)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR ;
		InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent, OldName);
	}
	void ItemSelected(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}
	void ItemActivated(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x5, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}
	void ItemReloaded(LPDISPATCH VBComponent)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBComponent);
	}

	// _dispVBComponentsEvents 属性
public:

};
