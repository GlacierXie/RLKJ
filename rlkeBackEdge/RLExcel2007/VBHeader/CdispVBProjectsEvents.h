// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CdispVBProjectsEvents 包装类

class CdispVBProjectsEvents : public COleDispatchDriver
{
public:
	CdispVBProjectsEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CdispVBProjectsEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CdispVBProjectsEvents(const CdispVBProjectsEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _dispVBProjectsEvents 方法
public:
	void ItemAdded(LPDISPATCH VBProject)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBProject);
	}
	void ItemRemoved(LPDISPATCH VBProject)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x2, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBProject);
	}
	void ItemRenamed(LPDISPATCH VBProject, LPCTSTR OldName)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_BSTR ;
		InvokeHelper(0x3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBProject, OldName);
	}
	void ItemActivated(LPDISPATCH VBProject)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x4, DISPATCH_METHOD, VT_EMPTY, NULL, parms, VBProject);
	}

	// _dispVBProjectsEvents 属性
public:

};
