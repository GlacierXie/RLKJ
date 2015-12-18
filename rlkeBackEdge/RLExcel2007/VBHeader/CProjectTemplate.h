// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CProjectTemplate 包装类

class CProjectTemplate : public COleDispatchDriver
{
public:
	CProjectTemplate(){} // 调用 COleDispatchDriver 默认构造函数
	CProjectTemplate(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CProjectTemplate(const CProjectTemplate& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _ProjectTemplate 方法
public:
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

	// _ProjectTemplate 属性
public:

};
