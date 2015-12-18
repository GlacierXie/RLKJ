// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类
#pragma  once
#ifndef   CCODE_MODULE_H
#define   CCODE_MODULE_H


// CCodeModule 包装类

class CCodeModule : public COleDispatchDriver
{
public:
	CCodeModule(){} // 调用 COleDispatchDriver 默认构造函数
	CCodeModule(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCodeModule(const CCodeModule& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 特性
public:

	// 操作
public:


	// _CodeModule 方法
public:
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x60020000, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_VBE()
	{
		LPDISPATCH result;
		InvokeHelper(0x60020001, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Name(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void AddFromString(LPCTSTR String)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x60020004, DISPATCH_METHOD, VT_EMPTY, NULL, parms, String);
	}
	void AddFromFile(LPCTSTR FileName)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x60020005, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FileName);
	}
	CString get_Lines(long StartLine, long Count)
	{
		CString result;
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x60020006, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, StartLine, Count);
		return result;
	}
	long get_CountOfLines()
	{
		long result;
		InvokeHelper(0x60020007, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void InsertLines(long Line, LPCTSTR String)
	{
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x60020008, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Line, String);
	}
	void DeleteLines(long StartLine, long Count)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x60020009, DISPATCH_METHOD, VT_EMPTY, NULL, parms, StartLine, Count);
	}
	void ReplaceLine(long Line, LPCTSTR String)
	{
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x6002000a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Line, String);
	}
	long get_ProcStartLine(LPCTSTR ProcName, long ProcKind)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_I4 ;
		InvokeHelper(0x6002000b, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, parms, ProcName, ProcKind);
		return result;
	}
	long get_ProcCountLines(LPCTSTR ProcName, long ProcKind)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_I4 ;
		InvokeHelper(0x6002000c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, parms, ProcName, ProcKind);
		return result;
	}
	long get_ProcBodyLine(LPCTSTR ProcName, long ProcKind)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_I4 ;
		InvokeHelper(0x6002000d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, parms, ProcName, ProcKind);
		return result;
	}
	CString get_ProcOfLine(long Line, long * ProcKind)
	{
		CString result;
		static BYTE parms[] = VTS_I4 VTS_PI4 ;
		InvokeHelper(0x6002000e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, parms, Line, ProcKind);
		return result;
	}
	long get_CountOfDeclarationLines()
	{
		long result;
		InvokeHelper(0x6002000f, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long CreateEventProc(LPCTSTR EventName, LPCTSTR ObjectName)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x60020010, DISPATCH_METHOD, VT_I4, (void*)&result, parms, EventName, ObjectName);
		return result;
	}
	BOOL Find(LPCTSTR Target, long * StartLine, long * StartColumn, long * EndLine, long * EndColumn, BOOL WholeWord, BOOL MatchCase, BOOL PatternSearch)
	{
		BOOL result;
		static BYTE parms[] = VTS_BSTR VTS_PI4 VTS_PI4 VTS_PI4 VTS_PI4 VTS_BOOL VTS_BOOL VTS_BOOL ;
		InvokeHelper(0x60020011, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Target, StartLine, StartColumn, EndLine, EndColumn, WholeWord, MatchCase, PatternSearch);
		return result;
	}
	LPDISPATCH get_CodePane()
	{
		LPDISPATCH result;
		InvokeHelper(0x60020012, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// _CodeModule 属性
public:

};

#endif   //CCODE_MODULE_H