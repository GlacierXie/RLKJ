/********************************************************************
	created:	2015/10/09
	file base:	ZWMainRibPanelWrapper
	file ext:	h

	purpose:	����ribon�˵�������
*********************************************************************/
#ifndef __CAcadRibPanelWrapper_H_
#define __CAcadRibPanelWrapper_H_

#include "tchar.h"
#include "resource.h"       // main symbols
#include <AtlBase.h>
#pragma warning( disable : 4278 )
#import "acax19ENU.tlb" no_implementation raw_interfaces_only named_guids
#pragma warning( default : 4278 )


//����·��
/////////////////////////////////////////////////////////////////////////////
typedef struct _SPopuMenuItemStyle
{
	_SPopuMenuItemStyle():strName(_T("")),
						 strMacro(_T("")),
						 isSeparator(false)
	{}
	_SPopuMenuItemStyle( CString menuName, CString menuMacro,bool isSep= false)
	{
		strName = menuName;
		strMacro = menuMacro;
		isSeparator  = isSep;
	}

	_SPopuMenuItemStyle(const _SPopuMenuItemStyle& rhl)
	{
		*this = rhl;
	}

	_SPopuMenuItemStyle& operator =(const _SPopuMenuItemStyle&rhl )
	{
		strName = rhl.strName;
		strMacro = rhl.strMacro;
		isSeparator = rhl.isSeparator;
		return *this;
	}

	CString strName;					//������
	CString strMacro;					//����
	BOOL	isSeparator;				//�Ƿ��Ƿָ���

} SPopuMenuItemStyle;

typedef CArray<SPopuMenuItemStyle> CArrayMenuPopuItem;
class  CAcadRibPanelWrapper
{
public:
    CAcadRibPanelWrapper();
	~CAcadRibPanelWrapper();

public:
	//{@ ������صĲ���

	//	@bref: ��ʼ����ػ���
	bool InitMenuPanel();

	//	@bref: �Ƴ����в˵�
	bool RremoveAllMenus();

	//	@bref: ���ز˵��������ռ�
	bool LoadMenuFromCuix(const CString& strCuixPath, const CString& workSpaceName);

	//	@bref: ����������
	void CreateToolBar(const CString& GroupName, const CArrayMenuPopuItem& menuItems);

	//	@bref: �õ�menuBar�ϲ˵�����
	int GetMenuCntByMenuBar();

	//	@bref: ����һ����ݲ˵�
	//			menuItems->MenuGroups:  ��������
	//			menuItems->MenuMacros:  ִ�����������꣬��ARX���ֶ�����������һ��
	//	  ����: �Ƿ��Ѿ����������ز˵�
	bool CreatePopuMenu( CString strMenuTitle,const CArrayMenuPopuItem& menuItems);

	BOOL creatMenuOnCurren(CString strMenuTitle, const CArrayMenuPopuItem& menuItems);

	void BrushRib();
	void GetRibStyle();
	void SetRibStyle();
	void ModifyRibStyle();
	//@}

protected:
	AutoCAD::IAcadMenuGroup * FindMenuGroupByName(const CString& menuName);
private:
    //short mNumber;
	AutoCAD::IAcadApplication *m_pAcApplication;
	AutoCAD::IAcadMenuGroups  *m_pAcMenuGroups ;
	AutoCAD::IAcadMenuBar	  *m_pAcMenuBar;

	static bool m_bIsMenuLoaded;
};

#endif
