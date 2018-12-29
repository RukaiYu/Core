/********************************************************************************
*					Tangram Library - version 10.0.0							*
*********************************************************************************
* Copyright (C) 2002-2018 by Tangram Team.   All Rights Reserved.				*
*
* THIS SOURCE FILE IS THE PROPERTY OF TANGRAM TEAM AND IS NOT TO
* BE RE-DISTRIBUTED BY ANY MEANS WHATSOEVER WITHOUT THE EXPRESSED
* WRITTEN CONSENT OF TANGRAM TEAM.
*
* THIS SOURCE CODE CAN ONLY BE USED UNDER THE TERMS AND CONDITIONS
* OUTLINED IN THE TANGRAM LICENSE AGREEMENT.TANGRAM TEAM
* GRANTS TO YOU (ONE SOFTWARE DEVELOPER) THE LIMITED RIGHT TO USE
* THIS SOFTWARE ON A SINGLE COMPUTER.
*
* CONTACT INFORMATION:
* mailto:tangramteam@outlook.com
* https://www.tangramteam.com
*
********************************************************************************/


// WndNode.h : Declaration of the CWndNode

#include "browser.h"
#include "WndPage.h"

#pragma once
class CNodeWnd;
class CWndNodeEvents;
class CTangramHtmlTreeWnd;
class CWndNodeWBEvent : 
	public CWebPageExtender,
	public CWebbrowserEvent
{
public:
	CWndNodeWBEvent(CWndNode* pNode);
	IHTMLDocument2*	m_pHTMLDocument2;

private:
	CWndNode* m_pWndNode;
	void __stdcall OnDocumentComplete(IDispatch *pDisp, VARIANT *URL);
	void __stdcall OnNavigateComplete2(IDispatch* pDisp, VARIANT* URL);
};

class ATL_NO_VTABLE CEventProxy : 
	public CComObjectRootBase,
	public IConnectionPointContainerImpl < CEventProxy >,
	public IConnectionPointImpl<CEventProxy, &DIID__IEventProxy>,
	public IDispatchImpl<IEventProxy, &IID_IEventProxy, &LIBID_Tangram,  1,  0>
{
public:
	BEGIN_COM_MAP(CEventProxy)
		COM_INTERFACE_ENTRY(IEventProxy)
		COM_INTERFACE_ENTRY(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(CEventProxy)
		CONNECTION_POINT_ENTRY(DIID__IEventProxy)
	END_CONNECTION_POINT_MAP()

	void Lock(){}
	void Unlock(){}

protected:
	ULONG InternalAddRef(){return 1;}
	ULONG InternalRelease(){return 1;}
};

class CTangramNodeCommonData
{
public:
	CTangramNodeCommonData();
	~CTangramNodeCommonData();

	CWndPage*							m_pPage;
	CWndFrame*							m_pFrame;
	CWndFrame*							m_pOldFrame;
	IDispatch*							m_pOfficeObj;
	CNodeWnd*							m_pHostClientView;
	CTangramXmlParse*					m_pTangramParse;
	map<CString, CWndNode*>				m_mapLayoutNodes;
	map<CString, CWndNode*>				m_mapAxNodes;
	map<CString, CWndNode*>				m_mapCLRNodes;
	map<CString, CWndNode*>				m_mapCppNodes;
	CMapStringToPtr						m_PlugInDispDictionary;
};

// CWndNode 
class ATL_NO_VTABLE CWndNode : 
	public CComObjectRootBase,
	public IConnectionPointContainerImpl<CWndNode>,
	public IConnectionPointImpl<CWndNode, &__uuidof(_IWndNodeEvents)>,
	public IDispatchImpl<IWndNode, &IID_IWndNode, &LIBID_Tangram,  1,  0>
{
public:
	CWndNode();
	virtual ~CWndNode();

	BOOL								m_bTopObj;
	BOOL								m_bCreated;
	BOOL								m_bWebInit;
	BOOL								m_bNodeDocComplete;

	ViewType							m_nViewType;
	int									m_nID;
	int									m_nActivePage;
	int									m_nWidth;
	int									m_nHeigh;
	int									m_nRow;
	int									m_nCol;
	int									m_nRows;
	int									m_nCols;
	HWND								m_hHostWnd;
	HWND								m_hChildHostWnd;

	CString 							m_strID;
	CString 							m_strURL;
	CString								m_strKey;
	CString								m_strName;
	CString 							m_strCnnID;
	CString 							m_strCaption;
	CString 							m_strXmlFileFromVS;

	CString								m_strNodeName;
	CString 							m_strExtenderID;

	IDispatch*							m_pDisp;
	CWndNode* 							m_pRootObj;
	CWndNode* 							m_pParentObj;
	CWndNode*							m_pVisibleXMLObj;
	CTangramXmlParse* 					m_pDocXmlParseNode;
	CTangramXmlParse*					m_pHostParse;

	CTangramNodeCommonData*				m_pTangramNodeCommonData;
	CTangramMessageObj*					m_pCurMsgObj;

	CWnd*								m_pHostWnd;
	IWndFrame*							m_pHostFrame;
	CRuntimeClass*						m_pObjClsInfo;

	CWndNodeEvents*						m_pCLREventConnector;
	CWndNodeWBEvent*					m_pTangramNodeWBEvent;

	VARIANT								m_varTag;
	IDispatch*							m_pExtender;
	IVSDocument*						m_pVSDocument;
	CTangramNodeVector					m_vChildNodes;
	CWndNode*							m_pCurrentExNode;
	map<CString, CWndFrame*>			m_mapSubFrame;
	map<CTangramAppProxy*, CTangramWndNodeProxy*> m_mapWndNodeProxy;
	CComObject<CWndNodeCollection>*		m_pChildNodeCollection;

	void	InitWndNode();
	BOOL	Create(DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext);
	BOOL	PreTranslateMessage(MSG* pMsg);
	BOOL	AddChildNode(CWndNode* pNode);
	BOOL	RemoveChildNode(CWndNode* pNode);
	CString GetNames();

	HRESULT Fire_ExtendComplete();
	HRESULT Fire_Destroy();
	HRESULT Fire_NodeAddInCreated(IDispatch * pAddIndisp, BSTR bstrAddInID, BSTR bstrAddInXml);
	HRESULT Fire_NodeAddInsCreated();
	HRESULT Fire_NodeDocumentComplete(IDispatch * ExtenderDisp, BSTR bstrURL);
	HRESULT Fire_ControlNotify(IWndNode * sender, LONG NotifyCode, LONG CtrlID, LONGLONG CtrlHandle, BSTR CtrlClassName);
	HRESULT Fire_TabChange(LONG ActivePage, LONG OldPage);
	HRESULT Fire_IPCMsg(BSTR bstrXml, BSTR bstrData);

	void Lock(){}
	void Unlock(){}
protected:
	ULONG InternalAddRef(){ return 1; }
	ULONG InternalRelease(){ return 1; }

public:
	STDMETHOD(get_XObject)(VARIANT* pVar);
	STDMETHOD(get_Tag)(VARIANT* pVar);
	STDMETHOD(put_Tag)(VARIANT var);
	STDMETHOD(get_AxPlugIn)(BSTR bstrName, IDispatch** pVal);
	STDMETHOD(get_Name)(BSTR* pVal);
	STDMETHOD(put_Name)(BSTR bstrName);
	STDMETHOD(get_Caption)(BSTR* pVal);
	STDMETHOD(put_Caption)(BSTR bstrCaption);
	STDMETHOD(get_Attribute)(BSTR bstrKey, BSTR* pVal);
	STDMETHOD(put_Attribute)(BSTR bstrKey, BSTR bstrVal);
	STDMETHOD(get_Handle)(LONGLONG* hWnd);
	STDMETHOD(get_XML)(BSTR* pVal);
	STDMETHOD(get_ChildNodes)(IWndNodeCollection** );
	STDMETHOD(get_Row)(long* nRow);
	STDMETHOD(get_Col)(long* nCol);
	STDMETHOD(get_Objects)(long nType, IWndNodeCollection** );
	STDMETHOD(get_NodeType)(ViewType* nType);
	STDMETHOD(get_ParentNode)(IWndNode** ppNode);
	STDMETHOD(get_RootNode)(IWndNode** ppNode);
	STDMETHOD(get_OuterXml)(BSTR* pVal);
	STDMETHOD(get_Key)(BSTR* pVal);
	STDMETHOD(get_Frame)(IWndFrame** pVal);
	STDMETHOD(get_HostFrame)(IWndFrame** pVal);
	STDMETHOD(get_Height)(LONG* pVal);
	STDMETHOD(get_Width)(LONG* pVal);
	STDMETHOD(get_Extender)(IDispatch** pVal);
	STDMETHOD(put_Extender)(IDispatch* newVal);
	STDMETHOD(get_VSDocument)(IVSDocument** pVal);
	//STDMETHOD(put_VSDocument)(IVSDocument* newVal);
	STDMETHOD(get_WndPage)(IWndPage** pVal);
	STDMETHOD(get_HTMLWindow)(IDispatch** pVal);
	STDMETHOD(get_HtmlDocument)(IDispatch** pVal);
	STDMETHOD(get_NameAtWindowPage)(BSTR* pVal);
	STDMETHOD(get_DocXml)(BSTR* pVal);
	STDMETHOD(get_HostNode)(IWndNode** pVal);
	STDMETHOD(put_HostNode)(IWndNode* newVal);
	STDMETHOD(get_ActivePage)(int* pVal);
	STDMETHOD(put_ActivePage)(int newVal);
	STDMETHOD(get_OfficeObj)(IDispatch** pVal);

	STDMETHOD(get_Rows)(long* nRows);
	STDMETHOD(get_Cols)(long* nCols);
	STDMETHOD(get_Hmin)(int* pVal);
	STDMETHOD(put_Hmin)(int newVal);
	STDMETHOD(get_Hmax)(int* pVal);
	STDMETHOD(put_Hmax)(int newVal);
	STDMETHOD(get_Vmin)(int* pVal);
	STDMETHOD(put_Vmin)(int newVal);
	STDMETHOD(get_Vmax)(int* pVal);
	STDMETHOD(put_Vmax)(int newVal);
	STDMETHOD(get_rgbMiddle)(OLE_COLOR* pVal);
	STDMETHOD(put_rgbMiddle)(OLE_COLOR newVal);
	STDMETHOD(get_rgbRightBottom)(OLE_COLOR* pVal);
	STDMETHOD(put_rgbRightBottom)(OLE_COLOR newVal);
	STDMETHOD(get_rgbLeftTop)(OLE_COLOR* pVal);
	STDMETHOD(put_rgbLeftTop)(OLE_COLOR newVal);
	STDMETHOD(get_MasterRow)(int* pVal);
	STDMETHOD(put_MasterRow)(int newVal);
	STDMETHOD(get_MasterCol)(int* pVal);
	STDMETHOD(put_MasterCol)(int newVal);
	STDMETHOD(put_SaveToConfigFile)(VARIANT_BOOL newVal);

	STDMETHOD(ExtendEx)(int nRow, int nCol, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode);
	STDMETHOD(ActiveTabPage)(IWndNode* pNode);
	STDMETHOD(Extend)(BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode);
	STDMETHOD(GetNode)(long nRow, long nCol,IWndNode** ppTangramNode);
	STDMETHOD(GetNodes)(BSTR bstrName, IWndNode** ppNode, IWndNodeCollection** ppNodes, long* pCount);
	STDMETHOD(GetCtrlByName)(BSTR bstrName, VARIANT_BOOL bFindInChild, IDispatch** ppRetDisp);
	STDMETHOD(GetCtrlValueByName)(BSTR bstrName, VARIANT_BOOL bFindInChild, BSTR* bstrVal);
	STDMETHOD(SetCtrlValueByName)(BSTR bstrName, VARIANT_BOOL bFindInChild, BSTR bstrVal);
	STDMETHOD(Refresh)(void);
	STDMETHOD(LoadXML)(int nType, BSTR bstrXML);
	STDMETHOD(Show)();
	STDMETHOD(GetNodeByName)(BSTR bstrName, IWndNodeCollection** pVal);
	STDMETHOD(get_DockObj)(BSTR bstrName, LONGLONG* pVal);
	STDMETHOD(put_DockObj)(BSTR bstrName, LONGLONG newVal);
	STDMETHOD(SendUCMAMsg)(BSTR bstrSipFrom, BSTR bstrTo, BSTR bstrMsg);
	STDMETHOD(get_CurMsgObj)(IMessageObj** pVal);

	BEGIN_COM_MAP(CWndNode)
		COM_INTERFACE_ENTRY(IWndNode)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(CWndNode)
		CONNECTION_POINT_ENTRY(__uuidof(_IWndNodeEvents))
	END_CONNECTION_POINT_MAP()

	HWND CreateView(HWND hParentWnd, CString strTag);

private:
	CString _GetNames(CWndNode* pNode);
	void _get_Objects(CWndNode* pNode, UINT32& nType, CWndNodeCollection* pNodeColletion);
	int _getNodes(CWndNode* pNode, CString& strName, CWndNode**ppRetNode, CWndNodeCollection* pNodes);
public:
	STDMETHOD(NavigateURL)(BSTR bstrURL, IDispatch* dispObjforScript);
	STDMETHOD(get_URL)(BSTR* pVal);
	STDMETHOD(put_URL)(BSTR newVal);
};

// CWndNodeCollection

class ATL_NO_VTABLE CWndNodeCollection :
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDispatchImpl<IWndNodeCollection, &IID_IWndNodeCollection, &LIBID_Tangram, 1, 0>
{
public:
	CWndNodeCollection();
	~CWndNodeCollection();

	BEGIN_COM_MAP(CWndNodeCollection)
		COM_INTERFACE_ENTRY(IWndNodeCollection)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()

	STDMETHOD(get_NodeCount)(long* pCount);
	STDMETHOD(get_Item)(long iIndex, IWndNode **ppTangramNode);
	STDMETHOD(get__NewEnum)(IUnknown** ppVal);
	CTangramNodeVector*	m_pNodes;

private:
	CTangramNodeVector	m_vNodes;
};
