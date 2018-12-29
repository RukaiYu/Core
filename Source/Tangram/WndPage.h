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

// WndPage.h : CWndPage implementation file

#pragma once

#include <assert.h>
#include <wininet.h>
#include "WebPageExtender.h"

#include "EclipsePlus\EclipseAddin.h"

// CWndPage
class ATL_NO_VTABLE CWndPage : 
	public CWebPageExtender,
	public CComObjectRootBase,
	public IPropertyNotifySink,
	public IConnectionPointContainerImpl <CWndPage>,
	public IConnectionPointImpl<CWndPage, &__uuidof(_ITangramObjEvents)>,
	public IDispatchImpl<IWndPage, &IID_IWndPage, &LIBID_Tangram,  1,  0>
{
	friend CTangram;
public:
	CWndPage();
	virtual ~CWndPage();

	int											m_nWebViewCount;
	BOOL										m_bDoc;
	BOOL										m_bPageDataLoaded;
	HWND										m_hWnd;

	CString										m_strPageFilePath;
	CString										m_strPageFileName;
	CString										m_strConfigFileNodeName;
	CString										m_strXmlHeader;
	CString										m_strXmlBottom;
	map<CTangramAppProxy*, CTangramWndPageProxy*> m_mapWndPageProxy;

	CWndFrame*									m_pBKFrame;
	CTangramDoc*								m_pActiveDoc;
	CTangramAppProxy*							m_pTangramAppProxy;
	CTangramDocTemplate*						m_pTangramDocTemplate;
	map<CString, CString>						m_strMapKey;
	map<HWND, CWndFrame*>						m_mapCtrlBarFrame;
	map<CString, HWND>							m_mapWnd;
	map<HWND, CString>							m_mapVisualDesignWnd;
	map<HWND, CWndFrame*>						m_mapFrame;
	map<HWND, CWndFrame*>						m_mapNeedSaveFrame;
	map<CString, CWndNode*>						m_mapNode;
	map<CString, IDispatch*>					m_mapExternalObj;
	map<IHTMLDocument2*, CWebPageExtender*>		m_mapWebPageExtender;
	map<HWND, CEclipseCtrl*>					m_mapNotifyCtrl;
	map<CString, CString>						m_mapXtml;
	map<__int64, CTangramEventObj*>				m_mapEventObj;

	BEGIN_COM_MAP(CWndPage)
		COM_INTERFACE_ENTRY(IWndPage)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(IPropertyNotifySink)
		COM_INTERFACE_ENTRY(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(CWndPage)
		CONNECTION_POINT_ENTRY(__uuidof(_ITangramObjEvents))
	END_CONNECTION_POINT_MAP()

	void Lock(){}
	void Unlock(){}
	void BeforeDestory();
	void UpdateMapKey(CString);
	HRESULT Fire_PageLoaded(IDispatch* sender, BSTR url);
	HRESULT Fire_NodeCreated(IWndNode * pNodeCreated);
	HRESULT Fire_AddInCreated(IWndNode * pRootNode, IDispatch * pAddIn, BSTR bstrID, BSTR bstrAddInXml);
	HRESULT Fire_BeforeExtendXml(BSTR bstrXml, LONGLONG hWnd);
	HRESULT Fire_ExtendXmlComplete(BSTR bstrXml, LONGLONG hWnd, IWndNode * pRetRootNode);
	HRESULT Fire_Destroy();
	HRESULT Fire_NodeMouseActivate(IWndNode * pActiveNode);
	HRESULT Fire_ClrControlCreated(IWndNode * Node, IDispatch * Ctrl, BSTR CtrlName, LONGLONG CtrlHandle);
	HRESULT Fire_TabChange(IWndNode* sender, LONG ActivePage, LONG OldPage);
	HRESULT Fire_TangramEvent(ITangramEventObj* pEventObj);
	HRESULT Fire_IPCMsg(IWndFrame* pSender, BSTR bstrType, BSTR bstrContent, BSTR bstrFeature);

	void OnNodeDocComplete(WPARAM);

protected:
	ULONG InternalAddRef(){ return 1; }
	ULONG InternalRelease(){ return 1; }

public:
	STDMETHOD(get_Count)(long* pCount);
	STDMETHOD(get_Frame)(VARIANT vIndex, IWndFrame **ppFrame);
	STDMETHOD(get__NewEnum)(IUnknown** ppVal);
	STDMETHOD(get_External)(IDispatch** ppVal);
	STDMETHOD(put_External)(IDispatch* newVal);
	STDMETHOD(get_URL)(BSTR* pVal);
	STDMETHOD(put_URL)(BSTR newVal);
	STDMETHOD(get_Handle)(LONGLONG* pVal);
	STDMETHOD(get_Extender)(BSTR bstrExtenderName, IDispatch** pVal);
	STDMETHOD(put_Extender)(BSTR bstrExtenderName, IDispatch* newVal);
	STDMETHOD(get_FrameName)(LONGLONG hHwnd, BSTR* pVal);
	STDMETHOD(get_WndFrame)(BSTR bstrFrameName, IWndFrame** pVal);
	STDMETHOD(get_Node)(BSTR bstrNodeName, IWndNode** pVal);
	STDMETHOD(get_NodeNames)(BSTR* pVal);
	STDMETHOD(get_XObject)(BSTR bstrName, IDispatch** pVal);
	STDMETHOD(get_HTMLWindow)(BSTR NodeName, IDispatch** pVal);
	STDMETHOD(get_HtmlDocument)(IDispatch** pVal);
	STDMETHOD(get_Parent)(IWndPage** pVal);
	STDMETHOD(get_EnableSinkCLRCtrlCreatedEvent)(VARIANT_BOOL* pVal);
	STDMETHOD(put_EnableSinkCLRCtrlCreatedEvent)(VARIANT_BOOL newVal);
	STDMETHOD(get_Width)(long* pVal);
	STDMETHOD(put_Width)(long newVal);
	STDMETHOD(get_Height)(long* pVal);
	STDMETHOD(put_Height)(long newVal);
	STDMETHOD(get_xtml)(BSTR strKey, BSTR* pVal);
	STDMETHOD(put_xtml)(BSTR strKey, BSTR newVal);
	STDMETHOD(get_PageXML)(BSTR* pVal);
	STDMETHOD(put_ConfigName)(BSTR newVal);

	STDMETHOD(AttachNodeEvent)(BSTR bstrNames, IDispatch* pWndDisp);
	STDMETHOD(AttachObjEvent)(IDispatch* HTMLWindow, IDispatch* SourceObj, BSTR bstrEventName, BSTR AliasName);
	STDMETHOD(AttachNodeSubCtrlEvent)(IDispatch* HtmlWndDisp, VARIANT Node, VARIANT CtrlName, BSTR EventName, BSTR AliasName);
	STDMETHOD(AddObjToHtml)(BSTR strObjName, VARIANT_BOOL bConnectEvent, IDispatch* pObjDisp);
	STDMETHOD(AttachEvent)(BSTR bstrName, IDispatch* pDoc, VARIANT_BOOL* bResult);
	STDMETHOD(AddNodesToPage)(IDispatch* pHtmlDoc, BSTR bstrNodeNames, VARIANT_BOOL bAdd, VARIANT_BOOL* bSuccess);
	STDMETHOD(AddObject)(IDispatch* pHtmlDoc, IDispatch* pObject, BSTR bstrObjName, VARIANT_BOOL bSinkEvent, VARIANT_BOOL* bResult);
	STDMETHOD(CreateFrame)(VARIANT ParentObj, VARIANT HostWnd, BSTR bstrFrameName, IWndFrame** pRetFrame);
	STDMETHOD(GetWndNode)(BSTR bstrFrameName, BSTR bstrNodeName, IWndNode** pRetNode);
	STDMETHOD(GetFrameFromCtrl)(IDispatch* ctrl, IWndFrame** ppFrame);
	STDMETHOD(GetHtmlDocument)(IDispatch* HtmlWindow, IDispatch** ppHtmlDoc);
	STDMETHOD(GetCtrlInNode)(BSTR NodeName, BSTR CtrlName, IDispatch** ppCtrl);
	STDMETHOD(Extend)(IDispatch* Parent, BSTR CtrlName, BSTR FrameName, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode);
	STDMETHOD(ExtendCtrl)(VARIANT MdiForm, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode);
	STDMETHOD(ConnectTangramCtrl)(ITangramCtrl* eventSource);
	STDMETHOD(CreateFrameWithDefaultNode)(ULONGLONG hFrameWnd, BSTR bstrFrameName, BSTR bstrDefaultNodeKey, BSTR bstrXml, VARIANT_BOOL bSaveToConfig, IWndNode** ppNode);
	STDMETHOD(ExtendFrames)(BSTR bstrFrames, BSTR bstrKey, BSTR bstrXml, VARIANT_BOOL bSaveToConfigFile);
	STDMETHOD(get_CurrentDesignFrameType)(WndFrameType* pVal);
	STDMETHOD(get_CurrentDesignNode)(IWndNode** pVal);

private:
	bool										m_bIsBlank;
	bool										m_bIsDestory;
	bool										m_bDocComplete;

	VARIANT_BOOL								m_bEnableSinkCLRCtrlCreatedEvent;

	DWORD										m_nThreadId;
	DWORD										m_dwCookie;
	HRESULT										m_hrConnected;
	READYSTATE									m_lReadyState;
	INTERNET_SCHEME								m_nScheme;
	CString										m_strURL;

	IDispatch*									m_pExternalDisp;
	IHTMLDocument2*								m_pHtmlDocument2;
	LPCONNECTIONPOINT							m_pCP;

	HRESULT Init(CString szURL);
	HRESULT UnLoad();
	HRESULT LoadURLFromFile();
	HRESULT LoadURLFromMoniker();

	// IPropertyNotifySink methods
	STDMETHOD(OnChanged)(DISPID dispID);
	STDMETHOD(OnRequestEdit)(DISPID dispID);
};

