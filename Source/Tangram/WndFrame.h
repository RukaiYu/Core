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

#pragma once
#include "chromium/BrowserWnd.h"
#include "chromium/HtmlWnd.h"

using namespace ChromePlus;

class CWndFrame;
class CWPFView;
class CTangramDoc;
class CEclipseWnd;
class CTangramDocFrame;
class CTangramEventObj;
class CTangramHtmlTreeWnd;
#define WM_TANGRAMNOTIFY WM_NOTIFY+WM_REFLECT_BASE
class CTangramCommonCtrl :
	public CWindowImpl<CTangramCommonCtrl, CWindow>
{
public:
	CTangramCommonCtrl(void);
	virtual ~CTangramCommonCtrl(void);
	BEGIN_MSG_MAP(CTangramCommonCtrl)
		MESSAGE_HANDLER(WM_NOTIFY, OnNotify)
	END_MSG_MAP()
	virtual void OnFinalMessage(HWND hWnd);
	LRESULT OnNotify(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
};

class CTangramAFXHelperWnd :
	public CWindowImpl<CTangramAFXHelperWnd, CWindow>
{
public:
	CTangramAFXHelperWnd(void);
	virtual ~CTangramAFXHelperWnd(void);
	HWND m_hFrame;
	HWND m_hParent;
	BEGIN_MSG_MAP(CTangramAFXHelperWnd)
		MESSAGE_HANDLER(WM_SHOWWINDOW, OnShowWindow)
		MESSAGE_HANDLER(WM_TANGRAMMSG, OnTangramMg)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanging)
	END_MSG_MAP()

private:
	void OnFinalMessage(HWND hWnd);
	LRESULT OnShowWindow(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnTangramMg(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnWindowPosChanging(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
};

class CTangramDocTemplate;
class CTangramMDIChildWnd;
class CTangramMDIMainWnd :
	public CWindowImpl<CTangramMDIMainWnd, CWindow>
{
public:
	CTangramMDIMainWnd(void);
	virtual ~CTangramMDIMainWnd(void);

	HWND									m_hMDIClient;
	CWndPage*								m_pPage;
	CTangramDocTemplate*					m_pDocTemplate;
	map<HWND, CString>						m_mapDesignableWnd;
	map<CString, CTangramDocTemplate*>		m_mapTangramDocTemplate;
	BEGIN_MSG_MAP(CTangramMDIMainWnd)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_COMMAND, OnCommand)
		MESSAGE_HANDLER(WM_CONTROLBARCREATED, OnControlBarCreated)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGING, OnWindowPosChanging)
	END_MSG_MAP()
	
	void OnFinalMessage(HWND hWnd);

	void OnCreateDoc(CString strDocData);
	void NavigateKey(CString strKey);
	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnCommand(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL&);
	LRESULT OnControlBarCreated(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL&);
	LRESULT OnWindowPosChanging(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
};

class CTangramMDIChildWnd :
	public CWindowImpl<CTangramMDIChildWnd, CWindow>
{
public:
	CTangramMDIChildWnd(void);
	virtual ~CTangramMDIChildWnd(void);
	BOOL				m_bNoDocView;
	HWND				m_hChild;

	CString				m_strCaption;
	CString				m_strTemplatePath;
	CString				m_strXml;
	CWndFrame*			m_pFrame;
	map<HWND, CString> m_mapDesignableWnd;

	CTangramDocTemplate* m_pDocTemplate;
	BEGIN_MSG_MAP(CTangramMDIChildWnd)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_COMMAND, OnCommand)
		MESSAGE_HANDLER(WM_ACTIVATE, OnActivate)
		MESSAGE_HANDLER(WM_DPICHANGED, OnDpiChanged)
		MESSAGE_HANDLER(WM_TANGRAMMSG, OnTangramMsg)
		MESSAGE_HANDLER(WM_MDIACTIVATE, OnMDIActivate)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGING, OnWindowPosChanged)
		MESSAGE_HANDLER(WM_CONTROLBARCREATED, OnControlBarCreated)
	END_MSG_MAP()

	void OnCreateDoc(CString strDocData);
private:
	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnCommand(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL&);
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnActivate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnMDIActivate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& );
	LRESULT OnDpiChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL&);
	LRESULT OnTangramMsg(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& );
	LRESULT OnWindowPosChanged(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL&);
	LRESULT OnControlBarCreated(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL&);
	void OnFinalMessage(HWND hWnd);
};

class CTangramDocWnd :
	public CWindowImpl<CTangramDocWnd, CWindow>
{
public:
	CTangramDocWnd(void);
	virtual ~CTangramDocWnd(void);

	HWND m_hView;
	CString	m_strKey;
	CString m_strWndID;
	CWndFrame* m_pWndFrame;
	CWndPage* m_pParentPage;
	CTangramDocFrame* m_pDocFrame;

	map<CString, HWND> m_mapCtrlBar;
	BEGIN_MSG_MAP(CTangramDocWnd)
		MESSAGE_HANDLER(WM_CLOSE, OnClose)
		MESSAGE_HANDLER(WM_COMMAND, OnCommand)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_TANGRAMMSG, OnTangramMsg)
		MESSAGE_HANDLER(WM_TANGRAMGETXML, OnTangramGetXml)
		MESSAGE_HANDLER(WM_CONTROLBARCREATED, OnCtrlBarCreated)
	END_MSG_MAP()

	void OnFinalMessage(HWND hWnd);

private:
	void CtrlBarDocActiveNotify(ITangramDoc* pDoc, IWndNode* pNodeInDoc, IWndNode* pNodeInCtrlBar, HWND hCtrlBar);
	LRESULT OnClose(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& );
	LRESULT OnCommand(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& );
	LRESULT OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& );
	LRESULT OnTangramMsg(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& );
	LRESULT OnTangramGetXml(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& );
	LRESULT OnCtrlBarCreated(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& );
};

class ATL_NO_VTABLE CWndFrame : 
	public CComObjectRootBase,	
	public CWindowImpl<CWndFrame, CWindow>,
	public IDispatchImpl<IWndFrame, &IID_IWndFrame>
{
public:
	CWndFrame();           
	virtual ~CWndFrame();           

	BOOL									m_bTabbedMDIClient;
	BOOL									m_bDetached;
	BOOL									m_bDesignerState;
	BOOL									m_bMDIChild;
	BOOL									m_bChromeFrame;
	WndFrameType							m_nWndFrameType;

	HWND									m_hPWnd;
	HWND									m_hHostWnd;

	CString									m_strAsynKeys;
	CString									m_strFrameName;	
	CString									m_strCurrentKey;
	CString									m_strLastKey;
	CTangramDocTemplate*					m_pTangramDocTemplate;
	CEclipseWnd*							m_pWorkBenchFrame;
	map<CTangramAppProxy*, CTangramWndFrameProxy*> m_mapWndFrameProxy;

	ChromePlus::CHtmlWnd*					m_pWebWnd;
	CBKWnd*									m_pBKWnd;
	CWndPage*								m_pPage;
	CWndNode*								m_pParentNode;
	CWndNode*								m_pWorkNode;
	CWndNode*								m_pHostNode;
	CWndNode*								m_pBindingNode;
	WndFrameInfo*							m_pWndFrameInfo;
	CTangramEventObj*						m_pInitEventObj;
	CTangramDoc*							m_pDoc;
	map<CString, CWndNode*>					m_mapNode;
	map<CString, CWndNode*>					m_mapNeedSaveToConfigNode;
	map<CString, VARIANT>					m_mapVal;
	map<HWND, CWPFView*>					m_mapWPFView;
	map<HWND, CWPFView*>					m_mapVisibleWPFView;
	map<CString, CWndNode*>					m_mapNodeScript;
	CComObject<CWndNodeCollection>*			m_pRootNodes;

	void Lock(){}
	void Unlock(){}
	void Destroy();
	void HostPosChanged();
	void UpdateDesignerTreeInfo();
	void CreateBKWnd();
	void UpdateVisualWPFMap(HWND, BOOL);

	CTangramXmlParse* UpdateWndNode();
	BOOL CreateWndPage();
	CWndNode* OpenXtmlDocument(CTangramXmlParse* pParse, CString strKey, CString	strFile);
	int Dispatch(CString strChannel, CString strArg1, CString strArg2);

	STDMETHOD(get_FrameXML)(BSTR* pVal);
	STDMETHOD(ModifyHost)(LONGLONG hHostWnd);
	STDMETHOD(Extend)(BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode);
	STDMETHOD(GetXml)(BSTR bstrRootName, BSTR* bstrRet);
	STDMETHOD(SendToBackgroundWebProxy)(BSTR bstrType, BSTR bstrKey, BSTR bstrData);

	BEGIN_COM_MAP(CWndFrame)
		COM_INTERFACE_ENTRY(IWndFrame)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()

	BEGIN_MSG_MAP(CWndFrame)
		MESSAGE_HANDLER(WM_SHOWWINDOW, OnShowWindow)
		MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
		MESSAGE_HANDLER(WM_HSCROLL, OnHScroll)
		MESSAGE_HANDLER(WM_VSCROLL, OnHScroll)
		MESSAGE_HANDLER(WM_TANGRAMDATA, OnGetMe)
		MESSAGE_HANDLER(WM_NCDESTROY, OnNcDestroy)
		MESSAGE_HANDLER(WM_TANGRAMMSG, OnTangramMsg)
		MESSAGE_HANDLER(WM_CHROMEOPENWINDOWMSG, OnIPCMsg)
		MESSAGE_HANDLER(WM_TANGRAMACTIVEPAGE, OnTabChanged)
		MESSAGE_HANDLER(WM_PARENTNOTIFY, OnParentNotify)
		MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
		MESSAGE_HANDLER(WM_QUERYAPPPROXY, OnQueryAppProxy)
		MESSAGE_HANDLER(WM_WINDOWPOSCHANGING, OnWindowPosChanging)
		MESSAGE_HANDLER(WM_BACKGROUNDWEBPROXY_MSG, OnBackgroundWebProxyMsg)
	END_MSG_MAP()

protected:
	ULONG InternalAddRef(){ return 1; }
	ULONG InternalRelease(){ return 1; }

private:
	LRESULT OnGetMe(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnIPCMsg(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnHScroll(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnNcDestroy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnTabChanged(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnTangramMsg(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnParentNotify(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnShowWindow(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& );
	LRESULT OnMouseActivate(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnQueryAppProxy(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnWindowPosChanging(UINT, WPARAM, LPARAM, BOOL&);
	LRESULT OnBackgroundWebProxyMsg(UINT, WPARAM, LPARAM, BOOL&);

	STDMETHOD(get_Count)(long* pCount);
	STDMETHOD(get_WndNode)(VARIANT vIndex, IWndNode **ppNode);
	STDMETHOD(get__NewEnum)(IUnknown** ppVal);
	STDMETHOD(get_HWND)(LONGLONG* pVal);
	STDMETHOD(get_WndPage)(IWndPage** pVal);
	STDMETHOD(get_CurrentNavigateKey)(BSTR* pVal);
	STDMETHOD(get_VisibleNode)(IWndNode** pVal);
	STDMETHOD(get_RootNodes)(IWndNodeCollection** pNodeColletion);
	STDMETHOD(get_FrameData)(BSTR bstrKey, VARIANT* pVal);
	STDMETHOD(put_FrameData)(BSTR bstrKey, VARIANT newVal);
	STDMETHOD(get_DesignerState)(VARIANT_BOOL* pVal);
	STDMETHOD(put_DesignerState)(VARIANT_BOOL newVal);
	STDMETHOD(get_TangramDoc)(ITangramDoc** pVal);
	STDMETHOD(get_FrameType)(WndFrameType* pVal);
	STDMETHOD(get_Name)(BSTR* pVal);

	STDMETHOD(Attach)(void);
	STDMETHOD(Detach)(void);
	STDMETHOD(ExtendFromWeb)(BSTR bstrURLBase, BSTR bstrKey, BSTR bstrOrg, BSTR bstrRepo, BSTR bstrBranch, BSTR bstrPath);
	STDMETHOD(SendIPCMessage)(int MessageType, BSTR bstrMessage, BSTR bstrMessageData);

	void OnFinalMessage(HWND hWnd);
};
