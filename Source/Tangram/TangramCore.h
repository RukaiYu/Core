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

// TangramCore.h : Declaration of the CTangram

#include "tangram.h"
#include "ObjSafe.h"
#include "wpfview.h"
#include "metahost.h"
//
////using namespace UCCollaborationLib;
//using namespace OfficePlus;
//using namespace OfficePlus::LyncPlus;
#pragma once
//https://github.com/eclipse/rt.equinox.framework/tree/master/features/org.eclipse.equinox.executable.feature/library/win32

using namespace TangramCommon;

class CHelperWnd;
class CEclipseWnd;
class CTangramDocWnd;
class CTangramAppCtrl;
class CTangramEventObj;
class CTangramDocFrame;
class CTangramMDIMainWnd;
class CTangramCommonCtrl;
class CTangramHtmlTreeWnd;
class CTangramAFXHelperWnd;

struct TangramThreadInfo
{
	HHOOK						m_hGetMessageHook;
	map<HWND, CWndFrame*>		m_mapTangramFrame;
};

struct TangramWndInfo
{
	HWND m_hBindWnd;
	CString m_strInfo;
};

struct TangramFrameInfo
{
	CString m_strKey;
	CString m_strXml;
};

class ATL_NO_VTABLE CTangramEventObj :
	public CComObjectRootBase,
	public IDispatchImpl<ITangramEventObj, &IID_ITangramEventObj, &LIBID_Tangram, 1, 0>
{
public:
	CTangramEventObj();
	virtual ~CTangramEventObj();

	BEGIN_COM_MAP(CTangramEventObj)
		COM_INTERFACE_ENTRY(ITangramEventObj)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()

	void Lock() {}
	void Unlock() {}

	void Init(CString strEventName, int nEventIndex, IDispatch* pSourceDisp)
	{
		m_nEventIndex = nEventIndex;
		m_strEventName = strEventName;
		if (pSourceDisp)
		{
			m_pSourceObj = pSourceDisp;
			m_pSourceObj->AddRef();
		}
	}
	BOOL m_bNotFired;
	int m_nEventIndex;
	CString m_strEventName;
	IDispatch*	m_pSourceObj;
	map<int, IDispatch*> m_mapDisp;
	map<int, VARIANT> m_mapVar;
protected:
	ULONG InternalAddRef() { return 1; }
	ULONG InternalRelease() { return 1; }
public:
	STDMETHOD(get_EventName)(BSTR* pVal);
	STDMETHOD(put_EventName)(BSTR newVal);
	STDMETHOD(get_Object)(int nIndex, IDispatch** pVal);
	STDMETHOD(put_Object)(int nIndex, IDispatch* newVal);
	STDMETHOD(get_Value)(int nIndex, VARIANT* pVal);
	STDMETHOD(put_Value)(int nIndex, VARIANT newVal);
	STDMETHOD(get_eventSource)(IDispatch** pVal);
	STDMETHOD(put_eventSource)(IDispatch* pSource);
	STDMETHOD(get_Index)(int* nVal);
	STDMETHOD(put_Index)(int newVal);
};

class ATL_NO_VTABLE CTangramDoc :
	public CComObjectRootBase,
	public IDispatchImpl<ITangramDoc, &IID_ITangramDoc, &LIBID_Tangram, 1, 0>
{
public:
	CTangramDoc();
	virtual ~CTangramDoc();
	BEGIN_COM_MAP(CTangramDoc)
		COM_INTERFACE_ENTRY(ITangramDoc)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()

	int								m_nState;

	LONGLONG						m_llDocID;
	CString							m_strDocID;
	CString							m_strCurrentWndID;
	CString							m_strTemplateXml;
	CString							m_strMainFrameID;
	CString							m_strPath;

	CTangramAppProxy*				m_pAppProxy;
	CTangramDocWnd*					m_pActiveWnd;
	CTangramDocProxy*				m_pDocProxy;
	map<CString, CString>			m_mapWndScript;
	map<CString, CTangramDocFrame*> m_mapFrame;

	void Lock() {}
	void Unlock() {}

protected:
	ULONG InternalAddRef() { return 1; }
	ULONG InternalRelease() { return 1; }

private:
	STDMETHOD(get_TemplateXml)(BSTR* bstrDocData);
	STDMETHOD(put_TemplateXml)(BSTR newVal);
	STDMETHOD(put_DocType)(BSTR newVal);
	STDMETHOD(get_DocID)(LONGLONG* pVal);
	STDMETHOD(put_DocID)(LONGLONG newVal);
	STDMETHOD(GetFrameWndXml)(BSTR bstrWndID, BSTR* bstrWndScriptVal);
};

class CTangramMDIChildWnd;
class ATL_NO_VTABLE CTangramDocTemplate :
	public CComObjectRootBase,
	public IDispatchImpl<ITangramDocTemplate, &IID_ITangramDocTemplate, &LIBID_Tangram, 1, 0>
{
public:
	CTangramDocTemplate();
	virtual ~CTangramDocTemplate();
	BEGIN_COM_MAP(CTangramDocTemplate)
		COM_INTERFACE_ENTRY(ITangramDocTemplate)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()

	CString							m_strKey;
	CString							m_strClientKey;
	CString							m_strDocTemplatePath;
	map<CString, CString>			m_mapXml;
	map<HWND, CWndNode*>			m_mapMainPageNode;
	map<HWND, CWndFrame*>			m_mapConnectedFrame;
	map<HWND, CTangramMDIChildWnd*> m_mapTangramMDIChildWnd;

	void InitXmlData();
	bool SaveXmlData();

	void Lock() {}
	void Unlock() {}

protected:
	ULONG InternalAddRef() { return 1; }
	ULONG InternalRelease() { return 1; }

private:
	STDMETHOD(get_TemplateXml)(BSTR* bstrDocData);
	STDMETHOD(put_TemplateXml)(BSTR newVal);
	STDMETHOD(put_DocType)(BSTR newVal);
	STDMETHOD(get_DocID)(LONGLONG* pVal);
	STDMETHOD(put_DocID)(LONGLONG newVal);
	STDMETHOD(GetFrameWndXml)(BSTR bstrWndID, BSTR* bstrWndScriptVal);
};

class ATL_NO_VTABLE CTangramMessageObj :
	public CComObjectRootBase,
	public IDispatchImpl<IMessageObj, &IID_IMessageObj, &LIBID_Tangram, 1, 0>
{
public:
	CTangramMessageObj();
	virtual ~CTangramMessageObj();
	BEGIN_COM_MAP(CTangramMessageObj)
		COM_INTERFACE_ENTRY(IMessageObj)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()
	CString			m_strSipFrom;
	CString			m_strSipTo;
	CString			m_strMsg;
	IDispatch*		m_pHostDisp;
	IWndNode*		m_pSenderNode;
	void Lock() {}
	void Unlock() {}
	virtual void ProcessMsg(CString strMsg);
protected:
	ULONG InternalAddRef() { return 1; }
	ULONG InternalRelease() { return 1; }

private:
public:
	STDMETHOD(get_SenderSip)(BSTR* pVal);
	STDMETHOD(get_TargetSip)(BSTR* pVal);
	STDMETHOD(get_ReceivedMsg)(BSTR* pVal);
	STDMETHOD(get_Sender)(IDispatch** pVal);
	STDMETHOD(get_Node)(IWndNode** pVal);
};

class CTangramDocFrame
{
public:
	CTangramDocFrame();
	~CTangramDocFrame();

	CString							m_strWndID;
	CWndFrame*						m_pHostFrame;
	CTangramDoc*					m_pTangramDoc;
	CTangramDocWnd*					m_pCurrentWnd;

	map<HWND, CTangramDocWnd*>		m_mapWnd;
};

// CTangram
class ATL_NO_VTABLE CTangram : 
	public CTangramProxyBase,
	public CComObjectRootBase,
	public IConnectionPointContainerImpl<CTangram>,
	public IConnectionPointImpl<CTangram, &__uuidof(_ITangram)>,
	public IDispatchImpl<ITangram, &IID_ITangram, &LIBID_Tangram,  1,  0>
{
	friend class CWndFrame;
	friend class CEclipseWnd;
	friend class CTangramApp;
public:
	CTangram();
	virtual ~CTangram();

	BOOL									m_bOfficeApp;
	bool									m_bEclipse;
	bool									m_bLoadEclipseDelay;
	bool									m_bTangramInit;
	bool									m_bTangramCLRLoaded;
	bool									m_bEclipseInited;
	BOOL									m_bAdmin;
	bool									m_bCanClose;
	bool									m_bFrameDefaultState;
	bool									m_bOfficeAddinUnLoad;
	BOOL									m_bWinFormActived;
	BOOL									m_bEnableProcessFormTabKey;
	BOOL									m_bDeleteWndPage;
	BOOL									m_bCLRStart;
	BOOL									m_bNewFile;
	BOOL									m_bFirstDocCreated;
	BOOL									m_bUsingDefaultAppDocTemplate;
	BOOL									m_bInitChrome;
	int										m_nRef;
	int										m_nAppID;
	int 									launchMode;

	int										m_nTangramObj;
#ifdef _DEBUG
	int										m_nTangram;
	int										m_nJsObj;
	int										m_nTangramCtrl;
	int										m_nTangramFrame;
	int										m_nTangramNodeCommonData;
	int										m_nOfficeDocs;
	int										m_nOfficeDocsSheet;
#endif
	DWORD									m_nAppType;
	DWORD									m_dwThreadID;
	HWND									m_hTemplateWnd;
	HWND									m_hTemplateChildWnd;
	HWND									m_hActiveWnd;
	HWND									m_hEclipseHideWnd;
	HHOOK									m_hCBTHook;
	HHOOK									m_hForegroundIdleHook;
	TCHAR									m_szBuffer[MAX_PATH];

	LPCTSTR									m_lpszSplitterClass;

	CString									m_strLibs;
	CString									m_strAsynKeys;
	CString									m_strCreatorIDs;
	CString									m_strCurrentKey;
	CString									m_strWorkBenchStrs;
	CString									m_strStartupCLRObj;
	CString									m_strCurrentEclipsePagePath;

	CImageList								m_DocImageList;
	CImageList								m_DocTemplateImageList;

	CWndFrame*								m_pDocTemplateFrame;
	ChromeEclipseProxy*						m_pChromeEclipseProxy;
	CTangramDocWnd*							m_pActiveDocWnd;
	CTangramMDIMainWnd*						m_pMDIMainWnd;
	CTangramDocTemplate*					m_pActiveTemplate;
	CTangramMDIChildWnd*					m_pActiveMDIChildWnd;
	TangramDocTemplateInfo*					m_pTangramDocTemplateInfo;

	CWndNode*								m_pActiveNode;
	CWndNode*								m_pDesignRootNode;
	CWndNode*								m_pDesignWindowNode;
	CWndNode*								m_pHostDesignUINode;
	CWndFrame*								m_pWndFrame;
	CWndFrame*								m_pDesignerFrame;
	CWndFrame*								m_pDesigningFrame;
	CWndPage*								m_pPage;
	CWndPage*								m_pDesignerWndPage;

	CEventProxy*							m_pEventProxy;
	ILyncExtender*							m_pLyncAppProxy;
	CComPtr<ITypeInfo>						m_pEventTypeInfo;
	CTangramAppCtrl*						m_pTangramAppCtrl;

	CTangramHtmlTreeWnd*					m_pDocDOMTree;
	CEclipseWnd*							m_pActiveEclipseWnd;

	CLSID									m_RemoteObjClsid;
	CWindow									m_RemoteObjHelperWnd;

	//.NET Version 4: 
	ICLRRuntimeHost*						m_pClrHost;

	map<CString, CString>					m_mapHeaders;
	map<CString, CTangramDoc*>				m_mapOpenDoc;
	map<CString, WindowEventType>			m_mapEventType;
	map<void*, IUnknown*>					m_mapObjects;
	map<HWND, CWndFrame*>					m_mapBKFrame;
	map<CString, IDispatch*>				m_mapUCMABot;
	map<HWND, CTangramDocWnd*>				m_mapMDTFrame;
	map<HWND, CWndPage*>					m_mapFramePage;
	map<HWND, CWndPage*>					m_mapWindowPage;
	map<IDispatch*, CString>				m_mapObjEventDic;
	map<HWND, CEclipseWnd*>					m_mapWorkBenchWnd;
	map<LONGLONG, CTangramEventObj*>		m_mapTangramEvent;
	map<HWND, CTangramCommonCtrl*>			m_mapTangramCommonCtrl;
	vector<HWND>							m_vecEclipseHideTopWnd;
	map<LONGLONG, CTangramMessageObj*>		m_mapTangramMessageObj;
	map<HWND, CTangramMDIChildWnd*>			m_mapTangramMDIChildWnd;
	map<HWND, CTangramAFXHelperWnd*>		m_mapTangramAFXHelperWnd;
	CMapStringToPtr							m_TabWndClassInfoDictionary;
	map<CString, CHelperWnd*>				m_mapRemoteTangramHelperWnd;

	BEGIN_COM_MAP(CTangram)
		COM_INTERFACE_ENTRY(ITangram)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(CTangram)
		CONNECTION_POINT_ENTRY(__uuidof(_ITangram))
	END_CONNECTION_POINT_MAP()

	BEGIN_CATEGORY_MAP(CTangram)
		IMPLEMENTED_CATEGORY(CATID_SafeForInitializing)
		IMPLEMENTED_CATEGORY(CATID_SafeForScripting)
	END_CATEGORY_MAP()

	STDMETHOD(get_Application)(IDispatch** pVal);
	STDMETHOD(put_Application)(IDispatch* newVal);
	STDMETHOD(get_RootNodes)(IWndNodeCollection** pNodeColletion);
	STDMETHOD(get_CurrentActiveNode)(IWndNode** pVal);
	STDMETHOD(get_CreatingNode)(IWndNode** pVal);
	STDMETHOD(get_DesignNode)(IWndNode** pVal);
	STDMETHOD(get_AppExtender)(BSTR bstrKey, IDispatch** pVal);
	STDMETHOD(put_AppExtender)(BSTR bstrKey, IDispatch* newVal);
	STDMETHOD(get_AppKeyValue)(BSTR bstrKey, VARIANT* pVal);
	STDMETHOD(put_AppKeyValue)(BSTR bstrKey, VARIANT newVal);
	STDMETHOD(get_RemoteHelperHWND)(LONGLONG* pVal);
	STDMETHOD(put_RemoteHelperHWND)(LONGLONG newVal) { m_hHostWnd = (HWND)newVal; return S_OK; };
	STDMETHOD(get_LyncExtender)(ILyncExtender** pVal);
	STDMETHOD(get_TangramDoc)(LONGLONG AppProxy, LONGLONG nDocID, ITangramDoc** pVal);
	STDMETHOD(get_HostWnd)(LONGLONG* pVal);
	STDMETHOD(get_RemoteTangram)(BSTR bstrID, ITangram** pVal);
	STDMETHOD(get_Extender)(ITangramExtender** pVal);
	STDMETHOD(put_Extender)(ITangramExtender* newVal);
	STDMETHOD(get_DocTemplate)(BSTR bstrKey, LONGLONG* pVal);
	STDMETHOD(get_ActiveChromeBrowserWnd)(IChromeWebBrowser** ppChromeWebBrowser) { return S_OK; };

	STDMETHOD(ActiveCLRMethod)(BSTR bstrObjID, BSTR bstrMethod, BSTR bstrParam, BSTR bstrData);
	STDMETHOD(AttachObjEvent)(IDispatch* pDisp, int nEventIndex);
	STDMETHOD(CreateOfficeDocument)(BSTR bstrXml);// { return S_OK; };
	STDMETHOD(CreateCLRObj)(BSTR bstrObjID, IDispatch** ppDisp);
	STDMETHOD(CreateWndPage)(LONGLONG hWnd, IWndPage** ppTangram);
	STDMETHOD(CreateTangramCtrl)(BSTR bstrAppID, ITangramCtrl** ppRetCtrl);
	STDMETHOD(CreateTangramEventObj)(ITangramEventObj** ppTangramEventObj);
	STDMETHOD(DownLoadFile)(BSTR strFileURL, BSTR bstrTargetFile, BSTR bstrActionXml);
	STDMETHOD(Encode)(BSTR bstrSRC, VARIANT_BOOL bEncode, BSTR* bstrRet);
	STDMETHOD(ExportOfficeObjXml)(IDispatch* OfficeObject, BSTR* bstrXml){ return S_OK; };
	STDMETHOD(FireTangramEventObj)(ITangramEventObj* pTangramEventObj);
	STDMETHOD(GetWndFrame)(LONGLONG hHostWnd, IWndFrame** ppFrame);
	STDMETHOD(GetNodeFromHandle)(LONGLONG hWnd, IWndNode** ppRetNode);
	STDMETHOD(GetCtrlByName)(IDispatch* pCtrl, BSTR bstrName, VARIANT_BOOL bFindInChild, IDispatch** ppRetDisp);
	STDMETHOD(GetCtrlValueByName)(IDispatch* pCtrl, BSTR bstrName, VARIANT_BOOL bFindInChild, BSTR* bstrVal);
	STDMETHOD(GetDocTemplateXml)(BSTR bstrCaption, BSTR bstrPath, BSTR bstrFilter, BSTR* bstrTemplatePath);
	STDMETHOD(GetWindowClientDefaultNode)(IDispatch* pAddDisp, LONGLONG hParent, BSTR bstrWndClsName, BSTR bstrWndPageName, IWndNode** ppNode);
	STDMETHOD(GetItemText)(IWndNode* pNode, long nCtrlID, LONG nMaxLengeh, BSTR* bstrRet);
	STDMETHOD(GetCLRControl)(IDispatch* CtrlDisp, BSTR bstrNames, IDispatch** ppRetDisp);
	STDMETHOD(InitLyncApp)();
	STDMETHOD(MessageBox)(LONGLONG hWnd, BSTR bstrContext, BSTR bstrCaption, long nStyle, int* nRet);
	STDMETHOD(NewGUID)(BSTR* retVal);
	STDMETHOD(StartApplication)(BSTR bstrAppID, BSTR bstrXml);
	STDMETHOD(SetCtrlValueByName)(IDispatch* pCtrl, BSTR bstrName, VARIANT_BOOL bFindInChild, BSTR bstrVal);
	STDMETHOD(SetItemText)(IWndNode* pNode, long nCtrlID, BSTR bstrRet);
	STDMETHOD(SetHostFocus)(void);
	STDMETHOD(UpdateWndNode)(IWndNode* pNode);
	STDMETHOD(TangramGetObject)(IDispatch* SourceDisp, IDispatch** ResultDisp);
	STDMETHOD(TangramCommand)(IDispatch* RibbonControl) { return S_OK; };
	STDMETHOD(TangramGetImage)(BSTR strValue, IPictureDisp ** ppDispImage) { return S_OK; };
	STDMETHOD(TangramGetVisible)(IDispatch* RibbonControl, VARIANT* varVisible) { return S_OK; };
	STDMETHOD(TangramOnLoad)(IDispatch* RibbonControl) { return S_OK; };
	STDMETHOD(TangramGetItemCount)(IDispatch* RibbonControl, long* nCount) { return S_OK; };
	STDMETHOD(TangramGetItemLabel)(IDispatch* RibbonControl, long nIndex, BSTR* bstrLabel) { return S_OK; };
	STDMETHOD(TangramGetItemID)(IDispatch* RibbonControl, long nIndex, BSTR* bstrID) { return S_OK; };
	STDMETHOD(ExtendFrames)(LONGLONG hWnd, BSTR bstrFrames, BSTR bstrKey, BSTR bstrXml, VARIANT_BOOL bSave);
	STDMETHOD(DeletePage)(LONGLONG PageHandle);
	STDMETHOD(LoadDocComponent)(BSTR bstrLib, LONGLONG* llAppProxy);
	STDMETHOD(OpenTangramFile)(ITangramDoc** ppDoc);
	STDMETHOD(OpenTangramDocFile)(BSTR bstrFilePath, ITangramDoc** ppDoc);
	STDMETHOD(NewWorkBench)(BSTR bstrTangramDoc, IWorkBenchWindow** ppWorkBenchWindow);
	STDMETHOD(CreateOutLookObj)(BSTR bstrObjType, int nType, BSTR bstrURL, IDispatch** ppRetDisp);
	STDMETHOD(SendUCMAMessage)(BSTR sipFrom, BSTR sipTo, IDispatch* SenderDispObj, IWndNode* pSender, BSTR bstrMsg);
	STDMETHOD(ClearHeader)();
	STDMETHOD(InitEclipseApp)();
	STDMETHOD(ReadTextFromWeb)(BSTR bstrURL, BSTR bstrOrg, BSTR bstrRepo, BSTR bstrBranch, BSTR bstrFile, BSTR bstrTarget, LONGLONG hNotify);
	STDMETHOD(DeleteFrame)(IWndFrame* pWndFrame);
	STDMETHOD(InitCLRApp)(BSTR strInitXml, LONGLONG* llHandle);
	//STDMETHOD(Connect)(BSTR bstrURI);
	//STDMETHOD(SendMessage)(BSTR batrMsg);
	//STDMETHOD(Receive)();
	//STDMETHOD(CloseWS)();

	void Init();
	void InitChromeComponent();
	void Lock(){ EnterCriticalSection(&m_csTangramCriticalSection); }
	void Unlock(){ LeaveCriticalSection(&m_csTangramCriticalSection); }
	void InitEventDic();
	void ExitInstance();
	void InitDesignerTreeCtrl(CString strXml);
	void FireTangramAppEvent(CTangramEventObj*);
	void InitTangramDocManager();
	void TangramInit();
	bool ImportTangramDocTemplate(CString strFile);
	BOOL IsUserAdministrator();
	BOOL LoadImageFromResource(ATL::CImage *pImage, HMODULE hMod, CString strResID, LPCTSTR lpTyp);
	BOOL LoadImageFromResource(ATL::CImage *pImage, HMODULE hMod, UINT nResID, LPCTSTR lpTyp);
	int	 LoadCLR();
	CString EncodeFileToBase64(CString strSRC);
	CString InitEclipse(_TCHAR* jarFile);
	CString ComputeHash(CString source);
	CString GetDesignerInfo(CString);
	CString GetXmlData(CString strName, CString strXml);
	CString GetDesignerData(CWndNode* pNode);
	CString GetDocTemplateXml(CString strCaption, CString strPath, CString strFilter);
	CString GetPropertyFromObject(IDispatch* pObj, CString strPropertyName);
	CString	BuildSipURICodeStr(CString strURI, CString strPrev, CString strFix, CString strData, int n1);
	CString	GetDataFromStr(CString strCoded, CString& strTime, CString strPrev, CString strFix, int n1);
	LRESULT Close(void);
	CWndNode* ExtendEx(long hHostMainWnd, CString strExXml, CString strXTMLFile);
	TangramThreadInfo* GetThreadInfo(DWORD dwInfo = 0);

	virtual void CreateCommonDesignerToolBar();
	virtual void ProcessMsg(LPMSG lpMsg);
	virtual void OnOpenDoc(WPARAM) {};
	virtual void UpdateOfficeObj(IDispatch* pObj, CString strXml, CString strName) {};
	virtual void WindowCreated(CString strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd) {};
	virtual void WindowDestroy(HWND hWnd) {};
	virtual HRESULT COMObjCreated(REFCLSID rclsid, LPVOID pv) { return 0; };
	virtual HRESULT RemoteObjCreated(CString strID, void** ppvObject) { return 0; };

	void _addObject(void* pThis, IUnknown* pUnknown)
	{
		m_mapObjects.insert(make_pair(pThis, pUnknown));
	}

	void _removeObject(void* pThis)
	{
		auto it = m_mapObjects.find(pThis);
		if (it != m_mapObjects.end())
		{
			m_mapObjects.erase(it);
		}
	}

	void _clearObjects()
	{
		map<void*, IUnknown*>::iterator it;
		while ((it = m_mapObjects.begin()) != m_mapObjects.end())
		{
			IUnknown* pUnknown = it->second;
			CCommonFunction::ClearObject<IUnknown>(pUnknown);
			m_mapObjects.erase(it);
		}
	}

	CString GetNewGUID();

protected:
	CString	m_strDefaultTemplateXml;
	ULONG InternalAddRef(){ return 1; }
	ULONG InternalRelease();

private:
	CString								m_strOfficeAppIDs;
	CString								m_strExcludeAppExtenderIDs;
	CWindow								m_HelperWnd;
	CComObject<CWndNodeCollection>*		m_pRootNodes;
	map<DWORD, TangramThreadInfo*>		m_mapThreadInfo;
	CRITICAL_SECTION					m_csTangramCriticalSection;

	void EclipseInit();
	bool CheckUrl(CString&   url);
	void AddClsInfo(CString m_strObjID, int nType, CRuntimeClass* pClsInfo);
	void AttachNode(void* pNodeEvents);
	void OnEvent(IEventProxy* pEvent, IDispatch* pCtrlDisp, IDispatch* pArgDisp);
	CString Encode(CString strSRC, BOOL bEnCode);
	CString GetNewLayoutNodeName(BSTR strCnnID, IWndNode* pDesignNode);
	IWndFrame* ConnectPage(HWND, CString, IWndPage* pPage, WndFrameInfo*);
	IWndPage* ExtendFrame(HWND, CString strName, CString strKey);
	IWndNode* ExtendCtrl(__int64 handle, CString name, CString NodeTag);
	bool IsMDIFrameNode(IWndNode*);

	int CalculateByteMD5(BYTE* pBuffer, int BufferSize, CString &MD5);
	void FireNodeEvent(int nIndex, CWndNode* pNode, CTangramEventObj* pObj);
	void GetTangramInfo(CString strFile, TangramDocInfo*);

	CString RemoveUTF8BOM(CString strUTF8);
};
