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
* OUTLINED IN THE GPL LICENSE AGREEMENT.TANGRAM TEAM
* GRANTS TO YOU (ONE SOFTWARE DEVELOPER) THE LIMITED RIGHT TO USE
* THIS SOFTWARE ON A SINGLE COMPUTER.
*
* CONTACT INFORMATION:
* mailto:tangramteam@outlook.com
* https://www.tangramteam.com
*
*
********************************************************************************/

// WndNode.cpp : Implementation of CWndNode

#include "stdafx.h"
#include "TangramApp.h"
#include "TangramCore.h"
#include "NodeWnd.h"
#include "SplitterWnd.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "TangramHtmlTreeWnd.h"
#include "TangramHtmlTreeExWnd.h"
#include "ProgressFX.h"
#include "HourglassFX.h"
#include "TangramTreeView.h"
#include "TangramCoreEvents.h"
#include "SplitterWnd.h"
#include "EclipsePlus\EclipseAddin.h"
#include "TangramCtrl.h"
#include "TangramListView.h"
#include "WPFView.h"
#include "tangram.c"
#include "TangramTabCtrl.h"
#include "OfficePlus\OfficeAddin.h"
#include "chromium/BrowserWnd.h"
#include "chromium/HtmlWnd.h"

// CWndNodeWBEvent
CWndNodeWBEvent::CWndNodeWBEvent(CWndNode* pNode)
{
	//_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);
	//_CrtSetBreakAlloc(578);
	m_pWndNode			= pNode;
	m_pHTMLDocument2	= nullptr;
	DispEventAdvise(m_pWndNode->m_pDisp);
}

void CWndNodeWBEvent::OnNavigateComplete2(IDispatch* pDisp, VARIANT* URL)
{
	if (m_pWndNode->m_pDisp == pDisp)
	{
		ConnectJSEngine(NULL);
		CWndPage* pPage = m_pWndNode->m_pTangramNodeCommonData->m_pFrame->m_pPage;
		if (m_pHTMLDocument2)
		{
			auto it = pPage->m_mapWebPageExtender.find(m_pHTMLDocument2);
			if (it != pPage->m_mapWebPageExtender.end())
			{
				pPage->m_mapWebPageExtender.erase(it);
			}
		}
		CComQIPtr<IWebBrowser2> pWeb2(pDisp);
		CComPtr<IDispatch> pDispDoc;
		pWeb2->get_Document(&pDispDoc);
		CComQIPtr<IHTMLDocument2> pHtmlDocument2(pDispDoc);
		if (pHtmlDocument2&&pPage)
		{
			m_pHTMLDocument2 = pHtmlDocument2.Detach();
			pPage->m_mapWebPageExtender[m_pHTMLDocument2] = this;
			CComPtr<IDispatch> pScriptDisp;
			CComPtr<IDispatchEx> _pScriptEx;
			if (m_pHTMLDocument2->get_Script(&pScriptDisp) == S_OK&&pScriptDisp->QueryInterface<IDispatchEx>(&_pScriptEx) == S_OK)
			{
				ConnectJSEngine(_pScriptEx);
				AddObject(_T("Tangram"), g_pTangram);
				AddObject(_T("WndPage"), (IWndPage*)pPage);
				SinkObject(_T("WndPage_On"), (IWndPage*)pPage);

				for (auto it2 : pPage->m_mapExternalObj)
				{
					AddObject(it2.first, (IDispatch*)it2.second);
					SinkObject(it2.first + _T("_On"), (IWndPage*)it2.second);
				}
			}
		}
	}
}

void CWndNodeWBEvent::OnDocumentComplete(IDispatch *pDisp, VARIANT *URL)
{
	if (m_pWndNode->m_pDisp == pDisp)
	{
		CWndPage* pPage = m_pWndNode->m_pTangramNodeCommonData->m_pPage;
		if (pPage)
		{
			bool bState = false;

			for (auto it: pPage->m_mapFrame)
			{
				for (auto it2 : it.second->m_mapNode)
				{
					if (it2.second->m_bCreated == false)
					{
						::PostMessage(pPage->m_hWnd, WM_TANGRAM_WEBNODEDOCCOMPLETE, 0, 0);
						return;
					}
				}
			}
			if (m_pHTMLDocument2)
			{
				m_pWndNode->m_bNodeDocComplete = true;
				CComPtr<IHTMLWindow2> pWinDisp;
				m_pHTMLDocument2->get_parentWindow(&pWinDisp);
				pPage->Fire_PageLoaded(pWinDisp.p, CComBSTR(m_pWndNode->m_strURL));
				CWndFrame* pFrame = m_pWndNode->m_pTangramNodeCommonData->m_pFrame;
				if (pFrame->m_pInitEventObj&&pFrame->m_pInitEventObj->m_bNotFired)
				{
					pFrame->m_pInitEventObj->m_bNotFired = false;
					pPage->Fire_TangramEvent(pFrame->m_pInitEventObj);
					pFrame->m_pInitEventObj = nullptr;
				}
			}
		}
	}
}

CWndNode::CWndNode()
{
	g_pTangram->m_nTangramObj++;
	m_nID						= 0;
	m_nRow						= 0;
	m_nCol						= 0;
	m_nRows						= 1;
	m_nCols						= 1;
	m_nViewType					= BlankView;

	m_bTopObj					= false;
	m_bWebInit					= false;
	m_bCreated					= false;
	m_bNodeDocComplete			= false;
	m_varTag.vt					= VT_EMPTY;
	m_strKey					= _T("");
	m_strXmlFileFromVS			= _T("");
	m_strURL					= _T("");
	m_strNodeName				= _T("");
	m_strExtenderID				= _T("");
	m_hHostWnd					= NULL;
	m_hChildHostWnd				= NULL;
	m_pDisp						= nullptr;
	m_pVSDocument				= nullptr;
	m_pHostFrame				= nullptr;
	m_pRootObj					= nullptr;
	m_pHostWnd					= nullptr;
	m_pExtender					= nullptr;
	m_pParentObj				= nullptr;
	m_pObjClsInfo				= nullptr;
	m_pVisibleXMLObj			= nullptr;
	m_pCLREventConnector		= nullptr;
	m_pTangramNodeWBEvent		= nullptr;
	m_pChildNodeCollection		= nullptr;
	m_pCurrentExNode			= nullptr;
	m_pHostParse				= nullptr;
	m_pDocXmlParseNode			= nullptr;
	m_pTangramNodeCommonData	= nullptr;
	m_pCurMsgObj				= nullptr;
	g_pTangram->m_pActiveNode	= this;
	if (g_pTangram->m_pTangramPackageProxy)
	{
		m_strXmlFileFromVS = g_pTangram->m_pTangramPackageProxy->m_strCurrentXtmlFilePath;
	}
}


void CWndNode::InitWndNode()
{
	m_pTangramNodeCommonData = m_pRootObj->m_pTangramNodeCommonData;
	ASSERT(m_pTangramNodeCommonData != nullptr);
	m_nHeigh = m_pHostParse->attrInt(TGM_HEIGHT, 0);
	m_nWidth = m_pHostParse->attrInt(TGM_WIDTH, 0);

	m_nActivePage = m_pHostParse->attrInt(TGM_ACTIVE_PAGE, 0);
	m_strCaption = m_pHostParse->attr(TGM_CAPTION, _T(""));
	m_strName = m_pHostParse->attr(TGM_NAME, _T(""));
	if (m_strName == _T(""))
	{
		m_strName.Format(_T("Node_%p"), (LONGLONG)this);
	}
	if (m_pTangramNodeCommonData->m_pFrame&&m_pTangramNodeCommonData->m_pFrame->m_pPage)
	{
		m_strNodeName = m_strName + _T("@") + g_pTangram->m_strCurrentKey + _T("@") + m_pTangramNodeCommonData->m_pFrame->m_strFrameName;
		auto it2 = m_pTangramNodeCommonData->m_pPage->m_mapNode.find(m_strNodeName);
		if (it2 == m_pTangramNodeCommonData->m_pPage->m_mapNode.end())
		{
			m_pTangramNodeCommonData->m_pPage->m_mapNode[m_strNodeName] = this;
		}
	}
	m_strID = m_pHostParse->attr(TGM_NODE_TYPE, _T(""));
	m_strID.MakeLower();
	m_strID.Trim();
	m_strCnnID = m_pHostParse->attr(TGM_CNN_ID, _T(""));
	m_strCnnID.MakeLower();
	m_strCnnID.Trim();
	if (m_strID==_T("")&&m_strCnnID==_T("")&&m_strXmlFileFromVS != _T(""))
	{
		CString strLocalurl = m_pHostParse->attr(_T("localurl"), _T(""));
		if (strLocalurl != _T("") && ::PathFileExists(m_strXmlFileFromVS))
		{
			int nPos = m_strXmlFileFromVS.ReverseFind('\\');
			CString strPath = m_strXmlFileFromVS.Left(nPos + 1) + strLocalurl;
			if (::PathFileExists(strPath))
			{
				m_strID = _T("activex");
				m_strCnnID = strPath;
			}
		}
	}

	TangramWndClsInfo* pWndClsInfo = nullptr;
	if (g_pTangram->m_TabWndClassInfoDictionary.Lookup(LPCTSTR(m_strID), (void *&)pWndClsInfo))
		m_pObjClsInfo = pWndClsInfo->m_pTabWndClsInfo;

	if (m_pObjClsInfo == nullptr)
		m_pObjClsInfo = RUNTIME_CLASS(CNodeWnd);

	for (auto it : g_pTangram->m_mapTangramAppProxy)
	{
		CTangramWndNodeProxy* pTangramWndNodeProxy = it.second->OnTangramNodeInit(this);
		if (pTangramWndNodeProxy)
			m_mapWndNodeProxy[it.second] = pTangramWndNodeProxy;
	}
}

CWndNode::~CWndNode()
{
	if (g_pTangram->m_pActiveNode == this)
		g_pTangram->m_pActiveNode = nullptr;
	if (m_pTangramNodeCommonData->m_pOldFrame)
		m_pTangramNodeCommonData->m_pFrame = m_pTangramNodeCommonData->m_pOldFrame;
	CWndFrame* pFrame = m_pTangramNodeCommonData->m_pFrame;
	if (pFrame->m_pPage)
	{
		auto it = pFrame->m_pPage->m_mapNode.find(m_strNodeName);
		if (it != pFrame->m_pPage->m_mapNode.end())
			pFrame->m_pPage->m_mapNode.erase(it);
	}
	if (pFrame->m_pWorkNode == this)
		pFrame->m_pWorkNode = nullptr;
	if (m_strKey!=_T(""))
	{
		auto it = pFrame->m_mapNode.find(m_strKey);
		if (it != pFrame->m_mapNode.end())
		{
			//BOOL bDeleteFrame = FALSE;
			pFrame->m_mapNode.erase(it);
			if (pFrame->m_mapNode.size() == 0)
			{
				if (::IsWindow(pFrame->m_hWnd))
				{
					pFrame->UnsubclassWindow(true);
					pFrame->m_hWnd = NULL;
				}
				delete pFrame;
			}
			delete m_pTangramNodeCommonData;
			m_pTangramNodeCommonData = nullptr;
		}
	}

	if (m_pCLREventConnector)
	{
		delete m_pCLREventConnector;
		m_pCLREventConnector = nullptr;
	}

	g_pTangram->m_nTangramObj--;
#ifdef _DEBUG
#endif
	HRESULT hr = S_OK;
	DWORD dw = 0;

	if (m_pExtender)
	{
		dw = m_pExtender->Release();
		m_pExtender = nullptr;
	}

	if (m_nViewType != TangramTreeView&&m_nViewType != Splitter&&m_pDisp)
		CCommonFunction::ClearObject<IUnknown>(m_pDisp);

	m_vChildNodes.clear();

	if (m_pChildNodeCollection != nullptr)
	{
		dw = m_pChildNodeCollection->Release();
		while(dw)
			dw = m_pChildNodeCollection->Release();
		m_pChildNodeCollection = nullptr;
	}
	for (auto it : m_mapWndNodeProxy)
	{
		if (it.second->m_bAutoDelete)
			delete it.second;
	}
	m_mapWndNodeProxy.clear();
	ATLTRACE(_T("delete CWndNode:%x\n"), this);
}

CString CWndNode::GetNames()
{
	CString strRet = _T("");
	CWndNode* pNode = this;
	if (pNode)
	{
		pNode = pNode->m_pRootObj;
		strRet += pNode->m_strName;
		strRet += _T(",");
		strRet += _GetNames(pNode);
	}
	return strRet;
}

CString CWndNode::_GetNames(CWndNode* pNode)
{
	CString strRet = _T("");
	if (pNode)
	{
		for (auto it : pNode->m_vChildNodes)
		{
			CWndNode* pChildNode = it;
			strRet += pChildNode->m_strName;
			strRet += _T(",");
			strRet += _GetNames(pChildNode);
		}
	}
	return strRet;
}

BOOL CWndNode::PreTranslateMessage(MSG* pMsg)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	if(m_pHostWnd&&pMsg->message != WM_MOUSEMOVE)
	{
		m_pHostWnd->PreTranslateMessage(pMsg);

		for(auto it : m_vChildNodes)
		{
			it->PreTranslateMessage(pMsg);
		}
	}
	return true;
}

STDMETHODIMP CWndNode::LoadXML(int nType, BSTR bstrXML)
{
	if (m_strID.CompareNoCase(_T("TreeView")) == 0)
	{
		CTangramHtmlTreeWnd* _pXHtmlTree = ((CNodeWnd*)m_pHostWnd)->m_pXHtmlTree;
		if (nType)
		{
			_pXHtmlTree->LoadXmlFromFile(OLE2T(bstrXML), CTangramHtmlTreeWnd::ConvertToUnicode);
		}
		else
		{
			_pXHtmlTree->LoadXmlFromString(OLE2T(bstrXML), CTangramHtmlTreeWnd::ConvertToUnicode);
		}
	}

	return S_OK;
}

STDMETHODIMP CWndNode::ActiveTabPage(IWndNode* _pNode)
{
	return S_OK;
}

STDMETHODIMP CWndNode::Extend(BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode)
{
	switch (m_nViewType)
	{
	case ActiveX:
	case CLRCtrl:
	case BlankView:
		{
			if (m_nViewType == BlankView)
			{
				if (m_pParentObj&&m_pParentObj->m_nViewType == Splitter)
				{
					return m_pParentObj->ExtendEx(m_nRow, m_nCol, bstrKey, bstrXml, ppRetNode);
				}
			}
			if (m_pTangramNodeCommonData->m_pPage)
			{
				if (m_nViewType == BlankView&& m_pParentObj&&m_pParentObj->m_nViewType == Splitter)
				{
					return m_pParentObj->ExtendEx(m_nRow, m_nCol, bstrKey, bstrXml, ppRetNode);
				}
				if (m_pHostFrame == nullptr)
				{
					CString strName = m_strNodeName;
					strName += _T("_Frame");

					if (m_nViewType == BlankView)
					{
						RECT rc;
						::GetClientRect(m_pHostWnd->m_hWnd, &rc);
						m_hHostWnd = ::CreateWindowEx(NULL, L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, rc.right, rc.bottom, m_pHostWnd->m_hWnd, NULL, AfxGetInstanceHandle(), NULL);
					}
					else
					{
						m_hHostWnd = ::GetWindow(m_pHostWnd->m_hWnd,GW_CHILD);
					}
					m_pTangramNodeCommonData->m_pPage->CreateFrame(CComVariant(0),CComVariant((long)m_hHostWnd),strName.AllocSysString(),&m_pHostFrame);
				}
				if (m_pHostFrame&&::IsWindow(m_hHostWnd))
				{
					CWndFrame* _pFrame = (CWndFrame*)m_pHostFrame;
					_pFrame->m_pHostNode = this;
					if (m_pTangramNodeCommonData->m_pFrame->m_pParentNode)
					{
						_pFrame->m_pParentNode = m_pTangramNodeCommonData->m_pFrame->m_pParentNode;
					}
					return m_pHostFrame->Extend(bstrKey, bstrXml, ppRetNode);
				}
			}
		}
		break;
	case Splitter:
		break;
	}
	return S_OK;
}

STDMETHODIMP CWndNode::ExtendEx(int nRow, int nCol, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode)
{
	if (m_pTangramNodeCommonData->m_pPage&&m_nViewType== Splitter)
	{
		IWndNode* pNode = nullptr;
		GetNode(nRow, nCol, &pNode);
		if (pNode == nullptr)
			return S_OK;
		CWndNode* pWindowNode = (CWndNode*)pNode;
		if (pWindowNode->m_pHostFrame == nullptr)
		{
			CString strName = pWindowNode->m_strNodeName;
			strName += _T("_Frame");

			::SetWindowLong(pWindowNode->m_pHostWnd->m_hWnd,GWL_ID,1);
			m_pTangramNodeCommonData->m_pPage->CreateFrame(CComVariant(0),CComVariant((long)pWindowNode->m_pHostWnd->m_hWnd),strName.AllocSysString(),&pWindowNode->m_pHostFrame);
			CWndFrame* _pFrame = (CWndFrame*)pWindowNode->m_pHostFrame;
			if (m_pTangramNodeCommonData->m_pFrame->m_pParentNode)
			{
				_pFrame->m_pParentNode = m_pTangramNodeCommonData->m_pFrame->m_pParentNode;
			}
			if (m_pTangramNodeCommonData->m_pFrame->m_bChromeFrame)
			{
				_pFrame->m_bChromeFrame = true;
				_pFrame->m_pWebWnd = (ChromePlus::CHtmlWnd*)::GetWindowLongPtr(m_pTangramNodeCommonData->m_pFrame->m_hWnd, GWLP_USERDATA);
				//::SetWindowLongPtr(pWindowNode->m_pHostWnd->m_hWnd, GWLP_USERDATA, (LONG_PTR)::GetWindowLongPtr(m_pTangramNodeCommonData->m_pFrame->m_hWnd, GWLP_USERDATA));
			}
			((CWndFrame*)pWindowNode->m_pHostFrame)->m_pDoc = m_pTangramNodeCommonData->m_pFrame->m_pDoc;
		}
		CWndFrame* _pFrame = (CWndFrame*)pWindowNode->m_pHostFrame;
		if (pWindowNode->m_pHostFrame)
		{
			if (pWindowNode->m_pCurrentExNode)
			{
				::SetWindowLong(pWindowNode->m_pCurrentExNode->m_pHostWnd->m_hWnd, GWL_ID,1);
			}
			HRESULT hr = pWindowNode->m_pHostFrame->Extend(bstrKey, bstrXml, ppRetNode);
			long dwID = AFX_IDW_PANE_FIRST + nRow * 16 + nCol;
			if (hr != S_OK)
			{
				if(pWindowNode->m_pCurrentExNode)
					::SetWindowLong(pWindowNode->m_pCurrentExNode->m_pHostWnd->m_hWnd, GWL_ID, dwID);
				else
					::SetWindowLong(pWindowNode->m_pHostWnd->m_hWnd, GWL_ID, dwID);
				ChromePlus::CHtmlWnd* pWebWnd = _pFrame->m_pWebWnd;
				if (pWebWnd)
				{
					::PostMessage(::GetParent(pWebWnd->m_hWnd), WM_TANGRAMMSG, 20181013, 1);
				}
				return S_OK;
			}
			CWndNode* pRootNode = (CWndNode*)*ppRetNode;
			pWindowNode->m_pCurrentExNode = pRootNode;
			::SetWindowLongPtr(pRootNode->m_pHostWnd->m_hWnd, GWLP_ID, dwID);
			ChromePlus::CHtmlWnd* pWebWnd = _pFrame->m_pWebWnd;
			if (pWebWnd)
			{
				if (_pFrame->m_pBindingNode)
					pWebWnd->m_hHostWnd = _pFrame->m_pBindingNode->m_pHostWnd->m_hWnd;
				else
					pWebWnd->m_hHostWnd = NULL;
				::PostMessage(::GetParent(pWebWnd->m_hWnd), WM_TANGRAMMSG, 20181013, 1);
			}
			return hr;
		}
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Tag(VARIANT* pVar)
{
	*pVar = m_varTag;

	if (m_varTag.vt == VT_DISPATCH && m_varTag.pdispVal)
		m_varTag.pdispVal->AddRef();
	return S_OK;
}

STDMETHODIMP CWndNode::put_Tag(VARIANT var)
{
	m_varTag = var;
	return S_OK;
}

STDMETHODIMP CWndNode::get_XObject(VARIANT* pVar)
{
	pVar->vt = VT_EMPTY;
	if (m_pDisp)
	{
		pVar->vt = VT_DISPATCH;
		pVar->pdispVal = m_pDisp;
		pVar->pdispVal->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_AxPlugIn(BSTR bstrPlugInName, IDispatch** pVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString strObjName = OLE2T(bstrPlugInName);
	strObjName.Trim();
	strObjName.MakeLower();
	IDispatch* pDisp = nullptr;
	if(m_pTangramNodeCommonData->m_PlugInDispDictionary.Lookup(LPCTSTR(strObjName),(void*&)pDisp))
	{
		* pVal = pDisp;
		(* pVal)->AddRef();
	}
	else
		*pVal = nullptr;
	return S_OK;
} 

STDMETHODIMP CWndNode::get_Name(BSTR* pVal)
{
	*pVal = m_strName.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndNode::put_Name(BSTR bstrNewName)
{
	CString strName = OLE2T(bstrNewName);
	strName.Trim();
	strName.Replace(_T(","),_T(""));
	if (m_pHostParse != NULL)
	{
		CString _strName = _T(",");
		_strName += GetNames();
		CString _strName2 = _T(",");
		_strName2 += strName;
		_strName2 += _T(",");
		int nPos = _strName.Find(_strName2);
		if (nPos == -1)
		{
			m_pHostParse->put_attr(L"name", strName);
			m_strName = strName;
		}
		else
		{
			::MessageBox(NULL,_T("Modify name failed!"),_T("Tangram"),MB_OK);
		}
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Attribute(BSTR bstrKey, BSTR* pVal)
{
	if (m_pHostParse != nullptr)
	{
		CString strVal = m_pHostParse->attr(OLE2T(bstrKey), _T(""));
		*pVal = strVal.AllocSysString();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_Attribute(BSTR bstrKey, BSTR bstrVal)
{
	if (m_pHostParse != nullptr)
	{
		CString strID = OLE2T(bstrKey);
		CString strVal = OLE2T(bstrVal);
		ATLTRACE(_T("Modify CWndNode Attribute: ID£º%s Value: %s\n"), strID, strVal);
		CWndFrame* pFrame = nullptr;
		if (strVal.CompareNoCase(_T("hostview")) == 0 && g_pTangram->m_pDesignWindowNode)
		{
			pFrame = g_pTangram->m_pDesignWindowNode->m_pRootObj->m_pTangramNodeCommonData->m_pFrame;
			if (g_pTangram->m_pMDIMainWnd&&g_pTangram->m_pActiveTemplate == nullptr&&pFrame->m_hWnd == g_pTangram->m_pMDIMainWnd->m_hMDIClient)
			{
				::MessageBox(nullptr, _T("Default UI Don't have a MDI Client!"), _T("Tangram"), MB_OK);
				return S_FALSE;
			}
			if (g_pTangram->m_pDesignWindowNode&&pFrame->m_pBindingNode)
			{
				CWndNode* pOldNode = pFrame->m_pBindingNode;
				if (pOldNode->m_pTangramNodeCommonData->m_pOldFrame)
				{
					pOldNode->m_pTangramNodeCommonData->m_pFrame = pOldNode->m_pTangramNodeCommonData->m_pOldFrame;
					pOldNode->m_pTangramNodeCommonData->m_pOldFrame = nullptr;
				}
				CWndNode* pParent = pOldNode->m_pParentObj;
				if (pParent&&pParent->m_nViewType == Splitter)
				{
					if (pOldNode != this)
					{
						CSplitterNodeWnd* pWnd = (CSplitterNodeWnd*)pParent->m_pHostWnd;
						pWnd->m_pHostNode = nullptr;
						if (m_pParentObj == pParent)
							pWnd->m_pHostNode = this;
					}
				}
				if (m_pParentObj&&m_pParentObj->m_nViewType == Splitter)
				{
					CSplitterNodeWnd* pWnd = (CSplitterNodeWnd*)m_pParentObj->m_pHostWnd;
					pWnd->m_pHostNode = this;
				}
				pOldNode->m_strID = _T("");
				if (pOldNode->m_pRootObj == g_pTangram->m_pDesignWindowNode->m_pRootObj)
					pOldNode->m_pHostParse->put_attr(_T("id"), _T(""));
				ATLTRACE(_T("Modify CWndNode HostView Attribute: ID£º%s Value: %s\n"), strID, strVal);
				pOldNode->m_pHostWnd->Invalidate();
				g_pTangram->UpdateWndNode(g_pTangram->m_pDesignWindowNode->m_pRootObj);
				g_pTangram->put_AppKeyValue(CComBSTR(L"TangramDesignerXml"), CComVariant(g_pTangram->m_pDesignWindowNode->m_pRootObj->m_pTangramNodeCommonData->m_pTangramParse->xml()));
			}
			m_strID = _T("hostview");
			CWndNode* pTopNode = m_pRootObj;
			pTopNode->m_pTangramNodeCommonData->m_pHostClientView = (CNodeWnd*)m_pHostWnd;
			while (pTopNode != pTopNode->m_pRootObj)
			{
				pTopNode->m_pTangramNodeCommonData->m_pFrame->m_pBindingNode = this;
				pTopNode->m_pTangramNodeCommonData->m_pHostClientView = pTopNode->m_pTangramNodeCommonData->m_pHostClientView;
				pTopNode = pTopNode->m_pRootObj;
			}
			m_pHostParse->put_attr(_T("id"), _T("hostview"));
			if (g_pTangram->m_pDesignWindowNode)
			{
				pFrame->m_pBindingNode = this;
				g_pTangram->m_pDesignWindowNode->m_pTangramNodeCommonData->m_pOldFrame = g_pTangram->m_pDesignWindowNode->m_pTangramNodeCommonData->m_pFrame;
				g_pTangram->m_pDesignWindowNode->m_pTangramNodeCommonData->m_pFrame = m_pRootObj->m_pTangramNodeCommonData->m_pFrame;
				g_pTangram->m_pDesignWindowNode->m_pTangramNodeCommonData->m_pHostClientView = m_pRootObj->m_pTangramNodeCommonData->m_pHostClientView;
			}

			if (m_pParentObj&&m_pParentObj->m_nViewType == Splitter)
				m_pHostWnd->ModifyStyleEx(WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE, 0);
			m_pTangramNodeCommonData->m_pFrame->HostPosChanged();
			if (m_pTangramNodeCommonData->m_pFrame->m_bChromeFrame)
			{
				ChromePlus::CHtmlWnd* pWebWnd = (ChromePlus::CHtmlWnd*)::GetWindowLongPtr(m_pTangramNodeCommonData->m_pFrame->m_hWnd, GWLP_USERDATA);
				CChromeTangram* pChromeTangram = (CChromeTangram*)g_pTangram;
				auto it = pChromeTangram->m_mapBrowserWnd.find(::GetParent(pWebWnd->m_hWnd));
				if (it != pChromeTangram->m_mapBrowserWnd.end()) {
					it->second->m_pBrowser->LayoutBrowser();
					it->second->ChromeDraw();
					::SendMessage(::GetParent(m_pTangramNodeCommonData->m_pFrame->m_hWnd), WM_CHROMEDRAW, 0, 0);
				}
			}
		}
		m_pHostParse->put_attr(strID, strVal);
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Caption(BSTR* pVal)
{
	*pVal = m_strCaption.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndNode::put_Caption(BSTR bstrCaption)
{
	CString str(bstrCaption);

	m_strCaption = str;

	if (m_pParentObj != nullptr && m_pParentObj->m_pHostWnd != nullptr)
	{
		m_pParentObj->m_pHostWnd->SendMessage(WM_TGM_SET_CAPTION,m_nCol,(LPARAM)str.GetBuffer());
	}

	if (m_pHostParse != nullptr)
	{
		m_pHostParse->put_attr(L"caption",str);
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Handle(LONGLONG* pVal)
{
	if(m_pHostWnd)	
		*pVal = (LONGLONG)m_pHostWnd->m_hWnd;
	return S_OK;
}

STDMETHODIMP CWndNode::get_OuterXml(BSTR* pVal)
{
	*pVal = m_pTangramNodeCommonData->m_pTangramParse->xml().AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndNode::get_Key(BSTR* pVal)
{
	*pVal = m_pRootObj->m_strKey.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndNode::get_XML(BSTR* pVal)
{
	*pVal = m_pHostParse->xml().AllocSysString();
	return S_OK;
}

BOOL CWndNode::Create(DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext)
{
	HWND hWnd = 0;
	BOOL bRet = false;

	CNodeWnd* pTangramDesignView = (CNodeWnd*)m_pHostWnd;
	BOOL isAppWnd = false;
	if (m_strID == _T("activex") || m_strID == _T("clrctrl"))
	{
		if (m_strID == _T("clrctrl") || m_strCnnID.Find(_T(",")) != -1)
		{
			g_pTangram->LoadCLR();
			m_nViewType = CLRCtrl;
		}
		else
			m_nViewType = ActiveX;
		if (m_strCnnID.Find(_T("//")) == -1 && ::PathFileExists(m_strCnnID) == false)
		{
			CString strPath = g_pTangram->m_strAppPath + _T("TangramWebPage\\") + m_strCnnID;
			if (::PathFileExists(strPath))
				m_strCnnID = strPath;
		}

		hWnd = CreateView(pParentWnd->m_hWnd, m_strCnnID);
		if (m_pDisp)
		{
			CComBSTR bstrExtenderID(L"");
			get_Attribute(_T("extender"), &bstrExtenderID);
			m_strExtenderID = OLE2T(bstrExtenderID);
			m_strExtenderID.Trim();
			if (m_strExtenderID != _T(""))
			{
				CComPtr<IDispatch> pDisp;
				pDisp.CoCreateInstance(bstrExtenderID);
				if (pDisp)
				{
					m_pExtender = pDisp.Detach();
					m_pExtender->AddRef();
				}
			}

			pTangramDesignView->m_bCreateExternal = true;
		}
		bRet = true;
	}
	else
	{
		CString str = g_pTangram->m_strExeName;
		int nPos = m_strID.Find(_T("@"));
		if (nPos != -1)
		{
			CTangramAppProxy* pProxy = nullptr;
			CString strKey = m_strID.Mid(nPos + 1);
			auto it = g_pTangram->m_mapTangramAppProxy.find(strKey);
			if (it != g_pTangram->m_mapTangramAppProxy.end())
			{
				pProxy = it->second;
			}
			else
			{
				CString strPath = g_pTangram->m_strAppPath;
				nPos = strKey.Find(_T("."));
				CString strAppName = strKey.Left(nPos);
				HMODULE hHandle = nullptr;
				CString strdll = strPath + strKey + _T("\\") + strAppName + _T(".dll");
				if (::PathFileExists(strdll))
					hHandle = ::LoadLibrary(strdll);
				if (hHandle == nullptr)
				{
					strdll = g_pTangram->m_strAppCommonDocPath2 + strKey + _T("\\") + strAppName + _T(".dll");
					if (::PathFileExists(strdll))
						hHandle = ::LoadLibrary(strdll);
				}
				if (hHandle)
				{
					it = g_pTangram->m_mapTangramAppProxy.find(strKey.MakeLower());
					if (it != g_pTangram->m_mapTangramAppProxy.end())
					{
						pProxy = it->second;
					}
				}
				if (pProxy)
				{
					str = m_strID.Left(nPos);
					hWnd = pProxy->CreateWindowObj(LPCTSTR(str), this, pParentWnd->m_hWnd);
					if (hWnd)
					{
						isAppWnd = true;
					}
				}
			}
		}

		if(::IsWindow(hWnd)==false)
		{
			ICreator* pViewFactoryDisp = nullptr;
			auto it = g_pTangram->m_mapObjDic.find(m_strCnnID);
			if (it != g_pTangram->m_mapObjDic.end())
			{
				pViewFactoryDisp = (ICreator*)it->second;
			}
			else
			{
				CString _strKey = m_strCnnID;
				_strKey += _T(",");
				if (g_pTangram->m_strCreatorIDs.Find(_strKey) == -1)
				{
					if (m_strCnnID == _T("hostapp"))
					{
						if (g_pTangram->m_pAppDisp)
						{
							g_pTangram->m_pAppDisp->QueryInterface(IID_ICreator, (void**)&pViewFactoryDisp);
							if (pViewFactoryDisp)
								g_pTangram->m_mapObjDic[m_strCnnID] = pViewFactoryDisp;
						}
					}
					else
					{
						if (m_strID.CompareNoCase(_T("TreeView")))
						{
							CComPtr<ICreator> pFactoryDisp;
							HRESULT hr = pFactoryDisp.CoCreateInstance(m_strCnnID.AllocSysString());
							if (hr == S_OK)
							{
								pViewFactoryDisp = pFactoryDisp.Detach();
								pViewFactoryDisp->AddRef();
								g_pTangram->m_mapObjDic[m_strCnnID] = pViewFactoryDisp;
							}
						}
					}
					g_pTangram->m_strCreatorIDs += _strKey;
				}
			}

			if (pViewFactoryDisp)
			{
				long _nRet = 0;
				pViewFactoryDisp->Create(pParentWnd ? (long)pParentWnd->m_hWnd : 0, this, &_nRet);
				hWnd = (HWND)_nRet;
				m_nID = ::GetWindowLong(hWnd, GWL_ID);
			}
		}
	}

	if (!::IsWindow(m_pHostWnd->m_hWnd) && hWnd&&pTangramDesignView->SubclassWindow(hWnd))
	{
		if(isAppWnd==false)
			::SetWindowLong(hWnd, GWL_STYLE, dwStyle | WS_CHILD | /*WS_VISIBLE | */WS_CLIPCHILDREN | WS_CLIPSIBLINGS);
		else
		{
			if(m_pParentObj&&m_pParentObj->m_nViewType==Splitter)
				m_pHostWnd->ModifyStyleEx(0, WS_EX_CLIENTEDGE);
		}
		::SetWindowLong(hWnd, GWL_ID, nID);

		pTangramDesignView->m_bCreateExternal = true;
		bRet = true;
	}

	if (hWnd == 0)
	{
		hWnd = CreateWindow(L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, pParentWnd->m_hWnd, (HMENU)nID, AfxGetInstanceHandle(), NULL);
		if (::IsWindow(m_pHostWnd->m_hWnd) == false)
			bRet = m_pHostWnd->SubclassWindow(hWnd);
		if(m_nViewType==BlankView)
		{
			CString strURL = m_pHostParse->attr(_T("url"), _T(""));
			if (strURL != _T(""))
			{
				g_pTangram->m_pActiveNode = this;
				g_pTangram->InitChromeComponent();
				g_pTangram->m_pCLRProxy->NavigateURL(this, strURL, nullptr);
				CString strPath = g_pTangram->m_strAppPath + _T("TangramBrowser.dll");
				if (::PathFileExists(strPath))
				{
					DWORD dwRetCode = 0;
					HRESULT hrStart = g_pTangram->m_pClrHost->ExecuteInDefaultAppDomain(
						strPath,
						_T("TangramBrowser.TangramChromeProxy"),
						_T("ChromeInit"),
						CComBSTR(""),
						&dwRetCode);
				}
			}
		}
	}

	if (m_strID.CompareNoCase(_T("TreeView")) == 0)
	{
		CComBSTR bstrStyle(L"");
		get_Attribute(CComBSTR(L"Style"), &bstrStyle);
		m_nViewType = TangramTreeView;
		CString _strStyle = OLE2T(bstrStyle);
		if (_strStyle != _T(""))
		{
			pTangramDesignView->m_pXHtmlTree = new CProgressFX< CHourglassFX< CTangramHtmlTreeEx2Wnd > >;
		}
		else
			pTangramDesignView->m_pXHtmlTree = new CTangramHtmlTreeWnd();
		m_pDisp = pTangramDesignView->m_pXHtmlTree->m_pObj;
		m_pDisp->AddRef();
		pTangramDesignView->m_pXHtmlTree->m_pHostWnd = pTangramDesignView;
		bRet = true;

		DWORD dwStyle = TVS_HASBUTTONS | TVS_HASLINES | TVS_LINESATROOT |
			TVS_SHOWSELALWAYS | /*TVS_EDITLABELS |TVS_NOTOOLTIPS |*/
			WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP;

		CRect rect(0, 0, 100, 100);
		VERIFY(pTangramDesignView->m_pXHtmlTree->Create(dwStyle, rect, pTangramDesignView, 100));

		CComBSTR bstrCheckBox(L"");
		get_Attribute(CComBSTR(L"CheckBoxes"), &bstrCheckBox);
		CString strCheckBox = OLE2T(bstrCheckBox);
		CComBSTR bstrSmartCheckBox(L"");
		get_Attribute(CComBSTR(L"SmartCheckBox"), &bstrSmartCheckBox);
		CString strSmartCheckBox = OLE2T(bstrSmartCheckBox);
		CComBSTR bstrSetHtml(L"");
		get_Attribute(CComBSTR(L"SetHtml"), &bstrSetHtml);
		CString strSetHtml = OLE2T(bstrSetHtml);
		CComBSTR bstrStripHtml(L"");
		get_Attribute(CComBSTR(L"StripHtml"), &bstrStripHtml);
		CString strStripHtml = OLE2T(bstrStripHtml);
		CComBSTR bstrImages(L"");
		get_Attribute(CComBSTR(L"Images"), &bstrImages);
		CString strImages = OLE2T(bstrImages);
		CComBSTR bstrReadOnly(L"");
		get_Attribute(CComBSTR(L"ReadOnly"), &bstrReadOnly);
		CString strReadOnly = OLE2T(bstrReadOnly);

		int r = 0, g = 0, b = 0;
		CComBSTR bstrBKColor(L"");
		get_Attribute(CComBSTR(L"BKColor"), &bstrBKColor);
		CString strBKColor = OLE2T(bstrBKColor);

		COLORREF colorBK = RGB(255, 255, 255);
		if (strBKColor != _T(""))
		{
			_stscanf_s(strBKColor, _T("RGB(%d,%d,%d)"), &r, &g, &b);
			colorBK = RGB(r, g, b);
		}

		CComBSTR bstrSeparatorColor(L"");
		get_Attribute(CComBSTR(L"SeparatorColor"), &bstrSeparatorColor);
		CString strSeparatorColor = OLE2T(bstrSeparatorColor);

		COLORREF colorSeparator = RGB(0, 0, 255);
		if (bstrSeparatorColor != _T(""))
		{
			_stscanf_s(strSeparatorColor, _T("RGB(%d,%d,%d)"), &r, &g, &b);
			colorSeparator = RGB(r, g, b);
		}
		pTangramDesignView->m_pXHtmlTree->Initialize(strCheckBox != _T("") ? true : false, true)
			.SetSmartCheckBox(strSmartCheckBox != _T("") ? true : false)
			.SetHtml(true)
			//.SetHtml(strSetHtml != _T("") ? true : false)
			.SetStripHtml(strStripHtml != _T("") ? true : false)
			.SetImages(strImages != _T("") ? true : false)
			.SetReadOnly(strReadOnly != _T("") ? true : false)
			.SetDragOps(XHTMLTREE_DO_DEFAULT)
			.SetSeparatorColor(colorSeparator).SetBkColor(colorBK);
		//.SetDropCursors(IDC_NODROP, IDC_DROPCOPY, IDC_DROPMOVE);
		if (strImages != _T(""))
		{
			CComBSTR bstrImageURL(L"");
			get_Attribute(CComBSTR(L"ImageURL"), &bstrImageURL);
			pTangramDesignView->m_pXHtmlTree->m_strImageURL = OLE2T(bstrImageURL);
			get_Attribute(CComBSTR(L"ImageTarget"), &bstrImageURL);
			CString strImage = OLE2T(bstrImageURL);
			if (strImage != _T(""))
			{
				//
				CString strPath = g_pTangram->m_strTempPath;
				strPath += _T("TangramTreeNode");// OLE2T(bstrImageURL);
				strPath += strImage;
				pTangramDesignView->m_pXHtmlTree->m_strImageTarget = strPath;
				int nPos = strPath.ReverseFind('\\');
				CString strDir = strPath.Left(nPos);
				SHCreateDirectory(NULL, strDir);
				if (::PathFileExists(strPath) == false)
				{
				}
				else
				{
					CImage image;
					image.Load(strPath);
					int nWidet = image.GetWidth();
					int nHeight = image.GetHeight();

					COLORREF color = image.GetTransparentColor();

					HBITMAP hBitmap;
					CBitmap* pBitmap;
					hBitmap = (HBITMAP)::LoadImage(::AfxGetInstanceHandle(), (LPCTSTR)strPath, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
					pBitmap = new CBitmap;
					pBitmap->Attach(hBitmap);
					pTangramDesignView->m_pXHtmlTree->m_Images.Add(pBitmap, color);
					delete pBitmap;
				}
			}
		}

		int nCol = m_pHostParse->GetCount();
		if (nCol)
		{
			CTangramXmlParse* pTreeNodeParse = m_pHostParse->GetChild(_T("TreeNode"));
			if (pTreeNodeParse)
			{
				CComBSTR bstrTag(L"");
				get_Attribute(CComBSTR(L"doctag"), &bstrTag);
				CString strTag = OLE2T(bstrTag);
				if (strTag.CompareNoCase(_T("tangramdesigner")) == 0)
				{
					CString strKey = L"tangramdesignerxml";
					auto it = g_pTangram->m_mapValInfo.find(strKey);
					if (it != g_pTangram->m_mapValInfo.end())
					{
						CString strXml = OLE2T(it->second.bstrVal);
						::VariantClear(&it->second);
						g_pTangram->m_mapValInfo.erase(it);
						//m_pRootObj->m_pDocTreeCtrl = pTangramDesignView->m_pXHtmlTree;
						CTangramXmlParse* pParse = new CTangramXmlParse();
						if (pParse->LoadXml(strXml))
						{
							pTangramDesignView->m_pXHtmlTree->m_pHostXmlParse = pParse;
							m_pRootObj->m_pDocXmlParseNode = pParse;
							pTangramDesignView->m_pXHtmlTree->m_hFirstRoot = pTangramDesignView->m_pXHtmlTree->LoadXmlFromXmlParse(pParse);
						}
					}
				}
				else
				{
					CComBSTR bstrType(L"");
					get_Attribute(CComBSTR(L"doctype"), &bstrType);
					CString strType = OLE2T(bstrType);
					if (bstrType != _T(""))
					{
						CXHtmlDraw::XHTMLDRAW_APP_COMMAND AppCommands[] =
						{
							{ pTangramDesignView->m_pXHtmlTree->m_hWnd, WM_TANGRAMDESIGNERCMD, 1992, _T("WM_TANGRAM_DESIGNERCMD") },
							{ pTangramDesignView->m_pXHtmlTree->m_hWnd, WM_TANGRAMDESIGNERCMD, 1963, _T("WM_TANGRAM_DESIGNERCMD2") },
						};

						pTangramDesignView->m_pXHtmlTree->m_Links.SetAppCommands(AppCommands, sizeof(AppCommands) / sizeof(AppCommands[0]));

						if (g_pTangram->m_pDocDOMTree == nullptr)
							g_pTangram->m_pDocDOMTree = pTangramDesignView->m_pXHtmlTree;
					}
					else
					{
						CString strXml = pTreeNodeParse->xml();
						CTangramXmlParse* pParse = new CTangramXmlParse();
						pParse->LoadXml(strXml);
						pTangramDesignView->m_pXHtmlTree->m_hFirstRoot = pTangramDesignView->m_pXHtmlTree->LoadXmlFromXmlParse(pParse);
					}
				}
			}
		}
	}
	bRet = true;

	//Very important:
	if (m_pHostWnd&&::IsWindow(m_pHostWnd->m_hWnd))
		m_pHostWnd->SendMessage(WM_INITIALUPDATE);

	////////////////////////////////////////////////////////////////////////////////////////////////
	if (g_pTangram->m_pActiveNode&&g_pTangram->m_pActiveNode->m_pTangramNodeCommonData->m_pPage)
		g_pTangram->m_pActiveNode->m_pTangramNodeCommonData->m_pPage->Fire_NodeCreated(this);

	if (m_pTangramNodeCommonData->m_pFrame->m_bChromeFrame)
	{
		IPCMsg* msg = new IPCMsg;
		msg->m_strType = _T("WndNode_Created");
		msg->m_strKey = m_strName;// m_pHostParse->xml();
		msg->m_strKey += _T("@");
		msg->m_strKey += m_pTangramNodeCommonData->m_pFrame->m_strCurrentKey;
		CString s = _T("");
		s.Format(_T("%I64d"), m_pHostWnd->m_hWnd);
		msg->m_strData = s;
		::PostMessage(m_pTangramNodeCommonData->m_pFrame->m_hWnd, WM_BACKGROUNDWEBPROXY_MSG, (WPARAM)msg, (WPARAM)0);
	}

	return bRet;
}

HWND CWndNode::CreateView(HWND hParentWnd, CString strTag)
{
	BOOL bWebCtrl = false;
	CString strURL = _T("");
	CString strID = strTag;
	auto hWnd = ::CreateWindowEx(NULL,L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE| WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, hParentWnd, NULL, AfxGetInstanceHandle(), NULL);
	CComBSTR bstr2;
	get_Attribute(CComBSTR("name"), &bstr2);
	CString strName = OLE2T(bstr2);
	switch (m_nViewType)
	{
	case ActiveX:
		{
			auto it = m_pTangramNodeCommonData->m_mapAxNodes.find(strName);
			if (it == m_pTangramNodeCommonData->m_mapAxNodes.end())
			{
				m_pTangramNodeCommonData->m_mapAxNodes[strName] = this;
			}
			else
			{
				int nCount = m_pTangramNodeCommonData->m_mapAxNodes.size();
				CString str = _T("");
				str.Format(_T("%s%d"), strName,nCount);
				it = m_pTangramNodeCommonData->m_mapAxNodes.find(str);
				while (it != m_pTangramNodeCommonData->m_mapAxNodes.end())
				{
					nCount++;
					str.Format(_T("%s%d"), strName, nCount);
					it = m_pTangramNodeCommonData->m_mapAxNodes.find(str);
				}
				m_pTangramNodeCommonData->m_mapAxNodes[str] = this;
				put_Attribute(CComBSTR("name"), str.AllocSysString());
			}
			strID.MakeLower();
			auto nPos = strID.Find(_T("http:"));
			if (nPos == -1)
				nPos = strID.Find(_T("https:"));
			if (m_pTangramNodeCommonData->m_pFrame)
			{
				CComBSTR bstr;
				get_Attribute(CComBSTR("initInfo"), &bstr);
				CString str = _T("");
				str += bstr;
				if (str != _T("") && m_pTangramNodeCommonData->m_pPage)
				{
					LRESULT hr = ::SendMessage(m_pTangramNodeCommonData->m_pPage->m_hWnd, WM_GETNODEINFO, (WPARAM)OLE2T(bstr), (LPARAM)hParentWnd);
					if (hr)
					{
						CString strInfo = _T("");
						CWindow m_wnd;
						m_wnd.Attach(hParentWnd);
						m_wnd.GetWindowText(strInfo);
						strID += strInfo;
						m_wnd.Detach();
					}
				}
			}

			if (strID.Find(_T("http://")) != -1 || strID.Find(_T("https://")) != -1)
			{
				bWebCtrl = true;
				strURL = strID;
				strID = _T("shell.explorer.2");
			}

			if (strID.Find(_T("res://")) != -1 || ::PathFileExists(strID))
			{
				bWebCtrl = true;
				strURL = strID;
				strID = _T("shell.explorer.2");
			}

			if (strID.CompareNoCase(_T("about:blank")) ==0 )
			{
				bWebCtrl = true;
				strURL = strID;
				strID = _T("shell.explorer.2");
			}

			if (m_pDisp == nullptr)
			{
				CComPtr<IDispatch> pDisp;
				HRESULT hr = pDisp.CoCreateInstance(CComBSTR(strID));
				if (hr == S_OK)
				{
					m_pDisp = pDisp.Detach();
				}
			}
		}
		break;
	case CLRCtrl:
		{
			auto it = m_pTangramNodeCommonData->m_mapCLRNodes.find(strName);
			if (it == m_pTangramNodeCommonData->m_mapCLRNodes.end())
			{
				m_pTangramNodeCommonData->m_mapCLRNodes[strName] = this;
			}
			else
			{
				int nCount = m_pTangramNodeCommonData->m_mapCLRNodes.size();
				CString str = _T("");
				str.Format(_T("%s%d"), strName, nCount);
				it = m_pTangramNodeCommonData->m_mapCLRNodes.find(str);
				while (it != m_pTangramNodeCommonData->m_mapCLRNodes.end())
				{
					nCount++;
					str.Format(_T("%s%d"), strName, nCount);
					it = m_pTangramNodeCommonData->m_mapCLRNodes.find(str);
				}
				m_pTangramNodeCommonData->m_mapCLRNodes[str] = this;
				put_Attribute(CComBSTR("name"), str.AllocSysString());
			}
			if (g_pTangram->m_pCLRProxy)
			{
				//get_Attribute(CComBSTR(""),)
				m_pDisp = g_pTangram->m_pCLRProxy->TangramCreateObject(strTag.AllocSysString(), (long)hParentWnd, this);
				CNodeWnd* pWnd = (CNodeWnd*)m_pHostWnd;
				if (m_pDisp&&pWnd->m_mapDockCtrl.size())
				{
					HWND hPage = m_pTangramNodeCommonData->m_pPage->m_hWnd;
					::SendMessage(hPage, WM_TANGRAMMSG, (WPARAM)m_pHostWnd, 1963);
				}
				if(m_pDisp==nullptr)
				{
					((CNodeWnd*)m_pHostWnd)->m_bCreateExternal = false;
					m_nViewType = BlankView;
				}
			}
		}
		break;
	}
	if (m_pDisp)
	{
		CAxWindow m_Wnd;
		m_Wnd.Attach(hWnd);
		CComPtr<IUnknown> pUnk;
		m_Wnd.AttachControl(m_pDisp, &pUnk);
		if (m_nViewType == ActiveX)
		{
			((CNodeWnd*)m_pHostWnd)->m_pWndNode = this;
			CComQIPtr<IWebBrowser2> pWebDisp(m_pDisp);
			if (pWebDisp)
			{
				bWebCtrl = true;
				m_strURL = strURL;
				if (m_strURL == _T(""))
					m_strURL = strID;
				if (m_pTangramNodeCommonData->m_pFrame)
				{
					if (m_pTangramNodeCommonData->m_pPage)
						m_pTangramNodeCommonData->m_pPage->m_nWebViewCount++;

					m_pTangramNodeWBEvent = new CWndNodeWBEvent(this);
				}
				CComPtr<IAxWinAmbientDispatch> spHost;
				LRESULT hr = m_Wnd.QueryHost(&spHost);
				if (SUCCEEDED(hr))
				{
					CComBSTR bstr;
					get_Attribute(CComBSTR("scrollbar"), &bstr);
					CString str = OLE2T(bstr);
					if (str == _T("1"))
						spHost->put_DocHostFlags(DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_THEME);//DOCHOSTUIFLAG_DIALOG|
					else
						spHost->put_DocHostFlags(/*DOCHOSTUIFLAG_DIALOG|*/DOCHOSTUIFLAG_SCROLL_NO | DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_THEME);

					if (m_strURL != _T(""))
					{
						pWebDisp->Navigate2(&CComVariant(m_strURL), &CComVariant(navNoReadFromCache | navNoWriteToCache), nullptr, nullptr, nullptr);
						m_bWebInit = true;
					}
				}
			}
			HWND hPage = m_pTangramNodeCommonData->m_pPage->m_hWnd;
			::SendMessage(hPage, WM_TANGRAMMSG, (WPARAM)((CNodeWnd*)m_pHostWnd), 1963);
		}

		CComQIPtr<IOleInPlaceActiveObject> pIOleInPlaceActiveObject(m_pDisp);
		if (pIOleInPlaceActiveObject)
			((CNodeWnd*)m_pHostWnd)->m_pOleInPlaceActiveObject = pIOleInPlaceActiveObject.Detach();
		m_Wnd.Detach();
		m_pTangramNodeCommonData->m_mapLayoutNodes[m_strName] = this;
	}

	return hWnd;
}

STDMETHODIMP CWndNode::get_ChildNodes(IWndNodeCollection** pNodeColletion)
{
	if (m_pChildNodeCollection == nullptr)
	{
		CComObject<CWndNodeCollection>::CreateInstance(&m_pChildNodeCollection);
		m_pChildNodeCollection->AddRef();		
		m_pChildNodeCollection->m_pNodes = &m_vChildNodes;
	}
	return m_pChildNodeCollection->QueryInterface(IID_IWndNodeCollection,(void**)pNodeColletion);
}

int CWndNode::_getNodes(CWndNode* pNode, CString& strName,CWndNode**ppRetNode, CWndNodeCollection* pNodes)
{
	int iCount = 0;
	if (pNode->m_strName.CompareNoCase(strName) == 0)
	{
		if (pNodes != nullptr)
			pNodes->m_pNodes->push_back(pNode);

		if (ppRetNode != nullptr && (*ppRetNode) == nullptr)
			*ppRetNode = pNode;
		return 1;
	}

	for(auto it : pNode->m_vChildNodes)
	{
		iCount += _getNodes(it,strName,ppRetNode, pNodes);
	}
	return iCount;
}

STDMETHODIMP CWndNode::Show()
{
	CWndNode* pChild = this;
	CWndNode* pParent = pChild->m_pParentObj;

	while(pParent != nullptr)
	{
		pParent->m_pHostWnd->SendMessage(WM_ACTIVETABPAGE,(WPARAM)pChild->m_nCol,(LPARAM)1);

		pChild = pParent;
		pParent = pChild->m_pParentObj;
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_RootNode(IWndNode** ppNode)
{
	if (m_pRootObj != nullptr)
		*ppNode = m_pRootObj;
	return S_OK;
}

STDMETHODIMP CWndNode::get_ParentNode(IWndNode** ppNode)
{
	*ppNode = nullptr;
	if (m_pParentObj != nullptr)
		*ppNode = m_pParentObj;

	return S_OK;
}

STDMETHODIMP CWndNode::get_NodeType(ViewType* nType)
{
	*nType = m_nViewType;
	return S_OK;
}

void CWndNode::_get_Objects(CWndNode* pNode, UINT32& nType, CWndNodeCollection* pNodeColletion)
{
	if (pNode->m_nViewType & nType)
	{
		pNodeColletion->m_pNodes->push_back(pNode);
	}

	CWndNode* pChildNode = nullptr;
	for(auto it : pNode->m_vChildNodes)
	{
		pChildNode = it;
		_get_Objects(pChildNode,nType,pNodeColletion);
	}
}

STDMETHODIMP CWndNode::get_Objects(long nType, IWndNodeCollection** ppNodeColletion)
{
	CComObject<CWndNodeCollection>* pNodes = nullptr;
	CComObject<CWndNodeCollection>::CreateInstance(&pNodes);

	pNodes->AddRef();	

	UINT32 u = nType;
	_get_Objects(this,u,pNodes);
	HRESULT hr = pNodes->QueryInterface(IID_IWndNodeCollection,(void**)ppNodeColletion);

	pNodes->Release();

	return hr;
}

STDMETHODIMP CWndNode::get_Rows(long* nRows)
{
	*nRows = m_nRows;
	return S_OK;
}

STDMETHODIMP CWndNode::get_Cols(long* nCols)
{
	*nCols = m_nCols;
	return S_OK;
}

STDMETHODIMP CWndNode::get_Row(long* nRow)
{
	*nRow = m_nRow;
	return S_OK;
}

STDMETHODIMP CWndNode::get_Col(long* nCol)
{
	*nCol = m_nCol;
	return S_OK;
}

STDMETHODIMP CWndNode::GetNode(long nRow, long nCol,IWndNode** ppTangramNode)
{
	CWndNode* pRet = nullptr;
	auto bFound = false;

	*ppTangramNode = nullptr;
	if (nRow < 0 || nCol < 0 || nRow >= m_nRows || nCol >= m_nCols) return E_INVALIDARG;

	for(auto it : m_vChildNodes)
	{
		pRet = it;
		if (pRet->m_nCol == nCol && pRet->m_nRow == nRow)
		{
			bFound = true;
			break;
		}
	}

	HRESULT hr = S_OK;
	if (bFound)
	{
		hr = pRet->QueryInterface(IID_IWndNode,(void**)ppTangramNode);
	}
	return hr;
}

STDMETHODIMP CWndNode::GetNodeByName(BSTR bstrName, IWndNodeCollection** ppNodes)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString strName(bstrName);

	CWndNode* pRetNode = nullptr;

	CComObject<CWndNodeCollection>* pNodes = nullptr;
	if (ppNodes != nullptr)
	{
		*ppNodes = nullptr;
		CComObject<CWndNodeCollection>::CreateInstance(&pNodes);
		pNodes->AddRef();
	}

	int iCount = _getNodes(this, strName, &pRetNode, pNodes);

	if (ppNodes != nullptr)
		pNodes->QueryInterface(IID_IWndNodeCollection, (void**)ppNodes);

	if (pNodes != nullptr)
		pNodes->Release();

	return S_OK;
}

STDMETHODIMP CWndNode::GetNodes(BSTR bstrName,IWndNode** ppNode,IWndNodeCollection** ppNodes,long* pCount)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CString strName(bstrName);

	CWndNode* pRetNode = nullptr;

	if (ppNode != nullptr)
		*ppNode = nullptr;

	CComObject<CWndNodeCollection>* pNodes = nullptr;
	if (ppNodes != nullptr)
	{
		*ppNodes = nullptr;
		CComObject<CWndNodeCollection>::CreateInstance(&pNodes);
		pNodes->AddRef();
	}

	int iCount = _getNodes(this,strName,&pRetNode,pNodes);

	*pCount = iCount;

	if ((iCount > 0) && (ppNode != nullptr))
		pRetNode->QueryInterface(IID_IWndNode,(void**)ppNode);	

	if (ppNodes != nullptr)
		pNodes->QueryInterface(IID_IWndNodeCollection,(void**)ppNodes);

	if (pNodes != nullptr)
		pNodes->Release();

	return S_OK;
}

BOOL CWndNode::AddChildNode(CWndNode* pNode)
{
	m_vChildNodes.push_back(pNode);
	pNode->m_pParentObj = this;
	pNode->m_pRootObj = m_pRootObj;
	return true;
}

BOOL CWndNode::RemoveChildNode(CWndNode* pNode)
{
	auto it = find(m_vChildNodes.begin(), m_vChildNodes.end(), pNode);
	if (it != m_vChildNodes.end())
	{
		m_vChildNodes.erase(it);
		return true;
	}
	return false;
}

STDMETHODIMP CWndNode::get_Frame(IWndFrame** pVal)
{
	if(m_pTangramNodeCommonData->m_pFrame)
		*pVal = m_pTangramNodeCommonData->m_pFrame;

	return S_OK;
}

STDMETHODIMP CWndNode::get_HostFrame(IWndFrame** pVal)
{
	if(m_pHostFrame)
		*pVal = m_pHostFrame;

	return S_OK;
}

STDMETHODIMP CWndNode::Refresh(void)
{
	if(m_pDisp)
	{
		CComQIPtr<IWebBrowser2> pWebCtrl(m_pDisp);
		if(pWebCtrl)
			pWebCtrl->Refresh();
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_Height(LONG* pVal)
{
	RECT rc;
	::GetClientRect(m_pHostWnd->m_hWnd,&rc);
	*pVal = rc.bottom;
	return S_OK;
}

STDMETHODIMP CWndNode::get_Width(LONG* pVal)
{
	RECT rc;
	::GetClientRect(m_pHostWnd->m_hWnd, &rc);
	*pVal = rc.right;

	return S_OK;
}

STDMETHODIMP CWndNode::get_OfficeObj(IDispatch** pVal)
{
	if (m_pTangramNodeCommonData->m_pOfficeObj)
	{
		*pVal = m_pTangramNodeCommonData->m_pOfficeObj;
		(* pVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Extender(IDispatch** pVal)
{
	if(m_pExtender)
	{
		* pVal = m_pExtender;
		(* pVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_Extender(IDispatch* newVal)
{
	if(m_pExtender)
		m_pExtender->Release();
	m_pExtender = newVal;
	m_pExtender->AddRef();

	return S_OK;
}

STDMETHODIMP CWndNode::get_VSDocument(IVSDocument** pVal)
{
	if(m_pVSDocument)
	{
		* pVal = m_pVSDocument;
	}
	return S_OK;
}
//
//STDMETHODIMP CWndNode::put_VSDocument(IVSDocument* newVal)
//{
//	m_pVSDocument = newVal;
//	return S_OK;
//}

STDMETHODIMP CWndNode::get_WndPage(IWndPage** pVal)
{
	*pVal = (IWndPage*)m_pTangramNodeCommonData->m_pFrame->m_pPage;
	return S_OK;
}

STDMETHODIMP CWndNode::get_HTMLWindow(IDispatch** pVal)
{
	if (m_pTangramNodeWBEvent&&m_pTangramNodeWBEvent->m_pHTMLDocument2)
	{
		CComPtr<IHTMLWindow2> pHTMLWindow2;
		m_pTangramNodeWBEvent->m_pHTMLDocument2->get_parentWindow(&pHTMLWindow2);
		*pVal = pHTMLWindow2.Detach();
		(*pVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_HtmlDocument(IDispatch** pVal)
{
	if (m_pTangramNodeWBEvent&&m_pTangramNodeWBEvent->m_pHTMLDocument2)
	{
		*pVal = m_pTangramNodeWBEvent->m_pHTMLDocument2;
		(*pVal)->AddRef();
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_NameAtWindowPage(BSTR* pVal)
{
	*pVal = m_strNodeName.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndNode::GetCtrlByName(BSTR bstrName, VARIANT_BOOL bFindInChild, IDispatch** ppRetDisp)
{
	if (g_pTangram->m_pCLRProxy&&m_nViewType==CLRCtrl&&m_pDisp)
		*ppRetDisp=g_pTangram->m_pCLRProxy->GetCtrlByName(m_pDisp, bstrName, bFindInChild?true:false);

	return S_OK;
}

STDMETHODIMP CWndNode::GetCtrlValueByName(BSTR bstrName, VARIANT_BOOL bFindInChild, BSTR* bstrVal)
{
	if (g_pTangram->m_pCLRProxy&&m_nViewType==CLRCtrl&&m_pDisp)
	{
		*bstrVal = g_pTangram->m_pCLRProxy->GetCtrlValueByName(m_pDisp, bstrName, bFindInChild ? true : false);
	}
	return S_OK;
}

STDMETHODIMP CWndNode::SetCtrlValueByName(BSTR bstrName, VARIANT_BOOL bFindInChild, BSTR bstrVal)
{
	if (g_pTangram->m_pCLRProxy&&m_nViewType==CLRCtrl&&m_pDisp)
	{
		g_pTangram->m_pCLRProxy->SetCtrlValueByName(m_pDisp, bstrName, bFindInChild ? true : false, bstrVal);
	}
	return S_OK;
}

CWndNodeCollection::CWndNodeCollection()
{
	m_pNodes = &m_vNodes;
	//theApp._addObject(this, this);
}

CWndNodeCollection::~CWndNodeCollection()
{
	//theApp._removeObject(this);
	if (g_pTangram->m_pCLRProxy)
	{
		//g_pTangram->m_pCLRProxy->ReleaseTangramObj((IWndNodeCollection*)this);
	}
	m_vNodes.clear();
}

STDMETHODIMP CWndNodeCollection::get_NodeCount(long* pCount)
{
	*pCount = (int)m_pNodes->size();
	return S_OK;
}

STDMETHODIMP CWndNodeCollection::get_Item(long iIndex, IWndNode **ppNode)
{
	if (iIndex < 0 || iIndex >= (int)m_pNodes->size()) return E_INVALIDARG;

	CWndNode* pNode = m_pNodes->operator [](iIndex);

	return pNode->QueryInterface(IID_IWndNode, (void**)ppNode);
}

STDMETHODIMP CWndNodeCollection::get__NewEnum(IUnknown** ppVal)
{
	*ppVal = nullptr;

	struct _CopyVariantFromIUnkown
	{
		static HRESULT copy(VARIANT* p1, CWndNode* const* p2)
		{
			CWndNode* pNode = *p2;
			p1->vt = VT_UNKNOWN;
			return pNode->QueryInterface(IID_IUnknown, (void**)&(p1->punkVal));
		}

		static void init(VARIANT* p)	{ VariantInit(p); }
		static void destroy(VARIANT* p)	{ VariantClear(p); }
	};

	typedef CComEnumOnSTL<IEnumVARIANT, &IID_IEnumVARIANT, VARIANT, _CopyVariantFromIUnkown, CTangramNodeVector>
		CComEnumVariantOnVector;

	CComObject<CComEnumVariantOnVector> *pe = 0;
	HRESULT hr = CComObject<CComEnumVariantOnVector>::CreateInstance(&pe);

	if (SUCCEEDED(hr))
	{
		hr = pe->AddRef();
		hr = pe->Init(GetUnknown(), *m_pNodes);

		if (SUCCEEDED(hr))
			hr = pe->QueryInterface(ppVal);

		hr = pe->Release();
	}

	return hr;
}

STDMETHODIMP CWndNode::get_DocXml(BSTR* pVal)
{
	g_pTangram->UpdateWndNode(m_pRootObj);
	CString strXml = m_pTangramNodeCommonData->m_pTangramParse->xml();
	*pVal = strXml.AllocSysString();
	strXml.ReleaseBuffer();

	return S_OK;
}

STDMETHODIMP CWndNode::get_rgbMiddle(OLE_COLOR* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = OLE_COLOR(pSplitter->rgbMiddle);
	}
	else
	{
		*pVal = OLE_COLOR(RGB(240,240,240));
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_rgbMiddle(OLE_COLOR newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		OleTranslateColor(newVal, NULL, &pSplitter->rgbMiddle);
		BYTE Red = GetRValue(pSplitter->rgbMiddle); 
		BYTE Green = GetGValue(pSplitter->rgbMiddle); 
		BYTE Blue = GetBValue(pSplitter->rgbMiddle); 
		CString strRGB = _T("");
		strRGB.Format(_T("RGB(%d,%d,%d)"), Red, Green, Blue);
		put_Attribute(CComBSTR(L"middlecolor"), strRGB.AllocSysString());
		pSplitter->Invalidate();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_rgbLeftTop(OLE_COLOR* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = OLE_COLOR(pSplitter->rgbLeftTop);
	}
	else
	{
		*pVal = OLE_COLOR(RGB(240, 240, 240));
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_rgbLeftTop(OLE_COLOR newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		OleTranslateColor(newVal, NULL, &pSplitter->rgbLeftTop);
		CString strRGB = _T("");
		strRGB.Format(_T("RGB(%d,%d,%d)"), GetRValue(pSplitter->rgbLeftTop), GetGValue(pSplitter->rgbLeftTop), GetBValue(pSplitter->rgbLeftTop));
		put_Attribute(CComBSTR(L"lefttopcolor"), strRGB.AllocSysString());
		pSplitter->Invalidate();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_rgbRightBottom(OLE_COLOR* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = OLE_COLOR(pSplitter->rgbRightBottom);
	}
	else
		*pVal = OLE_COLOR(RGB(240, 240, 240));
	
	return S_OK;
}

STDMETHODIMP CWndNode::put_rgbRightBottom(OLE_COLOR newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		OleTranslateColor(newVal, NULL, &pSplitter->rgbRightBottom);
		BYTE Red = GetRValue(pSplitter->rgbRightBottom);
		BYTE Green = GetGValue(pSplitter->rgbRightBottom);
		BYTE Blue = GetBValue(pSplitter->rgbRightBottom);
		CString strRGB = _T("");
		strRGB.Format(_T("RGB(%d,%d,%d)"), Red, Green, Blue);
		put_Attribute(CComBSTR(L"rightbottomcolor"), strRGB.AllocSysString());
		pSplitter->Invalidate();
	}
	return S_OK;
}

STDMETHODIMP CWndNode::get_Hmin(int* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = pSplitter->m_Hmin;
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_Hmin(int newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		pSplitter->m_Hmin = min(pSplitter->m_Hmax, newVal);
		CString strVal = _T("");
		strVal.Format(_T("%d"), pSplitter->m_Hmin);
		put_Attribute(CComBSTR(L"hmin"), strVal.AllocSysString());
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_Hmax(int* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = pSplitter->m_Hmax;
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_Hmax(int newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		pSplitter->m_Hmax = max(pSplitter->m_Hmin,newVal);
		CString strVal = _T("");
		strVal.Format(_T("%d"), pSplitter->m_Hmax);
		put_Attribute(CComBSTR(L"hmax"), strVal.AllocSysString());
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_Vmin(int* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = pSplitter->m_Vmin;
	}

	return S_OK;
}

STDMETHODIMP CWndNode::put_Vmin(int newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		pSplitter->m_Vmin = min(pSplitter->m_Vmax, newVal);
		CString strVal = _T("");
		strVal.Format(_T("%d"), pSplitter->m_Vmin);
		put_Attribute(CComBSTR(L"vmin"), strVal.AllocSysString());
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_Vmax(int* pVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		*pVal = pSplitter->m_Vmax;
	}

	return S_OK;
}

STDMETHODIMP CWndNode::put_Vmax(int newVal)
{
	if (m_nViewType == Splitter)
	{
		CSplitterNodeWnd* pSplitter = (CSplitterNodeWnd*)m_pHostWnd;
		pSplitter->m_Vmax = max(pSplitter->m_Vmin, newVal);
		CString strVal = _T("");
		strVal.Format(_T("%d"), pSplitter->m_Vmax);
		put_Attribute(CComBSTR(L"vmax"), strVal.AllocSysString());
	}

	return S_OK;
}


STDMETHODIMP CWndNode::get_HostNode(IWndNode** pVal)
{
	if (m_pTangramNodeCommonData->m_pHostClientView)
		*pVal = m_pTangramNodeCommonData->m_pHostClientView->m_pWndNode;

	return S_OK;
}


STDMETHODIMP CWndNode::put_HostNode(IWndNode* newVal)
{
	return S_OK;
}


STDMETHODIMP CWndNode::get_ActivePage(int* pVal)
{
	return S_OK;
}


STDMETHODIMP CWndNode::put_ActivePage(int newVal)
{
	return S_OK;
}

STDMETHODIMP CWndNode::get_MasterRow(int* pVal)
{
	return S_OK;
}

STDMETHODIMP CWndNode::put_MasterRow(int newVal)
{
	return S_OK;
}

STDMETHODIMP CWndNode::get_MasterCol(int* pVal)
{
	return S_OK;
}

STDMETHODIMP CWndNode::put_MasterCol(int newVal)
{
	return S_OK;
}

HRESULT CWndNode::Fire_ExtendComplete()
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		DISPPARAMS params = { NULL, NULL, 0, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}

	for (auto it : m_mapWndNodeProxy)
	{
		it.second->OnExtendComplete();
	}

	return hr;
}

HRESULT CWndNode::Fire_Destroy()
{
	if(m_pRootObj==this)
	{
		if (m_pTangramNodeCommonData->m_pFrame->m_pPage)
		{
			CString strKey = m_strKey + _T("@") + m_pTangramNodeCommonData->m_pFrame->m_strFrameName + _T("@") + m_pTangramNodeCommonData->m_pFrame->m_pPage->m_strConfigFileNodeName;
			auto it = m_pTangramNodeCommonData->m_pFrame->m_mapNeedSaveToConfigNode.find(m_strKey);
			if (it != m_pTangramNodeCommonData->m_pFrame->m_mapNeedSaveToConfigNode.end())
			{
				if (it->second == this)
				{
					CString strFile = m_pTangramNodeCommonData->m_pFrame->m_pPage->m_strPageFilePath;
					CString strXml = _T("");
					CTangramXmlParse m_Parse;
					CTangramXmlParse xml;
					CTangramXmlParse* m_pTangramPageParse = nullptr;
					if (::PathFileExists(strFile))
					{
						if (m_Parse.LoadFile(strFile))
						{
							m_pTangramPageParse = m_Parse.GetChild(_T("tangrampage"));
							if (m_pTangramPageParse == nullptr)
							{
								m_Parse.AddNode(_T("tangrampage"));
								m_pTangramPageParse = m_Parse.GetChild(_T("tangrampage"));
							}
							if (m_pTangramPageParse)
							{
								CTangramXmlParse* pTangramPageParse = m_pTangramPageParse->GetChild(m_pTangramNodeCommonData->m_pPage->m_strConfigFileNodeName);
								if (pTangramPageParse == nullptr)
								{
									pTangramPageParse = m_pTangramPageParse->AddNode(m_pTangramNodeCommonData->m_pPage->m_strConfigFileNodeName);
								}
								if(pTangramPageParse)
								{
									CString strFrameName = m_pTangramNodeCommonData->m_pFrame->m_strFrameName;

									CTangramXmlParse* pTangramFrameParse = pTangramPageParse->GetChild(strFrameName);
									if (pTangramFrameParse == nullptr)
										pTangramFrameParse = pTangramPageParse->AddNode(strFrameName);
									if (pTangramFrameParse)
									{
										CComQIPtr<IWndContainer> pContainer(m_pDisp);
										if (pContainer)
										{
											if (m_nActivePage >0)
											{
												CString strVal = _T("");
												strVal.Format(_T("%d"), m_nActivePage);
												m_pHostParse->put_attr(_T("activepage"), strVal);
											}
											pContainer->Save();
										}
										if (m_nViewType == Splitter)
										{
											((CSplitterNodeWnd*)m_pHostWnd)->Save();
										}

										for (auto it2 : m_vChildNodes)
										{
											g_pTangram->UpdateWndNode(it2);
										}
										CTangramXmlParse* pParse = pTangramFrameParse->GetChild(m_strKey);
										if (pParse)
											pTangramFrameParse->RemoveNode(m_strKey);
									
										strXml = m_pTangramNodeCommonData->m_pTangramParse->xml();
										CString _strName = m_pTangramNodeCommonData->m_pTangramParse->name();
										if (_strName != m_strKey)
										{
											CString strName = _T("<") + _strName;
											int nPos = strXml.ReverseFind('<');
											CString str = strXml.Left(nPos);
											nPos = str.Find(strName);
											str = str.Mid(nPos + strName.GetLength());
											strXml = _T("<");
											strXml += m_strKey;
											strXml += str;
											strXml += _T("</");
											strXml += m_strKey;
											strXml += _T(">");
										}
										xml.LoadXml(strXml);
										if (pTangramFrameParse->AddNode(&xml, _T("")))
											m_Parse.SaveFile(strFile);
									}
								}
							}
						}
					}
					if (strXml!=_T(""))
					{
						CTangramXmlParse m_Parse;
						if (m_Parse.LoadFile(strFile))
						{
							m_pTangramPageParse = m_Parse.GetChild(_T("tangrampage"));
							if (m_pTangramPageParse == nullptr)
							{
								m_Parse.AddNode(_T("tangrampage"));
								m_pTangramPageParse = m_Parse.GetChild(_T("tangrampage"));
							}
							if (m_pTangramPageParse)
							{
								CTangramXmlParse* pPageParse = m_pTangramPageParse->GetChild(m_pTangramNodeCommonData->m_pPage->m_strConfigFileNodeName);
								if (pPageParse)
								{
									CString strFrameName = m_pTangramNodeCommonData->m_pFrame->m_strFrameName;

									CTangramXmlParse* pFrameParse = pPageParse->GetChild(strFrameName);
									if (pFrameParse == nullptr)
										pFrameParse = pPageParse->AddNode(strFrameName);
									if (pFrameParse)
									{
										CTangramXmlParse* pParse = pFrameParse->GetChild(m_strKey);
										if (pParse)
											pFrameParse->RemoveNode(m_strKey);
										if (pFrameParse->AddNode(&xml, _T("")))
											m_Parse.SaveFile(strFile);
									}
								}
							}
						}
					}
				}
			}
		}
	}
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		DISPPARAMS params = { NULL, NULL, 0, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(2, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	for (auto it : m_mapWndNodeProxy)
	{
		it.second->OnDestroy();
	}

	if (g_pTangram->m_pCLRProxy)
	{
		g_pTangram->m_pCLRProxy->ReleaseTangramObj((IWndNode*)this);
	}

	return hr;
}

HRESULT CWndNode::Fire_NodeAddInCreated(IDispatch * pAddIndisp, BSTR bstrAddInID, BSTR bstrAddInXml)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[3];
		avarParams[2] = pAddIndisp;
		avarParams[2].vt = VT_DISPATCH;
		avarParams[1] = bstrAddInID;
		avarParams[1].vt = VT_BSTR;
		avarParams[0] = bstrAddInXml;
		avarParams[0].vt = VT_BSTR;
		DISPPARAMS params = { avarParams, NULL, 3, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(3, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}

	for (auto it : m_mapWndNodeProxy)
	{
		it.second->OnNodeAddInCreated(pAddIndisp, OLE2T(bstrAddInID), OLE2T(bstrAddInXml));
	}

	return hr;
}

HRESULT CWndNode::Fire_NodeAddInsCreated()
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		DISPPARAMS params = { NULL, NULL, 0, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	for (auto it : m_mapWndNodeProxy)
	{
		it.second->OnNodeAddInsCreated();
	}
	return hr;
}

HRESULT CWndNode::Fire_NodeDocumentComplete(IDispatch * ExtenderDisp, BSTR bstrURL)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[2];
		avarParams[1] = ExtenderDisp;
		avarParams[1].vt = VT_DISPATCH;
		avarParams[0] = bstrURL;
		avarParams[0].vt = VT_BSTR;
		DISPPARAMS params = { avarParams, NULL, 2, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(5, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}

	for (auto it : m_mapWndNodeProxy)
	{
		it.second->OnNodeDocumentComplete(ExtenderDisp, OLE2T(bstrURL));
	}

	return hr;
}

HRESULT CWndNode::Fire_ControlNotify(IWndNode * sender, LONG NotifyCode, LONG CtrlID, LONGLONG CtrlHandle, BSTR CtrlClassName)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[5];
		avarParams[4] = sender;
		avarParams[4].vt = VT_DISPATCH;
		avarParams[3] = NotifyCode;
		avarParams[3].vt = VT_I4;
		avarParams[2] = CtrlID;
		avarParams[2].vt = VT_I4;
		avarParams[1] = CtrlHandle;
		avarParams[1].vt = VT_I8;
		avarParams[0] = CtrlClassName;
		avarParams[0].vt = VT_BSTR;
		DISPPARAMS params = { avarParams, NULL, 5, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();

			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);

			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(6, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}

	if (g_pTangram->m_bEclipse&&m_pTangramNodeCommonData->m_pPage&&m_pTangramNodeCommonData->m_pPage->m_mapNotifyCtrl.size())
	{
		map<HWND, CEclipseCtrl*> m_mapTemp;
		for (auto it : m_pTangramNodeCommonData->m_pPage->m_mapNotifyCtrl)
		{
			CEclipseCtrl* pCtrl = it.second;
			if (::IsWindow(it.first))
				pCtrl->Fire_ControlNotify(sender, NotifyCode, CtrlID, CtrlHandle, CtrlClassName);
			else
				m_mapTemp[it.first] = pCtrl;
		}
		if (m_mapTemp.size())
		{
			for (auto it : m_mapTemp)
				m_pTangramNodeCommonData->m_pPage->m_mapNotifyCtrl.erase(m_pTangramNodeCommonData->m_pPage->m_mapNotifyCtrl.find(it.first));
			m_mapTemp.clear();
		}
	}
	for (auto it : m_mapWndNodeProxy)
	{
		it.second->OnControlNotify(sender, NotifyCode, CtrlID, (HWND)CtrlHandle, OLE2T(CtrlClassName));
	}
	return hr;
}

HRESULT CWndNode::Fire_TabChange(LONG ActivePage, LONG OldPage)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[2];
		avarParams[1] = ActivePage;
		avarParams[1].vt = VT_I4;
		avarParams[0] = OldPage;
		avarParams[0].vt = VT_I4;
		DISPPARAMS params = { avarParams, NULL, 2, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			IUnknown* punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();
			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection);

			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(7, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	for (auto it : m_mapWndNodeProxy)
	{
		it.second->OnTabChange(ActivePage, OldPage);
	}
	return hr;
}

HRESULT CWndNode::Fire_IPCMsg(BSTR bstrXml, BSTR bstrData)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[2];
		avarParams[1] = bstrXml;
		avarParams[1].vt = VT_BSTR;
		avarParams[0] = bstrData;
		avarParams[0].vt = VT_BSTR;
		DISPPARAMS params = { avarParams, NULL, 2, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			IUnknown* punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();
			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection);

			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(8, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	//for (auto it : m_mapWndNodeProxy)
	//{
	//	it.second->OnTabChange(ActivePage, OldPage);
	//}
	return hr;
}

STDMETHODIMP CWndNode::put_SaveToConfigFile(VARIANT_BOOL newVal)
{
	if (m_pRootObj==this)
	{
		CString strKey = m_strKey + _T("@") + m_pTangramNodeCommonData->m_pFrame->m_strFrameName + _T("@") + m_pTangramNodeCommonData->m_pFrame->m_pPage->m_strConfigFileNodeName;
		auto it = m_pTangramNodeCommonData->m_pFrame->m_mapNeedSaveToConfigNode.find(m_strKey);
		if (newVal)
		{
			if (it == m_pTangramNodeCommonData->m_pFrame->m_mapNeedSaveToConfigNode.end())
			{
				m_pTangramNodeCommonData->m_pFrame->m_mapNeedSaveToConfigNode[m_strKey] = this;
				if (m_pTangramNodeCommonData->m_pPage->m_strConfigFileNodeName == _T(""))
				{
					m_pTangramNodeCommonData->m_pPage->put_ConfigName(CComBSTR(L""));
				}
				auto it2 = m_pTangramNodeCommonData->m_pPage->m_mapNeedSaveFrame.find(m_pTangramNodeCommonData->m_pFrame->m_hWnd);
				if (it2 == m_pTangramNodeCommonData->m_pPage->m_mapNeedSaveFrame.end())
					m_pTangramNodeCommonData->m_pPage->m_mapNeedSaveFrame[m_pTangramNodeCommonData->m_pFrame->m_hWnd] = m_pTangramNodeCommonData->m_pFrame;
			}
		}
		else if (it != m_pTangramNodeCommonData->m_pFrame->m_mapNeedSaveToConfigNode.end())
		{
			m_pTangramNodeCommonData->m_pFrame->m_mapNeedSaveToConfigNode.erase(it);
			if (m_pTangramNodeCommonData->m_pFrame->m_mapNeedSaveToConfigNode.size() == 0)
			{
				auto it2 = m_pTangramNodeCommonData->m_pPage->m_mapNeedSaveFrame.find(m_pTangramNodeCommonData->m_pFrame->m_hWnd);
				if (it2 != m_pTangramNodeCommonData->m_pPage->m_mapNeedSaveFrame.end())
					m_pTangramNodeCommonData->m_pPage->m_mapNeedSaveFrame.erase(it2);
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndNode::get_DockObj(BSTR bstrName, LONGLONG* pVal)
{
	CString strName = OLE2T(bstrName);
	if (m_nViewType == CLRCtrl)
	{
		CNodeWnd* pWnd = (CNodeWnd*)m_pHostWnd;
		auto it = pWnd->m_mapDockCtrl.find(strName);
		if (it != pWnd->m_mapDockCtrl.end())
		{
			*pVal = (LONGLONG)it->second;
		}
	}
	return S_OK;
}

STDMETHODIMP CWndNode::put_DockObj(BSTR bstrName, LONGLONG newVal)
{
	CString strName = OLE2T(bstrName);
	if (/*m_nViewType == CLRCtrl&&*/::IsWindow((HWND)newVal)&&strName!=_T(""))
	{
		CNodeWnd* pWnd = (CNodeWnd*)m_pHostWnd;
		auto it = pWnd->m_mapDockCtrl.find(strName);
		if (it == pWnd->m_mapDockCtrl.end())
		{
			pWnd->m_mapDockCtrl[strName] = (HWND)newVal;
		}
	}
	return S_OK;
}

STDMETHODIMP CWndNode::SendUCMAMsg(BSTR bstrSipFrom, BSTR bstrTo, BSTR bstrMsg)
{
	if (m_pTangramNodeCommonData->m_pOfficeObj)
	{
		OfficePlus::COfficeAddin* pAddin = (OfficePlus::COfficeAddin*)g_pTangram;
		pAddin->SendUCMAMessage(bstrSipFrom, bstrTo, m_pTangramNodeCommonData->m_pOfficeObj, this, bstrMsg);
	}
	else
		g_pTangram->SendUCMAMessage(bstrSipFrom, bstrTo, this, this, bstrMsg);

	return S_OK;
}

STDMETHODIMP CWndNode::get_CurMsgObj(IMessageObj** pVal)
{
	if (m_pCurMsgObj)
		* pVal = m_pCurMsgObj;

	return S_OK;
}

STDMETHODIMP CWndNode::NavigateURL(BSTR bstrURL, IDispatch* dispObjforScript)
{
	if (m_nViewType == BlankView)
	{
		if (g_pTangram->m_bInitChrome == FALSE)
		{
			g_pTangram->InitChromeComponent();
		}
		return g_pTangram->m_pCLRProxy->NavigateURL(this, OLE2T(bstrURL), dispObjforScript);
	}
	return S_OK;
}


STDMETHODIMP CWndNode::get_URL(BSTR* pVal)
{
	if (m_pHostParse != nullptr)
	{
		CString strVal = m_pHostParse->attr(_T("url"), _T(""));
		*pVal = strVal.AllocSysString();
	}
	return S_OK;
}


STDMETHODIMP CWndNode::put_URL(BSTR newVal)
{
	if (m_pHostParse != nullptr&&m_nViewType==BlankView)
	{
		put_Attribute(CComBSTR(L"url"), OLE2T(newVal));
	}
	return S_OK;
}
