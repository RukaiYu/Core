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

// TangramView.cpp : implementation file
//

#include "stdafx.h"
#include "TangramApp.h"
#include "NodeWnd.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "SplitterWnd.h"
#include "TangramHtmlTreeWnd.h"
#include "OfficePlus\ExcelPlus\ExcelAddin.h"
#include "OfficePlus\ExcelPlus\ExcelPlusWnd.h"

#include "chromium/BrowserWnd.h"
#include "chromium/HtmlWnd.h"

// CNodeWnd

IMPLEMENT_DYNCREATE(CNodeWnd, CWnd)

CNodeWnd::CNodeWnd()
{
	m_bNoMove = false;
	m_bBKWnd = false;
	m_bCreateExternal = false;
	m_bEraseBkgnd = true;
	m_pWndNode = nullptr;
	m_pXHtmlTree = nullptr;
	m_pParentNode = nullptr;
	m_pOleInPlaceActiveObject = nullptr;
}

CNodeWnd::~CNodeWnd()
{
}

BEGIN_MESSAGE_MAP(CNodeWnd, CWnd)
	ON_WM_SIZE()
	ON_WM_DESTROY()
	ON_WM_ERASEBKGND()
	ON_WM_MOUSEACTIVATE()
	ON_WM_WINDOWPOSCHANGED()
	ON_WM_WINDOWPOSCHANGING()
	ON_MESSAGE(WM_TANGRAMMSG, OnTangramMsg)
	ON_MESSAGE(WM_CHROMEWNDNODEMSG, OnWndNodeIPCMsg)
	ON_MESSAGE(WM_TANGRAMGETNODE, OnGetTangramObj)
	ON_MESSAGE(WM_TGM_SETACTIVEPAGE, OnActiveTangramObj)
	ON_MESSAGE(WM_SPLITTERREPOSITION, OnSplitterReposition)
END_MESSAGE_MAP()

// CNodeWnd diagnostics

#ifdef _DEBUG
void CNodeWnd::AssertValid() const
{
	///CView::AssertValid();
}

#ifndef _WIN32_WCE
void CNodeWnd::Dump(CDumpContext& dc) const
{
	CWnd::Dump(dc);
}
#endif
#endif //_DEBUG

//CNodeWnd message handlers
BOOL CNodeWnd::Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext)
{
	m_pWndNode = g_pTangram->m_pActiveNode;
	m_pWndNode->m_nID = nID;
	m_pWndNode->m_pHostWnd = this;

	if (m_pWndNode->m_strID.CompareNoCase(_T("HostView")) == 0)
	{
		CWndFrame* pFrame = m_pWndNode->m_pTangramNodeCommonData->m_pFrame;
		pFrame->m_pBindingNode = m_pWndNode;

		m_pWndNode->m_pTangramNodeCommonData->m_pHostClientView = this;
		CWndPage* pPage = pFrame->m_pPage;
		HWND hWnd = CreateWindow(L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, pParentWnd->m_hWnd, (HMENU)nID, AfxGetInstanceHandle(), NULL);
		BOOL bRet = SubclassWindow(hWnd);
		if (m_pWndNode->m_pParentObj)
		{
			if (m_pWndNode->m_pParentObj->m_nViewType == Splitter)
			{
				::PostMessage(::GetParent(m_hWnd), WM_HOSTNODEFORSPLITTERCREATED, m_pWndNode->m_nRow, m_pWndNode->m_nCol);
				ModifyStyleEx(WS_EX_WINDOWEDGE | WS_EX_CLIENTEDGE, 0);
			}
			else if (g_pTangram->m_pMDIMainWnd&&pFrame->m_nWndFrameType == WndFrameType::MDIFrame&&pFrame->m_pBindingNode)
			{
				RECT rc = { 0,0,0,0 };
				if (::SendMessage(g_pTangram->m_pMDIMainWnd->m_hWnd, WM_QUERYAPPPROXY, (WPARAM)&rc, 19921989) == 19921989)
				{
					::SetWindowPos(pFrame->m_pWorkNode->m_pHostWnd->m_hWnd, m_hWnd, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,  /*SWP_NOSENDCHANGING| SWP_NOZORDER |*/ SWP_NOACTIVATE | SWP_FRAMECHANGED);
				}
			}
		}

		::GetClassName(pFrame->m_hWnd, g_pTangram->m_szBuffer, MAX_PATH);
		CString strName = g_pTangram->m_szBuffer;
		pFrame->m_bChromeFrame = strName == _T("Chrome Extended Window Class");
		return bRet;
	}
	return m_pWndNode->Create(dwStyle, rect, pParentWnd, nID, pContext);
}

LRESULT CNodeWnd::OnSplitterReposition(WPARAM wParam, LPARAM lParam)
{
	switch (m_pWndNode->m_nViewType)
	{
	case CLRCtrl:
	case ActiveX:
	case TabbedWnd:
		m_pWndNode->m_pTangramNodeCommonData->m_pFrame->HostPosChanged();
		break;
		break;
	default:
		break;
	}

	return CWnd::DefWindowProc(WM_SPLITTERREPOSITION, wParam, lParam);
}

LRESULT CNodeWnd::OnWndNodeIPCMsg(WPARAM wParam, LPARAM lParam)
{
	if (lParam)
	{
		ChromePlus::IPCMsg* pIPCInfo = (ChromePlus::IPCMsg*)lParam;
		m_pWndNode->Fire_IPCMsg(CComBSTR(pIPCInfo->m_strKey), CComBSTR(pIPCInfo->m_strData));
		//switch (m_pWndNode->m_nViewType)
		//{
		//case CLRCtrl:
		//case ActiveX:
		//case TabbedWnd:
		//	m_pWndNode->m_pTangramNodeCommonData->m_pFrame->HostPosChanged();
		//	break;
		//break;
		//default:
		//	break;
		//}
	}

	return CWnd::DefWindowProc(WM_CHROMEWNDNODEMSG, wParam, lParam);
}

int CNodeWnd::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message)
{
	CWndFrame* pFrame = m_pWndNode->m_pRootObj->m_pTangramNodeCommonData->m_pFrame;
	if (pFrame->m_pPage->m_pTangramAppProxy)
	{
		HWND hMenuWnd = pFrame->m_pPage->m_pTangramAppProxy->GetActivePopupMenu(nullptr);
		if (::IsWindow(hMenuWnd))
			::PostMessage(hMenuWnd, WM_CLOSE, 0, 0);
	}
	else if (g_pTangram->m_pActiveAppProxy)
	{
		HWND hMenuWnd = g_pTangram->m_pActiveAppProxy->GetActivePopupMenu(nullptr);
		if (::IsWindow(hMenuWnd))
			::PostMessage(hMenuWnd, WM_CLOSE, 0, 0);
	}
	BOOL b = pFrame->m_bDesignerState;
	if (g_pTangram->m_pDesignerWndPage)
		if (pFrame->m_pPage == g_pTangram->m_pDesignerWndPage || m_pWndNode->m_strCnnID != _T(""))
			b = false;
	if (m_pWndNode->m_nViewType == BlankView && m_pWndNode->m_strCnnID == _T(""))
		b = true;
	if (g_pTangram->m_pDocDOMTree&&g_pTangram->m_pCLRProxy)
	{
		if (b)
		{
			if (g_pTangram->m_hVSToolBoxWnd)
			{
				if (::IsChild(g_pTangram->m_hVSToolBoxWnd, m_hWnd) == false)
					g_pTangram->m_pCLRProxy->SelectNode(m_pWndNode);
			}
			else if (::IsChild(g_pTangram->m_hHostWnd, m_hWnd) == false)
				g_pTangram->m_pCLRProxy->SelectNode(m_pWndNode);
		}
	}
	//if(m_pWndNode->m_nViewType == ActiveX&&b)
	//{
	//	g_pTangram->m_pCLRProxy->SelectNode(m_pWndNode);
	//}
	if (m_pWndNode&&m_pWndNode->m_pTangramNodeCommonData->m_pPage)
		m_pWndNode->m_pTangramNodeCommonData->m_pPage->Fire_NodeMouseActivate(m_pWndNode);

	if ((m_pWndNode->m_nViewType == TabbedWnd || m_pWndNode->m_nViewType == Splitter))
	{
		if (g_pTangram->m_pWndFrame != m_pWndNode->m_pTangramNodeCommonData->m_pFrame)
			::SetFocus(m_hWnd);
		g_pTangram->m_pActiveNode = m_pWndNode;
		g_pTangram->m_bWinFormActived = false;
		return MA_ACTIVATE;
	}

	if (!m_bCreateExternal)
		Invalidate(true);
	else
	{
		if ((m_pWndNode->m_nViewType != TabbedWnd && m_pWndNode->m_nViewType != Splitter))
		{
			if (g_pTangram->m_pWndFrame != m_pWndNode->m_pTangramNodeCommonData->m_pFrame || g_pTangram->m_pActiveNode != m_pWndNode)
				::SetFocus(m_hWnd);
		}
	}
	g_pTangram->m_pActiveNode = m_pWndNode;
	g_pTangram->m_bWinFormActived = false;
	if (m_pWndNode->m_pParentObj)
	{
		if (m_pWndNode->m_pParentObj->m_nViewType & TabbedWnd)
			m_pWndNode->m_pParentObj->m_pVisibleXMLObj = m_pWndNode;
	}
	g_pTangram->m_pWndFrame = m_pWndNode->m_pTangramNodeCommonData->m_pFrame;

	if ((m_pWndNode->m_nViewType == ActiveX || m_pWndNode->m_nViewType == CLRCtrl))
	{
		return MA_ACTIVATE;
	}

	if (b&&m_bCreateExternal == false)
	{
		g_pTangram->CreateCommonDesignerToolBar();
		if (g_pTangram->m_pHostViewDesignerNode)
		{
			HWND hwnd = pFrame->m_hWnd;
			::GetClassName(::GetParent(hwnd), g_pTangram->m_szBuffer, MAX_PATH);
			CString strName = CString(g_pTangram->m_szBuffer);
			if (strName.Find(_T("Afx:ControlBar:")) != -1)
			{
				pFrame->m_nWndFrameType = ControlBarFrame;
			}
			::GetClassName(hwnd, g_pTangram->m_szBuffer, MAX_PATH);
			strName = CString(g_pTangram->m_szBuffer);
			CTangramMDIChildWnd* pWnd = (CTangramMDIChildWnd*)::SendMessage(::GetParent(pFrame->m_hWnd), WM_TANGRAMMSG, 0, 19631222);
			if (pWnd)
			{
				g_pTangram->m_pHostViewDesignerNode->put_Caption(CComBSTR(L"Show Host MFC View for MFC Document"));
			}
			else if (strName == _T("MDIClient"))
			{
				g_pTangram->m_pHostViewDesignerNode->put_Caption(CComBSTR(L"Show MDI Client"));
			}
			else if (pFrame->m_pWndFrameInfo)
			{
				CString s = _T("");
				s.Format(_T("Show .NET Ctrl \"%s\" at Control '%s'"), pFrame->m_pWndFrameInfo->m_strCtrlName, pFrame->m_pWndFrameInfo->m_strParentCtrlName);
				g_pTangram->m_pHostViewDesignerNode->put_Caption(CComBSTR(s));
			}
			else if (pFrame->m_nWndFrameType == WndFrameType::ControlBarFrame)
			{
				g_pTangram->m_pHostViewDesignerNode->put_Caption(CComBSTR(L"Show Control Bar Host Object"));
			}
			else
			{
				g_pTangram->m_pHostViewDesignerNode->put_Caption(CComBSTR(L"Create Host Node"));
			}
		}
		if (m_pWndNode->m_strID.CompareNoCase(_T("hostview")) && (m_bCreateExternal == false && m_pWndNode->m_pDisp == NULL))
		{
			if (::IsWindow(g_pTangram->m_hHostWnd) == false)
			{
				g_pTangram->m_hHostWnd = ::CreateWindowEx(NULL, L"Tangram Window Class", _T("Tangram Designer Helper Window"), WS_OVERLAPPED | WS_CAPTION, 0, 0, 0, 0, NULL, NULL, theApp.m_hInstance, NULL);
				g_pTangram->m_mapValInfo[_T("hostwindow")] = CComVariant((LONGLONG)g_pTangram->m_hHostWnd);
			}
			if (g_pTangram->m_pDesignWindowNode&&g_pTangram->m_pDesignWindowNode != m_pWndNode)
			{
				CNodeWnd* pWnd = ((CNodeWnd*)g_pTangram->m_pDesignWindowNode->m_pHostWnd);
				if (pWnd&&::IsWindow(pWnd->m_hWnd))
				{
					pWnd->Invalidate(true);
				}
			}
			g_pTangram->m_pDesignWindowNode = m_pWndNode;
			Invalidate(true);
		}

		if (g_pTangram->m_pDesigningFrame != pFrame)
		{
			g_pTangram->m_pHostDesignUINode = m_pWndNode->m_pRootObj;
			g_pTangram->m_pDesigningFrame = pFrame;
			pFrame->UpdateDesignerTreeInfo();
		}
	}

	if (m_bCreateExternal == false)
		return MA_ACTIVATEANDEAT;
	else
		return CWnd::OnMouseActivate(pDesktopWnd, nHitTest, message);
}

BOOL CNodeWnd::OnEraseBkgnd(CDC* pDC)
{
	CWndFrame* pFrame = m_pWndNode->m_pTangramNodeCommonData->m_pFrame;
	BOOL bInDesignState = pFrame->m_bDesignerState;
	CBitmap bit;
	RECT rt;
	GetClientRect(&rt);
	if (m_pWndNode->m_strID.CompareNoCase(_T("HostView")) == 0)
	{
		HWND hWnd = pFrame->m_hWnd;
		if (pFrame->m_bDesignerState)
		{
		}
		if (::IsWindow(hWnd) && !::IsWindowVisible(hWnd))
		{
			pFrame->HostPosChanged();
			return true;
		}
	}
	else if ((m_pWndNode->m_nViewType == ActiveX || m_pWndNode->m_nViewType == CLRCtrl) && m_pWndNode->m_pDisp == nullptr)
	{
		bit.LoadBitmap(IDB_BITMAP_DESIGNER);
		CBrush br(&bit);
		pDC->FillRect(&rt, &br);
		pDC->SetBkMode(TRANSPARENT);
		CComBSTR bstrCaption(L"");
		m_pWndNode->get_Attribute(CComBSTR(L"caption"), &bstrCaption);
		CString strInfo = _T("");
		strInfo = strInfo +
			_T("\n  ----.NET Assembly don't exists or AxtiveX Control not Support!----");

		pDC->SetTextColor(RGB(255, 0, 255));
		pDC->DrawText(strInfo, &rt, DT_LEFT);
	}

	if (m_pWndNode->m_strID.CompareNoCase(_T("hostview")) && (m_bCreateExternal == false && m_pWndNode->m_pDisp == NULL) && m_bEraseBkgnd)
	{
		CString strText = _T("");
		if (bInDesignState&&g_pTangram->m_pDesignWindowNode == m_pWndNode)
		{
			bit.LoadBitmap(IDB_BITMAP_DESIGNER);
			CBrush br(&bit);
			pDC->FillRect(&rt, &br);
			pDC->SetTextColor(RGB(255, 0, 255));
			CString str = g_pTangram->GetDesignerInfo(_T("SelectedText"));
			if (str == _T(""))
				strText = _T("\n\n  ") + g_pTangram->m_strNodeSelectedText;
			else
				strText = _T("\n\n  ") + str;
		}
		else
		{
			bit.LoadBitmap(IDB_BITMAP_GRID);
			CBrush br(&bit);
			pDC->FillRect(&rt, &br);
			CComBSTR bstrCaption(L"");
			m_pWndNode->get_Attribute(CComBSTR(L"caption"), &bstrCaption);
			CString strInfo = _T("\n\n  ");
			if (bInDesignState)
			{
				CString str = g_pTangram->GetDesignerInfo(_T("DesignerTip1"));
				if (str == _T(""))
					strInfo = strInfo + g_pTangram->m_strDesignerTip1;
				else
					strInfo = strInfo + str;
			}
			CString str = g_pTangram->GetDesignerInfo(_T("DesignerTip2"));
			if (str == _T(""))
				strInfo = strInfo + _T("\n  ") + g_pTangram->m_strDesignerTip2;
			else
				strInfo = strInfo + _T("\n  ") + str;
			if (pFrame->m_pParentNode)
			{
				CString strTip = _T("");
				if (pFrame->m_pParentNode->m_nViewType == CLRCtrl)
				{
					strTip.Format(_T("Sub Frame for .NET CLRCtrl: %s"), pFrame->m_pParentNode->m_strName);
					strInfo = strInfo + _T("\n  ") + strTip;
					strTip.Format(_T("CLRCtrl Type: %s"), pFrame->m_pParentNode->m_strCnnID);
					strInfo = strInfo + _T("\n  ") + strTip;
				}
				else if (m_pWndNode->m_pTangramNodeCommonData->m_pFrame->m_pParentNode->m_nViewType == ActiveX)
				{
					strTip.Format(_T("Sub Frame for ActiveX Control: %s"), pFrame->m_pParentNode->m_strName);
					strInfo = strInfo + _T("\n  ") + strTip;
					strTip.Format(_T("ActiveX ID: %s"), pFrame->m_pParentNode->m_strCnnID);
					strInfo = strInfo + _T("\n  ") + strTip;
				}
			}
			strText.Format(strInfo, m_pWndNode->m_strName, CString(OLE2T(bstrCaption)));
			pDC->SetTextColor(RGB(255, 255, 255));
		}

		strText += g_pTangram->m_strDesignerInfo;
		CString strKey = m_pWndNode->m_pRootObj->m_strKey;
		switch (pFrame->m_nWndFrameType)
		{
		case MDIFrame:
		{
			CWndPage* pPage = pFrame->m_pPage;
			if (pPage)
			{
				LRESULT lRes = ::SendMessage(pPage->m_hWnd, WM_QUERYAPPPROXY, 0, 0);
				if (lRes > 0)
				{
					CTangramAppProxy* pProxy = (CTangramAppProxy*)lRes;
					auto it = g_pTangram->m_mapTemplateInfo.find(strKey);
					if (it != g_pTangram->m_mapTemplateInfo.end())
					{
						CString strInfo = _T("");
						strInfo.Format(_T("  MDIFrame for Document Template : %s"), strKey);
						strText += _T("  ") + strInfo;
					}
				}
			}
		}
		break;
		case MDIChildFrame:
		{
			if (pFrame->m_pDoc)
			{
				CString strInfo = _T("");
				strInfo.Format(_T("  MDIChildFrame for Document Template : %s"), pFrame->m_pDoc->m_strDocID);
				strText += _T("  ") + strInfo;

			}
		}
		break;
		case SDIFrame:
		{
			strText += _T("\n  ") + strKey;
		}
		break;
		case ControlBarFrame:
		{
			CWndPage* pPage = pFrame->m_pPage;
			if (pPage)
			{
				LRESULT lRes = ::SendMessage(pPage->m_hWnd, WM_QUERYAPPPROXY, 0, 0);
				if (lRes > 0)
				{
					CTangramAppProxy* pProxy = (CTangramAppProxy*)lRes;
					auto it = g_pTangram->m_mapTemplateInfo.find(strKey);
					if (it != g_pTangram->m_mapTemplateInfo.end())
					{
						CString strInfo = _T("");
						strInfo.Format(_T("  Control Bar for Document Template: %s"), strKey);
						strText += _T("\n  ") + strInfo;
					}
				}
			}
		}
		break;
		case WinFormMDIFrame:
		{
		}
		break;
		case WinFormMDIChildFrame:
		{
		}
		break;
		case WinFormFrame:
		{
		}
		break;
		case EclipseWorkBenchFrame:
		{
			strText += _T("\n  ") + strKey;
		}
		break;
		case EclipseViewFrame:
		{
		}
		break;
		case EclipseSWTFrame:
		{
		}
		break;
		case WinFormControlFrame:
		{
		}
		break;
		default:
			break;
		}
		pDC->SetBkMode(TRANSPARENT);
		pDC->DrawText(strText, &rt, DT_LEFT);
	}
	return true;
}

BOOL CNodeWnd::PreTranslateMessage(MSG* pMsg)
{
	if (m_pXHtmlTree)
	{
		return m_pXHtmlTree->PreTranslateMessage(pMsg);
	}

	if (m_pOleInPlaceActiveObject)
	{
		LRESULT hr = m_pOleInPlaceActiveObject->TranslateAccelerator((LPMSG)pMsg);
		if (hr == S_OK)
			return true;
		else
		{
			if (m_pWndNode->m_nViewType == CLRCtrl)
			{
				VARIANT_BOOL bShiftKey = false;
				if (::GetKeyState(VK_SHIFT) < 0)  // shift pressed
					bShiftKey = true;
				hr = g_pTangram->m_pCLRProxy->ProcessCtrlMsg(::GetWindow(m_hWnd, GW_CHILD), bShiftKey ? true : false);
				if (hr == S_OK)
					return true;
			}
			else
			{
				pMsg->hwnd = 0;
				return true;
			}
		}
	}
	if (IsDialogMessage(pMsg))
		return true;
	return CWnd::PreTranslateMessage(pMsg);
}

void CNodeWnd::OnDestroy()
{
	if (g_pTangram->m_pDesignWindowNode == m_pWndNode)
	{
		if (g_pTangram->m_pCLRProxy)
			g_pTangram->m_pCLRProxy->SelectNode(NULL);
		g_pTangram->m_pDesignWindowNode = NULL;
	}

	m_pWndNode->Fire_Destroy();
	CWnd::OnDestroy();
}

void CNodeWnd::PostNcDestroy()
{
	delete m_pWndNode;
	CWnd::PostNcDestroy();
	delete this;
}

LRESULT CNodeWnd::OnTangramMsg(WPARAM wParam, LPARAM lParam)
{
	if (wParam == 0 && lParam)//Create CLRCtrl Node
	{
		if (lParam == 19631222||lParam == 20180115)
		{
			return CWnd::DefWindowProc(WM_TANGRAMMSG, wParam, lParam);
		}
		if (lParam == 19920612 && g_pTangram->m_pCLRProxy)
		{
			CString strName = g_pTangram->m_pCLRProxy->m_strObjTypeName;
			CWndNode* pNode = m_pWndNode->m_pRootObj;
			CWndNode* pParent = m_pWndNode->m_pParentObj;
			if (pParent)
			{
				strName = pParent->m_strName + _T("_") + strName;
			}
			auto it = pNode->m_pTangramNodeCommonData->m_mapLayoutNodes.find(strName);
			if (it != pNode->m_pTangramNodeCommonData->m_mapLayoutNodes.end())
			{
				BOOL bGetNew = false;
				int nIndex = 0;
				while (bGetNew == false)
				{
					CString strNewName = _T("");
					strNewName.Format(_T("%s%d"), strName, nIndex);
					it = pNode->m_pTangramNodeCommonData->m_mapLayoutNodes.find(strNewName);
					if (it == pNode->m_pTangramNodeCommonData->m_mapLayoutNodes.end())
					{
						strName = strNewName;
						break;
					}
					nIndex++;
				}
			}
			m_pWndNode->put_Attribute(CComBSTR(L"name"), strName.AllocSysString());
			m_pWndNode->m_strName = strName;
			return 0;
		}
		CString strCnnID = (LPCTSTR)lParam;
		if (strCnnID.Find(_T(",")) != -1 && g_pTangram->m_pCLRProxy)
		{
			m_pWndNode->m_pDisp = g_pTangram->m_pCLRProxy->TangramCreateObject(strCnnID.AllocSysString(), (long)::GetParent(m_hWnd), m_pWndNode);

			if (m_pWndNode->m_pDisp)
			{
				m_pWndNode->m_pTangramNodeCommonData->m_mapLayoutNodes[m_pWndNode->m_strName] = m_pWndNode;
				m_pWndNode->m_nViewType = CLRCtrl;
				CAxWindow m_Wnd;
				m_Wnd.Attach(m_hWnd);
				CComPtr<IUnknown> pUnk;
				m_Wnd.AttachControl(m_pWndNode->m_pDisp, &pUnk);
				CComQIPtr<IOleInPlaceActiveObject> pIOleInPlaceActiveObject(m_pWndNode->m_pDisp);
				if (pIOleInPlaceActiveObject)
					m_pOleInPlaceActiveObject = pIOleInPlaceActiveObject.Detach();
				m_Wnd.Detach();
				if (m_mapDockCtrl.size())
				{
					HWND hPage = m_pWndNode->m_pTangramNodeCommonData->m_pPage->m_hWnd;
					::SendMessage(hPage, WM_TANGRAMMSG, (WPARAM)this, 1963);
				}
			}
			else
			{
				m_bCreateExternal = true;
			}
		}
		else
		{
			CString strName = strCnnID;
			int _nPos = strName.Find(_T("."));
			if (_nPos != -1)
				strName = strName.Mid(_nPos + 1);
			_nPos = strName.ReverseFind('.');
			if (_nPos != -1)
				strName = strName.Left(_nPos);
			CWndNode* pNode = m_pWndNode->m_pRootObj;
			CWndNode* pParent = m_pWndNode->m_pParentObj;
			if (pParent)
			{
				strName = pParent->m_strName + _T("_") + strName;
			}
			auto it = pNode->m_pTangramNodeCommonData->m_mapAxNodes.find(strName);
			if (it != pNode->m_pTangramNodeCommonData->m_mapAxNodes.end())
			{
				BOOL bGetNew = false;
				int nIndex = 0;
				while (bGetNew == false)
				{
					CString strNewName = _T("");
					strNewName.Format(_T("%s%d"), strName, nIndex);
					it = pNode->m_pTangramNodeCommonData->m_mapAxNodes.find(strNewName);
					if (it == pNode->m_pTangramNodeCommonData->m_mapAxNodes.end())
					{
						pNode->m_pTangramNodeCommonData->m_mapAxNodes[strNewName] = m_pWndNode;
						strName = strNewName;
						bGetNew = true;
						break;
					}
					nIndex++;
				}
			}
			else
			{
				pNode->m_pTangramNodeCommonData->m_mapAxNodes[strName] = m_pWndNode;
			}
			m_pWndNode->put_Attribute(CComBSTR(L"name"), strName.AllocSysString());
			m_pWndNode->m_strName = strName;
			BOOL bWebCtrl = false;
			CString strURL = _T("");
			strCnnID.MakeLower();
			auto nPos = strCnnID.Find(_T("http:"));
			if (nPos == -1)
				nPos = strCnnID.Find(_T("https:"));
			if (m_pWndNode->m_pTangramNodeCommonData->m_pFrame)
			{
				CComBSTR bstr;
				m_pWndNode->get_Attribute(CComBSTR("InitInfo"), &bstr);
				CString str = _T("");
				str += bstr;
				if (str != _T("") && m_pWndNode->m_pTangramNodeCommonData->m_pPage)
				{
					LRESULT hr = ::SendMessage(m_pWndNode->m_pTangramNodeCommonData->m_pPage->m_hWnd, WM_GETNODEINFO, (WPARAM)OLE2T(bstr), (LPARAM)::GetParent(m_hWnd));
					if (hr)
					{
						CString strInfo = _T("");
						CWindow m_wnd;
						m_wnd.Attach(::GetParent(m_hWnd));
						m_wnd.GetWindowText(strInfo);
						strCnnID += strInfo;
						m_wnd.Detach();
					}
				}
			}

			if (strCnnID.Find(_T("http://")) != -1 || strCnnID.Find(_T("https://")) != -1)
			{
				bWebCtrl = true;
				strURL = strCnnID;
				strCnnID = _T("shell.explorer.2");
			}

			if (strCnnID.Find(_T("res://")) != -1 || ::PathFileExists(strCnnID))
			{
				bWebCtrl = true;
				strURL = strCnnID;
				strCnnID = _T("shell.explorer.2");
				if (strCnnID.Find(_T("res://")) != -1)
				{
					CString strTemp = strURL;
					strTemp.Replace(_T("res://"), _T(""));
					int nPos = strTemp.ReverseFind('/');
					CString strTemp2 = strTemp.Mid(nPos);
					strTemp = strTemp.Left(nPos);
					if (::PathFileExists(strTemp) == false)
					{
						strTemp = g_pTangram->m_strAppPath + strTemp;
						if (::PathFileExists(strTemp) == false)
						{
							strTemp = g_pTangram->m_strProgramFilePath + _T("Tangram\\") + strTemp;
							if (::PathFileExists(strTemp) == false)
							{
								strURL = _T("about:blank");
							}
							else
								strURL = _T("res://") + strTemp + strTemp2;
						}
						else
							strURL = _T("res://") + strTemp + strTemp2;
					}
				}
			}

			if (strCnnID.CompareNoCase(_T("about:blank")) == 0)
			{
				bWebCtrl = true;
				strURL = strCnnID;
				strCnnID = _T("shell.explorer.2");
			}

			if (m_pWndNode->m_pDisp == NULL)
			{
				CComPtr<IDispatch> pDisp;
				HRESULT hr = pDisp.CoCreateInstance(CComBSTR(strCnnID));
				if (hr == S_OK)
				{
					m_pWndNode->m_pDisp = pDisp.Detach();
				}
			}
			if (m_pWndNode->m_pDisp)
			{
				m_pWndNode->m_pTangramNodeCommonData->m_mapLayoutNodes[m_pWndNode->m_strName] = m_pWndNode;
				m_pWndNode->m_nViewType = ActiveX;
				CAxWindow m_Wnd;
				m_Wnd.Attach(m_hWnd);
				CComPtr<IUnknown> pUnk;
				m_Wnd.AttachControl(m_pWndNode->m_pDisp, &pUnk);
				CComQIPtr<IWebBrowser2> pWebDisp(m_pWndNode->m_pDisp);
				if (pWebDisp)
				{
					bWebCtrl = true;
					m_pWndNode->m_strURL = strURL;
					if (m_pWndNode->m_strURL == _T(""))
						m_pWndNode->m_strURL = strCnnID;
					if (m_pWndNode->m_pTangramNodeCommonData->m_pFrame)
					{
						if (m_pWndNode->m_pTangramNodeCommonData->m_pPage)
							m_pWndNode->m_pTangramNodeCommonData->m_pPage->m_nWebViewCount++;

						m_pWndNode->m_pTangramNodeWBEvent = new CWndNodeWBEvent(m_pWndNode);
					}
					CComPtr<IAxWinAmbientDispatch> spHost;
					LRESULT hr = m_Wnd.QueryHost(&spHost);
					if (SUCCEEDED(hr))
					{
						CComBSTR bstr;
						m_pWndNode->get_Attribute(CComBSTR("scrollbar"), &bstr);
						CString str = OLE2T(bstr);
						if (str == _T("1"))
							spHost->put_DocHostFlags(DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_THEME);//DOCHOSTUIFLAG_DIALOG|
						else
							spHost->put_DocHostFlags(/*DOCHOSTUIFLAG_DIALOG|*/DOCHOSTUIFLAG_SCROLL_NO | DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE | DOCHOSTUIFLAG_THEME);

						if (m_pWndNode->m_strURL != _T(""))
						{
							pWebDisp->Navigate2(&CComVariant(m_pWndNode->m_strURL), &CComVariant(navNoReadFromCache | navNoWriteToCache), NULL, NULL, NULL);
							m_pWndNode->m_bWebInit = true;
						}
					}
				}
				((CNodeWnd*)m_pWndNode->m_pHostWnd)->m_pWndNode = m_pWndNode;
				HWND hPage = m_pWndNode->m_pTangramNodeCommonData->m_pPage->m_hWnd;
				::SendMessage(hPage, WM_TANGRAMMSG, (WPARAM)((CNodeWnd*)m_pWndNode->m_pHostWnd), 1963);
				CComQIPtr<IOleInPlaceActiveObject> pIOleInPlaceActiveObject(m_pWndNode->m_pDisp);
				if (pIOleInPlaceActiveObject)
					m_pOleInPlaceActiveObject = pIOleInPlaceActiveObject.Detach();
				m_Wnd.Detach();
			}
		}
		return 0;
	}

	CString str = (LPCTSTR)lParam;
	IWndNode* _pNode = nullptr;
	CWndNode* pOldNode = (CWndNode*)m_pWndNode;
	if (m_pWndNode->m_hHostWnd == 0)
	{
		RECT rc;
		::GetClientRect(m_hWnd, &rc);
		m_pWndNode->m_hHostWnd = ::CreateWindowEx(NULL, L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, rc.right, rc.bottom, m_hWnd, NULL, AfxGetInstanceHandle(), NULL);
		m_pWndNode->m_hChildHostWnd = ::CreateWindowEx(NULL, L"Tangram Window Class", NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, rc.right, rc.bottom, m_pWndNode->m_hHostWnd, NULL, AfxGetInstanceHandle(), NULL);
		IWndFrame* pFrame = nullptr;
		m_pWndNode->m_pTangramNodeCommonData->m_pPage->CreateFrame(CComVariant(0), CComVariant((long)pOldNode->m_hChildHostWnd), CComBSTR(L"Design"), &pFrame);
		((CWndFrame*)pFrame)->m_pDoc = m_pWndNode->m_pTangramNodeCommonData->m_pFrame->m_pDoc;
		IWndNode* pNode = nullptr;
		pFrame->Extend(CComBSTR(L""), str.AllocSysString(), &pNode);
		m_bEraseBkgnd = false;

		((CWndNode*)pNode)->m_pRootObj = m_pWndNode->m_pRootObj;
		((CWndNode*)pNode)->m_pParentObj = m_pWndNode->m_pParentObj;
		m_pWndNode->m_pTangramNodeCommonData->m_mapLayoutNodes[((CWndNode*)pNode)->m_strName] = (CWndNode*)pNode;
		((CWndFrame*)pFrame)->m_bDesignerState = m_pWndNode->m_pTangramNodeCommonData->m_pFrame->m_bDesignerState;
		CString strXml = ((CWndNode*)pNode)->m_pHostParse->xml();
		CTangramXmlParse* pNew = ((CWndNode*)pNode)->m_pHostParse;
		CTangramXmlParse* pOld = pOldNode->m_pHostParse;
		CTangramXmlParse* pParent = m_pWndNode->m_pHostParse->m_pParentParse;
		CTangramXmlParse* pRet = nullptr;
		if (pParent)
		{
			pRet = pParent->ReplaceNode(pOld, pNew, _T(""));
			CString str = pRet->xml();
			int nCount = pRet->GetCount();
			((CWndNode*)pNode)->m_pHostParse = pRet;
			m_pWndNode->m_pHostParse = pRet;

			CWndNode* pChildNode = nullptr;
			for (auto it2 : ((CWndNode*)pNode)->m_vChildNodes)
			{
				pChildNode = it2;
				pChildNode->m_pRootObj = m_pWndNode->m_pRootObj;
				CString strName = pChildNode->m_strName;
				for (int i = 0; i < nCount; i++)
				{
					CTangramXmlParse* child = pRet->GetChild(i);
					CString _strName = child->attr(_T("name"), _T(""));
					if (_strName.CompareNoCase(strName) == 0)
					{
						pChildNode->m_pHostParse = child;
						break;
					}
				}
			}
			m_pWndNode->m_vChildNodes.push_back(((CWndNode*)pNode));
		}

		strXml = m_pWndNode->m_pTangramNodeCommonData->m_pTangramParse->xml();
		g_pTangram->m_pDesignWindowNode = m_pWndNode;
		g_pTangram->m_pHostDesignUINode = m_pWndNode->m_pRootObj;
		if (g_pTangram->m_pHostDesignUINode)
		{
			CTangramHtmlTreeWnd* pTreeCtrl = g_pTangram->m_pDocDOMTree;
			if (pTreeCtrl)
			{
				pTreeCtrl->DeleteItem(g_pTangram->m_pDocDOMTree->m_hFirstRoot);

				if (pTreeCtrl->m_pHostXmlParse)
				{
					delete pTreeCtrl->m_pHostXmlParse;
				}
				pTreeCtrl->m_pHostXmlParse = new CTangramXmlParse();
				pTreeCtrl->m_pHostXmlParse->LoadXml(strXml);
				pTreeCtrl->m_hFirstRoot = pTreeCtrl->LoadXmlFromXmlParse(pTreeCtrl->m_pHostXmlParse);
				pTreeCtrl->ExpandAll();
			}
		}
		auto it = m_pWndNode->m_pTangramNodeCommonData->m_pFrame->m_pPage->m_mapFrame.find(pOldNode->m_hChildHostWnd);
		if (it != m_pWndNode->m_pTangramNodeCommonData->m_pFrame->m_pPage->m_mapFrame.end())
			m_pWndNode->m_pTangramNodeCommonData->m_pFrame->m_pPage->m_mapFrame.erase(it);
	}
	return -1;
}

LRESULT CNodeWnd::OnActiveTangramObj(WPARAM wParam, LPARAM lParam)
{
	if (m_pWndNode->m_nViewType == CLRCtrl)
		::SetWindowLong(m_hWnd, GWL_STYLE, WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS);

	m_pWndNode->m_pTangramNodeCommonData->m_pFrame->HostPosChanged();
	::InvalidateRect(::GetParent(m_hWnd), nullptr, true);
	return CWnd::DefWindowProc(WM_TGM_SETACTIVEPAGE, wParam, lParam);
}

LRESULT CNodeWnd::OnGetTangramObj(WPARAM wParam, LPARAM lParam)
{
	if (m_pWndNode)
		return (LRESULT)m_pWndNode;
	return (long)CWnd::DefWindowProc(WM_TANGRAMGETNODE, wParam, lParam);;
}

LRESULT CNodeWnd::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	if (m_bCreateExternal)
	{
		switch (message)
		{
		case WM_COMMAND:
		{
			WNDPROC* lplpfn = GetSuperWndProcAddr();
			LRESULT res = CallWindowProc(*lplpfn, m_hWnd, message, wParam, lParam);
			HWND hWnd = (HWND)lParam;
			if (::IsWindow(hWnd) && m_pWndNode)
			{
				::GetClassName(hWnd, g_pTangram->m_szBuffer, MAX_PATH);
				m_pWndNode->Fire_ControlNotify(m_pWndNode, HIWORD(wParam), LOWORD(wParam), lParam, CComBSTR(g_pTangram->m_szBuffer));
			}

			return res;
		}
		case WM_ACTIVATE:
		case WM_ERASEBKGND:
		case WM_SETFOCUS:
		case WM_CONTEXTMENU:
		case WM_LBUTTONDOWN:
		case WM_RBUTTONDOWN:
		case WM_LBUTTONUP:
		case WM_RBUTTONUP:
		case WM_LBUTTONDBLCLK:
		{
			WNDPROC* lplpfn = GetSuperWndProcAddr();
			return CallWindowProc(*lplpfn, m_hWnd, message, wParam, lParam);
		}
		case WM_MOUSEACTIVATE:
		{
			break;
		}
		case WM_SHOWWINDOW:
		{
			if (wParam&&m_pWndNode->m_strURL != _T(""))
			{
				CComQIPtr<IWebBrowser2> pWebCtrl(m_pWndNode->m_pDisp);
				if (pWebCtrl)
				{
					if (m_pWndNode->m_bWebInit == false)
					{
						pWebCtrl->Navigate2(&CComVariant(m_pWndNode->m_strURL), &CComVariant(navNoReadFromCache | navNoWriteToCache), NULL, NULL, NULL);
						m_pWndNode->m_bWebInit = true;
					}
				}
			}
			break;
		}
		}
	}

	return CWnd::WindowProc(message, wParam, lParam);
}

CBKWnd::CBKWnd(void)
{
	m_pFrame = nullptr;
	m_hChild = nullptr;
	m_pWndNode = nullptr;
	m_strURL = _T("");
}

void CBKWnd::OnFinalMessage(HWND hWnd)
{
	CWindowImpl<CBKWnd, CWindow>::OnFinalMessage(hWnd);
	delete this;
}

LRESULT CBKWnd::OnMouseActivate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	if (m_pWndNode)
	{
		g_pTangram->m_pActiveNode = m_pWndNode;
		g_pTangram->m_bWinFormActived = false;
	}
	return DefWindowProc(uMsg, wParam, lParam);
}

LRESULT CBKWnd::OnMdiClientCreated(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	if (m_hChild)
	{
		CWndPage* pPage = m_pWndNode->m_pTangramNodeCommonData->m_pPage;
		IWndFrame* pFrame = nullptr;
		pPage->CreateFrame(CComVariant(0), CComVariant((LONGLONG)m_hChild), CComBSTR(L"ClientFrame"), &pFrame);
		CString strXml = _T("");
		strXml.Format(_T("<mdiclient><window><node name=\"mdiclient\" id=\"activex\" cnnid=\"%s\" /></window></mdiclient>"), m_strURL);
		IWndNode* pNode = nullptr;
		pFrame->Extend(CComBSTR(L"default"), strXml.AllocSysString(), &pNode);
		m_pFrame = (CWndFrame*)pFrame;
		g_pTangram->m_mapBKFrame[m_hChild] = m_pFrame;
	}
	return 1;
}

LRESULT CBKWnd::OnWindowPosChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	WINDOWPOS* lpwndpos = (WINDOWPOS*)lParam;
	if (lpwndpos->cx == 0 && lpwndpos->cy == 0)
	{
		RECT rc;
		::GetClientRect(::GetParent(m_hWnd), &rc);
		lpwndpos->cx = rc.right;
		lpwndpos->cy = rc.bottom;
	}
	lpwndpos->flags |= SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOREDRAW;

	if (::IsWindow(m_hChild))
	{
		::SetWindowPos(m_hChild, HWND_BOTTOM, 0, 0, lpwndpos->cx, lpwndpos->cy, SWP_NOREDRAW | SWP_NOACTIVATE);
	}
	return DefWindowProc(uMsg, wParam, lParam);
}

void CNodeWnd::OnWindowPosChanging(WINDOWPOS* lpwndpos)
{
	if (m_pWndNode->m_strID.CompareNoCase(_T("hostview")) == 0)
	{
		ATLTRACE(_T("CNodeWnd::OnWindowPosChanging: %d,%d, width:%d height:%d\n"),lpwndpos->x, lpwndpos->y, lpwndpos->cx, lpwndpos->cy);
	}
	CWnd::OnWindowPosChanging(lpwndpos);
	if (m_bCreateExternal)
		Invalidate(true);
}

void CNodeWnd::OnWindowPosChanged(WINDOWPOS* lpwndpos)
{
	if (m_pWndNode&&m_pWndNode->m_nViewType == CLRCtrl && m_pWndNode->m_hHostWnd)
	{
		if (m_bNoMove)
		{
			lpwndpos->flags |= SWP_NOMOVE;
			CNodeWnd* pNodeWnd = (CNodeWnd*)m_pWndNode->m_pHostWnd;
			if (m_pWndNode->m_pParentObj->m_nViewType == Splitter)
			{
				CSplitterNodeWnd* pWnd = (CSplitterNodeWnd*)m_pWndNode->m_pParentObj->m_pHostWnd;
				::PostMessage(pWnd->m_hWnd, WM_TANGRAMMSG, 3, 1993);
			}
		}
	}
	CWnd::OnWindowPosChanged(lpwndpos);
	if (m_pWndNode&&m_pWndNode->m_nViewType == CLRCtrl && m_pWndNode->m_hHostWnd)
	{
		if (m_bNoMove)
		{
			m_bNoMove = false;
		}
		return;
	}
	if (m_pWndNode&&m_pWndNode->m_nViewType == BlankView && m_pWndNode->m_hHostWnd)
	{
		::SetWindowPos(m_pWndNode->m_hHostWnd, HWND_BOTTOM, 0, 0, lpwndpos->cx, lpwndpos->cy, SWP_NOACTIVATE | SWP_NOREDRAW);
		if (m_pWndNode->m_hChildHostWnd)
			::SetWindowPos(m_pWndNode->m_hChildHostWnd, HWND_BOTTOM, 0, 0, lpwndpos->cx, lpwndpos->cy, SWP_NOACTIVATE | SWP_NOREDRAW);
	}
	if (m_pWndNode->m_strID.CompareNoCase(_T("hostview")) && (m_bCreateExternal == false && m_pWndNode->m_pDisp == NULL) && m_pWndNode != m_pWndNode->m_pRootObj)
	{
		return;
	}
	if (m_pWndNode->m_nViewType == TangramTreeView)
	{
		lpwndpos->flags &= ~SWP_NOREDRAW;
		::SetWindowPos(m_pXHtmlTree->m_hWnd, NULL, 0, 0, lpwndpos->cx, lpwndpos->cy, SWP_NOACTIVATE);
	}

	if (m_bCreateExternal || m_pWndNode == m_pWndNode->m_pRootObj)
	{
		Invalidate(true);
	}
}

void CNodeWnd::OnSize(UINT nType, int cx, int cy)
{
	CWnd::OnSize(nType, cx, cy);

	if (m_pWndNode->m_nViewType == TangramTreeView && m_pWndNode->m_pParentObj == NULL)
	{
		::SetWindowPos(m_pXHtmlTree->m_hWnd, NULL, 0, 0, cx, cy,/*SWP_NOREDRAW|*/SWP_NOACTIVATE);
	}
}
