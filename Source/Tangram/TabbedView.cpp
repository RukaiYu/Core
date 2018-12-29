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

// TangramTabbedFormView.cpp : implementation file
//

#include "stdafx.h"
#include "TangramApp.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "TabbedView.h"

// CTabbedNodeView

IMPLEMENT_DYNCREATE_ATL(CTabbedNodeView, CFormView)

CTabbedNodeView::CTabbedNodeView()
: CFormView(CTabbedNodeView::IDD)
{
	m_pWndNode = nullptr;
}

CTabbedNodeView::~CTabbedNodeView()
{
}

void CTabbedNodeView::DoDataExchange(CDataExchange* pDX)
{
	CFormView::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_TAB, m_TangramTabCtrl);
}

BEGIN_MESSAGE_MAP(CTabbedNodeView, CFormView)
	ON_NOTIFY(TCN_SELCHANGE, IDC_TAB, &CTabbedNodeView::OnTcnSelchangeTab)
	ON_MESSAGE(WM_TGM_SETACTIVEPAGE,OnActiveTangramObj)
	ON_MESSAGE(WM_MODIFYTABPAGE,OnModifyPage)
	ON_MESSAGE(WM_ACTIVETABPAGE,OnActivePage)
	ON_MESSAGE(WM_TGM_SET_CAPTION,OnTgmSetCaption)
	ON_WM_SIZE()
	ON_WM_MOUSEACTIVATE()
	ON_WM_NCDESTROY()
	ON_WM_DESTROY()
END_MESSAGE_MAP()


// CTabbedNodeView diagnostics

#ifdef _DEBUG
void CTabbedNodeView::AssertValid() const
{
	//CFormView::AssertValid();
}

#ifndef _WIN32_WCE
void CTabbedNodeView::Dump(CDumpContext& dc) const
{
	CFormView::Dump(dc);
}
#endif
#endif //_DEBUG


// CTabbedNodeView message handlers

LRESULT CTabbedNodeView::OnActiveTangramObj(WPARAM wParam, LPARAM lParam)
{
	m_pWndNode->m_pTangramNodeCommonData->m_pFrame->HostPosChanged();
	return -1;
}

void CTabbedNodeView::OnTcnSelchangeTab(NMHDR *pNMHDR, LRESULT *pResult)
{
	int nSel = m_TangramTabCtrl.GetCurSel();
	TC_ITEM tcItem;
	tcItem.mask = TCIF_PARAM;
	m_TangramTabCtrl.GetItem(nSel,&tcItem);

	if(m_pWndNode->m_pVisibleXMLObj)
	{
		m_pWndNode->m_pVisibleXMLObj->m_pHostWnd->ShowWindow(SW_HIDE);
		m_pWndNode->m_pVisibleXMLObj = (CWndNode*)tcItem.lParam;
		m_pVisibleWnd = m_pWndNode->m_pVisibleXMLObj->m_pHostWnd;
		m_pVisibleWnd->ShowWindow(SW_SHOW);
		if(m_pWndNode->m_pVisibleXMLObj->m_nViewType & TangramView)
			m_pVisibleWnd->ModifyStyleEx(0,WS_EX_CLIENTEDGE);
		else
			m_pVisibleWnd->ModifyStyleEx(WS_EX_CLIENTEDGE,0);

		m_pVisibleWnd->BringWindowToTop();

		AdjustPageSize();

		m_pWndNode->m_pTangramNodeCommonData->m_pFrame->HostPosChanged();

		//CWndNode* pWindowNode = theApp.FindActiveNode(m_pWndNode);
		//if (pWindowNode&&pWindowNode->m_nViewType == CLRCtrl&&g_pTangram->m_pCLRProxy)
		//{
		//	LRESULT hr = S_FALSE;
		//	hr = g_pTangram->m_pCLRProxy->CtrlMouseClick(pWindowNode->m_pDisp);
		//}
	}
	*pResult = 0;
}

BOOL CTabbedNodeView::Create(LPCTSTR lpszClassName, LPCTSTR lpszWindowName, DWORD dwStyle, const RECT& rect, CWnd* pParentWnd, UINT nID, CCreateContext* pContext)
{
	m_pWndNode = g_pTangram->m_pActiveNode;
	BOOL bRet =true;
	if(m_pWndNode)
	{
		m_pWndNode -> m_nViewType = TabbedWnd;
		m_pWndNode -> m_pHostWnd = this;
		AFX_MANAGE_STATE(AfxGetStaticModuleState());
		bRet = CFormView::Create(lpszClassName, lpszWindowName, dwStyle, rect, pParentWnd, nID, pContext);
		OnInitialUpdate();
		ASSERT(::IsWindow(m_TangramTabCtrl.m_hWnd));
	}

	AdjustTabCtrlSize();

	CRect rtPage(GetPageSize());

	if(m_pWndNode)
	{
		int nCol = m_pWndNode->m_pHostParse->GetCount();

		m_pWndNode->m_nRows = 1;
		m_pWndNode->m_nCols = nCol;

		if(m_pWndNode->m_nActivePage<0||m_pWndNode->m_nActivePage>nCol-1)
			m_pWndNode->m_nActivePage=0;

		CWnd* pView = NULL;
		CWndNode* pObj = NULL;

		int j=0;
		CTangramXmlParse* pChild = NULL;
		for (int i = 0; i < nCol;i++)
		{
			pChild = m_pWndNode->m_pHostParse->GetChild(i);
			CString strName = pChild->attr(TGM_NAME,_T(""));
			strName.Trim();
			strName.MakeLower();
			pObj = new CComObject<CWndNode>;
			pObj -> m_pRootObj = m_pWndNode -> m_pRootObj;
			pObj->m_pHostParse = pChild;
			m_pWndNode->AddChildNode(pObj);
			pObj->m_pParentObj = m_pWndNode;
			pObj->InitWndNode();
			pObj->m_nCol = j;

			if(pObj->m_pObjClsInfo)
			{
				pContext->m_pNewViewClass = pObj->m_pObjClsInfo;
				///////////////////////////////////////////////////////////////////////////////////////////////
				TCITEM tcItem;
				tcItem.mask = TCIF_TEXT | TCIF_PARAM;
				tcItem.pszText = pObj->m_strCaption.GetBuffer();

				pView = (CWnd*)pObj->m_pObjClsInfo->CreateObject();
				UINT nID2 = WS_CHILD|WS_VISIBLE;
				if(j!=m_pWndNode->m_nActivePage)nID2 = WS_CHILD;
				
				pView->Create(NULL,_T(""),nID2,rtPage,this,0,pContext);
				pView->ShowWindow(SW_HIDE);
				m_pWndNode->m_pVisibleXMLObj = pObj; 

				tcItem.lParam = (LPARAM)pObj;
				m_TangramTabCtrl.InsertItem(j, &tcItem);
				///////////////////////////////////////////////////////////////////////////////////////////////
			}
			j++;
		}

		OnActivePage(m_pWndNode->m_nActivePage,NULL);
	}
	return bRet;
}

void CTabbedNodeView::OnSize(UINT nType, int cx, int cy)
{
	CFormView::OnSize(nType, cx, cy);
	if(::IsWindow(m_TangramTabCtrl.m_hWnd))
	{
		AdjustTabCtrlSize(CRect(0,0,cx,cy));
		AdjustPageSize();
	}
}

LRESULT CTabbedNodeView::OnActivePage(WPARAM wParam, LPARAM lParam)
{
	m_TangramTabCtrl.SetCurSel(wParam);

	LRESULT lr;
	OnTcnSelchangeTab(NULL,&lr);
	int iIndex = (int)wParam;
	IWndNode* pActiveNode = NULL;
	if (m_pWndNode)
	{
		CComPtr<IWndNodeCollection> pCol;
		m_pWndNode->get_ChildNodes(&pCol);
		CComPtr<IWndNode> pNode;
		pCol->get_Item(iIndex, &pNode);
		if (pNode)
		{
			pNode->ActiveTabPage(pNode);
		}
	}
	return 0;
}

LRESULT CTabbedNodeView::OnModifyPage(WPARAM wParam, LPARAM lParam)
{
	if(m_pWndNode->m_pVisibleXMLObj)
	{
		m_pVisibleWnd = m_pWndNode->m_pVisibleXMLObj->m_pHostWnd;
		m_pVisibleWnd->BringWindowToTop();
		AdjustPageSize();
	}
	return 0;
}

LRESULT CTabbedNodeView::OnTgmSetCaption(WPARAM wParam,LPARAM lParam)
{
	int iIndex = (int)wParam;
	LPTSTR lpStr = (LPTSTR)lParam;

	int iPageCount = m_TangramTabCtrl.GetItemCount();

	if (iIndex >= iPageCount) return 0;

	TCITEM tcItem = {0};


	tcItem.mask = TCIF_TEXT;
	tcItem.pszText = lpStr;

	m_TangramTabCtrl.SetItem(iIndex,&tcItem);

	return 0;

}

BOOL CTabbedNodeView::PreCreateWindow(CREATESTRUCT& cs)
{
	AfxRegisterWndClass(0);
	cs.style|=WS_CHILD|WS_CLIPSIBLINGS;

	return CFormView::PreCreateWindow(cs);
}

int CTabbedNodeView::OnMouseActivate(CWnd* pDesktopWnd, UINT nHitTest, UINT message)
{
	if(m_pWndNode->m_pParentObj)
	{
		if (m_pWndNode->m_pParentObj->m_nViewType & TabbedWnd)
			m_pWndNode->m_pParentObj->m_pVisibleXMLObj = m_pWndNode;
	}
	return CFormView::OnMouseActivate(pDesktopWnd, nHitTest, message);
}

RECT CTabbedNodeView::GetPageSize()
{	
	CRect rc;
	m_TangramTabCtrl.GetWindowRect(&rc);
	m_TangramTabCtrl.AdjustRect(false,&rc);
	ScreenToClient(&rc);
	return rc;
}

BOOL CTabbedNodeView::AdjustPageSize()
{
	return AdjustPageSize(GetPageSize());
}

BOOL CTabbedNodeView::AdjustTabCtrlSize()
{
	CRect rc;
	GetClientRect(rc);
	return AdjustTabCtrlSize(rc);
}

BOOL CTabbedNodeView::AdjustTabCtrlSize(RECT& rc)
{
	return m_TangramTabCtrl.SetWindowPos(NULL,2,2,rc.right-rc.left-4,rc.bottom-rc.top-4,
		SWP_NOACTIVATE|SWP_NOZORDER);
}

BOOL CTabbedNodeView::AdjustPageSize(RECT& rc)
{
	CRect rt(rc);

	if(m_pWndNode->m_pVisibleXMLObj)
	{
		::SetWindowPos(m_pVisibleWnd->m_hWnd, NULL,
			rt.left,rt.top+2,rt.Width()-1,rt.Height()-2,SWP_SHOWWINDOW|SWP_FRAMECHANGED|SWP_NOREDRAW);
	}
	return true;
}

STDMETHODIMP CTabbedNodeView::Save()
{
	CString str;
	str.Format(_T("%d"),m_TangramTabCtrl.GetCurSel());
	m_pWndNode->put_Attribute(L"activepage",CComBSTR(str));
	return S_OK;
}


void CTabbedNodeView::OnNcDestroy()
{
	__super::OnNcDestroy();
	delete m_pWndNode;
}


void CTabbedNodeView::OnDestroy()
{
	m_pWndNode->Fire_Destroy();
	__super::OnDestroy();
}
