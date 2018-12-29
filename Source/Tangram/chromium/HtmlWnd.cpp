/********************************************************************************
 *					Tangram Library - version 10.0.0
 **
 *********************************************************************************
 * Copyright (C) 2002-2018 by Tangram Team.   All Rights Reserved.
 **
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

#include "../stdafx.h"
#include "../TangramApp.h"
#include "../WndNode.h"
#include "../WndFrame.h"
#include "../NodeWnd.h"
#include "ChromeBrowserProxy.h"
#include "HtmlWnd.h"
#include "BrowserWnd.h"

namespace ChromePlus {
	CHtmlWnd::CHtmlWnd() {
		m_pWebWnd = nullptr;
		m_pDevToolWnd = nullptr;
		m_bDevToolWnd = false;
		m_strCurKey = _T("");
		m_pFrame = nullptr;
		m_hHostWnd=m_hExtendWnd = m_hChildWnd = NULL;
	}

	CHtmlWnd::~CHtmlWnd() {
	}

	LRESULT CHtmlWnd::OnChromeIPCMessage(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
	{
		ChromePlus::IPCMsg* pIPCInfo = (ChromePlus::IPCMsg*)wParam;

		if (m_pFrame)
		{
			int result = m_pFrame->Dispatch(pIPCInfo->m_strType, pIPCInfo->m_strKey, pIPCInfo->m_strData);
			if (result == 0)
			{
				// TODO: Handled by Third-party developers
			}
		}

		LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
		return lRes;
	}

	LRESULT CHtmlWnd::OnOpenDOMWindow(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
	{
		if (lParam) {
			ChromePlus::CChromeTangram* pChromeTangram = (ChromePlus::CChromeTangram*)g_pTangram;
			m_pChromeRenderFrameHost = (CChromeRenderFrameHostBase*)wParam;
			ChromePlus::IPCMsg* pIPCInfo = (ChromePlus::IPCMsg*)lParam;
			CString strFeatures = pIPCInfo->m_strData;
			int nPos = strFeatures.Find(_T(":"));
			if (nPos != -1) {
				CString strID = strFeatures.Mid(nPos + 1);
				CString strType = strFeatures.Left(nPos);
				if (strType.CompareNoCase(_T("app")) == 0)
				{
					g_pTangram->StartApplication(CComBSTR(strID), CComBSTR(pIPCInfo->m_strType));
					LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
					return lRes;
				}
				if (strID == _T("tangram"))
				{
					CString strKey = pIPCInfo->m_strKey;
					strKey.MakeLower();
					CString strXml = pIPCInfo->m_strType;
					if (strKey != _T("")) {
						// for tangram developer
						if (m_hExtendWnd == nullptr)
						{
							m_hExtendWnd = ::CreateWindowEx(NULL, _T("Chrome Extended Window Class"), L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, ::GetParent(m_hWnd), NULL, theApp.m_hInstance, NULL);
							m_hChildWnd = ::CreateWindowEx(NULL, _T("Chrome Extended Window Class"), L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, m_hExtendWnd, (HMENU)AFX_IDW_PANE_FIRST, theApp.m_hInstance, NULL);

							::SetWindowLongPtr(m_hExtendWnd, GWLP_USERDATA, (LONG_PTR)m_hChildWnd);
							::SetWindowLongPtr(m_hChildWnd, GWLP_USERDATA, (LONG_PTR)this);
						}
						if (m_hExtendWnd)
						{
							if (m_pFrame == nullptr) {
								IWndPage* pPage = nullptr;
								pChromeTangram->CreateWndPage((__int64)m_hExtendWnd, &pPage);
								if (pPage) {
									IWndFrame* pFrame = nullptr;
									pPage->CreateFrame(CComVariant((__int64)0), CComVariant((__int64)m_hChildWnd), CComBSTR("default"), &pFrame);
									if (pFrame)
									{
										m_pFrame = (CWndFrame*)pFrame;
										m_pFrame->m_pWebWnd = this;
									}
								}
							}
							if (m_pFrame)
							{
								strKey.MakeLower();
								IWndNode* pNode = nullptr;
								m_pFrame->Extend(CComBSTR(strKey), CComBSTR(strXml), &pNode);
								if (pNode)
								{
									m_strCurKey = strKey;
									if (m_pFrame->m_pBindingNode)
										m_hHostWnd = m_pFrame->m_pBindingNode->m_pHostWnd->m_hWnd;
									else
										m_hHostWnd = NULL;
									if (::IsWindowVisible(m_hWnd))
									{
										auto it = pChromeTangram->m_mapBrowserWnd.find(::GetParent(m_hWnd));
										if (it != pChromeTangram->m_mapBrowserWnd.end())
										{
											CBrowserWnd* pParent = it->second;
											pParent->ChromeDraw();
											pParent->m_pBrowser->LayoutBrowser();
											::PostMessage(pParent->m_hWnd, WM_CHROMEDRAW, 2, 1);//| SWP_FRAMECHANGED
											if (pParent->m_pBrowser&&m_pFrame->m_pBindingNode == nullptr) {
												pParent->m_pBrowser->LayoutBrowser();
											}
										}
									}
								}
							}
						}
					}
					LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
					return lRes;
				}
			}
			else
			{
				::SendMessage(m_hChildWnd, WM_CHROMEOPENWINDOWMSG, wParam, lParam);
			}
		}

		LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
		return lRes;
	}

	LRESULT CHtmlWnd::OnTangramMsg(UINT uMsg,
		WPARAM wParam,
		LPARAM lParam,
		BOOL&) {
		switch (wParam)
		{
		case 20181220:
			{
				if (m_pChromeRenderFrameHost->m_pMsg == nullptr)
					return 0;
				int nPos = m_pChromeRenderFrameHost->m_pMsg->m_strKey.Find(_T("|"));
				CString strKey = m_pChromeRenderFrameHost->m_pMsg->m_strKey.Left(nPos);
				CString strXml = m_pChromeRenderFrameHost->m_pMsg->m_strData;

				CBrowserWnd* pParent = nullptr;
				ChromePlus::CChromeTangram* pChromeTangram = (ChromePlus::CChromeTangram*)g_pTangram;
				auto it = pChromeTangram->m_mapBrowserWnd.find(::GetParent(m_hWnd));
				if (it != pChromeTangram->m_mapBrowserWnd.end())
				{
					pParent = it->second;
					if (strKey != _T("")) {
						if (m_hExtendWnd == nullptr)
						{
							m_hExtendWnd = ::CreateWindowEx(NULL, _T("Chrome Extended Window Class"), L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, ::GetParent(m_hWnd), NULL, theApp.m_hInstance, NULL);
							m_hChildWnd = ::CreateWindowEx(NULL, _T("Chrome Extended Window Class"), L"", WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, 0, 0, 0, 0, m_hExtendWnd, (HMENU)AFX_IDW_PANE_FIRST, theApp.m_hInstance, NULL);

							::SetWindowLongPtr(m_hExtendWnd, GWLP_USERDATA, (LONG_PTR)m_hChildWnd);
							::SetWindowLongPtr(m_hChildWnd, GWLP_USERDATA, (LONG_PTR)this);
						}
						if (m_hExtendWnd)
						{
							if (m_pFrame == nullptr) {
								IWndPage* pPage = nullptr;
								pChromeTangram->CreateWndPage((__int64)m_hExtendWnd, &pPage);
								if (pPage) {
									IWndFrame* pFrame = nullptr;
									pPage->CreateFrame(CComVariant((__int64)0), CComVariant((__int64)m_hChildWnd), CComBSTR("default"), &pFrame);
									if (pFrame)
									{
										m_pFrame = (CWndFrame*)pFrame;
										m_pFrame->m_pWebWnd = this;
									}
								}
							}
							if (m_pFrame) {
								IWndNode* pNode = nullptr;
								m_pFrame->Extend(CComBSTR(strKey), CComBSTR(strXml), &pNode);
								if (pNode)
								{
									if (m_pFrame->m_pBindingNode)
										m_hHostWnd = m_pFrame->m_pBindingNode->m_pHostWnd->m_hWnd;
									else
										m_hHostWnd = NULL;
									m_strCurKey = g_pTangram->m_strKey;
									g_pTangram->m_strKey = _T("");
									pParent->ChromeDraw();
									pParent->m_pBrowser->LayoutBrowser();
									::PostMessage(pParent->m_hWnd, WM_CHROMEDRAW, 2, 1);//| SWP_FRAMECHANGED
									if (pParent->m_pBrowser&&m_pFrame->m_pBindingNode == nullptr) {
										pParent->m_pBrowser->LayoutBrowser();
									}
								}
							}
						}
					}
				}
				//delete m_pChromeRenderFrameHost->m_pMsg;
				//m_pChromeRenderFrameHost->m_pMsg = nullptr;
				IPCMsg* pMsg = m_pChromeRenderFrameHost->m_pMsg;
				m_pChromeRenderFrameHost->m_pMsg = nullptr;
				::PostAppMessage(::GetCurrentThreadId(), WM_TANGRAMMSG, (WPARAM)pMsg,
					20181008);
		}
			break;
		}
		LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
		return lRes;
	}

	LRESULT CHtmlWnd::OnParentChanged(UINT uMsg,
		WPARAM wParam,
		LPARAM lParam,
		BOOL&) {
		if (lParam) {
			ChromePlus::CChromeTangram* pChromeTangram = (ChromePlus::CChromeTangram*)g_pTangram;
			HWND hNewPWnd = (HWND)lParam;
			HWND hNewPWnd2 = ::GetParent(m_hWnd);
			bool bNewParent = false;
			if (hNewPWnd != hNewPWnd2)
			{
				hNewPWnd = hNewPWnd2;
				bNewParent = true;
			}
			::GetClassName(hNewPWnd, pChromeTangram->m_szBuffer, 256);
			CString strName = CString(pChromeTangram->m_szBuffer);
			if (strName == _T("Chrome_WidgetWin_1")) {
				ChromePlus::CBrowserWnd* pChromeBrowserWnd = nullptr;
				auto it = pChromeTangram->m_mapBrowserWnd.find(hNewPWnd);
				if (it == pChromeTangram->m_mapBrowserWnd.end()) {
					pChromeBrowserWnd = new CComObject<ChromePlus::CBrowserWnd>();
					pChromeBrowserWnd->SubclassWindow(hNewPWnd);
					pChromeTangram->m_mapBrowserWnd[hNewPWnd] = pChromeBrowserWnd;
					pChromeBrowserWnd->m_pBrowser = theApp.m_pActiveBrowser;
					if (pChromeBrowserWnd->m_pBrowser)
						pChromeBrowserWnd->m_pBrowser->m_pProxy = pChromeBrowserWnd;
				}
				else {
					pChromeBrowserWnd = it->second;
				}
				theApp.m_pActiveBrowser = pChromeBrowserWnd->m_pBrowser;
				if (pChromeBrowserWnd && m_hExtendWnd) {
					::SetParent(m_hExtendWnd,hNewPWnd);
					if (::IsWindowVisible(m_hWnd)) {
						pChromeBrowserWnd->m_pVisibleWebWnd = this;
						if (bNewParent)
						{
							::ShowWindow(m_hExtendWnd,SW_SHOW);
							theApp.m_pActiveBrowser = pChromeBrowserWnd->m_pBrowser;
							theApp.m_pActiveBrowser->m_pProxy = pChromeBrowserWnd;
							//pChromeBrowserWnd->ChromeDraw();
							pChromeBrowserWnd->m_pBrowser->LayoutBrowser();
						}
					}
				}
			}
		}
		LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
		return lRes;
	}

	LRESULT CHtmlWnd::OnDestroy(UINT uMsg,
		WPARAM wParam,
		LPARAM lParam,
		BOOL& /*bHandled*/) {
		if (m_hExtendWnd)
			::DestroyWindow(m_hExtendWnd);

		m_hExtendWnd = nullptr;

		if (m_bDevToolWnd) {
			if (m_pWebWnd) {
				m_pWebWnd->m_pDevToolWnd = nullptr;
				::PostMessage(::GetParent(m_pWebWnd->m_hWnd), WM_CHROMEDRAW, 0, 1);
			}
		}
		else {
			ChromePlus::CChromeTangram* pChromeTangram = (ChromePlus::CChromeTangram*)g_pTangram;
			CBrowserWnd* pPWnd = nullptr;
			auto it2 = pChromeTangram->m_mapBrowserWnd.find(::GetParent(m_hWnd));
			if (it2 != pChromeTangram->m_mapBrowserWnd.end())
			{
				pPWnd = it2->second;
				if (pPWnd)
				{
					if (pPWnd->m_pVisibleWebWnd == this)
						pPWnd->m_pVisibleWebWnd = nullptr;
				}
			}
			auto it = pChromeTangram->m_mapHtmlWnd.find(m_hWnd);
			if (it != pChromeTangram->m_mapHtmlWnd.end())
			{
				pChromeTangram->m_mapHtmlWnd.erase(it);
			}

			if (theApp.m_pActiveBrowser)
			{
				CBrowserWnd* pWnd = (CBrowserWnd*)(theApp.m_pActiveBrowser->m_pProxy);
				if (pWnd)
				{
					if (pWnd->m_pVisibleWebWnd == this)
						pWnd->m_pVisibleWebWnd = nullptr;
				}
			}
		}
		LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
		return lRes;
	}

	LRESULT CHtmlWnd::OnShowWindow(UINT uMsg,
		WPARAM wParam,
		LPARAM lParam,
		BOOL&) {
		CChromeTangram* pChromeTangram = (CChromeTangram*)g_pTangram;
		CBrowserWnd* pParent = nullptr;
		HWND hPWnd = ::GetParent(m_hWnd);
		auto it = pChromeTangram->m_mapBrowserWnd.find(hPWnd);
		if (it != pChromeTangram->m_mapBrowserWnd.end()) {
			pParent = it->second;
			if (wParam) {
				if (!m_bDevToolWnd)
					pParent->m_pVisibleWebWnd = this;

				if (m_pFrame)
				{
					pParent->m_pBrowser->LayoutBrowser();
					for (auto it : pChromeTangram->m_mapHtmlWnd)
					{
						if (it.second != this && it.second->m_hExtendWnd)
						{
							::SetParent(it.second->m_hExtendWnd, ::GetParent(it.first));
						}
					}
					::ShowWindow(m_hExtendWnd,SW_SHOW);
					::PostMessage(pParent->m_hWnd, WM_TANGRAMMSG, 20181013, 1);
				}
			}
		}
		LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
		return lRes;
	}

	void CHtmlWnd::OnFinalMessage(HWND hWnd) {
		CWindowImpl::OnFinalMessage(hWnd);
		delete this;
	}
}  // namespace ChromePlus



