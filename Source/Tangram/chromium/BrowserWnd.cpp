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
#include "../WndFrame.h"
#include "BrowserWnd.h"
#include "ChromeBrowserProxy.h"
#include "HtmlWnd.h"

namespace ChromePlus {
	CBrowserWnd::CBrowserWnd() {
		m_hDrawWnd = 0;
		m_pBrowser = nullptr;
		m_fdevice_scale_factor = 1.0f;
		m_strCurKey = _T("");
		m_pVisibleWebWnd = nullptr;
	}

	CBrowserWnd::~CBrowserWnd() {}

	LRESULT CBrowserWnd::ChromeDraw() {
		if (m_pVisibleWebWnd) {
			if (m_hDrawWnd == 0) {
				m_hDrawWnd = ::FindWindowEx(m_hWnd, nullptr, _T("Intermediate D3D Window"), nullptr);
			}
			if (m_hDrawWnd == 0) {
				m_hDrawWnd = ::FindWindowEx(m_hWnd, nullptr, _T("Intermediate Software Window"), nullptr);
			}
			CHtmlWnd* pWebWnd = m_pVisibleWebWnd;
			RECT rc;
			if (pWebWnd->m_hExtendWnd) {
				RECT rcWndRng;
				HWND hTangramHostWnd = pWebWnd->m_hChildWnd;

				if (pWebWnd->m_pDevToolWnd == nullptr) {
					::GetWindowRect(pWebWnd->m_hWnd, &rcWndRng);
				}
				else {
					if (pWebWnd->GetParent().m_hWnd == ::GetParent(pWebWnd->m_pDevToolWnd->m_hWnd))
						::GetWindowRect(pWebWnd->m_pDevToolWnd->m_hWnd, &rcWndRng);
					else
						::GetWindowRect(pWebWnd->m_hWnd, &rcWndRng);
				}
				ScreenToClient(&rcWndRng);
				if (::IsChild(m_hWnd, pWebWnd->m_hExtendWnd) == false) {
					::SetParent(pWebWnd->m_hExtendWnd, m_hWnd);
				}

				if (m_hDrawWnd) {
					if (hTangramHostWnd && pWebWnd) {
						::GetWindowRect(m_hWnd, &rc);
						::ScreenToClient(m_hDrawWnd, (LPPOINT)&rc);
						::ScreenToClient(m_hDrawWnd, ((LPPOINT)&rc) + 1);

						HRGN hGPUWndRgn = ::CreateRectRgn(rc.left, rc.top, rc.right, rc.bottom);
						HRGN hWebExtendWndRgn = ::CreateRectRgn(m_Rect.left, m_Rect.top, m_Rect.right, m_Rect.bottom);
						HRGN hTemp = ::CreateRectRgn(0, 0, 0, 0);
						HRGN hTemp2 = ::CreateRectRgn(0, 0, 0, 0);
						::CombineRgn(hTemp2, hGPUWndRgn, hWebExtendWndRgn, RGN_DIFF);
						HRGN hWndRgn = ::CreateRectRgn(rcWndRng.left, rcWndRng.top, rcWndRng.right, rcWndRng.bottom);
						::CombineRgn(hTemp, hTemp2, hWndRgn, RGN_OR);
						::SetWindowRgn(m_hDrawWnd, hTemp, true);
						::DeleteObject(hTemp2);
						::DeleteObject(hGPUWndRgn);
						::DeleteObject(hWndRgn);
						::DeleteObject(hWebExtendWndRgn);
					}
					else
					{
						::SetWindowRgn(m_hDrawWnd, ::CreateRectRgn(rcWndRng.left, rcWndRng.top,
							rcWndRng.right, rcWndRng.bottom), true);
					}
				}
				else
				{

				}
			}
			else {
				if (m_hDrawWnd) {
					GetWindowRect(&rc);
					CWindow wndGPU;
					wndGPU.Attach(m_hDrawWnd);
					wndGPU.ScreenToClient(&rc);
					::SetWindowRgn(m_hDrawWnd, ::CreateRectRgn(rc.left, rc.top, rc.right, rc.bottom), true);
				}
			}
		}

		return 0;
	}

	void CBrowserWnd::UpdateContentRect(RECT& rc, int nTopFix) {
		if (m_pVisibleWebWnd)
		{
			m_Rect.left = rc.left;
			m_Rect.top = nTopFix * m_fdevice_scale_factor;
			m_Rect.right = rc.left+ rc.right*m_fdevice_scale_factor;
			m_Rect.bottom = (nTopFix +rc.bottom - rc.top)*m_fdevice_scale_factor;
			HWND hExtendWnd = m_pVisibleWebWnd->m_hExtendWnd;
			if (hExtendWnd) {
				::SetWindowPos(hExtendWnd, HWND_TOP, rc.left, nTopFix *m_fdevice_scale_factor, rc.right*m_fdevice_scale_factor, (rc.bottom - rc.top)*m_fdevice_scale_factor, SWP_SHOWWINDOW | SWP_NOREDRAW | SWP_NOACTIVATE);
				HWND hHostWnd = m_pVisibleWebWnd->m_hHostWnd;
				if(hHostWnd==NULL)
					hHostWnd = m_pVisibleWebWnd->m_hChildWnd;
				if (::IsWindowVisible(hHostWnd) == false) {
					rc.bottom = rc.top + 1;
					rc.right = rc.left + 1;
					return;
				}
				RECT rc2;
				RECT rect;
				::GetWindowRect(hHostWnd, &rc2);
				if (::ScreenToClient(hExtendWnd, (LPPOINT)&rc2))
				{
					::ScreenToClient(hExtendWnd, ((LPPOINT)&rc2) + 1);
					::GetClientRect(hExtendWnd, &rect);

					rc.left += rc2.left / m_fdevice_scale_factor;
					rc.top += (rc2.top-rect.top) / m_fdevice_scale_factor;
					rc.right -= (rect.right - rc2.right) / m_fdevice_scale_factor;
					rc.bottom -= (rect.bottom - rc2.bottom) / m_fdevice_scale_factor;

					if (rc.bottom <= rc.top||rc.right<=rc.left)
					{
						rc.bottom = rc.top + 1;
						rc.right = rc.left + 1;
					}
				}
			}
		}
	};

	LRESULT CBrowserWnd::OnActivate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&) {
		LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
		if (LOWORD(wParam) != WA_INACTIVE) {
			if (m_pBrowser) {
				theApp.m_pActiveBrowser = m_pBrowser;
				theApp.m_pActiveBrowser->m_pProxy = this;
				::PostMessage(m_hWnd, WM_CHROMEDRAW, 0, 1);
			}
		}
		return lRes;
	}

	LRESULT CBrowserWnd::OnDeviceScaleFactorChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&) {
		LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
		m_fdevice_scale_factor = (float)lParam / 100;
		ChromeDraw();
		::PostMessage(m_hWnd, WM_CHROMEDRAW, 0, 1);
		return lRes;
	}

	LRESULT CBrowserWnd::OnTangramMsg(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&) {
		LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
		HWND hWnd = (HWND)lParam;
		switch (wParam) {
		case 0: {
			ChromePlus::CChromeTangram* pChromeTangram = (ChromePlus::CChromeTangram*)g_pTangram;
			pChromeTangram->m_pHtmlWndCreated = new CComObject<CHtmlWnd>;
			pChromeTangram->m_pHtmlWndCreated->SubclassWindow(hWnd);
			if (pChromeTangram->m_bCreatingDevTool == false)
			{
				m_pVisibleWebWnd = pChromeTangram->m_pHtmlWndCreated;
				m_pVisibleWebWnd->m_bDevToolWnd = false;
				pChromeTangram->m_mapHtmlWnd[hWnd] = m_pVisibleWebWnd;
			}
			else
			{
				pChromeTangram->m_bCreatingDevTool = false;
				pChromeTangram->m_pHtmlWndCreated->m_bDevToolWnd = true;
				if (m_pVisibleWebWnd) {
					m_pVisibleWebWnd->m_pDevToolWnd = pChromeTangram->m_pHtmlWndCreated;
					pChromeTangram->m_pHtmlWndCreated->m_pWebWnd = m_pVisibleWebWnd;
				}
			}
		} break;
		case 20181216:
			return (LRESULT)((CChromeBrowserProxy*)this);
			break;
		case 20181013:
		{
			m_pBrowser->LayoutBrowser();
			ChromeDraw();
		}break;
		}
		return lRes;
	}

	LRESULT CBrowserWnd::OnWindowPosChanged(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
	{
		LRESULT LRes = DefWindowProc(uMsg, wParam, lParam);
		if (m_pBrowser)
		{
			ChromeDraw();
		}
		return LRes;
	}

	LRESULT CBrowserWnd::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
	{
		CChromeTangram* pChromeTangram = (CChromeTangram*)g_pTangram;
		if (m_pVisibleWebWnd)
		{
			if (m_pVisibleWebWnd->m_hExtendWnd)
			{
				HWND hWnd = ::GetParent(m_pVisibleWebWnd->m_hWnd);
				if (hWnd != m_hWnd)
					::SetParent(m_pVisibleWebWnd->m_hExtendWnd, hWnd);
			}
			for (auto it : pChromeTangram->m_mapHtmlWnd)
			{
				if (it.second != m_pVisibleWebWnd && it.second->m_hExtendWnd)
				{
					::SetParent(it.second->m_hExtendWnd, ::GetParent(it.first));
				}
			}
		}
		
		if (g_pTangram->m_pCLRProxy)
		{
			IChromeWebBrowser* pIChromeWebBrowser = nullptr;
			QueryInterface(__uuidof(IChromeWebBrowser), (void**)&pIChromeWebBrowser);
			g_pTangram->m_pCLRProxy->OnDestroyChromeBrowser(pIChromeWebBrowser);
		}

		m_pVisibleWebWnd = nullptr;
		auto it = pChromeTangram->m_mapBrowserWnd.find(m_hWnd);
		if (it != pChromeTangram->m_mapBrowserWnd.end()) {
			pChromeTangram->m_mapBrowserWnd.erase(it);
		}

		if (pChromeTangram->m_mapBrowserWnd.size() == 0) {
			if (g_pTangram->m_pCLRProxy)
			{
				if (g_pTangram->m_pTangramCLRAppProxy)
					g_pTangram->m_pTangramCLRAppProxy->OnTangramClose();

				theApp.m_bClose = true;
			}

			::PostAppMessage(::GetCurrentThreadId(), WM_CHROMEHELPWND, 0, 0);
			if (pChromeTangram->m_hCBTHook) {
				UnhookWindowsHookEx(pChromeTangram->m_hCBTHook);
				pChromeTangram->m_hCBTHook = nullptr;
			}
		}
		return DefWindowProc(uMsg, wParam, lParam);
	}

	LRESULT CBrowserWnd::OnChromeDraw(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
	{
		DefWindowProc(uMsg, wParam, lParam);
		if (m_pVisibleWebWnd)
		{
			switch (lParam)
			{
			case 1:
			{
				if (m_pBrowser)
				{
					switch (wParam)
					{
					case 1:
					{
						if (m_pVisibleWebWnd->m_pFrame)
						{
							IWndNode* pNode = nullptr;
							m_pVisibleWebWnd->m_pFrame->Extend(m_pVisibleWebWnd->m_strCurKey.AllocSysString(), CComBSTR(""), &pNode);
							return 0;
						}
					}
						break;
					case 2:
						if (m_pVisibleWebWnd->m_pFrame)
							ChromeDraw();
							//::PostMessage(m_hWnd, WM_CHROMEDRAW, 0, 2);
						break;
					}
					ChromeDraw();
					m_pBrowser->LayoutBrowser();
				}
			}
			break;
			}
		}
		return 0;
	}

	void CBrowserWnd::OnFinalMessage(HWND hWnd) {
		CWindowImpl::OnFinalMessage(hWnd);
		delete this;
	}

	STDMETHODIMP CBrowserWnd::OpenURL(BSTR bstrURL, BrowserWndOpenDisposition nDisposition, BSTR bstrKey, BSTR bstrXml)
	{
		if (bstrURL != L"")
		{
			m_pBrowser->OpenURL(OLE2W(bstrURL), nDisposition, nullptr);
		}
		return S_OK;
	}
}  // namespace ChromePlus

