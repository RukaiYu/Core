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

#include "../stdafx.h"
#include "ChromeProxy.h"

namespace ChromePlus
{
	class ATL_NO_VTABLE CHtmlWnd :
		public CWindowImpl<CHtmlWnd, CWindow>,
		public CChromeRenderFrameHostProxyBase,
		public CComObjectRootEx<CComSingleThreadModel>,
		public IDispatchImpl<IChromeWebContent, &_uuidof(IChromeWebContent)>
	{
	public:
		CHtmlWnd();
		~CHtmlWnd() override;

		BOOL			m_bDevToolWnd;
		HWND			m_hExtendWnd;
		HWND			m_hChildWnd;
		HWND			m_hHostWnd;
		CString			m_strURL;
		CString			m_strCurKey;
		CWndFrame*		m_pFrame;

		CHtmlWnd*		m_pWebWnd;
		CHtmlWnd*		m_pDevToolWnd;

		BEGIN_MSG_MAP(CHtmlWnd)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_TANGRAMMSG, OnTangramMsg)
			MESSAGE_HANDLER(WM_SHOWWINDOW, OnShowWindow)
			MESSAGE_HANDLER(WM_CHROMEOPENWINDOWMSG, OnOpenDOMWindow)
			MESSAGE_HANDLER(WM_CHROMEWNDPARENTCHANGED, OnParentChanged)
			MESSAGE_HANDLER(WM_CHROMEIPCMSG, OnChromeIPCMessage)
		END_MSG_MAP()

		BEGIN_COM_MAP(CHtmlWnd)
			COM_INTERFACE_ENTRY(IChromeWebContent)
			COM_INTERFACE_ENTRY(IDispatch)
		END_COM_MAP()
	protected:
		ULONG InternalAddRef() { return 1; }
		ULONG InternalRelease() { return 1; }

	private:
		void OnFinalMessage(HWND hWnd) override;
		LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
		LRESULT OnShowWindow(UINT, WPARAM, LPARAM, BOOL&);
		LRESULT OnTangramMsg(UINT, WPARAM, LPARAM, BOOL&);
		LRESULT OnOpenDOMWindow(UINT, WPARAM, LPARAM, BOOL&);
		LRESULT OnParentChanged(UINT, WPARAM, LPARAM, BOOL&);
		LRESULT OnChromeIPCMessage(UINT, WPARAM, LPARAM, BOOL&);
	};
};