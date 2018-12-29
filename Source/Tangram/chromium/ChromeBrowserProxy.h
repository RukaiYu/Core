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
#include "../TangramCore.h"
#include "ChromeProxy.h"

namespace ChromePlus
{
	class CBrowserWnd;
	class CHtmlWnd;
}

namespace ChromePlus
{
	class  CChromeTangram :
		public CTangram
	{
	public:
		CChromeTangram();
		virtual ~CChromeTangram();

		BOOL									m_bCreatingDevTool;

		CString									m_strDefaultXml;
		CHtmlWnd*								m_pHtmlWndCreated;
		map<HWND, CHtmlWnd*>					m_mapHtmlWnd;
		map<HWND, CBrowserWnd*>					m_mapBrowserWnd;

		void CreateCommonDesignerToolBar();		

		STDMETHOD(CreateOfficeDocument)(BSTR bstrXml);
		STDMETHOD(get_ActiveChromeBrowserWnd)(IChromeWebBrowser** ppChromeWebBrowser) ;
	};
}

