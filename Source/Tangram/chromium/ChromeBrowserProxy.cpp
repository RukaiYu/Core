#include "../stdafx.h"
#include "../TangramApp.h"
#include "../TangramCore.h"
#include "../nodewnd.h"
#include "BrowserWnd.h"
#include "ChromeBrowserProxy.h"

namespace ChromePlus
{
	CChromeTangram::CChromeTangram()
	{
		m_bCreatingDevTool = false;
		m_pHtmlWndCreated = nullptr;
		m_strDefaultXml = _T("<default><window><node name=\"tangram\" id=\"HostView\"/></window></default>");
	}

	CChromeTangram::~CChromeTangram()
	{
	}

	void CChromeTangram::CreateCommonDesignerToolBar()
	{

	}

	STDMETHODIMP CChromeTangram::CreateOfficeDocument(BSTR bstrXml)
	{
		return S_OK;
	}

	STDMETHODIMP CChromeTangram::get_ActiveChromeBrowserWnd(IChromeWebBrowser** ppChromeWebBrowser)
	{
		if (theApp.m_pActiveBrowser->m_pProxy)
		{
			CBrowserWnd* pCBrowserWnd = (CBrowserWnd*)theApp.m_pActiveBrowser->m_pProxy;
			pCBrowserWnd->QueryInterface(__uuidof(IChromeWebBrowser), (void**)ppChromeWebBrowser);
		}
		return S_OK;
	}
}
