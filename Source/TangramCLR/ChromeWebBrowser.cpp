#include "stdafx.h"
#include "ChromeWebBrowser.h"

namespace TangramCLR
{

	ChromeWebBrowser::ChromeWebBrowser(IChromeWebBrowser* pChromeWebBrowser)
	{
		m_pChromeWebBrowserHost = pChromeWebBrowser;
	}

	void ChromeWebBrowser::OpenURL(String^ url, Disposition disposition, String^ key, String^ xml)
	{
		BSTR bstrUrl = STRING2BSTR(url);
		BSTR bstrKey = STRING2BSTR(key);
		BSTR bstrXml = STRING2BSTR(xml);
		if (m_pChromeWebBrowserHost)
		{
			m_pChromeWebBrowserHost->OpenURL(bstrUrl, (BrowserWndOpenDisposition)disposition, bstrKey, bstrXml);
		}
	}

}