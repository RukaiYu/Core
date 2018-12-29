#pragma once

#include "BrowserWndOpenDisposition.h"

using namespace System;
using System::Runtime::InteropServices::Marshal;

namespace TangramCLR
{

	public ref class ChromeWebBrowser
	{
	public:
		IChromeWebBrowser*				m_pChromeWebBrowserHost;
		ChromeWebBrowser(IChromeWebBrowser* pChromeWebBrowser);
		void OpenURL(String^ url, Disposition disposition, String^ key, String^ xml);
	};

}