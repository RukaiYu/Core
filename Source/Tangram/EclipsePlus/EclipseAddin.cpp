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
//https://github.com/eclipse/rt.equinox.framework/tree/master/features/org.eclipse.equinox.executable.feature

#include "../stdafx.h"
#include "../TangramApp.h"
#include "../WndNode.h"
#include "../WndFrame.h"
#include "../TangramCtrl.h"
#include "EclipseAddin.h"

#include "atlcom.h"

extern jstring newJavaString(JNIEnv *env, _TCHAR * str);
extern CTangramApplicationImpl* m_pTangramApplicationImpl;

CEclipseExtender::CEclipseExtender()
{
}

CEclipseExtender::~CEclipseExtender()
{
}

STDMETHODIMP CEclipseExtender::Close()
{
	delete this;
	return S_OK;
}

STDMETHODIMP CEclipseExtender::get_ActiveWorkBenchWindow(BSTR bstrID, IWorkBenchWindow** pVal)
{
	CString strID = OLE2T(bstrID);
	strID.Trim();
	if (strID==_T(""))
	{
		if (g_pTangram->m_pActiveEclipseWnd)
		{
			*pVal = (IWorkBenchWindow*)g_pTangram->m_pActiveEclipseWnd;
			(*pVal)->AddRef();
		}
	}
	else
	{
		ITangram* pTangram = nullptr;
		g_pTangram->get_RemoteTangram(bstrID, &pTangram);
		if (pTangram)
		{
			IWorkBenchWindow* pRet = nullptr;
			ITangramExtender* pExtender = nullptr;
			pTangram->get_Extender(&pExtender);
			if (pExtender)
			{
				CComQIPtr<IEclipseExtender> pEclipse(pExtender);
				if(pEclipse)
					pEclipse->get_ActiveWorkBenchWindow(bstrID, &pRet);
				if (pRet)
				{
					*pVal = pRet;
					(*pVal)->AddRef();
				}
			}
		}
	}

	return S_OK;
}

CEclipseWnd::CEclipseWnd(void)
{
	m_bInit = false;
	m_pDoc = nullptr;
	m_pPage = nullptr;
	m_pFrame = nullptr;
	m_pCurNode = nullptr;
	m_pHostNode = nullptr;
	m_strPath = _T("");
	m_strNodeName = _T("");
	m_strDocKey = _T("");
	m_strFrameID = _T("");
	m_pAppProxy = nullptr;
	g_pTangram->m_pActiveEclipseWnd = this;
}

CEclipseWnd::~CEclipseWnd(void) 
{
	if (m_pPage)
	{
		HWND hWnd = m_pPage->m_hWnd;
		auto it = g_pTangram->m_mapWindowPage.find(hWnd);
		if (it != g_pTangram->m_mapWindowPage.end())
		delete m_pPage;
	}
	ATLTRACE(_T("delete CEclipseWnd:%x\n"), this);
}

void CEclipseWnd::OnFinalMessage(HWND hWnd)
{
	auto it = g_pTangram->m_mapWorkBenchWnd.find(hWnd);
	if (it != g_pTangram->m_mapWorkBenchWnd.end())
		g_pTangram->m_mapWorkBenchWnd.erase(it);
	if (g_pTangram->m_pCLRProxy)
	{
		g_pTangram->m_pCLRProxy->ReleaseTangramObj((IWorkBenchWindow*)this);
	}
	CWindowImpl::OnFinalMessage(hWnd);
	delete this;
}

void CEclipseWnd::CreatePage(BOOL bSaveToConfigFile)
{
	if (m_pPage == nullptr)
	{
		IWndPage* pPage = nullptr;
		g_pTangram->CreateWndPage((LONGLONG)m_hWnd, &pPage);
		if (pPage == nullptr)
			return;
		m_pPage = (CWndPage*)pPage;
		HMENU hMenu = ::GetMenu(m_hWnd);
		if (hMenu)
		{
			MENUITEMINFO mii;
			mii.fMask = MIIM_STRING;
			mii.cbSize = sizeof(MENUITEMINFO);
			int nCount = ::GetMenuItemCount(hMenu);
			for (int i = nCount - 1; i >0; i--)
			{
				mii.cch = 256;
				mii.fType = MFT_STRING;
				mii.dwTypeData = g_pTangram->m_szBuffer;
				GetMenuItemInfo(hMenu, i, true, &mii);
				CString str = mii.dwTypeData;
				if (str == _T("&Window"))
				{
					HMENU hSubMenu = ::GetSubMenu(hMenu, i);
					int nCount = ::GetMenuItemCount(hSubMenu);
					for (int i = 0; i<nCount; i++)
					{
						mii.fMask = MIIM_STRING | MIIM_ID;
						mii.cch = 256;
						mii.dwTypeData = g_pTangram->m_szBuffer;
						GetMenuItemInfo(hSubMenu, i, true, &mii);
						str = mii.dwTypeData;
						if (str == _T("&New Window"))
						{
							m_nNewWinCmdID = mii.wID;
							break;
						}
					}
					break;
				}
			}
		}
		if (::IsWindowVisible(m_hWnd))
		{
			m_strDocKey = g_pTangram->m_strCurrentEclipsePagePath;
			g_pTangram->m_strCurrentEclipsePagePath = _T("");
			Show(m_strDocKey);
			auto it2 = g_pTangram->m_mapValInfo.find(_T("newmdtframe"));
			if (it2 != g_pTangram->m_mapValInfo.end())
			{
				m_strFrameID = OLE2T(it2->second.bstrVal);
				g_pTangram->m_mapValInfo.erase(it2);
			}
		}
	}
}

STDMETHODIMP CEclipseWnd::get_Count(long* pCount)
{
	*pCount = (long)m_mapCtrl.size();
	return S_OK;
}

STDMETHODIMP CEclipseWnd::get_Ctrl(VARIANT vIndex, IEclipseCtrl **ppCtrl)
{
	if (vIndex.vt == VT_I4)
	{
		long lCount = m_mapCtrl.size();
		int iIndex = vIndex.lVal;
		if (iIndex < 0 || iIndex >= lCount) return E_INVALIDARG;

		auto it = m_mapCtrl.begin();
		int iPos = 0;
		while (it != m_mapCtrl.end())
		{
			if (iPos == iIndex) break;
			iPos++;
			it++;
		}

		*ppCtrl = it->second;
		return S_OK;
	}

	return S_OK;
}

STDMETHODIMP CEclipseWnd::CloseTangramUI()
{
	if (m_pFrame)
	{
		HWND _hWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), L"", WS_CHILD, 0, 0, 0, 0, g_pTangram->m_hHostWnd, NULL, AfxGetInstanceHandle(), NULL);
		HWND _hChildWnd = ::CreateWindowEx(NULL, _T("Tangram Window Class"), L"", WS_CHILD, 0, 0, 0, 0, (HWND)_hWnd, NULL, AfxGetInstanceHandle(), NULL);
		if (::IsWindow(m_hWnd))
		{
			m_pFrame->ModifyHost((LONGLONG)_hChildWnd);
			::DestroyWindow(_hWnd);
		}
		m_pFrame = nullptr;
		delete m_pPage;
		m_pPage = nullptr;
	}

	return S_OK;
}

STDMETHODIMP CEclipseWnd::get__NewEnum(IUnknown** ppVal)
{
	typedef CComEnum<IEnumVARIANT, &IID_IEnumVARIANT, VARIANT, _Copy<VARIANT>>
		CComEnumVariant;

	CComObject<CComEnumVariant> *pe = 0;
	HRESULT hr = pe->CreateInstance(&pe);

	if (SUCCEEDED(hr))
	{
		pe->AddRef();
		long nLen = (long)m_mapCtrl.size();
		VARIANT* rgvar = new VARIANT[nLen];
		ZeroMemory(rgvar, sizeof(VARIANT)*nLen);
		VARIANT* pItem = rgvar;
		for (auto it : m_mapCtrl)
		{
			IUnknown* pDisp = nullptr;
			CEclipseCtrl* pObj = it.second;
			hr = pObj->QueryInterface(IID_IUnknown, (void**)&pDisp);
			pItem->vt = VT_UNKNOWN;
			pItem->punkVal = pDisp;
			if (pItem->punkVal != nullptr)
				pItem->punkVal->AddRef();
			pItem++;
			pDisp->Release();
		}

		hr = pe->Init(rgvar, &rgvar[nLen], 0, AtlFlagTakeOwnership);
		if (SUCCEEDED(hr))
			hr = pe->QueryInterface(IID_IUnknown, (void**)ppVal);
		pe->Release();
	}
	return S_OK;
}

STDMETHODIMP CEclipseWnd::get_Handle(LONGLONG* pVal)
{
	*pVal = (LONGLONG)m_hClient;
	return S_OK;
}

STDMETHODIMP CEclipseWnd::get_WndPage(IWndPage** pVal)
{
	if (m_pPage)
		*pVal = m_pPage;

	return S_OK;
}

STDMETHODIMP CEclipseWnd::get_WndFrame(IWndFrame** pVal)
{
	if (m_pFrame)
		*pVal = m_pFrame;
	return S_OK;
}

STDMETHODIMP CEclipseWnd::Extend(BSTR bstrKey, BSTR bstrXml, IWndNode** ppNode)
{
	CString strKey = OLE2T(bstrKey);
	strKey.Trim();
	HRESULT hr = S_FALSE;
	if (m_hClient == NULL)
		return S_FALSE;
	if (m_pPage == nullptr)
	{
		IWndPage* pPage = nullptr;
		g_pTangram->CreateWndPage((LONGLONG)m_hWnd, &pPage);
		if (pPage == nullptr)
			return S_FALSE;
		m_pPage = (CWndPage*)pPage;
		if (m_pFrame == nullptr)
		{
			IWndFrame* pFrame = nullptr;
			m_pPage->CreateFrame(CComVariant(0), CComVariant((LONGLONG)m_hClient), CComBSTR(L"default"), &pFrame);
			if (pFrame == nullptr)
			{
				delete m_pPage;
				m_pPage = nullptr;
			}
			m_pFrame = (CWndFrame*)pFrame;
		}
	}
	else
	{
		m_pFrame = (CWndFrame*)::SendMessage(m_hClient, WM_TANGRAMDATA, 0, 1992);
		if (m_pFrame == nullptr)
		{
			auto it = m_pPage->m_mapFrame.find(m_hClient);
			if (it != m_pPage->m_mapFrame.end())
				m_pFrame = it->second;
			else
			{
				IWndFrame* pFrame = nullptr;
				m_pPage->CreateFrame(CComVariant(0), CComVariant((LONGLONG)m_hClient), CComBSTR(L"default"), &pFrame);
				if (pFrame == nullptr)
				{
					delete m_pPage;
					m_pPage = nullptr;
				}
				m_pFrame = (CWndFrame*)pFrame;
			}
		}
	}
	if (m_pFrame)
	{
		hr = m_pFrame->Extend(bstrKey, bstrXml, ppNode);
		if (*ppNode)
		{
			m_pCurNode = (CWndNode*)*ppNode;
		}
	}
	return hr;
} 

STDMETHODIMP CEclipseWnd::ExtendEx(BSTR bstrKey, BSTR bstrXml, IWndNode** ppNode)
{
	HRESULT hr = Extend(bstrKey, bstrXml, ppNode);
	if (*ppNode)
	{
		m_pCurNode = (CWndNode*)*ppNode;
		m_pCurNode->put_SaveToConfigFile(true);
	}
	return hr;
} 

STDMETHODIMP CEclipseWnd::ExtendInView(int nIndex, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode)
{
	IEclipseCtrl* pCtrl = nullptr;
	get_Ctrl(CComVariant((int)nIndex), &pCtrl);
	if (pCtrl)
	{
		IWndNode* pNode = nullptr;
		return pCtrl->Extend(CComBSTR(L"EclipseView"), bstrKey, bstrXml, ppRetNode);
	}

	return S_OK;
}

STDMETHODIMP CEclipseWnd::Active()
{
	SetFocus();

	return S_OK;
}

STDMETHODIMP CEclipseWnd::get_TangramCtrl(LONGLONG hWnd, IEclipseCtrl** pVal)
{
	auto it = m_mapCtrl.find((HWND)hWnd);
	if (it != m_mapCtrl.end())
	{
		*pVal = it->second;
		(*pVal)->AddRef();
	}
	return S_OK;
}

void CEclipseWnd::Show(CString strID)
{
	//::MessageBox(nullptr, strID, _T(""), MB_OK);
	if (m_pDoc)
		return;
	LONG_PTR data = 0;
	if(::IsWindow(m_hClient))
		data = ::GetWindowLongPtr(m_hClient, GWLP_USERDATA);

	if (data==0&&g_pTangram->m_pTangramAppProxy)
	{
		m_strPath = strID;
		if (m_strPath == _T(""))
		{
			m_strPath = g_pTangram->m_strAppDataPath + _T("default.tangramdoc");
		}
		if (::PathFileExists(m_strPath)&&::IsChild(m_hWnd,m_hClient))
		{
			TangramDocInfo m_TangramDocInfo;
			m_TangramDocInfo.m_strAppProxyID = m_TangramDocInfo.m_strDocID = m_TangramDocInfo.m_strMainFrameID = m_TangramDocInfo.m_strTangramData = m_TangramDocInfo.m_strTangramID = _T("");
			g_pTangram->GetTangramInfo(m_strPath, &m_TangramDocInfo);
			if (m_TangramDocInfo.m_strTangramID == _T("19631222199206121965060119820911"))
			{
				CTangramAppProxy* pProxy = nullptr;
				auto it = g_pTangram->m_mapTangramAppProxy.find(m_TangramDocInfo.m_strAppProxyID.MakeLower());
				if (it != g_pTangram->m_mapTangramAppProxy.end())
					pProxy = it->second;
				else
				{
					CString strPath = g_pTangram->m_strAppPath;
					HMODULE hHandle = nullptr;
					CString strAppName = _T("");
					int nPos = m_TangramDocInfo.m_strAppProxyID.Find(_T("."));
					if (nPos != -1)
						strAppName = m_TangramDocInfo.m_strAppProxyID.Left(nPos);
					CString strdll = strPath + m_TangramDocInfo.m_strAppProxyID + _T("\\") + strAppName + _T(".dll");
					g_pTangram->m_strCurrentFrameID = m_TangramDocInfo.m_strMainFrameID;
					if (::PathFileExists(strdll))
						hHandle = ::LoadLibrary(strdll);
					if (hHandle == nullptr)
					{
						strdll = g_pTangram->m_strAppCommonDocPath2 + m_TangramDocInfo.m_strAppProxyID + _T("\\") + strAppName + _T(".dll");
						if (::PathFileExists(strdll))
							hHandle = ::LoadLibrary(strdll);
					}
					if (hHandle)
					{
						it = g_pTangram->m_mapTangramAppProxy.find(m_TangramDocInfo.m_strAppProxyID.MakeLower());
						if (it != g_pTangram->m_mapTangramAppProxy.end())
						{
							pProxy = it->second;
						}
					}
					g_pTangram->m_strCurrentFrameID = _T("");
				}
				if (pProxy)
				{
					m_pAppProxy = pProxy;
					g_pTangram->m_pActiveAppProxy = pProxy;
					auto it2 = g_pTangram->m_mapTemplateInfo.find(m_TangramDocInfo.m_strDocID.MakeLower());
					if (it2 != g_pTangram->m_mapTemplateInfo.end())
					{
						pProxy->m_hCreatingView = m_hClient;
						::SetWindowText(m_hClient, g_pTangram->m_strAppKey);
						if (data == 0)
						{
							::SetWindowLongPtr(m_hClient, GWLP_USERDATA, (LONG_PTR)1);
							m_pDoc = (CTangramDoc*)it->second->OpenDocument(it2->second, m_strPath, true);
							if (m_pDoc)
							{
								m_pDoc->m_strPath = m_strPath;
							}
						}
					}
				}
			}
			auto it2 = g_pTangram->m_mapValInfo.find(_T("newmdtframe"));
			if (it2 != g_pTangram->m_mapValInfo.end())
			{
				m_strFrameID = OLE2T(it2->second.bstrVal);
				g_pTangram->m_mapValInfo.erase(it2);
			}
		}
	}
}

LRESULT CEclipseWnd::OnShowWindow(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& )
{
	if (m_pDoc == nullptr)
	{
		if (g_pTangram->m_strWorkBenchStrs != _T(""))
		{
			int nPos = g_pTangram->m_strWorkBenchStrs.Find(_T("|"));
			m_strDocKey = g_pTangram->m_strWorkBenchStrs.Left(nPos);
			g_pTangram->m_strWorkBenchStrs = g_pTangram->m_strWorkBenchStrs.Mid(nPos + 1);
		}

		if (g_pTangram->m_strCurrentEclipsePagePath != _T(""))
		{
			m_strDocKey = g_pTangram->m_strCurrentEclipsePagePath;
			g_pTangram->m_strCurrentEclipsePagePath = _T("");
		}
		if(::IsWindow(m_hClient))
			Show(m_strDocKey);
	
		if (wParam)
			::PostMessage(g_pTangram->m_hTangramWnd, WM_TANGRAMAPPINIT, 0, 1992);
	}
	LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
	return lRes;
}

LRESULT CEclipseWnd::OnCommand(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	if (wParam == m_nNewWinCmdID)
	{
		if (g_pTangram->m_strCurrentEclipsePagePath == _T(""))
		{
			CComBSTR bstrTemplate(L"");
			g_pTangram->GetDocTemplateXml(CComBSTR("Please select New Eclipse Window Bench Template:"), CComBSTR(g_pTangram->m_strAppDataPath+_T("workbench\\")), _T(".eclipse"), &bstrTemplate);
			g_pTangram->m_strCurrentEclipsePagePath = OLE2T(bstrTemplate);
		}
	}
	LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
	return lRes;
}

LRESULT CEclipseWnd::OnTangramMsg(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& )
{
	switch (wParam)
	{
	case 1963:
		{
			return ::SendMessage(::GetParent(m_hClient), WM_QUERYAPPPROXY, wParam, 0);
		}
		break;
	case 19820911:
		{
			if (lParam == 0)
			{
				::PostMessage(m_hWnd, WM_TANGRAMMSG, 19820911, 1);
			}
			if (lParam == 1)
			{
				CWndFrame* pFrame = (CWndFrame*)::SendMessage(m_hClient, WM_TANGRAMDATA, 0, 1992);
				if (pFrame)
				{
					pFrame->m_nWndFrameType = EclipseWorkBenchFrame;
					pFrame->m_pWorkBenchFrame = this;
				}
				else
				{
					if (m_pDoc == nullptr)
					{
						if (g_pTangram->m_strWorkBenchStrs != _T(""))
						{
							int nPos = g_pTangram->m_strWorkBenchStrs.Find(_T("|"));
							m_strDocKey = g_pTangram->m_strWorkBenchStrs.Left(nPos);
							g_pTangram->m_strWorkBenchStrs = g_pTangram->m_strWorkBenchStrs.Mid(nPos + 1);
						}
						if (g_pTangram->m_strCurrentEclipsePagePath != _T(""))
						{
							m_strDocKey = g_pTangram->m_strCurrentEclipsePagePath;
							g_pTangram->m_strCurrentEclipsePagePath = _T("");
						}
						LONG_PTR data = 0;
						if (::IsWindow(m_hClient))
						{
							data = ::GetWindowLongPtr(m_hClient, GWLP_USERDATA);
							if (data == 0)
							{
								Show(m_strDocKey);
								break;
							}
						}

						if (wParam&&data==0)
						{
							::PostMessage(g_pTangram->m_hTangramWnd, WM_TANGRAMAPPINIT, 0, 1992);
						}
					}
				}
			}
		}
		break;
	}

	LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
	return lRes;
}

LRESULT CEclipseWnd::OnClose(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& )
{
	if (m_strPath!=_T(""))
		::SendMessage(::GetParent(m_hClient),WM_QUERYAPPPROXY,0,19631222);

	LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
	return lRes;
}

LRESULT CEclipseWnd::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& )
{
	CTangramXmlParse xmlParse;
	if (xmlParse.LoadFile(g_pTangram->m_strConfigDataFile))
	{
		CTangramXmlParse* pParse = xmlParse.GetChild(_T("openedworkbench"));
		if (pParse == nullptr)
		{
			pParse = xmlParse.AddNode(_T("openedworkbench"));
		}
		if (pParse)
		{
			CString strWorkBenchStrs = m_strDocKey + _T("|");
			strWorkBenchStrs += pParse->text();
			pParse->put_text(strWorkBenchStrs);
			xmlParse.SaveFile(g_pTangram->m_strConfigDataFile);
		}
	}
	if (g_pTangram->m_pActiveEclipseWnd == this)
		g_pTangram->m_pActiveEclipseWnd = nullptr;
	auto it = g_pTangram->m_mapWorkBenchWnd.find(m_hWnd);
	if (it != g_pTangram->m_mapWorkBenchWnd.end())
	{
		g_pTangram->m_mapWorkBenchWnd.erase(it);
		if (g_pTangram->m_mapWorkBenchWnd.size() == 0&&::IsWindow(g_pTangram->m_hHostWnd))
		{
			::DestroyWindow(g_pTangram->m_hHostWnd);
		}
	}
	LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
	return lRes;
}

LRESULT CEclipseWnd::OnActivate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& )
{
	LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
	if (LOWORD(wParam) != WA_INACTIVE)
	{
		g_pTangram->m_pActiveEclipseWnd = this;
		g_pTangram->m_pActiveAppProxy = m_pAppProxy;
	}
	if(LOWORD(wParam)==WA_INACTIVE&&::IsWindow(m_hClient))
		SendMessage(::GetParent(m_hClient), WM_ACTIVATE, wParam, lParam);
	return lRes;
}

LRESULT CEclipseWnd::OnQueryAppProxy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& )
{
	LRESULT lRes = ::SendMessage(::GetParent(m_hClient), uMsg, wParam, 19631992);
	if (lRes)
	{
		if (m_pAppProxy==nullptr)
			m_pAppProxy = (CTangramAppProxy*)lRes;

		return lRes;
	}
	
	lRes = DefWindowProc(uMsg, wParam, lParam);
	return lRes;
}

LRESULT CEclipseWnd::OnUCMAMsg(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& )
{
	UCMAMSGInfo* pUCMAMSGInfo = (UCMAMSGInfo*)wParam;
	IWndNode* pNode = nullptr;
	Extend(pUCMAMSGInfo->m_strMsg.AllocSysString(), pUCMAMSGInfo->m_strMsg.AllocSysString(), &pNode);
	delete pUCMAMSGInfo;
	LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
	return lRes;
}

CEclipseCtrl::CEclipseCtrl()
{
	m_varTag.vt = VT_BSTR;
	m_varTag.bstrVal = ::SysAllocString(L"");
	m_hEclipseViewWnd = NULL;
	m_pCurNode = nullptr;
	m_pEclipseWnd = nullptr;
	m_pTangramWndPageProxy = nullptr;
}

STDMETHODIMP CEclipseCtrl::get_AppKeyValue(BSTR bstrKey, VARIANT* pVal)
{
	if (g_pTangram)
	{
		return g_pTangram->get_AppKeyValue(bstrKey, pVal);
	}
	return S_FALSE;
}

STDMETHODIMP CEclipseCtrl::put_AppKeyValue(BSTR bstrKey, VARIANT newVal)
{
	if (g_pTangram)
	{
		return g_pTangram->put_AppKeyValue(bstrKey, newVal);
	}

	return S_OK;
}

STDMETHODIMP CEclipseCtrl::get_tag(VARIANT* pVal)
{
	*pVal = m_varTag;
	return S_OK;
}

STDMETHODIMP CEclipseCtrl::put_tag(VARIANT newVal)
{
	::VariantClear(&m_varTag);
	m_varTag = newVal;
	return S_OK;
}

HRESULT CEclipseCtrl::Fire_PageLoaded(IDispatch* sender, BSTR url)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[2];
		avarParams[1] = sender;
		avarParams[1].vt = VT_DISPATCH;
		avarParams[0] = url;
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
				hr = pConnection->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnPageLoaded(sender, OLE2T(url));
	return hr;
}

HRESULT CEclipseCtrl::Fire_NodeCreated(IWndNode * pNodeCreated)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[1];
		avarParams[0] = pNodeCreated;
		avarParams[0].vt = VT_DISPATCH;
		DISPPARAMS params = { avarParams, NULL, 1, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			IUnknown* punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();
			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection);
			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(2, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnNodeCreated(pNodeCreated);
	return hr;
}

HRESULT CEclipseCtrl::Fire_AddInCreated(IWndNode * pRootNode, IDispatch * pAddIn, BSTR bstrID, BSTR bstrAddInXml)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[4];
		avarParams[3] = pRootNode;
		avarParams[2] = pAddIn;
		avarParams[2].vt = VT_DISPATCH;
		avarParams[1] = bstrID;
		avarParams[1].vt = VT_BSTR;
		avarParams[0] = bstrAddInXml;
		avarParams[0].vt = VT_BSTR;
		DISPPARAMS params = { avarParams, NULL, 4, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			IUnknown* punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();
			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection);
			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(3, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnAddInCreated(pRootNode, pAddIn, OLE2T(bstrID), OLE2T(bstrAddInXml));
	return hr;
}

HRESULT CEclipseCtrl::Fire_BeforeExtendXml(BSTR bstrXml, LONGLONG hWnd)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[2];
		avarParams[1] = bstrXml;
		avarParams[1].vt = VT_BSTR;
		avarParams[0] = hWnd;
		avarParams[0].vt = VT_I8;
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
				hr = pConnection->Invoke(4, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnBeforeExtendXml(OLE2T(bstrXml), (HWND)hWnd);
	return hr;
}

HRESULT CEclipseCtrl::Fire_ExtendXmlComplete(BSTR bstrXml, LONGLONG hWnd, IWndNode * pRetRootNode)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[3];
		avarParams[2] = bstrXml;
		avarParams[2].vt = VT_BSTR;
		avarParams[1] = hWnd;
		avarParams[1].vt = VT_I8;
		avarParams[0] = pRetRootNode;
		avarParams[0].vt = VT_DISPATCH;
		DISPPARAMS params = { avarParams, NULL, 3, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			IUnknown* punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();
			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection);
			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(5, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnExtendXmlComplete(OLE2T(bstrXml), (HWND)hWnd, pRetRootNode);

	return hr;
}

HRESULT CEclipseCtrl::Fire_Destroy()
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		DISPPARAMS params = { NULL, NULL, 0, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			IUnknown* punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();
			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection);
			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(6, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnDestroy();

	return hr;
}

HRESULT CEclipseCtrl::Fire_NodeMouseActivate(IWndNode * pActiveNode)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[1];
		avarParams[0] = pActiveNode;
		avarParams[0].vt = VT_DISPATCH;
		DISPPARAMS params = { avarParams, NULL, 1, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			CComPtr<IUnknown> punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();
			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection.p);
			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(7, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}

	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnNodeMouseActivate(pActiveNode);
	return hr;
}

HRESULT CEclipseCtrl::Fire_ClrControlCreated(IWndNode * Node, IDispatch * Ctrl, BSTR CtrlName, LONGLONG CtrlHandle)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[4];
		avarParams[3] = Node;
		avarParams[3].vt = VT_DISPATCH;
		avarParams[2] = Ctrl;
		avarParams[2].vt = VT_DISPATCH;
		avarParams[1] = CtrlName;
		avarParams[1].vt = VT_BSTR;
		avarParams[0] = CtrlHandle;
		avarParams[0].vt = VT_I8;
		DISPPARAMS params = { avarParams, NULL, 4, 0 };
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

	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnClrControlCreated(Node, Ctrl, OLE2T(CtrlName), (HWND)CtrlHandle);
	return hr;
}

HRESULT CEclipseCtrl::Fire_TabChange(IWndNode* sender, LONG ActivePage, LONG OldPage)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[3];
		avarParams[2] = sender;
		avarParams[2].vt = VT_DISPATCH;
		avarParams[1] = ActivePage;
		avarParams[1].vt = VT_I4;
		avarParams[0] = OldPage;
		avarParams[0].vt = VT_I4;
		DISPPARAMS params = { avarParams, NULL, 3, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			IUnknown* punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();
			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection);
			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(9, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}

	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnTabChange(sender, ActivePage, OldPage);
	return hr;
}

HRESULT CEclipseCtrl::Fire_Event(IWndNode* sender, IDispatch* EventArg)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[2];
		avarParams[1] = sender;
		avarParams[1].vt = VT_DISPATCH;
		avarParams[0] = EventArg;
		avarParams[0].vt = VT_DISPATCH;
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
				hr = pConnection->Invoke(10, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}

	return hr;
}

HRESULT CEclipseCtrl::Fire_ControlNotify(IWndNode * sender, LONG NotifyCode, LONG CtrlID, LONGLONG CtrlHandle, BSTR CtrlClassName)
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
				hr = pConnection->Invoke(11, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnControlNotify(sender, NotifyCode, CtrlID, (HWND)CtrlHandle, OLE2T(CtrlClassName));
	return hr;
}

HRESULT CEclipseCtrl::Fire_TangramEvent(ITangramEventObj* pEventObj)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[1];
		avarParams[0] = pEventObj;
		avarParams[0].vt = VT_DISPATCH;
		DISPPARAMS params = { avarParams, NULL, 1, 0 };
		for (int iConnection = 0; iConnection < cConnections; iConnection++)
		{
			g_pTangram->Lock();
			IUnknown* punkConnection = m_vec.GetAt(iConnection);
			g_pTangram->Unlock();
			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection);
			if (pConnection)
			{
				CComVariant varResult;
				hr = pConnection->Invoke(12, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}
	if (m_pTangramWndPageProxy)
		m_pTangramWndPageProxy->OnTangramEvent(pEventObj);
	return hr;
}

STDMETHODIMP CEclipseCtrl::InitCtrl(BSTR bstrXml)
{
	CString strXml = OLE2T(bstrXml);
	if (strXml != _T(""))
	{
		CTangramXmlParse m_Parse;
		if (!m_Parse.LoadXml(strXml))
		{
			CString strPath = g_pTangram->m_strAppPath + strXml;
			if (m_Parse.LoadFile(strPath) == false)
				return S_OK;
		}
		int nCount = m_Parse.GetCount();
		CTangramXmlParse* pChild = nullptr;
		for (int i = 0; i < nCount; i++)
		{
			pChild = m_Parse.GetChild(i);
			CString strName = pChild->name();
			strName.MakeLower();
			auto it = m_mapTangramInfo.find(strName);
			if (it == m_mapTangramInfo.end())
			{
				m_mapTangramInfo[strName] = pChild->xml();
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CEclipseCtrl::put_TangramHandle(BSTR bstrHandleName, LONGLONG newVal)
{
	CString strName = OLE2T(bstrHandleName);
	HWND hWnd = (HWND)newVal;
	if (strName != _T("") && ::IsWindow(hWnd))
	{
		auto it = m_mapTangramHandle.find(strName);
		if (it != m_mapTangramHandle.end())
			m_mapTangramHandle.erase(strName);
		m_mapTangramHandle[strName] = hWnd;
	}

	return S_OK;
}

STDMETHODIMP CEclipseCtrl::Extend(BSTR bstrFrameName, BSTR bstrKey, BSTR bstrXml, IWndNode** ppNode)
{
	CString strFrameName = OLE2T(bstrFrameName);
	if (strFrameName == _T(""))
		strFrameName = _T("EclipseView");
	if (strFrameName.CompareNoCase(_T("TopView")) == 0)
	{
		return m_pEclipseWnd->Extend(bstrKey, bstrXml, ppNode);
	}
	auto it = m_mapTangramHandle.find(strFrameName);
	if (it == m_mapTangramHandle.end())
		return S_FALSE;
	HWND hWnd = it->second;
	CString strKey = OLE2T(bstrKey);
	if (strKey == _T(""))
		strKey= _T("default");
	CString strXml = OLE2T(bstrXml);
	IWndNode* pNode = nullptr;
	auto it2 = m_mapTangramFrame.find(strFrameName);
	if (it2 == m_mapTangramFrame.end())
	{
		HWND hTop = ::GetAncestor(m_hWnd, GA_ROOT);
		IWndPage* pPage = nullptr;
		g_pTangram->CreateWndPage((LONGLONG)hTop, &pPage);
		if (pPage == nullptr)
			return S_OK;
		//m_pPage = (CWndPage*)pPage;m_pEclipseWnd->m_
		IWndFrame* pFrame = nullptr;
		pPage->CreateFrame(CComVariant(0), CComVariant((long)hWnd), bstrFrameName, &pFrame);
		if (pFrame == nullptr)
		{
			return S_FALSE;
		}
		m_mapTangramFrame[strFrameName] = (CWndFrame*)pFrame;
		HRESULT hr = pFrame->Extend(bstrKey, bstrXml, ppNode);
		if (hr == S_OK&&strFrameName.CompareNoCase(_T("EclipseView")) == 0)
		{
			if (strFrameName.CompareNoCase(_T("EclipseView")) == 0)
			{
				((CWndFrame*)pFrame)->m_nWndFrameType = EclipseViewFrame;
			}
			else
				((CWndFrame*)pFrame)->m_nWndFrameType = EclipseSWTFrame;
			(*ppNode)->put_SaveToConfigFile(true);
		}
		return hr;
	}
	HRESULT hr = it2->second->Extend(bstrKey, bstrXml, ppNode);
	if(hr==S_OK&&strFrameName.CompareNoCase(_T("EclipseView")) == 0)
		(*ppNode)->put_SaveToConfigFile(true);
	return hr;
}
	
STDMETHODIMP CEclipseCtrl::ExtendEx(BSTR bstrFrameName, BSTR bstrKey, BSTR bstrXml, IWndNode** ppNode)
{
	HRESULT hr = Extend(bstrFrameName, bstrKey, bstrXml, ppNode);
	if (hr == S_OK&&*ppNode != nullptr)
		(*ppNode)->put_SaveToConfigFile(true);
	return S_OK;
}

void CEclipseCtrl::OnFinalMessage(HWND hWnd)
{
	::VariantClear(&m_varTag);
	if (m_pEclipseWnd)
	{
		auto it = m_pEclipseWnd->m_mapCtrl.find(hWnd);
		if (it != m_pEclipseWnd->m_mapCtrl.end())
			m_pEclipseWnd->m_mapCtrl.erase(it);
	}
	__super::OnFinalMessage(hWnd);
	if (g_pTangram->m_bEclipse)
		Release();
}

STDMETHODIMP CEclipseCtrl::get_EclipseViewHandle(LONGLONG* pVal)
{
	*pVal = (LONGLONG)m_hEclipseViewWnd;
	return S_OK;
}

LRESULT CEclipseCtrl::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	LRESULT lr = DefWindowProc(uMsg, wParam, lParam);
	HWND hTop = ::GetAncestor(m_hWnd, GA_ROOT);
	HWND hClient = ::GetWindow(hTop, GW_CHILD);
	m_mapTangramHandle[_T("TopView")] = hClient;
	m_hEclipseViewWnd = ::GetParent(::GetParent(::GetParent(m_hWnd)));
	m_mapTangramHandle[_T("EclipseView")] = m_hEclipseViewWnd;
	::PostMessage(m_hWnd, WM_TANGRAMMSG, 0, 0);
	return lr;
}

LRESULT CEclipseCtrl::OnTangramMsg(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
{
	LRESULT lr = DefWindowProc(uMsg, wParam, lParam);
	HWND hTop = ::GetAncestor(m_hWnd, GA_ROOT);

	auto it = g_pTangram->m_mapWorkBenchWnd.find(hTop);
	if (it != g_pTangram->m_mapWorkBenchWnd.end())
	{
		m_pEclipseWnd = it->second;
		auto it2 = m_pEclipseWnd->m_mapCtrl.find(m_hWnd);
		if (it2 == m_pEclipseWnd->m_mapCtrl.end())
			m_pEclipseWnd->m_mapCtrl[m_hWnd] = this;
		if (m_pEclipseWnd->m_pPage)
		{
			auto it = m_pEclipseWnd->m_pPage->m_mapNotifyCtrl.find(m_hWnd);
			if (it == m_pEclipseWnd->m_pPage->m_mapNotifyCtrl.end())
				m_pEclipseWnd->m_pPage->m_mapNotifyCtrl[m_hWnd] = this;
		}
		m_pTangramApplicationImpl->OnTangramCtrlCreated(this);
	}
	return lr;
}

STDMETHODIMP CEclipseCtrl::get_HWND(LONGLONG* pVal)
{
	*pVal = (LONGLONG)m_hWnd;
	return S_OK;
}

STDMETHODIMP CEclipseCtrl::get_Tangram(ITangram** pVal)
{
	*pVal = g_pTangram;
	return S_OK;
}

STDMETHODIMP CEclipseCtrl::get_WndPage(IWndPage** pVal)
{
	if (m_pEclipseWnd->m_pPage)
		*pVal = m_pEclipseWnd->m_pPage;
	return S_OK;
}

STDMETHODIMP CEclipseCtrl::get_WorkBenchWindow(IWorkBenchWindow** pVal)
{
	if (m_pEclipseWnd)
		*pVal = m_pEclipseWnd;
	return S_OK;
}

STDMETHODIMP CEclipseCtrl::AttachCLRObjEvent(IDispatch* Sender, WindowEventType nEventType, VARIANT_BOOL bAttachEvent)
{
	if (g_pTangram->m_pCLRProxy)
	{
		g_pTangram->m_pCLRProxy->AttachCLRObjEvent(Sender, nEventType, m_hWnd, bAttachEvent);
	}
	return S_OK;
}

STDMETHODIMP CEclipseCtrl::get_TopPage(IWndPage** pVal)
{
	if (m_pEclipseWnd)
	{
		*pVal = m_pEclipseWnd->m_pPage;
	}

	return S_OK;
}

STDMETHODIMP CEclipseCtrl::get_TopFrame(IWndFrame** pVal)
{
	if (m_pEclipseWnd)
	{
		*pVal = m_pEclipseWnd->m_pFrame;
	}

	return S_OK;
}

STDMETHODIMP CEclipseCtrl::get_ActiveTopNode(IWndNode** pVal)
{
	if (m_pEclipseWnd)
	{
		*pVal = m_pEclipseWnd->m_pFrame->m_pWorkNode;
	}

	return S_OK;
}

HRESULT CEclipseCtrl::FinalConstruct()
{
	CString strKey = _T("startdata");
	auto it = g_pTangram->m_mapValInfo.find(strKey);
	if (it != g_pTangram->m_mapValInfo.end())
	{
		CString strData = OLE2T(it->second.bstrVal);
		int nPos = strData.Find(_T("|"));
		if (nPos != -1)
		{
			g_pTangram->m_strStartView = strData.Left(nPos);
			g_pTangram->m_strStartXml = strData.Mid(nPos + 1);
			if (::IsWindow(g_pTangram->m_hHostWnd))
			{
				::PostMessage(g_pTangram->m_hHostWnd, WM_TANGRAMAPPQUIT, 0, 0);
			}
		}
		::VariantClear(&it->second);
		g_pTangram->m_mapValInfo.erase(it);
	}

	return S_OK;
}
