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
********************************************************************************/

// TangramObject.cpp : Implementation of CWndPage

#include "stdafx.h"
#include "WndNode.h"
#include "WndFrame.h"
#include "TangramCtrl.h"
#include "TangramApp.h"
#include "TangramCore.h"

CTangramNodeCommonData::CTangramNodeCommonData()
{
	m_pOldFrame			= nullptr;
	m_pOfficeObj		= nullptr;
	m_pTangramParse		= nullptr;
	m_pHostClientView	= nullptr;
#ifdef _DEBUG
	g_pTangram->m_nTangramNodeCommonData++;
#endif	
};

CTangramNodeCommonData::~CTangramNodeCommonData()
{
	if (m_pTangramParse)
		delete m_pTangramParse;
#ifdef _DEBUG
	g_pTangram->m_nTangramNodeCommonData--;
#endif	
};

// CWndPage

CWndPage::CWndPage()
{
	m_hWnd								= 0;
	m_nWebViewCount						= 0;
	m_nThreadId							= GetCurrentThreadId();
	m_bIsBlank							= false;
	m_bPageDataLoaded					= false;
	m_bIsDestory						= false;
	m_bDocComplete						= false;
	m_bDoc								= false;
	m_bEnableSinkCLRCtrlCreatedEvent	= true;
	m_strURL							= _T("");
	m_strXmlHeader						= _T("");
	m_strPageFileName					= _T("");
	m_strPageFilePath					= _T("");
	m_strConfigFileNodeName				= _T("tangramdefaultpage");
	m_pCP								= nullptr;
	m_pExternalDisp						= nullptr;
	m_pHtmlDocument2					= nullptr;
	m_pBKFrame							= nullptr;
	m_pActiveDoc						= nullptr;
	g_pTangram->m_pPage					= this;
	m_pTangramAppProxy					= nullptr;
	m_pTangramDocTemplate				= nullptr;

#ifdef _DEBUG
	g_pTangram->m_nTangram++;
#endif	
}

CWndPage::~CWndPage()
{
	if (this == g_pTangram->m_pDesignerWndPage)
		g_pTangram->m_pDesignerWndPage = nullptr;
#ifdef _DEBUG
	g_pTangram->m_nTangram--;
#endif	

	UnLoad();

	for (auto it2 : m_mapExternalObj)
	{
		it2.second->Release();
	}
	m_mapExternalObj.erase(m_mapExternalObj.begin(), m_mapExternalObj.end());

	m_mapFrame.erase(m_mapFrame.begin(), m_mapFrame.end());
	m_mapNode.erase(m_mapNode.begin(), m_mapNode.end());
	auto it = g_pTangram->m_mapWindowPage.find(m_hWnd);
	if (it != g_pTangram->m_mapWindowPage.end())
	{
		g_pTangram->m_mapWindowPage.erase(it);
	}
	if (g_pTangram->m_mapWindowPage.size() == 0)
		g_pTangram->Close();
	else
	{
		if (g_pTangram->m_mapWindowPage.size() == 1 && g_pTangram->m_pDesignerWndPage)
		{
			if (g_pTangram->m_bDeleteWndPage == FALSE)
				::DestroyWindow(g_pTangram->m_hHostWnd);
			else
				g_pTangram->m_bDeleteWndPage = FALSE;
		}
	}
	for (auto it : m_mapWndPageProxy)
	{
		if (it.second->m_bAutoDelete)
			delete it.second;
	}
	m_mapWndPageProxy.clear();
	for (auto it : m_mapEventObj)
		delete it.second;
	m_mapEventObj.clear();
}

STDMETHODIMP CWndPage::get_Count(long* pCount)
{
	*pCount = (long)m_mapFrame.size();
	return S_OK;
}

STDMETHODIMP CWndPage::get_Frame(VARIANT vIndex, IWndFrame **ppFrame)
{
	if (vIndex.vt == VT_I4)
	{
		long lCount = m_mapFrame.size();
		int iIndex = vIndex.lVal;
		HWND hWnd = (HWND)iIndex;
		if (::IsWindow(hWnd))
		{
			auto it = m_mapFrame.find(hWnd);
			if (it != m_mapFrame.end())
			{
				*ppFrame = it->second;
				return S_OK;
			}
		}
		if (iIndex < 0 || iIndex >= lCount) return E_INVALIDARG;

		auto it = m_mapFrame.begin();
		int iPos = 0;
		while (it != m_mapFrame.end())
		{
			if (iPos == iIndex)
			{
				*ppFrame = it->second;
				return S_OK;
			};
			iPos++;
			it++;
		}
		return S_OK;
	}

	if (vIndex.vt == VT_BSTR)
	{
		CString strKey = OLE2T(vIndex.bstrVal);
		auto it = m_mapWnd.find(strKey);
		if (it != m_mapWnd.end())
		{
			auto it2 = m_mapFrame.find(it->second);
			if (it2 != m_mapFrame.end())
			{
				*ppFrame = it2->second;
				return S_OK;
			}
		}
		return E_INVALIDARG;
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get__NewEnum(IUnknown** ppVal)
{
	typedef CComEnum<IEnumVARIANT, &IID_IEnumVARIANT, VARIANT, _Copy<VARIANT>>
		CComEnumVariant;

	CComObject<CComEnumVariant> *pe = 0;
	HRESULT hr = pe->CreateInstance(&pe);

	if (SUCCEEDED(hr))
	{
		pe->AddRef();
		long nLen = (long)m_mapFrame.size();
		VARIANT* rgvar = new VARIANT[nLen];
		ZeroMemory(rgvar, sizeof(VARIANT)*nLen);
		VARIANT* pItem = rgvar;
		for (auto it : m_mapFrame)
		{
			IUnknown* pDisp = nullptr;
			CWndFrame* pObj = it.second;
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

STDMETHODIMP CWndPage::get_External(IDispatch** ppVal)
{
	if (m_pExternalDisp)
	{
		*ppVal = m_pExternalDisp;
		(*ppVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndPage::put_External(IDispatch* newVal)
{
	if (m_pExternalDisp)
	{
		UnSinkObject(m_pExternalDisp, nullptr);
		m_pExternalDisp->Release();
	}
	if (newVal == nullptr)
	{
		m_pExternalDisp = nullptr;
		return S_OK;
	}
	m_pExternalDisp = newVal;
	m_pExternalDisp->AddRef();
	return S_OK;
}

STDMETHODIMP CWndPage::CreateFrame(VARIANT ParentObj, VARIANT HostWnd, BSTR bstrFrameName, IWndFrame** pRetFrame)
{
	HWND h = 0; 
	CString strFrameName = OLE2T(bstrFrameName);
	//auto it = m_mapWnd.find(strFrameName);
	//if (it != m_mapWnd.end())
	//{
	//	int i = 0;
	//	CString s = _T("");
	//	s.Format(_T("%s%d"), strFrameName,i);
	//	auto it = m_mapWnd.find(s);
	//	while (it != m_mapWnd.end())
	//	{
	//		i++;
	//		s.Format(_T("%s%d"), strFrameName, i);
	//		it = m_mapWnd.find(s);
	//		if (it == m_mapWnd.end())
	//			break;
	//	}
	//	strFrameName = s;
	//}
	BSTR bstrName = strFrameName.AllocSysString();
	if (ParentObj.vt == VT_DISPATCH&&HostWnd.vt == VT_BSTR)
	{
		if (g_pTangram->m_pCLRProxy)
		{
			IDispatch* pDisp = nullptr;
			pDisp = g_pTangram->m_pCLRProxy->GetCtrlByName(ParentObj.pdispVal, HostWnd.bstrVal, true);
			if (pDisp)
			{
				h = g_pTangram->m_pCLRProxy->GetCtrlHandle(pDisp);
				if (h)
				{
					HRESULT hr = CreateFrame(CComVariant(0), CComVariant((long)h), bstrName, pRetFrame);
					::SysFreeString(bstrName);
					return hr;
				}
			}
		}
		::SysFreeString(bstrName);
		return S_FALSE;
	}
	if (HostWnd.vt == VT_DISPATCH)
	{
		if (g_pTangram->m_pCLRProxy)
		{
			h = g_pTangram->m_pCLRProxy->IsCtrlCanNavigate(HostWnd.pdispVal);
			if (h)
			{
				CString strName = strFrameName;
				if (strName == _T(""))
				{
					::SysFreeString(bstrName);
					bstrFrameName = g_pTangram->m_pCLRProxy->GetCtrlName(HostWnd.pdispVal);
					strName = OLE2T(bstrFrameName);
					if (strName == _T(""))
						bstrFrameName = CComBSTR(L"Default");
					
					strFrameName = OLE2T(bstrFrameName);
					auto it = m_mapWnd.find(strFrameName);
					if (it != m_mapWnd.end())
					{
						int i = 0;
						CString s = _T("");
						s.Format(_T("%s%d"), strFrameName, i);
						auto it = m_mapWnd.find(s);
						while (it != m_mapWnd.end())
						{
							i++;
							s.Format(_T("%s%d"), strFrameName, i);
							it = m_mapWnd.find(s);
							if (it == m_mapWnd.end())
								break;
						}
						strFrameName = s;
					}

				}
				auto it = m_mapFrame.find((HWND)h);
				if (it == m_mapFrame.end())
					return CreateFrame(CComVariant(0), CComVariant((long)h), CComBSTR(strFrameName), pRetFrame);
				else
					*pRetFrame = it->second;
			}
		}
	}
	else if (HostWnd.vt == VT_I2||HostWnd.vt == VT_I4 || HostWnd.vt == VT_I8)
	{
		BOOL bIsMDI = FALSE;
		HWND _hWnd = NULL;
		if(HostWnd.vt == VT_I4)
			_hWnd = (HWND)HostWnd.lVal;
		if(HostWnd.vt == VT_I8)
			_hWnd = (HWND)HostWnd.llVal;
		if (_hWnd == 0)
		{
			_hWnd = ::FindWindowEx(m_hWnd, NULL, _T("MDIClient"), NULL);
			if (_hWnd == NULL)
				_hWnd = ::GetWindow(m_hWnd, GW_CHILD);
			else
				bIsMDI = TRUE;
		}
		if (_hWnd&&::IsWindow(_hWnd))
		{
			auto it = m_mapFrame.find(_hWnd);
			if (it == m_mapFrame.end())
			{
				DWORD dwID = ::GetWindowThreadProcessId(_hWnd, NULL);
				TangramThreadInfo* pThreadInfo = g_pTangram->GetThreadInfo(dwID);

				CWndFrame* m_pExtenderFrame = new CComObject<CWndFrame>();
				CString strName = strFrameName;
				if (strName == _T(""))
					strName = _T("default");
				m_pExtenderFrame->m_strFrameName = strName;
				if (strName.CompareNoCase(_T("System.Windows.Forms.MdiClient")) == 0)
					m_pExtenderFrame->m_nWndFrameType = WinFormMDIFrame;
				else if(bIsMDI)
					m_pExtenderFrame->m_nWndFrameType = MDIFrame;
				::GetClassName(::GetParent(_hWnd), g_pTangram->m_szBuffer, MAX_PATH);
				CString strClassName = CString(g_pTangram->m_szBuffer);
				if (strClassName.Find(_T("Afx:ControlBar:")) == 0)
				{
					if (g_pTangram->m_pMDIMainWnd)
					{
						auto it = g_pTangram->m_pMDIMainWnd->m_mapDesignableWnd.find(_hWnd);
						if (it == g_pTangram->m_pMDIMainWnd->m_mapDesignableWnd.end())
						{
							g_pTangram->m_pMDIMainWnd->m_mapDesignableWnd[_hWnd] = strName;
						}
					}
					m_pExtenderFrame->m_nWndFrameType = ControlBarFrame;
				}
				else if (strClassName.Find(_T("MDIClient")) == 0)
				{
					m_pExtenderFrame->m_nWndFrameType = MDIFrame;
				}
				m_pExtenderFrame->m_pPage = this;
				m_pExtenderFrame->m_hHostWnd = _hWnd;
				if (ParentObj.vt == VT_I8 || ParentObj.vt == VT_I4)
				{
					HWND hPWnd = (HWND)ParentObj.llVal;
					if (::IsWindow(hPWnd))
					{
						m_pExtenderFrame->m_hPWnd = hPWnd;
					}
				}
				pThreadInfo->m_mapTangramFrame[_hWnd] = m_pExtenderFrame;
				m_mapFrame[_hWnd] = m_pExtenderFrame;
				m_mapWnd[strName] = _hWnd;

				for (auto it : g_pTangram->m_mapTangramAppProxy)
				{
					CTangramWndFrameProxy* pTangramWndFrameProxy = it.second->OnWndFrameCreated(m_pExtenderFrame);
					if (pTangramWndFrameProxy)
					{
						m_pExtenderFrame->m_mapWndFrameProxy[it.second] = pTangramWndFrameProxy;
					}
				}

				if (strName.CompareNoCase(_T("MdiClient"))==0)
				{
					//m_pExtenderFrame->CreateBKWnd();
				}
				*pRetFrame = m_pExtenderFrame;
			}
			else
				*pRetFrame = it->second;
		}
	}
		
	return S_OK;
}


STDMETHODIMP CWndPage::GetFrameFromCtrl(IDispatch* CtrlObj, IWndFrame** ppFrame)
{
	if (g_pTangram->m_pCLRProxy)
	{
		HWND h = g_pTangram->m_pCLRProxy->IsCtrlCanNavigate(CtrlObj);
		if (h)
		{
			auto it = m_mapFrame.find(h);
			if (it != m_mapFrame.end())
				*ppFrame = it->second;
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::GetWndNode(BSTR bstrFrameName, BSTR bstrNodeName, IWndNode** pRetNode)
{
	CString strKey = OLE2T(bstrFrameName);
	CString strName = OLE2T(bstrNodeName);
	if (strKey == _T("") || strName == _T(""))
		return S_FALSE;
	auto it = m_mapWnd.find(strKey);
	if (it != m_mapWnd.end())
	{
		HWND hWnd = it->second;
		if (::IsWindow(hWnd))
		{
			auto it = m_mapFrame.find(hWnd);
			if (it != m_mapFrame.end())
			{
				CWndFrame* pFrame = it->second;
				strName = strName.MakeLower();
				auto it2 = pFrame->m_mapNode.find(strName);
				if (it2 != pFrame->m_mapNode.end())
					*pRetNode = (IWndNode*)it2->second;
				else
				{
					it2 = m_mapNode.find(strName);
					if (it2 != m_mapNode.end())
						*pRetNode = (IWndNode*)it2->second;
				}
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_URL(BSTR* pVal)
{
	*pVal = m_strURL.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndPage::put_URL(BSTR newVal)
{
	m_strURL = OLE2T(newVal);
	CString s = _T("");
	s.Format(_T("%s?xx=%x"), m_strURL,this);
	UnLoad();
	Init(s);
	return S_OK;
}

STDMETHODIMP CWndPage::AddObjToHtml(BSTR strObjName, VARIANT_BOOL bConnectEvent, IDispatch* pObjDisp)
{
	CString strName = OLE2T(strObjName);
	if (strName!=_T(""))
	{
		CWebPageExtender::AddObject(strName, pObjDisp);
		if (bConnectEvent)
		{
			CComQIPtr<IConnectionPointContainer> pContainer(pObjDisp);
			if (pContainer)
			{
				strName += _T("_On");
				SinkObject(strName, pObjDisp);
			}
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_Extender(BSTR bstrExtenderName, IDispatch** pVal)
{
	CString strName = OLE2T(bstrExtenderName);
	if (strName == _T(""))
		return S_OK;
	auto it = m_mapExternalObj.find(strName);
	if (it != m_mapExternalObj.end())
	{
		*pVal = it->second;
		(*pVal)->AddRef();
	}
	return S_OK;
}


STDMETHODIMP CWndPage::put_Extender(BSTR bstrExtenderName, IDispatch* newVal)
{
	CString strName = OLE2T(bstrExtenderName);
	if (strName == _T(""))
		return S_OK;
	auto it = m_mapExternalObj.find(strName);
	if (it == m_mapExternalObj.end())
	{
		m_mapExternalObj[strName] = newVal;
		newVal->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_FrameName(LONGLONG hHwnd, BSTR* pVal)
{
	HWND _hWnd = (HWND)hHwnd;
	if (_hWnd)
	{
		auto it = m_mapFrame.find(_hWnd);
		if (it!=m_mapFrame.end())
			*pVal = it->second->m_strFrameName.AllocSysString();
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_WndFrame(BSTR bstrFrameName, IWndFrame** pVal)
{
	CString strName = OLE2T(bstrFrameName);
	if (strName != _T(""))
	{
		auto it = m_mapWnd.find(strName);
		if (it != m_mapWnd.end())
		{
			HWND h = it->second;
			auto it2 = m_mapFrame.find(h);
			if (it2 != m_mapFrame.end())
				*pVal = it2->second;
		}
	}
	return S_OK;
}

void CWndPage::OnNodeDocComplete(WPARAM wParam)
{
	bool bState = false;
	for (auto it1 : m_mapFrame)
	{
		for (auto it2 : it1.second->m_mapNode)
		{
			if (it2.second->m_bCreated == false)
			{
				::PostMessage(m_hWnd, WM_TANGRAM_WEBNODEDOCCOMPLETE, wParam, 0);
				return;
			}
		}
	}

	switch (wParam)
	{
	case 0:
		{
			int nIndex = 0;
			for (auto it : m_mapNode)
			{
				if (it.second->m_bNodeDocComplete==false)
				{
					it.second->m_bNodeDocComplete = true;
					if (it.second->m_pTangramNodeWBEvent)
					{
						BSTR bstrURL = it.second->m_strURL.AllocSysString();
						it.second->Fire_NodeDocumentComplete(it.second->m_pExtender, bstrURL);
						CComPtr<IHTMLWindow2> pWinDisp;
						it.second->m_pTangramNodeWBEvent->m_pHTMLDocument2->get_parentWindow(&pWinDisp);
						Fire_PageLoaded(pWinDisp.p, bstrURL);
						::SysFreeString(bstrURL);
					}
					nIndex++;
				}
			}
			if (nIndex < m_nWebViewCount)
				::PostMessage(m_hWnd, WM_TANGRAM_WEBNODEDOCCOMPLETE, wParam, 0);
		}
		break;
	case 1:
	default:
		if (m_bDocComplete == false)
		{
			m_bDocComplete = true;
			CComQIPtr<IHTMLDocument3> pDoc(m_pHtmlDocument2);
			CComPtr<IHTMLElementCollection> pColl;
			HRESULT hr = pDoc->getElementsByTagName(CComBSTR(L"xml"), &pColl);
			if (hr == S_OK)
			{
				long nCount = 0;
				pColl->get_length(&nCount);
				if (nCount)
				{
					CComBSTR bstrID(L"");
					BSTR bstrXml=::SysAllocString(L"");
					CString strID = _T("");
					CString strXml = _T("");
					int nIndex = 0;
					for (int i = 0; i < nCount; i++)
					{
						CComVariant vIdx((long)i, VT_I4);
						CComPtr<IDispatch> pDisp;
						hr = pColl->item(vIdx, vIdx, &pDisp);
						if (pDisp)
						{
							CComQIPtr<IHTMLElement> pElem(pDisp);
							pElem->get_innerHTML(&bstrXml);
							pElem->get_id(&bstrID);
							strID = OLE2T(bstrID);
							if (strID == _T("tangram"))
							{
								g_pTangram->Encode(bstrXml, false, &bstrXml);
								strXml = OLE2T(bstrXml);
								::SysFreeString(bstrXml);
								CTangramXmlParse m_Parse;
								if (m_Parse.LoadXml(strXml))
								{
									int nCount = m_Parse.GetCount();
									for (int i = 0; i < nCount; i++)
									{
										CString strName = m_Parse[i].name();
										CString _strXml = m_Parse[i].xml();
										if (strName.CompareNoCase(_T("xtmlSelectedText"))==0)
										{
											g_pTangram->m_strNodeSelectedText = m_Parse[i].text();
										}
										else if (strName.CompareNoCase( _T("tangramdesignertoolbar"))==0)
										{
											g_pTangram->m_strDesignerXml = _strXml;
										}
										else
											m_mapXtml[strName] = _strXml;
									}
								}
							}
							else
							{
								strXml = OLE2T(bstrXml);
								if (strID == _T(""))
								{
									strID.Format(_T("xtml%d"), nIndex);
									nIndex++;
								}
								if (strID == _T("xtmlSelectedText"))
								{
									g_pTangram->m_strNodeSelectedText = strXml;
								}
								else
									m_mapXtml[strID] = strXml;
							}
						}
					}
				}
				else
				{
					//CString strPath = g_pTangram->m_strAppDataPath + theApp.m_pszExeName+_T("\\tangraminit.xml");
					//if (::PathFileExists(strPath))
					//{
					//	CTangramXmlParse m_Parse;
					//	if (m_Parse.LoadFile(strPath))
					//	{
					//		int nCount = m_Parse.GetCount();
					//		for (int i = 0; i < nCount; i++)
					//		{
					//			CTangramXmlParse* pParse = m_Parse.GetChild(i);
					//			CString strID = pParse->attr(_T("id"), _T(""));
					//			CString strXml = pParse->GetChild(0)->xml();
					//			if (strID == _T("xmlRibbon"))
					//			{
					//				CString strPath = g_pTangram->m_strAppCommonDocPath + _T("OfficeRibbon\\") + g_pTangram->m_strExeName + _T("\\ribbon.xml");
					//				CTangramXmlParse m_Parse2;
					//				if (m_Parse2.LoadXml(strXml))
					//					m_Parse2.SaveFile(strPath);
					//			}
					//			if (strID == _T("tangramdesigner"))
					//				g_pTangram->m_strDesignerXml = strXml;
					//			else
					//			{
					//				strID.MakeLower();
					//				if (strID == _T("newtangramdocument"))
					//				{
					//					g_pTangram->m_strNewDocXml = strXml;
					//				}
					//				else
					//				{
					//					g_pTangram->m_mapValInfo[strID] = CComVariant(strXml);
					//				}
					//			}
					//		}
					//	}
					//}
				}
				if (g_pTangram->m_strDesignerXml != _T(""))
				{
					if (::GetModuleHandle(_T("TangramPackage.dll")) == nullptr/*&&g_pTangram->m_bOfficeApp == false*/)
					{
						auto it = m_mapFrame.find(g_pTangram->m_hChildHostWnd);
						if (it == m_mapFrame.end())
						{
							IWndFrame* pFrame = nullptr;
							HRESULT hr = CreateFrame(CComVariant(0), CComVariant((long)g_pTangram->m_hChildHostWnd), CComBSTR(L"DeignerTool"), &pFrame);
							if (pFrame)
							{
								IWndNode* pNode = nullptr;
								pFrame->Extend(CComBSTR(L"DeignerToolBox"), g_pTangram->m_strDesignerXml.AllocSysString(), &pNode);
								CWndFrame* _pFrame = (CWndFrame*)pFrame;
								_pFrame->m_bDesignerState = false;
							}
						}
					}
					else
					{
						CWebPageExtender::AddObject(_T("WndPage"), (IWndPage*)this);
						SinkObject(_T("WndPage_On"), (IWndPage*)this);
						CComPtr<IHTMLWindow2> pWinDisp;
						m_pHtmlDocument2->get_parentWindow(&pWinDisp);
						Fire_PageLoaded(pWinDisp.p, CComBSTR(m_strURL));
					}
				}
				else
				{
					CWebPageExtender::AddObject(_T("WndPage"), (IWndPage*)this);
					SinkObject(_T("WndPage_On"), (IWndPage*)this);
					CComPtr<IHTMLWindow2> pWinDisp;
					m_pHtmlDocument2->get_parentWindow(&pWinDisp);
					Fire_PageLoaded(pWinDisp.p, CComBSTR(m_strURL));
				}
			}
			else
			{
				UnLoad();
			}
		}
		break;
	}
}

void CWndPage::UpdateMapKey(CString strXml)
{
	if (m_strXmlHeader != _T(""))
		return;
	CTangramXmlParse m_Parse;
	if (m_Parse.LoadXml(strXml)||m_Parse.LoadFile(strXml))
	{
		strXml = m_Parse.xml();
		int nPos = strXml.Find(_T(">"));
		m_strXmlHeader = strXml.Left(nPos + 1);
		nPos = strXml.ReverseFind('<');
		if (nPos != -1)
			m_strXmlBottom = strXml.Mid(nPos);
		CString strMainKey = _T("tangramdefaultpage");// m_Parse.name();
		int nCount = m_Parse.GetCount();
		for (int i = 0; i < nCount; i++)
		{
			CTangramXmlParse* pChild = m_Parse.GetChild(i);
			CString strFrameName = pChild->name();
			int nCount = pChild->GetCount();
			for (int i = 0; i < nCount; i++)
			{
				CTangramXmlParse* _pChild = pChild->GetChild(i);
				CString strKey = _pChild->name();
				if (strKey.CompareNoCase(_T("tangram")) == 0)
					strKey = _T("default");
				CString _strKey = strKey + _T("@") + strFrameName + _T("@") + strMainKey;
				m_strMapKey[_strKey] = _pChild->xml();
			}
		}

		if (m_bPageDataLoaded == false)
		{
			m_bPageDataLoaded = true;
			CTangramXmlParse* pParse = (CTangramXmlParse*)m_Parse.GetChild(_T("pagedata"));
			if (pParse&&::PathFileExists(g_pTangram->m_strTemplatePath) == TRUE)
			{
				int nCount = pParse->GetCount();
				if (nCount)
				{
					for (int i = 0; i < nCount; i++)
					{
						CString strPath = g_pTangram->m_strTemplatePath;
						CTangramXmlParse* pChild = pParse->GetChild(i);
						CString strName = pChild->name();
						m_mapXtml[strName] = pChild->xml();
						int nPos = strPath.ReverseFind('\\');
						CString strFileName = strPath.Mid(nPos + 1);
						int nPos1 = strFileName.ReverseFind('.');
						strFileName = strFileName.Left(nPos1);
						strPath = strPath.Left(nPos + 1) + strFileName + _T("_") + strName + _T(".bmp");
						if (::PathFileExists(strPath))
						{
							m_mapXtml[strName + _T("_imagePath")] = strPath;
							CString strData = g_pTangram->EncodeFileToBase64(strPath);
							m_mapXtml[strName + _T("_image")] = strData;
						}
					}
				}
			}
		}
	}
}

void CWndPage::BeforeDestory()
{
	if (!m_bIsDestory)
	{
		m_bIsDestory = true;
		Fire_Destroy();
		for (auto it : m_mapNode)
		{
			if (it.second->m_pTangramNodeWBEvent)
			{
				it.second->m_pTangramNodeWBEvent->ConnectJSEngine(nullptr);
				it.second->m_pTangramNodeWBEvent->DispEventUnadvise(it.second->m_pDisp);
				delete it.second->m_pTangramNodeWBEvent;
				it.second->m_pTangramNodeWBEvent = nullptr;
			}
		}

		for (auto it: m_mapFrame)
			it.second->Destroy();
		if(m_pJSEngine)
			ConnectJSEngine(nullptr);
		if (/*g_pTangram->m_bChromeShutdown==false&&*/g_pTangram->m_pCLRProxy)
		{
			g_pTangram->m_pCLRProxy->ReleaseTangramObj((IWndPage*)this);
		}
	}
}

STDMETHODIMP CWndPage::get_Node(BSTR bstrNodeName, IWndNode** pVal)
{
	CString strName = OLE2T(bstrNodeName);
	if (strName == _T(""))
		return S_OK;
	auto it2 = m_mapNode.find(strName);
	if (it2 == m_mapNode.end())
		return S_OK;

	if (it2->second)
		*pVal = it2->second;

	return S_OK;
}

STDMETHODIMP CWndPage::get_XObject(BSTR bstrName, IDispatch** pVal)
{
	CString strName = OLE2T(bstrName);
	if (strName == _T(""))
		return S_OK;
	auto it2 = m_mapNode.find(strName);
	if (it2 == m_mapNode.end())
		return S_OK;
	if (it2->second->m_pDisp)
	{
		*pVal = it2->second->m_pDisp;
		(*pVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndPage::AttachEvent(BSTR bstrName, IDispatch* pWndDisp, VARIANT_BOOL* bResult)
{
	CString strNames = OLE2T(bstrName);
	if (strNames==_T(""))
		return S_FALSE;
	CComQIPtr<IHTMLWindow2> pWindow2(pWndDisp);
	if (pWindow2 == nullptr)
		return S_FALSE;

	CWebPageExtender* _pJSExtender = nullptr;
	for (auto it : m_mapWebPageExtender)
	{
		CComPtr<IHTMLWindow2> pWindow;
		it.first->get_parentWindow(&pWindow);
		if (pWindow.p == pWindow2.p)
		{
			_pJSExtender = it.second;
			break;
		}
	}
	if (_pJSExtender == nullptr)
		return S_FALSE;

	int nPos = strNames.Find(_T(","));
	if (nPos == -1)
	{
		auto it = m_mapNode.find(strNames);
		if (it == m_mapNode.end())
		{
			return S_OK;
		}
		strNames += _T("_On");
		_pJSExtender->SinkObject(strNames, it->second->m_pDisp);
		*bResult = true;
	}
	while (nPos != -1)
	{
		CString strName = strNames.Left(nPos);
		strNames = strNames.Mid(nPos + 1);
		nPos = strNames.Find(_T(","));
		auto it = m_mapNode.find(strName);
		if (it != m_mapNode.end() && it->second->m_pDisp)
		{

			strName += _T("_On");
			_pJSExtender->SinkObject(strName, it->second->m_pDisp);
			*bResult = true;
		}
		if (nPos == -1)
		{
			it = m_mapNode.find(strNames);
			if (it != m_mapNode.end())
			{
				strNames += _T("_On");
				_pJSExtender->SinkObject(strNames, it->second->m_pDisp);
				*bResult = true;
				return S_OK;
			}
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_Width(long* pVal)
{
	if (m_hWnd)
	{
		RECT rc;
		::GetWindowRect(m_hWnd, &rc);
		*pVal = rc.right - rc.left;
	}

	return S_OK;
}

STDMETHODIMP CWndPage::put_Width(long newVal)
{
	if (m_hWnd&&newVal)
	{
		RECT rc;
		::GetWindowRect(m_hWnd, &rc);
		rc.right = rc.left + newVal;
		::SetWindowPos(m_hWnd, NULL, rc.left, rc.top, newVal, rc.bottom - rc.top, SWP_NOACTIVATE | SWP_NOREDRAW);
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_Height(long* pVal)
{
	if (m_hWnd)
	{
		RECT rc;
		::GetWindowRect(m_hWnd, &rc);
		*pVal = rc.bottom - rc.top;
	}
	return S_OK;
}

STDMETHODIMP CWndPage::put_Height(long newVal)
{
	if (m_hWnd&&newVal)
	{
		RECT rc;
		::GetWindowRect(m_hWnd, &rc);
		rc.right = rc.left + newVal;
		::SetWindowPos(m_hWnd, NULL, rc.left, rc.top, rc.right - rc.left, newVal, SWP_NOACTIVATE | SWP_NOREDRAW);
	}
	return S_OK; 
}

STDMETHODIMP CWndPage::AddNodesToPage(IDispatch* pHtmlDoc, BSTR bstrNodeNames, VARIANT_BOOL bAdd, VARIANT_BOOL* bSuccess)
{
	CString strNames = OLE2T(bstrNodeNames);
	if (strNames != _T(""))
	{
		int nPos = strNames.Find(_T(","));
		if (nPos == -1)
		{
			auto it = m_mapNode.find(strNames);
			if (it != m_mapNode.end() && it->second->m_pDisp)
			{
				AddObject(pHtmlDoc, it->second->m_pDisp, bstrNodeNames, false, bSuccess);
			}
			return S_OK;
		}
		CString strName = _T("");
		while (nPos != -1)
		{
			strName = strNames.Left(nPos);
			strNames = strNames.Mid(nPos + 1);
			nPos = strNames.Find(_T(","));
			auto it = m_mapNode.find(strName);
			if (it != m_mapNode.end() && it->second->m_pDisp)
			{
				AddObject(pHtmlDoc, it->second->m_pDisp, strName.AllocSysString(), false, bSuccess);
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_NodeNames(BSTR* pVal)
{
	CString strNames = _T("");
	for (auto it : m_mapNode)
	{
		strNames += it.first;
		strNames += _T(",");
	}
	int nPos = strNames.ReverseFind(',');
	strNames = strNames.Left(nPos);
	*pVal = strNames.AllocSysString();
	return S_OK;
}

STDMETHODIMP CWndPage::AddObject(IDispatch* pHtmlDoc, IDispatch* pObject, BSTR bstrObjName, VARIANT_BOOL bSinkEvent, VARIANT_BOOL* bResult)
{
	CString strName = OLE2T(bstrObjName);
	CComQIPtr<IHTMLDocument2> pDoc(pHtmlDoc);
	if (strName != _T("") && pHtmlDoc)
	{
		auto it = m_mapWebPageExtender.find(pDoc.p);
		if (it != m_mapWebPageExtender.end())
		{
			CWebPageExtender* _pJSExtender = it->second;
			if (_pJSExtender)
			{
				_pJSExtender->AddObject(strName, pObject);
				if (bSinkEvent)
				{
					CComQIPtr<IConnectionPointContainer> pContainer(pObject);
					if (pContainer)
					{
						strName += _T("_On");
						_pJSExtender->SinkObject(strName, pObject);
						*bResult = true;
					}
				}
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_HTMLWindow(BSTR NodeName, IDispatch** pVal)
{
	CString strName = OLE2T(NodeName);
	if (strName != _T(""))
	{
		auto it = m_mapNode.find(strName);
		if (it != m_mapNode.end() && it->second->m_pTangramNodeWBEvent->m_pHTMLDocument2)
		{
			CComPtr<IHTMLWindow2> pHTMLWindow2;
			it->second->m_pTangramNodeWBEvent->m_pHTMLDocument2->get_parentWindow(&pHTMLWindow2);
			*pVal = pHTMLWindow2.Detach();
			(*pVal)->AddRef();
		}
	}
	else if(m_pHtmlDocument2)
	{
		CComPtr<IHTMLWindow2> pHTMLWindow2;
		m_pHtmlDocument2->get_parentWindow(&pHTMLWindow2);
		*pVal = pHTMLWindow2.Detach();
		(*pVal)->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_HtmlDocument(IDispatch** pVal)
{
	if (m_pHtmlDocument2)
	{
		*pVal = m_pHtmlDocument2;
		(*pVal)->AddRef();
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_Parent(IWndPage** pVal)
{
	HWND hWnd = ::GetParent(m_hWnd);
	if (hWnd == NULL)
		return S_OK;

	auto it = g_pTangram->m_mapWindowPage.find(hWnd);
	while (it == g_pTangram->m_mapWindowPage.end())
	{
		hWnd = ::GetParent(hWnd);
		if (hWnd == NULL)
			return S_OK;
		it = g_pTangram->m_mapWindowPage.find(hWnd);
		if (it != g_pTangram->m_mapWindowPage.end())
		{
			*pVal = it->second;
			return S_OK;
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::GetHtmlDocument(IDispatch* HtmlWindow, IDispatch** ppHtmlDoc)
{
	CComQIPtr<IHTMLWindow2> pWindow2(HtmlWindow);
	if (pWindow2 == nullptr)
		return S_OK;

	for (auto it : m_mapWebPageExtender)
	{
		CComPtr<IHTMLWindow2> pWindow;
		it.first->get_parentWindow(&pWindow);
		if (pWindow.p == pWindow2.p)
		{
			*ppHtmlDoc = it.first;
			(*ppHtmlDoc)->AddRef();
			return S_OK;
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::AttachNodeEvent(BSTR bstrNames, IDispatch* pWndDisp)
{
	CString strNames = OLE2T(bstrNames);
	if (strNames == _T(""))
		return S_FALSE;
	CComQIPtr<IHTMLWindow2> pWindow2(pWndDisp);
	if (pWindow2 == NULL)
		return S_FALSE;

	CWebPageExtender* _pJSExtender = nullptr;
	for (auto it0 : m_mapWebPageExtender)
	{
		CComPtr<IHTMLWindow2> pWindow;
		it0.first->get_parentWindow(&pWindow);
		if (pWindow.p == pWindow2.p)
		{
			_pJSExtender = it0.second;
			break;
		}
	}
	if (_pJSExtender == nullptr)
		return S_FALSE;

	int nPos = strNames.Find(_T(","));
	if (nPos == -1)
	{
		auto it = m_mapNode.find(strNames);
		if (it == m_mapNode.end())
		{
			return S_OK;
		}
		strNames += _T("_Node_On");
		_pJSExtender->SinkObject(strNames, (IWndNode*)it->second);
	}
	CString strName = _T("");
	while (nPos != -1)
	{
		strName = strNames.Left(nPos);
		strNames = strNames.Mid(nPos + 1);
		nPos = strNames.Find(_T(","));
		auto it = m_mapNode.find(strName);
		if (it != m_mapNode.end() && it->second->m_pDisp)
		{
			strName += _T("_Node_On");
			_pJSExtender->SinkObject(strName, (IWndNode*)it->second);
		}
		if (nPos == -1)
		{
			it = m_mapNode.find(strNames);
			if (it != m_mapNode.end())
			{
				strNames += _T("_Node_On");
				_pJSExtender->SinkObject(strNames, (IWndNode*)it->second);
				return S_OK;
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::AttachObjEvent(IDispatch* HTMLWindow,IDispatch* SourceObj, BSTR bstrEventName,BSTR AliasName)
{
	if (g_pTangram->m_mapEventType.size() == 0)
	{
		g_pTangram->InitEventDic();
	}

	CString strAliasName = OLE2T(AliasName);
	CString strEventName = OLE2T(bstrEventName);
	if (strEventName == _T(""))
		return S_FALSE;
	auto it = g_pTangram->m_mapEventType.find(strEventName);
	if (it == g_pTangram->m_mapEventType.end())
		return S_FALSE;

	CComQIPtr<IHTMLWindow2> pWindow2(HTMLWindow);
	if (pWindow2 == NULL)
	{
		if (g_pTangram->m_pCLRProxy)
		{
			g_pTangram->m_pEventProxy = new CComObject < CEventProxy >;
			CComBSTR bstrName(L"");
			bstrName = g_pTangram->m_pCLRProxy->AttachObjEvent((IEventProxy*)g_pTangram->m_pEventProxy, SourceObj, it->second, HTMLWindow);
			if (strAliasName == _T(""))
				strAliasName = OLE2T(bstrName);
		}
		return S_OK;
	}

	for (auto it0 : m_mapWebPageExtender)
	{
		CComPtr<IHTMLWindow2> pWindow;
		it0.first->get_parentWindow(&pWindow);
		if (pWindow.p == pWindow2.p)
		{
			if (g_pTangram->m_pCLRProxy)
			{
				g_pTangram->m_pEventProxy = new CComObject < CEventProxy >;
				CComBSTR bstrName(L"");
				bstrName = g_pTangram->m_pCLRProxy->AttachObjEvent((IEventProxy*)g_pTangram->m_pEventProxy, SourceObj, it->second, HTMLWindow);
				if (strAliasName == _T(""))
					strAliasName = OLE2T(bstrName);
				return it0.second->SinkObject(strAliasName, g_pTangram->m_pEventProxy);
			}
			break;
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::AttachNodeSubCtrlEvent(IDispatch* HtmlWndDisp, VARIANT Node, VARIANT CtrlName, BSTR EventName, BSTR AliasName)
{
	if (g_pTangram->m_mapEventType.size() == 0)
	{
		g_pTangram->InitEventDic();
	}

	CString strEventName = OLE2T(EventName);
	if (strEventName == _T(""))
		return S_FALSE;

	auto it1 = g_pTangram->m_mapEventType.find(strEventName);
	if (it1 == g_pTangram->m_mapEventType.end())
		return S_FALSE;

	CString strName = _T("");
	CWndNode* pNode = nullptr;
	if (Node.vt == VT_BSTR)
	{
		strName = OLE2T(Node.bstrVal);
		if (strName == _T(""))
			return S_OK;
	}
	else if (Node.vt == VT_DISPATCH)
	{
		CComQIPtr<IWndNode> _pNode(Node.pdispVal);
		if (_pNode)
		{
			pNode = (CWndNode*)_pNode.p;
			if (pNode->m_nViewType != CLRCtrl)
				return S_FALSE;
			strName = pNode->m_strNodeName;
		}
		else
			return S_OK;
	}
	else
		return S_OK;
	if (strName == _T("") && pNode == nullptr)
		return S_OK;
	CString strCtrlName = _T("");
	CComPtr<IDispatch> pDisp;
	if (CtrlName.vt == VT_BSTR)
		strCtrlName = OLE2T(CtrlName.bstrVal);
	else if (CtrlName.vt == VT_DISPATCH)
		pDisp.Attach(CtrlName.pdispVal);
	if (strCtrlName == _T("") && pDisp == nullptr)
		return S_FALSE;

	if (pNode == nullptr)
	{
		auto it2 = m_mapNode.find(strName);
		if (it2 == m_mapNode.end())
			return S_OK;
		pNode = it2->second;
	}
	if (pNode)
	{
		if (g_pTangram->m_pCLRProxy)
		{
			if (pDisp)
			{
				CComBSTR bstrName(L"");
				bstrName = g_pTangram->m_pCLRProxy->GetCtrlName(pDisp);
				strCtrlName = OLE2T(bstrName);
			}
		}
		else
			return S_OK;

		if (pDisp == nullptr)
			pNode->GetCtrlByName(strCtrlName.AllocSysString(), true, &pDisp);

		if (pDisp)
		{
			CString strAliasName = OLE2T(AliasName);
			if (strAliasName == _T(""))
			{
				strAliasName = strName;
				strAliasName += _T("_");
				strAliasName += strCtrlName;
				strAliasName += _T("_");
				strAliasName += strEventName;
			}

			CWebPageExtender* _pJSExtender = nullptr;
			CComQIPtr<IHTMLWindow2> pWindow2(HtmlWndDisp);
			if (pWindow2)
			{
				for (auto it0 : m_mapWebPageExtender)
				{
					CComPtr<IHTMLWindow2> pWindow;
					it0.first->get_parentWindow(&pWindow);
					if (pWindow.p == pWindow2.p)
					{
						_pJSExtender = it0.second;
						break;
					}
				}
				if (_pJSExtender == nullptr)
					return S_FALSE;
			}
			g_pTangram->m_pEventProxy = new CComObject < CEventProxy >;
			CComBSTR bstrName = g_pTangram->m_pCLRProxy->AttachObjEvent((IEventProxy*)g_pTangram->m_pEventProxy, pDisp, it1->second, HtmlWndDisp);
			if(pWindow2&&_pJSExtender)
				return _pJSExtender->SinkObject(strAliasName, g_pTangram->m_pEventProxy);
			return S_OK;
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::get_Handle(LONGLONG* pVal)
{
	if (m_hWnd)
		*pVal = (LONGLONG)m_hWnd;
	return S_OK;
}

STDMETHODIMP CWndPage::GetCtrlInNode(BSTR NodeName, BSTR CtrlName, IDispatch** ppCtrl)
{
	CString strName = OLE2T(NodeName);
	if (strName == _T(""))
		return S_OK;
	auto it2 = m_mapNode.find(strName);
	if (it2 == m_mapNode.end())
		return S_OK;

	CWndNode* pNode = it2->second;
	if (pNode)
		pNode->GetCtrlByName(CtrlName, true, ppCtrl);

	return S_OK;
}

STDMETHODIMP CWndPage::get_xtml(BSTR strKey, BSTR* pVal)
{
	map<CString, CString>::iterator it = m_mapXtml.find(OLE2T(strKey));
	if (it != m_mapXtml.end())
		*pVal = it->second.AllocSysString();

	return S_OK;
}

STDMETHODIMP CWndPage::put_xtml(BSTR strKey, BSTR newVal)
{
	CString strkey = OLE2T(strKey);
	CString strVal = OLE2T(newVal);
	if (strkey == _T("") || strVal == _T(""))
		return S_OK;
	auto it = m_mapXtml.find(strkey);
	if (it != m_mapXtml.end())
		m_mapXtml.erase(it);

	m_mapXtml[strkey] = strVal;
	return S_OK;
}

STDMETHODIMP CWndPage::Extend(IDispatch* Parent, BSTR CtrlName, BSTR FrameName, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode)
{
	if (g_pTangram->m_pCLRProxy)
	{
		IDispatch* pDisp = nullptr;
		pDisp =g_pTangram->m_pCLRProxy->GetCtrlByName(Parent, CtrlName, true);
		if (pDisp)
		{
			HWND h = 0;
			h = g_pTangram->m_pCLRProxy->IsCtrlCanNavigate(pDisp);
			if (h)
			{
				CString strFrameName = OLE2T(FrameName);
				CString strKey = OLE2T(bstrKey);
				if (strFrameName == _T(""))
					FrameName = CtrlName;
				if (strKey == _T(""))
					bstrKey = CComBSTR(L"Default");
				auto it = m_mapFrame.find((HWND)h);
				if (it == m_mapFrame.end())
				{
					IWndFrame* pFrame = nullptr;
					CreateFrame(CComVariant(0), CComVariant((long)h), FrameName, &pFrame);
					return pFrame->Extend(bstrKey, bstrXml, ppRetNode);
				}
				else
				{
					return it->second->Extend(bstrKey, bstrXml, ppRetNode);
				}
			}
		}
	}
	return S_OK;
}

STDMETHODIMP CWndPage::ExtendCtrl(VARIANT MdiForm, BSTR bstrKey, BSTR bstrXml, IWndNode** ppRetNode)
{
	HWND h = 0;
	bool bMDI = false;
	if (MdiForm.vt == VT_DISPATCH)
	{
		if (g_pTangram->m_pCLRProxy)
		{
			h = g_pTangram->m_pCLRProxy->GetMDIClientHandle(MdiForm.pdispVal);
			if (h)
				bMDI = true;
			else
			{
				h = g_pTangram->m_pCLRProxy->IsCtrlCanNavigate(MdiForm.pdispVal);
				if (h)
				{
					CComBSTR bstrName(L"");
					bstrName = g_pTangram->m_pCLRProxy->GetCtrlName(MdiForm.pdispVal);
					CString strKey = OLE2T(bstrKey);
					if (strKey == _T(""))
						bstrKey = CComBSTR(L"Default");
					IWndFrame* pFrame = nullptr;
					map<HWND, CWndFrame*>::iterator it = m_mapFrame.find((HWND)h);
					if (it == m_mapFrame.end())
					{
						CreateFrame(CComVariant(0), CComVariant((long)h), bstrName, &pFrame);
						return pFrame->Extend(bstrKey, bstrXml, ppRetNode);
					}
					else
					{
						return it->second->Extend(bstrKey, bstrXml, ppRetNode);
					}
				}
			}
		}
	}
	else if (MdiForm.vt == VT_I4 || MdiForm.vt == VT_I8)
	{
		HWND hWnd = NULL;
		if (MdiForm.vt == VT_I4)
		{
			if (MdiForm.lVal == 0)
			{
				hWnd = ::FindWindowEx(m_hWnd, NULL, _T("MDIClient"), NULL);
				if (hWnd)
				{
					bMDI = true;
				}
				else
				{
					hWnd = ::GetWindow(m_hWnd, GW_CHILD);
				}
			}
			else
			{
				hWnd = (HWND)MdiForm.lVal;
				if (::IsWindow(hWnd) == false)
					hWnd = ::GetWindow(m_hWnd, GW_CHILD);
			}
		}
		if (MdiForm.vt == VT_I8)
		{
			if (MdiForm.llVal == 0)
			{
				hWnd = ::FindWindowEx(m_hWnd, NULL, _T("MDIClient"), NULL);
				if (hWnd)
				{
					bMDI = true;
				}
				else
				{
					hWnd = ::GetWindow(m_hWnd, GW_CHILD);
				}
			}
			else
			{
				hWnd = (HWND)MdiForm.llVal;
				if (::IsWindow(hWnd) == false)
					hWnd = ::GetWindow(m_hWnd, GW_CHILD);

			}
		}
		h = hWnd;
	}
	if (h)
	{
		CString strKey = OLE2T(bstrKey);
		if (strKey == _T(""))
			bstrKey = CComBSTR(L"Default");
		IWndFrame* pFrame = nullptr;
		if (bMDI)
		{
			HRESULT hr = CreateFrame(CComVariant(0), CComVariant((long)h), CComBSTR(L"TangramMdiFrame"), &pFrame);
			if (pFrame)
			{
				return pFrame->Extend(bstrKey, bstrXml, ppRetNode);
			}
		}
		else
		{
			CString strKey = OLE2T(bstrKey);
			if (strKey == _T(""))
				bstrKey = CComBSTR(L"Default");

			IWndFrame* pFrame = nullptr;
			auto it = m_mapFrame.find((HWND)h);
			if (it == m_mapFrame.end())
			{
				TCHAR szBuffer[MAX_PATH];
				::GetWindowText((HWND)h, szBuffer, MAX_PATH);
				CString strText = szBuffer;
				if (strText == _T(""))
				{
					CString s = _T("");
					s.Format(_T("Frame_%p"), h);
					strText = s;
				}
				CreateFrame(CComVariant(0), CComVariant((long)h), CComBSTR(strText), &pFrame);
				return pFrame->Extend(bstrKey, bstrXml, ppRetNode);
			}
			else
			{
				return it->second->Extend(bstrKey, bstrXml, ppRetNode);
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_EnableSinkCLRCtrlCreatedEvent(VARIANT_BOOL* pVal)
{
	*pVal = m_bEnableSinkCLRCtrlCreatedEvent;
	return S_OK;
}

STDMETHODIMP CWndPage::put_EnableSinkCLRCtrlCreatedEvent(VARIANT_BOOL newVal)
{
	m_bEnableSinkCLRCtrlCreatedEvent = newVal;
	return S_OK;
}

STDMETHODIMP CWndPage::OnChanged(DISPID dispID)
{
	if (DISPID_READYSTATE == dispID)
	{
		VARIANT vResult = { 0 };
		EXCEPINFO excepInfo;
		UINT uArgErr;
		DISPPARAMS dp = { NULL, NULL, 0, 0 };
		if (SUCCEEDED(m_pHtmlDocument2->Invoke(DISPID_READYSTATE, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dp, &vResult, &excepInfo, &uArgErr)))
		{
			assert(VT_I4 == V_VT(&vResult));
			m_lReadyState = (READYSTATE)V_I4(&vResult);
			switch (m_lReadyState)
			{
			case READYSTATE_COMPLETE:
				PostThreadMessage(m_nThreadId, WM_TANGRAM_WEBNODEDOCCOMPLETE, (WPARAM)m_hWnd, (LPARAM)0);
				break;
			default:
				break;
			}

			VariantClear(&vResult);
		}
	}

	return NOERROR;
}

STDMETHODIMP CWndPage::OnRequestEdit(DISPID dispID)
{
	return NOERROR;
}

HRESULT CWndPage::Init(CString szURL)
{
	HRESULT hr;
	LPCONNECTIONPOINTCONTAINER pCPC = nullptr;
	CComPtr<IDispatch> pScriptDisp;

	if (szURL == _T(""))
	{
		m_nScheme = INTERNET_SCHEME_HTTP;
	}
	else
	{
		URL_COMPONENTS urlComponents;
		DWORD dwFlags = 0;
		m_nScheme = INTERNET_SCHEME_UNKNOWN;
		ZeroMemory((PVOID)&urlComponents, sizeof(URL_COMPONENTS));
		urlComponents.dwStructSize = sizeof(URL_COMPONENTS);

		if (szURL)
		{
			if (InternetCrackUrl(szURL, 0, dwFlags, &urlComponents))
			{
				m_nScheme = urlComponents.nScheme;
			}
		}
	}

	if (FAILED(hr = CoCreateInstance(CLSID_HTMLDocument, nullptr, CLSCTX_INPROC_SERVER, IID_IHTMLDocument2, (LPVOID*)&m_pHtmlDocument2)))
	{
		goto Error;
	}

	if (FAILED(hr = m_pHtmlDocument2->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&pCPC)))
	{
		goto Error;
	}

	if (FAILED(hr = pCPC->FindConnectionPoint(IID_IPropertyNotifySink, &m_pCP)))
	{
		goto Error;
	}

	m_hrConnected = m_pCP->Advise((LPUNKNOWN)(IPropertyNotifySink*)this, &m_dwCookie);

	if (m_pHtmlDocument2->get_Script(&pScriptDisp) == S_OK)
	{
		CComPtr<IDispatchEx> _pScriptEx;
		hr = pScriptDisp->QueryInterface<IDispatchEx>(&_pScriptEx);
		if (hr == S_OK)
		{
			ConnectJSEngine(_pScriptEx);
			CWebPageExtender::AddObject(_T("Tangram"), (ITangram*)g_pTangram);
		}
	}

	m_mapWebPageExtender[m_pHtmlDocument2] = this;

	if (m_strURL.CompareNoCase(_T("about:blank")) == 0)
		m_bIsBlank = true;

	switch (m_nScheme)
	{
	case INTERNET_SCHEME_HTTP:
	case INTERNET_SCHEME_FTP:
	case INTERNET_SCHEME_RES:
	case INTERNET_SCHEME_GOPHER:
	case INTERNET_SCHEME_HTTPS:
	case INTERNET_SCHEME_FILE:
		// load URL using IPersistMoniker
		hr = LoadURLFromMoniker();
		break;
	case INTERNET_SCHEME_NEWS:
	case INTERNET_SCHEME_MAILTO:
	case INTERNET_SCHEME_SOCKS:
		// we don't handle these
		return E_FAIL;
		break;
	default:
		if (m_bIsBlank)
		{
			hr = LoadURLFromMoniker();
			break;
		}
		hr = LoadURLFromFile();
		break;
	}

	//if (SUCCEEDED(hr) || E_PENDING == hr)
	//{
	//	MSG msg;
	//	while (GetMessage(&msg, NULL, 0, 0))
	//	{
	//		if (WM_TANGRAM_WEBNODEDOCCOMPLETE == msg.message && NULL == msg.hwnd)
	//		{
	//			theApp.SetHook(::GetCurrentThreadId());
	//			::PostMessage(m_hWnd, WM_TANGRAM_WEBNODEDOCCOMPLETE, 1, 0);
	//			return 1;
	//		}
	//		else
	//		{
	//			DispatchMessage(&msg);
	//		}
	//	}
	//}
	return S_OK;
Error:
	if (pCPC)
		pCPC->Release();

	return hr;
}

HRESULT CWndPage::LoadURLFromMoniker()
{
	HRESULT hr;
	LPMONIKER pMk			= nullptr;
	LPBINDCTX pBCtx			= nullptr;
	LPPERSISTMONIKER pPMk	= nullptr;

	if (FAILED(hr = CreateURLMonikerEx(nullptr, CComBSTR(m_strURL), &pMk, URL_MK_UNIFORM)))
	{
		return hr;
	}

	if (FAILED(hr = CreateBindCtx(0, &pBCtx)))
	{
		goto Error;
	}

	if (SUCCEEDED(hr = m_pHtmlDocument2->QueryInterface(IID_IPersistMoniker, (LPVOID*)&pPMk)))
	{
		hr = pPMk->Load(false, pMk, pBCtx, STGM_READ);
	}

Error:
	if (pMk) pMk->Release();
	if (pBCtx) pBCtx->Release();
	if (pPMk) pPMk->Release();
	return hr;
}

HRESULT CWndPage::LoadURLFromFile()
{
	HRESULT hr;
	LPPERSISTFILE  pPF;
	if (SUCCEEDED(hr = m_pHtmlDocument2->QueryInterface(IID_IPersistFile, (LPVOID*)&pPF)))
	{
		hr = pPF->Load(CComBSTR(m_strURL), 0);
		pPF->Release();
	}

	return hr;
}

HRESULT CWndPage::UnLoad()
{
	HRESULT hr = NOERROR;

	if (m_pCP)
	{
		if (m_dwCookie)
		{
			hr = m_pCP->Unadvise(m_dwCookie);
			m_dwCookie = 0;
		}

		m_pCP->Release();
		m_pCP = nullptr;
	}

	if (m_pHtmlDocument2)
	{
		CCommonFunction::ClearObject<IUnknown>(m_pHtmlDocument2);
		m_pHtmlDocument2 = nullptr;
	}

	return NOERROR;
}

STDMETHODIMP CWndPage::ConnectTangramCtrl(ITangramCtrl* eventSource)
{
	if (g_pTangram->m_bEclipse)
	{
		if (eventSource != nullptr)
		{
			CEclipseCtrl* pCtrl = (CEclipseCtrl*)eventSource;
			HWND hCtrl = pCtrl->m_hWnd;
			auto it = m_mapNotifyCtrl.find(hCtrl);
			if (it == m_mapNotifyCtrl.end())
				m_mapNotifyCtrl[hCtrl] = pCtrl;
		}
	}
	return S_OK;
}

HRESULT CWndPage::Fire_PageLoaded(IDispatch* sender, BSTR url)
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

	if (g_pTangram->m_bEclipse)
	{
		int nSize = m_mapNotifyCtrl.size();
		if (nSize)
		{
			auto it = m_mapNotifyCtrl.begin();
			map<HWND, CEclipseCtrl*> m_mapTemp;
			for (it = m_mapNotifyCtrl.begin(); it != m_mapNotifyCtrl.end(); it++)
			{
				CEclipseCtrl* pCtrl = it->second;
				if (::IsWindow(it->first))
				{
					it->second->Fire_PageLoaded(sender, url);
				}
				else
				{
					m_mapTemp[it->first] = it->second;
				}
			}
			nSize = m_mapTemp.size();
			if (nSize)
			{
				for (it = m_mapTemp.begin(); it != m_mapTemp.end(); it++)
				{
					auto it2 = m_mapNotifyCtrl.find(it->first);
					if (it2 != m_mapNotifyCtrl.end())
						m_mapNotifyCtrl.erase(it2);
				}
				m_mapTemp.clear();
			}
		}
	}
	for (auto it : m_mapWndPageProxy)
	{
		it.second->OnPageLoaded(sender, OLE2T(url));
	}
	return hr;
}

HRESULT CWndPage::Fire_NodeCreated(IWndNode * pNodeCreated)
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

	if (g_pTangram->m_bEclipse)
	{
		int nSize = m_mapNotifyCtrl.size();
		if (nSize)
		{
			auto it = m_mapNotifyCtrl.begin();
			map<HWND, CEclipseCtrl*> m_mapTemp;
			for (it = m_mapNotifyCtrl.begin(); it != m_mapNotifyCtrl.end(); it++)
			{
				CEclipseCtrl* pCtrl = it->second;
				if (::IsWindow(it->first))
				{
					it->second->Fire_NodeCreated(pNodeCreated);
				}
				else
				{
					m_mapTemp[it->first] = it->second;
				}
			}
			nSize = m_mapTemp.size();
			if (nSize)
			{
				for (it = m_mapTemp.begin(); it != m_mapTemp.end(); it++)
				{
					auto it2 = m_mapNotifyCtrl.find(it->first);
					if (it2 != m_mapNotifyCtrl.end())
						m_mapNotifyCtrl.erase(it2);
				}
				m_mapTemp.clear();
			}
		}
	}
	for (auto it : m_mapWndPageProxy)
	{
		it.second->OnNodeCreated(pNodeCreated);
	}
	return hr;
}

HRESULT CWndPage::Fire_AddInCreated(IWndNode * pRootNode, IDispatch * pAddIn, BSTR bstrID, BSTR bstrAddInXml)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[4];
		avarParams[3] = pRootNode;
		avarParams[3].vt = VT_DISPATCH;
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

	if (g_pTangram->m_bEclipse)
	{
		int nSize = m_mapNotifyCtrl.size();
		if (nSize)
		{
			auto it = m_mapNotifyCtrl.begin();
			map<HWND, CEclipseCtrl*> m_mapTemp;
			for (it = m_mapNotifyCtrl.begin(); it != m_mapNotifyCtrl.end(); it++)
			{
				CEclipseCtrl* pCtrl = it->second;
				if (::IsWindow(it->first))
				{
					it->second->Fire_AddInCreated(pRootNode, pAddIn, bstrID, bstrAddInXml);
				}
				else
				{
					m_mapTemp[it->first] = it->second;
				}
			}
			nSize = m_mapTemp.size();
			if (nSize)
			{
				for (it = m_mapTemp.begin(); it != m_mapTemp.end(); it++)
				{
					auto it2 = m_mapNotifyCtrl.find(it->first);
					if (it2 != m_mapNotifyCtrl.end())
						m_mapNotifyCtrl.erase(it2);
				}
				m_mapTemp.clear();
			}
		}
	}
	for (auto it : m_mapWndPageProxy)
	{
		it.second->OnAddInCreated(pRootNode, pAddIn, OLE2T(bstrID), OLE2T(bstrAddInXml));
	}
	return hr;
}

HRESULT CWndPage::Fire_BeforeExtendXml(BSTR bstrXml, LONGLONG hWnd)
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
	if (g_pTangram->m_bEclipse)
	{
		int nSize = m_mapNotifyCtrl.size();
		if (nSize)
		{
			auto it = m_mapNotifyCtrl.begin();
			map<HWND, CEclipseCtrl*> m_mapTemp;
			for (it = m_mapNotifyCtrl.begin(); it != m_mapNotifyCtrl.end(); it++)
			{
				CEclipseCtrl* pCtrl = it->second;
				if (::IsWindow(it->first))
				{
					it->second->Fire_BeforeExtendXml(bstrXml, hWnd);
				}
				else
				{
					m_mapTemp[it->first] = it->second;
				}
			}
			nSize = m_mapTemp.size();
			if (nSize)
			{
				for (it = m_mapTemp.begin(); it != m_mapTemp.end(); it++)
				{
					auto it2 = m_mapNotifyCtrl.find(it->first);
					if (it2 != m_mapNotifyCtrl.end())
						m_mapNotifyCtrl.erase(it2);
				}
				m_mapTemp.clear();
			}
		}
	}
	for (auto it : m_mapWndPageProxy)
	{
		it.second->OnBeforeExtendXml(OLE2T(bstrXml), (HWND)hWnd);
	}
	return hr;
}

HRESULT CWndPage::Fire_ExtendXmlComplete(BSTR bstrXml, LONGLONG hWnd, IWndNode * pRetRootNode)
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

	if (g_pTangram->m_bEclipse)
	{
		int nSize = m_mapNotifyCtrl.size();
		if (nSize)
		{
			auto it = m_mapNotifyCtrl.begin();
			map<HWND, CEclipseCtrl*> m_mapTemp;
			for (it = m_mapNotifyCtrl.begin(); it != m_mapNotifyCtrl.end(); it++)
			{
				CEclipseCtrl* pCtrl = it->second;
				if (::IsWindow(it->first))
				{
					it->second->Fire_ExtendXmlComplete(bstrXml, hWnd, pRetRootNode);
				}
				else
				{
					m_mapTemp[it->first] = it->second;
				}
			}
			nSize = m_mapTemp.size();
			if (nSize)
			{
				for (it = m_mapTemp.begin(); it != m_mapTemp.end(); it++)
				{
					auto it2 = m_mapNotifyCtrl.find(it->first);
					if (it2 != m_mapNotifyCtrl.end())
						m_mapNotifyCtrl.erase(it2);
				}
				m_mapTemp.clear();
			}
		}
	}
	for (auto it : m_mapWndPageProxy)
	{
		it.second->OnExtendXmlComplete(OLE2T(bstrXml), (HWND)hWnd, pRetRootNode);
	}
	return hr;
}

HRESULT CWndPage::Fire_Destroy()
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

	if (g_pTangram->m_bEclipse)
	{
		int nSize = m_mapNotifyCtrl.size();
		if (nSize)
		{
			auto it = m_mapNotifyCtrl.begin();
			map<HWND, CEclipseCtrl*> m_mapTemp;
			for (it = m_mapNotifyCtrl.begin(); it != m_mapNotifyCtrl.end(); it++)
			{
				CEclipseCtrl* pCtrl = it->second;
				if (::IsWindow(it->first))
				{
					it->second->Fire_Destroy();
				}
				else
				{
					m_mapTemp[it->first] = it->second;
				}
			}
			nSize = m_mapTemp.size();
			if (nSize)
			{
				for (it = m_mapTemp.begin(); it != m_mapTemp.end(); it++)
				{
					auto it2 = m_mapNotifyCtrl.find(it->first);
					if (it2 != m_mapNotifyCtrl.end())
						m_mapNotifyCtrl.erase(it2);
				}
				m_mapTemp.clear();
			}
		}
	}
	for (auto it : m_mapWndPageProxy)
	{
		it.second->OnDestroy();
	}
	return hr;
}

HRESULT CWndPage::Fire_NodeMouseActivate(IWndNode * pActiveNode)
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

	if (g_pTangram->m_bEclipse)
	{
		int nSize = m_mapNotifyCtrl.size();
		if (nSize)
		{
			map<HWND, CEclipseCtrl*>::iterator it = m_mapNotifyCtrl.begin();
			map<HWND, CEclipseCtrl*> m_mapTemp;
			for (it = m_mapNotifyCtrl.begin(); it != m_mapNotifyCtrl.end(); it++)
			{
				CEclipseCtrl* pCtrl = it->second;
				if (::IsWindow(it->first))
				{
					it->second->Fire_NodeMouseActivate(pActiveNode);
				}
				else
				{
					m_mapTemp[it->first] = it->second;
				}
			}
			nSize = m_mapTemp.size();
			if (nSize)
			{
				for (it = m_mapTemp.begin(); it != m_mapTemp.end(); it++)
				{
					auto it2 = m_mapNotifyCtrl.find(it->first);
					if (it2 != m_mapNotifyCtrl.end())
						m_mapNotifyCtrl.erase(it2);
				}
				m_mapTemp.clear();
			}
		}
	}
	for (auto it : m_mapWndPageProxy)
	{
		it.second->OnNodeMouseActivate(pActiveNode);
	}
	return hr;
}

HRESULT CWndPage::Fire_ClrControlCreated(IWndNode * Node, IDispatch * Ctrl, BSTR CtrlName, LONGLONG CtrlHandle)
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

	if (g_pTangram->m_bEclipse)
	{
		int nSize = m_mapNotifyCtrl.size();
		if (nSize)
		{
			auto it = m_mapNotifyCtrl.begin();
			map<HWND, CEclipseCtrl*> m_mapTemp;
			for (it = m_mapNotifyCtrl.begin(); it != m_mapNotifyCtrl.end(); it++)
			{
				CEclipseCtrl* pCtrl = it->second;
				if (::IsWindow(it->first))
				{
					it->second->Fire_ClrControlCreated(Node, Ctrl, CtrlName, CtrlHandle);
				}
				else
				{
					m_mapTemp[it->first] = it->second;
				}
			}
			nSize = m_mapTemp.size();
			if (nSize)
			{
				for (it = m_mapTemp.begin(); it != m_mapTemp.end(); it++)
				{
					auto it2 = m_mapNotifyCtrl.find(it->first);
					if (it2 != m_mapNotifyCtrl.end())
						m_mapNotifyCtrl.erase(it2);
				}
				m_mapTemp.clear();
			}
		}
	}
	for (auto it : m_mapWndPageProxy)
	{
		it.second->OnClrControlCreated(Node, Ctrl, OLE2T(CtrlName), (HWND)CtrlHandle);
	}
	return hr;
}

HRESULT CWndPage::Fire_IPCMsg(IWndFrame* pSender, BSTR bstrType, BSTR bstrContent, BSTR bstrFeature)
{
	HRESULT hr = S_OK;
	int cConnections = m_vec.GetSize();
	if (cConnections)
	{
		CComVariant avarParams[4];
		avarParams[3] = pSender;
		avarParams[3].vt = VT_DISPATCH;
		avarParams[2] = bstrType;
		avarParams[2].vt = VT_BSTR;
		avarParams[1] = bstrContent;
		avarParams[1].vt = VT_BSTR;
		avarParams[0] = bstrFeature;
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
				hr = pConnection->Invoke(13, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
			}
		}
	}

	//if (g_pTangram->m_bEclipse)
	//{
	//	int nSize = m_mapNotifyCtrl.size();
	//	if (nSize)
	//	{
	//		auto it = m_mapNotifyCtrl.begin();
	//		map<HWND, CEclipseCtrl*> m_mapTemp;
	//		for (it = m_mapNotifyCtrl.begin(); it != m_mapNotifyCtrl.end(); it++)
	//		{
	//			CEclipseCtrl* pCtrl = it->second;
	//			if (::IsWindow(it->first))
	//			{
	//				it->second->Fire_ClrControlCreated(Node, Ctrl, CtrlName, CtrlHandle);
	//			}
	//			else
	//			{
	//				m_mapTemp[it->first] = it->second;
	//			}
	//		}
	//		nSize = m_mapTemp.size();
	//		if (nSize)
	//		{
	//			for (it = m_mapTemp.begin(); it != m_mapTemp.end(); it++)
	//			{
	//				auto it2 = m_mapNotifyCtrl.find(it->first);
	//				if (it2 != m_mapNotifyCtrl.end())
	//					m_mapNotifyCtrl.erase(it2);
	//			}
	//			m_mapTemp.clear();
	//		}
	//	}
	//}
	//for (auto it : m_mapWndPageProxy)
	//{
	//	it.second->OnClrControlCreated(Node, Ctrl, OLE2T(CtrlName), (HWND)CtrlHandle);
	//}
	return hr;
}

HRESULT CWndPage::Fire_TabChange(IWndNode* sender, LONG ActivePage, LONG OldPage)
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

	if (g_pTangram->m_bEclipse)
	{
		int nSize = m_mapNotifyCtrl.size();
		if (nSize)
		{
			auto it = m_mapNotifyCtrl.begin();
			map<HWND, CEclipseCtrl*> m_mapTemp;
			for (it = m_mapNotifyCtrl.begin(); it != m_mapNotifyCtrl.end(); it++)
			{
				CEclipseCtrl* pCtrl = it->second;
				if (::IsWindow(it->first))
				{
					it->second->Fire_TabChange(sender, ActivePage, OldPage);
				}
				else
				{
					m_mapTemp[it->first] = it->second;
				}
			}
			nSize = m_mapTemp.size();
			if (nSize)
			{
				for (it = m_mapTemp.begin(); it != m_mapTemp.end(); it++)
				{
					auto it2 = m_mapNotifyCtrl.find(it->first);
					if (it2 != m_mapNotifyCtrl.end())
						m_mapNotifyCtrl.erase(it2);
				}
				m_mapTemp.clear();
			}
		}
	}
	for (auto it : m_mapWndPageProxy)
	{
		it.second->OnTabChange(sender, ActivePage, OldPage);
	}
	return hr;
}

HRESULT CWndPage::Fire_TangramEvent(ITangramEventObj* pEventObj)
{
	CTangramEventObj* pEvent = (CTangramEventObj*)pEventObj;
	m_mapEventObj[(__int64)pEvent] = pEvent;
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
	return hr;
}

//
//HRESULT CWndPage::Fire_DesignTangramNode(WndFrameType DeignNodeFrameType, IWndNode* DesignNode)
//{
//	HRESULT hr = S_OK;
//	int cConnections = m_vec.GetSize();
//	if (cConnections)
//	{
//		CComVariant avarParams[2];
//		avarParams[1] = DeignNodeFrameType;
//		avarParams[1].vt = VT_I4;
//		avarParams[0] = DesignNode;
//		avarParams[0].vt = VT_DISPATCH;
//		DISPPARAMS params = { avarParams, NULL, 2, 0 };
//		for (int iConnection = 0; iConnection < cConnections; iConnection++)
//		{
//			g_pTangram->Lock();
//			IUnknown* punkConnection = m_vec.GetAt(iConnection);
//			g_pTangram->Unlock();
//			IDispatch * pConnection = static_cast<IDispatch *>(punkConnection);
//			if (pConnection)
//			{
//				CComVariant varResult;
//				hr = pConnection->Invoke(13, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &varResult, NULL, NULL);
//			}
//		}
//	}
//
//	return hr;
//}

STDMETHODIMP CWndPage::get_PageXML(BSTR* pVal)
{
	CString strData = m_strXmlHeader;
	if (strData == _T(""))
		strData = _T("<tangramdefaultpage>");
	CString strIndex = _T("@");
	for (auto it : m_mapFrame)
	{
		CComBSTR bstrXml(L"");
		strIndex += it.second->m_strFrameName;
		strIndex += _T("@");
		it.second->get_FrameXML(&bstrXml);
		strData += OLE2T(bstrXml);
	}
	map<CString, CString> m_mapTemp;
	for (auto it : m_strMapKey)
	{
		CString strKey = it.first;
		int nPos = strKey.Find(_T("@"));
		if (nPos != -1)
		{
			strKey = strKey.Mid(nPos);
			nPos = strKey.ReverseFind('@');
			if (nPos != -1)
			{
				strKey = strKey.Left(nPos + 1);
				if (strIndex.Find(strKey) == -1)
				{
					strKey.Replace(_T("@"), _T(""));
					auto it2 = m_mapTemp.find(strKey);
					if (strKey != _T(""))
					{
						if (it2 == m_mapTemp.end())
						{
							m_mapTemp[strKey] = it.second;
						}
						else
						{
							m_mapTemp[strKey] = it2->second+it.second;
						}
					}
				}
			}
		}
	}
	for (auto it : m_mapTemp)
	{
		CString strXml = _T("");
		strXml.Format(_T("<%s>%s</%s>"), it.first, it.second, it.first);
		strData += strXml;
	}
	if (m_strXmlBottom != _T(""))
		strData += m_strXmlBottom;
	else
		strData += _T("</tangramdefaultpage>");
	*pVal = CComBSTR(strData);
	return S_OK;
}

STDMETHODIMP CWndPage::put_ConfigName(BSTR newVal)
{
	m_strConfigFileNodeName = OLE2T(newVal);
	m_strConfigFileNodeName.Trim();
	if (m_strConfigFileNodeName == _T(""))
		m_strConfigFileNodeName = _T("default");
	if (m_strConfigFileNodeName.Find(_T(" ")))
		m_strConfigFileNodeName.Replace(_T(" "), _T("_"));

	if (m_strPageFileName == _T("")&&::GetParent(m_hWnd)==nullptr)
	{
		m_strPageFileName = g_pTangram->m_strExeName;
		m_strPageFilePath = g_pTangram->m_strConfigDataFile;
	}
	
	if (::PathFileExists(m_strPageFilePath))
	{
		CTangramXmlParse m_Parse2;
		if (m_Parse2.LoadFile(m_strPageFilePath))
		{
			CTangramXmlParse* m_pTangramPageParse = m_Parse2.GetChild(_T("tangrampage"));
			if (m_pTangramPageParse == nullptr)
			{
				m_Parse2.AddNode(_T("tangrampage"));
				m_pTangramPageParse = m_Parse2.GetChild(_T("tangrampage"));
			}
			if (m_pTangramPageParse)
			{
				CTangramXmlParse* m_pTangramPageParse2 = m_pTangramPageParse->GetChild(m_strConfigFileNodeName);
				if (m_pTangramPageParse2 == nullptr)
				{
					if(m_pTangramPageParse->AddNode(m_strConfigFileNodeName)!=nullptr)
						m_Parse2.SaveFile(m_strPageFilePath);
				}
				if (m_pTangramPageParse2)
				{
					int nCount = m_pTangramPageParse2->GetCount();
					for (int i = 0; i < nCount; i++)
					{
						CTangramXmlParse* _pParse = m_pTangramPageParse2->GetChild(i);
						CString _str = _T("@") + _pParse->name() + _T("@") + m_strConfigFileNodeName;
						int nCount2 = _pParse->GetCount();
						for (int i = 0; i < nCount2; i++)
						{
							CTangramXmlParse* _pParse2 = _pParse->GetChild(i);
							m_strMapKey[_pParse2->name() + _str] = _pParse2->xml();
						}
					}
				}
			}
		}
	}
	else
	{
		if (::PathIsDirectory(g_pTangram->m_strAppDataPath) == false)
		{
			if (::SHCreateDirectoryEx(NULL, g_pTangram->m_strAppDataPath, NULL))
				return S_FALSE;
		}
		CString strXml = _T("");
		strXml.Format(_T("<%s><tangrampage><%s /></tangrampage></%s>"), m_strPageFileName, m_strConfigFileNodeName, m_strPageFileName);
		CTangramXmlParse m_Parse2;
		m_Parse2.LoadXml(strXml);
		m_Parse2.SaveFile(m_strPageFilePath);
	}
	return S_OK;
}

STDMETHODIMP CWndPage::CreateFrameWithDefaultNode(ULONGLONG hFrameWnd, BSTR bstrFrameName, BSTR bstrDefaultNodeKey, BSTR bstrXml, VARIANT_BOOL bSaveToConfig, IWndNode** ppNode)
{
	CString strXml = OLE2T(bstrXml);
	strXml.Trim();
	if (strXml == _T(""))
		strXml = _T("<default><window><node name=\"Start\" /></window></default>");
	IWndFrame* pFrame = nullptr;
	CreateFrame(CComVariant(0), CComVariant((LONGLONG)hFrameWnd), bstrFrameName, &pFrame);
	if (pFrame)
	{
		if (m_pTangramDocTemplate)
			((CWndFrame*)pFrame)->m_pTangramDocTemplate = m_pTangramDocTemplate;
		pFrame->Extend(bstrDefaultNodeKey, strXml.AllocSysString(), ppNode);
		if (*ppNode&&bSaveToConfig)
		{
			(*ppNode)->put_SaveToConfigFile(true);
			//CWndFrame* pFrame2 = (CWndFrame*)::SendMessage(((CWndFrame*)pFrame)->m_hWnd, WM_TANGRAMDATA, 0, 1992);
			//if (pFrame)
			//{
			//	pFrame2->m_nWndFrameType = EclipseWorkBenchFrame;
			//	//pFrame->m_pWorkBenchFrame = this;
			//}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::ExtendFrames(BSTR bstrFrames, BSTR bstrKey, BSTR bstrXml, VARIANT_BOOL bSaveToConfigFile)
{
	CString strFrames = OLE2T(bstrFrames);
	CString strKey = OLE2T(bstrKey);
	CString strXml = OLE2T(bstrXml);
	if (strFrames == _T(""))
	{
		for (auto it : m_mapFrame)
		{
			if (it.second != m_pBKFrame)
			{
				IWndNode* pNode = nullptr;
				it.second->Extend(bstrKey, bstrXml, &pNode);
				if (pNode&&bSaveToConfigFile)
					pNode->put_SaveToConfigFile(true);
			}
		}
	}
	else
	{
		strFrames = _T(",") + strFrames;
		for (auto it : m_mapFrame)
		{
			CString strName = _T(",") + it.second->m_strFrameName + _T(",");
			if (strFrames.Find(strName) != -1)
			{
				IWndNode* pNode = nullptr;
				it.second->Extend(bstrKey, bstrXml, &pNode);
				if (pNode&&bSaveToConfigFile)
					pNode->put_SaveToConfigFile(true);
			}
		}
	}

	return S_OK;
}

STDMETHODIMP CWndPage::get_CurrentDesignFrameType(WndFrameType* pVal)
{
	if (g_pTangram->m_pDesignWindowNode)
	{
		CWndFrame* pFrame = g_pTangram->m_pDesignWindowNode->m_pTangramNodeCommonData->m_pFrame;
		*pVal = pFrame->m_nWndFrameType;
	}
	else
		*pVal = NOFrame;

	return S_OK;
}

STDMETHODIMP CWndPage::get_CurrentDesignNode(IWndNode** pVal)
{
	if (g_pTangram->m_pDesignWindowNode)
		*pVal = g_pTangram->m_pDesignWindowNode;

	return S_OK;
}
