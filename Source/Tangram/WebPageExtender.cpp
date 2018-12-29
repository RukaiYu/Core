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

#include "StdAfx.h"
#include "TangramApp.h"
#include "TangramCore.h"
#include "WndNode.h"

CJSProxyObj::CJSProxyObj(CWebPageExtender* _pJSExtender, IDispatch *pDisp, const GUID iid, CString strPrefix) :
m_dwCookie(0xFEFEFEFE), m_pDisp(pDisp), m_EventIID(iid)
{
#ifdef _DEBUG
	g_pTangram->m_nJsObj++;
#endif	
	m_pEventObj = nullptr;
	ITypeInfo* pTypeInfo = nullptr;
	m_pWebPageExtender = _pJSExtender;
	if (g_pTangram->m_pEventProxy)
	{
		m_pEventObj = g_pTangram->m_pEventProxy;
		g_pTangram->m_pEventProxy = nullptr;
	}
	else
	{
		pDisp->GetTypeInfo(0, GetUserDefaultLCID(), &pTypeInfo);
		if (pTypeInfo)
		{
			CComPtr<ITypeLib> pTypeLib;
			UINT nIndex = 0;
			pTypeInfo->GetContainingTypeLib(&pTypeLib, &nIndex);
			pTypeLib->GetTypeInfoOfGuid(iid, &m_pTypeInfo);
			pTypeInfo->Release();
		}
	}

	CString strName = strPrefix;
	auto it = m_pWebPageExtender->m_mapTangramJSObject.find(strName);
	if (it != m_pWebPageExtender->m_mapTangramJSObject.end())
	{
		int nSize = m_pWebPageExtender->m_mapTangramJSObject.size();
		CString s = _T("");
		s.Format(_T("%s_%d_"), strName, nSize);
		strName = s;
	}
	m_bstrPrefix = CComBSTR(strName);
	SetEventNames();
	m_pWebPageExtender->m_mapTangramJSObject[strName] = this;

	HRESULT hr = AtlAdvise(m_pDisp, this, m_EventIID, &m_dwCookie);
}

CJSProxyObj::~CJSProxyObj()
{
#ifdef _DEBUG
	g_pTangram->m_nJsObj--;
#endif	
	if (m_pDisp && m_dwCookie != 0xFEFEFEFE)
	{
		HRESULT hr = AtlUnadvise(m_pDisp, m_EventIID, m_dwCookie);

		m_pDisp = nullptr;
		m_dwCookie = 0xFEFEFEFE;
	}
	if (m_pEventObj)
		delete m_pEventObj;
	for (int i = 0; i < m_mapDispId.GetSize(); i++)
		SysFreeString(m_mapDispId.GetValueAt(i));
}

CWebPageExtender::CWebPageExtender(void)
{
	OutputDebugString(_T("CWebPageExtender\n"));
	m_pJSEngine = nullptr;
}

CWebPageExtender::~CWebPageExtender(void)
{
	OutputDebugString(_T("~CWebPageExtender\n"));

	for (auto it : m_mapTangramJSObject)
		delete it.second;
	m_mapTangramJSObject.erase(m_mapTangramJSObject.begin(), m_mapTangramJSObject.end());
	m_mapJSObject.erase(m_mapJSObject.begin(), m_mapJSObject.end());
}

HRESULT CJSProxyObj::QueryInterface(REFIID riid, void **ppvObject)
{
	ATLASSERT(ppvObject != NULL);
	if (ppvObject == nullptr)
		return E_POINTER;
	*ppvObject = nullptr;

	if (InlineIsEqualGUID(riid, IID_NULL))
		return E_NOINTERFACE;

	if (InlineIsEqualUnknown(riid) ||
		InlineIsEqualGUID(riid, __uuidof(IDispatch)) ||
		InlineIsEqualGUID(riid, m_EventIID))
	{
		*ppvObject = this;
		AddRef();
		return S_OK;
	}
	else
		return E_NOINTERFACE;
}

HRESULT CJSProxyObj::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (m_pWebPageExtender == nullptr)
	{
		ATLTRACE(_T("No TangramJSExtender pointer set\r\n"));
		return S_OK;
	}

	HRESULT hr;
	BSTR bstrMember = NULL;
	CComBSTR bstrFunction = m_bstrPrefix;

	bstrMember = m_mapDispId.Lookup(dispIdMember);
	if (!bstrMember || SysStringLen(bstrMember) == 0)
	{
		ATLTRACE(_T("Tangram Could not find DISPID in lookup table\r\n"));
		return S_OK;
	}

	if (m_pEventObj == nullptr)
		bstrFunction += bstrMember;

	DISPID dispIdEx;
	hr = m_pWebPageExtender->m_pJSEngine->GetDispID(bstrFunction, fdexNameCaseSensitive, &dispIdEx);
	if (FAILED(hr))
	{
		hr = m_pWebPageExtender->m_pJSEngine->GetDispID(bstrMember, fdexNameCaseSensitive, &dispIdEx);
		if (FAILED(hr))
		{
			ATLTRACE(_T("Tangram Could not find matching function '%s' in Tangram\r\n"), OLE2T(bstrFunction));
			return S_OK;
		}
	}
	m_pWebPageExtender->m_pJSEngine->InvokeEx(dispIdEx, lcid, DISPATCH_METHOD, pDispParams, pVarResult, pExcepInfo, NULL);
	return S_OK;
}

HRESULT CJSProxyObj::SetEventNames()
{
	ITypeInfo* pInfo;
	if (m_pTypeInfo == nullptr)
		pInfo = g_pTangram->m_pEventTypeInfo.p;
	else
		pInfo = m_pTypeInfo.p;
	TYPEATTR *pta = nullptr;;
	HRESULT hr = pInfo->GetTypeAttr(&pta);
	if (FAILED(hr))
		return hr;
	BSTR bstrName = ::SysAllocString(L"");
	FUNCDESC *pfd = nullptr;
	unsigned int uiNames = 0;
	for (int i = 0; i < pta->cFuncs; i++)
	{
		hr = pInfo->GetFuncDesc(i, &pfd);
		if (FAILED(hr))
			continue;

		hr = pInfo->GetNames(pfd->memid, &bstrName, 1, &uiNames);
		if (SUCCEEDED(hr) && bstrName && SysStringLen(bstrName))
		{
			m_mapDispId.Add(pfd->memid, bstrName);
			ATLTRACE(_T("*********Add function '%s' in Tangram*******\r\n"), OLE2T(bstrName));
		}
		//hr = pInfo->GetNames(pfd->memid, &bstrName, 1, &uiNames);
		//if (SUCCEEDED(hr) && bstrName && SysStringLen(bstrName))
		//{

		//	CString strEventName = OLE2T(m_bstrPrefix);
		//	strEventName += OLE2T(bstrName);
		//	m_mapDispId.Add(pfd->memid, CComBSTR(strEventName));
		//	m_mapDispID[pfd->memid] = strEventName;
		//	ATLTRACE(_T("*********Add function '%s' index: %d in Tangram*******\r\n"), strEventName, pfd->memid);
		//}
		pInfo->ReleaseFuncDesc(pfd);
		pfd = nullptr;
	}
	pInfo->ReleaseTypeAttr(pta);
	return hr;
}

BOOL CWebPageExtender::ConnectJSEngine(IDispatchEx* _pDispEx)
{
	if (m_pJSEngine != nullptr)
	{
		auto it = m_mapJSObject.begin();
		while (it != m_mapJSObject.end())
		{
			m_pJSEngine->DeleteMemberByName(it->first.AllocSysString(), fdexNameCaseSensitive);
			m_mapJSObject.erase(it);
			it = m_mapJSObject.begin();
		}

		for (auto it : m_mapTangramJSObject)
			delete it.second;
		m_mapTangramJSObject.erase(m_mapTangramJSObject.begin(), m_mapTangramJSObject.end());

		DWORD dw = -1;
		if (m_pJSEngine != nullptr)
			dw = m_pJSEngine->Release();

		m_pJSEngine = nullptr;
	}
	if (_pDispEx != nullptr)
	{
		m_pJSEngine = _pDispEx;
		m_pJSEngine->AddRef();
		return true;
	}
	return false;
}

HRESULT CWebPageExtender::AddObject(CString _strName, IDispatch* _pDisp)
{
	if (m_pJSEngine != nullptr)
	{
		if (_strName==_T("") || !_pDisp)	
			return E_INVALIDARG;

		CString strName = _strName;
		auto it = m_mapJSObject.find(strName);
		int nSize = m_mapJSObject.size();
		if (it != m_mapJSObject.end())
		{
			CString str = _T("");
			str.Format(_T("%s%d"), strName, nSize);
			strName = str;
		}
		m_mapJSObject[strName] = _pDisp;

		DISPID dispIdThis = 0;
		HRESULT hr = m_pJSEngine->GetDispID(CComBSTR(strName), fdexNameEnsure | fdexNameCaseSensitive, &dispIdThis);

		if (SUCCEEDED(hr) && dispIdThis != 0)
		{
			ATLTRACE(_T("=================Add Object '%s' to WebPage in Tangram============\r\n"), strName);
			DISPID dispIdPut = DISPID_PROPERTYPUT;
			DISPPARAMS params;
			VARIANTARG vtArg;
			EXCEPINFO ei;
			ZeroMemory(&params, sizeof(DISPPARAMS));

			VariantInit(&vtArg);
			vtArg.vt = VT_DISPATCH;
			vtArg.pdispVal = _pDisp;
			params.cArgs = 1;
			params.rgvarg = &vtArg;
			params.cNamedArgs = 1;
			params.rgdispidNamedArgs = &dispIdPut;
			hr = m_pJSEngine->InvokeEx(dispIdThis, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, NULL, &ei, NULL);
			vtArg.pdispVal = nullptr;
			VariantClear(&vtArg);
			return S_OK;
		}
		return hr;
	}
	return E_FAIL;
}

HRESULT CWebPageExtender::SinkObject(CString strPrefix, IDispatch *pUnk, LCID lcid)
{
	if (!pUnk || strPrefix==_T(""))
		return E_INVALIDARG;

	HRESULT hr = S_FALSE;
	GUID iid = GUID_NULL;

	if (g_pTangram->m_pEventProxy)
		g_pTangram->m_pEventProxy->GetConnectionInterface(&iid);
	else
	{
		CComQIPtr<IConnectionPointContainer> pContainer(pUnk);
		if (pContainer == nullptr)
			return 0;
		GUID libid;
		WORD wMajorVer = 1;
		WORD wMinorVer = 0;
		hr = AtlGetObjectSourceInterface(pUnk, &libid, &iid, &wMajorVer, &wMinorVer);
		if (iid == GUID_NULL)
		{
			IConnectionPoint* pConnectionPoint = nullptr;
			IEnumConnectionPoints* pEnum = nullptr;
			HRESULT hr = pContainer->EnumConnectionPoints(&pEnum);
			if (pEnum)
			{
				ULONG uFetched;
				while (S_OK == (pEnum->Next(1, &pConnectionPoint, &uFetched)) && uFetched >= 1)
				{
					hr = pConnectionPoint->GetConnectionInterface(&iid);
					if (hr == S_OK)
					{
						pConnectionPoint->Release();
						pConnectionPoint = nullptr;
						break;
					}
					pConnectionPoint->Release();
					pConnectionPoint = nullptr;
				}
				pEnum->Release();
			}
		}
	}
	if (iid != GUID_NULL)
		new CJSProxyObj(this, pUnk, iid, strPrefix);

	return S_OK;
}

HRESULT CWebPageExtender::UnSinkObject(IUnknown *pUnk, const GUID *piid)
{
	for (auto it = m_mapTangramJSObject.begin(); it != m_mapTangramJSObject.end(); it++)
	{
		if (pUnk == it->second->m_pDisp)
		{
			delete it->second;
			m_mapTangramJSObject.erase(it);
			return S_OK;
		}
	}

	return E_INVALIDARG;
}

