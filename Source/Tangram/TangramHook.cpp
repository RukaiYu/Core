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

// TangramHook.cpp : Defines the exported functions for the DLL application.
//


#include "stdafx.h"
#include "TangramApp.h"
#include "TangramCore.h"
#include "TangramHook.h"
#include "resource.h"
#pragma comment(lib, "oleacc.lib") 
#include <windows.h>
#include "tchar.h"
#include "mhook-lib/mhook.h"
#include "OfficePlus\LyncPlus\uccapi.h"
#include "OfficePlus\LyncPlus\UccAPIEvent.h"
#include "OfficePlus\LyncPlus\lyncaddin.h"
#include "OfficePlus\LyncPlus\lyncevent.h"
using namespace UCCAPILib;
using namespace UCCollaborationLib;
using namespace OfficePlus::LyncPlus;
using namespace OfficePlus::LyncPlus::UccApiEvent;
using namespace OfficePlus::LyncPlus::LyncClientEvent;



HRESULT(WINAPI *ORI_IUccEndpoint_Enable)(UCCAPILib::IUccEndpoint* pThis, UCCAPILib::IUccOperationContext* pContext) = NULL;

HRESULT WINAPI HOOK_IUccEndpoint_Enable(UCCAPILib::IUccEndpoint* pThis, UCCAPILib::IUccOperationContext* pContext)
{
	HRESULT hr = S_OK;
	//IUccContext * _pContext = NULL;
	//pContext->get_Context(&_pContext);
	//if(_pContext)
	//{
	//	CComPtr<IUccProperty> pIUccProperty;
	//	hr = _pContext->AddNamedProperty(CComBSTR(L"Tangram"),CComVariant(L"Tangram"),&pIUccProperty);
	//	if(hr==S_OK)
	//	{
	//	}
	//}
	hr = ORI_IUccEndpoint_Enable(pThis, pContext);
	CLyncAppProxy* pProxy = (CLyncAppProxy*)g_pTangram->m_pLyncAppProxy;
	if (hr == S_OK)
	{
		//if (pProxy->m_bSinkClient == FALSE)
		//{
		//	hr = ((CTangramLyncClientEvents*)pProxy)->DispEventAdvise(pProxy->m_pLyncClient);
		//	if (hr == S_OK)
		//		pProxy->m_bSinkClient = TRUE;
		//}
		pProxy->SinkClientEvent(TRUE);
		if (pProxy->m_spSessionManager == nullptr)
		{
			//hr = ((CTangramLyncClientEvents*)pProxy)->DispEventAdvise(pProxy->m_pLyncClient);
			//hr = pThis->QueryInterface(IID_IUccSignalingChannelManager, (void**)&pProxy->m_spSignalingChannelManager);
			hr = pThis->QueryInterface(IID_IUccSessionManager, (void**)&pProxy->m_spSessionManager);
			//hr = ((CUccSessionManagerEvents*)pProxy)->DispEventAdvise(pProxy->m_spSessionManager);
			//pProxy->m_spSessionManager->AddRef();
			//pProxy->m_spSessionManager->AddRef();
			//CComQIPtr<IUccServerSignalingSettings> pIUccServerSignalingSettings(pThis);
			//if(pIUccServerSignalingSettings)
			//{
			//	CComPtr<IUccCredentialCache> pUccCredentialCache = NULL;
			//	pIUccServerSignalingSettings->get_CredentialCache(&pUccCredentialCache);
			//	if(pUccCredentialCache)
			//	{
			//		pUccCredentialCache->get_DefaultCredential(&theApp.m_spUccCredential);
			//		CComBSTR b(L"");
			//		CComBSTR b1(L"");
			//		if(theApp.m_spUccCredential)
			//		{
			//			theApp.m_spUccCredential->get_Domain(&b);
			//			theApp.m_spUccCredential->get_UserName(&b1);
			//		}
			//	}
			//}
			//if (ORI_IUccSessionManager_CreateSession == NULL)
			//{
			//	IUnknown* pUnk = (IUnknown*)theApp.m_spSessionManager;
			//	PVOID pVmt = *(PVOID*)pUnk;
			//	ORI_IUccSessionManager_CreateSession = (HRESULT(__stdcall *)(UCCAPILib::IUccSessionManager*, enum UCCAPILib::UCC_SESSION_TYPE, UCCAPILib::IUccContext*, UCCAPILib::IUccSession**))((FARPROC*)pVmt)[3];
			//	Mhook_SetHook((PVOID*)&ORI_IUccSessionManager_CreateSession, HOOK_IUccSessionManager_CreateSession);
			//}
		}
		if (pProxy->m_spUriManager == NULL)
		{
			HRESULT hr = CoCreateInstance(CLSID_UccUriManager, NULL, CLSCTX_INPROC_SERVER, __uuidof(IUccUriManager), (LPVOID *)&pProxy->m_spUriManager);
			if (hr == S_OK)
			{
				pProxy->m_spUriManager->AddRef();
			}
		}
		//if(theApp.m_strUCMAScript!=_T(""))
		//{
		//}
		//theApp.m_pAppExtensionProxy->InitNewEndpoint(pThis);
	}
	return hr;
}

HRESULT(WINAPI *ORI_IUccEndpoint_Disable)(UCCAPILib::IUccEndpoint* pThis, UCCAPILib::IUccOperationContext* pContext) = NULL;

HRESULT WINAPI HOOK_IUccEndpoint_Disable(UCCAPILib::IUccEndpoint* pThis, UCCAPILib::IUccOperationContext* pContext)
{
	CLyncAppProxy* pProxy = (CLyncAppProxy*)g_pTangram->m_pLyncAppProxy;
	if (pProxy)
	{
		pProxy->SinkClientEvent(FALSE);
		//if (pProxy->m_bSinkSessionManager)
		//{
		//	hr = ((CUccSessionManagerEvents*)pProxy)->DispEventUnadvise(m_spSessionManager);
		//	m_bSinkSessionManager = FALSE;
		//}
		//if (pProxy->m_spSignalingChannelManager)
		//{
		//	pProxy->m_spSignalingChannelManager->Release();
		//	pProxy->m_spSignalingChannelManager = NULL;
		//}
		if (pProxy->m_spSessionManager)
		{
			//((CUccSessionManagerEvents*)g_pTangram->m_pLyncAppProxy)->DispEventUnadvise(pProxy->m_spSessionManager);
			pProxy->m_spSessionManager->Release();
			pProxy->m_spSessionManager = NULL;
		}
		//if (pProxy->m_pLyncAutomation)
		//{
		//	DWORD m_nRefCount = pProxy->m_pLyncAutomation->Release();
		//	while (m_nRefCount>0)
		//	{
		//		m_nRefCount = pProxy->m_pLyncAutomation->Release();
		//	}
		//	pProxy->m_pLyncAutomation = nullptr;
		//}
	}
	HRESULT hr = ORI_IUccEndpoint_Disable(pThis, pContext);
	return hr;
}

HRESULT(WINAPI *ORI_CreateEndpoint)(UCCAPILib::IUccPlatform* This, UCCAPILib::UCC_ENDPOINT_TYPE eType, UCCAPILib::IUccUri *pUri, BSTR bstrEndpointId, UCCAPILib::IUccContext * pContext, UCCAPILib::IUccEndpoint** ppEndpoint) = NULL;
HRESULT WINAPI NEW_CreateEndpoint(UCCAPILib::IUccPlatform* This, UCCAPILib::UCC_ENDPOINT_TYPE eType, UCCAPILib::IUccUri *pUri, BSTR bstrEndpointId, UCCAPILib::IUccContext * pContext, UCCAPILib::IUccEndpoint** ppEndpoint)
{
	HRESULT hr = S_OK;

	hr = ORI_CreateEndpoint(This, eType, pUri, bstrEndpointId, pContext, ppEndpoint);

	if (*ppEndpoint)
	{
		if (ORI_IUccEndpoint_Enable == NULL)
		{
			CComBSTR bstrUser(L"");
			pUri->get_UserAtHost(&bstrUser);
			CLyncAppProxy* pProxy = (CLyncAppProxy*)g_pTangram->m_pLyncAppProxy;
			pProxy->m_strUserUri = OLE2T(bstrUser);
			pProxy->m_spUccPlatform = This;
			//((CUccPlatformEvents*)pProxy)->DispEventAdvise(pProxy->m_spUccPlatform);
			UCCAPILib::IUccEndpoint* pTS = NULL;
			IUnknown* pUnk = *ppEndpoint;
			PVOID pVmt = *(PVOID*)pUnk;
			ORI_IUccEndpoint_Enable = (HRESULT(__stdcall *)(UCCAPILib::IUccEndpoint*, UCCAPILib::IUccOperationContext *))((FARPROC*)pVmt)[7];
			Mhook_SetHook((PVOID*)&ORI_IUccEndpoint_Enable, HOOK_IUccEndpoint_Enable);
			ORI_IUccEndpoint_Disable = (HRESULT(__stdcall *)(UCCAPILib::IUccEndpoint*, UCCAPILib::IUccOperationContext *))((FARPROC*)pVmt)[8];
			Mhook_SetHook((PVOID*)&ORI_IUccEndpoint_Disable, HOOK_IUccEndpoint_Disable);
		}

		//if(eType == UCCET_PRINCIPAL_SERVER_BASED)
		//{
		//	CComBSTR bstrUri(L"");
		//	pUri->get_UserAtHost(&bstrUri);
		//	theApp.m_pAppExtensionProxy->NewEndpoint(OLE2T(bstrUri), *ppEndpoint);
		//}
	}
	return hr;
}

typedef HRESULT (STDAPICALLTYPE * PCoCreateInstance)(
	REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext, REFIID riid, LPVOID FAR* ppv);
PCoCreateInstance ORI_CoCreateInstance = NULL;

HRESULT WINAPI NEW_CoCreateInstance(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext, REFIID riid, LPVOID FAR* ppv)
{
	//if(!theApp.m_bNeedExtension)
	//{
	//	return ORI_CoCreateInstance(rclsid,pUnkOuter,dwClsContext,riid,ppv);  
	//}
	//if(theApp.m_bNeedExtension&&theApp.m_pAppExtensionProxy==NULL&&theApp.m_strInitExtensionObjID != _T(""))
	//{
	//	ITangramExtensionProxy* pTangramExtensionProxy=NULL;
	//	CLSID clsid;
	//	HRESULT hr = ::CLSIDFromProgID(CComBSTR(theApp.m_strInitExtensionObjID),&clsid);
	//	if(clsid==CLSID_TangramManager)
	//	{
	//		theApp.m_pAppExtensionProxy = (CAppExtensionProxy*)&_AtlModule;
	//	}
	//	else
	//	{
	//		HRESULT hr = ORI_CoCreateInstance(clsid,NULL,CLSCTX_INPROC_SERVER,__uuidof(ITangramExtensionProxy),(LPVOID *)&pTangramExtensionProxy);
	//		if(hr==S_OK)
	//		{
	//			pTangramExtensionProxy->GetAppProxy((LONGLONG*)&theApp.m_pAppExtensionProxy);
	//		}
	//	}
	//}
	if(rclsid==UCCAPILib::CLSID_UccPlatform)
	{
		OutputDebugString(_T("------------------Begin Create UccPlatform------------------------\n"));
	}
	HRESULT hr = ORI_CoCreateInstance(rclsid,pUnkOuter,dwClsContext,riid,ppv);  
	if(rclsid==UCCAPILib::CLSID_UccPlatform)
	{
		OutputDebugString(_T("------------------End Create UccPlatform------------------------\n"));
		if(ORI_CreateEndpoint == NULL)
		{
			UCCAPILib::IUccPlatform* pTS = NULL;
			IUnknown* pUnk = (IUnknown*)*ppv;
			pUnk->QueryInterface(UCCAPILib::IID_IUccPlatform,(void**)&(pTS));
			PVOID pVmt = *(PVOID*)pTS;
			ORI_CreateEndpoint =(HRESULT (__stdcall *)(UCCAPILib::IUccPlatform *,UCCAPILib::UCC_ENDPOINT_TYPE eType, UCCAPILib::IUccUri *pUri,BSTR bstrEndpointId,UCCAPILib::IUccContext * pContext,UCCAPILib::IUccEndpoint** ppEndpoint))((FARPROC*)pVmt)[4];
			Mhook_SetHook((PVOID*)&ORI_CreateEndpoint, NEW_CreateEndpoint);
			pTS->Release();
		}
	}
	g_pTangram->COMObjCreated(rclsid, *ppv);
	//if(theApp.m_pAppExtensionProxy)
	//{
	//	if(rclsid==UCCAPILib::CLSID_UccPlatform)
	//	{
	//		if(ORI_CreateEndpoint == NULL)
	//		{
	//			UCCAPILib::IUccPlatform* pTS = NULL;
	//			IUnknown* pUnk = (IUnknown*)*ppv;
	//			pUnk->QueryInterface(UCCAPILib::IID_IUccPlatform,(void**)&(pTS));
	//			PVOID pVmt = *(PVOID*)pTS;
	//			ORI_CreateEndpoint =(HRESULT (__stdcall *)(UCCAPILib::IUccPlatform *,UCCAPILib::UCC_ENDPOINT_TYPE eType, UCCAPILib::IUccUri *pUri,BSTR bstrEndpointId,UCCAPILib::IUccContext * pContext,UCCAPILib::IUccEndpoint** ppEndpoint))((FARPROC*)pVmt)[4];
	//			Mhook_SetHook((PVOID*)&ORI_CreateEndpoint, NEW_CreateEndpoint);
	//			pTS->Release();
	//		}
	//	}
	//	theApp.m_pAppExtensionProxy->COMObjCreated(rclsid,*ppv);
	//}
	return hr;
}

FARPROC (WINAPI *ORI_GetProcAddress)(HMODULE hModule,LPCSTR lpProcName);
FARPROC WINAPI HOOK_GetProcAddress (HMODULE hModule,LPCSTR lpProcName)
{
	FARPROC pRet = ORI_GetProcAddress(hModule,lpProcName);
	if((int)lpProcName >= 40000)
	{
		//if(strcmp(lpProcName,"HttpOpenRequest") == 0 && ORI_HttpOpenRequest == NULL)
		//{
		//	ORI_HttpOpenRequest =(HINTERNET (__cdecl *)(HINTERNET,LPCWSTR,LPCWSTR,LPCWSTR,LPCWSTR,LPCWSTR *,DWORD,DWORD_PTR)) pRet;
		//	Mhook_SetHook((PVOID*)&ORI_HttpOpenRequest, HOOK_HttpOpenRequest);
		//}
	}
	return pRet;
}

HRESULT (WINAPI *ORI_CreateInstance)(IClassFactory* pThis,IUnknown *pUnkOuter,REFIID riid,void **ppvObject) = NULL;
HRESULT WINAPI HOOK_CreateInstance(IClassFactory* pThis,IUnknown *pUnkOuter,REFIID riid,void **ppvObject)
{
	HRESULT hr = S_FALSE;
	CString strID = _T("");
	if(::IsWindow(g_pTangram->m_RemoteObjHelperWnd.m_hWnd)==false)
	{
		HWND hWnd = ::FindWindowEx(NULL,NULL,_T("Tangram Lync Window Class"),NULL);
		//if(hWnd==0)
		//{
		//	return 	lpCreateInstance(This,pUnkOuter,riid,ppvObject);
		//}
		g_pTangram->m_RemoteObjHelperWnd.Attach(hWnd);
	}
	if(::IsWindow(g_pTangram->m_RemoteObjHelperWnd.m_hWnd))
	{
		g_pTangram->m_RemoteObjHelperWnd.GetWindowText(strID);
		g_pTangram->m_RemoteObjHelperWnd.SetWindowText(_T(""));
		if (strID != _T(""))
		{
			hr = g_pTangram->RemoteObjCreated(strID, ppvObject);
			if (hr == S_OK&&*ppvObject)
			{
				return hr;
			}
		}
	}
	//strID.Replace(_T("/"),_T(""));
	//int nPos = strID.Find(_T(":"));
	//CString strAppID = _T("");
	//if(nPos!=-1)
	//{
	//	strAppID = strID.Left(nPos);
	//	strID = strID.Mid(nPos+1);
	//}
	//if (strID!=_T("")&&strID.CompareNoCase(_T("shell.explorer.2")))
	//{
	//	if(strID.CompareNoCase(_T("TangramRemoteConnector"))==0)
	//	{
	//	//	CComObject<CTangramRemoteConnector>* pObj = new CComObject<CTangramRemoteConnector>;
	//	//	theApp.m_mapRemoteTangramRemoteConnector[strAppID] = pObj;
	//	//	pObj->m_strKey = strAppID;
	//	//	*ppvObject = (IDispatch*)pObj;
	//	//	pObj->AddRef();
	//	//	if(theApp.m_pAppExtensionProxy)
	//	//	{
	//	//		theApp.m_pAppExtensionProxy->RemoteObjCreated(strAppID,strID, (IDispatch*)pObj);
	//	//	}
	//	//	return S_OK;
	//	}
	//	else
	//	{
	//		CComPtr<IDispatch> pDisp;
	//		hr = pDisp.CoCreateInstance(CComBSTR(strID));
	//		*ppvObject = pDisp.p;
	//		//if(theApp.m_pAppExtensionProxy)
	//		//{
	//		//	theApp.m_pAppExtensionProxy->RemoteObjCreated(strAppID,strID, pDisp.p);
	//		//}
	//		pDisp.p->AddRef();
	//	}
	//}
	//else
	//{
	//	hr = ORI_CreateInstance(pThis,pUnkOuter,riid,ppvObject);
	//}	
	hr = ORI_CreateInstance(pThis, pUnkOuter, riid, ppvObject);
	return hr;
}

HRESULT (WINAPI *ORI_CoRegisterClassObject)(REFCLSID rclsid,
										 IUnknown * pUnk,
										 DWORD dwClsContext,
										 DWORD flags,
										 LPDWORD lpdwRegister);
HRESULT WINAPI NEW_CoRegisterClassObject(REFCLSID rclsid,
										 IUnknown * pUnk,
										 DWORD dwClsContext,
										 DWORD flags,
										 LPDWORD lpdwRegister)
{
	HRESULT hr = ORI_CoRegisterClassObject(rclsid,pUnk,dwClsContext,flags,lpdwRegister);

	if(rclsid== g_pTangram->m_RemoteObjClsid&&ORI_CreateInstance==NULL)
	{
		IClassFactory* pClassFactory = NULL;

		if (pUnk->QueryInterface(IID_IClassFactory,(void**)&pClassFactory) == S_OK)
		{
			PVOID pVmt = *(PVOID*)pClassFactory;
			ORI_CreateInstance =(HRESULT (__stdcall *)(IClassFactory *,IUnknown *,REFIID ,void **))((FARPROC*)pVmt)[3];
			Mhook_SetHook((PVOID*)&ORI_CreateInstance, HOOK_CreateInstance);
			//lpCreateInstance = HookVmt(pClassFactory,3,NEW_CreateInstance);
			pClassFactory->Release();
		}				
		//g_pClsidProxyObject[g_iAppIndex] = rclsid;
		//g_pdwAppID[g_iAppIndex] = ::GetCurrentProcessId();
		//g_iAppIndex++;
	}
	return hr;
}

BOOL HookApi()
{
	/////////////////////////////////////////////////////////////////////////////

	//ORI_CreateFile = CreateFile;
	//Mhook_SetHook((PVOID*)&ORI_CreateFile, HOOK_CreateFile);
	//ORI_GetProcAddress = GetProcAddress;
	//Mhook_SetHook((PVOID*)&ORI_GetProcAddress, HOOK_GetProcAddress);
	//ORI_HttpOpenRequest = HttpOpenRequest;
	//Mhook_SetHook((PVOID*)&ORI_HttpOpenRequest, HOOK_HttpOpenRequest);
	//ORI_TrackPopupMenuEx = TrackPopupMenuEx;
	//Mhook_SetHook((PVOID*)&ORI_TrackPopupMenuEx, HOOK_TrackPopupMenuEx);
	//ORI_SetWindowPos = SetWindowPos;
	//Mhook_SetHook((PVOID*)&ORI_SetWindowPos, NEW_SetWindowPos);

	//ORI_FindResource = FindResource;
	//Mhook_SetHook((PVOID*)&ORI_FindResource, HOOK_FindResource);
	//ORI_LoadResource = LoadResource;
	//Mhook_SetHook((PVOID*)&ORI_LoadResource, HOOK_LoadResource);
	//ORI_SizeofResource = SizeofResource;
	//Mhook_SetHook((PVOID*)&ORI_SizeofResource, HOOK_SizeofResource);
	//ORI_LoadIcon = LoadIcon;
	//Mhook_SetHook((PVOID*)&ORI_LoadIcon, NEW_LoadIcon);
	//ORI_SetWindowText = SetWindowText;
	//Mhook_SetHook((PVOID*)&ORI_SetWindowText, NEW_SetWindowText);
	//ORI_LoadImage = LoadImage;
	//Mhook_SetHook((PVOID*)&ORI_LoadImage, NEW_LoadImage);
	//ORI_LoadBitmap = LoadBitmap;
	//Mhook_SetHook((PVOID*)&ORI_LoadBitmap, NEW_LoadBitmap);
	ORI_CoCreateInstance = CoCreateInstance;
	Mhook_SetHook((PVOID*)&ORI_CoCreateInstance, NEW_CoCreateInstance);
	ORI_CoRegisterClassObject = CoRegisterClassObject;
	Mhook_SetHook((PVOID*)&ORI_CoRegisterClassObject, NEW_CoRegisterClassObject);
	return TRUE;
}

BOOL UnHookApi()
{
	//Mhook_Unhook((PVOID*)&ORI_GetProcAddress);
	//Mhook_Unhook((PVOID*)&ORI_TrackPopupMenuEx);
	//Mhook_Unhook((PVOID*)&ORI_FindResource);
	//Mhook_Unhook((PVOID*)&ORI_LoadResource);
	//Mhook_Unhook((PVOID*)&ORI_SizeofResource);
	//Mhook_Unhook((PVOID*)&ORI_LoadIcon);
	//Mhook_Unhook((PVOID*)&ORI_SetWindowText);
	//Mhook_Unhook((PVOID*)&ORI_LoadBitmap);
	//Mhook_Unhook((PVOID*)&ORI_LoadImage);
	if(ORI_CoCreateInstance)
		Mhook_Unhook((PVOID*)&ORI_CoCreateInstance);
	if(ORI_CoRegisterClassObject)
		Mhook_Unhook((PVOID*)&ORI_CoRegisterClassObject);
	if(ORI_CreateInstance)
		Mhook_Unhook((PVOID*)&ORI_CreateInstance);
	if(ORI_CreateEndpoint)
		Mhook_Unhook((PVOID*)&ORI_CreateEndpoint);
	if(ORI_IUccEndpoint_Enable)
		Mhook_Unhook((PVOID*)&ORI_IUccEndpoint_Enable);
	if(ORI_IUccEndpoint_Disable)
		Mhook_Unhook((PVOID*)&ORI_IUccEndpoint_Disable);
	//if(ORI_CreateTextServices)
	//	Mhook_Unhook((PVOID*)&ORI_CreateTextServices);
	//if(ORI_ITextServices_Release)
	//	Mhook_Unhook((PVOID*)&ORI_ITextServices_Release);
	////if(ORI_IUccPlatform_Shutdown)
	////	Mhook_Unhook((PVOID*)&ORI_IUccPlatform_Shutdown);
	////if(ORI_IUccSession_AddParticipant)
	////	Mhook_Unhook((PVOID*)&ORI_IUccSession_AddParticipant);
	//if(ORI_ITextServices_TxGetText)
	//	Mhook_Unhook((PVOID*)&ORI_ITextServices_TxGetText);
	//if(ORI_IUccSessionManager_CreateSession)
	//	Mhook_Unhook((PVOID*)&ORI_IUccSessionManager_CreateSession);
	return TRUE;
}

BOOL HookCoCreateInstance(BOOL bHook)
{
	if (bHook)
	{
		ORI_CoCreateInstance = CoCreateInstance;
		return Mhook_SetHook((PVOID*)&ORI_CoCreateInstance, NEW_CoCreateInstance);
	}
	else
	{
		return Mhook_Unhook((PVOID*)&ORI_CoCreateInstance);
	}
}

BOOL HookCoRegisterClassObject(BOOL bHook)
{
	if (bHook)
	{
		ORI_CoRegisterClassObject = CoRegisterClassObject;
		return Mhook_SetHook((PVOID*)&ORI_CoRegisterClassObject, NEW_CoRegisterClassObject);
	}
	else
	{
		if (ORI_CreateInstance)
		{
			Mhook_Unhook((PVOID*)&ORI_CreateInstance);
		}
		return Mhook_Unhook((PVOID*)&ORI_CoRegisterClassObject);
	}
}
