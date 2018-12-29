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

#include "../../stdafx.h"
#include "../../TangramApp.h"
#include "../../TangramHook.h"
#include "LyncAddin.h"
#include "LyncEvent.cpp"
#include "uccapievent.cpp"
#include "LyncCommon.h"

namespace OfficePlus
{
	namespace LyncPlus
	{
		CLyncAddin::CLyncAddin() : CTangram()
		{
			//m_nAppID = 9;
			m_nRichEditCount = 0;
			m_bOfficeApp = true;
			m_hMainWnd	 = nullptr;
			m_hMainWnd2	 = nullptr;
			m_hTabFrameWnd = nullptr;
			HookCoCreateInstance(true);
			m_pLyncAppProxy = new CComObject<CLyncAppProxy>;
			m_hCBTHook = SetWindowsHookEx(WH_CBT, CTangramApp::CBTProc, NULL, GetCurrentThreadId());
		}

		CLyncAddin::~CLyncAddin()
		{
			UnHookApi();
		}

		HRESULT CLyncAddin::COMObjCreated(REFCLSID rclsid, LPVOID pv)
		{
			return 0;
		};

		void CLyncAddin::WindowCreated(CString strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd)
		{ 
			if (strClassName.CompareNoCase(_T("CommunicatorMainWindowClass")) == 0)
			{
				m_hMainWnd = hWnd;
				return;
			}
			if (strClassName.CompareNoCase(_T("NetUINativeHWNDHost")) == 0)
			{
				if (m_hMainWnd2 == nullptr)
				{
					SubclassWindow(m_hMainWnd);
					m_hMainWnd2 = hWnd;
					Init();
					m_pLyncAppProxy->InitLyncApp();
				}
				else
				{
					::GetClassName(hPWnd, m_szBuffer, MAX_PATH);
					CString strClass = CString(m_szBuffer);
					if (strClass.CompareNoCase(_T("LyncConversationWindowClass")) == 0)
					{
						//if (theApp.m_pCurUCPlusIMWnd&&::IsWindow(m_pCurUCPlusIMWnd->m_hWnd) == false)
						//	m_pCurUCPlusIMWnd->SubclassWindow(hParent);
						//if (m_pCurUCPlusIMWnd&&m_pCurUCPlusIMWnd->m_hConWnd == NULL)
						//	m_pCurUCPlusIMWnd->m_hConWnd = hWnd;
						::PostMessage(m_hMainWnd, WM_LYNCIMWNDCREATED, (WPARAM)hWnd, (LPARAM)hPWnd);
					}
				}
				return;
			}
			if (strClassName.CompareNoCase(_T("RICHEDIT60W")) == 0)
			{
				if (::IsChild(m_hMainWnd, hWnd))
					m_nRichEditCount++;
				return;
			}
			if (strClassName.CompareNoCase(_T("LyncVdiVideoPlaceHolderWindowClass")) == 0)
			{
				CLyncAppProxy* pProxy = (CLyncAppProxy*)m_pLyncAppProxy;
				if (pProxy->m_pLyncConversationManager)
				{
					CComPtr<IConversationCollection> pCol;
					pProxy->m_pLyncConversationManager->get_Conversations(&pCol);
					//pProxy->m_pLyncConversationManager->
					//long nCount = 0;
					//pCol->get_Count(&nCount);
					//if (nCount)
					//{

					//}
				}
				return;
			}
			if (strClassName.CompareNoCase(_T("LyncTabFrameHostWindowClass")) == 0)
			{
				m_hTabFrameWnd = hWnd;
				return;
			}
			if (strClassName.CompareNoCase(_T("LyncConversationWindowClass")) == 0)
			{
				return;
			}
		}

		void CLyncAddin::WindowDestroy(HWND hWnd)
		{
			::GetClassName(hWnd, m_szBuffer, MAX_PATH);
			CString strClassName = CString(m_szBuffer);
			if (hWnd == m_hMainWnd)
			{
			}
			if (strClassName.CompareNoCase(_T("RICHEDIT60W")) == 0)
			{
				if (::IsChild(m_hMainWnd, hWnd))
					m_nRichEditCount--;
				return;
			}
		}

		LRESULT CLyncAddin::OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
		{
			m_pLyncAppProxy->Close();
			LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
			return lRes;
		}

		LRESULT CLyncAddin::OnConsationWndCreated(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL&)
		{
			CLyncAppProxy* pProxy = (CLyncAppProxy*)m_pLyncAppProxy;
			if (pProxy->m_pLyncConversationManager)
			{
				CComPtr<IConversationCollection> pCol;
				pProxy->m_pLyncConversationManager->get_Conversations(&pCol);
				if (pCol)
				{
					long nCount = 0;
					pCol->get_Count(&nCount);
					if (nCount)
					{

					}
				}
			}
			LRESULT lRes = DefWindowProc(uMsg, wParam, lParam);
			return lRes;
		}

		CLyncAppProxy::CLyncAppProxy()
		{
			m_bSinkClient					= FALSE;
			m_bSinkSessionManager			= FALSE;
			m_nLyncState					= ucClientStateSignedOut;
			m_strUserUri					= _T("");
			m_pLyncClient					= nullptr;
			m_pLyncAutomation				= nullptr;
			m_spUriManager					= nullptr;
			m_pLyncRoomManager				= nullptr;
			m_pUCOfficeIntegration			= nullptr;
			m_pLyncConversationManager		= nullptr;
			m_spUccPlatform					= nullptr;
			m_spSessionManager				= nullptr;
			//m_spSignalingChannelManager	= nullptr;
			g_pTangram->m_pLyncAppProxy		= nullptr;

		}

		CLyncAppProxy::~CLyncAppProxy()
		{
			HRESULT hr = S_OK;
			DWORD m_nRefCount = -1;
			if (m_spUriManager)
			{
				m_nRefCount = m_spUriManager->Release();
				while (m_nRefCount>0)
				{
					m_nRefCount = m_spUriManager->Release();
				}
				m_spUriManager = nullptr;
			}
			if (m_pLyncAutomation)
			{
				m_nRefCount = m_pLyncAutomation->Release();
				while (m_nRefCount>0)
				{
					m_nRefCount = m_pLyncAutomation->Release();
				}
				m_pLyncAutomation = nullptr;
			}
			if (m_pLyncRoomManager)
			{
				((CLyncRoomManagerEvents*)this)->DispEventUnadvise(m_pLyncRoomManager);
				m_nRefCount = m_pLyncRoomManager->Release();
				while (m_nRefCount>0)
				{
					m_nRefCount = m_pLyncRoomManager->Release();
				}
				m_pLyncRoomManager = nullptr;
			}
			if (m_pLyncConversationManager)
			{
				((CLyncConversationManagerEvents*)this)->DispEventUnadvise(m_pLyncConversationManager);
				m_nRefCount = m_pLyncConversationManager->Release();
				while (m_nRefCount>0)
				{
					m_nRefCount = m_pLyncConversationManager->Release();
				}
				m_pLyncConversationManager = nullptr;
			}
			if (m_spSessionManager)
			{
				if (m_bSinkSessionManager)
				{
					hr = ((CUccSessionManagerEvents*)this)->DispEventUnadvise(m_spSessionManager);
					m_bSinkSessionManager = FALSE;
				}
				m_nRefCount = m_spSessionManager->Release();
				while (m_nRefCount>0)
				{
					m_nRefCount = m_spSessionManager->Release();
				}
				m_spSessionManager = nullptr;
			}
			//if (m_spSignalingChannelManager)
			//{
			//	DWORD m_nRefCount = m_spSignalingChannelManager->Release();
			//	while (m_nRefCount>0)
			//	{
			//		m_nRefCount = m_spSignalingChannelManager->Release();
			//	}
			//	m_spSignalingChannelManager = nullptr;
			//}
			if (m_pLyncClient)
			{
				((CTangramLyncClientEvents*)this)->DispEventUnadvise(m_pLyncClient);
				m_nRefCount = m_pLyncClient->Release();
				while (m_nRefCount>0)
				{
					m_nRefCount = m_pLyncClient->Release();
				}
				m_pLyncClient = nullptr;
			}
			//g_pTangram->Lock();
			//g_pTangram->Unlock();
		}

		HRESULT	CLyncAppProxy::SinkClientEvent(BOOL bSink)
		{
			HRESULT hr = S_OK;
			if (m_bSinkClient == FALSE)
			{
				if (m_pLyncClient == nullptr)
					_InitLyncApp();

				//hr = ((CTangramLyncClientEvents*)this)->DispEventAdvise(m_pLyncClient);
				//if (hr == S_OK)
				//{
				//	m_bSinkClient = TRUE;
				//}
			}
			else
			{
				//hr = ((CTangramLyncClientEvents*)this)->DispEventUnadvise(m_pLyncClient);
				//if (hr == S_OK)
				//{
				//	m_bSinkClient = FALSE;
				//}
			}
			return hr;
		}

		STDMETHODIMP CLyncAppProxy::InitLyncApp()
		{
			CLyncAppProxy* pProxy = (CLyncAppProxy*)g_pTangram->m_pLyncAppProxy;
			//if (hr == S_OK)
			{
				//if (pProxy->m_bSinkClient == FALSE)
				//{
				//	hr = ((CTangramLyncClientEvents*)pProxy)->DispEventAdvise(pProxy->m_pLyncClient);
				//	if (hr == S_OK)
				//		pProxy->m_bSinkClient = TRUE;
				//}
				//pProxy->SinkClientEvent(TRUE);
				//if (pProxy->m_spSessionManager == nullptr)
				//{
				//	//hr = ((CTangramLyncClientEvents*)pProxy)->DispEventAdvise(pProxy->m_pLyncClient);
				//	//hr = pThis->QueryInterface(IID_IUccSignalingChannelManager, (void**)&pProxy->m_spSignalingChannelManager);
				//	QueryInterface(IID_IUccSessionManager, (void**)&pProxy->m_spSessionManager);
				//	//hr = ((CUccSessionManagerEvents*)pProxy)->DispEventAdvise(pProxy->m_spSessionManager);
				//	//pProxy->m_spSessionManager->AddRef();
				//	//pProxy->m_spSessionManager->AddRef();
				//	//CComQIPtr<IUccServerSignalingSettings> pIUccServerSignalingSettings(pThis);
				//	//if(pIUccServerSignalingSettings)
				//	//{
				//	//	CComPtr<IUccCredentialCache> pUccCredentialCache = NULL;
				//	//	pIUccServerSignalingSettings->get_CredentialCache(&pUccCredentialCache);
				//	//	if(pUccCredentialCache)
				//	//	{
				//	//		pUccCredentialCache->get_DefaultCredential(&theApp.m_spUccCredential);
				//	//		CComBSTR b(L"");
				//	//		CComBSTR b1(L"");
				//	//		if(theApp.m_spUccCredential)
				//	//		{
				//	//			theApp.m_spUccCredential->get_Domain(&b);
				//	//			theApp.m_spUccCredential->get_UserName(&b1);
				//	//		}
				//	//	}
				//	//}
				//	//if (ORI_IUccSessionManager_CreateSession == NULL)
				//	//{
				//	//	IUnknown* pUnk = (IUnknown*)theApp.m_spSessionManager;
				//	//	PVOID pVmt = *(PVOID*)pUnk;
				//	//	ORI_IUccSessionManager_CreateSession = (HRESULT(__stdcall *)(UCCAPILib::IUccSessionManager*, enum UCCAPILib::UCC_SESSION_TYPE, UCCAPILib::IUccContext*, UCCAPILib::IUccSession**))((FARPROC*)pVmt)[3];
				//	//	Mhook_SetHook((PVOID*)&ORI_IUccSessionManager_CreateSession, HOOK_IUccSessionManager_CreateSession);
				//	//}
				//}
				//if (pProxy->m_spUriManager == NULL)
				//{
				//	HRESULT hr = CoCreateInstance(CLSID_UccUriManager, NULL, CLSCTX_INPROC_SERVER, __uuidof(IUccUriManager), (LPVOID *)&pProxy->m_spUriManager);
				//	if (hr == S_OK)
				//	{
				//		pProxy->m_spUriManager->AddRef();
				//	}
				//}
				//if(theApp.m_strUCMAScript!=_T(""))
				//{
				//}
				//theApp.m_pAppExtensionProxy->InitNewEndpoint(pThis);
			}
			//g_pTangram->Init();
			//if (_InitLyncApp() == false)
			//{
			//	return S_FALSE;
			//}
			return S_OK;
		}

		STDMETHODIMP CLyncAppProxy::Close()
		{
			delete this;
			g_pTangram->m_pLyncAppProxy = nullptr;
			return S_OK;
		}

		STDMETHODIMP CLyncAppProxy::get_ActiveWorkBenchWindow(BSTR bstrID, IWorkBenchWindow** pVal)
		{
			//CString strID = OLE2T(bstrID);
			//strID.Trim();
			//if (strID != _T(""))
			//{
			//	ITangram* pTangram = nullptr;
			//	m_pAddin->get_RemoteTangram(bstrID, &pTangram);
			//	if (pTangram)
			//	{
			//		IWorkBenchWindow* pRet = nullptr;
			//		ITangramExtender* pExtender = nullptr;
			//		pTangram->get_Extender(&pExtender);
			//		if (pExtender)
			//		{
			//			CComQIPtr<IEclipseExtender> pEclipse(pExtender);
			//			if (pEclipse)
			//				pEclipse->get_ActiveWorkBenchWindow(bstrID, &pRet);
			//			if (pRet)
			//			{
			//				*pVal = pRet;
			//				(*pVal)->AddRef();
			//			}
			//		}
			//	}
			//}

			return S_OK;
		}

		BOOL CLyncAppProxy::_InitLyncApp()
		{
			BSTR bstrVer;
			CComPtr<IUCOfficeIntegration> _pUCOfficeIntegration;
			CLSID cls;
			::CLSIDFromProgID(L"lync.UCOfficeIntegration.1", &cls);
			if (_pUCOfficeIntegration.CoCreateInstance(CComBSTR("lync.UCOfficeIntegration.1"), 0, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER) != S_OK)
			{
				if (_pUCOfficeIntegration.CoCreateInstance(CLSID_UCOfficeIntegration, 0, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER) == S_OK)
					bstrVer = ::SysAllocString(L"14.0.0.0");
				else
					return FALSE;
			}
			else
				bstrVer = ::SysAllocString(L"15.0.0.0");
			if (_pUCOfficeIntegration)
			{
				IDispatch* pLyncClient = NULL;
				IDispatch* pLyncAuto = NULL;
				_pUCOfficeIntegration->GetInterface(bstrVer, oiInterfaceILyncClient, (IDispatch * *)&pLyncClient);
				_pUCOfficeIntegration->GetInterface(bstrVer, oiInterfaceIAutomation, (IDispatch * *)&pLyncAuto);
				HRESULT hr = pLyncClient->QueryInterface(UCCollaborationLib::IID_ILyncClient, (void**)&m_pLyncClient);
				m_pLyncClient->AddRef();

				hr = pLyncAuto->QueryInterface(UCCollaborationLib::IID_IAutomation, (void**)&m_pLyncAutomation);
				m_pLyncAutomation->AddRef();
				//m_pLyncAutomation->AddRef();
				::SysFreeString(bstrVer);
				ClientType type;
				m_pLyncClient->get_Type(&type);
				hr = ((CTangramLyncClientEvents*)this)->DispEventAdvise(m_pLyncClient);
				hr = m_pLyncClient->get_ConversationManager(&m_pLyncConversationManager);
				if (hr == S_OK)
					hr = ((CLyncConversationManagerEvents*)this)->DispEventAdvise(m_pLyncConversationManager);

				CComQIPtr<IClient2>pLyncClient2(pLyncClient);
				if (pLyncClient2)
				{
					hr = pLyncClient2->get_RoomManager(&m_pLyncRoomManager);
					m_pLyncRoomManager->AddRef();
					((CLyncRoomManagerEvents*)this)->DispEventAdvise(m_pLyncRoomManager);
				}
				return TRUE;
			}
			return FALSE;
		}

		void CLyncAppProxy::OnConversationAdded(IConversationManager* _eventSource, IConversationManagerEventData* _eventData) 
		{
		}

		void CLyncAppProxy::OnConversationRemoved(IConversationManager* _eventSource, IConversationManagerEventData* _eventData)
		{
		}

		void CLyncAppProxy::OnStateChanged(IClient* _eventSource, IClientStateChangedEventData* _eventData)
		{
			_eventData->get_NewState(&m_nLyncState);
			switch (m_nLyncState)
			{
			case ClientState::ucClientStateSignedIn:
				{
					if (m_bSinkSessionManager==FALSE)
					{
						HRESULT hr = ((CUccSessionManagerEvents*)this)->DispEventAdvise(m_spSessionManager);
						if (hr == S_OK)
							m_bSinkSessionManager = TRUE;
					}
				}
				break;
			case ClientState::ucClientStateSignedOut:
				{
					if (m_bSinkSessionManager)
					{
						HRESULT hr = ((CUccSessionManagerEvents*)this)->DispEventUnadvise(m_spSessionManager);
						m_bSinkSessionManager = FALSE;
					}
				}
				break;
			case ClientState::ucClientStateSigningIn:
				if (m_bSinkSessionManager)
				{
					//HRESULT hr = ((CUccSessionManagerEvents*)this)->DispEventUnadvise(m_spSessionManager);
					//m_bSinkSessionManager = FALSE;
				}
				break;
			case ClientState::ucClientStateSigningOut:
				break;
			default:
				break;
			}
		}

		void CLyncAppProxy::OnShutdown(IUccPlatform* pEventSource, IUccOperationProgressEvent* pEventData)
		{
			((CUccPlatformEvents*)this)->DispEventUnadvise(m_spUccPlatform);
		}

		HRESULT CLyncAppProxy::OnIncomingSession(IUccEndpoint* pEventSource, IUccIncomingSessionEvent* pEventData)
		{
			CComPtr<IUccSession> pUccSession;
			pEventData->get_Session(&pUccSession);
			CComPtr<IUccSessionParticipant> pIUccSessionParticipant;
			pEventData->get_Inviter(&pIUccSessionParticipant);
			return S_OK;
		}

		HRESULT CLyncAppProxy::OnOutgoingSession(IUccEndpoint* pEventSource, IUccOutgoingSessionEvent* pEventData)
		{
			CComPtr<IUccSession> pUccSession;
			pEventData->get_Session(&pUccSession);
			return S_OK;
		}

		//CLyncRoomManagerEvents:
		void CLyncAppProxy::OnFollowedRoomAdded(IRoomManager* _eventSource, IFollowedRoomsChangedEventData* _eventData)
		{
			CComPtr<IRoom> pRoom;
			_eventData->get_Room(&pRoom);
			if (pRoom)
			{
				//RoomProperty m_nRoomProperty;
				CComPtr<IRoomPropertyDictionary> pIRoomPropertyDictionary;
				pRoom->get_Properties(&pIRoomPropertyDictionary);
				if (pIRoomPropertyDictionary)
				{
					CComVariant var;
					HRESULT hr = pIRoomPropertyDictionary->get_Item(RoomProperty::ucRoomUri, &var);
					if (hr == S_OK)
					{
						CString strURI = OLE2T(var.bstrVal);
						CLyncRoomObj* pLyncRoomObj = new CLyncRoomObj();
						pLyncRoomObj->m_pRoom = pRoom.p;
						pLyncRoomObj->DispEventAdvise(pRoom.p);
						m_mapRoom[strURI] = pLyncRoomObj;
					}
				}
			}
		}

		void CLyncAppProxy::OnFollowedRoomRemoved(IRoomManager* _eventSource, IFollowedRoomsChangedEventData* _eventData)
		{
			CComPtr<IRoom> pRoom;
			_eventData->get_Room(&pRoom);
			if (pRoom)
			{
				//RoomProperty m_nRoomProperty;
				CComPtr<IRoomPropertyDictionary> pIRoomPropertyDictionary;
				pRoom->get_Properties(&pIRoomPropertyDictionary);
				if (pIRoomPropertyDictionary)
				{
					CComVariant var;
					HRESULT hr = pIRoomPropertyDictionary->get_Item(RoomProperty::ucRoomUri, &var);
					if (hr == S_OK)
					{
						CString strURI = OLE2T(var.bstrVal);
						auto it = m_mapRoom.find(strURI);
						if (it != m_mapRoom.end())
						{
							it->second->DispEventUnadvise(it->second->m_pRoom);
							delete it->second;
							m_mapRoom.erase(it);
						}
					}
				}
			}
		}

		void CLyncAppProxy::OnRoomManagerStateChanged(IRoomManager* _eventSource, IRoomManagerStateChangedEventData* _eventData)
		{
			RoomManagerState m_nState;
			_eventData->get_NewState(&m_nState);
			switch (m_nState)
			{
			case RoomManagerState::ucRoomManagerDisabled:
				break;
			case RoomManagerState::ucRoomManagerEnabled:
				break;
			}
		}
	}
}

