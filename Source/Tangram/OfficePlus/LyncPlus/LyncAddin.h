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
#include "CommonFunction.h"
#include "tangrambase.h"
#include "uccapievent.h"
#include "LyncEvent.h"
#include "..\Tangram\TangramCore.h"
#include "lync.h"
#include "uccapi.h"
#include "UccAPIEvent.h"

using namespace UCCAPILib;
using namespace UCCollaborationLib;
using namespace OfficePlus::LyncPlus::UccApiEvent;
using namespace OfficePlus::LyncPlus::LyncClientEvent;

namespace OfficePlus
{
	namespace LyncPlus
	{
		class CLyncRoomObj;
		class CLyncAppProxy :
			public CTangramLyncClientEvents,
			public CComObjectRootBase,
			public CLyncRoomManagerEvents,
			public CUccPlatformEvents,
			public CUccSessionManagerEvents,
			public CLyncConversationManagerEvents,
			public IDispatchImpl<ILyncExtender, &IID_ILyncExtender, &LIBID_Tangram, /*wMajor =*/ 1, /*wMinor =*/ 0>
		{
		public:
			CLyncAppProxy();
			virtual ~CLyncAppProxy();

			void Lock() {}
			void Unlock() {}

			BOOL									m_bSinkClient;
			BOOL									m_bSinkSessionManager;
			ClientState								m_nLyncState;
			CString									m_strUserUri;
			ILyncClient*							m_pLyncClient;
			IAutomation*							m_pLyncAutomation;
			IUccUriManager*							m_spUriManager;
			IRoomManager*							m_pLyncRoomManager;
			IUCOfficeIntegration*					m_pUCOfficeIntegration;
			IConversationManager*					m_pLyncConversationManager;
			IUccPlatform*							m_spUccPlatform;
			IUccSessionManager*						m_spSessionManager;

			map<CString, CLyncRoomObj*>				m_mapRoom;
			BOOL _InitLyncApp();
			HRESULT	SinkClientEvent(BOOL bSink);

			BEGIN_COM_MAP(CLyncAppProxy)
				COM_INTERFACE_ENTRY(ILyncExtender)
				COM_INTERFACE_ENTRY(IDispatch)
			END_COM_MAP()

			STDMETHOD(Close)();
			STDMETHOD(get_ActiveWorkBenchWindow)(BSTR bstrID, IWorkBenchWindow** pVal);
			STDMETHOD(InitLyncApp)();
			
			void __stdcall OnShutdown(IUccPlatform* pEventSource, IUccOperationProgressEvent* pEventData);
			void __stdcall OnStateChanged(IClient* _eventSource, IClientStateChangedEventData* _eventData);
			//CUccSessionManagerEvents£º
			HRESULT __stdcall OnIncomingSession(IUccEndpoint* pEventSource, IUccIncomingSessionEvent* pEventData);
			HRESULT __stdcall OnOutgoingSession(IUccEndpoint* pEventSource, IUccOutgoingSessionEvent* pEventData);
			//CLyncConversationManagerEvents:
			void __stdcall OnConversationAdded(IConversationManager* _eventSource, IConversationManagerEventData* _eventData);
			void __stdcall OnConversationRemoved(IConversationManager* _eventSource, IConversationManagerEventData* _eventData);
			//CLyncRoomManagerEvents:
			void __stdcall OnFollowedRoomAdded(IRoomManager* _eventSource, IFollowedRoomsChangedEventData* _eventData);
			void __stdcall OnFollowedRoomRemoved(IRoomManager* _eventSource, IFollowedRoomsChangedEventData* _eventData);
			void __stdcall OnRoomManagerStateChanged(IRoomManager* _eventSource, IRoomManagerStateChangedEventData* _eventData);
		protected:
			ULONG InternalAddRef() { return 1; }
			ULONG InternalRelease() { return 1; }
		};

		class CLyncAddin :
			public CTangram,
			public CWindowImpl<CTangram, CWindow>
		{
		public:
			CLyncAddin();
			virtual ~CLyncAddin();
			int										m_nRichEditCount;
			HWND									m_hMainWnd;
			HWND									m_hMainWnd2;
			HWND									m_hTabFrameWnd;

			//CTangram:
			void WindowDestroy(HWND hWnd);
			void WindowCreated(CString strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd);
			HRESULT COMObjCreated(REFCLSID rclsid, LPVOID pv);

			BEGIN_MSG_MAP(CLyncAddin)
				MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
				MESSAGE_HANDLER(WM_LYNCIMWNDCREATED, OnConsationWndCreated)
			END_MSG_MAP()
			LRESULT OnDestroy(UINT, WPARAM, LPARAM, BOOL&);
			LRESULT OnConsationWndCreated(UINT, WPARAM, LPARAM, BOOL&);
		};
	}
}


