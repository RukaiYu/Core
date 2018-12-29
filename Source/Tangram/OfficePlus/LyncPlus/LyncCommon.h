#pragma once
#include "..\Tangram\TangramCore.h"
#include "lync.h"
#include "uccapi.h"
#include "UccAPIEvent.h"
#include "LyncEvent.h"

using namespace UCCAPILib;
using namespace UCCollaborationLib;
using namespace OfficePlus::LyncPlus::UccApiEvent;
using namespace OfficePlus::LyncPlus::LyncClientEvent;

namespace OfficePlus
{
	namespace LyncPlus
	{
		class CLyncRoomObj : public CLyncRoomEvents
		{
		public:
			CLyncRoomObj();
			virtual ~CLyncRoomObj();

			IRoom*	m_pRoom;

			void __stdcall OnPropertyChanged(IRoom* _eventSource, IRoomPropertyChangedEventData* _eventData);
			void __stdcall OnUnreadMessageCountChanged(IRoom* _eventSource, IUnreadMessageCountChangedEventData* _eventData);
			void __stdcall OnIsSendingMessage(IRoom* _eventSource, IRoomMessageEventData* _eventData);
			void __stdcall OnMessagesReceived(IRoom* _eventSource, IRoomMessagesEventData* _eventData);
			void __stdcall OnParticipantAdded(IRoom* _eventSource, IRoomParticipantsEventData* _eventData);
			void __stdcall OnParticipantRemoved(IRoom* _eventSource, IRoomParticipantsEventData* _eventData);
			void __stdcall OnJoinStateChanged(IRoom* _eventSource, IRoomJoinStateChangedEventData* _eventData);
		};

		class CLyncConversationObj : public CLyncConversationEvents
		{
		public:
			CLyncConversationObj();
			virtual ~CLyncConversationObj();

			CString m_strConversationID;
			CString m_strConversationSubject;
			CString m_strConversationConferencingUri;
			IConversation*	m_pConversation;

			virtual void __stdcall OnStateChanged(IConversation* _eventSource, IConversationStateChangedEventData* _eventData);
			virtual void __stdcall OnPropertyChanged(IConversation* _eventSource, IConversationPropertyChangedEventData* _eventData);
			virtual void __stdcall OnParticipantAdded(IConversation* _eventSource, IParticipantCollectionChangedEventData* _eventData);
			virtual void __stdcall OnParticipantRemoved(IConversation* _eventSource, IParticipantCollectionChangedEventData* _eventData);
			virtual void __stdcall OnActionAvailabilityChanged(IConversation* _eventSource, IConversationActionAvailabilityEventData* _eventData);
			virtual void __stdcall OnConversationContextAdded(IConversation* _eventSource, IConversationContextCollectionEventData* _eventData);
			virtual void __stdcall OnConversationContextRemoved(IConversation* _eventSource, IConversationContextCollectionEventData* _eventData);
			virtual void __stdcall OnConversationContextLinkClicked(IConversation* _eventSource, IInitialContextEventData* _eventData);
			virtual void __stdcall OnInitialContextReceived(IConversation* _eventSource, IInitialContextEventData* _eventData);
			virtual void __stdcall OnInitialContextSent(IConversation* _eventSource, IInitialContextEventData* _eventData);
			virtual void __stdcall OnContextDataReceived(IConversation* _eventSource, IContextEventData* _eventData);
			virtual void __stdcall OnContextDataSent(IConversation* _eventSource, IContextEventData* _eventData);
		};


		class CLyncConversationWindowObj : public CLyncConversationWindowEvents, public CLyncConversationWindow2Events
		{
		public:
			CLyncConversationWindowObj();
			virtual ~CLyncConversationWindowObj();

			IConversationWindow* m_pConversationWindow;
			void __stdcall OnNeedsSizeChange(IConversationWindow* _eventSource, IConversationWindowNeedsSizeChangeEventData* _eventData);
			void __stdcall OnNeedsAttention(IConversationWindow* _eventSource, IConversationWindowNeedsAttentionEventData* _eventData);

			void __stdcall OnStateChanged(IConversationWindow* _eventSource, IConversationWindowStateChangedEventData* _eventData);
			void __stdcall OnInformationChanged(IConversationWindow* _eventSource, IConversationWindowInformationChangedEventData* _eventData);
			void __stdcall OnActionAvailabilityChanged(IConversationWindow* _eventSource, IConversationWindowActionAvailabilityChangedEventData* _eventData);
		};
	}
}


