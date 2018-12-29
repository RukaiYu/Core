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
#include "Lync.h"
using namespace UCCollaborationLib;

namespace OfficePlus
{
	namespace LyncPlus
	{
		namespace LyncClientEvent
		{
			extern _ATL_FUNC_INFO AccessEntryAdded;
			extern _ATL_FUNC_INFO AccessEntryRemoved;
			class CLyncAccessPermissionEvents : public IDispEventSimpleImpl< 1, CLyncAccessPermissionEvents, &__uuidof(_IAccessPermissionEvents)>
			{
			public:
				virtual void __stdcall OnAccessEntryAdded(IAccessPermission* _eventSource, IAccessEntryCollectionChangedEventData* _eventData) {}
				virtual void __stdcall OnAccessEntryRemoved(IAccessPermission* _eventSource, IAccessEntryCollectionChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncAccessPermissionEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IAccessPermissionEvents), 0x0000079e, OnAccessEntryAdded, &AccessEntryAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IAccessPermissionEvents), 0x0000079f, OnAccessEntryRemoved, &AccessEntryRemoved)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ModalityStateChanged;
			extern _ATL_FUNC_INFO ActionAvailabilityChanged;
			extern _ATL_FUNC_INFO LocalSharedResourcesChanged;
			extern _ATL_FUNC_INFO ParticipationStateChanged;
			extern _ATL_FUNC_INFO ControlRequestReceived;
			extern _ATL_FUNC_INFO ControllerChanged;
			extern _ATL_FUNC_INFO SharerChanged;
			class CLyncApplicationSharingModalityEvents : public IDispEventSimpleImpl< 1, CLyncApplicationSharingModalityEvents, &__uuidof(_IApplicationSharingModalityEvents)>
			{
			public:
				virtual void __stdcall OnModalityStateChanged(IModality* _eventSource, IModalityStateChangedEventData* _eventData) {}
				virtual void __stdcall OnActionAvailabilityChanged(IModality* _eventSource, IModalityActionAvailabilityChangedEventData* _eventData) {}
				virtual void __stdcall OnLocalSharedResourcesChanged(IApplicationSharingModality* eventSource, ILocalSharedResourcesChangedEventData* eventData) {}
				virtual void __stdcall OnParticipationStateChanged(IApplicationSharingModality* eventSource, IParticipationStateChangedEventData* eventData) {}
				virtual void __stdcall OnControlRequestReceived(IApplicationSharingModality* eventSource, IControlRequestReceivedEventData* eventData) {}
				virtual void __stdcall OnControllerChanged(IApplicationSharingModality* eventSource, IControllerChangedEventData* eventData) {}
				virtual void __stdcall OnSharerChanged(IApplicationSharingModality* eventSource, ISharerChangedEventData* eventData) {}

				BEGIN_SINK_MAP(CLyncApplicationSharingModalityEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IApplicationSharingModalityEvents), 0x000004b1, OnModalityStateChanged, &ModalityStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IApplicationSharingModalityEvents), 0x000004b4, OnActionAvailabilityChanged, &ActionAvailabilityChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IApplicationSharingModalityEvents), 0x00000598, OnLocalSharedResourcesChanged, &LocalSharedResourcesChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IApplicationSharingModalityEvents), 0x00000596, OnParticipationStateChanged, &ParticipationStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IApplicationSharingModalityEvents), 0x00000597, OnControlRequestReceived, &ControlRequestReceived)
					SINK_ENTRY_INFO(1, __uuidof(_IApplicationSharingModalityEvents), 0x00000599, OnControllerChanged, &ControllerChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IApplicationSharingModalityEvents), 0x0000059a, OnSharerChanged, &SharerChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ModalityStateChanged;
			extern _ATL_FUNC_INFO AVModalityPropertyChanged;
			extern _ATL_FUNC_INFO IsContributingChanged;
			extern _ATL_FUNC_INFO StreamStateChanged;
			extern _ATL_FUNC_INFO ActionAvailabilityChanged;
			class CLyncAVModalityEvents : public IDispEventSimpleImpl< 1, CLyncAVModalityEvents, &__uuidof(_IAVModalityEvents)>
			{
			public:
				virtual void __stdcall OnAVModalityStateChanged(IModality* _eventSource, IModalityStateChangedEventData* _eventData) {}
				virtual void __stdcall OnAVModalityPropertyChanged(IAVModality* _eventSource, IModalityPropertyChangedEventData* _eventData) {}
				virtual void __stdcall OnIsContributingChanged(IAVModality* _eventSource, IIsContributingChangedEventData* _eventData) {}
				virtual void __stdcall OnStreamStateChanged(IAVModality* _eventSource, IStreamStateChangedEventData* _eventData) {}
				virtual void __stdcall OnAVActionAvailabilityChanged(IModality* _eventSource, IModalityActionAvailabilityChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncAVModalityEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IAVModalityEvents), 0x000004b1, OnAVModalityStateChanged, &ModalityStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IAVModalityEvents), 0x000004b0, OnAVModalityPropertyChanged, &AVModalityPropertyChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IAVModalityEvents), 0x000004e2, OnIsContributingChanged, &IsContributingChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IAVModalityEvents), 0x000004e3, OnStreamStateChanged, &StreamStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IAVModalityEvents), 0x000004b4, OnAVActionAvailabilityChanged, &ActionAvailabilityChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO StateChanged;
			extern _ATL_FUNC_INFO ActionAvailabilityChanged;
			class CLyncChannelEvents : public IDispEventSimpleImpl< 1, CLyncChannelEvents, &__uuidof(_IChannelEvents)>
			{
			public:
				virtual void __stdcall OnStateChanged(IChannel* _eventSource, IChannelStateChangedEventData* _eventData) {}
				virtual void __stdcall OnActionAvailabilityChanged(IChannel* _eventSource, IChannelActionAvailabilityEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncChannelEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IChannelEvents), 0x00000514, OnStateChanged, &StateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IChannelEvents), 0x00000515, OnActionAvailabilityChanged, &ActionAvailabilityChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO StateChanged;
			class CLyncClientEvents : public IDispEventSimpleImpl< 1, CLyncClientEvents, &__uuidof(_IClientEvents)>
			{
			public:
				virtual void __stdcall OnStateChanged(IClient* _eventSource, IClientStateChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncClientEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IClientEvents), 0x00000064, OnStateChanged, &StateChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO SettingChanged;
			class CLyncClientSettingsEvents : public IDispEventSimpleImpl< 1, CLyncClientSettingsEvents, &__uuidof(_IClientSettingsEvents)>
			{
			public:
				virtual void __stdcall OnSettingChanged(IClientSettings* _eventSource, IClientSettingsChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncClientSettingsEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IClientSettingsEvents), 0x0000141e, OnSettingChanged, &SettingChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ContactInformationChanged;
			extern _ATL_FUNC_INFO SettingChanged;
			extern _ATL_FUNC_INFO UriChanged;
			class CLyncContactEvents : public IDispEventSimpleImpl< 1, CLyncContactEvents, &__uuidof(_IContactEvents)>
			{
			public:
				virtual void __stdcall OnContactInformationChanged(IContact* _eventSource, IContactInformationChangedEventData* _eventData) {}
				virtual void __stdcall OnSettingChanged(IContact* _eventSource, IContactSettingChangedEventData* _eventData) {}
				virtual void __stdcall OnUriChanged(IContact* _eventSource, IUriChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncContactEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IContactEvents), 0x00000190, OnContactInformationChanged, &ContactInformationChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IContactEvents), 0x00000191, OnSettingChanged, &SettingChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IContactEvents), 0x00000192, OnUriChanged, &UriChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO GroupAdded;
			extern _ATL_FUNC_INFO GroupRemoved;
			extern _ATL_FUNC_INFO SearchProviderStateChanged;
			class CLyncContactManagerEvents : public IDispEventSimpleImpl< 1, CLyncContactManagerEvents, &__uuidof(_IContactManagerEvents)>
			{
			public:
				virtual void __stdcall OnGroupAdded(IContactManager* _eventSource, IGroupCollectionChangedEventData* _eventData) {}
				virtual void __stdcall OnGroupRemoved(IContactManager* _eventSource, IGroupCollectionChangedEventData* _eventData) {}
				virtual void __stdcall OnSearchProviderStateChanged(IContactManager* _eventSource, ISearchProviderStateChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncContactManagerEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IContactManagerEvents), 0x000000c8, OnGroupAdded, &GroupAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IContactManagerEvents), 0x000000c9, OnGroupRemoved, &GroupRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IContactManagerEvents), 0x000000ca, OnSearchProviderStateChanged, &SearchProviderStateChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ModalityStateChanged;
			extern _ATL_FUNC_INFO ActionAvailabilityChanged;
			extern _ATL_FUNC_INFO ContentAdded;
			extern _ATL_FUNC_INFO ContentRemoved;
			extern _ATL_FUNC_INFO ActivePresenterChanged;
			extern _ATL_FUNC_INFO ActiveContentChanged;
			class CLyncContentSharingModalityEvents : public IDispEventSimpleImpl< 1, CLyncContentSharingModalityEvents, &__uuidof(_IContentSharingModalityEvents)>
			{
			public:
				virtual void __stdcall OnModalityStateChanged(IModality* _eventSource, IModalityStateChangedEventData* _eventData) {}
				virtual void __stdcall OnActionAvailabilityChanged(IModality* _eventSource, IModalityActionAvailabilityChangedEventData* _eventData) {}
				virtual void __stdcall OnContentAdded(IContentSharingModality* _eventSource, IContentCollectionChangedEventData* _eventData) {}
				virtual void __stdcall OnContentRemoved(IContentSharingModality* _eventSource, IContentCollectionChangedEventData* _eventData) {}
				virtual void __stdcall OnActivePresenterChanged(IContentSharingModality* _eventSource, IActivePresenterChangedEventData* _eventData) {}
				virtual void __stdcall OnActiveContentChanged(IContentSharingModality* _eventSource, IActiveContentChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncContentSharingModalityEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IContentSharingModalityEvents), 0x000004b1, OnModalityStateChanged, &ModalityStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IContentSharingModalityEvents), 0x000004b4, OnActionAvailabilityChanged, &ActionAvailabilityChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IContentSharingModalityEvents), 0x000013d8, OnContentAdded, &ContentAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IContentSharingModalityEvents), 0x000013d9, OnContentRemoved, &ContentRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IContentSharingModalityEvents), 0x000013da, OnActivePresenterChanged, &ActivePresenterChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IContentSharingModalityEvents), 0x000013db, OnActiveContentChanged, &ActiveContentChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO StateChanged;
			extern _ATL_FUNC_INFO PropertyChanged;
			extern _ATL_FUNC_INFO ParticipantAdded;
			extern _ATL_FUNC_INFO ParticipantRemoved;
			extern _ATL_FUNC_INFO ActionAvailabilityChanged;
			extern _ATL_FUNC_INFO ConversationContextAdded;
			extern _ATL_FUNC_INFO ConversationContextRemoved;
			extern _ATL_FUNC_INFO ConversationContextLinkClicked;
			extern _ATL_FUNC_INFO InitialContextReceived;
			extern _ATL_FUNC_INFO InitialContextSent;
			extern _ATL_FUNC_INFO ContextDataReceived;
			extern _ATL_FUNC_INFO ContextDataSent;
			class CLyncConversationEvents : public IDispEventSimpleImpl< 1, CLyncConversationEvents, &__uuidof(_IConversationEvents)>
			{
			public:
				virtual void __stdcall OnStateChanged(IConversation* _eventSource, IConversationStateChangedEventData* _eventData) {}
				virtual void __stdcall OnPropertyChanged(IConversation* _eventSource, IConversationPropertyChangedEventData* _eventData) {}
				virtual void __stdcall OnParticipantAdded(IConversation* _eventSource, IParticipantCollectionChangedEventData* _eventData) {}
				virtual void __stdcall OnParticipantRemoved(IConversation* _eventSource, IParticipantCollectionChangedEventData* _eventData) {}
				virtual void __stdcall OnActionAvailabilityChanged(IConversation* _eventSource, IConversationActionAvailabilityEventData* _eventData) {}
				virtual void __stdcall OnConversationContextAdded(IConversation* _eventSource, IConversationContextCollectionEventData* _eventData) {}
				virtual void __stdcall OnConversationContextRemoved(IConversation* _eventSource, IConversationContextCollectionEventData* _eventData) {}
				virtual void __stdcall OnConversationContextLinkClicked(IConversation* _eventSource, IInitialContextEventData* _eventData) {}
				virtual void __stdcall OnInitialContextReceived(IConversation* _eventSource, IInitialContextEventData* _eventData) {}
				virtual void __stdcall OnInitialContextSent(IConversation* _eventSource, IInitialContextEventData* _eventData) {}
				virtual void __stdcall OnContextDataReceived(IConversation* _eventSource, IContextEventData* _eventData) {}
				virtual void __stdcall OnContextDataSent(IConversation* _eventSource, IContextEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncConversationEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003e8, OnStateChanged, &StateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003e9, OnPropertyChanged, &PropertyChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003eb, OnParticipantAdded, &ParticipantAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003ec, OnParticipantRemoved, &ParticipantRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003ea, OnActionAvailabilityChanged, &ActionAvailabilityChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003ed, OnConversationContextAdded, &ConversationContextAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003ee, OnConversationContextRemoved, &ConversationContextRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003f1, OnConversationContextLinkClicked, &ConversationContextLinkClicked)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003ef, OnInitialContextReceived, &InitialContextReceived)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003f2, OnInitialContextSent, &InitialContextSent)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003f0, OnContextDataReceived, &ContextDataReceived)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationEvents), 0x000003f3, OnContextDataSent, &ContextDataSent)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO HistoryItemAdded;
			extern _ATL_FUNC_INFO PreviousHistoryRetrievalStateChanged;
			extern _ATL_FUNC_INFO HistoryActionAvailabilityChanged;


			class CLyncConversationHistoryEvents : public IDispEventSimpleImpl< 1, CLyncConversationHistoryEvents, &__uuidof(_IConversationHistoryEvents)>
			{
			public:
				virtual void __stdcall OnHistoryItemAdded(IConversationHistory* _eventSource, IConversationHistoryItemAddedEventData* _eventData) {}
				virtual void __stdcall OnPreviousHistoryRetrievalStateChanged(IConversationHistory* _eventSource, IPreviousHistoryRetrievalStateChangedEventData* _eventData) {}
				virtual void __stdcall OnHistoryActionAvailabilityChanged(IConversationHistory* _eventSource, IConversationHistoryActionAvailabilityEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncConversationHistoryEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationHistoryEvents), 0x000007d0, OnHistoryItemAdded, &HistoryItemAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationHistoryEvents), 0x000007d1, OnPreviousHistoryRetrievalStateChanged, &PreviousHistoryRetrievalStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationHistoryEvents), 0x000007d2, OnHistoryActionAvailabilityChanged, &HistoryActionAvailabilityChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ConversationAdded;
			extern _ATL_FUNC_INFO ConversationRemoved;
			extern _ATL_FUNC_INFO SummarizerStatesChanged;


			class CLyncConversationManagerEvents : public IDispEventSimpleImpl< 1, CLyncConversationManagerEvents, &__uuidof(_IConversationManagerEvents)>
			{
			public:
				virtual void __stdcall OnConversationAdded(IConversationManager* _eventSource, IConversationManagerEventData* _eventData) {}
				virtual void __stdcall OnConversationRemoved(IConversationManager* _eventSource, IConversationManagerEventData* _eventData) {}
				virtual void __stdcall OnSummarizerStatesChanged(IConversationManager* _eventSource, ISummarizerStatesChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncConversationManagerEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationManagerEvents), 0x00000384, OnConversationAdded, &ConversationAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationManagerEvents), 0x00000385, OnConversationRemoved, &ConversationRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationManagerEvents), 0x00000386, OnSummarizerStatesChanged, &SummarizerStatesChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO StateChanged;
			extern _ATL_FUNC_INFO InformationChanged;
			extern _ATL_FUNC_INFO ActionAvailabilityChanged;
			class CLyncConversationWindow2Events : public IDispEventSimpleImpl< 1, CLyncConversationWindow2Events, &__uuidof(_IConversationWindow2Events)>
			{
			public:
				virtual void __stdcall OnStateChanged(IConversationWindow* _eventSource, IConversationWindowStateChangedEventData* _eventData) {}
				virtual void __stdcall OnInformationChanged(IConversationWindow* _eventSource, IConversationWindowInformationChangedEventData* _eventData) {}
				virtual void __stdcall OnActionAvailabilityChanged(IConversationWindow* _eventSource, IConversationWindowActionAvailabilityChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncConversationWindow2Events)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationWindow2Events), 0x000013ba, OnStateChanged, &StateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationWindow2Events), 0x000013bb, OnInformationChanged, &InformationChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationWindow2Events), 0x000013bc, OnActionAvailabilityChanged, &ActionAvailabilityChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO NeedsSizeChange;
			extern _ATL_FUNC_INFO NeedsAttention;
			class CLyncConversationWindowEvents : public IDispEventSimpleImpl< 1, CLyncConversationWindowEvents, &__uuidof(_IConversationWindowEvents)>
			{
			public:
				virtual void __stdcall OnNeedsSizeChange(IConversationWindow* _eventSource, IConversationWindowNeedsSizeChangeEventData* _eventData) {}
				virtual void __stdcall OnNeedsAttention(IConversationWindow* _eventSource, IConversationWindowNeedsAttentionEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncConversationWindowEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationWindowEvents), 0x00000fa0, OnNeedsSizeChange, &NeedsSizeChange)
					SINK_ENTRY_INFO(1, __uuidof(_IConversationWindowEvents), 0x00000fa1, OnNeedsAttention, &NeedsAttention)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO NameChanged;
			extern _ATL_FUNC_INFO ContactAdded;
			extern _ATL_FUNC_INFO ContactRemoved;
			class CLyncCustomGroupEvents : public IDispEventSimpleImpl< 1, CLyncCustomGroupEvents, &__uuidof(_ICustomGroupEvents)>
			{
			public:
				virtual void __stdcall OnNameChanged(IGroup* _eventSource, IGroupNameChangedEventData* _eventData) {}
				virtual void __stdcall OnContactAdded(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}
				virtual void __stdcall OnContactRemoved(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncCustomGroupEvents)
					SINK_ENTRY_INFO(1, __uuidof(_ICustomGroupEvents), 0x000001f4, OnNameChanged, &NameChanged)
					SINK_ENTRY_INFO(1, __uuidof(_ICustomGroupEvents), 0x000001f5, OnContactAdded, &ContactAdded)
					SINK_ENTRY_INFO(1, __uuidof(_ICustomGroupEvents), 0x000001f6, OnContactRemoved, &ContactRemoved)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO StateChanged;
			class CLyncDelegatorClientEvents : public IDispEventSimpleImpl< 1, CLyncDelegatorClientEvents, &__uuidof(_IDelegatorClientEvents)>
			{
			public:
				virtual void __stdcall OnStateChanged(IClient* _eventSource, IClientStateChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncDelegatorClientEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IDelegatorClientEvents), 0x00000064, OnStateChanged, &StateChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO NameChanged;
			extern _ATL_FUNC_INFO ContactAdded;
			extern _ATL_FUNC_INFO ContactRemoved;
			extern _ATL_FUNC_INFO ExpansionStateChanged;
			extern _ATL_FUNC_INFO NestedGroupAdded;
			extern _ATL_FUNC_INFO NestedGroupRemoved;
			extern _ATL_FUNC_INFO DescriptionChanged;
			class CLyncDistributionGroupEvents : public IDispEventSimpleImpl< 1, CLyncDistributionGroupEvents, &__uuidof(_IDistributionGroupEvents)>
			{
			public:
				virtual void __stdcall OnNameChanged(IGroup* _eventSource, IGroupNameChangedEventData* _eventData) {}
				virtual void __stdcall OnContactAdded(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}
				virtual void __stdcall OnContactRemoved(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}
				virtual void __stdcall OnExpansionStateChanged(IDistributionGroup* _eventSource, IDistributionGroupExpansionStateChangedEventData* _eventData) {}
				virtual void __stdcall OnNestedGroupAdded(IDistributionGroup* _eventSource, IGroupCollectionChangedEventData* _eventData) {}
				virtual void __stdcall OnNestedGroupRemoved(IDistributionGroup* _eventSource, IGroupCollectionChangedEventData* _eventData) {}
				virtual void __stdcall OnDescriptionChanged(IDistributionGroup* _eventSource, IDescriptionChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncDistributionGroupEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IDistributionGroupEvents), 0x000001f4, OnNameChanged, &NameChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IDistributionGroupEvents), 0x000001f5, OnContactAdded, &ContactAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IDistributionGroupEvents), 0x000001f6, OnContactRemoved, &ContactRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IDistributionGroupEvents), 0x000001f7, OnExpansionStateChanged, &ExpansionStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IDistributionGroupEvents), 0x000001f9, OnNestedGroupAdded, &NestedGroupAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IDistributionGroupEvents), 0x000001fa, OnNestedGroupRemoved, &NestedGroupRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IDistributionGroupEvents), 0x000001fb, OnDescriptionChanged, &DescriptionChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO NameChanged;
			extern _ATL_FUNC_INFO ContactAdded;
			extern _ATL_FUNC_INFO ContactRemoved;
			class CLyncFavoriteContactsEvents : public IDispEventSimpleImpl< 1, CLyncFavoriteContactsEvents, &__uuidof(_IFavoriteContactsEvents)>
			{
			public:
				virtual void __stdcall OnNameChanged(IGroup* _eventSource, IGroupNameChangedEventData* _eventData) {}
				virtual void __stdcall OnContactAdded(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}
				virtual void __stdcall OnContactRemoved(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncFavoriteContactsEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IFavoriteContactsEvents), 0x000001f4, OnNameChanged, &NameChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IFavoriteContactsEvents), 0x000001f5, OnContactAdded, &ContactAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IFavoriteContactsEvents), 0x000001f6, OnContactRemoved, &ContactRemoved)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO NameChanged;
			extern _ATL_FUNC_INFO ContactAdded;
			extern _ATL_FUNC_INFO ContactRemoved;
			extern _ATL_FUNC_INFO ContactPositionChanged;
			class CLyncFrequentContactsEvents : public IDispEventSimpleImpl< 1, CLyncFrequentContactsEvents, &__uuidof(_IFrequentContactsEvents)>
			{
			public:
				virtual void __stdcall OnNameChanged(IGroup* _eventSource, IGroupNameChangedEventData* _eventData) {}
				virtual void __stdcall OnContactAdded(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}
				virtual void __stdcall OnContactRemoved(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}
				virtual void __stdcall OnContactPositionChanged(IFrequentContacts* _eventSource, IContactPositionChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncFrequentContactsEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IFrequentContactsEvents), 0x000001f4, OnNameChanged, &NameChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IFrequentContactsEvents), 0x000001f5, OnContactAdded, &ContactAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IFrequentContactsEvents), 0x000001f6, OnContactRemoved, &ContactRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IFrequentContactsEvents), 0x000001f8, OnContactPositionChanged, &ContactPositionChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO NameChanged;
			extern _ATL_FUNC_INFO ContactAdded;
			extern _ATL_FUNC_INFO ContactRemoved;
			class CLyncGroupEvents : public IDispEventSimpleImpl< 1, CLyncGroupEvents, &__uuidof(_IGroupEvents)>
			{
			public:
				virtual void __stdcall OnNameChanged(IGroup* _eventSource, IGroupNameChangedEventData* _eventData) {}
				virtual void __stdcall OnContactAdded(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}
				virtual void __stdcall OnContactRemoved(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncGroupEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IGroupEvents), 0x000001f4, OnNameChanged, &NameChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IGroupEvents), 0x000001f5, OnContactAdded, &ContactAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IGroupEvents), 0x000001f6, OnContactRemoved, &ContactRemoved)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ModalityStateChanged;
			extern _ATL_FUNC_INFO InstantMessagePropertyChanged;
			extern _ATL_FUNC_INFO InstantMessageReceived;
			extern _ATL_FUNC_INFO IsTypingChanged;
			extern _ATL_FUNC_INFO ActionAvailabilityChanged;
			class CLyncInstantMessageModalityEvents : public IDispEventSimpleImpl< 1, CLyncInstantMessageModalityEvents, &__uuidof(_IInstantMessageModalityEvents)>
			{
			public:
				virtual void __stdcall OnModalityStateChanged(IModality* _eventSource, IModalityStateChangedEventData* _eventData) {}
				virtual void __stdcall OnInstantMessagePropertyChanged(IInstantMessageModality* _eventSource, IModalityPropertyChangedEventData* _eventData) {}
				virtual void __stdcall OnInstantMessageReceived(IInstantMessageModality* _eventSource, IMessageSentEventData* _eventData) {}
				virtual void __stdcall OnIsTypingChanged(IInstantMessageModality* _eventSource, IIsTypingChangedEventData* _eventData) {}
				virtual void __stdcall OnActionAvailabilityChanged(IModality* _eventSource, IModalityActionAvailabilityChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncInstantMessageModalityEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IInstantMessageModalityEvents), 0x000004b1, OnModalityStateChanged, &ModalityStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IInstantMessageModalityEvents), 0x000004b0, OnInstantMessagePropertyChanged, &InstantMessagePropertyChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IInstantMessageModalityEvents), 0x000004b2, OnInstantMessageReceived, &InstantMessageReceived)
					SINK_ENTRY_INFO(1, __uuidof(_IInstantMessageModalityEvents), 0x000004b3, OnIsTypingChanged, &IsTypingChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IInstantMessageModalityEvents), 0x000004b4, OnActionAvailabilityChanged, &ActionAvailabilityChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO StateChanged;
			extern _ATL_FUNC_INFO NotificationReceived;
			extern _ATL_FUNC_INFO CredentialRequested;
			extern _ATL_FUNC_INFO SignInDelayed;
			extern _ATL_FUNC_INFO CapabilitiesChanged;
			extern _ATL_FUNC_INFO DelegatorClientAdded;
			extern _ATL_FUNC_INFO DelegatorClientRemoved;
			class CTangramLyncClientEvents : public IDispEventSimpleImpl< 1, CTangramLyncClientEvents, &__uuidof(_ILyncClientEvents)>
			{
			public:
				virtual void __stdcall OnStateChanged(IClient* _eventSource, IClientStateChangedEventData* _eventData) {}
				virtual void __stdcall OnNotificationReceived(ILyncClient* _eventSource, ILyncClientNotificationReceivedEventData* _eventData) {}
				virtual void __stdcall OnCredentialRequested(ILyncClient* _eventSource, ICredentialRequestedEventData* _eventData) {}
				virtual void __stdcall OnSignInDelayed(ILyncClient* _eventSource, ISignInDelayedEventData* _eventData) {}
				virtual void __stdcall OnCapabilitiesChanged(ILyncClient* _eventSource, IPreferredCapabilitiesChangedEventData* _eventData) {}
				virtual void __stdcall OnDelegatorClientAdded(ILyncClient* _eventSource, IDelegatorClientCollectionEventData* _eventData) {}
				virtual void __stdcall OnDelegatorClientRemoved(ILyncClient* _eventSource, IDelegatorClientCollectionEventData* _eventData) {}

				BEGIN_SINK_MAP(CTangramLyncClientEvents)
					SINK_ENTRY_INFO(1, __uuidof(_ILyncClientEvents), 0x00000064, OnStateChanged, &StateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_ILyncClientEvents), 0x00000065, OnNotificationReceived, &NotificationReceived)
					SINK_ENTRY_INFO(1, __uuidof(_ILyncClientEvents), 0x00000067, OnCredentialRequested, &CredentialRequested)
					SINK_ENTRY_INFO(1, __uuidof(_ILyncClientEvents), 0x00000066, OnSignInDelayed, &SignInDelayed)
					SINK_ENTRY_INFO(1, __uuidof(_ILyncClientEvents), 0x00000069, OnCapabilitiesChanged, &CapabilitiesChanged)
					SINK_ENTRY_INFO(1, __uuidof(_ILyncClientEvents), 0x0000006a, OnDelegatorClientAdded, &DelegatorClientAdded)
					SINK_ENTRY_INFO(1, __uuidof(_ILyncClientEvents), 0x0000006b, OnDelegatorClientRemoved, &DelegatorClientRemoved)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ModalityStateChanged;
			extern _ATL_FUNC_INFO ActionAvailabilityChanged;
			class CLyncModalityEvents : public IDispEventSimpleImpl< 1, CLyncModalityEvents, &__uuidof(_IModalityEvents)>
			{
			public:
				virtual void __stdcall OnModalityStateChanged(IModality* _eventSource, IModalityStateChangedEventData* _eventData) {}
				virtual void __stdcall OnActionAvailabilityChanged(IModality* _eventSource, IModalityActionAvailabilityChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncModalityEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IModalityEvents), 0x000004b1, OnModalityStateChanged, &ModalityStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IModalityEvents), 0x000004b4, OnActionAvailabilityChanged, &ActionAvailabilityChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO PropertyChanged;
			extern _ATL_FUNC_INFO IsMutedChanged;
			extern _ATL_FUNC_INFO ActionAvailabilityChanged;
			class CLyncParticipantEvents : public IDispEventSimpleImpl< 1, CLyncParticipantEvents, &__uuidof(_IParticipantEvents)>
			{
			public:
				virtual void __stdcall OnPropertyChanged(IParticipant* _eventSource, IParticipantPropertyChangedEventData* _eventData) {}
				virtual void __stdcall OnIsMutedChanged(IParticipant* _eventSource, IMutedChangedEventData* _eventData) {}
				virtual void __stdcall OnActionAvailabilityChanged(IParticipant* _eventSource, IParticipantActionAvailabilityChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncParticipantEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IParticipantEvents), 0x0000044c, OnPropertyChanged, &PropertyChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IParticipantEvents), 0x0000044d, OnIsMutedChanged, &IsMutedChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IParticipantEvents), 0x0000044e, OnActionAvailabilityChanged, &ActionAvailabilityChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ConnectionStateChanged;
			extern _ATL_FUNC_INFO NewItemCountSinceLastResetChanged;
			extern _ATL_FUNC_INFO ItemsAdded;
			extern _ATL_FUNC_INFO ItemsRemoved;
			extern _ATL_FUNC_INFO ItemsModified;
			class CLyncPreviousConversationsManagerEvents : public IDispEventSimpleImpl< 1, CLyncPreviousConversationsManagerEvents, &__uuidof(_IPreviousConversationsManagerEvents)>
			{
			public:
				virtual void __stdcall OnConnectionStateChanged(IPreviousConversationsManager* _eventSource, IPreviousConversationsManagerConnectionStateChangedEventData* _eventData) {}
				virtual void __stdcall OnNewItemCountSinceLastResetChanged(IPreviousConversationsManager* _eventSource, IPreviousConversationsManagerNewItemCountChangedEventData* _eventData) {}
				virtual void __stdcall OnItemsAdded(IPreviousConversationsManager* _eventSource, IPreviousConversationBatchEventData* _eventData) {}
				virtual void __stdcall OnItemsRemoved(IPreviousConversationsManager* _eventSource, IPreviousConversationBatchEventData* _eventData) {}
				virtual void __stdcall OnItemsModified(IPreviousConversationsManager* _eventSource, IPreviousConversationBatchEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncPreviousConversationsManagerEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IPreviousConversationsManagerEvents), 0x00000834, OnConnectionStateChanged, &ConnectionStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IPreviousConversationsManagerEvents), 0x00000835, OnNewItemCountSinceLastResetChanged, &NewItemCountSinceLastResetChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IPreviousConversationsManagerEvents), 0x00000837, OnItemsAdded, &ItemsAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IPreviousConversationsManagerEvents), 0x00000838, OnItemsRemoved, &ItemsRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IPreviousConversationsManagerEvents), 0x00000839, OnItemsModified, &ItemsModified)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO NameChanged;
			extern _ATL_FUNC_INFO ContactAdded;
			extern _ATL_FUNC_INFO ContactRemoved;
			class CLyncRecentContactsEvents : public IDispEventSimpleImpl< 1, CLyncRecentContactsEvents, &__uuidof(_IRecentContactsEvents)>
			{
			public:
				virtual void __stdcall OnNameChanged(IGroup* _eventSource, IGroupNameChangedEventData* _eventData) {}
				virtual void __stdcall OnContactAdded(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}
				virtual void __stdcall OnContactRemoved(IGroup* _eventSource, IGroupMemberChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncRecentContactsEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IRecentContactsEvents), 0x000001f4, OnNameChanged, &NameChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IRecentContactsEvents), 0x000001f5, OnContactAdded, &ContactAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IRecentContactsEvents), 0x000001f6, OnContactRemoved, &ContactRemoved)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO PropertyChanged;
			extern _ATL_FUNC_INFO UnreadMessageCountChanged;
			extern _ATL_FUNC_INFO IsSendingMessage;
			extern _ATL_FUNC_INFO MessagesReceived;
			extern _ATL_FUNC_INFO ParticipantAdded;
			extern _ATL_FUNC_INFO ParticipantRemoved;
			extern _ATL_FUNC_INFO JoinStateChanged;
			class CLyncRoomEvents : public IDispEventSimpleImpl< 1, CLyncRoomEvents, &__uuidof(_IRoomEvents)>
			{
			public:
				virtual void __stdcall OnPropertyChanged(IRoom* _eventSource, IRoomPropertyChangedEventData* _eventData) {}
				virtual void __stdcall OnUnreadMessageCountChanged(IRoom* _eventSource, IUnreadMessageCountChangedEventData* _eventData) {}
				virtual void __stdcall OnIsSendingMessage(IRoom* _eventSource, IRoomMessageEventData* _eventData) {}
				virtual void __stdcall OnMessagesReceived(IRoom* _eventSource, IRoomMessagesEventData* _eventData) {}
				virtual void __stdcall OnParticipantAdded(IRoom* _eventSource, IRoomParticipantsEventData* _eventData) {}
				virtual void __stdcall OnParticipantRemoved(IRoom* _eventSource, IRoomParticipantsEventData* _eventData) {}
				virtual void __stdcall OnJoinStateChanged(IRoom* _eventSource, IRoomJoinStateChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncRoomEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomEvents), 0x00001068, OnPropertyChanged, &PropertyChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomEvents), 0x00001069, OnUnreadMessageCountChanged, &UnreadMessageCountChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomEvents), 0x0000106a, OnIsSendingMessage, &IsSendingMessage)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomEvents), 0x0000106b, OnMessagesReceived, &MessagesReceived)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomEvents), 0x0000106c, OnParticipantAdded, &ParticipantAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomEvents), 0x0000106d, OnParticipantRemoved, &ParticipantRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomEvents), 0x0000106e, OnJoinStateChanged, &JoinStateChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO FollowedRoomAdded;
			extern _ATL_FUNC_INFO FollowedRoomRemoved;
			extern _ATL_FUNC_INFO RoomManagerStateChanged;
			class CLyncRoomManagerEvents : public IDispEventSimpleImpl< 1, CLyncRoomManagerEvents, &__uuidof(_IRoomManagerEvents)>
			{
			public:
				virtual void __stdcall OnFollowedRoomAdded(IRoomManager* _eventSource, IFollowedRoomsChangedEventData* _eventData) {}
				virtual void __stdcall OnFollowedRoomRemoved(IRoomManager* _eventSource, IFollowedRoomsChangedEventData* _eventData) {}
				virtual void __stdcall OnRoomManagerStateChanged(IRoomManager* _eventSource, IRoomManagerStateChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncRoomManagerEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomManagerEvents), 0x0000107c, OnFollowedRoomAdded, &FollowedRoomAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomManagerEvents), 0x0000107d, OnFollowedRoomRemoved, &FollowedRoomRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomManagerEvents), 0x0000107e, OnRoomManagerStateChanged, &RoomManagerStateChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO CanSendMessageChanged;
			class CLyncRoomUserEvents : public IDispEventSimpleImpl< 1, CLyncRoomUserEvents, &__uuidof(_IRoomUserEvents)>
			{
			public:
				virtual void __stdcall OnCanSendMessageChanged(IRoomUser* _eventSource, IRoomUserChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncRoomUserEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IRoomUserEvents), 0x00001090, OnCanSendMessageChanged, &CanSendMessageChanged)
				END_SINK_MAP()
			};


			extern _ATL_FUNC_INFO AlertLevelChanged;
			class CLyncSelf2Events : public IDispEventSimpleImpl< 1, CLyncSelf2Events, &__uuidof(_ISelf2Events)>
			{
			public:
				virtual void __stdcall OnAlertLevelChanged(ISelf* _eventSource, IAlertLevelChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncSelf2Events)
					SINK_ENTRY_INFO(1, __uuidof(_ISelf2Events), 0x0000139c, OnAlertLevelChanged, &AlertLevelChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO PhonesChanged;
			extern _ATL_FUNC_INFO ResiliencyModeChanged;
			class CLyncSelfEvents : public IDispEventSimpleImpl< 1, CLyncSelfEvents, &__uuidof(_ISelfEvents)>
			{
			public:
				virtual void __stdcall OnPhonesChanged(ISelf* _eventSource, IPhonesChangedEventData* _eventData) {}
				virtual void __stdcall OnResiliencyModeChanged(ISelf* _eventSource, IResiliencyModeChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncSelfEvents)
					SINK_ENTRY_INFO(1, __uuidof(_ISelfEvents), 0x000006a4, OnPhonesChanged, &PhonesChanged)
					SINK_ENTRY_INFO(1, __uuidof(_ISelfEvents), 0x000006a5, OnResiliencyModeChanged, &ResiliencyModeChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO StateChanged;
			extern _ATL_FUNC_INFO PropertyChanged;
			extern _ATL_FUNC_INFO ActionAvailabilityChanged;
			class CLyncShareableContentEvents : public IDispEventSimpleImpl< 1, CLyncShareableContentEvents, &__uuidof(_IShareableContentEvents)>
			{
			public:
				virtual void __stdcall OnStateChanged(IShareableContent* _eventSource, IShareableContentStateChangedEventData* _eventData) {}
				virtual void __stdcall OnPropertyChanged(IShareableContent* _eventSource, IShareableContentPropertyChangedEventData* _eventData) {}
				virtual void __stdcall OnActionAvailabilityChanged(IShareableContent* _eventSource, IShareableContentActionAvailabilityChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncShareableContentEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IShareableContentEvents), 0x000013e2, OnStateChanged, &StateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IShareableContentEvents), 0x000013e3, OnPropertyChanged, &PropertyChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IShareableContentEvents), 0x000013e4, OnActionAvailabilityChanged, &ActionAvailabilityChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ShuttingDown;
			class CLyncUCOfficeIntegrationEvents : public IDispEventSimpleImpl< 1, CLyncUCOfficeIntegrationEvents, &__uuidof(_IUCOfficeIntegrationEvents)>
			{
			public:
				virtual void __stdcall OnShuttingDown() {}


				BEGIN_SINK_MAP(CLyncUCOfficeIntegrationEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IUCOfficeIntegrationEvents), 0x00000032, OnShuttingDown, &ShuttingDown)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ExchangeECPUrlChanged;
			class CLyncUtilities2Events : public IDispEventSimpleImpl< 1, CLyncUtilities2Events, &__uuidof(_IUtilities2Events)>
			{
			public:
				virtual void __stdcall OnExchangeECPUrlChanged(IUtilities* _eventSource, IExchangeECPUrlChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncUtilities2Events)
					SINK_ENTRY_INFO(1, __uuidof(_IUtilities2Events), 0x000013c4, OnExchangeECPUrlChanged, &ExchangeECPUrlChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO InformationChanged;
			class CLyncVideoViewEvents : public IDispEventSimpleImpl< 1, CLyncVideoViewEvents, &__uuidof(_IVideoViewEvents)>
			{
			public:
				virtual void __stdcall OnInformationChanged(IVideoView* _eventSource, IVideoViewInformationChangedEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncVideoViewEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IVideoViewEvents), 0x00001400, OnInformationChanged, &InformationChanged)
				END_SINK_MAP()
			};

			extern _ATL_FUNC_INFO ConnectionStateChanged;
			extern _ATL_FUNC_INFO NewVoiceMailCountSinceLastResetChanged;
			extern _ATL_FUNC_INFO VoiceMailsAdded;
			extern _ATL_FUNC_INFO VoiceMailsRemoved;
			extern _ATL_FUNC_INFO VoiceMailsModified;
			class CLyncVoiceMailsManagerEvents : public IDispEventSimpleImpl< 1, CLyncVoiceMailsManagerEvents, &__uuidof(_IVoiceMailsManagerEvents)>
			{
			public:
				virtual void __stdcall OnConnectionStateChanged(IVoiceMailsManager* _eventSource, IPreviousConversationsManagerConnectionStateChangedEventData* _eventData) {}
				virtual void __stdcall OnNewVoiceMailCountSinceLastResetChanged(IVoiceMailsManager* _eventSource, IPreviousConversationsManagerNewItemCountChangedEventData* _eventData) {}
				virtual void __stdcall OnVoiceMailsAdded(IVoiceMailsManager* _eventSource, IVoiceMailBatchEventData* _eventData) {}
				virtual void __stdcall OnVoiceMailsRemoved(IVoiceMailsManager* _eventSource, IVoiceMailBatchEventData* _eventData) {}
				virtual void __stdcall OnVoiceMailsModified(IVoiceMailsManager* _eventSource, IVoiceMailBatchEventData* _eventData) {}

				BEGIN_SINK_MAP(CLyncVoiceMailsManagerEvents)
					SINK_ENTRY_INFO(1, __uuidof(_IVoiceMailsManagerEvents), 0x00000866, OnConnectionStateChanged, &ConnectionStateChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IVoiceMailsManagerEvents), 0x00000867, OnNewVoiceMailCountSinceLastResetChanged, &NewVoiceMailCountSinceLastResetChanged)
					SINK_ENTRY_INFO(1, __uuidof(_IVoiceMailsManagerEvents), 0x00000869, OnVoiceMailsAdded, &VoiceMailsAdded)
					SINK_ENTRY_INFO(1, __uuidof(_IVoiceMailsManagerEvents), 0x0000086a, OnVoiceMailsRemoved, &VoiceMailsRemoved)
					SINK_ENTRY_INFO(1, __uuidof(_IVoiceMailsManagerEvents), 0x0000086b, OnVoiceMailsModified, &VoiceMailsModified)
				END_SINK_MAP()
			};
		}
	}
}
