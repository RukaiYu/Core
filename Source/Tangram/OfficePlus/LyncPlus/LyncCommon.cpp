#include "../../stdafx.h"
#include "LyncCommon.h"


namespace OfficePlus
{
	namespace LyncPlus
	{
		CLyncRoomObj::CLyncRoomObj()
		{
			m_pRoom = nullptr;
		}

		CLyncRoomObj::~CLyncRoomObj()
		{
		}

		void CLyncRoomObj::OnPropertyChanged(IRoom* _eventSource, IRoomPropertyChangedEventData* _eventData) 
		{
		}

		void CLyncRoomObj::OnUnreadMessageCountChanged(IRoom* _eventSource, IUnreadMessageCountChangedEventData* _eventData)
		{
		}

		void CLyncRoomObj::OnIsSendingMessage(IRoom* _eventSource, IRoomMessageEventData* _eventData)
		{
		}

		void CLyncRoomObj::OnMessagesReceived(IRoom* _eventSource, IRoomMessagesEventData* _eventData)
		{
		}

		void CLyncRoomObj::OnParticipantAdded(IRoom* _eventSource, IRoomParticipantsEventData* _eventData)
		{
		}

		void CLyncRoomObj::OnParticipantRemoved(IRoom* _eventSource, IRoomParticipantsEventData* _eventData)
		{
		}

		void CLyncRoomObj::OnJoinStateChanged(IRoom* _eventSource, IRoomJoinStateChangedEventData* _eventData) 
		{
		}

		CLyncConversationObj::CLyncConversationObj()
		{
			m_strConversationID = _T("");
			m_strConversationSubject = _T("");
			m_strConversationConferencingUri = _T("");
			m_pConversation = nullptr;
		}

		CLyncConversationObj::~CLyncConversationObj()
		{
		}

		//struct __declspec(uuid("2a9385aa-cf61-4e47-b64c-214de4a4ad11"))
		//	IConversationStateChangedEventData : IDispatch
		//{
		//	//
		//	// Raw methods provided by interface
		//	//

		//	virtual HRESULT __stdcall get_OldState(
		//		/*[out,retval]*/ enum ConversationState * _oldState) = 0;
		//	virtual HRESULT __stdcall get_NewState(
		//		/*[out,retval]*/ enum ConversationState * _newState) = 0;
		//	virtual HRESULT __stdcall get_StatusCode(
		//		/*[out,retval]*/ long * _statusCode) = 0;
		//	virtual HRESULT __stdcall get_Properties(
		//		/*[out,retval]*/ struct IConversationStateChangePropertyDictionary * * _properties) = 0;
		//};

		//enum ConversationStateChangeProperty
		//{
		//	ucConversationStateChangeUnparkReason = 805896192,
		//	ucConversationStateChangeUnparkTarget = 589825
		//};
		void CLyncConversationObj::OnStateChanged(IConversation* _eventSource, IConversationStateChangedEventData* _eventData)
		{
			CComPtr<IConversationStateChangePropertyDictionary> pIConversationStateChangePropertyDictionary;
			_eventData->get_Properties(&pIConversationStateChangePropertyDictionary);
			if (pIConversationStateChangePropertyDictionary)
			{
				long nCount = 0;
				pIConversationStateChangePropertyDictionary->get_Count(&nCount);
				ConversationState nState;
				_eventData->get_NewState(&nState);
				switch (nState)
				{
				case ConversationState::ucConversationActive:
					break;
				case ConversationState::ucConversationInactive:
					break;
				case ConversationState::ucConversationParked:
					break;
				case ConversationState::ucConversationTerminated:
					break;
				}
			}
		}

		void CLyncConversationObj::OnPropertyChanged(IConversation* _eventSource, IConversationPropertyChangedEventData* _eventData)
		{
			ConversationProperty nProperty;
			VARIANT var;
			_eventData->get_Property(&nProperty);
			_eventData->get_Value(&var);
			switch (nProperty)
			{
			case ConversationProperty::ucConversationId:
			{
				m_strConversationID = OLE2T(var.bstrVal);
			}
				break;
			case ConversationProperty::ucConversationSubject:
			{
				m_strConversationSubject = OLE2T(var.bstrVal);
			}
				break;
			case ConversationProperty::ucConversationImportance:
				break;
			case ConversationProperty::ucConversationTransferredBy:
				break;
			case ConversationProperty::ucConversationReplaced:
				break;
			case ConversationProperty::ucConversationConferencingUri:
			{
				m_strConversationConferencingUri = OLE2T(var.bstrVal);
			}
				break;
			case ConversationProperty::ucConversationRepresentedBy:
				break;
			case ConversationProperty::ucConversationConferenceInviterRepresentationInfo:
				break;
			case ConversationProperty::ucConversationFollowUp:
				break;
			case ConversationProperty::ucConversationDirection:
				break;
			case ConversationProperty::ucConversationConferenceAcceptingParticipant:
				break;
			case ConversationProperty::ucConversationPreviousConversationId:
				break;
			case ConversationProperty::ucConversationAcceptanceState:
				break;
			case ConversationProperty::ucConversationIsUsbConversation:
				break;
			case ConversationProperty::ucConversationAutoTerminateOnIdle:
				break;
			case ConversationProperty::ucConversationConferenceEscalationProgress:
				break;
			case ConversationProperty::ucConversationConferenceEscalationResult:
				break;
			case ConversationProperty::ucConversationConferencingInvitedModes:
				break;
			case ConversationProperty::ucConversationInviter:
				break;
			case ConversationProperty::ucConversationConferencingLocked:
				break;
			case ConversationProperty::ucConversationConferencingFirstInstantMessage:
				break;
			case ConversationProperty::ucConversationConferenceAccessInformation:
				break;
			case ConversationProperty::ucConversationConferencingAccessType:
				break;
			case ConversationProperty::ucConversationCallParkOrbit:
				break;
			case ConversationProperty::ucConversationConferenceDisclaimer:
				break;
			case ConversationProperty::ucConversationConferenceDisclaimerAccepted:
				break;
			case ConversationProperty::ucConversationConferenceTerminateOnLeave:
				break;
			case ConversationProperty::ucConversationNumberOfParticipantsRecording:
				break;
			case ConversationProperty::ucConversationConferenceJoinDialogCompleted:
				break;
			case ConversationProperty::ucConversationLastActivityTimeStamp:
				break;
			case ConversationProperty::ucConversationConferenceDialogId:
				break;
			case ConversationProperty::ucConversationConferenceDialogFromTag:
				break;
			case ConversationProperty::ucConversationConferenceDialogToTag:
				break;
			case ConversationProperty::ucConversationConferenceEndorseEnabled:
				break;
			case ConversationProperty::ucConversationConferenceHardMute:
				break;
			case ConversationProperty::ucConversationConferenceAutoPromoteLevel:
				break;
			case ConversationProperty::ucConversationConferencePermittedAutoPromoteLevels:
				break;
			case ConversationProperty::ucConversationConferencePSTNBypassEnabled:
				break;
			case ConversationProperty::ucConversationConferencePermissionManager:
				break;
			case ConversationProperty::ucConversationPresentedItem:
				break;
			case ConversationProperty::ucConversationActivePresenter:
				break;
			case ConversationProperty::ucConversationConferenceGlobalAnnouncements:
				break;
			case ConversationProperty::ucConversationViewedItem:
				break;
			case ConversationProperty::ucConversationPresentationState:
				break;
			case ConversationProperty::ucConversationConferenceIsRosterLimited:
				break;
			}
			::VariantClear(&var);
		}

		void CLyncConversationObj::OnParticipantAdded(IConversation* _eventSource, IParticipantCollectionChangedEventData* _eventData)
		{
		}

		void CLyncConversationObj::OnParticipantRemoved(IConversation* _eventSource, IParticipantCollectionChangedEventData* _eventData)
		{
		}

		void CLyncConversationObj::OnActionAvailabilityChanged(IConversation* _eventSource, IConversationActionAvailabilityEventData* _eventData)
		{
			VARIANT_BOOL _isAvailable;
			_eventData->get_IsAvailable(&_isAvailable);
			if (_isAvailable)
			{
				ConversationAction _action;
				_eventData->get_Action(&_action);
				switch (_action)
				{
				case ConversationAction::ucConversationActionAddParticipant:
					break;
				case ConversationAction::ucConversationActionMerge:
					break;
				case ConversationAction::ucConversationActionPark:
					break;
				case ConversationAction::ucConversationActionRemoveParticipant:
					break;
				}
			}
		}

		void CLyncConversationObj::OnConversationContextAdded(IConversation* _eventSource, IConversationContextCollectionEventData* _eventData)
		{
		}

		void CLyncConversationObj::OnConversationContextRemoved(IConversation* _eventSource, IConversationContextCollectionEventData* _eventData)
		{
		}

		void CLyncConversationObj::OnConversationContextLinkClicked(IConversation* _eventSource, IInitialContextEventData* _eventData)
		{
		}

		void CLyncConversationObj::OnInitialContextReceived(IConversation* _eventSource, IInitialContextEventData* _eventData)
		{
		}

		void CLyncConversationObj::OnInitialContextSent(IConversation* _eventSource, IInitialContextEventData* _eventData)
		{
		}

		void CLyncConversationObj::OnContextDataReceived(IConversation* _eventSource, IContextEventData* _eventData)
		{
		}

		void CLyncConversationObj::OnContextDataSent(IConversation* _eventSource, IContextEventData* _eventData)
		{
		}

		CLyncConversationWindowObj::CLyncConversationWindowObj()
		{
			m_pConversationWindow = nullptr;
		}

		CLyncConversationWindowObj::~CLyncConversationWindowObj()
		{
		}

		void CLyncConversationWindowObj::OnNeedsSizeChange(IConversationWindow* _eventSource, IConversationWindowNeedsSizeChangeEventData* _eventData)
		{
		}

		void CLyncConversationWindowObj::OnNeedsAttention(IConversationWindow* _eventSource, IConversationWindowNeedsAttentionEventData* _eventData)
		{

		}

		void CLyncConversationWindowObj::OnStateChanged(IConversationWindow* _eventSource, IConversationWindowStateChangedEventData* _eventData)
		{
			ConversationWindowState m_OldState;
			_eventData->get_OldState(&m_OldState);
			ConversationWindowState m_NewState;
			_eventData->get_NewState(&m_NewState);
			switch (m_NewState)
			{
			case ConversationWindowState::uiaConversationWindowDestroyed:
				break;
			case ConversationWindowState::uiaConversationWindowInitialized:
				break;
			case ConversationWindowState::uiaConversationWindowNotInitialized:
				break;
			}
		}

		void CLyncConversationWindowObj::OnInformationChanged(IConversationWindow* _eventSource, IConversationWindowInformationChangedEventData* _eventData)
		{
			CComPtr<IConversationWindowInformationDictionary> pIConversationWindowInformationDictionary;
			_eventData->get_ChangedProperties(&pIConversationWindowInformationDictionary);
			if (pIConversationWindowInformationDictionary)
			{
				SAFEARRAY * pSAFEARRAY = nullptr;
				pIConversationWindowInformationDictionary->get_Keys(&pSAFEARRAY);

				//enum ConversationWindowInformationType
				//{
				//	uiaConversationWindowWidthMin = 537788416,
				//	uiaConversationWindowHeightMin = 537788417,
				//	uiaConversationWindowIsDocked = 269352962,
				//	uiaConversationWindowHasVideo = 269352963,
				//	uiaConversationWindowHasContentStage = 269352964,
				//	uiaConversationWindowHasExtensionPane = 269352965,
				//	uiaConversationWindowIsFullScreen = 269352966
				//};
			}
		}

		//IConversationWindowActionAvailabilityChangedEventData: IDispatch
		//{
		//	//
		//	// Raw methods provided by interface
		//	//

		//	virtual HRESULT __stdcall get_Action(
		//		/*[out,retval]*/ enum ConversationWindowAction * _action) = 0;
		//	virtual HRESULT __stdcall get_IsAvailable(
		//		/*[out,retval]*/ VARIANT_BOOL * _isAvailable) = 0;
		//};
		void CLyncConversationWindowObj::OnActionAvailabilityChanged(IConversationWindow* _eventSource, IConversationWindowActionAvailabilityChangedEventData* _eventData)
		{
			VARIANT_BOOL isAvailable;
			_eventData->get_IsAvailable(&isAvailable);
			if (isAvailable)
			{
				ConversationWindowAction action;
				_eventData->get_Action(&action);
				switch (action)
				{
				case ucConversationWindowActionAddOfficePowerPoint:
					break;
				case ucConversationWindowActionAddOfficeOneNote:
					break;
				case ucConversationWindowFullScreen:
					break;
				case ucConversationWindowActionMAX:
					break;
				}
			}
			else
			{

			}
		}
	}
}
