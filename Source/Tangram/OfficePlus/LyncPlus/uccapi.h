// Created by Microsoft (R) C/C++ Compiler Version 15.00.21022.08 (85218fb3).
//
// f:\ecp2009\communicatorshell\debug\uccapi.tlh
//
// C++ source equivalent of Win32 type library C:\Program Files\Microsoft Office Communicator\uccApi.dll
// compiler-generated file created 05/26/09 at 12:19:06 - DO NOT EDIT!

#pragma once
#pragma pack(push, 8)

#include <comdef.h>

namespace UCCAPILib {

//
// Forward references and typedefs
//

struct __declspec(uuid("ecd5307e-4419-43cf-8bda-c9946ac375cf"))
/* LIBID */ __UCCAPILib;
enum UCC_AUTHENTICATION_MODES;
enum UCC_CATEGORY_INSTANCE_EXPIRE_TYPE;
enum UCC_CATEGORY_CONTEXT_EVENT_CONTEXT;
enum UCC_CATEGORY_REMOVAL_REASON;
enum UCC_CATEGORY_INSTANCE_EVENT_CONTEXT;
enum UCC_CHANNEL_CONTEXT;
enum UCC_CONFERENCE_ADMISSION_TYPE;
enum UCC_CONFERENCE_ENTITY_PROPERTY;
enum UCC_CONFERENCE_ENTITY_SETTING_PROPERTY;
enum UCC_CONFERENCE_ENTITY_STATE_PROPERTY;
enum UCC_CONFERENCE_ENTITY_TYPE;
enum UCC_CONFERENCE_INFORMATION_PROPERTY;
enum UCC_CONFERENCE_MANAGER_OPERATION_COMPLETED_EVENT_PROPERTY;
enum UCC_CONFERENCE_MEDIA_CHANNEL_PROPERTY;
enum UCC_CONFERENCE_MEDIA_FILTERS_RULE_PROPERTY;
enum UCC_CONFERENCE_MEDIA_JOIN_TYPE;
enum UCC_CONFERENCE_MEDIUM_PROPERTY;
enum UCC_CONFERENCE_PARTICIPANT_AUTHENTICATION_TYPE;
enum UCC_CONFERENCE_PARTICIPANT_ENDPOINT_CONTEXT;
enum UCC_CONFERENCE_PARTICIPANT_ENDPOINT_PROPERTY;
enum UCC_CONFERENCE_PARTICIPANT_PROPERTY;
enum UCC_CONFERENCE_PARTICIPANT_ROLE;
enum UCC_CONFERENCE_RECORDED_MEDIA_PROPERTY;
enum UCC_CONFERENCE_SESSION_PROPERTY;
enum UCC_CONFERENCE_VIDEO_SWITCHING_MODE;
enum UCC_CONTACT_CARD_ADDRESS_TYPE;
enum UCC_CONTACT_CARD_PHONE_TYPE;
enum UCC_CONTACT_CARD_URL_TYPE;
enum UCC_CONTAINER_MEMBER_COMPARE_RESULT;
enum UCC_CONTAINER_MEMBER_MATCHING_TYPE;
enum UCC_CONTAINER_MEMBERSHIP_SCOPE;
enum UCC_DIAGNOSTIC_INFO_TYPE;
enum UCC_DTMF;
enum UCC_ENDPOINT_SIGNALING_CAPABILITIES;
enum UCC_FIND_MEDIA_CONNECTIVITY_SERVER_REQUEST_CONTEXT;
enum UCC_GROUP_TYPE;
enum UCC_INCOMING_REQUEST_EVENT_CONTEXT;
enum UCC_MEDIA_CONNECTIVITY_SERVER_CONFIGURATION_CONTEXT;
enum UCC_MEDIA_CONNECTIVITY_SERVER_LOCATION;
enum UCC_MEDIA_CONNECTIVITY_SERVER_TYPE;
enum UCC_MEDIA_DIRECTIONS;
enum UCC_MEDIA_ENCRYPTION_LEVEL;
enum UCC_MEDIA_ENCRYPTION_MODES;
enum UCC_MEDIA_STREAM_STATE;
enum UCC_MEDIA_CHANNEL_CONNECTION_STATE;
enum UCC_MEDIA_STREAM_STATE_CHANGE_REASON;
enum UCC_MEDIA_TYPES;
enum UCC_OPERATION_PROGRESS_EVENT_CONTEXT;
enum UCC_OPERATION_CONTEXT;
enum UCC_PARTICIPANT_STATE_CHANGE_REASON;
enum UCC_PLATFORM_CONTEXT;
enum UCC_PRESENCE_AVAILABILITY;
enum UCC_PRESENCE_DEVICE_CAPABILITIES;
enum UCC_PRESENCE_NOTE_TYPE;
enum UCC_PRESENCE_WORKING_HOURS_DAYS;
enum UCC_PRESENCE_WORKING_HOURS_TIMEZONE_TYPE;
enum UCC_ENDPOINT_TYPE;
enum UCC_PUBLICATION_OPERATION_TYPE;
enum UCC_REJECT_OR_TERMINATE_REASON;
enum UCC_RETARGET_REASON;
enum UCC_SERVER_ENDPOINT_CONTEXT;
enum UCC_SERVER_ENDPOINT_FLAGS;
enum UCC_SESSION_CONTEXT;
enum UCC_SESSION_REDIRECT_METHOD;
enum UCC_SESSION_TYPE;
enum UCC_SESSION_ENTITY_STATE;
enum UCC_SIGNALING_MESSAGE_TYPE;
enum UCC_SIGNALING_REQUEST_TYPE;
enum UCC_SOURCE_NETWORK_TYPE;
enum UCC_TRANSFER_CONTEXT;
enum UCC_TRANSPORT_MODE;
enum UCC_VIDEO_TYPE;
enum UCC_MEDIA_DEVICE_SETTING_PROPERTY;
enum UCC_TUNING_WIZARD_PAGE;
enum UCC_MEDIA_DEVICE_PROPERTY;
enum UCC_SUBSCRIPTION_CONTEXT;
enum UCC_PROXY_ENDPOINT_CONTEXT_EXTENSIONS;
enum UCC_USER_SEARCH_COLUMN;
enum UCC_USER_SEARCH_PREFERENCE;
enum UCC_PROXY_ENDPOINT_COLLECTION_REASON;
enum UCC_DELEGATE_PRIVILEGES;
enum UCC_SESSION_RETRY_PREFERENCE;
enum UCC_APPLICATION_SHARING_RESOURCE_TYPE;
enum UCC_APPLICATION_SHARING_RESOURCE_PROPERTY;
enum UCC_APPLICATION_SHARING_RESOURCE_STATUS;
enum UCC_APPLICATION_SHARING_CHANNEL_TYPE;
enum UCC_SESSION_PARTICIPANT_CONTEXT;
enum UCC_BOUNDING_RECTANGLE_PROPERTY;
struct __declspec(uuid("e8909d14-268c-4cb9-a9bf-48818f5c48bc"))
/* interface */ IUccApplicationSessionParticipant;
struct __declspec(uuid("afe09299-87d8-4737-ad38-fa6d78537f4a"))
/* interface */ IUccOperationContext;
struct __declspec(uuid("e66a1e6b-5c1c-40b6-a3ad-23c3699b9658"))
/* interface */ IUccContext;
struct __declspec(uuid("73fe78fc-efe6-422b-b9dd-4a33d9ed6395"))
/* interface */ IUccPropertyCollection;
struct __declspec(uuid("e97197bf-9939-4ab5-9056-5a1fdbd15390"))
/* interface */ IUccReadOnlyPropertyCollection;
struct __declspec(uuid("0814cabe-959d-47b4-841e-5b77b06d6046"))
/* dual interface */ IUccCollection;
struct __declspec(uuid("cadc8c3e-857e-4898-9083-8750559376af"))
/* interface */ IUccProperty;
struct __declspec(uuid("307c5595-c51d-414b-bccf-38c0cd227f52"))
/* dual interface */ IUccArray;
struct __declspec(uuid("5c3ed734-2a86-4b9a-8a03-837d4342f5d3"))
/* interface */ IUccAudioMediaChannel;
struct __declspec(uuid("a98332d2-2d80-4759-b8a8-cc5f4c79f460"))
/* interface */ IUccAudioMediaDevice;
struct __declspec(uuid("d0283fc6-4ab7-4d63-bb24-8cdd72fd897a"))
/* interface */ IUccAudioVideoSession;
struct __declspec(uuid("c3c42eac-9d5c-4daa-abcc-84019d3706ad"))
/* interface */ IUccAudioVideoSessionParticipant;
struct __declspec(uuid("faff2612-3bf1-44aa-acb6-80d39c0d991d"))
/* interface */ IUccMediaChannel;
struct __declspec(uuid("c2a50530-5635-41b5-8d30-d406489238fa"))
/* interface */ IUccConferenceEntityView;
struct __declspec(uuid("8b55d477-d96e-4670-b0a9-f1653b9d97ce"))
/* interface */ IUccUri;
enum UCC_URI_SCHEME;
enum UCC_URI_COMPARISON_OPTION;
struct __declspec(uuid("65d83c13-b0a2-400c-bfd8-bcc4890b7574"))
/* interface */ IUccConferenceInformation;
struct __declspec(uuid("9dcfc8ab-4b32-4078-8414-a59de529b327"))
/* interface */ IUccConferenceManager;
struct __declspec(uuid("052e68e9-3ea7-4b9c-972c-91013f044af7"))
/* interface */ IUccConferenceManagerSession;
struct __declspec(uuid("af404bf1-2495-4c3e-bf3b-be8359341f78"))
/* interface */ IUccConferenceMediaChannel;
struct __declspec(uuid("16b1749a-39ce-4fd1-945c-b13d569261b9"))
/* interface */ IUccSessionParticipantEndpoint;
struct __declspec(uuid("c520c114-7ebf-4a14-8416-d17adfa86227"))
/* interface */ IUccSessionParticipant;
struct __declspec(uuid("233693fb-ad38-4152-8d4f-688fdcc3fcde"))
/* interface */ IUccSession;
struct __declspec(uuid("a0897f48-a177-4c66-a223-f7b8d695b6b8"))
/* interface */ IUccConferenceApplicationViewerChannel;
struct __declspec(uuid("a8c81cfe-853d-413a-ab91-d5b1c6ee1099"))
/* interface */ IUccConferenceView;
struct __declspec(uuid("f8c1eee0-4272-4dd6-8986-786358e6a0e3"))
/* interface */ IUccConferenceSession;
struct __declspec(uuid("09717930-a1de-4b46-bae2-08f3e5c74fa2"))
/* interface */ IUccConferenceSessionParticipant;
struct __declspec(uuid("48461087-92f1-43bc-9d42-ae615fc39771"))
/* interface */ IUccConferenceSessionParticipantEndpoint;
struct __declspec(uuid("bfc14f2e-90fe-4a5e-894d-62a16435b4d8"))
/* interface */ IUccCredential;
struct __declspec(uuid("ab9f6a4a-a930-4b5e-bb77-ea6ec31346b3"))
/* interface */ IUccCredentialCache;
struct __declspec(uuid("ce70dd5e-c89e-4863-8a59-39221773c9e9"))
/* interface */ IUccDelegate;
struct __declspec(uuid("d3e9e4fa-e813-431e-aa53-b73cd7799843"))
/* interface */ IUccDelegatee;
struct __declspec(uuid("fab74d22-5830-4d14-a918-177f77ea0256"))
/* interface */ IUccDelegateEvent;
struct __declspec(uuid("7bcb8062-1931-42ca-8cc6-38ea95d1941f"))
/* interface */ IUccDiagnosticInfo;
struct __declspec(uuid("ce5f4b42-3d3f-4466-a3d5-b11e588fade8"))
/* interface */ IUccDiagnosticReportingSettings;
enum UCC_DIAGNOSTIC_REPORTING_METHODS;
struct __declspec(uuid("2dc3c92f-8fe2-4a84-8f76-bb1a14ff0f3a"))
/* interface */ IUccEndpoint;
struct __declspec(uuid("943bca5a-e22d-4c5d-88d0-d0021ee510d7"))
/* interface */ IUccEndpoint2;
enum UCC_ENDPOINT_STATE;
struct __declspec(uuid("6cfb1827-9feb-462a-97c6-b858702c4fe4"))
/* interface */ IUccOperationProgressEvent;
struct __declspec(uuid("91154bc1-8079-4bf7-8914-fc5b62aac3d6"))
/* interface */ IUccFindMediaConnectivityServersEvent;
struct __declspec(uuid("162809da-825d-4e75-9877-a8909897e567"))
/* interface */ IUccFindServerEvent;
struct __declspec(uuid("36e59b05-2fc3-4782-b3c9-7d5dddbe25b4"))
/* interface */ IUccHistoryInfoEntry;
struct __declspec(uuid("2f96ff5d-e7d3-4aef-a979-4a02d69b9dd0"))
/* interface */ IUccInstantMessagingSession;
struct __declspec(uuid("e5e2d6bc-910d-4348-9d6b-23ec74908e68"))
/* interface */ IUccInstantMessagingSessionParticipant;
struct __declspec(uuid("780eca62-e839-4bb0-84c1-005f5f4bc855"))
/* interface */ IUccMediaConnectivityServer;
struct __declspec(uuid("6a84605f-af6a-402b-9b0d-5ff881f81879"))
/* interface */ IUccMediaConnectivityServerConfiguration;
struct __declspec(uuid("8d538142-1fc9-4cd4-b53b-5601360a908e"))
/* interface */ IUccMediaChannelDevices;
struct __declspec(uuid("216d2747-60a0-43be-abb3-1993d263aa68"))
/* interface */ IUccMediaDevice;
struct __declspec(uuid("24b14557-850e-494f-9ca7-799559cd3eeb"))
/* interface */ IUccMediaDeviceManager;
enum UCC_AUDIO_VIDEO_DEVICE_CATEGORY;
struct __declspec(uuid("56a60090-39c2-48fb-ad07-43c847c55c4c"))
/* interface */ IUccMediaDeviceSettings;
struct __declspec(uuid("4d587df7-1adb-48cb-b6b4-914a85c26f6a"))
/* interface */ IUccMediaEndpointSettings;
struct __declspec(uuid("2df08b07-6f9c-4890-a2b0-dda68f430bfb"))
/* interface */ IUccMediaPlayerDevice;
struct __declspec(uuid("8ab54fb5-7656-4348-b5c5-32ca2efee7f4"))
/* interface */ IUccPlatform;
struct __declspec(uuid("8e17e673-96ec-4aea-aaec-5280ed162410"))
/* interface */ IUccProvisioningPolicyInstance;
struct __declspec(uuid("b22ef3c5-c0dc-473b-b2a4-5ee253e87f5e"))
/* interface */ IUccProxyEndpoint;
struct __declspec(uuid("5e4739c8-240d-4678-9574-9b342f8e6094"))
/* interface */ IUccEndpointRefreshEvent;
struct __declspec(uuid("3405d52a-1498-4bd5-9248-e423e97bf1fe"))
/* interface */ IUccServerSignalingSettings;
struct __declspec(uuid("20f08a08-e163-45f8-b43a-2dab579c51ae"))
/* interface */ IUccSignalingServer;
struct __declspec(uuid("2d237189-43ff-40ce-bdfb-1f560d56b4fa"))
/* interface */ IUccSessionCallControl;
struct __declspec(uuid("57061595-cea1-4c1d-af44-3dff48a7e0df"))
/* interface */ IUccSessionManager;
struct __declspec(uuid("83a209ca-296c-44d9-9331-f5dcf3e3c0ab"))
/* interface */ _IUccSessionDescriptionEvaluator;
struct __declspec(uuid("8e0d1060-df8e-45f9-a834-237f1e4418f0"))
/* interface */ IUccTraceSettings;
struct __declspec(uuid("bb782727-4c0a-42ba-a8f8-cb6bd7140bab"))
/* interface */ IUccVideoMediaChannel;
struct __declspec(uuid("56a868b4-0ad4-11ce-b03a-0020af0ba770"))
/* dual interface */ IVideoWindow;
struct __declspec(uuid("230175b3-791c-4904-bce9-8048cda85f87"))
/* interface */ IUccVideoMediaDevice;
struct __declspec(uuid("90a4b7cd-1dac-498f-a12b-fce949550694"))
/* interface */ IUccAudioVideoMediaChannelStreamsEvent;
struct __declspec(uuid("63bbfd31-c323-4a38-be5a-e9af55d5e6a7"))
/* interface */ IUccAudioVideoMediaChannel;
struct __declspec(uuid("2c46dc00-1382-4c65-8b79-3cedb358e5d0"))
/* interface */ IUccAudioVideoMediaChannelSourceEvent;
struct __declspec(uuid("345420c6-72b7-4640-a474-9eb877fcf60a"))
/* interface */ IUccUserSearchManager;
struct __declspec(uuid("5072148c-de7a-4826-965c-812ab676e0a4"))
/* interface */ IUccUserSearchQuery;
struct __declspec(uuid("54562fbc-5a84-4461-8bc9-590737e5de13"))
/* interface */ IUccUserSearchQueryEvent;
struct __declspec(uuid("ac628a4c-7595-4129-83f8-b37c2f8af813"))
/* interface */ IUccUserSearchResultRecord;
struct __declspec(uuid("8f77e9f1-ebe2-496d-96fd-a7ddd1adb329"))
/* interface */ IUccConferenceMediaChannelStateChangedEvent;
enum UCC_CONFERENCE_MEDIA_STATE;
struct __declspec(uuid("db49e9b0-4c96-46b6-8227-159481610c6b"))
/* dispinterface */ _IUccApplicationSessionParticipantEvents;
struct __declspec(uuid("6d362a12-a539-4581-a2ec-2647caff6a72"))
/* interface */ IUccOutgoingInvitationEvent;
struct __declspec(uuid("4896d466-934f-43a2-bf32-e6ccc6721524"))
/* interface */ IUccIncomingInvitationEvent;
struct __declspec(uuid("ceb0d9cb-78b9-4d19-b8d9-480a1bbf9c38"))
/* interface */ IUccIncomingRequestEvent;
struct __declspec(uuid("06eefc22-08be-4ea4-879e-115939e59e96"))
/* interface */ IUccInvitationAcceptedEvent;
struct __declspec(uuid("b6dc5e1a-18ca-47bb-bce7-156142ac5453"))
/* dispinterface */ _IUccAudioVideoSessionEvents;
struct __declspec(uuid("80bdd0ab-207b-4180-8511-c9967b0bf698"))
/* dispinterface */ _IUccAudioVideoSessionParticipantEvents;
struct __declspec(uuid("46ace669-1ede-478d-98cd-f1d1ddda116f"))
/* interface */ IUccIncomingMediaRequestEvent;
struct __declspec(uuid("da1582aa-99e3-424b-87df-d3d15a278e75"))
/* dispinterface */ _IUccAudioMediaChannelEvents;
struct __declspec(uuid("d82b8d0c-1e1d-4f13-8ae1-dd54f70ffb09"))
/* interface */ IUccAudioSignalLevelChangedEvent;
struct __declspec(uuid("ea82efbc-1da0-45cb-b4c2-d033b5d19559"))
/* dispinterface */ _IUccMediaChannelDevicesEvents;
struct __declspec(uuid("d8ba144f-efe7-4eb7-857c-53771fc7984e"))
/* interface */ IUccMediaDeviceChangeEvent;
enum UCC_MEDIA_DEVICE_CHANGE_REASON;
struct __declspec(uuid("d7db1218-a366-4482-af15-41d5bc859d94"))
/* dispinterface */ _IUccMediaChannelEvents;
struct __declspec(uuid("14689259-2a48-4eb3-9094-ec65019b4076"))
/* interface */ IUccMediaChannelEvent;
struct __declspec(uuid("c7eddd72-0fd5-422e-a206-cf19b09d0e40"))
/* dispinterface */ _IUccMediaEndpointEvents;
struct __declspec(uuid("bbfcb8d3-7be1-4525-916b-548a6a617ee4"))
/* dispinterface */ _IUccAudioVideoMediaChannelEvents;
struct __declspec(uuid("f891ace7-e9fc-4768-8a50-61489020bc8d"))
/* dispinterface */ _IUccConferenceManagerSessionEvents;
struct __declspec(uuid("7dbbf4e6-a70d-4889-86a0-2d95d3ef047b"))
/* dispinterface */ _IUccConferenceMediaChannelEvents;
struct __declspec(uuid("826ef7c1-84b3-486a-9790-14430f642d4f"))
/* interface */ IUccPropertyUpdateEvent;
struct __declspec(uuid("386ab1af-dd8b-41e9-ad87-8c09583e805e"))
/* dispinterface */ _IUccConferenceMediaChannelEvents2;
struct __declspec(uuid("e5837f40-2f55-458f-9adf-42c1bf973560"))
/* interface */ IUccConferenceMediaChannelSourceEvent;
struct __declspec(uuid("07b3bf54-c27f-4413-b6bf-841cef84e5ba"))
/* interface */ IUccNamedPropertyUpdateEvent;
struct __declspec(uuid("8ec5f518-6c4d-4c79-afa6-47b6d1039bd0"))
/* dispinterface */ _IUccConferenceSessionEvents;
struct __declspec(uuid("390a56e6-5868-433c-8cf3-401d104c59f3"))
/* dispinterface */ _IUccConferenceSessionEvents2;
struct __declspec(uuid("25c44946-2e3d-4bf3-a5be-9513e5f969e9"))
/* interface */ IUccConferenceSession2;
struct __declspec(uuid("c819c292-f45f-4854-ae66-c9e3cfe45195"))
/* dispinterface */ _IUccConferenceSessionParticipantEvents;
struct __declspec(uuid("c57a388e-c215-4bd4-8b2d-da5f341ef96b"))
/* dispinterface */ _IUccConferenceSessionParticipantEndpointEvents;
struct __declspec(uuid("a68a5d90-944e-40c0-9146-883f9c8556bd"))
/* interface */ IUccConferenceSessionParticipantEndpointStateChangedEvent;
struct __declspec(uuid("75025986-cdaa-42e7-9e06-642a8e9e4253"))
/* dispinterface */ _IUccConferenceMediaChannelCollectionEvents;
struct __declspec(uuid("56a2af6c-e402-4fc9-9461-d8a92abe1918"))
/* interface */ IUccMediaChannelCollectionEvent;
struct __declspec(uuid("53548bb8-662d-4117-9f82-26b85a30742a"))
/* dispinterface */ _IUccConferenceEntityViewCollectionEvents;
struct __declspec(uuid("be9e679a-d5af-485e-b1f6-8c74378a5359"))
/* interface */ IUccConferenceEntityViewCollectionEvent;
struct __declspec(uuid("67e8922c-51d8-409a-9219-30646c50baea"))
/* dispinterface */ _IUccConferenceEntityViewEvents;
struct __declspec(uuid("060bd8f0-5a2a-4d56-a065-2173894aefaf"))
/* dispinterface */ _IUccEndpointEvents;
struct __declspec(uuid("1a8a9e13-99e3-4edf-a2f6-9ddbb041023c"))
/* dispinterface */ _IUccEndpointEvents2;
struct __declspec(uuid("fdcb6c19-66cb-492a-9f4e-996d92a50407"))
/* dispinterface */ _IUccInstantMessagingSessionEvents;
struct __declspec(uuid("70e0fff6-dd6d-4420-8e19-2de91da99195"))
/* interface */ IUccSessionOperationEvent;
struct __declspec(uuid("98a51831-7b5b-49ba-84cf-2a4cde61fa35"))
/* dispinterface */ _IUccInstantMessagingSessionParticipantEvents;
struct __declspec(uuid("ca7acf03-537d-4654-92f6-b556fde0cd80"))
/* interface */ IUccIncomingInstantMessageEvent;
struct __declspec(uuid("11ebeff7-e4c7-4bc3-8752-493b947004fa"))
/* interface */ IUccInstantMessagingComposingEvent;
struct __declspec(uuid("a1faa7b6-0e6e-4d7e-9e4d-62c1629292b8"))
/* dispinterface */ _IUccPlatformEvents;
struct __declspec(uuid("e7a261a2-f96b-4ce9-a0cf-660a4b64fa70"))
/* dispinterface */ _IUccProxyEndpointCollectionEvents;
struct __declspec(uuid("03988bae-597c-402e-8b7d-2cd41aa25ecb"))
/* interface */ IUccProxyEndpointCollectionEvent;
struct __declspec(uuid("acee6559-9e58-4c91-bc24-eea3f8fce5d6"))
/* dispinterface */ _IUccServerSignalingSettingsEvents;
struct __declspec(uuid("ab81ffb3-80a5-4fe9-ac37-b06d5d631be1"))
/* dispinterface */ _IUccSessionEvents;
struct __declspec(uuid("3130f1a7-0569-4b49-a8c4-6b902194e55e"))
/* dispinterface */ _IUccSessionParticipantCollectionEvents;
struct __declspec(uuid("005a2d10-dd52-4ace-9ca1-9bdd699315f0"))
/* interface */ IUccSessionParticipantCollectionEvent;
struct __declspec(uuid("90bbe2b8-fb81-4e70-ae47-92b75431d4c1"))
/* dispinterface */ _IUccSessionManagerEvents;
struct __declspec(uuid("37cf9d37-c22c-4c5d-9023-f78ae5325036"))
/* interface */ IUccIncomingSessionEvent;
struct __declspec(uuid("e11e841b-da1d-4425-bb62-44df4418cb24"))
/* interface */ IUccOutgoingSessionEvent;
struct __declspec(uuid("9b48f770-4eac-4967-bcac-283b73557860"))
/* dispinterface */ _IUccSessionParticipantEvents;
struct __declspec(uuid("d0e8a9a5-2f04-42f7-96ea-89ab2b547da8"))
/* interface */ IUccSessionParticipantStateChangedEvent;
struct __declspec(uuid("8db557ba-e669-4de9-9922-8c6c53ad1ea3"))
/* dispinterface */ _IUccSessionParticipantEndpointCollectionEvents;
struct __declspec(uuid("683701e6-6c96-441a-8ab2-b5292dbb7104"))
/* interface */ IUccSessionParticipantEndpointCollectionEvent;
struct __declspec(uuid("00fb5f00-b2d5-4f3d-9e4b-feb4180d66fa"))
/* dispinterface */ _IUccMediaChannelCollectionEvents;
struct __declspec(uuid("79712662-75f4-44a0-8519-1e6013de4109"))
/* dispinterface */ _IUccSessionCallControlEvents;
struct __declspec(uuid("c31836bd-ad83-4a63-b53a-ac6161cea97f"))
/* interface */ IUccSessionTransferProgressEvent;
struct __declspec(uuid("7d3446dd-a657-437a-b47f-3080cbb02ca3"))
/* interface */ IUccIncomingSessionTransferEvent;
struct __declspec(uuid("6a8e8235-9734-4eb0-938c-a6de10deb4ea"))
/* dispinterface */ _IUccSignalingChannelEvents;
struct __declspec(uuid("23053341-87a9-44eb-bb22-800835bba620"))
/* interface */ IUccSignalingChannel;
struct __declspec(uuid("9ea6b70c-c59e-4823-9fdc-e7c85e6320a0"))
/* interface */ IUccSignalingRequest;
struct __declspec(uuid("9d45f185-46b0-48b4-9e9c-080313280711"))
/* interface */ IUccSignalingMessage;
struct __declspec(uuid("05bf6f94-bb11-4f38-a45a-1e1cbe177dbf"))
/* interface */ IUccSignalingHeaderField;
struct __declspec(uuid("5a9747c4-41ed-4812-8e73-8dffd9f5c97b"))
/* interface */ IUccSignalingHeaderFieldParameter;
struct __declspec(uuid("bc23ff63-d024-46df-b980-f8d7d76cd6a9"))
/* interface */ IUccIncomingSignalingMessageEvent;
struct __declspec(uuid("94f59d79-583a-4547-a620-ead932a2f2eb"))
/* dispinterface */ _IUccUserSearchQueryEvents;
struct __declspec(uuid("8c37fed9-ab4a-45bc-b018-5ceefb94f34f"))
/* interface */ IUccSessionParticipantOperationEvent;
struct __declspec(uuid("1713afd4-3ba6-48bc-853b-e1b7e1425ec8"))
/* interface */ IUccOperationManager;
struct __declspec(uuid("fb194526-2e78-4cca-927e-e967bf0101f7"))
/* interface */ IUccSubscriptionManager;
struct __declspec(uuid("f6e4f91e-40e2-4bf2-a808-37ec02daf67b"))
/* interface */ IUccSubscription;
struct __declspec(uuid("d0171dc1-7a36-494b-a8e5-cc51711a8cac"))
/* interface */ IUccPresentity;
struct __declspec(uuid("69f9efcf-5fea-46b1-8395-bd96212a8c51"))
/* interface */ IUccCategoryContext;
struct __declspec(uuid("72d6a6c3-ac51-4166-93f0-e941c3d784df"))
/* interface */ IUccCategoryInstance;
struct __declspec(uuid("646d6d11-c95c-49b2-a60f-77c5e91408d9"))
/* interface */ IUccContainer;
struct __declspec(uuid("0dd847ab-0f32-4449-9165-7425d11ddd89"))
/* interface */ IUccContainerMember;
struct __declspec(uuid("6ed14bd9-627d-4d9d-9405-ed9ad0f07023"))
/* interface */ IUccCategoryInstanceEvent;
struct __declspec(uuid("218ed275-cbe4-43f2-9108-242e5968b7e4"))
/* interface */ IUccCategoryTransformationManager;
struct __declspec(uuid("84b2232b-9ac5-4a70-bee3-9afb52b3ad25"))
/* interface */ IUccCategoryInstanceTransformationManifest;
struct __declspec(uuid("3afa976f-86ee-412e-986c-eb084a3d0347"))
/* interface */ IUccPresenceStateInstance;
enum UCC_PRESENCE_STATE_TYPE;
struct __declspec(uuid("890cf073-bb33-4dcb-aee8-08c300469b62"))
/* interface */ IUccPresenceActivity;
struct __declspec(uuid("6fdbdefe-8a31-49cd-87b8-cc61a49925fe"))
/* interface */ IUccLocaleString;
struct __declspec(uuid("b6608b42-8838-4ae6-af83-d0f0b46c7f1a"))
/* interface */ IUccPresenceStateInstance2;
struct __declspec(uuid("f13220e8-dec0-4e04-ac27-3a7785e9d8c5"))
/* interface */ IUccCategoryContextEvent;
struct __declspec(uuid("104cc179-5a41-4539-a57f-708bcbaaacd8"))
/* interface */ IUccSubscriptionEvent;
struct __declspec(uuid("8698350f-4b7e-4e38-a7e7-83038b69db7b"))
/* interface */ IUccPublicationManager;
struct __declspec(uuid("f1d73cad-6f6a-413b-bc2b-158dc2a7ccb2"))
/* interface */ IUccPublication;
enum UCC_PRESENCE_CATEGORY_TYPE;
struct __declspec(uuid("70c8f58f-a1ca-4614-8760-68399e111cfa"))
/* interface */ IUccContact;
struct __declspec(uuid("f36d99bb-5cb1-4dd6-9677-11787069e405"))
/* interface */ IUccGroup;
struct __declspec(uuid("0994e0c2-b0ac-46b5-b258-7df51fab8b2e"))
/* interface */ IUccContactCollectionEvent;
struct __declspec(uuid("679d5f13-90bb-40ab-98f7-6243b25b8d89"))
/* interface */ IUccContainerMemberCollectionEvent;
struct __declspec(uuid("e7122690-cbf5-4d7d-9f1b-bf28668d91a9"))
/* interface */ IUccContainerMembershipEvent;
struct __declspec(uuid("92700fb4-ff9b-46b7-af97-789d3d0e4ce3"))
/* interface */ IUccContainerMembershipManager;
struct __declspec(uuid("7377729f-2fa4-4e73-a6a9-cdb8b035676d"))
/* interface */ IUccGroupMembershipEvent;
struct __declspec(uuid("d450256a-adb1-4ea7-9cf1-3bd645b1c6d1"))
/* interface */ IUccSubscriber;
struct __declspec(uuid("573039b6-c916-4e9a-a6a2-4b0530248af6"))
/* interface */ IUccSubscriberEx;
struct __declspec(uuid("f5791faa-7225-41a7-8043-040dcea4517a"))
/* interface */ IUccPresenceContactCardIdentity;
struct __declspec(uuid("755b7ded-e5a2-456a-a5d3-43c1b94cfbcc"))
/* interface */ IUccPresenceContactCardAddress;
struct __declspec(uuid("9433ccce-7a69-4a71-9713-eae839eaf83c"))
/* interface */ IUccPresenceContactCardPhone;
struct __declspec(uuid("06e754c6-4bff-41fb-ba4d-af1a962be236"))
/* interface */ IUccPresenceContactCardUrl;
struct __declspec(uuid("fc3cfe6f-29d0-4efc-b3ec-6c69bbcbcd74"))
/* interface */ IUccPresenceContactCardInstance;
struct __declspec(uuid("9cce5b3c-7dd5-4452-84f7-cadaebc1ce53"))
/* interface */ IUccPresenceContactCardInstance2;
struct __declspec(uuid("088fea55-8dc9-4808-a0f5-b04f04c513af"))
/* interface */ IUccPresenceFreeBusyInformation;
struct __declspec(uuid("f3933dc2-1e04-43c9-a0ff-5d578a352ba8"))
/* interface */ IUccPresenceTimeZoneInformation;
struct __declspec(uuid("10659f8e-f79f-4d16-b26d-a7079014483b"))
/* interface */ IUccPresenceWorkingHoursBlock;
struct __declspec(uuid("d8bec519-3f8d-4fa3-a9c3-43be4b4e7c0b"))
/* interface */ IUccPresenceWorkingHoursInformation;
struct __declspec(uuid("1de93dcf-3bd3-411d-822c-6ec2adf9ffb7"))
/* interface */ IUccPresenceCapabilities;
struct __declspec(uuid("a035a856-83b7-43bc-b3ac-1aa1e2c14ff1"))
/* interface */ IUccPresenceCapabilityInfo;
struct __declspec(uuid("0fae5c66-18af-4d5d-9164-b0e41db04fb7"))
/* interface */ IUccPresenceCalendarDataInstance;
struct __declspec(uuid("eb59d99c-8e23-410a-9f49-e3a3585c8116"))
/* interface */ IUccPresenceServicesInstance;
struct __declspec(uuid("21c89f28-862d-42ce-b12c-090cd3adea82"))
/* interface */ IUccPresenceServiceElement;
struct __declspec(uuid("53c63418-b703-4d25-95d6-c2385c18ed6a"))
/* interface */ IUccPresenceDeviceInstance;
struct __declspec(uuid("6a332cd1-594b-413b-9d42-55efc5a0a605"))
/* interface */ IUccPresenceDeviceInstance2;
struct __declspec(uuid("f0e4beea-90a7-48f4-b8f7-f3cce0d8a8fa"))
/* interface */ IUccPresenceNoteData;
struct __declspec(uuid("cca44790-b93f-4564-9135-384fcd1ed934"))
/* interface */ IUccPresenceNoteInstance;
struct __declspec(uuid("9081029f-bcc4-4cf5-ac2d-440c97b8315b"))
/* interface */ IUccPresencePhoneLine;
enum UCC_PRESENCE_PHONE_LINE_TYPE;
struct __declspec(uuid("273819b7-8c74-4300-8c65-fb7e7dfc0146"))
/* interface */ IUccRoutingRuleInstance;
struct __declspec(uuid("b8bc87e3-1c8b-4f9e-8ab6-8be9a2c671af"))
/* interface */ IUccRoutingRuleElement;
enum UCC_ROUTING_RULE_ELEMENT_TYPE;
struct __declspec(uuid("f28f34f5-9e9e-4166-8d6c-c9d1e6262487"))
/* interface */ IUccRoutingPreamble;
struct __declspec(uuid("d0467f80-4d40-4668-8acd-aeef4f69b92c"))
/* interface */ IUccRoutingPreambleElement;
enum UCC_ROUTING_PREAMBLE_ELEMENT_TYPE;
struct __declspec(uuid("0255c189-db59-4732-bb8c-73534860711c"))
/* interface */ IUccRoutingPreambleListElement;
struct __declspec(uuid("d624a32e-76b3-4d75-b48e-87a97d5fd786"))
/* interface */ IUccRoutingTarget;
struct __declspec(uuid("27c32a76-49b9-4a1a-b293-772e3d5a9020"))
/* interface */ IUccRoutingPreambleWaitElement;
struct __declspec(uuid("1e8e354c-7dcf-4ca3-a009-ed3876015076"))
/* interface */ IUccRoutingPreambleFlagsElement;
struct __declspec(uuid("6689760d-24fa-4b3f-ac35-2042bfd70b89"))
/* interface */ IUccServerConfigurationCategory;
struct __declspec(uuid("bf78a2b4-cb02-4c00-875f-f51e3f0a74f2"))
/* interface */ IUccSignalingMessage2;
struct __declspec(uuid("3b4688bf-4d98-4d91-b3fd-c30fe0389597"))
/* interface */ IUccSignalingBodyPart;
struct __declspec(uuid("2285d731-7269-43f8-ac54-68a61ca8e1a8"))
/* interface */ IUccSignalingResponse;
struct __declspec(uuid("f17f94f7-bf07-4949-ba20-eb374f6d58c0"))
/* interface */ IUccUriManager;
struct __declspec(uuid("79f1c237-bd7d-4627-befd-813986373fcc"))
/* interface */ IUccApplicationSharingResource;
struct __declspec(uuid("32d73ef2-18a2-42f1-816c-ef4436719f40"))
/* interface */ IUccApplicationSharingChannelResource;
struct __declspec(uuid("5383c344-0684-4cdf-8899-3af3963fa1e1"))
/* interface */ IUccApplicationSharingChannelResourceEvent;
struct __declspec(uuid("74236283-2f10-422a-a99f-da2b61585a53"))
/* interface */ IUccApplicationSharerChannelResource;
struct __declspec(uuid("5247ac0e-a9f3-4514-baac-ca7e9f52180f"))
/* interface */ IUccApplicationViewerChannelResource;
struct __declspec(uuid("cd0d0e0f-79d8-422d-9bd5-e439812253a4"))
/* interface */ IUccApplicationSharingViewer;
struct __declspec(uuid("b7112800-b839-4f64-8ca0-5ce22d1d6666"))
/* interface */ IUccApplicationSharingResourceManager;
struct __declspec(uuid("7d6773b5-19a5-4668-8811-5510dce5bee2"))
/* interface */ IUccApplicationSharingSession;
struct __declspec(uuid("11ceee73-aad1-40f5-a09a-ff88e321049d"))
/* interface */ IUccApplicationSharingSessionParticipant;
struct __declspec(uuid("3f25cb45-e64f-4a0d-865b-0822d62ac2cc"))
/* interface */ IUccApplicationSharingChannel;
struct __declspec(uuid("f6eaba3c-69d7-41cf-8710-d1300e8425d1"))
/* interface */ IUccApplicationSharingControlEvent;
struct __declspec(uuid("882ef9b5-92e6-466d-b31b-e7f6a4c3b57f"))
/* interface */ IUccApplicationSharingControlRequestAllowedEvent;
struct __declspec(uuid("1c88a56d-8f5f-4793-a990-e5333bfe8d37"))
/* interface */ IUccApplicationSharerChannel;
struct __declspec(uuid("ad1a9ef1-2584-4693-83c4-1f479d8b4fba"))
/* interface */ IUccApplicationSharingRequestControlEvent;
struct __declspec(uuid("a5951db4-3d41-4af0-a360-40e2af383527"))
/* interface */ IUccApplicationViewerChannel;
struct __declspec(uuid("809f1175-58df-4b0b-8e7d-12c15050e0d4"))
/* interface */ IUccApplicationSharingConferenceChannel;
struct __declspec(uuid("5d87afee-c509-4400-b780-3da13f94f174"))
/* interface */ IUccApplicationSharingChannelResources;
struct __declspec(uuid("81c623e0-0167-457b-80e6-c089c9105eeb"))
/* interface */ IUccMediaChannelConnectionStateChangedEvent;
struct __declspec(uuid("0b804e06-2e95-40b6-9748-a2aaf7796585"))
/* dispinterface */ _IUccApplicationSharingChannelResourceEvents;
struct __declspec(uuid("3677165d-f78e-4d54-9039-5a082294584e"))
/* dispinterface */ _IUccApplicationSharingViewerEvents;
struct __declspec(uuid("a7d52eb2-8ae3-47e5-97af-c96a6f639c2a"))
/* dispinterface */ _IUccApplicationSharingResourceManagerEvents;
struct __declspec(uuid("74543c38-8c7f-4eef-85c2-1f1eb281b0d8"))
/* interface */ IUccApplicationSharingResourceChangeEvent;
enum UCC_APPLICATION_SHARING_RESOURCE_CHANGE_REASON;
struct __declspec(uuid("e9a00f18-3f9f-4094-8347-887fca4b2b20"))
/* dispinterface */ _IUccApplicationSharingSessionParticipantEvents;
struct __declspec(uuid("414c2bb7-3f99-41db-a369-b5d74c31b5e9"))
/* dispinterface */ _IUccApplicationSharingChannelEvents;
struct __declspec(uuid("58ad2cbd-83cf-4e58-8041-7d5d664a5287"))
/* dispinterface */ _IUccApplicationSharerChannelEvents;
struct __declspec(uuid("fd372aa6-ab45-4fe8-a425-4d9c21717147"))
/* dispinterface */ _IUccApplicationViewerChannelEvents;
struct __declspec(uuid("f4ddfc28-ee50-48df-92ae-09fc427cb855"))
/* dispinterface */ _IUccApplicationSharingChannelResourcesEvents;
struct __declspec(uuid("5de958ab-a06c-4539-bb11-1d43b0cdf8bc"))
/* dispinterface */ _IUccApplicationSharingChannelCollectionEvents;
struct __declspec(uuid("15f5e833-3dc1-4f9a-993e-6fac96a82bbe"))
/* dispinterface */ _IUccPublicationManagerEvents;
struct __declspec(uuid("15f8e9d8-097d-470b-9bad-965cbcfa1e03"))
/* dispinterface */ _IUccCategoryInstanceEvents;
struct __declspec(uuid("c9077987-cb46-4f9c-b3fd-254fc8fa3f98"))
/* dispinterface */ _IUccCategoryContextEvents;
struct __declspec(uuid("717efb8d-956d-467d-8ccf-2a461fd630da"))
/* dispinterface */ _IUccContactEvents;
struct __declspec(uuid("cff1d60f-10f4-4ca6-abc5-063fd96e988b"))
/* dispinterface */ _IUccDelegateEvents;
struct __declspec(uuid("dd825224-63e0-4a35-bd54-9d3cbbedaabf"))
/* dispinterface */ _IUccGroupEvents;
struct __declspec(uuid("bc909c17-90d1-4c2c-be68-f74728866afc"))
/* dispinterface */ _IUccSubscriptionEvents;
struct __declspec(uuid("5c7e7df6-0b5a-4ec2-9f04-a875329b4ec1"))
/* dispinterface */ _IUccPresentityEvents;
struct __declspec(uuid("7a3ca589-6955-4831-8795-6623c78f8de3"))
/* dispinterface */ _IUccPublicationEvent;
struct __declspec(uuid("507298f3-7975-4817-8e35-33447e837fff"))
/* dispinterface */ _IUccContainerEvents;
struct __declspec(uuid("70eb6c1a-9445-4d34-a9a5-4ca0423af676"))
/* dispinterface */ _IUccContainerMemberEvents;
struct __declspec(uuid("edb173ee-ee09-407b-b0fe-aef5a895542f"))
/* dispinterface */ _IUccOperationManagerEvents;
struct __declspec(uuid("cb2b887f-d037-4309-a504-31d4b3c41563"))
/* dispinterface */ _IUccMediaDeviceManagerEvents;
struct /* coclass */ UccPlatform;
struct /* coclass */ UccPropertyCollection;
struct /* coclass */ UccOperationContext;
struct /* coclass */ UccContext;
struct /* coclass */ UccSignalingBodyPart;
struct /* coclass */ UccSignalingRequest;
struct /* coclass */ UccSignalingResponse;
struct /* coclass */ UccUriManager;
struct /* coclass */ UccApplicationSession;
struct /* coclass */ UccApplicationSessionParticipant;
struct /* coclass */ UccAudioMediaChannel;
struct /* coclass */ UccAudioMediaDevice;
struct /* coclass */ UccAudioVideoSession;
struct /* coclass */ UccAudioVideoSessionParticipant;
struct /* coclass */ UccConferenceEntityView;
struct /* coclass */ UccConferenceInformation;
struct /* coclass */ UccConferenceManagerSession;
struct /* coclass */ UccConferenceMediaChannel;
struct /* coclass */ UccConferenceView;
struct /* coclass */ UccConferenceSession;
struct /* coclass */ UccConferenceSessionParticipant;
struct /* coclass */ UccConferenceSessionParticipantEndpoint;
struct /* coclass */ UccCredential;
struct /* coclass */ UccCredentialCache;
struct /* coclass */ UccDiagnosticInfo;
struct /* coclass */ UccHistoryInfoEntry;
struct /* coclass */ UccFindServerEvent;
struct /* coclass */ UccInstantMessagingSession;
struct /* coclass */ UccInstantMessagingSessionParticipant;
struct /* coclass */ UccMediaConnectivityServer;
struct /* coclass */ UccMediaPlayerDevice;
struct /* coclass */ UccProperty;
struct /* coclass */ UccProvisioningPolicyInstance;
struct /* coclass */ UccProxyEndpoint;
struct /* coclass */ UccRemoteCallControlSession;
struct /* coclass */ UccRemoteCallControlSessionParticipant;
struct /* coclass */ UccRemoteCallControlAudioChannel;
struct /* coclass */ UccServerEndpoint;
struct __declspec(uuid("af9b29ec-84a0-482b-8e8a-2bd307d1eb24"))
/* interface */ IUccSignalingChannelManager;
struct /* coclass */ UccPrincipalServerEndpoint;
struct /* coclass */ UccDelegateEndpoint;
struct /* coclass */ UccAssociatedUriEndpoint;
struct /* coclass */ UccSignalingChannel;
struct /* coclass */ UccVideoMediaChannel;
struct /* coclass */ UccVideoMediaDevice;
struct /* coclass */ UccAudioVideoMediaChannelStreamsEvent;
struct /* coclass */ UccAudioVideoMediaChannelSourceEvent;
struct /* coclass */ UccAudioSignalLevelChangedEvent;
struct /* coclass */ UccIncomingMediaRequestEvent;
struct /* coclass */ UccIncomingInstantMessageEvent;
struct /* coclass */ UccIncomingInvitationEvent;
struct /* coclass */ UccIncomingSessionEvent;
struct /* coclass */ UccIncomingSessionTransferEvent;
struct /* coclass */ UccInstantMessagingComposingEvent;
struct /* coclass */ UccInvitationAcceptedEvent;
struct /* coclass */ UccMediaChannelEvent;
struct /* coclass */ UccOperationProgressEvent;
struct /* coclass */ UccOutgoingInvitationEvent;
struct /* coclass */ UccOutgoingSessionEvent;
struct /* coclass */ UccPropertyUpdateEvent;
struct /* coclass */ UccNamedPropertyUpdateEvent;
struct /* coclass */ UccSessionOperationEvent;
struct /* coclass */ UccSessionParticipantOperationEvent;
struct /* coclass */ UccSessionParticipantCollectionEvent;
struct /* coclass */ UccSessionParticipantEndpointCollectionEvent;
struct /* coclass */ UccSessionParticipantStateChangedEvent;
struct /* coclass */ UccSessionTransferProgressEvent;
struct /* coclass */ UccConferenceEntityViewCollectionEvent;
struct /* coclass */ UccConferenceSessionParticipantEndpointStateChangedEvent;
struct /* coclass */ UccMediaDeviceChangeEvent;
struct /* coclass */ UccProxyEndpointCollectionEvent;
struct /* coclass */ UccPresenceStateInstance;
struct /* coclass */ UccCategoryContext;
struct /* coclass */ UccCategoryContextEvent;
struct /* coclass */ UccCategoryInstance;
struct /* coclass */ UccCategoryInstanceEvent;
struct /* coclass */ UccGroupMembershipEvent;
struct /* coclass */ UccContainerMemberCollectionEvent;
struct /* coclass */ UccContainerMembershipEvent;
struct /* coclass */ UccContactCollectionEvent;
struct /* coclass */ UccSubscription;
struct /* coclass */ UccPresenceContactCardIdentity;
struct /* coclass */ UccLocaleString;
struct /* coclass */ UccPresenceContactCardAddress;
struct /* coclass */ UccPresenceContactCardPhone;
struct /* coclass */ UccPresenceContactCardUrl;
struct /* coclass */ UccPresenceContactCardInstance;
struct /* coclass */ UccPresenceFreeBusyInformation;
struct /* coclass */ UccPresenceTimeZoneInformation;
struct /* coclass */ UccPresenceWorkingHoursBlock;
struct /* coclass */ UccPresenceWorkingHoursInformation;
struct /* coclass */ UccPresenceActivity;
struct /* coclass */ UccPresenceCapabilities;
struct /* coclass */ UccPresenceCapabilityInfo;
struct /* coclass */ UccPresenceCalendarDataInstance;
struct /* coclass */ UccPresenceServicesInstance;
struct /* coclass */ UccPresenceServiceElement;
struct /* coclass */ UccPresenceDeviceInstance;
struct /* coclass */ UccPresenceNoteData;
struct /* coclass */ UccPresenceNoteInstance;
struct /* coclass */ UccPresentity;
struct /* coclass */ UccServerConfigurationCategory;
struct /* coclass */ UccSignalingHeaderField;
struct /* coclass */ UccSignalingHeaderFieldParameter;
struct /* coclass */ UccUri;
struct /* coclass */ UccContact;
struct /* coclass */ UccGroup;
struct /* coclass */ UccDelegate;
struct /* coclass */ UccDelegatee;
struct /* coclass */ UccRoutingRuleInstance;
struct /* coclass */ UccRoutingPreamble;
struct /* coclass */ UccRoutingPreambleListElement;
struct /* coclass */ UccRoutingTarget;
struct /* coclass */ UccRoutingPreambleWaitElement;
struct /* coclass */ UccRoutingPreambleFlagsElement;
struct /* coclass */ UccSubscriptionEvent;
struct /* coclass */ UccPresencePhoneLine;
struct /* coclass */ UccUserSearchQuery;
struct /* coclass */ UccUserSearchQueryEvent;
struct /* coclass */ UccUserSearchResultRecord;
struct /* coclass */ UccCategoryInstanceTransformationManifest;
struct /* coclass */ UccEndpointRefreshEvent;
struct /* coclass */ UccApplicationSharingSession;
struct /* coclass */ UccApplicationSharingSessionParticipant;
struct /* coclass */ UccApplicationSharingResourceChangeEvent;
struct /* coclass */ UccApplicationSharingRequestControlEvent;
struct /* coclass */ UccApplicationSharingControlEvent;
struct /* coclass */ UccApplicationSharingControlRequestAllowedEvent;
struct /* coclass */ UccConferenceMediaChannelSourceEvent;
struct /* coclass */ UccMediaChannelConnectionStateChangedEvent;
struct /* coclass */ UccConferenceMediaChannelStateChangedEvent;
#if !defined(_WIN64)
typedef __w64 long LONG_PTR_1;
#else
typedef __int64 LONG_PTR_1;
#endif

//
// Smart pointer typedef declarations
//

_COM_SMARTPTR_TYPEDEF(IUccCollection, __uuidof(IUccCollection));
_COM_SMARTPTR_TYPEDEF(IUccProperty, __uuidof(IUccProperty));
_COM_SMARTPTR_TYPEDEF(IUccReadOnlyPropertyCollection, __uuidof(IUccReadOnlyPropertyCollection));
_COM_SMARTPTR_TYPEDEF(IUccPropertyCollection, __uuidof(IUccPropertyCollection));
_COM_SMARTPTR_TYPEDEF(IUccContext, __uuidof(IUccContext));
_COM_SMARTPTR_TYPEDEF(IUccOperationContext, __uuidof(IUccOperationContext));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSessionParticipant, __uuidof(IUccApplicationSessionParticipant));
_COM_SMARTPTR_TYPEDEF(IUccArray, __uuidof(IUccArray));
_COM_SMARTPTR_TYPEDEF(IUccAudioMediaChannel, __uuidof(IUccAudioMediaChannel));
_COM_SMARTPTR_TYPEDEF(IUccAudioMediaDevice, __uuidof(IUccAudioMediaDevice));
_COM_SMARTPTR_TYPEDEF(IUccAudioVideoSession, __uuidof(IUccAudioVideoSession));
_COM_SMARTPTR_TYPEDEF(IUccMediaChannel, __uuidof(IUccMediaChannel));
_COM_SMARTPTR_TYPEDEF(IUccAudioVideoSessionParticipant, __uuidof(IUccAudioVideoSessionParticipant));
_COM_SMARTPTR_TYPEDEF(IUccUri, __uuidof(IUccUri));
_COM_SMARTPTR_TYPEDEF(IUccConferenceEntityView, __uuidof(IUccConferenceEntityView));
_COM_SMARTPTR_TYPEDEF(IUccConferenceInformation, __uuidof(IUccConferenceInformation));
_COM_SMARTPTR_TYPEDEF(IUccConferenceManagerSession, __uuidof(IUccConferenceManagerSession));
_COM_SMARTPTR_TYPEDEF(IUccConferenceManager, __uuidof(IUccConferenceManager));
_COM_SMARTPTR_TYPEDEF(IUccConferenceView, __uuidof(IUccConferenceView));
_COM_SMARTPTR_TYPEDEF(IUccConferenceSession, __uuidof(IUccConferenceSession));
_COM_SMARTPTR_TYPEDEF(IUccConferenceSessionParticipantEndpoint, __uuidof(IUccConferenceSessionParticipantEndpoint));
_COM_SMARTPTR_TYPEDEF(IUccCredential, __uuidof(IUccCredential));
_COM_SMARTPTR_TYPEDEF(IUccCredentialCache, __uuidof(IUccCredentialCache));
_COM_SMARTPTR_TYPEDEF(IUccDelegate, __uuidof(IUccDelegate));
_COM_SMARTPTR_TYPEDEF(IUccDelegatee, __uuidof(IUccDelegatee));
_COM_SMARTPTR_TYPEDEF(IUccDelegateEvent, __uuidof(IUccDelegateEvent));
_COM_SMARTPTR_TYPEDEF(IUccDiagnosticInfo, __uuidof(IUccDiagnosticInfo));
_COM_SMARTPTR_TYPEDEF(IUccDiagnosticReportingSettings, __uuidof(IUccDiagnosticReportingSettings));
_COM_SMARTPTR_TYPEDEF(IUccEndpoint, __uuidof(IUccEndpoint));
_COM_SMARTPTR_TYPEDEF(IUccEndpoint2, __uuidof(IUccEndpoint2));
_COM_SMARTPTR_TYPEDEF(IUccOperationProgressEvent, __uuidof(IUccOperationProgressEvent));
_COM_SMARTPTR_TYPEDEF(IUccFindMediaConnectivityServersEvent, __uuidof(IUccFindMediaConnectivityServersEvent));
_COM_SMARTPTR_TYPEDEF(IUccFindServerEvent, __uuidof(IUccFindServerEvent));
_COM_SMARTPTR_TYPEDEF(IUccHistoryInfoEntry, __uuidof(IUccHistoryInfoEntry));
_COM_SMARTPTR_TYPEDEF(IUccInstantMessagingSession, __uuidof(IUccInstantMessagingSession));
_COM_SMARTPTR_TYPEDEF(IUccInstantMessagingSessionParticipant, __uuidof(IUccInstantMessagingSessionParticipant));
_COM_SMARTPTR_TYPEDEF(IUccMediaConnectivityServer, __uuidof(IUccMediaConnectivityServer));
_COM_SMARTPTR_TYPEDEF(IUccMediaConnectivityServerConfiguration, __uuidof(IUccMediaConnectivityServerConfiguration));
_COM_SMARTPTR_TYPEDEF(IUccMediaDevice, __uuidof(IUccMediaDevice));
_COM_SMARTPTR_TYPEDEF(IUccMediaChannelDevices, __uuidof(IUccMediaChannelDevices));
_COM_SMARTPTR_TYPEDEF(IUccMediaDeviceManager, __uuidof(IUccMediaDeviceManager));
_COM_SMARTPTR_TYPEDEF(IUccMediaDeviceSettings, __uuidof(IUccMediaDeviceSettings));
_COM_SMARTPTR_TYPEDEF(IUccMediaEndpointSettings, __uuidof(IUccMediaEndpointSettings));
_COM_SMARTPTR_TYPEDEF(IUccMediaPlayerDevice, __uuidof(IUccMediaPlayerDevice));
_COM_SMARTPTR_TYPEDEF(IUccPlatform, __uuidof(IUccPlatform));
_COM_SMARTPTR_TYPEDEF(IUccProvisioningPolicyInstance, __uuidof(IUccProvisioningPolicyInstance));
_COM_SMARTPTR_TYPEDEF(IUccProxyEndpoint, __uuidof(IUccProxyEndpoint));
_COM_SMARTPTR_TYPEDEF(IUccEndpointRefreshEvent, __uuidof(IUccEndpointRefreshEvent));
_COM_SMARTPTR_TYPEDEF(IUccSignalingServer, __uuidof(IUccSignalingServer));
_COM_SMARTPTR_TYPEDEF(IUccServerSignalingSettings, __uuidof(IUccServerSignalingSettings));
_COM_SMARTPTR_TYPEDEF(_IUccSessionDescriptionEvaluator, __uuidof(_IUccSessionDescriptionEvaluator));
_COM_SMARTPTR_TYPEDEF(IUccTraceSettings, __uuidof(IUccTraceSettings));
_COM_SMARTPTR_TYPEDEF(IVideoWindow, __uuidof(IVideoWindow));
_COM_SMARTPTR_TYPEDEF(IUccVideoMediaChannel, __uuidof(IUccVideoMediaChannel));
_COM_SMARTPTR_TYPEDEF(IUccVideoMediaDevice, __uuidof(IUccVideoMediaDevice));
_COM_SMARTPTR_TYPEDEF(IUccAudioVideoMediaChannelStreamsEvent, __uuidof(IUccAudioVideoMediaChannelStreamsEvent));
_COM_SMARTPTR_TYPEDEF(IUccAudioVideoMediaChannel, __uuidof(IUccAudioVideoMediaChannel));
_COM_SMARTPTR_TYPEDEF(IUccUserSearchQuery, __uuidof(IUccUserSearchQuery));
_COM_SMARTPTR_TYPEDEF(IUccUserSearchManager, __uuidof(IUccUserSearchManager));
_COM_SMARTPTR_TYPEDEF(IUccUserSearchQueryEvent, __uuidof(IUccUserSearchQueryEvent));
_COM_SMARTPTR_TYPEDEF(IUccUserSearchResultRecord, __uuidof(IUccUserSearchResultRecord));
_COM_SMARTPTR_TYPEDEF(IUccConferenceMediaChannelStateChangedEvent, __uuidof(IUccConferenceMediaChannelStateChangedEvent));
_COM_SMARTPTR_TYPEDEF(_IUccApplicationSessionParticipantEvents, __uuidof(_IUccApplicationSessionParticipantEvents));
_COM_SMARTPTR_TYPEDEF(IUccOutgoingInvitationEvent, __uuidof(IUccOutgoingInvitationEvent));
_COM_SMARTPTR_TYPEDEF(IUccIncomingRequestEvent, __uuidof(IUccIncomingRequestEvent));
_COM_SMARTPTR_TYPEDEF(IUccIncomingInvitationEvent, __uuidof(IUccIncomingInvitationEvent));
_COM_SMARTPTR_TYPEDEF(IUccInvitationAcceptedEvent, __uuidof(IUccInvitationAcceptedEvent));
_COM_SMARTPTR_TYPEDEF(_IUccAudioVideoSessionEvents, __uuidof(_IUccAudioVideoSessionEvents));
_COM_SMARTPTR_TYPEDEF(_IUccAudioVideoSessionParticipantEvents, __uuidof(_IUccAudioVideoSessionParticipantEvents));
_COM_SMARTPTR_TYPEDEF(IUccIncomingMediaRequestEvent, __uuidof(IUccIncomingMediaRequestEvent));
_COM_SMARTPTR_TYPEDEF(_IUccAudioMediaChannelEvents, __uuidof(_IUccAudioMediaChannelEvents));
_COM_SMARTPTR_TYPEDEF(IUccAudioSignalLevelChangedEvent, __uuidof(IUccAudioSignalLevelChangedEvent));
_COM_SMARTPTR_TYPEDEF(_IUccMediaChannelDevicesEvents, __uuidof(_IUccMediaChannelDevicesEvents));
_COM_SMARTPTR_TYPEDEF(IUccMediaDeviceChangeEvent, __uuidof(IUccMediaDeviceChangeEvent));
_COM_SMARTPTR_TYPEDEF(_IUccMediaChannelEvents, __uuidof(_IUccMediaChannelEvents));
_COM_SMARTPTR_TYPEDEF(IUccMediaChannelEvent, __uuidof(IUccMediaChannelEvent));
_COM_SMARTPTR_TYPEDEF(_IUccMediaEndpointEvents, __uuidof(_IUccMediaEndpointEvents));
_COM_SMARTPTR_TYPEDEF(_IUccAudioVideoMediaChannelEvents, __uuidof(_IUccAudioVideoMediaChannelEvents));
_COM_SMARTPTR_TYPEDEF(_IUccConferenceManagerSessionEvents, __uuidof(_IUccConferenceManagerSessionEvents));
_COM_SMARTPTR_TYPEDEF(_IUccConferenceMediaChannelEvents, __uuidof(_IUccConferenceMediaChannelEvents));
_COM_SMARTPTR_TYPEDEF(IUccPropertyUpdateEvent, __uuidof(IUccPropertyUpdateEvent));
_COM_SMARTPTR_TYPEDEF(_IUccConferenceMediaChannelEvents2, __uuidof(_IUccConferenceMediaChannelEvents2));
_COM_SMARTPTR_TYPEDEF(IUccNamedPropertyUpdateEvent, __uuidof(IUccNamedPropertyUpdateEvent));
_COM_SMARTPTR_TYPEDEF(_IUccConferenceSessionEvents, __uuidof(_IUccConferenceSessionEvents));
_COM_SMARTPTR_TYPEDEF(_IUccConferenceSessionEvents2, __uuidof(_IUccConferenceSessionEvents2));
_COM_SMARTPTR_TYPEDEF(IUccConferenceSession2, __uuidof(IUccConferenceSession2));
_COM_SMARTPTR_TYPEDEF(_IUccConferenceSessionParticipantEvents, __uuidof(_IUccConferenceSessionParticipantEvents));
_COM_SMARTPTR_TYPEDEF(_IUccConferenceSessionParticipantEndpointEvents, __uuidof(_IUccConferenceSessionParticipantEndpointEvents));
_COM_SMARTPTR_TYPEDEF(IUccConferenceSessionParticipantEndpointStateChangedEvent, __uuidof(IUccConferenceSessionParticipantEndpointStateChangedEvent));
_COM_SMARTPTR_TYPEDEF(_IUccConferenceMediaChannelCollectionEvents, __uuidof(_IUccConferenceMediaChannelCollectionEvents));
_COM_SMARTPTR_TYPEDEF(IUccMediaChannelCollectionEvent, __uuidof(IUccMediaChannelCollectionEvent));
_COM_SMARTPTR_TYPEDEF(_IUccConferenceEntityViewCollectionEvents, __uuidof(_IUccConferenceEntityViewCollectionEvents));
_COM_SMARTPTR_TYPEDEF(IUccConferenceEntityViewCollectionEvent, __uuidof(IUccConferenceEntityViewCollectionEvent));
_COM_SMARTPTR_TYPEDEF(_IUccConferenceEntityViewEvents, __uuidof(_IUccConferenceEntityViewEvents));
_COM_SMARTPTR_TYPEDEF(_IUccEndpointEvents, __uuidof(_IUccEndpointEvents));
_COM_SMARTPTR_TYPEDEF(_IUccEndpointEvents2, __uuidof(_IUccEndpointEvents2));
_COM_SMARTPTR_TYPEDEF(_IUccInstantMessagingSessionEvents, __uuidof(_IUccInstantMessagingSessionEvents));
_COM_SMARTPTR_TYPEDEF(IUccSessionOperationEvent, __uuidof(IUccSessionOperationEvent));
_COM_SMARTPTR_TYPEDEF(_IUccInstantMessagingSessionParticipantEvents, __uuidof(_IUccInstantMessagingSessionParticipantEvents));
_COM_SMARTPTR_TYPEDEF(_IUccPlatformEvents, __uuidof(_IUccPlatformEvents));
_COM_SMARTPTR_TYPEDEF(_IUccProxyEndpointCollectionEvents, __uuidof(_IUccProxyEndpointCollectionEvents));
_COM_SMARTPTR_TYPEDEF(IUccProxyEndpointCollectionEvent, __uuidof(IUccProxyEndpointCollectionEvent));
_COM_SMARTPTR_TYPEDEF(_IUccServerSignalingSettingsEvents, __uuidof(_IUccServerSignalingSettingsEvents));
_COM_SMARTPTR_TYPEDEF(_IUccSessionEvents, __uuidof(_IUccSessionEvents));
_COM_SMARTPTR_TYPEDEF(_IUccSessionParticipantCollectionEvents, __uuidof(_IUccSessionParticipantCollectionEvents));
_COM_SMARTPTR_TYPEDEF(_IUccSessionManagerEvents, __uuidof(_IUccSessionManagerEvents));
_COM_SMARTPTR_TYPEDEF(_IUccSessionParticipantEvents, __uuidof(_IUccSessionParticipantEvents));
_COM_SMARTPTR_TYPEDEF(IUccSessionParticipantStateChangedEvent, __uuidof(IUccSessionParticipantStateChangedEvent));
_COM_SMARTPTR_TYPEDEF(_IUccSessionParticipantEndpointCollectionEvents, __uuidof(_IUccSessionParticipantEndpointCollectionEvents));
_COM_SMARTPTR_TYPEDEF(_IUccMediaChannelCollectionEvents, __uuidof(_IUccMediaChannelCollectionEvents));
_COM_SMARTPTR_TYPEDEF(_IUccSessionCallControlEvents, __uuidof(_IUccSessionCallControlEvents));
_COM_SMARTPTR_TYPEDEF(IUccSessionTransferProgressEvent, __uuidof(IUccSessionTransferProgressEvent));
_COM_SMARTPTR_TYPEDEF(IUccIncomingSessionTransferEvent, __uuidof(IUccIncomingSessionTransferEvent));
_COM_SMARTPTR_TYPEDEF(_IUccSignalingChannelEvents, __uuidof(_IUccSignalingChannelEvents));
_COM_SMARTPTR_TYPEDEF(IUccSignalingHeaderFieldParameter, __uuidof(IUccSignalingHeaderFieldParameter));
_COM_SMARTPTR_TYPEDEF(IUccSignalingHeaderField, __uuidof(IUccSignalingHeaderField));
_COM_SMARTPTR_TYPEDEF(IUccSignalingMessage, __uuidof(IUccSignalingMessage));
_COM_SMARTPTR_TYPEDEF(IUccSignalingRequest, __uuidof(IUccSignalingRequest));
_COM_SMARTPTR_TYPEDEF(IUccSignalingChannel, __uuidof(IUccSignalingChannel));
_COM_SMARTPTR_TYPEDEF(IUccIncomingSignalingMessageEvent, __uuidof(IUccIncomingSignalingMessageEvent));
_COM_SMARTPTR_TYPEDEF(_IUccUserSearchQueryEvents, __uuidof(_IUccUserSearchQueryEvents));
_COM_SMARTPTR_TYPEDEF(IUccOperationManager, __uuidof(IUccOperationManager));
_COM_SMARTPTR_TYPEDEF(IUccCategoryInstanceTransformationManifest, __uuidof(IUccCategoryInstanceTransformationManifest));
_COM_SMARTPTR_TYPEDEF(IUccCategoryTransformationManager, __uuidof(IUccCategoryTransformationManager));
_COM_SMARTPTR_TYPEDEF(IUccLocaleString, __uuidof(IUccLocaleString));
_COM_SMARTPTR_TYPEDEF(IUccPresenceActivity, __uuidof(IUccPresenceActivity));
_COM_SMARTPTR_TYPEDEF(IUccPresenceStateInstance, __uuidof(IUccPresenceStateInstance));
_COM_SMARTPTR_TYPEDEF(IUccPresenceStateInstance2, __uuidof(IUccPresenceStateInstance2));
_COM_SMARTPTR_TYPEDEF(IUccSubscriber, __uuidof(IUccSubscriber));
_COM_SMARTPTR_TYPEDEF(IUccSubscriberEx, __uuidof(IUccSubscriberEx));
_COM_SMARTPTR_TYPEDEF(IUccPresenceContactCardIdentity, __uuidof(IUccPresenceContactCardIdentity));
_COM_SMARTPTR_TYPEDEF(IUccPresenceContactCardAddress, __uuidof(IUccPresenceContactCardAddress));
_COM_SMARTPTR_TYPEDEF(IUccPresenceContactCardPhone, __uuidof(IUccPresenceContactCardPhone));
_COM_SMARTPTR_TYPEDEF(IUccPresenceContactCardUrl, __uuidof(IUccPresenceContactCardUrl));
_COM_SMARTPTR_TYPEDEF(IUccPresenceContactCardInstance, __uuidof(IUccPresenceContactCardInstance));
_COM_SMARTPTR_TYPEDEF(IUccPresenceContactCardInstance2, __uuidof(IUccPresenceContactCardInstance2));
_COM_SMARTPTR_TYPEDEF(IUccPresenceFreeBusyInformation, __uuidof(IUccPresenceFreeBusyInformation));
_COM_SMARTPTR_TYPEDEF(IUccPresenceTimeZoneInformation, __uuidof(IUccPresenceTimeZoneInformation));
_COM_SMARTPTR_TYPEDEF(IUccPresenceWorkingHoursBlock, __uuidof(IUccPresenceWorkingHoursBlock));
_COM_SMARTPTR_TYPEDEF(IUccPresenceWorkingHoursInformation, __uuidof(IUccPresenceWorkingHoursInformation));
_COM_SMARTPTR_TYPEDEF(IUccPresenceCapabilityInfo, __uuidof(IUccPresenceCapabilityInfo));
_COM_SMARTPTR_TYPEDEF(IUccPresenceCapabilities, __uuidof(IUccPresenceCapabilities));
_COM_SMARTPTR_TYPEDEF(IUccPresenceCalendarDataInstance, __uuidof(IUccPresenceCalendarDataInstance));
_COM_SMARTPTR_TYPEDEF(IUccPresenceServicesInstance, __uuidof(IUccPresenceServicesInstance));
_COM_SMARTPTR_TYPEDEF(IUccPresenceServiceElement, __uuidof(IUccPresenceServiceElement));
_COM_SMARTPTR_TYPEDEF(IUccPresenceDeviceInstance, __uuidof(IUccPresenceDeviceInstance));
_COM_SMARTPTR_TYPEDEF(IUccPresenceDeviceInstance2, __uuidof(IUccPresenceDeviceInstance2));
_COM_SMARTPTR_TYPEDEF(IUccPresenceNoteData, __uuidof(IUccPresenceNoteData));
_COM_SMARTPTR_TYPEDEF(IUccPresenceNoteInstance, __uuidof(IUccPresenceNoteInstance));
_COM_SMARTPTR_TYPEDEF(IUccPresencePhoneLine, __uuidof(IUccPresencePhoneLine));
_COM_SMARTPTR_TYPEDEF(IUccRoutingRuleElement, __uuidof(IUccRoutingRuleElement));
_COM_SMARTPTR_TYPEDEF(IUccRoutingRuleInstance, __uuidof(IUccRoutingRuleInstance));
_COM_SMARTPTR_TYPEDEF(IUccRoutingPreambleElement, __uuidof(IUccRoutingPreambleElement));
_COM_SMARTPTR_TYPEDEF(IUccRoutingPreamble, __uuidof(IUccRoutingPreamble));
_COM_SMARTPTR_TYPEDEF(IUccRoutingTarget, __uuidof(IUccRoutingTarget));
_COM_SMARTPTR_TYPEDEF(IUccRoutingPreambleListElement, __uuidof(IUccRoutingPreambleListElement));
_COM_SMARTPTR_TYPEDEF(IUccRoutingPreambleWaitElement, __uuidof(IUccRoutingPreambleWaitElement));
_COM_SMARTPTR_TYPEDEF(IUccRoutingPreambleFlagsElement, __uuidof(IUccRoutingPreambleFlagsElement));
_COM_SMARTPTR_TYPEDEF(IUccServerConfigurationCategory, __uuidof(IUccServerConfigurationCategory));
_COM_SMARTPTR_TYPEDEF(IUccSignalingBodyPart, __uuidof(IUccSignalingBodyPart));
_COM_SMARTPTR_TYPEDEF(IUccSignalingMessage2, __uuidof(IUccSignalingMessage2));
_COM_SMARTPTR_TYPEDEF(IUccSignalingResponse, __uuidof(IUccSignalingResponse));
_COM_SMARTPTR_TYPEDEF(IUccUriManager, __uuidof(IUccUriManager));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingResource, __uuidof(IUccApplicationSharingResource));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingChannelResource, __uuidof(IUccApplicationSharingChannelResource));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingChannelResourceEvent, __uuidof(IUccApplicationSharingChannelResourceEvent));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharerChannelResource, __uuidof(IUccApplicationSharerChannelResource));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingViewer, __uuidof(IUccApplicationSharingViewer));
_COM_SMARTPTR_TYPEDEF(IUccApplicationViewerChannelResource, __uuidof(IUccApplicationViewerChannelResource));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingResourceManager, __uuidof(IUccApplicationSharingResourceManager));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingSession, __uuidof(IUccApplicationSharingSession));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingSessionParticipant, __uuidof(IUccApplicationSharingSessionParticipant));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingChannel, __uuidof(IUccApplicationSharingChannel));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingControlEvent, __uuidof(IUccApplicationSharingControlEvent));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingControlRequestAllowedEvent, __uuidof(IUccApplicationSharingControlRequestAllowedEvent));
_COM_SMARTPTR_TYPEDEF(IUccApplicationViewerChannel, __uuidof(IUccApplicationViewerChannel));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingConferenceChannel, __uuidof(IUccApplicationSharingConferenceChannel));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingChannelResources, __uuidof(IUccApplicationSharingChannelResources));
_COM_SMARTPTR_TYPEDEF(IUccMediaChannelConnectionStateChangedEvent, __uuidof(IUccMediaChannelConnectionStateChangedEvent));
_COM_SMARTPTR_TYPEDEF(_IUccApplicationSharingChannelResourceEvents, __uuidof(_IUccApplicationSharingChannelResourceEvents));
_COM_SMARTPTR_TYPEDEF(_IUccApplicationSharingViewerEvents, __uuidof(_IUccApplicationSharingViewerEvents));
_COM_SMARTPTR_TYPEDEF(_IUccApplicationSharingResourceManagerEvents, __uuidof(_IUccApplicationSharingResourceManagerEvents));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingResourceChangeEvent, __uuidof(IUccApplicationSharingResourceChangeEvent));
_COM_SMARTPTR_TYPEDEF(_IUccApplicationSharingSessionParticipantEvents, __uuidof(_IUccApplicationSharingSessionParticipantEvents));
_COM_SMARTPTR_TYPEDEF(_IUccApplicationSharingChannelEvents, __uuidof(_IUccApplicationSharingChannelEvents));
_COM_SMARTPTR_TYPEDEF(_IUccApplicationSharerChannelEvents, __uuidof(_IUccApplicationSharerChannelEvents));
_COM_SMARTPTR_TYPEDEF(_IUccApplicationViewerChannelEvents, __uuidof(_IUccApplicationViewerChannelEvents));
_COM_SMARTPTR_TYPEDEF(_IUccApplicationSharingChannelResourcesEvents, __uuidof(_IUccApplicationSharingChannelResourcesEvents));
_COM_SMARTPTR_TYPEDEF(_IUccApplicationSharingChannelCollectionEvents, __uuidof(_IUccApplicationSharingChannelCollectionEvents));
_COM_SMARTPTR_TYPEDEF(_IUccPublicationManagerEvents, __uuidof(_IUccPublicationManagerEvents));
_COM_SMARTPTR_TYPEDEF(_IUccCategoryInstanceEvents, __uuidof(_IUccCategoryInstanceEvents));
_COM_SMARTPTR_TYPEDEF(_IUccCategoryContextEvents, __uuidof(_IUccCategoryContextEvents));
_COM_SMARTPTR_TYPEDEF(_IUccContactEvents, __uuidof(_IUccContactEvents));
_COM_SMARTPTR_TYPEDEF(_IUccDelegateEvents, __uuidof(_IUccDelegateEvents));
_COM_SMARTPTR_TYPEDEF(_IUccGroupEvents, __uuidof(_IUccGroupEvents));
_COM_SMARTPTR_TYPEDEF(_IUccSubscriptionEvents, __uuidof(_IUccSubscriptionEvents));
_COM_SMARTPTR_TYPEDEF(_IUccPresentityEvents, __uuidof(_IUccPresentityEvents));
_COM_SMARTPTR_TYPEDEF(_IUccPublicationEvent, __uuidof(_IUccPublicationEvent));
_COM_SMARTPTR_TYPEDEF(_IUccContainerEvents, __uuidof(_IUccContainerEvents));
_COM_SMARTPTR_TYPEDEF(_IUccContainerMemberEvents, __uuidof(_IUccContainerMemberEvents));
_COM_SMARTPTR_TYPEDEF(_IUccOperationManagerEvents, __uuidof(_IUccOperationManagerEvents));
_COM_SMARTPTR_TYPEDEF(_IUccMediaDeviceManagerEvents, __uuidof(_IUccMediaDeviceManagerEvents));
_COM_SMARTPTR_TYPEDEF(IUccSignalingChannelManager, __uuidof(IUccSignalingChannelManager));
_COM_SMARTPTR_TYPEDEF(IUccConferenceMediaChannel, __uuidof(IUccConferenceMediaChannel));
_COM_SMARTPTR_TYPEDEF(IUccSessionParticipantEndpoint, __uuidof(IUccSessionParticipantEndpoint));
_COM_SMARTPTR_TYPEDEF(IUccConferenceApplicationViewerChannel, __uuidof(IUccConferenceApplicationViewerChannel));
_COM_SMARTPTR_TYPEDEF(IUccConferenceSessionParticipant, __uuidof(IUccConferenceSessionParticipant));
_COM_SMARTPTR_TYPEDEF(IUccAudioVideoMediaChannelSourceEvent, __uuidof(IUccAudioVideoMediaChannelSourceEvent));
_COM_SMARTPTR_TYPEDEF(IUccConferenceMediaChannelSourceEvent, __uuidof(IUccConferenceMediaChannelSourceEvent));
_COM_SMARTPTR_TYPEDEF(IUccIncomingInstantMessageEvent, __uuidof(IUccIncomingInstantMessageEvent));
_COM_SMARTPTR_TYPEDEF(IUccInstantMessagingComposingEvent, __uuidof(IUccInstantMessagingComposingEvent));
_COM_SMARTPTR_TYPEDEF(IUccSessionParticipantEndpointCollectionEvent, __uuidof(IUccSessionParticipantEndpointCollectionEvent));
_COM_SMARTPTR_TYPEDEF(IUccSessionParticipantOperationEvent, __uuidof(IUccSessionParticipantOperationEvent));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharerChannel, __uuidof(IUccApplicationSharerChannel));
_COM_SMARTPTR_TYPEDEF(IUccApplicationSharingRequestControlEvent, __uuidof(IUccApplicationSharingRequestControlEvent));
_COM_SMARTPTR_TYPEDEF(IUccSessionParticipant, __uuidof(IUccSessionParticipant));
_COM_SMARTPTR_TYPEDEF(IUccSession, __uuidof(IUccSession));
_COM_SMARTPTR_TYPEDEF(IUccSessionCallControl, __uuidof(IUccSessionCallControl));
_COM_SMARTPTR_TYPEDEF(IUccSessionManager, __uuidof(IUccSessionManager));
_COM_SMARTPTR_TYPEDEF(IUccSessionParticipantCollectionEvent, __uuidof(IUccSessionParticipantCollectionEvent));
_COM_SMARTPTR_TYPEDEF(IUccIncomingSessionEvent, __uuidof(IUccIncomingSessionEvent));
_COM_SMARTPTR_TYPEDEF(IUccOutgoingSessionEvent, __uuidof(IUccOutgoingSessionEvent));
_COM_SMARTPTR_TYPEDEF(IUccSubscriptionManager, __uuidof(IUccSubscriptionManager));
_COM_SMARTPTR_TYPEDEF(IUccSubscription, __uuidof(IUccSubscription));
_COM_SMARTPTR_TYPEDEF(IUccPresentity, __uuidof(IUccPresentity));
_COM_SMARTPTR_TYPEDEF(IUccSubscriptionEvent, __uuidof(IUccSubscriptionEvent));
_COM_SMARTPTR_TYPEDEF(IUccCategoryContext, __uuidof(IUccCategoryContext));
_COM_SMARTPTR_TYPEDEF(IUccCategoryContextEvent, __uuidof(IUccCategoryContextEvent));
_COM_SMARTPTR_TYPEDEF(IUccCategoryInstance, __uuidof(IUccCategoryInstance));
_COM_SMARTPTR_TYPEDEF(IUccCategoryInstanceEvent, __uuidof(IUccCategoryInstanceEvent));
_COM_SMARTPTR_TYPEDEF(IUccPublication, __uuidof(IUccPublication));
_COM_SMARTPTR_TYPEDEF(IUccContainer, __uuidof(IUccContainer));
_COM_SMARTPTR_TYPEDEF(IUccContainerMember, __uuidof(IUccContainerMember));
_COM_SMARTPTR_TYPEDEF(IUccPublicationManager, __uuidof(IUccPublicationManager));
_COM_SMARTPTR_TYPEDEF(IUccContainerMemberCollectionEvent, __uuidof(IUccContainerMemberCollectionEvent));
_COM_SMARTPTR_TYPEDEF(IUccContainerMembershipEvent, __uuidof(IUccContainerMembershipEvent));
_COM_SMARTPTR_TYPEDEF(IUccContainerMembershipManager, __uuidof(IUccContainerMembershipManager));
_COM_SMARTPTR_TYPEDEF(IUccContact, __uuidof(IUccContact));
_COM_SMARTPTR_TYPEDEF(IUccGroup, __uuidof(IUccGroup));
_COM_SMARTPTR_TYPEDEF(IUccContactCollectionEvent, __uuidof(IUccContactCollectionEvent));
_COM_SMARTPTR_TYPEDEF(IUccGroupMembershipEvent, __uuidof(IUccGroupMembershipEvent));

//
// Type library items
//
    const long UCCAM_ALL = 15;
    const long UCCMD_NONE = 0;
    const long UCCMD_BOTH = 3;
    const long UCC_DEFAULT_CONTAINER_ID = 0;
    const long UCC_SELF_CONTAINER_ID = 1;
    const long UCCCMJT_NONE = 0;
    const LPWSTR UCC_NULL_SIP_URI = (wchar_t*) L"sip:null@invalid";
    const LPWSTR UCC_REALM_ALL = (wchar_t*) L"*";
    const long UCCCPAT_NONE = 0;
    const long UCCDRM_ALL = 3;
    const long UCCTWP_ALL_PAGES = 15;
    const long UCCMSC_MIN_AUDIO_LEVEL = 0;
    const long UCCMSC_MAX_AUDIO_LEVEL = 32767;
    const long UCCMSC_MIN_AUDIO_VOLUME = 0;
    const long UCCMSC_MAX_AUDIO_VOLUME = 65535;

enum UCC_AUTHENTICATION_MODES
{
    UCCAM_NONE = 0,
    UCCAM_BASIC = 1,
    UCCAM_DIGEST = 2,
    UCCAM_NTLM = 4,
    UCCAM_KERBEROS = 8
};

enum UCC_CATEGORY_INSTANCE_EXPIRE_TYPE
{
    UCCCIET_NONE = 0,
    UCCCIET_USER = 1,
    UCCCIET_DEVICE = 2,
    UCCCIET_STATIC = 3,
    UCCCIET_TIME = 4
};

enum UCC_CATEGORY_CONTEXT_EVENT_CONTEXT
{
    UCCCCEC_REMOVAL_REASON = 1
};

enum UCC_CATEGORY_REMOVAL_REASON
{
    UCCCRR_RECEIVED_NOTIFY = 1,
    UCCCRR_UNSUBSCRIBED_BY_USER = 2,
    UCCCRR_SUBSCRIPTION_TERMINATED = 3
};

enum UCC_CATEGORY_INSTANCE_EVENT_CONTEXT
{
    UCCCCIC_REMOVAL_REASON = 1
};

enum UCC_CHANNEL_CONTEXT
{
    UCCCC_VIDEO_TYPE = 266241
};

enum UCC_CONFERENCE_ADMISSION_TYPE
{
    UCCCAT_OPEN_AUTHENTICATED = 1,
    UCCCAT_CLOSED_AUTHENTICATED = 2,
    UCCCAT_ANONYMOUS = 3
};

enum UCC_CONFERENCE_ENTITY_PROPERTY
{
    UCCCEP_CAPABILITY = 1153025,
    UCCCEP_STATE = 1153026
};

enum UCC_CONFERENCE_ENTITY_SETTING_PROPERTY
{
    UCCCESP_TYPE = 1511425,
    UCCCESP_PRESENTATION_MODE = 1509378,
    UCCCESP_DATA_MCU_COLLABORATION_MODE = 1509379
};

enum UCC_CONFERENCE_ENTITY_STATE_PROPERTY
{
    UCCCESTP_DISPLAY_TEXT = 1213441,
    UCCCESTP_USER_COUNT = 1215490,
    UCCCESTP_ACTIVE = 1214467,
    UCCCESTP_LOCKED = 1214468,
    UCCCESTP_RECORDING_ID = 1213445,
    UCCCESTP_RECORDING_ACTIVE = 1214470,
    UCCCESTP_RECORDING_ERROR_CODE = 1215495,
    UCCCESTP_RECORDING_ERROR_DESCRIPTION = 1213448,
    UCCCESTP_RECORDING_PAUSED = 1214473,
    UCCCESTP_MEDIA_COLLECTION = 1218570,
    UCCCESTP_MEDIA_FILTERS_RULE_COLLECTION = 1218571,
    UCCCESTP_RECORDED_MEDIA_COLLECTION = 1218572,
    UCCCESTP_ACP_TOLL_NUMBER = 1213453,
    UCCCESTP_ACP_TOLL_FREE_NUMBER = 1213454,
    UCCCESTP_ACP_PARTICIPANT_PASSCODE = 1213455,
    UCCCESTP_ACP_ENTRY_EXIT_ANNOUNCEMENT = 1213456,
    UCCCESTP_AV_VIDEO_MODE = 1216529,
    UCCCESTP_AV_DOMINANT_SPEAKER_SOURCE = 1213458,
    UCCCESTP_AV_INTENDED_PRIME_PRESENTER_SOURCE = 1213459,
    UCCCESTP_AV_INTENDED_SECONDARY_PRESENTER_SOURCE = 1213460,
    UCCCESTP_DATA_MCU_GUID = 1213461
};

enum UCC_CONFERENCE_ENTITY_TYPE
{
    UCCCET_CONTROL = 1,
    UCCCET_AUDIO_VIDEO = 2,
    UCCCET_DATA = 3,
    UCCCET_INSTANT_MESSAGING = 4,
    UCCCET_TELEPHONY = 5,
    UCCCET_APPLICATION_SHARING = 6
};

enum UCC_CONFERENCE_INFORMATION_PROPERTY
{
    UCCCIP_SUBJECT = 1442817,
    UCCCIP_ADMISSION_TYPE = 1445890,
    UCCCIP_EXPIRY_TIME = 1448963,
    UCCCIP_ORGANIZER_ROAMING_DATA = 1442820,
    UCCCIP_NOTIFICATION_DATA = 1442821,
    UCCCIP_PARTICIPANTS = 1447942,
    UCCCIP_MCUS = 1447943,
    UCCCIP_EXTERNAL_WEB_LOCATION = 1442824,
    UCCCIP_INTERNAL_WEB_LOCATION = 1442825,
    UCCCIP_TELEPHONY_ACCESS_ID = 1442826
};

enum UCC_CONFERENCE_MANAGER_OPERATION_COMPLETED_EVENT_PROPERTY
{
    UCCCMOCEP_CONFERENCE_URI = 1606657,
    UCCCMOCEP_CONFERENCE_INFORMATION = 1610754,
    UCCCMOCEP_CONFERENCE_LIST = 1611779,
    UCCCMOCEP_AVAILABLE_MCUS = 1611780
};

enum UCC_CONFERENCE_MEDIA_CHANNEL_PROPERTY
{
    UCCCMCPID_MUTE = 1050625,
    UCCCMCPID_VIDEO_TYPE = 1085442,
    UCCCMCPID_APPLICATION_SHARING_TYPE = 1085443
};

enum UCC_CONFERENCE_MEDIA_FILTERS_RULE_PROPERTY
{
    UCCCMFRP_ROLE = 1347585,
    UCCCMFRP_MAY_MODIFY_OWN_FILTERS = 1345538,
    UCCCMFRP_INGRESS_FILTER = 1344515,
    UCCCMFRP_EGRESS_FILTER = 1344516
};

enum UCC_CONFERENCE_MEDIA_JOIN_TYPE
{
    UCCCMJT_DIAL_IN = 1,
    UCCCMJT_DIAL_OUT = 2
};

enum UCC_CONFERENCE_MEDIUM_PROPERTY
{
    UCCCMP_LABEL = 1278977,
    UCCCMP_DISPLAY_TEXT = 1278978,
    UCCCMP_TYPE = 1282051,
    UCCCMP_STATUS = 1281028,
    UCCCMP_VIDEO_MODE = 1282053,
    UCCCMP_PRIMARY_VIDEO_SOURCE = 1278982
};

enum UCC_CONFERENCE_PARTICIPANT_AUTHENTICATION_TYPE
{
    UCCCPAT_FULL = 1,
    UCCCPAT_PIN = 2,
    UCCCPAT_EXTERNAL = 3,
    UCCCPAT_FEDERATED = 4,
    UCCCPAT_NOT_AUTHED = 5
};

enum UCC_CONFERENCE_PARTICIPANT_ENDPOINT_CONTEXT
{
    UCCCPEC_ROLE = 856065,
    UCCCPEC_CONVERSATION_IDENTITY = 852994,
    UCCCPEC_CONFERENCE_MEDIA_JOIN_PREFERENCE = 856067,
    UCCCPEC_LEADER_PASSCODE = 852996,
    UCCCPEC_POST_DIAL = 852997,
    UCCCPEC_REPLACES = 857094
};

enum UCC_CONFERENCE_PARTICIPANT_ENDPOINT_PROPERTY
{
    UCCCPEPID_ENDPOINT_URI = 820225,
    UCCCPEPID_ROLE = 790530,
    UCCCPEPID_TELEPHONY_ROLE = 790531,
    UCCCPEPID_AUTHENTICATION_TYPE = 823300,
    UCCCPEPID_LEADER_PASSCODE = 820229,
    UCCCPEPID_JOINING_METHOD = 823302,
    UCCCPEPID_JOINING_INFO_BY = 820231,
    UCCCPEPID_JOINING_INFO_WHEN = 820232,
    UCCCPEPID_JOINING_INFO_REASON = 820233,
    UCCCPEPID_ON_BEHALF_OF_URI = 824330
};

enum UCC_CONFERENCE_PARTICIPANT_PROPERTY
{
    UCCCPPID_URI = 951297,
    UCCCPPID_DISPLAY_NAME = 951298,
    UCCCPPID_ROLE = 921603
};

enum UCC_CONFERENCE_PARTICIPANT_ROLE
{
    UCCCPR_UNDEFINED = 0,
    UCCCPR_PARTICIPANT = 1,
    UCCCPR_ADMIN = 2
};

enum UCC_CONFERENCE_RECORDED_MEDIA_PROPERTY
{
    UCCCRMP_TYPE = 1413121,
    UCCCRMP_ERROR_CODE = 1412098,
    UCCCRMP_ERROR_DESCRIPTION = 1410051
};

enum UCC_CONFERENCE_SESSION_PROPERTY
{
    UCCCSPID_MEDIA = 1018881,
    UCCCSPID_ACTIVE = 1017858,
    UCCCSPID_LOCKED = 985091,
    UCCCSPID_RECORDING = 1017860,
    UCCCSPID_SOURCE_NETWORK_TYPE = 1019909,
    UCCCSPID_SUBJECT = 984070,
    UCCCSPID_VIDEO_SWITCHING_MODE = 987143,
    UCCCSPID_PRIMARY_VIDEO_SOURCE = 988168,
    UCCCSPID_CONFERENCE_AUDIO_VIDEO_URI = 1016841,
    UCCCSPID_CONFERENCE_DATA_URI = 1016842,
    UCCCSPID_CONFERENCE_INSTANT_MESSAGING_URI = 1016843,
    UCCCSPID_CONFERENCE_TELEPHONY_URI = 1016844,
    UCCCSPID_CONFERENCE_APPLICATION_SHARING_URI = 1016845,
    UCCCSPID_CONFERENCE_VIEW = 1020942,
    UCCCSPID_NOTIFICATION_DATA = 1016847,
    UCCCSPID_SIGNALING_SUSPENDED = 1017872
};

enum UCC_CONFERENCE_VIDEO_SWITCHING_MODE
{
    UCCCVSM_DOMINANT_SPEAKER_SWITCHED = 1,
    UCCCVSM_MANUAL_SELECTION = 2
};

enum UCC_CONTACT_CARD_ADDRESS_TYPE
{
    UCCCCAT_UNKNOWN = 0,
    UCCCCAT_WORK = 1,
    UCCCCAT_HOME = 2,
    UCCCCAT_OTHER = 3
};

enum UCC_CONTACT_CARD_PHONE_TYPE
{
    UCCCCPT_UNKNOWN = 0,
    UCCCCPT_WORK = 1,
    UCCCCPT_HOME = 2,
    UCCCCPT_FAX = 3,
    UCCCCPT_MOBILE = 4,
    UCCCCPT_OTHER = 5
};

enum UCC_CONTACT_CARD_URL_TYPE
{
    UCCCCUT_UNKNOWN = 0,
    UCCCCUT_SHAREPOINT = 1,
    UCCCCUT_VOICEMAIL = 2,
    UCCCCUT_OTHER = 3
};

enum UCC_CONTAINER_MEMBER_COMPARE_RESULT
{
    UCCCMCR_SAME = 0,
    UCCCMCR_SUPERSET = 1,
    UCCCMCR_SUBSET = 2,
    UCCCMCR_DIFFERENT = 3
};

enum UCC_CONTAINER_MEMBER_MATCHING_TYPE
{
    UCCCMMT_EXACT = 1,
    UCCCMMT_BEST = 2,
    UCCCMMT_EXCLUDE_EXACT = 3
};

enum UCC_CONTAINER_MEMBERSHIP_SCOPE
{
    UCCCMS_USER = 0,
    UCCCMS_DOMAIN = 1,
    UCCCMS_COMPANY = 2,
    UCCCMS_FEDERATED_ENTERPRISE = 3,
    UCCCMS_FEDERATED_PUBLIC = 4,
    UCCCMS_EVERYONE = 5
};

enum UCC_DIAGNOSTIC_INFO_TYPE
{
    UCCDIT_SERVER_DIAGNOSTIC = 0,
    UCCDIT_SERVER_DIAGNOSTIC_PUBLIC = 1,
    UCCDIT_CLIENT_DIAGNOSTIC = 2,
    UCCDIT_WARNING = 3,
    UCCDIT_STATUS = 4,
    UCCDIT_HRESULT = 5
};

enum UCC_DTMF
{
    UCCDTMF_0 = 0,
    UCCDTMF_1 = 1,
    UCCDTMF_2 = 2,
    UCCDTMF_3 = 3,
    UCCDTMF_4 = 4,
    UCCDTMF_5 = 5,
    UCCDTMF_6 = 6,
    UCCDTMF_7 = 7,
    UCCDTMF_8 = 8,
    UCCDTMF_9 = 9,
    UCCDTMF_STAR = 10,
    UCCDTMF_POUND = 11,
    UCCDTMF_A = 12,
    UCCDTMF_B = 13,
    UCCDTMF_C = 14,
    UCCDTMF_D = 15,
    UCCDTMF_FLASH = 16
};

enum UCC_ENDPOINT_SIGNALING_CAPABILITIES
{
    UCCESC_NONE = 0,
    UCCESC_SESSION_ANSWER = 1,
    UCCESC_SESSION_HOLD = 2,
    UCCESC_SESSION_UNHOLD = 4,
    UCCESC_SESSION_FORWARD = 8,
    UCCESC_SESSION_CONFERENCE = 16,
    UCCESC_SESSION_BLIND_TRANSFER = 32,
    UCCESC_SESSION_CONSULTATIVE_TRANSFER = 64,
    UCCESC_SESSION_ALTERNATE = 128,
    UCCESC_SESSION_RECONNECT = 256,
    UCCESC_SESSION_GENERATE_DTMF = 512
};

enum UCC_FIND_MEDIA_CONNECTIVITY_SERVER_REQUEST_CONTEXT
{
    UCCFMCSR_LOCATION = 1
};

enum UCC_GROUP_TYPE
{
    UCCGT_USER_DEFINED_GROUP = 0,
    UCCGT_DISTRIBUTION_GROUP = 1
};

enum UCC_INCOMING_REQUEST_EVENT_CONTEXT
{
    UCCIRECE_SIGNALING_REQUEST = 660481,
    UCCIRECE_SIGNALING_RESPONSE = 660482
};

enum UCC_MEDIA_CONNECTIVITY_SERVER_CONFIGURATION_CONTEXT
{
    UCCMCSC_LOCATION = 1,
    UCCMCSC_DURATION = 2
};

enum UCC_MEDIA_CONNECTIVITY_SERVER_LOCATION
{
    UCCMCSL_UNSPECIFIED = 0,
    UCCMCSL_INTRANET = 1,
    UCCMCSL_INTERNET = 2
};

enum UCC_MEDIA_CONNECTIVITY_SERVER_TYPE
{
    UCCMCST_TURN = 1,
    UCCMCST_HTTP_PROXY = 2
};

enum UCC_MEDIA_DIRECTIONS
{
    UCCMD_SEND = 1,
    UCCMD_RECEIVE = 2
};

enum UCC_MEDIA_ENCRYPTION_LEVEL
{
    UCCMEL_NONE = 1,
    UCCMEL_OPTIONAL = 2,
    UCCMEL_REQUIRED = 3
};

enum UCC_MEDIA_ENCRYPTION_MODES
{
    UCCMEM_DES = 1,
    UCCMEM_SRTP = 2
};

enum UCC_MEDIA_STREAM_STATE
{
    UCCMSS_STARTED = 1,
    UCCMSS_ACTIVE = 2,
    UCCMSS_INACTIVE = 3,
    UCCMSS_STOPPED = 4
};

enum UCC_MEDIA_CHANNEL_CONNECTION_STATE
{
    UCCMCCS_IDLE = 0,
    UCCMCCS_CONNECTED = 1,
    UCCMCCS_DISCONNECTED = 2
};

enum UCC_MEDIA_STREAM_STATE_CHANGE_REASON
{
    UCCMSSCR_NORMAL = 0,
    UCCMSSCR_TIMEOUT = 1,
    UCCMSSCR_NO_DEVICE = 2,
    UCCMSSCR_NO_PORT = 3,
    UCCMSSCR_PORT_MAPPING_FAILURE = 4,
    UCCMSSCR_REMOTE_REQUEST = 5,
    UCCMSSCR_MEDIA_CONNECTIVITY_FAILURE = 6,
    UCCMSSCR_NETWORK_BANDWIDTH_LIMITATION = 7,
    UCCMSSCR_DEVICE_ERROR = 8,
    UCCMSSCR_ERROR = 9
};

enum UCC_MEDIA_TYPES
{
    UCCMT_AUDIO = 1,
    UCCMT_VIDEO = 2,
    UCCMT_MESSAGE = 4,
    UCCMT_TELEPHONY = 8,
    UCCMT_DATA = 16,
    UCCMT_APPLICATION_SHARING = 32
};

enum UCC_OPERATION_PROGRESS_EVENT_CONTEXT
{
    UCCOPEC_DIAGNOSTIC_DETAILS = 365569,
    UCCOPEC_REDIRECT_TARGETS = 365570,
    UCCOPEC_RETRY_AFTER = 363523,
    UCCOPEC_SIGNALING_RESPONSE = 365572,
    UCCOPEC_ACCEPTED_USER_URI = 365573,
    UCCOPEC_HISTORY_INFORMATION = 365574,
    UCCOPEC_SIGNALING_REQUEST = 365575,
    UCCOPEC_CONFERENCE_INFORMATION = 365576,
    UCCOPEC_REDIRECTED_TO_SAME_PARTICIPANT = 362505
};

enum UCC_OPERATION_CONTEXT
{
    UCCOCE_SIGNALING_REQUEST = 398337,
    UCCOCE_SIGNALING_RESPONSE = 398338
};

enum UCC_PARTICIPANT_STATE_CHANGE_REASON
{
    UCCPSCR_NORMAL = 0,
    UCCPSCR_RINGING = 1,
    UCCPSCR_TRYING = 2,
    UCCPSCR_REMOTE_HANGUP_ACTION = 3,
    UCCPSCR_FORWARD_DISCONNECT = 4,
    UCCPSCR_REDIRECT_ACTION = 5,
    UCCPSCR_CONFERENCE_ACTION = 6,
    UCCPSCR_TRANSFER_ACTION = 7,
    UCCPSCR_FAILED = 8,
    UCCPSCR_ERROR = 9,
    UCCPSCR_ENDPOINT_NOT_EXIST = 10,
    UCCPSCR_AUTOMATIC_RETRY = 11,
    UCCPSCR_SUSPENDED = 12
};

enum UCC_PLATFORM_CONTEXT
{
    UCCPC_DISABLE_MEDIA = 67585,
    UCCPC_SUPPORT_HID = 67586
};

enum UCC_PRESENCE_AVAILABILITY
{
    UCCPA_NOT_SET = 0,
    UCCPA_FREE = 3000,
    UCCPA_IDLE_FREE = 4500,
    UCCPA_BUSY = 6000,
    UCCPA_IDLE_BUSY = 7500,
    UCCPA_DO_NOT_DISTURB = 9000,
    UCCPA_TEMP_UNALERTABLE = 12000,
    UCCPA_UNALERTABLE = 15000,
    UCCPA_UNAVAILABLE = 18000
};

enum UCC_PRESENCE_DEVICE_CAPABILITIES
{
    UCCDP_NONE = 0,
    UCCDP_CAN_RENDER = 1,
    UCCDP_CAN_CAPTURE = 2
};

enum UCC_PRESENCE_NOTE_TYPE
{
    UCCPNT_PERSONAL = 0,
    UCCPNT_OOF = 1,
    UCCPNT_UNKNOWN = 2
};

enum UCC_PRESENCE_WORKING_HOURS_DAYS
{
    UCCPWHD_SUNDAY = 1,
    UCCPWHD_MONDAY = 2,
    UCCPWHD_TUESDAY = 4,
    UCCPWHD_WEDNESDAY = 8,
    UCCPWHD_THURSDAY = 16,
    UCCPWHD_FRIDAY = 32,
    UCCPWHD_SATURDAY = 64
};

enum UCC_PRESENCE_WORKING_HOURS_TIMEZONE_TYPE
{
    UCCPWHTT_STANDARD = 0,
    UCCPWHTT_DAYLIGHT = 1
};

enum UCC_ENDPOINT_TYPE
{
    UCCET_PRINCIPAL_SERVER_BASED = 1,
    UCCET_PROXY_TELEPHONY = 65536,
    UCCET_PROXY_DELEGATE = 65537,
    UCCET_PROXY_ASSOCIATED = 65538
};

enum UCC_PUBLICATION_OPERATION_TYPE
{
    UCCPOT_NONE = 0,
    UCCPOT_ADD = 1,
    UCCPOT_REMOVE = 2
};

enum UCC_REJECT_OR_TERMINATE_REASON
{
    UCCROTR_NOT_SPECIFIED = 0,
    UCCROTR_ACCEPTED_ELSEWHERE = 1,
    UCCROTR_FORWARDED = 303,
    UCCROTR_TIMEOUT = 408,
    UCCROTR_UNAVAILABLE = 480,
    UCCROTR_BUSY = 486,
    UCCROTR_NOT_ACCEPTABLE = 488,
    UCCROTR_DECLINE = 603,
    UCCROTR_DECLINE_EVERYWHERE = 605
};

enum UCC_RETARGET_REASON
{
    UCCRR_NONE = 0,
    UCCRR_FORWARDING = 1,
    UCCRR_TEAM_CALL = 2,
    UCCRR_DELEGATION = 3,
    UCCRR_HUNT_GROUP = 4,
    UCCRR_REFEREE = 5,
    UCCRR_REFERER = 6,
    UCCRR_OTHERS = 7
};

enum UCC_SERVER_ENDPOINT_CONTEXT
{
    UCCSEC_SERVER_ENDPOINT_FLAGS = 135169,
    UCCSEC_APPLICATION_VERSION = 132098,
    UCCSEC_DISPLAY_NAME = 132099,
    UCCSEC_REMOTE_REGISTRATION = 165892,
    UCCSEC_SERVER_VERSION = 164869,
    UCCSEC_GRUU = 168966,
    UCCSEC_HTTPS_PROXY_ADDRESS = 164871,
    UCCSEC_LOCAL_NETWORK_ADDRESS = 164872,
    UCCSEC_SERVER_NETWORK_ADDRESS = 164873,
    UCCSEC_AUTOMATIC_PROXY_CREATION = 133130,
    UCCSEC_SIGNALING_SUSPENDED = 165899
};

enum UCC_SERVER_ENDPOINT_FLAGS
{
    UCCSEF_DISABLE_REGISTRATION = 1,
    UCCSEF_DISABLE_SUBSCRIPTIONS = 2,
    UCCSEF_DISABLE_SESSIONS = 4,
    UCCSEF_DISABLE_STRICT_DNS = 8,
    UCCSEF_ENABLE_EVENT_LOGGING = 16,
    UCCSEF_DISABLE_HTTP_FALLBACK = 32
};

enum UCC_SESSION_CONTEXT
{
    UCCSC_DISABLE_MEDIA_CONNECTIVITY = 198657,
    UCCSC_PREFER_DIRECT_MEDIA_CONNECTIVITY = 198658,
    UCCSC_DISABLE_AUTOMATIC_REDIRECT = 198659,
    UCCSC_TRANSFER_CONTEXT = 201732,
    UCCSC_ASSOCIATED_CONFERENCE_SESSION = 201733,
    UCCSC_AUTOMATIC_SESSION_RETRY_PREFERENCE = 200710,
    UCCSC_REPLACES_SESSION = 201735
};

enum UCC_SESSION_REDIRECT_METHOD
{
    UCCSRM_CLIENT_REDIRECT = 302,
    UCCSRM_SERVER_REDIRECT = 303
};

enum UCC_SESSION_TYPE
{
    UCCST_INSTANT_MESSAGING = 0,
    UCCST_AUDIO_VIDEO = 1,
    UCCST_CONFERENCE = 2,
    UCCST_APPLICATION = 3,
    UCCST_APPLICATION_SHARING = 4,
    UCCST_TPCP_SERVER_CONVERSATION = 5,
    UCCST_DATA_SHARING = 6
};

enum UCC_SESSION_ENTITY_STATE
{
    UCCSES_IDLE = 1,
    UCCSES_CONNECTING = 2,
    UCCSES_CONNECTED = 3,
    UCCSES_DISCONNECTING = 4,
    UCCSES_DISCONNECTED = 5
};

enum UCC_SIGNALING_MESSAGE_TYPE
{
    UCCSMT_REQUEST = 0,
    UCCSMT_RESPONSE = 1
};

enum UCC_SIGNALING_REQUEST_TYPE
{
    UCCSRT_UNKNOWN = 0,
    UCCSRT_ACK = 1,
    UCCSRT_BYE = 2,
    UCCSRT_CANCEL = 3,
    UCCSRT_INFO = 4,
    UCCSRT_INVITE = 5,
    UCCSRT_MESSAGE = 6,
    UCCSRT_NOTIFY = 7,
    UCCSRT_OPTIONS = 8,
    UCCSRT_REFER = 9,
    UCCSRT_REGISTER = 10,
    UCCSRT_SERVICE = 11,
    UCCSRT_SUBSCRIBE = 12,
    UCCSRT_UPDATE = 13
};

enum UCC_SOURCE_NETWORK_TYPE
{
    UCCSNT_COMPANY = 2,
    UCCSNT_FEDERATED_ENTERPRISE = 3,
    UCCSNT_FEDERATED_PUBLIC = 4,
    UCCSNT_UNSPECIFIED = 5
};

enum UCC_TRANSFER_CONTEXT
{
    UCCTC_TRANSFERRING_PARTICIPANT = 463873,
    UCCTC_TRANSFER_PROGRESS_NOTIFIER = 463874,
    UCCTC_TRANSFER_TARGET_REQUEST_SIGNALING_MESSAGE = 463875
};

enum UCC_TRANSPORT_MODE
{
    UCCTM_UDP = 1,
    UCCTM_TCP = 2,
    UCCTM_TLS = 3
};

enum UCC_VIDEO_TYPE
{
    UCCVT_NORMAL = 0,
    UCCVT_PANORAMIC = 1
};

enum UCC_MEDIA_DEVICE_SETTING_PROPERTY
{
    UCCMDSP_PREFERRED_HANDSET_AUDIO_CAPTURE = 525313,
    UCCMDSP_PREFERRED_HANDSET_AUDIO_RENDER = 525314,
    UCCMDSP_PREFERRED_HANDSET_AUDIO_IS_MANUAL = 526339,
    UCCMDSP_PREFERRED_AUDIO_CAPTURE = 525316,
    UCCMDSP_PREFERRED_AUDIO_RENDER = 525317,
    UCCMDSP_PREFERRED_AUDIO_IS_MANUAL = 526342,
    UCCMDSP_PREFERRED_NOTIFICATION_AUDIO_RENDER = 525319,
    UCCMDSP_PREFERRED_NOTIFICATION_AUDIO_IS_MANUAL = 526344,
    UCCMDSP_PREFERRED_VIDEO_CAPTURE = 525321,
    UCCMDSP_PREFERRED_VIDEO_IS_MANUAL = 526346,
    UCCMDSP_AUDIO_CAPTURE_VOLUME_COLLECTION = 530443,
    UCCMDSP_AUDIO_RENDER_VOLUME_COLLECTION = 530444
};

enum UCC_TUNING_WIZARD_PAGE
{
    UCCTWP_AUDIO = 1,
    UCCTWP_AUDIO_NONPRIVATE = 2,
    UCCTWP_AUDIO_NOTIFICATION = 4,
    UCCTWP_WEBCAM = 8
};

enum UCC_MEDIA_DEVICE_PROPERTY
{
    UCCMDP_TELEPHONY_DEVICE_ID = 623617,
    UCCMDP_TELEPHONY_DEVICE_NAME = 623618,
    UCCMDP_TELEPHONY_DEVICE_SUB_TYPE = 626691,
    UCCMDP_PRODUCT_VENDOR_ID = 623620,
    UCCMDP_DEVICE_PNP_NAME = 623621,
    UCCMDP_AUDIO_DEVICE_WAVE_ID = 625670,
    UCCMDP_AUDIO_DEVICE_DSOUND_GUID = 623623
};

enum UCC_SUBSCRIPTION_CONTEXT
{
    UCCSC_RETRY_TIMEOUT = 1,
    UCCSC_THROTTLE_LIMIT = 2
};

enum UCC_PROXY_ENDPOINT_CONTEXT_EXTENSIONS
{
    UCCPECE_PRINCIPAL_SESSION = 10000
};

enum UCC_USER_SEARCH_COLUMN
{
    UCCUSC_URI = 0,
    UCCUSC_DISPLAYNAME = 1,
    UCCUSC_TITLE = 2,
    UCCUSC_OFFICE = 3,
    UCCUSC_PHONE = 4,
    UCCUSC_COMPANY = 5,
    UCCUSC_CITY = 6,
    UCCUSC_STATE = 7,
    UCCUSC_COUNTRY = 8,
    UCCUSC_EMAIL = 9
};

enum UCC_USER_SEARCH_PREFERENCE
{
    UCCUSP_MAX_MATCHES = 0,
    UCCUSP_TIME_LIMIT = 1
};

enum UCC_PROXY_ENDPOINT_COLLECTION_REASON
{
    UCCPECR_USER_REQUEST = 1,
    UCCPECR_SERVER_REQUEST = 2
};

enum UCC_DELEGATE_PRIVILEGES
{
    UCCDP_CALLS = 1,
    UCCDP_PUBLISH = 2,
    UCCDP_SUBSCRIBE = 4,
    UCCDP_REDELEGATE = 8
};

enum UCC_SESSION_RETRY_PREFERENCE
{
    UCCSRP_NO_RETRY = 1,
    UCCSRP_CURRENT_VERSION = 2,
    UCCSRP_LEGACY_VERSION = 3
};

enum UCC_APPLICATION_SHARING_RESOURCE_TYPE
{
    UASRT_MONITOR = 0,
    UASRT_DESKTOP = 1,
    UASRT_PROCESS = 2,
    UASRT_WINDOW = 3
};

enum UCC_APPLICATION_SHARING_RESOURCE_PROPERTY
{
    UASRP_RESOURCE_NAME = 0,
    UASRP_RESOURCE_ID = 1,
    UASRP_RESOURCE_LEFT = 2,
    UASRP_RESOURCE_TOP = 3,
    UASRP_RESOURCE_BOTTOM = 4,
    UASRP_RESOURCE_RIGHT = 5
};

enum UCC_APPLICATION_SHARING_RESOURCE_STATUS
{
    UASRS_ACTIVE = 0,
    UASRS_INACTIVE = 1
};

enum UCC_APPLICATION_SHARING_CHANNEL_TYPE
{
    UASCT_SHARER = 0,
    UASCT_VIEWER = 1
};

enum UCC_SESSION_PARTICIPANT_CONTEXT
{
    UCCSPC_ON_BEHALF_OF = 1803265,
    UCCSPC_SIGNALING_SUSPENDED = 1804290
};

enum UCC_BOUNDING_RECTANGLE_PROPERTY
{
    UCCBRP_WIDTH = 0,
    UCCBRP_HEIGHT = 1
};

struct __declspec(uuid("0814cabe-959d-47b4-841e-5b77b06d6046"))
IUccCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * lCount ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long Index,
        /*[out,retval]*/ VARIANT * pVariant ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * ppVariantEnumerator ) = 0;
};

struct __declspec(uuid("cadc8c3e-857e-4898-9083-8750559376af"))
IUccProperty : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ long * pPropertyId ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * bstrPropertyName ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ VARIANT * pvPropertyValue ) = 0;
      virtual HRESULT __stdcall put_Value (
        /*[in]*/ VARIANT pvPropertyValue ) = 0;
      virtual HRESULT __stdcall get_StringValue (
        /*[out,retval]*/ BSTR * pbstrPropertyValue ) = 0;
      virtual HRESULT __stdcall put_StringValue (
        /*[in]*/ BSTR pbstrPropertyValue ) = 0;
      virtual HRESULT __stdcall get_NumericValue (
        /*[out,retval]*/ long * plPropertyValue ) = 0;
      virtual HRESULT __stdcall put_NumericValue (
        /*[in]*/ long plPropertyValue ) = 0;
      virtual HRESULT __stdcall get_BoolValue (
        /*[out,retval]*/ VARIANT_BOOL * pbPropertyValue ) = 0;
      virtual HRESULT __stdcall put_BoolValue (
        /*[in]*/ VARIANT_BOOL pbPropertyValue ) = 0;
      virtual HRESULT __stdcall get_ObjectValue (
        /*[out,retval]*/ IUnknown * * ppPropertyValue ) = 0;
      virtual HRESULT __stdcall put_ObjectValue (
        /*[in]*/ IUnknown * ppPropertyValue ) = 0;
};

struct __declspec(uuid("e97197bf-9939-4ab5-9056-5a1fdbd15390"))
IUccReadOnlyPropertyCollection : IUccCollection
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Property (
        /*[in]*/ long lPropertyId,
        /*[out,retval]*/ struct IUccProperty * * ppProperty ) = 0;
      virtual HRESULT __stdcall IsPropertySet (
        /*[in]*/ long lPropertyId,
        /*[out,retval]*/ VARIANT_BOOL * pbSet ) = 0;
      virtual HRESULT __stdcall get_NamedProperty (
        /*[in]*/ BSTR bstrPropertyId,
        /*[out,retval]*/ struct IUccProperty * * ppProperty ) = 0;
      virtual HRESULT __stdcall IsNamedPropertySet (
        /*[in]*/ BSTR bstrPropertyName,
        /*[out,retval]*/ VARIANT_BOOL * pbSet ) = 0;
};

struct __declspec(uuid("73fe78fc-efe6-422b-b9dd-4a33d9ed6395"))
IUccPropertyCollection : IUccReadOnlyPropertyCollection
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall AddProperty (
        /*[in]*/ long lPropertyId,
        /*[in]*/ VARIANT vValue,
        /*[out,retval]*/ struct IUccProperty * * ppProperty ) = 0;
      virtual HRESULT __stdcall RemoveProperty (
        /*[in]*/ long lPropertyId ) = 0;
      virtual HRESULT __stdcall AddNamedProperty (
        /*[in]*/ BSTR bstrPropertyName,
        /*[in]*/ VARIANT vValue,
        /*[out,retval]*/ struct IUccProperty * * ppProperty ) = 0;
      virtual HRESULT __stdcall RemoveNamedProperty (
        /*[in]*/ BSTR bstrPropertyId ) = 0;
};

struct __declspec(uuid("e66a1e6b-5c1c-40b6-a3ad-23c3699b9658"))
IUccContext : IUccPropertyCollection
{};

struct __declspec(uuid("afe09299-87d8-4737-ad38-fa6d78537f4a"))
IUccOperationContext : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall Initialize (
        /*[in]*/ LONG_PTR_1 lOperationId,
        /*[in]*/ struct IUccContext * pOperation ) = 0;
      virtual HRESULT __stdcall get_OperationId (
        /*[out,retval]*/ LONG_PTR_1 * plOperationId ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * ppContext ) = 0;
};

struct __declspec(uuid("e8909d14-268c-4cb9-a9bf-48818f5c48bc"))
IUccApplicationSessionParticipant : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall ReNegotiate (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("307c5595-c51d-414b-bccf-38c0cd227f52"))
IUccArray : IUccCollection
{};

struct __declspec(uuid("5c3ed734-2a86-4b9a-8a03-837d4342f5d3"))
IUccAudioMediaChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Volume (
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enDirection,
        /*[out,retval]*/ long * plVolume ) = 0;
      virtual HRESULT __stdcall put_Volume (
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enDirection,
        /*[in]*/ long plVolume ) = 0;
      virtual HRESULT __stdcall get_IsMuted (
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enDirection,
        /*[out,retval]*/ VARIANT_BOOL * pbIsMuted ) = 0;
      virtual HRESULT __stdcall put_IsMuted (
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enDirection,
        /*[in]*/ VARIANT_BOOL pbIsMuted ) = 0;
      virtual HRESULT __stdcall get_SilenceSuppression (
        /*[out,retval]*/ VARIANT_BOOL * pvbEnable ) = 0;
      virtual HRESULT __stdcall put_SilenceSuppression (
        /*[in]*/ VARIANT_BOOL pvbEnable ) = 0;
      virtual HRESULT __stdcall get_SignalLevel (
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enDirection,
        /*[out,retval]*/ long * plSignalLevel ) = 0;
      virtual HRESULT __stdcall SendDtmf (
        /*[in]*/ enum UCC_DTMF enDtmf ) = 0;
};

struct __declspec(uuid("a98332d2-2d80-4759-b8a8-cc5f4c79f460"))
IUccAudioMediaDevice : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Volume (
        /*[out,retval]*/ long * plVolume ) = 0;
      virtual HRESULT __stdcall put_Volume (
        /*[in]*/ long plVolume ) = 0;
      virtual HRESULT __stdcall get_IsMuted (
        /*[out,retval]*/ VARIANT_BOOL * pbIsMuted ) = 0;
      virtual HRESULT __stdcall put_IsMuted (
        /*[in]*/ VARIANT_BOOL pbIsMuted ) = 0;
};

struct __declspec(uuid("d0283fc6-4ab7-4d63-bb24-8cdd72fd897a"))
IUccAudioVideoSession : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall SetEncryptionParameters (
        /*[in]*/ enum UCC_MEDIA_ENCRYPTION_LEVEL enEncryptionLevel,
        /*[in]*/ long lEncryptionMode ) = 0;
      virtual HRESULT __stdcall GetEncryptionParameters (
        /*[out]*/ enum UCC_MEDIA_ENCRYPTION_LEVEL * penEncryptionLevel,
        /*[out]*/ long * plnEncryptionMode ) = 0;
      virtual HRESULT __stdcall PublishSessionMetrics (
        /*[in]*/ BSTR bstrMetricServer,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall get_MediaConnectivityDisabled (
        /*[out,retval]*/ VARIANT_BOOL * pvbDisableICE ) = 0;
      virtual HRESULT __stdcall put_MediaConnectivityDisabled (
        /*[in]*/ VARIANT_BOOL pvbDisableICE ) = 0;
      virtual HRESULT __stdcall get_AcousticEchoCancellationEnabled (
        /*[out,retval]*/ VARIANT_BOOL * pvbEnableAEC ) = 0;
      virtual HRESULT __stdcall put_AcousticEchoCancellationEnabled (
        /*[in]*/ VARIANT_BOOL pvbEnableAEC ) = 0;
      virtual HRESULT __stdcall get_MaximumBandwidth (
        /*[out,retval]*/ unsigned long * pdwMaximumBitsPerSecond ) = 0;
      virtual HRESULT __stdcall put_MaximumBandwidth (
        /*[in]*/ unsigned long pdwMaximumBitsPerSecond ) = 0;
};

struct __declspec(uuid("faff2612-3bf1-44aa-acb6-80d39c0d991d"))
IUccMediaChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_MediaType (
        /*[out,retval]*/ enum UCC_MEDIA_TYPES * penMediaType ) = 0;
      virtual HRESULT __stdcall get_PreferredMedia (
        /*[out,retval]*/ long * plMediaDirections ) = 0;
      virtual HRESULT __stdcall put_PreferredMedia (
        /*[in]*/ long plMediaDirections ) = 0;
      virtual HRESULT __stdcall get_NegotiatedMedia (
        /*[out,retval]*/ long * plMediaDirections ) = 0;
};

struct __declspec(uuid("c3c42eac-9d5c-4daa-abcc-84019d3706ad"))
IUccAudioVideoSessionParticipant : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateChannel (
        /*[in]*/ enum UCC_MEDIA_TYPES enMediaType,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccMediaChannel * * ppChannel ) = 0;
      virtual HRESULT __stdcall AddChannel (
        /*[in]*/ struct IUccMediaChannel * pChannel ) = 0;
      virtual HRESULT __stdcall RemoveChannel (
        /*[in]*/ struct IUccMediaChannel * pChannel ) = 0;
      virtual HRESULT __stdcall get_Channels (
        /*[out,retval]*/ struct IUccCollection * * ppChannels ) = 0;
      virtual HRESULT __stdcall UpdateChannels (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall get_EncryptionModes (
        /*[out,retval]*/ long * plEncryptionMode ) = 0;
      virtual HRESULT __stdcall get_EncryptionLevel (
        /*[out,retval]*/ enum UCC_MEDIA_ENCRYPTION_LEVEL * penEncryptionLevel ) = 0;
};

enum UCC_URI_SCHEME
{
    UCCUS_SIP = 1,
    UCCUS_TEL = 2
};

enum UCC_URI_COMPARISON_OPTION
{
    UCCUCO_DEFAULT = 0,
    UCCUCO_EXACT_MATCH = 1,
    UCCUCO_USER_HOST = 2,
    UCCUCO_CONTAINED = 3
};

struct __declspec(uuid("8b55d477-d96e-4670-b0a9-f1653b9d97ce"))
IUccUri : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Scheme (
        /*[out,retval]*/ enum UCC_URI_SCHEME * penScheme ) = 0;
      virtual HRESULT __stdcall get_User (
        /*[out,retval]*/ BSTR * pbstrUser ) = 0;
      virtual HRESULT __stdcall get_Host (
        /*[out,retval]*/ BSTR * pbstrHost ) = 0;
      virtual HRESULT __stdcall get_Port (
        /*[out,retval]*/ long * plPort ) = 0;
      virtual HRESULT __stdcall get_Transport (
        /*[out,retval]*/ enum UCC_TRANSPORT_MODE * penTransport ) = 0;
      virtual HRESULT __stdcall get_UserAtHost (
        /*[out,retval]*/ BSTR * pbstrUserAtHost ) = 0;
      virtual HRESULT __stdcall get_AddressOfRecord (
        /*[out,retval]*/ BSTR * pbstrAOR ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ BSTR * pbstrValue ) = 0;
      virtual HRESULT __stdcall get_Parameters (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * ppParameters ) = 0;
      virtual HRESULT __stdcall get_HeaderFields (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * ppHeaders ) = 0;
      virtual HRESULT __stdcall FindHeaderFields (
        /*[in]*/ BSTR bstrHeaderName,
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * ppSignallingHeaders ) = 0;
      virtual HRESULT __stdcall IsEqual (
        /*[in]*/ struct IUccUri * pSipUri,
        /*[in]*/ enum UCC_URI_COMPARISON_OPTION enOption,
        /*[out,retval]*/ VARIANT_BOOL * pfIsEqual ) = 0;
};

struct __declspec(uuid("c2a50530-5635-41b5-8d30-d406489238fa"))
IUccConferenceEntityView : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_CONFERENCE_ENTITY_TYPE * penType ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * pVal ) = 0;
};

struct __declspec(uuid("65d83c13-b0a2-400c-bfd8-bcc4890b7574"))
IUccConferenceInformation : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ConferenceId (
        /*[out,retval]*/ BSTR * pbstrConfId ) = 0;
      virtual HRESULT __stdcall get_ConferenceUri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccPropertyCollection * * ppVal ) = 0;
      virtual HRESULT __stdcall get_ConferenceKey (
        /*[out,retval]*/ BSTR * pbstrConfKey ) = 0;
      virtual HRESULT __stdcall put_ConferenceKey (
        /*[in]*/ BSTR pbstrConfKey ) = 0;
};

struct __declspec(uuid("052e68e9-3ea7-4b9c-972c-91013f044af7"))
IUccConferenceManagerSession : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateConferenceInformation (
        /*[out,retval]*/ struct IUccConferenceInformation * * ppConfInfo ) = 0;
      virtual HRESULT __stdcall ScheduleConference (
        /*[in]*/ struct IUccConferenceInformation * pConfInfo,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall GetConference (
        /*[in]*/ BSTR bstrConfId,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall GetConferenceList (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall ModifyConference (
        /*[in]*/ struct IUccConferenceInformation * pConfInfo,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall DeleteConference (
        /*[in]*/ BSTR bstrConfId,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall GetAvailableMcuList (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("9dcfc8ab-4b32-4078-8414-a59de529b327"))
IUccConferenceManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateConferenceManagerSession (
        /*[in]*/ struct IUccUri * pUri,
        /*[out,retval]*/ struct IUccConferenceManagerSession * * ppConfMgrSession ) = 0;
};

struct __declspec(uuid("a8c81cfe-853d-413a-ab91-d5b1c6ee1099"))
IUccConferenceView : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Entities (
        /*[out,retval]*/ struct IUccCollection * * pEntities ) = 0;
};

struct __declspec(uuid("f8c1eee0-4272-4dd6-8986-786358e6a0e3"))
IUccConferenceSession : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * pVal ) = 0;
      virtual HRESULT __stdcall Enter (
        /*[in]*/ struct IUccUri * pUri,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall Leave (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall SetProperty (
        /*[in]*/ enum UCC_CONFERENCE_SESSION_PROPERTY enPropertyId,
        /*[in]*/ VARIANT vPropertyValue,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("48461087-92f1-43bc-9d42-ae615fc39771"))
IUccConferenceSessionParticipantEndpoint : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ enum UCC_SESSION_ENTITY_STATE * penState ) = 0;
      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_CONFERENCE_ENTITY_TYPE * penType ) = 0;
      virtual HRESULT __stdcall get_Channels (
        /*[out,retval]*/ struct IUccCollection * * ppChannels ) = 0;
      virtual HRESULT __stdcall CreateChannel (
        /*[in]*/ enum UCC_MEDIA_TYPES enMediaType,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccMediaChannel * * ppChannel ) = 0;
      virtual HRESULT __stdcall AddChannel (
        /*[in]*/ struct IUccMediaChannel * pChannel ) = 0;
      virtual HRESULT __stdcall RemoveChannel (
        /*[in]*/ struct IUccMediaChannel * pChannel ) = 0;
      virtual HRESULT __stdcall UpdateChannels (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * pVal ) = 0;
      virtual HRESULT __stdcall SetProperty (
        /*[in]*/ enum UCC_CONFERENCE_PARTICIPANT_ENDPOINT_PROPERTY enPropertyId,
        /*[in]*/ VARIANT vPropertyValue,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("bfc14f2e-90fe-4a5e-894d-62a16435b4d8"))
IUccCredential : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_UserName (
        /*[out,retval]*/ BSTR * pbstrUserName ) = 0;
      virtual HRESULT __stdcall get_Password (
        /*[out,retval]*/ BSTR * pbstrPassword ) = 0;
      virtual HRESULT __stdcall get_Domain (
        /*[out,retval]*/ BSTR * pVal ) = 0;
};

struct __declspec(uuid("ab9f6a4a-a930-4b5e-bb77-ea6ec31346b3"))
IUccCredentialCache : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_DefaultCredential (
        /*[out,retval]*/ struct IUccCredential * * ppCredential ) = 0;
      virtual HRESULT __stdcall CreateCredential (
        /*[in]*/ BSTR bstrUserName,
        /*[in]*/ BSTR bstrPassword,
        /*[in]*/ BSTR bstrDomain,
        /*[out,retval]*/ struct IUccCredential * * ppCredential ) = 0;
      virtual HRESULT __stdcall SetCredential (
        /*[in]*/ BSTR bstrRealm,
        /*[in]*/ struct IUccCredential * pCredential ) = 0;
      virtual HRESULT __stdcall RemoveCredential (
        /*[in]*/ BSTR bstrRealm ) = 0;
      virtual HRESULT __stdcall TryGetCredential (
        /*[in]*/ BSTR bstrRealm,
        /*[out]*/ struct IUccCredential * * ppCredential,
        /*[out,retval]*/ VARIANT_BOOL * pbFound ) = 0;
};

struct __declspec(uuid("ce70dd5e-c89e-4863-8a59-39221773c9e9"))
IUccDelegate : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ struct IUccUri * * ppUccUri ) = 0;
      virtual HRESULT __stdcall put_Uri (
        /*[in]*/ struct IUccUri * ppUccUri ) = 0;
      virtual HRESULT __stdcall get_Privileges (
        /*[out,retval]*/ long * pPrivileges ) = 0;
      virtual HRESULT __stdcall put_Privileges (
        /*[in]*/ long pPrivileges ) = 0;
};

struct __declspec(uuid("d3e9e4fa-e813-431e-aa53-b73cd7799843"))
IUccDelegatee : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ struct IUccUri * * ppUccUri ) = 0;
      virtual HRESULT __stdcall put_Uri (
        /*[in]*/ struct IUccUri * ppUccUri ) = 0;
};

struct __declspec(uuid("fab74d22-5830-4d14-a918-177f77ea0256"))
IUccDelegateEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldPrivileges (
        /*[out,retval]*/ long * pPrivileges ) = 0;
      virtual HRESULT __stdcall get_NewPrivileges (
        /*[out,retval]*/ long * pPrivileges ) = 0;
};

struct __declspec(uuid("7bcb8062-1931-42ca-8cc6-38ea95d1941f"))
IUccDiagnosticInfo : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_DIAGNOSTIC_INFO_TYPE * penType ) = 0;
      virtual HRESULT __stdcall get_Code (
        /*[out,retval]*/ long * plCode ) = 0;
      virtual HRESULT __stdcall get_SubCodes (
        /*[out,retval]*/ struct IUccArray * * ppErrorCodes ) = 0;
      virtual HRESULT __stdcall get_Source (
        /*[out,retval]*/ BSTR * pbstrSource ) = 0;
      virtual HRESULT __stdcall get_Text (
        /*[out,retval]*/ BSTR * pbstrText ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * ppProperties ) = 0;
};

enum UCC_DIAGNOSTIC_REPORTING_METHODS
{
    UCCDRM_EVENT_LOGGING = 1,
    UCCDRM_SERVER_CDR_LOGGING = 2
};

struct __declspec(uuid("ce5f4b42-3d3f-4466-a3d5-b11e588fade8"))
IUccDiagnosticReportingSettings : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall EnableReportingMethods (
        /*[in]*/ enum UCC_DIAGNOSTIC_REPORTING_METHODS enMethods ) = 0;
      virtual HRESULT __stdcall DisableReportingMethods (
        /*[in]*/ enum UCC_DIAGNOSTIC_REPORTING_METHODS enMethods ) = 0;
      virtual HRESULT __stdcall IsReportingMethodEnabled (
        /*[in]*/ enum UCC_DIAGNOSTIC_REPORTING_METHODS enMethods,
        /*[out,retval]*/ VARIANT_BOOL * pfVal ) = 0;
};

struct __declspec(uuid("2dc3c92f-8fe2-4a84-8f76-bb1a14ff0f3a"))
IUccEndpoint : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_ENDPOINT_TYPE * peType ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * ppContext ) = 0;
      virtual HRESULT __stdcall Enable (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall Disable (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall get_SignalingCapabilities (
        /*[out,retval]*/ long * pVal ) = 0;
      virtual HRESULT __stdcall get_MediaCapabilities (
        /*[out,retval]*/ long * pVal ) = 0;
};

enum UCC_ENDPOINT_STATE
{
    UCCES_IDLE = 0,
    UCCES_ENABLING = 1,
    UCCES_ENABLED = 2,
    UCCES_DISABLING = 3,
    UCCES_DISABLED = 4
};

struct __declspec(uuid("943bca5a-e22d-4c5d-88d0-d0021ee510d7"))
IUccEndpoint2 : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ProxyEndpoints (
        /*[out,retval]*/ struct IUccCollection * * ppCollection ) = 0;
      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ enum UCC_ENDPOINT_STATE * penState ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * pVal ) = 0;
};

struct __declspec(uuid("6cfb1827-9feb-462a-97c6-b858702c4fe4"))
IUccOperationProgressEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OriginalOperationContext (
        /*[out,retval]*/ struct IUccOperationContext * * ppOperationContext ) = 0;
      virtual HRESULT __stdcall get_IsComplete (
        /*[out,retval]*/ VARIANT_BOOL * pfIsComplete ) = 0;
      virtual HRESULT __stdcall get_StatusCode (
        /*[out,retval]*/ HRESULT * pVal ) = 0;
      virtual HRESULT __stdcall get_StatusText (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * pProperties ) = 0;
};

struct __declspec(uuid("91154bc1-8079-4bf7-8914-fc5b62aac3d6"))
IUccFindMediaConnectivityServersEvent : IUccOperationProgressEvent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_MediaConnectivityServerConfigurations (
        /*[out,retval]*/ struct IUccCollection * * ppCollection ) = 0;
};

struct __declspec(uuid("162809da-825d-4e75-9877-a8909897e567"))
IUccFindServerEvent : IUccOperationProgressEvent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_SignalingServers (
        /*[out,retval]*/ struct IUccCollection * * ppCollection ) = 0;
};

struct __declspec(uuid("36e59b05-2fc3-4782-b3c9-7d5dddbe25b4"))
IUccHistoryInfoEntry : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall get_StatusCode (
        /*[out,retval]*/ HRESULT * pVal ) = 0;
      virtual HRESULT __stdcall get_StatusText (
        /*[out,retval]*/ BSTR * pbstrStatusText ) = 0;
      virtual HRESULT __stdcall get_Index (
        /*[out,retval]*/ BSTR * pbstrIndex ) = 0;
      virtual HRESULT __stdcall get_RedirectionReason (
        /*[out,retval]*/ enum UCC_RETARGET_REASON * penReason ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * ppProperties ) = 0;
};

struct __declspec(uuid("2f96ff5d-e7d3-4aef-a979-4a02d69b9dd0"))
IUccInstantMessagingSession : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall SendMessage (
        /*[in]*/ BSTR bstrContentType,
        /*[in]*/ BSTR bstrMessage,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall StartComposing (
        /*[in]*/ struct IUccContext * pComposingContext ) = 0;
      virtual HRESULT __stdcall StopComposing (
        /*[in]*/ struct IUccContext * pComposingContext ) = 0;
};

struct __declspec(uuid("e5e2d6bc-910d-4348-9d6b-23ec74908e68"))
IUccInstantMessagingSessionParticipant : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_IsComposing (
        /*[out,retval]*/ VARIANT_BOOL * pbComposing ) = 0;
      virtual HRESULT __stdcall get_ComposingContentType (
        /*[out,retval]*/ BSTR * pbstrContentType ) = 0;
};

struct __declspec(uuid("780eca62-e839-4bb0-84c1-005f5f4bc855"))
IUccMediaConnectivityServer : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_MEDIA_CONNECTIVITY_SERVER_TYPE * pVal ) = 0;
      virtual HRESULT __stdcall get_Address (
        /*[out,retval]*/ BSTR * pbstrAddress ) = 0;
      virtual HRESULT __stdcall get_Port (
        /*[out,retval]*/ short * pnPort ) = 0;
      virtual HRESULT __stdcall get_TransportModes (
        /*[out,retval]*/ enum UCC_TRANSPORT_MODE * penTransportMode ) = 0;
      virtual HRESULT __stdcall get_Priority (
        /*[out,retval]*/ long * plPriority ) = 0;
      virtual HRESULT __stdcall get_Realm (
        /*[out,retval]*/ BSTR * pbstrRealm ) = 0;
      virtual HRESULT __stdcall put_Realm (
        /*[in]*/ BSTR pbstrRealm ) = 0;
};

struct __declspec(uuid("6a84605f-af6a-402b-9b0d-5ff881f81879"))
IUccMediaConnectivityServerConfiguration : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_MediaConnectivityServer (
        /*[out,retval]*/ struct IUccMediaConnectivityServer * * pServer ) = 0;
      virtual HRESULT __stdcall get_Credential (
        /*[out,retval]*/ struct IUccCredential * * ppCredential ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * pContext ) = 0;
};

struct __declspec(uuid("216d2747-60a0-43be-abb3-1993d263aa68"))
IUccMediaDevice : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_MediaType (
        /*[out,retval]*/ enum UCC_MEDIA_TYPES * penMediaType ) = 0;
      virtual HRESULT __stdcall get_MediaDirection (
        /*[out,retval]*/ enum UCC_MEDIA_DIRECTIONS * penMediaDirection ) = 0;
      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * pbstrId ) = 0;
      virtual HRESULT __stdcall get_Description (
        /*[out,retval]*/ BSTR * pbstrDescription ) = 0;
      virtual HRESULT __stdcall get_IsNull (
        /*[out,retval]*/ VARIANT_BOOL * pbIsNull ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * ppProperties ) = 0;
};

struct __declspec(uuid("8d538142-1fc9-4cd4-b53b-5601360a908e"))
IUccMediaChannelDevices : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetDevice (
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enDirection,
        /*[out,retval]*/ struct IUccMediaDevice * * ppDevice ) = 0;
      virtual HRESULT __stdcall SetDevice (
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enDirection,
        /*[in]*/ struct IUccMediaDevice * pDevice ) = 0;
};

enum UCC_AUDIO_VIDEO_DEVICE_CATEGORY
{
    UCCAVDC_NONE = 0,
    UCCAVDC_AUDIO_VIDEO_CALL_DEVICE = 1,
    UCCAVDC_AUDIO_VIDEO_CALL_NONPRIVATE_DEVICE = 2,
    UCCAVDC_NOTIFICATION_DEVICE = 3
};

struct __declspec(uuid("24b14557-850e-494f-9ca7-799559cd3eeb"))
IUccMediaDeviceManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetDevices (
        /*[in]*/ enum UCC_MEDIA_TYPES enMediaType,
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enMediaDirection,
        /*[out,retval]*/ struct IUccCollection * * ppMediaDevices ) = 0;
      virtual HRESULT __stdcall GetRecommendedDevice (
        /*[in]*/ enum UCC_AUDIO_VIDEO_DEVICE_CATEGORY enDeviceCategory,
        /*[in]*/ enum UCC_MEDIA_TYPES enMediaType,
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enMediaDirection,
        /*[out,retval]*/ struct IUccMediaDevice * * ppMediaDevice ) = 0;
      virtual HRESULT __stdcall GetSelectedDevice (
        /*[in]*/ enum UCC_AUDIO_VIDEO_DEVICE_CATEGORY enDeviceCategory,
        /*[in]*/ enum UCC_MEDIA_TYPES enMediaType,
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enMediaDirection,
        /*[out,retval]*/ struct IUccMediaDevice * * ppMediaDevice ) = 0;
      virtual HRESULT __stdcall SetSelectedDevice (
        /*[in]*/ enum UCC_AUDIO_VIDEO_DEVICE_CATEGORY enDeviceCategory,
        /*[in]*/ enum UCC_MEDIA_TYPES enMediaType,
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enMediaDirection,
        /*[in]*/ struct IUccMediaDevice * pMediaDevice ) = 0;
      virtual HRESULT __stdcall GetNullDevice (
        /*[in]*/ enum UCC_MEDIA_TYPES enMediaType,
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enMediaDirection,
        /*[out]*/ struct IUccMediaDevice * * ppMediaDevice ) = 0;
};

struct __declspec(uuid("56a60090-39c2-48fb-ad07-43c847c55c4c"))
IUccMediaDeviceSettings : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetDeviceSettings (
        /*[out,retval]*/ struct IUccPropertyCollection * * ppProperties ) = 0;
      virtual HRESULT __stdcall SetDeviceSettings (
        /*[in]*/ struct IUccPropertyCollection * pProperties ) = 0;
      virtual HRESULT __stdcall ResetDeviceSettings ( ) = 0;
      virtual HRESULT __stdcall InvokeTuningWizard (
        /*[in]*/ LONG_PTR_1 hwndParent,
        /*[in]*/ long lPagesToShow ) = 0;
};

struct __declspec(uuid("4d587df7-1adb-48cb-b6b4-914a85c26f6a"))
IUccMediaEndpointSettings : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall FindMediaConnectivityServers (
        /*[in]*/ BSTR bstrMediaInfoServer,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall AddMediaConnectivityServerWithCredential (
        /*[in]*/ struct IUccMediaConnectivityServerConfiguration * pServerConfig ) = 0;
      virtual HRESULT __stdcall AddMediaConnectivityServer (
        /*[in]*/ enum UCC_MEDIA_CONNECTIVITY_SERVER_TYPE enType,
        /*[in]*/ BSTR bstrAddress,
        /*[in]*/ short nPort,
        /*[in]*/ enum UCC_TRANSPORT_MODE enTransportMode,
        /*[in]*/ BSTR bstrRealm,
        /*[in]*/ long lPriority,
        /*[out,retval]*/ struct IUccMediaConnectivityServer * * ppServer ) = 0;
      virtual HRESULT __stdcall RemoveMediaConnectivityServer (
        /*[in]*/ struct IUccMediaConnectivityServer * Server ) = 0;
      virtual HRESULT __stdcall get_MediaConnectivityServers (
        /*[out,retval]*/ struct IUccCollection * * pVal ) = 0;
      virtual HRESULT __stdcall get_CredentialCache (
        /*[out,retval]*/ struct IUccCredentialCache * * ppCredentialCache ) = 0;
      virtual HRESULT __stdcall SetMediaPortRange (
        /*[in]*/ short nPortMin,
        /*[in]*/ short nPortMax ) = 0;
      virtual HRESULT __stdcall GetMediaPortRange (
        /*[out]*/ short * pnPortMin,
        /*[out]*/ short * pnPortMax ) = 0;
};

struct __declspec(uuid("2df08b07-6f9c-4890-a2b0-dda68f430bfb"))
IUccMediaPlayerDevice : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall put_LoopingEnabled (
        /*[in]*/ VARIANT_BOOL pvbLoop ) = 0;
      virtual HRESULT __stdcall get_LoopingEnabled (
        /*[out,retval]*/ VARIANT_BOOL * pvbLoop ) = 0;
      virtual HRESULT __stdcall get_Duration (
        /*[out,retval]*/ unsigned long * plDuration ) = 0;
      virtual HRESULT __stdcall get_FileSource (
        /*[out,retval]*/ BSTR * pbstrFileName ) = 0;
      virtual HRESULT __stdcall SetStartPosition (
        /*[in]*/ unsigned long ulPosition ) = 0;
};

struct __declspec(uuid("8ab54fb5-7656-4348-b5c5-32ca2efee7f4"))
IUccPlatform : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall Initialize (
        /*[in]*/ BSTR bstrApplicationName,
        /*[in]*/ struct IUccContext * pContext ) = 0;
      virtual HRESULT __stdcall CreateEndpoint (
        /*[in]*/ enum UCC_ENDPOINT_TYPE eType,
        /*[in]*/ struct IUccUri * pUri,
        /*[in]*/ BSTR bstrEndpointId,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccEndpoint * * ppEndpoint ) = 0;
      virtual HRESULT __stdcall CreateProxyEndpoint (
        /*[in]*/ enum UCC_ENDPOINT_TYPE eType,
        /*[in]*/ struct IUccEndpoint * pControllingEndpoint,
        /*[in]*/ struct IUccUri * pUri,
        /*[in]*/ BSTR bstrEndpointId,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccEndpoint * * ppEndpoint ) = 0;
      virtual HRESULT __stdcall Shutdown (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("8e17e673-96ec-4aea-aaec-5280ed162410"))
IUccProvisioningPolicyInstance : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * ppProperties ) = 0;
};

struct __declspec(uuid("b22ef3c5-c0dc-473b-b2a4-5ee253e87f5e"))
IUccProxyEndpoint : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_PrincipalEndpoint (
        /*[out,retval]*/ struct IUccEndpoint * * pPrincipleEndpoint ) = 0;
};

struct __declspec(uuid("5e4739c8-240d-4678-9574-9b342f8e6094"))
IUccEndpointRefreshEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * pProperties ) = 0;
};

struct __declspec(uuid("20f08a08-e163-45f8-b43a-2dab579c51ae"))
IUccSignalingServer : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ServerAddress (
        /*[out,retval]*/ BSTR * pbstrServer ) = 0;
      virtual HRESULT __stdcall get_TransportMode (
        /*[out,retval]*/ enum UCC_TRANSPORT_MODE * penMode ) = 0;
};

struct __declspec(uuid("3405d52a-1498-4bd5-9248-e423e97bf1fe"))
IUccServerSignalingSettings : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Server (
        /*[out,retval]*/ struct IUccSignalingServer * * ppServer ) = 0;
      virtual HRESULT __stdcall put_Server (
        /*[in]*/ struct IUccSignalingServer * ppServer ) = 0;
      virtual HRESULT __stdcall get_AllowedAuthenticationModes (
        /*[out,retval]*/ long * plModes ) = 0;
      virtual HRESULT __stdcall put_AllowedAuthenticationModes (
        /*[in]*/ long plModes ) = 0;
      virtual HRESULT __stdcall get_AuthenticationMode (
        /*[out,retval]*/ enum UCC_AUTHENTICATION_MODES * penMode ) = 0;
      virtual HRESULT __stdcall get_SupportedAuthenticationModes (
        /*[out,retval]*/ long * plModes ) = 0;
      virtual HRESULT __stdcall get_CredentialCache (
        /*[out,retval]*/ struct IUccCredentialCache * * ppCredentialCache ) = 0;
      virtual HRESULT __stdcall FindServer (
        /*[in]*/ BSTR bstrDomain,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall CreateSignalingServer (
        /*[in]*/ BSTR bstrServerAddr,
        /*[in]*/ enum UCC_TRANSPORT_MODE enTransportMode,
        /*[out,retval]*/ struct IUccSignalingServer * * pServer ) = 0;
};

struct __declspec(uuid("83a209ca-296c-44d9-9331-f5dcf3e3c0ab"))
_IUccSessionDescriptionEvaluator : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall Evaluate (
        /*[in]*/ BSTR bstrContentType,
        /*[in]*/ BSTR bstrSessionDescription,
        /*[out,retval]*/ VARIANT_BOOL * pbIsApplicationSession ) = 0;
};

struct __declspec(uuid("8e0d1060-df8e-45f9-a834-237f1e4418f0"))
IUccTraceSettings : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall EnableTracing ( ) = 0;
      virtual HRESULT __stdcall DisableTracing ( ) = 0;
      virtual HRESULT __stdcall get_TracingEnabled (
        /*[out,retval]*/ VARIANT_BOOL * pfVal ) = 0;
};

struct __declspec(uuid("56a868b4-0ad4-11ce-b03a-0020af0ba770"))
IVideoWindow : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall put_Caption (
        /*[in]*/ BSTR strCaption ) = 0;
      virtual HRESULT __stdcall get_Caption (
        /*[out,retval]*/ BSTR * strCaption ) = 0;
      virtual HRESULT __stdcall put_WindowStyle (
        /*[in]*/ long WindowStyle ) = 0;
      virtual HRESULT __stdcall get_WindowStyle (
        /*[out,retval]*/ long * WindowStyle ) = 0;
      virtual HRESULT __stdcall put_WindowStyleEx (
        /*[in]*/ long WindowStyleEx ) = 0;
      virtual HRESULT __stdcall get_WindowStyleEx (
        /*[out,retval]*/ long * WindowStyleEx ) = 0;
      virtual HRESULT __stdcall put_AutoShow (
        /*[in]*/ long AutoShow ) = 0;
      virtual HRESULT __stdcall get_AutoShow (
        /*[out,retval]*/ long * AutoShow ) = 0;
      virtual HRESULT __stdcall put_WindowState (
        /*[in]*/ long WindowState ) = 0;
      virtual HRESULT __stdcall get_WindowState (
        /*[out,retval]*/ long * WindowState ) = 0;
      virtual HRESULT __stdcall put_BackgroundPalette (
        /*[in]*/ long pBackgroundPalette ) = 0;
      virtual HRESULT __stdcall get_BackgroundPalette (
        /*[out,retval]*/ long * pBackgroundPalette ) = 0;
      virtual HRESULT __stdcall put_Visible (
        /*[in]*/ long pVisible ) = 0;
      virtual HRESULT __stdcall get_Visible (
        /*[out,retval]*/ long * pVisible ) = 0;
      virtual HRESULT __stdcall put_Left (
        /*[in]*/ long pLeft ) = 0;
      virtual HRESULT __stdcall get_Left (
        /*[out,retval]*/ long * pLeft ) = 0;
      virtual HRESULT __stdcall put_Width (
        /*[in]*/ long pWidth ) = 0;
      virtual HRESULT __stdcall get_Width (
        /*[out,retval]*/ long * pWidth ) = 0;
      virtual HRESULT __stdcall put_Top (
        /*[in]*/ long pTop ) = 0;
      virtual HRESULT __stdcall get_Top (
        /*[out,retval]*/ long * pTop ) = 0;
      virtual HRESULT __stdcall put_Height (
        /*[in]*/ long pHeight ) = 0;
      virtual HRESULT __stdcall get_Height (
        /*[out,retval]*/ long * pHeight ) = 0;
      virtual HRESULT __stdcall put_Owner (
        /*[in]*/ LONG_PTR_1 Owner ) = 0;
      virtual HRESULT __stdcall get_Owner (
        /*[out,retval]*/ LONG_PTR_1 * Owner ) = 0;
      virtual HRESULT __stdcall put_MessageDrain (
        /*[in]*/ LONG_PTR_1 Drain ) = 0;
      virtual HRESULT __stdcall get_MessageDrain (
        /*[out,retval]*/ LONG_PTR_1 * Drain ) = 0;
      virtual HRESULT __stdcall get_BorderColor (
        /*[out,retval]*/ long * Color ) = 0;
      virtual HRESULT __stdcall put_BorderColor (
        /*[in]*/ long Color ) = 0;
      virtual HRESULT __stdcall get_FullScreenMode (
        /*[out,retval]*/ long * FullScreenMode ) = 0;
      virtual HRESULT __stdcall put_FullScreenMode (
        /*[in]*/ long FullScreenMode ) = 0;
      virtual HRESULT __stdcall SetWindowForeground (
        /*[in]*/ long Focus ) = 0;
      virtual HRESULT __stdcall NotifyOwnerMessage (
        /*[in]*/ LONG_PTR_1 hwnd,
        /*[in]*/ long uMsg,
        /*[in]*/ LONG_PTR_1 wParam,
        /*[in]*/ LONG_PTR_1 lParam ) = 0;
      virtual HRESULT __stdcall SetWindowPosition (
        /*[in]*/ long Left,
        /*[in]*/ long Top,
        /*[in]*/ long Width,
        /*[in]*/ long Height ) = 0;
      virtual HRESULT __stdcall GetWindowPosition (
        /*[out]*/ long * pLeft,
        /*[out]*/ long * pTop,
        /*[out]*/ long * pWidth,
        /*[out]*/ long * pHeight ) = 0;
      virtual HRESULT __stdcall GetMinIdealImageSize (
        /*[out]*/ long * pWidth,
        /*[out]*/ long * pHeight ) = 0;
      virtual HRESULT __stdcall GetMaxIdealImageSize (
        /*[out]*/ long * pWidth,
        /*[out]*/ long * pHeight ) = 0;
      virtual HRESULT __stdcall GetRestorePosition (
        /*[out]*/ long * pLeft,
        /*[out]*/ long * pTop,
        /*[out]*/ long * pWidth,
        /*[out]*/ long * pHeight ) = 0;
      virtual HRESULT __stdcall HideCursor (
        /*[in]*/ long HideCursor ) = 0;
      virtual HRESULT __stdcall IsCursorHidden (
        /*[out]*/ long * CursorHidden ) = 0;
};

struct __declspec(uuid("bb782727-4c0a-42ba-a8f8-cb6bd7140bab"))
IUccVideoMediaChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_VideoType (
        /*[out,retval]*/ enum UCC_VIDEO_TYPE * penVideoType ) = 0;
      virtual HRESULT __stdcall get_VideoWindow (
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS enMediaDirection,
        /*[out,retval]*/ struct IVideoWindow * * ppVideoWindow ) = 0;
};

struct __declspec(uuid("230175b3-791c-4904-bce9-8048cda85f87"))
IUccVideoMediaDevice : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_VideoType (
        /*[out,retval]*/ enum UCC_VIDEO_TYPE * penVideoType ) = 0;
      virtual HRESULT __stdcall StartPreview ( ) = 0;
      virtual HRESULT __stdcall get_PreviewWindow (
        /*[out,retval]*/ struct IVideoWindow * * ppPreviewWindow ) = 0;
      virtual HRESULT __stdcall StopPreview ( ) = 0;
};

struct __declspec(uuid("90a4b7cd-1dac-498f-a12b-fce949550694"))
IUccAudioVideoMediaChannelStreamsEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum UCC_MEDIA_STREAM_STATE * pVal ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum UCC_MEDIA_STREAM_STATE * pVal ) = 0;
      virtual HRESULT __stdcall get_MediaDirection (
        /*[out,retval]*/ enum UCC_MEDIA_DIRECTIONS * peDirections ) = 0;
      virtual HRESULT __stdcall get_Reason (
        /*[out,retval]*/ enum UCC_MEDIA_STREAM_STATE_CHANGE_REASON * penReason ) = 0;
};

struct __declspec(uuid("63bbfd31-c323-4a38-be5a-e9af55d5e6a7"))
IUccAudioVideoMediaChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetMediaStreamState (
        /*[in]*/ enum UCC_MEDIA_DIRECTIONS eMediaDirections,
        /*[out,retval]*/ enum UCC_MEDIA_STREAM_STATE * peMediaStreamState ) = 0;
};

struct __declspec(uuid("5072148c-de7a-4826-965c-812ab676e0a4"))
IUccUserSearchQuery : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall put_SearchTerm (
        /*[in]*/ BSTR bstrName,
        /*[in]*/ BSTR pbstrValue ) = 0;
      virtual HRESULT __stdcall get_SearchTerm (
        /*[in]*/ BSTR bstrName,
        /*[out,retval]*/ BSTR * pbstrValue ) = 0;
      virtual HRESULT __stdcall get_SearchTerms (
        /*[out,retval]*/ BSTR * pbstrNames ) = 0;
      virtual HRESULT __stdcall put_SearchPreference (
        /*[in]*/ enum UCC_USER_SEARCH_PREFERENCE enPreference,
        /*[in]*/ unsigned long pdwValue ) = 0;
      virtual HRESULT __stdcall get_SearchPreference (
        /*[in]*/ enum UCC_USER_SEARCH_PREFERENCE enPreference,
        /*[out,retval]*/ unsigned long * pdwValue ) = 0;
      virtual HRESULT __stdcall put_SearchDomain (
        /*[in]*/ BSTR pbstrDomain ) = 0;
      virtual HRESULT __stdcall get_SearchDomain (
        /*[out,retval]*/ BSTR * pbstrDomain ) = 0;
      virtual HRESULT __stdcall Execute (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("345420c6-72b7-4640-a474-9eb877fcf60a"))
IUccUserSearchManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateQuery (
        /*[out,retval]*/ struct IUccUserSearchQuery * * ppQuery ) = 0;
};

struct __declspec(uuid("54562fbc-5a84-4461-8bc9-590737e5de13"))
IUccUserSearchQueryEvent : IUccOperationProgressEvent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Results (
        /*[out,retval]*/ struct IUccCollection * * ppCollection ) = 0;
      virtual HRESULT __stdcall get_MoreAvailable (
        /*[out,retval]*/ VARIANT_BOOL * pfMoreAvailable ) = 0;
};

struct __declspec(uuid("ac628a4c-7595-4129-83f8-b37c2f8af813"))
IUccUserSearchResultRecord : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Value (
        /*[in]*/ enum UCC_USER_SEARCH_COLUMN enColumn,
        /*[out,retval]*/ BSTR * pbstrValue ) = 0;
};

enum UCC_CONFERENCE_MEDIA_STATE
{
    UCCCMS_DISCONNECTED = 0,
    UCCCMS_PENDING = 1,
    UCCCMS_OFFERING = 2,
    UCCCMS_TRYING = 3,
    UCCCMS_PROCEEDING = 4,
    UCCCMS_RINGING = 5,
    UCCCMS_JOINING = 6,
    UCCCMS_CONNECTED = 7,
    UCCCMS_HOLD = 8,
    UCCCMS_FORWARDING = 9,
    UCCCMS_TRANSFERING = 10
};

struct __declspec(uuid("8f77e9f1-ebe2-496d-96fd-a7ddd1adb329"))
IUccConferenceMediaChannelStateChangedEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum UCC_CONFERENCE_MEDIA_STATE * penState ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum UCC_CONFERENCE_MEDIA_STATE * penState ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * ppVal ) = 0;
};

struct __declspec(uuid("db49e9b0-4c96-46b6-8227-159481610c6b"))
_IUccApplicationSessionParticipantEvents : IDispatch
{};

struct __declspec(uuid("6d362a12-a539-4581-a2ec-2647caff6a72"))
IUccOutgoingInvitationEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ContentType (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall put_ContentType (
        /*[in]*/ BSTR pVal ) = 0;
      virtual HRESULT __stdcall get_SessionDescription (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall put_SessionDescription (
        /*[in]*/ BSTR pVal ) = 0;
      virtual HRESULT __stdcall Send ( ) = 0;
};

struct __declspec(uuid("ceb0d9cb-78b9-4d19-b8d9-480a1bbf9c38"))
IUccIncomingRequestEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall Accept ( ) = 0;
      virtual HRESULT __stdcall Reject (
        /*[in]*/ enum UCC_REJECT_OR_TERMINATE_REASON enReason ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * pVal ) = 0;
};

struct __declspec(uuid("4896d466-934f-43a2-bf32-e6ccc6721524"))
IUccIncomingInvitationEvent : IUccIncomingRequestEvent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_RemoteContentType (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall get_RemoteSessionDescription (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall get_LocalContentType (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall put_LocalContentType (
        /*[in]*/ BSTR pVal ) = 0;
      virtual HRESULT __stdcall get_LocalSessionDescription (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall put_LocalSessionDescription (
        /*[in]*/ BSTR pVal ) = 0;
};

struct __declspec(uuid("06eefc22-08be-4ea4-879e-115939e59e96"))
IUccInvitationAcceptedEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_RemoteContentType (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall get_RemoteSessionDescription (
        /*[out,retval]*/ BSTR * pVal ) = 0;
};

struct __declspec(uuid("b6dc5e1a-18ca-47bb-bce7-156142ac5453"))
_IUccAudioVideoSessionEvents : IDispatch
{};

struct __declspec(uuid("80bdd0ab-207b-4180-8511-c9967b0bf698"))
_IUccAudioVideoSessionParticipantEvents : IDispatch
{};

struct __declspec(uuid("46ace669-1ede-478d-98cd-f1d1ddda116f"))
IUccIncomingMediaRequestEvent : IUccIncomingRequestEvent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_MediaChannels (
        /*[out,retval]*/ struct IUccCollection * * ppChannels ) = 0;
      virtual HRESULT __stdcall RemoveChannel (
        /*[in]*/ struct IUccMediaChannel * pChannel ) = 0;
};

struct __declspec(uuid("da1582aa-99e3-424b-87df-d3d15a278e75"))
_IUccAudioMediaChannelEvents : IDispatch
{};

struct __declspec(uuid("d82b8d0c-1e1d-4f13-8ae1-dd54f70ffb09"))
IUccAudioSignalLevelChangedEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Level (
        /*[out,retval]*/ long * pVal ) = 0;
      virtual HRESULT __stdcall get_MinLevel (
        /*[out,retval]*/ long * pVal ) = 0;
      virtual HRESULT __stdcall get_MaxLevel (
        /*[out,retval]*/ long * pVal ) = 0;
      virtual HRESULT __stdcall get_Source (
        /*[out,retval]*/ enum UCC_MEDIA_DIRECTIONS * penDirection ) = 0;
};

struct __declspec(uuid("ea82efbc-1da0-45cb-b4c2-d033b5d19559"))
_IUccMediaChannelDevicesEvents : IDispatch
{};

enum UCC_MEDIA_DEVICE_CHANGE_REASON
{
    UCCMDCR_USER_INITIATED_CHANGE = 1,
    UCCMDCR_SYSTEM_INITIATED_CHANGE = 2,
    UCCMDCR_DEVICE_FAILURE = 3
};

struct __declspec(uuid("d8ba144f-efe7-4eb7-857c-53771fc7984e"))
IUccMediaDeviceChangeEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_MediaDevice (
        /*[out,retval]*/ struct IUccMediaDevice * * ppVal ) = 0;
      virtual HRESULT __stdcall get_DeviceCategory (
        /*[out,retval]*/ enum UCC_AUDIO_VIDEO_DEVICE_CATEGORY * penDeviceCategory ) = 0;
      virtual HRESULT __stdcall get_Reason (
        /*[out,retval]*/ enum UCC_MEDIA_DEVICE_CHANGE_REASON * penReason ) = 0;
};

struct __declspec(uuid("d7db1218-a366-4482-af15-41d5bc859d94"))
_IUccMediaChannelEvents : IDispatch
{};

struct __declspec(uuid("14689259-2a48-4eb3-9094-ec65019b4076"))
IUccMediaChannelEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldMediaDirection (
        /*[out,retval]*/ long * pVal ) = 0;
      virtual HRESULT __stdcall get_NewMediaDirection (
        /*[out,retval]*/ long * pVal ) = 0;
};

struct __declspec(uuid("c7eddd72-0fd5-422e-a206-cf19b09d0e40"))
_IUccMediaEndpointEvents : IDispatch
{};

struct __declspec(uuid("bbfcb8d3-7be1-4525-916b-548a6a617ee4"))
_IUccAudioVideoMediaChannelEvents : IDispatch
{};

struct __declspec(uuid("f891ace7-e9fc-4768-8a50-61489020bc8d"))
_IUccConferenceManagerSessionEvents : IDispatch
{};

struct __declspec(uuid("7dbbf4e6-a70d-4889-86a0-2d95d3ef047b"))
_IUccConferenceMediaChannelEvents : IDispatch
{};

struct __declspec(uuid("826ef7c1-84b3-486a-9790-14430f642d4f"))
IUccPropertyUpdateEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_PropertyId (
        /*[out,retval]*/ long * pVal ) = 0;
      virtual HRESULT __stdcall get_OldPropertyValue (
        /*[out,retval]*/ VARIANT * pVal ) = 0;
      virtual HRESULT __stdcall get_NewPropertyValue (
        /*[out,retval]*/ VARIANT * pVal ) = 0;
};

struct __declspec(uuid("386ab1af-dd8b-41e9-ad87-8c09583e805e"))
_IUccConferenceMediaChannelEvents2 : IDispatch
{};

struct __declspec(uuid("07b3bf54-c27f-4413-b6bf-841cef84e5ba"))
IUccNamedPropertyUpdateEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_PropertyName (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall get_OldPropertyValue (
        /*[out,retval]*/ VARIANT * pVal ) = 0;
      virtual HRESULT __stdcall get_NewPropertyValue (
        /*[out,retval]*/ VARIANT * pVal ) = 0;
};

struct __declspec(uuid("8ec5f518-6c4d-4c79-afa6-47b6d1039bd0"))
_IUccConferenceSessionEvents : IDispatch
{};

struct __declspec(uuid("390a56e6-5868-433c-8cf3-401d104c59f3"))
_IUccConferenceSessionEvents2 : IDispatch
{};

struct __declspec(uuid("25c44946-2e3d-4bf3-a5be-9513e5f969e9"))
IUccConferenceSession2 : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetJoinInformation (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("c819c292-f45f-4854-ae66-c9e3cfe45195"))
_IUccConferenceSessionParticipantEvents : IDispatch
{};

struct __declspec(uuid("c57a388e-c215-4bd4-8b2d-da5f341ef96b"))
_IUccConferenceSessionParticipantEndpointEvents : IDispatch
{};

struct __declspec(uuid("a68a5d90-944e-40c0-9146-883f9c8556bd"))
IUccConferenceSessionParticipantEndpointStateChangedEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum UCC_SESSION_ENTITY_STATE * pVal ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum UCC_SESSION_ENTITY_STATE * pVal ) = 0;
};

struct __declspec(uuid("75025986-cdaa-42e7-9e06-642a8e9e4253"))
_IUccConferenceMediaChannelCollectionEvents : IDispatch
{};

struct __declspec(uuid("56a2af6c-e402-4fc9-9461-d8a92abe1918"))
IUccMediaChannelCollectionEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_MediaChannel (
        /*[out,retval]*/ struct IUccMediaChannel * * pChannel ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * pVal ) = 0;
};

struct __declspec(uuid("53548bb8-662d-4117-9f82-26b85a30742a"))
_IUccConferenceEntityViewCollectionEvents : IDispatch
{};

struct __declspec(uuid("be9e679a-d5af-485e-b1f6-8c74378a5359"))
IUccConferenceEntityViewCollectionEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_EntityView (
        /*[out,retval]*/ struct IUccConferenceEntityView * * ppVal ) = 0;
};

struct __declspec(uuid("67e8922c-51d8-409a-9219-30646c50baea"))
_IUccConferenceEntityViewEvents : IDispatch
{};

struct __declspec(uuid("060bd8f0-5a2a-4d56-a065-2173894aefaf"))
_IUccEndpointEvents : IDispatch
{};

struct __declspec(uuid("1a8a9e13-99e3-4edf-a2f6-9ddbb041023c"))
_IUccEndpointEvents2 : IDispatch
{};

struct __declspec(uuid("fdcb6c19-66cb-492a-9f4e-996d92a50407"))
_IUccInstantMessagingSessionEvents : IDispatch
{};

struct __declspec(uuid("70e0fff6-dd6d-4420-8e19-2de91da99195"))
IUccSessionOperationEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OriginalOperationContext (
        /*[out,retval]*/ struct IUccOperationContext * * ppOperationContext ) = 0;
      virtual HRESULT __stdcall get_ParticipantResults (
        /*[out,retval]*/ struct IUccCollection * * ppIUccSessionParticipantOperationEventCollection ) = 0;
};

struct __declspec(uuid("98a51831-7b5b-49ba-84cf-2a4cde61fa35"))
_IUccInstantMessagingSessionParticipantEvents : IDispatch
{};

struct __declspec(uuid("a1faa7b6-0e6e-4d7e-9e4d-62c1629292b8"))
_IUccPlatformEvents : IDispatch
{};

struct __declspec(uuid("e7a261a2-f96b-4ce9-a0cf-660a4b64fa70"))
_IUccProxyEndpointCollectionEvents : IDispatch
{};

struct __declspec(uuid("03988bae-597c-402e-8b7d-2cd41aa25ecb"))
IUccProxyEndpointCollectionEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Endpoint (
        /*[out,retval]*/ struct IUccEndpoint * * pVal ) = 0;
      virtual HRESULT __stdcall get_Reason (
        /*[out,retval]*/ enum UCC_PROXY_ENDPOINT_COLLECTION_REASON * eReason ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * pVal ) = 0;
};

struct __declspec(uuid("acee6559-9e58-4c91-bc24-eea3f8fce5d6"))
_IUccServerSignalingSettingsEvents : IDispatch
{};

struct __declspec(uuid("ab81ffb3-80a5-4fe9-ac37-b06d5d631be1"))
_IUccSessionEvents : IDispatch
{};

struct __declspec(uuid("3130f1a7-0569-4b49-a8c4-6b902194e55e"))
_IUccSessionParticipantCollectionEvents : IDispatch
{};

struct __declspec(uuid("90bbe2b8-fb81-4e70-ae47-92b75431d4c1"))
_IUccSessionManagerEvents : IDispatch
{};

struct __declspec(uuid("9b48f770-4eac-4967-bcac-283b73557860"))
_IUccSessionParticipantEvents : IDispatch
{};

struct __declspec(uuid("d0e8a9a5-2f04-42f7-96ea-89ab2b547da8"))
IUccSessionParticipantStateChangedEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum UCC_SESSION_ENTITY_STATE * pVal ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum UCC_SESSION_ENTITY_STATE * pVal ) = 0;
      virtual HRESULT __stdcall get_Reason (
        /*[out,retval]*/ enum UCC_PARTICIPANT_STATE_CHANGE_REASON * penReason ) = 0;
};

struct __declspec(uuid("8db557ba-e669-4de9-9922-8c6c53ad1ea3"))
_IUccSessionParticipantEndpointCollectionEvents : IDispatch
{};

struct __declspec(uuid("00fb5f00-b2d5-4f3d-9e4b-feb4180d66fa"))
_IUccMediaChannelCollectionEvents : IDispatch
{};

struct __declspec(uuid("79712662-75f4-44a0-8519-1e6013de4109"))
_IUccSessionCallControlEvents : IDispatch
{};

struct __declspec(uuid("c31836bd-ad83-4a63-b53a-ac6161cea97f"))
IUccSessionTransferProgressEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_TransferState (
        /*[out,retval]*/ enum UCC_SESSION_ENTITY_STATE * penTransferState ) = 0;
};

struct __declspec(uuid("7d3446dd-a657-437a-b47f-3080cbb02ca3"))
IUccIncomingSessionTransferEvent : IUccIncomingRequestEvent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ReferredBy (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall get_ReferredTo (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall get_TransferContext (
        /*[out,retval]*/ struct IUccContext * * ppTransferContext ) = 0;
};

struct __declspec(uuid("6a8e8235-9734-4eb0-938c-a6de10deb4ea"))
_IUccSignalingChannelEvents : IDispatch
{};

struct __declspec(uuid("5a9747c4-41ed-4812-8e73-8dffd9f5c97b"))
IUccSignalingHeaderFieldParameter : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall put_Value (
        /*[in]*/ BSTR pVal ) = 0;
};

struct __declspec(uuid("05bf6f94-bb11-4f38-a45a-1e1cbe177dbf"))
IUccSignalingHeaderField : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall put_Value (
        /*[in]*/ BSTR pVal ) = 0;
      virtual HRESULT __stdcall get_Parameters (
        /*[out,retval]*/ struct IUccCollection * * pVal ) = 0;
      virtual HRESULT __stdcall AddParameter (
        /*[in]*/ BSTR bstrName,
        /*[in]*/ BSTR bstrValue,
        /*[out,retval]*/ struct IUccSignalingHeaderFieldParameter * * ppParameter ) = 0;
      virtual HRESULT __stdcall RemoveParameter (
        /*[in]*/ struct IUccSignalingHeaderFieldParameter * pParameter ) = 0;
      virtual HRESULT __stdcall FindParameter (
        /*[in]*/ BSTR bstrName,
        /*[out,retval]*/ struct IUccCollection * * ppParameters ) = 0;
};

struct __declspec(uuid("9d45f185-46b0-48b4-9e9c-080313280711"))
IUccSignalingMessage : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_SIGNALING_MESSAGE_TYPE * pVal ) = 0;
      virtual HRESULT __stdcall get_HeaderFields (
        /*[out,retval]*/ struct IUccCollection * * pVal ) = 0;
      virtual HRESULT __stdcall get_Body (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall put_Body (
        /*[in]*/ BSTR pVal ) = 0;
      virtual HRESULT __stdcall get_ContentType (
        /*[out,retval]*/ BSTR * pContentType ) = 0;
      virtual HRESULT __stdcall put_ContentType (
        /*[in]*/ BSTR pContentType ) = 0;
      virtual HRESULT __stdcall AddHeaderField (
        /*[in]*/ BSTR bstrName,
        /*[in]*/ BSTR bstrValue,
        /*[out,retval]*/ struct IUccSignalingHeaderField * * ppHeader ) = 0;
      virtual HRESULT __stdcall RemoveHeaderField (
        /*[in]*/ struct IUccSignalingHeaderField * pHeader ) = 0;
      virtual HRESULT __stdcall FindHeaderFields (
        /*[in]*/ BSTR bstrName,
        /*[out,retval]*/ struct IUccCollection * * ppHeaders ) = 0;
};

struct __declspec(uuid("9ea6b70c-c59e-4823-9fdc-e7c85e6320a0"))
IUccSignalingRequest : IUccSignalingMessage
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_RequestType (
        /*[out,retval]*/ enum UCC_SIGNALING_REQUEST_TYPE * pVal ) = 0;
      virtual HRESULT __stdcall put_RequestType (
        /*[in]*/ enum UCC_SIGNALING_REQUEST_TYPE pVal ) = 0;
      virtual HRESULT __stdcall get_RequestUri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall put_RequestUri (
        /*[in]*/ struct IUccUri * ppUri ) = 0;
};

struct __declspec(uuid("23053341-87a9-44eb-bb22-800835bba620"))
IUccSignalingChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall SendRequest (
        /*[in]*/ struct IUccSignalingRequest * pRequest,
        /*[in]*/ struct IUccOperationContext * pOpContext ) = 0;
};

struct __declspec(uuid("bc23ff63-d024-46df-b980-f8d7d76cd6a9"))
IUccIncomingSignalingMessageEvent : IUccIncomingRequestEvent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Message (
        /*[out,retval]*/ struct IUccSignalingMessage * * pVal ) = 0;
};

struct __declspec(uuid("94f59d79-583a-4547-a620-ead932a2f2eb"))
_IUccUserSearchQueryEvents : IDispatch
{};

struct __declspec(uuid("1713afd4-3ba6-48bc-853b-e1b7e1425ec8"))
IUccOperationManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CancelOperation (
        /*[in]*/ struct IUccOperationContext * pCancelledOperationContext,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("84b2232b-9ac5-4a70-bee3-9afb52b3ad25"))
IUccCategoryInstanceTransformationManifest : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_CategoryName (
        /*[out,retval]*/ BSTR * pbstrCategoryName ) = 0;
      virtual HRESULT __stdcall get_ContainerId (
        /*[out,retval]*/ long * plContainerId ) = 0;
      virtual HRESULT __stdcall get_InstanceId (
        /*[out,retval]*/ long * plInstanceId ) = 0;
      virtual HRESULT __stdcall get_TransformationManifest (
        /*[out,retval]*/ BSTR * pbstrManifest ) = 0;
      virtual HRESULT __stdcall put_TransformationManifest (
        /*[in]*/ BSTR pbstrManifest ) = 0;
};

struct __declspec(uuid("218ed275-cbe4-43f2-9108-242e5968b7e4"))
IUccCategoryTransformationManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateTransformationManifest (
        /*[in]*/ BSTR bstrCategoryName,
        /*[in]*/ long lContainerID,
        /*[in]*/ long lInstanceId,
        /*[out,retval]*/ struct IUccCategoryInstanceTransformationManifest * * ppManifest ) = 0;
      virtual HRESULT __stdcall AddTransformationManifest (
        /*[in]*/ struct IUccCategoryInstanceTransformationManifest * pTransform ) = 0;
      virtual HRESULT __stdcall RemoveTransformationManifest (
        /*[in]*/ BSTR bstrCategoryName,
        /*[in]*/ long lContainerID,
        /*[in]*/ long lInstanceId ) = 0;
      virtual HRESULT __stdcall get_TransformationManifests (
        /*[out,retval]*/ struct IUccCollection * * ppCollection ) = 0;
};

enum UCC_PRESENCE_STATE_TYPE
{
    UCCPST_UNKNOWN = 0,
    UCCPST_AGGREGATE_STATE = 1,
    UCCPST_AGGREGATE_MACHINE_STATE = 2,
    UCCPST_USER_STATE = 3,
    UCCPST_MACHINE_STATE = 4,
    UCCPST_PHONE_STATE = 5,
    UCCPST_CALENDAR_STATE = 6,
    UCCPST_CONFERENCE_STATE = 7
};

struct __declspec(uuid("6fdbdefe-8a31-49cd-87b8-cc61a49925fe"))
IUccLocaleString : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_String (
        /*[out,retval]*/ BSTR * pbstrString ) = 0;
      virtual HRESULT __stdcall put_String (
        /*[in]*/ BSTR pbstrString ) = 0;
      virtual HRESULT __stdcall get_Lcid (
        /*[out,retval]*/ long * pLcid ) = 0;
      virtual HRESULT __stdcall put_Lcid (
        /*[in]*/ long pLcid ) = 0;
};

struct __declspec(uuid("890cf073-bb33-4dcb-aee8-08c300469b62"))
IUccPresenceActivity : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Token (
        /*[out,retval]*/ BSTR * pbstrToken ) = 0;
      virtual HRESULT __stdcall put_Token (
        /*[in]*/ BSTR pbstrToken ) = 0;
      virtual HRESULT __stdcall GetAvailabilityRange (
        /*[out]*/ long * plMinAvailability,
        /*[out]*/ long * plMaxAvailability ) = 0;
      virtual HRESULT __stdcall SetAvailabilityRange (
        /*[in]*/ long lMinAvailability,
        /*[in]*/ long lMaxAvailability ) = 0;
      virtual HRESULT __stdcall CreateCustomString (
        /*[in]*/ BSTR bstrString,
        /*[out]*/ struct IUccLocaleString * * ppCustomString ) = 0;
      virtual HRESULT __stdcall AddCustomString (
        /*[in]*/ struct IUccLocaleString * pCustomString ) = 0;
      virtual HRESULT __stdcall RemoveCustomString (
        /*[in]*/ struct IUccLocaleString * pCustomString ) = 0;
      virtual HRESULT __stdcall get_CustomStrings (
        /*[out,retval]*/ struct IUccCollection * * ppAddresses ) = 0;
};

struct __declspec(uuid("3afa976f-86ee-412e-986c-eb084a3d0347"))
IUccPresenceStateInstance : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_PRESENCE_STATE_TYPE * penType ) = 0;
      virtual HRESULT __stdcall put_Type (
        /*[in]*/ enum UCC_PRESENCE_STATE_TYPE penType ) = 0;
      virtual HRESULT __stdcall get_Availability (
        /*[out,retval]*/ long * plAvailability ) = 0;
      virtual HRESULT __stdcall put_Availability (
        /*[in]*/ long plAvailability ) = 0;
      virtual HRESULT __stdcall get_EndpointLocation (
        /*[out,retval]*/ BSTR * pbstrCustomEndpointLocation ) = 0;
      virtual HRESULT __stdcall put_EndpointLocation (
        /*[in]*/ BSTR pbstrCustomEndpointLocation ) = 0;
      virtual HRESULT __stdcall get_MeetingSubject (
        /*[out,retval]*/ BSTR * pbstrMeetingSubject ) = 0;
      virtual HRESULT __stdcall put_MeetingSubject (
        /*[in]*/ BSTR pbstrMeetingSubject ) = 0;
      virtual HRESULT __stdcall get_MeetingLocation (
        /*[out,retval]*/ BSTR * pbstrMeetingLocation ) = 0;
      virtual HRESULT __stdcall put_MeetingLocation (
        /*[in]*/ BSTR pbstrMeetingLocation ) = 0;
      virtual HRESULT __stdcall get_IsManual (
        /*[out,retval]*/ VARIANT_BOOL * pfIsManual ) = 0;
      virtual HRESULT __stdcall put_IsManual (
        /*[in]*/ VARIANT_BOOL pfIsManual ) = 0;
      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * pbstrURI ) = 0;
      virtual HRESULT __stdcall put_Uri (
        /*[in]*/ BSTR pbstrURI ) = 0;
      virtual HRESULT __stdcall get_IsMachineStateForLocalEndpoint (
        /*[out,retval]*/ VARIANT_BOOL * pfIsLocalId ) = 0;
      virtual HRESULT __stdcall get_LastActive (
        /*[out,retval]*/ DATE * pdLastActive ) = 0;
      virtual HRESULT __stdcall get_StartTime (
        /*[out,retval]*/ DATE * pdStartTime ) = 0;
      virtual HRESULT __stdcall put_StartTime (
        /*[in]*/ DATE pdStartTime ) = 0;
      virtual HRESULT __stdcall CreateActivity (
        /*[out,retval]*/ struct IUccPresenceActivity * * ppActivity ) = 0;
      virtual HRESULT __stdcall AddActivity (
        /*[in]*/ struct IUccPresenceActivity * pActivity ) = 0;
      virtual HRESULT __stdcall RemoveActivity (
        /*[in]*/ struct IUccPresenceActivity * pActivity ) = 0;
      virtual HRESULT __stdcall get_Activities (
        /*[out,retval]*/ struct IUccCollection * * ppActivities ) = 0;
      virtual HRESULT __stdcall RemoveAllActivities ( ) = 0;
};

struct __declspec(uuid("b6608b42-8838-4ae6-af83-d0f0b46c7f1a"))
IUccPresenceStateInstance2 : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_TimeZoneName (
        /*[out,retval]*/ BSTR * pbstrTimeZoneName ) = 0;
      virtual HRESULT __stdcall put_TimeZoneName (
        /*[in]*/ BSTR pbstrTimeZoneName ) = 0;
      virtual HRESULT __stdcall get_TimeZoneAbbreviation (
        /*[out,retval]*/ BSTR * pbstrTimeZoneAbbreviation ) = 0;
      virtual HRESULT __stdcall put_TimeZoneAbbreviation (
        /*[in]*/ BSTR pbstrTimeZoneAbbreviation ) = 0;
      virtual HRESULT __stdcall get_TimeZoneBias (
        /*[out,retval]*/ long * plTimeZoneBias ) = 0;
      virtual HRESULT __stdcall put_TimeZoneBias (
        /*[in]*/ long plTimeZoneBias ) = 0;
      virtual HRESULT __stdcall get_DeviceType (
        /*[out,retval]*/ BSTR * pbstrDeviceType ) = 0;
      virtual HRESULT __stdcall put_DeviceType (
        /*[in]*/ BSTR pbstrDeviceType ) = 0;
};

enum UCC_PRESENCE_CATEGORY_TYPE
{
    UCCPCT_STATE = 0,
    UCCPCT_NOTE = 1,
    UCCPCT_CALENDARDATA = 2,
    UCCPCT_DEVICE = 3,
    UCCPCT_SERVICE = 4,
    UCCPCT_CONTACT_CARD = 5,
    UCCPCT_SERVER_CONFIGURATION = 6,
    UCCPCT_MEETING_POLICY = 7,
    UCCPCT_UC_POLICY = 8,
    UCCPCT_UC_USERPROPERTIES = 9,
    UCCPCT_UC_ROUTING = 10
};

struct __declspec(uuid("d450256a-adb1-4ea7-9cf1-3bd645b1c6d1"))
IUccSubscriber : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall put_Uri (
        /*[in]*/ struct IUccUri * ppUri ) = 0;
      virtual HRESULT __stdcall get_SourceType (
        /*[out,retval]*/ enum UCC_SOURCE_NETWORK_TYPE * penSourceType ) = 0;
      virtual HRESULT __stdcall get_Acknowledged (
        /*[out,retval]*/ VARIANT_BOOL * pfAcknowledged ) = 0;
      virtual HRESULT __stdcall put_Acknowledged (
        /*[in]*/ VARIANT_BOOL pfAcknowledged ) = 0;
};

struct __declspec(uuid("573039b6-c916-4e9a-a6a2-4b0530248af6"))
IUccSubscriberEx : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_SubscriberContext (
        /*[out,retval]*/ BSTR * pbstrContext ) = 0;
};

struct __declspec(uuid("f5791faa-7225-41a7-8043-040dcea4517a"))
IUccPresenceContactCardIdentity : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_DisplayName (
        /*[out,retval]*/ BSTR * pName ) = 0;
      virtual HRESULT __stdcall get_EmailAddresses (
        /*[out,retval]*/ struct IUccCollection * * ppAddresses ) = 0;
};

struct __declspec(uuid("755b7ded-e5a2-456a-a5d3-43c1b94cfbcc"))
IUccPresenceContactCardAddress : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_CONTACT_CARD_ADDRESS_TYPE * penType ) = 0;
      virtual HRESULT __stdcall get_StreetAddress (
        /*[out,retval]*/ BSTR * pbstrStreetAddress ) = 0;
      virtual HRESULT __stdcall put_StreetAddress (
        /*[in]*/ BSTR pbstrStreetAddress ) = 0;
      virtual HRESULT __stdcall get_City (
        /*[out,retval]*/ BSTR * pbstrCity ) = 0;
      virtual HRESULT __stdcall put_City (
        /*[in]*/ BSTR pbstrCity ) = 0;
      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ BSTR * pbstrState ) = 0;
      virtual HRESULT __stdcall put_State (
        /*[in]*/ BSTR pbstrState ) = 0;
      virtual HRESULT __stdcall get_PostalCode (
        /*[out,retval]*/ BSTR * pbstrPostalCode ) = 0;
      virtual HRESULT __stdcall put_PostalCode (
        /*[in]*/ BSTR pbstrPostalCode ) = 0;
      virtual HRESULT __stdcall get_Country (
        /*[out,retval]*/ BSTR * pbstrCountry ) = 0;
      virtual HRESULT __stdcall put_Country (
        /*[in]*/ BSTR pbstrCountry ) = 0;
};

struct __declspec(uuid("9433ccce-7a69-4a71-9713-eae839eaf83c"))
IUccPresenceContactCardPhone : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_CONTACT_CARD_PHONE_TYPE * penType ) = 0;
      virtual HRESULT __stdcall get_Url (
        /*[out,retval]*/ BSTR * pbstrUrl ) = 0;
      virtual HRESULT __stdcall put_Url (
        /*[in]*/ BSTR pbstrUrl ) = 0;
};

struct __declspec(uuid("06e754c6-4bff-41fb-ba4d-af1a962be236"))
IUccPresenceContactCardUrl : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_CONTACT_CARD_URL_TYPE * pbstrType ) = 0;
      virtual HRESULT __stdcall get_Url (
        /*[out,retval]*/ BSTR * pbstrUrl ) = 0;
      virtual HRESULT __stdcall put_Url (
        /*[in]*/ BSTR pbstrUrl ) = 0;
      virtual HRESULT __stdcall get_LastUpdated (
        /*[out,retval]*/ DATE * pdLastUpdated ) = 0;
      virtual HRESULT __stdcall put_LastUpdated (
        /*[in]*/ DATE pdLastUpdated ) = 0;
};

struct __declspec(uuid("fc3cfe6f-29d0-4efc-b3ec-6c69bbcbcd74"))
IUccPresenceContactCardInstance : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall RemoveAllData ( ) = 0;
      virtual HRESULT __stdcall get_Identity (
        /*[out,retval]*/ struct IUccPresenceContactCardIdentity * * ppIdentity ) = 0;
      virtual HRESULT __stdcall CreateAddress (
        /*[in]*/ enum UCC_CONTACT_CARD_ADDRESS_TYPE enType,
        /*[out,retval]*/ struct IUccPresenceContactCardAddress * * ppAddress ) = 0;
      virtual HRESULT __stdcall AddAddress (
        /*[in]*/ struct IUccPresenceContactCardAddress * pAddress ) = 0;
      virtual HRESULT __stdcall RemoveAddress (
        /*[in]*/ struct IUccPresenceContactCardAddress * pAddress ) = 0;
      virtual HRESULT __stdcall get_Addresses (
        /*[out,retval]*/ struct IUccCollection * * ppAddresses ) = 0;
      virtual HRESULT __stdcall RemoveAllAddresses ( ) = 0;
      virtual HRESULT __stdcall get_Company (
        /*[out,retval]*/ BSTR * pbstrCompany ) = 0;
      virtual HRESULT __stdcall put_Company (
        /*[in]*/ BSTR pbstrCompany ) = 0;
      virtual HRESULT __stdcall get_Department (
        /*[out,retval]*/ BSTR * pbstrDepartment ) = 0;
      virtual HRESULT __stdcall put_Department (
        /*[in]*/ BSTR pbstrDepartment ) = 0;
      virtual HRESULT __stdcall get_Title (
        /*[out,retval]*/ BSTR * pbstrTitle ) = 0;
      virtual HRESULT __stdcall put_Title (
        /*[in]*/ BSTR pbstrTitle ) = 0;
      virtual HRESULT __stdcall get_OfficeLocation (
        /*[out,retval]*/ BSTR * pbstrOfficeLocation ) = 0;
      virtual HRESULT __stdcall put_OfficeLocation (
        /*[in]*/ BSTR pbstrOfficeLocation ) = 0;
      virtual HRESULT __stdcall CreateUrl (
        /*[in]*/ enum UCC_CONTACT_CARD_URL_TYPE enUrlType,
        /*[out,retval]*/ struct IUccPresenceContactCardUrl * * ppUrl ) = 0;
      virtual HRESULT __stdcall AddUrl (
        /*[in]*/ struct IUccPresenceContactCardUrl * pUrl ) = 0;
      virtual HRESULT __stdcall RemoveUrl (
        /*[in]*/ struct IUccPresenceContactCardUrl * pUrl ) = 0;
      virtual HRESULT __stdcall get_Urls (
        /*[out,retval]*/ struct IUccCollection * * ppUrls ) = 0;
      virtual HRESULT __stdcall RemoveAllUrls ( ) = 0;
      virtual HRESULT __stdcall CreatePhone (
        /*[in]*/ enum UCC_CONTACT_CARD_PHONE_TYPE enPhoneType,
        /*[out,retval]*/ struct IUccPresenceContactCardPhone * * ppPhone ) = 0;
      virtual HRESULT __stdcall AddPhone (
        /*[in]*/ struct IUccPresenceContactCardPhone * pPhone ) = 0;
      virtual HRESULT __stdcall RemovePhone (
        /*[in]*/ struct IUccPresenceContactCardPhone * pPhone ) = 0;
      virtual HRESULT __stdcall get_Phones (
        /*[out,retval]*/ struct IUccCollection * * ppPhones ) = 0;
      virtual HRESULT __stdcall RemoveAllPhones ( ) = 0;
      virtual HRESULT __stdcall get_IsAutomaton (
        /*[out,retval]*/ VARIANT_BOOL * pfFlag ) = 0;
      virtual HRESULT __stdcall put_IsAutomaton (
        /*[in]*/ VARIANT_BOOL pfFlag ) = 0;
};

struct __declspec(uuid("9cce5b3c-7dd5-4452-84f7-cadaebc1ce53"))
IUccPresenceContactCardInstance2 : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall put_PresentityType (
        /*[in]*/ BSTR pbstrPresentityType ) = 0;
      virtual HRESULT __stdcall get_PresentityType (
        /*[out,retval]*/ BSTR * pbstrPresentityType ) = 0;
      virtual HRESULT __stdcall put_PresentityDescription (
        /*[in]*/ BSTR pbstrPresentityDesc ) = 0;
      virtual HRESULT __stdcall get_PresentityDescription (
        /*[out,retval]*/ BSTR * pbstrPresentityDesc ) = 0;
      virtual HRESULT __stdcall put_UCEnabled (
        /*[in]*/ VARIANT_BOOL pfFlag ) = 0;
      virtual HRESULT __stdcall get_UCEnabled (
        /*[out,retval]*/ VARIANT_BOOL * pfFlag ) = 0;
};

struct __declspec(uuid("088fea55-8dc9-4808-a0f5-b04f04c513af"))
IUccPresenceFreeBusyInformation : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_FreeBusyDataVersion (
        /*[out,retval]*/ BSTR * pbstrFreeBusyDataVersion ) = 0;
      virtual HRESULT __stdcall put_FreeBusyDataVersion (
        /*[in]*/ BSTR pbstrFreeBusyDataVersion ) = 0;
      virtual HRESULT __stdcall get_FreeBusyData (
        /*[out,retval]*/ BSTR * pbstrFreeBusyData ) = 0;
      virtual HRESULT __stdcall put_FreeBusyData (
        /*[in]*/ BSTR pbstrFreeBusyData ) = 0;
      virtual HRESULT __stdcall get_Granularity (
        /*[out,retval]*/ long * plGranularity ) = 0;
      virtual HRESULT __stdcall put_Granularity (
        /*[in]*/ long plGranularity ) = 0;
      virtual HRESULT __stdcall get_StartTime (
        /*[out,retval]*/ DATE * pdStartTime ) = 0;
      virtual HRESULT __stdcall put_StartTime (
        /*[in]*/ DATE pdStartTime ) = 0;
};

struct __declspec(uuid("f3933dc2-1e04-43c9-a0ff-5d578a352ba8"))
IUccPresenceTimeZoneInformation : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Bias (
        /*[out,retval]*/ long * plBias ) = 0;
      virtual HRESULT __stdcall put_Bias (
        /*[in]*/ long plBias ) = 0;
      virtual HRESULT __stdcall get_Time (
        /*[out,retval]*/ long * plTime ) = 0;
      virtual HRESULT __stdcall put_Time (
        /*[in]*/ long plTime ) = 0;
      virtual HRESULT __stdcall get_DayOrder (
        /*[out,retval]*/ long * plDayOrder ) = 0;
      virtual HRESULT __stdcall put_DayOrder (
        /*[in]*/ long plDayOrder ) = 0;
      virtual HRESULT __stdcall get_Month (
        /*[out,retval]*/ long * plMonth ) = 0;
      virtual HRESULT __stdcall put_Month (
        /*[in]*/ long plMonth ) = 0;
      virtual HRESULT __stdcall get_DayOfWeek (
        /*[out,retval]*/ enum UCC_PRESENCE_WORKING_HOURS_DAYS * plDays ) = 0;
      virtual HRESULT __stdcall put_DayOfWeek (
        /*[in]*/ enum UCC_PRESENCE_WORKING_HOURS_DAYS plDays ) = 0;
};

struct __declspec(uuid("10659f8e-f79f-4d16-b26d-a7079014483b"))
IUccPresenceWorkingHoursBlock : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ApplicableDays (
        /*[out,retval]*/ long * plDays ) = 0;
      virtual HRESULT __stdcall put_ApplicableDays (
        /*[in]*/ long plDays ) = 0;
      virtual HRESULT __stdcall get_StartTime (
        /*[out,retval]*/ long * plStartTime ) = 0;
      virtual HRESULT __stdcall put_StartTime (
        /*[in]*/ long plStartTime ) = 0;
      virtual HRESULT __stdcall get_EndTime (
        /*[out,retval]*/ long * plEndTime ) = 0;
      virtual HRESULT __stdcall put_EndTime (
        /*[in]*/ long plEndTime ) = 0;
};

struct __declspec(uuid("d8bec519-3f8d-4fa3-a9c3-43be4b4e7c0b"))
IUccPresenceWorkingHoursInformation : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_TimeZoneBias (
        /*[out,retval]*/ long * plTimeZoneBias ) = 0;
      virtual HRESULT __stdcall put_TimeZoneBias (
        /*[in]*/ long plTimeZoneBias ) = 0;
      virtual HRESULT __stdcall get_TimeZoneInformation (
        /*[in]*/ enum UCC_PRESENCE_WORKING_HOURS_TIMEZONE_TYPE enType,
        /*[out,retval]*/ struct IUccPresenceTimeZoneInformation * * ppTimeZoneInfo ) = 0;
      virtual HRESULT __stdcall put_TimeZoneInformation (
        /*[in]*/ enum UCC_PRESENCE_WORKING_HOURS_TIMEZONE_TYPE enType,
        /*[in]*/ struct IUccPresenceTimeZoneInformation * ppTimeZoneInfo ) = 0;
      virtual HRESULT __stdcall CreateTimeZoneInformation (
        /*[out,retval]*/ struct IUccPresenceTimeZoneInformation * * ppTimeZoneInfo ) = 0;
      virtual HRESULT __stdcall AddWorkingHoursBlock (
        /*[in]*/ struct IUccPresenceWorkingHoursBlock * pBlock ) = 0;
      virtual HRESULT __stdcall RemoveWorkingHoursBlock (
        /*[in]*/ struct IUccPresenceWorkingHoursBlock * pBlock ) = 0;
      virtual HRESULT __stdcall RemoveAllWorkingHours ( ) = 0;
      virtual HRESULT __stdcall CreateWorkingHoursBlock (
        /*[out,retval]*/ struct IUccPresenceWorkingHoursBlock * * ppBlock ) = 0;
      virtual HRESULT __stdcall get_WorkingHoursBlocks (
        /*[out,retval]*/ struct IUccCollection * * ppBlock ) = 0;
};

struct __declspec(uuid("a035a856-83b7-43bc-b3ac-1aa1e2c14ff1"))
IUccPresenceCapabilityInfo : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pbstrCapabilityName ) = 0;
      virtual HRESULT __stdcall get_Preferred (
        /*[out,retval]*/ VARIANT_BOOL * pfPreferred ) = 0;
      virtual HRESULT __stdcall get_CapabilityProperties (
        /*[out,retval]*/ struct IUccPropertyCollection * * pPropertyCollection ) = 0;
      virtual HRESULT __stdcall put_CapabilityProperties (
        /*[in]*/ struct IUccPropertyCollection * pPropertyCollection ) = 0;
      virtual HRESULT __stdcall CreateCapabilityProperties (
        /*[out,retval]*/ struct IUccPropertyCollection * * ppCapabilityProperties ) = 0;
};

struct __declspec(uuid("1de93dcf-3bd3-411d-822c-6ec2adf9ffb7"))
IUccPresenceCapabilities : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * pbstrURI ) = 0;
      virtual HRESULT __stdcall get_Preferred (
        /*[out,retval]*/ VARIANT_BOOL * pfPreferred ) = 0;
      virtual HRESULT __stdcall put_Preferred (
        /*[in]*/ VARIANT_BOOL pfPreferred ) = 0;
      virtual HRESULT __stdcall get_CapabilityInfoCollection (
        /*[out,retval]*/ struct IUccCollection * * ppCapabilityInfoCollection ) = 0;
      virtual HRESULT __stdcall AddCapabilityInfo (
        /*[in]*/ struct IUccPresenceCapabilityInfo * pCapabilityInfo ) = 0;
      virtual HRESULT __stdcall RemoveCapabilityInfo (
        /*[in]*/ struct IUccPresenceCapabilityInfo * pCapabilityInfo ) = 0;
      virtual HRESULT __stdcall CreateCapabilityInfo (
        /*[in]*/ BSTR bstrCapabilityName,
        /*[out,retval]*/ struct IUccPresenceCapabilityInfo * * ppCapabilityProperties ) = 0;
};

struct __declspec(uuid("0fae5c66-18af-4d5d-9164-b0e41db04fb7"))
IUccPresenceCalendarDataInstance : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_MailboxUri (
        /*[out,retval]*/ BSTR * pbstrMailboxURI ) = 0;
      virtual HRESULT __stdcall put_MailboxUri (
        /*[in]*/ BSTR pbstrMailboxURI ) = 0;
      virtual HRESULT __stdcall get_FreeBusyInformation (
        /*[out,retval]*/ struct IUccPresenceFreeBusyInformation * * ppFreeBusyInfo ) = 0;
      virtual HRESULT __stdcall put_FreeBusyInformation (
        /*[in]*/ struct IUccPresenceFreeBusyInformation * ppFreeBusyInfo ) = 0;
      virtual HRESULT __stdcall CreateFreeBusyInformation (
        /*[out,retval]*/ struct IUccPresenceFreeBusyInformation * * ppFreeBusyInfo ) = 0;
      virtual HRESULT __stdcall get_WorkingHoursInformation (
        /*[out,retval]*/ struct IUccPresenceWorkingHoursInformation * * ppWorkingHoursInfo ) = 0;
      virtual HRESULT __stdcall put_WorkingHoursInformation (
        /*[in]*/ struct IUccPresenceWorkingHoursInformation * ppWorkingHoursInfo ) = 0;
      virtual HRESULT __stdcall CreateWorkingHoursInformation (
        /*[out,retval]*/ struct IUccPresenceWorkingHoursInformation * * ppWorkingHoursInfo ) = 0;
};

struct __declspec(uuid("eb59d99c-8e23-410a-9f49-e3a3585c8116"))
IUccPresenceServicesInstance : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ServiceElements (
        /*[out,retval]*/ struct IUccCollection * * ppElements ) = 0;
};

struct __declspec(uuid("21c89f28-862d-42ce-b12c-090cd3adea82"))
IUccPresenceServiceElement : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * pbstrURI ) = 0;
      virtual HRESULT __stdcall get_PreferredEndpoint (
        /*[out,retval]*/ BSTR * pbstrPreferredEndpoint ) = 0;
      virtual HRESULT __stdcall get_CapabilitiesCollection (
        /*[out,retval]*/ struct IUccCollection * * ppCapabilities ) = 0;
};

struct __declspec(uuid("53c63418-b703-4d25-95d6-c2385c18ed6a"))
IUccPresenceDeviceInstance : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Gruu (
        /*[out,retval]*/ BSTR * pbstrGruu ) = 0;
      virtual HRESULT __stdcall get_CapabilitiesCollection (
        /*[out,retval]*/ struct IUccCollection * * ppCapabilities ) = 0;
      virtual HRESULT __stdcall CreateCapabilities (
        /*[in]*/ BSTR bstrUri,
        /*[out,retval]*/ struct IUccPresenceCapabilities * * ppCapabilities ) = 0;
      virtual HRESULT __stdcall AddCapabilities (
        /*[in]*/ struct IUccPresenceCapabilities * pCapabilities ) = 0;
      virtual HRESULT __stdcall RemoveCapabilities (
        /*[in]*/ struct IUccPresenceCapabilities * pCapabilities ) = 0;
      virtual HRESULT __stdcall RemoveAllCapabilities ( ) = 0;
      virtual HRESULT __stdcall get_MachineName (
        /*[out,retval]*/ BSTR * pbstrMachineName ) = 0;
      virtual HRESULT __stdcall put_MachineName (
        /*[in]*/ BSTR pbstrMachineName ) = 0;
};

struct __declspec(uuid("6a332cd1-594b-413b-9d42-55efc5a0a605"))
IUccPresenceDeviceInstance2 : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_DeviceUri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
};

struct __declspec(uuid("f0e4beea-90a7-48f4-b8f7-f3cce0d8a8fa"))
IUccPresenceNoteData : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_PRESENCE_NOTE_TYPE * penType ) = 0;
      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * pbstrURI ) = 0;
      virtual HRESULT __stdcall put_Uri (
        /*[in]*/ BSTR pbstrURI ) = 0;
      virtual HRESULT __stdcall get_StartTime (
        /*[out,retval]*/ DATE * pdStartTime ) = 0;
      virtual HRESULT __stdcall put_StartTime (
        /*[in]*/ DATE pdStartTime ) = 0;
      virtual HRESULT __stdcall get_EndTime (
        /*[out,retval]*/ DATE * pdEndTime ) = 0;
      virtual HRESULT __stdcall put_EndTime (
        /*[in]*/ DATE pdEndTime ) = 0;
      virtual HRESULT __stdcall get_Content (
        /*[out,retval]*/ struct IUccLocaleString * * pbstrContent ) = 0;
      virtual HRESULT __stdcall put_Content (
        /*[in]*/ struct IUccLocaleString * pbstrContent ) = 0;
      virtual HRESULT __stdcall CreateContent (
        /*[out,retval]*/ struct IUccLocaleString * * ppContent ) = 0;
};

struct __declspec(uuid("cca44790-b93f-4564-9135-384fcd1ed934"))
IUccPresenceNoteInstance : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateNote (
        /*[in]*/ enum UCC_PRESENCE_NOTE_TYPE enNoteType,
        /*[out,retval]*/ struct IUccPresenceNoteData * * ppNote ) = 0;
      virtual HRESULT __stdcall AddNote (
        /*[in]*/ struct IUccPresenceNoteData * pNote ) = 0;
      virtual HRESULT __stdcall RemoveNote (
        /*[in]*/ struct IUccPresenceNoteData * pNote ) = 0;
      virtual HRESULT __stdcall RemoveAllNotes ( ) = 0;
      virtual HRESULT __stdcall get_Notes (
        /*[out,retval]*/ struct IUccCollection * * ppNotes ) = 0;
};

enum UCC_PRESENCE_PHONE_LINE_TYPE
{
    UCCPPLT_UNIFIED_COMMUNCATIONS = 0,
    UCCPPLT_REMOTE_CALL_CONTROL = 1,
    UCCPPLT_DUAL = 2,
    UCCPPLT_NONE = 3
};

struct __declspec(uuid("9081029f-bcc4-4cf5-ac2d-440c97b8315b"))
IUccPresencePhoneLine : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_LineType (
        /*[out,retval]*/ enum UCC_PRESENCE_PHONE_LINE_TYPE * penType ) = 0;
      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * pbstrURI ) = 0;
      virtual HRESULT __stdcall get_LineServer (
        /*[out,retval]*/ BSTR * pbstrLineServer ) = 0;
};

enum UCC_ROUTING_RULE_ELEMENT_TYPE
{
    UCCRRET_PREAMBLE = 1
};

struct __declspec(uuid("b8bc87e3-1c8b-4f9e-8ab6-8be9a2c671af"))
IUccRoutingRuleElement : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_ROUTING_RULE_ELEMENT_TYPE * penType ) = 0;
};

struct __declspec(uuid("273819b7-8c74-4300-8c65-fb7e7dfc0146"))
IUccRoutingRuleInstance : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall put_Name (
        /*[in]*/ BSTR pbstrName ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pbstrName ) = 0;
      virtual HRESULT __stdcall put_RuleVersion (
        /*[in]*/ long lVersion ) = 0;
      virtual HRESULT __stdcall get_RuleVersion (
        /*[out,retval]*/ long * lVersion ) = 0;
      virtual HRESULT __stdcall get_Elements (
        /*[out,retval]*/ struct IUccCollection * * ppProperties ) = 0;
      virtual HRESULT __stdcall AddRoutingRuleElement (
        /*[in]*/ struct IUccRoutingRuleElement * pElement ) = 0;
      virtual HRESULT __stdcall RemoveRoutingRuleElement (
        /*[in]*/ struct IUccRoutingRuleElement * pElement ) = 0;
      virtual HRESULT __stdcall CreateRoutingRuleElement (
        /*[in]*/ enum UCC_ROUTING_RULE_ELEMENT_TYPE enType,
        /*[out,retval]*/ struct IUccRoutingRuleElement * * ppElement ) = 0;
};

enum UCC_ROUTING_PREAMBLE_ELEMENT_TYPE
{
    UCCRPET_LIST = 1,
    UCCRPET_WAIT = 2,
    UCCRPET_FLAGS = 3
};

struct __declspec(uuid("d0467f80-4d40-4668-8acd-aeef4f69b92c"))
IUccRoutingPreambleElement : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pbstrName ) = 0;
      virtual HRESULT __stdcall put_Name (
        /*[in]*/ BSTR pbstrName ) = 0;
      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_ROUTING_PREAMBLE_ELEMENT_TYPE * penType ) = 0;
};

struct __declspec(uuid("f28f34f5-9e9e-4166-8d6c-c9d1e6262487"))
IUccRoutingPreamble : IUccRoutingRuleElement
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Elements (
        /*[out,retval]*/ struct IUccCollection * * ppElements ) = 0;
      virtual HRESULT __stdcall AddElement (
        /*[in]*/ struct IUccRoutingPreambleElement * pElement ) = 0;
      virtual HRESULT __stdcall CreatePreambleElement (
        /*[in]*/ enum UCC_ROUTING_PREAMBLE_ELEMENT_TYPE enType,
        /*[out,retval]*/ struct IUccRoutingPreambleElement * * ppElement ) = 0;
};

struct __declspec(uuid("d624a32e-76b3-4d75-b48e-87a97d5fd786"))
IUccRoutingTarget : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * pbstrURI ) = 0;
      virtual HRESULT __stdcall get_Application (
        /*[out,retval]*/ BSTR * pbstrApplication ) = 0;
};

struct __declspec(uuid("0255c189-db59-4732-bb8c-73534860711c"))
IUccRoutingPreambleListElement : IUccRoutingPreambleElement
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Targets (
        /*[out,retval]*/ struct IUccCollection * * ppTargets ) = 0;
      virtual HRESULT __stdcall CreateTarget (
        /*[in]*/ BSTR bstrUri,
        /*[in]*/ BSTR bstrApplication,
        /*[out,retval]*/ struct IUccRoutingTarget * * ppTarget ) = 0;
      virtual HRESULT __stdcall AddTarget (
        /*[in]*/ struct IUccRoutingTarget * pTarget ) = 0;
      virtual HRESULT __stdcall RemoveTarget (
        /*[in]*/ struct IUccRoutingTarget * pTarget ) = 0;
};

struct __declspec(uuid("27c32a76-49b9-4a1a-b293-772e3d5a9020"))
IUccRoutingPreambleWaitElement : IUccRoutingPreambleElement
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Seconds (
        /*[out,retval]*/ long * plSeconds ) = 0;
      virtual HRESULT __stdcall put_Seconds (
        /*[in]*/ long plSeconds ) = 0;
};

struct __declspec(uuid("1e8e354c-7dcf-4ca3-a009-ed3876015076"))
IUccRoutingPreambleFlagsElement : IUccRoutingPreambleElement
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ BSTR * pbstr ) = 0;
      virtual HRESULT __stdcall put_Value (
        /*[in]*/ BSTR pbstr ) = 0;
};

struct __declspec(uuid("6689760d-24fa-4b3f-ac35-2042bfd70b89"))
IUccServerConfigurationCategory : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_AddressBookServerUrl (
        /*[in]*/ VARIANT_BOOL fExternal,
        /*[out,retval]*/ BSTR * pbstrAddressBookServerUri ) = 0;
      virtual HRESULT __stdcall get_DistributionGroupExpansionUrl (
        /*[in]*/ VARIANT_BOOL fExternal,
        /*[out,retval]*/ BSTR * pbstrDistributionGroupExpansionUri ) = 0;
      virtual HRESULT __stdcall get_ConsoleInstallationUrl (
        /*[in]*/ VARIANT_BOOL fExternal,
        /*[out,retval]*/ BSTR * pbstrConsoleInstallationUri ) = 0;
      virtual HRESULT __stdcall get_ConferenceTroubleshootingUrl (
        /*[in]*/ VARIANT_BOOL fExternal,
        /*[out,retval]*/ BSTR * pbstrConferenceTroubleshootingUri ) = 0;
      virtual HRESULT __stdcall get_OrganizationName (
        /*[out,retval]*/ BSTR * pbstrOrganizationName ) = 0;
      virtual HRESULT __stdcall get_FocusFactory (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall get_MeetingDisclaimer (
        /*[out,retval]*/ BSTR * pbstrMeetingDisclaimer ) = 0;
      virtual HRESULT __stdcall get_VoicemailUri (
        /*[out,retval]*/ BSTR * pbstrVoicemailUri ) = 0;
      virtual HRESULT __stdcall get_MediaRelayAuthenticationServerUri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall get_LocationProfile (
        /*[out,retval]*/ BSTR * pbstrUcLocationProfile ) = 0;
      virtual HRESULT __stdcall get_IsDistributionGroupExpansionEnabled (
        /*[out,retval]*/ VARIANT_BOOL * pbDistributionGroupExpansionEnabled ) = 0;
      virtual HRESULT __stdcall get_AudioVideoEncryption (
        /*[out,retval]*/ BSTR * pbstrAVEncryption ) = 0;
      virtual HRESULT __stdcall get_QualityMonitoringServerUri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
};

struct __declspec(uuid("3b4688bf-4d98-4d91-b3fd-c30fe0389597"))
IUccSignalingBodyPart : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_IsRequired (
        /*[out,retval]*/ VARIANT_BOOL * pbIsRequired ) = 0;
      virtual HRESULT __stdcall get_Accepted (
        /*[out,retval]*/ VARIANT_BOOL * pbAccepted ) = 0;
      virtual HRESULT __stdcall put_Accepted (
        /*[in]*/ VARIANT_BOOL pbAccepted ) = 0;
      virtual HRESULT __stdcall get_ContentType (
        /*[out,retval]*/ BSTR * pbstrContentType ) = 0;
      virtual HRESULT __stdcall put_ContentType (
        /*[in]*/ BSTR pbstrContentType ) = 0;
      virtual HRESULT __stdcall get_Body (
        /*[out,retval]*/ BSTR * pbstrBody ) = 0;
      virtual HRESULT __stdcall put_Body (
        /*[in]*/ BSTR pbstrBody ) = 0;
      virtual HRESULT __stdcall AddBodyPart (
        /*[in]*/ struct IUccSignalingBodyPart * pBodyPart ) = 0;
      virtual HRESULT __stdcall RemoveBodyPart (
        /*[in]*/ struct IUccSignalingBodyPart * pBodyPart ) = 0;
      virtual HRESULT __stdcall get_BodyParts (
        /*[out,retval]*/ struct IUccCollection * * ppBodyParts ) = 0;
      virtual HRESULT __stdcall AddHeaderField (
        /*[in]*/ BSTR bstrName,
        /*[in]*/ BSTR bstrValue,
        /*[out]*/ struct IUccSignalingHeaderField * * ppHeaderField ) = 0;
      virtual HRESULT __stdcall RemoveHeaderField (
        /*[in]*/ struct IUccSignalingHeaderField * pHeaderField ) = 0;
      virtual HRESULT __stdcall get_HeaderFields (
        /*[out,retval]*/ struct IUccCollection * * ppHeaderFields ) = 0;
      virtual HRESULT __stdcall get_ContentID (
        /*[out,retval]*/ BSTR * pbstrContentID ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * pContext ) = 0;
};

struct __declspec(uuid("bf78a2b4-cb02-4c00-875f-f51e3f0a74f2"))
IUccSignalingMessage2 : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateBodyPart (
        /*[in]*/ VARIANT_BOOL vbRequired,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccSignalingBodyPart * * pBodyPart ) = 0;
      virtual HRESULT __stdcall AddBodyPart (
        /*[in]*/ struct IUccSignalingBodyPart * pBodyPart ) = 0;
      virtual HRESULT __stdcall RemoveBodyPart (
        /*[in]*/ struct IUccSignalingBodyPart * pBodyPart ) = 0;
      virtual HRESULT __stdcall get_BodyParts (
        /*[out,retval]*/ struct IUccCollection * * ppBodyParts ) = 0;
};

struct __declspec(uuid("2285d731-7269-43f8-ac54-68a61ca8e1a8"))
IUccSignalingResponse : IUccSignalingMessage
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_StatusCode (
        /*[out,retval]*/ short * pVal ) = 0;
      virtual HRESULT __stdcall get_StatusText (
        /*[out,retval]*/ BSTR * pVal ) = 0;
};

struct __declspec(uuid("f17f94f7-bf07-4949-ba20-eb374f6d58c0"))
IUccUriManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall ParseUri (
        /*[in]*/ BSTR pbstrURI,
        /*[out,retval]*/ struct IUccUri * * ppUccUri ) = 0;
      virtual HRESULT __stdcall ConvertUri (
        /*[in]*/ BSTR bstrUri,
        /*[in]*/ BSTR bstrDomain,
        /*[out,retval]*/ struct IUccUri * * ppUccUri ) = 0;
};

struct __declspec(uuid("79f1c237-bd7d-4627-befd-813986373fcc"))
IUccApplicationSharingResource : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_APPLICATION_SHARING_RESOURCE_TYPE * penType ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccPropertyCollection * * ppProperties ) = 0;
      virtual HRESULT __stdcall get_Resources (
        /*[out,retval]*/ struct IUccCollection * * ppResources ) = 0;
      virtual HRESULT __stdcall get_Owner (
        /*[out,retval]*/ struct IUccApplicationSharingResource * * ppOwner ) = 0;
};

struct __declspec(uuid("32d73ef2-18a2-42f1-816c-ef4436719f40"))
IUccApplicationSharingChannelResource : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_NegotiatedStatus (
        /*[out,retval]*/ enum UCC_APPLICATION_SHARING_RESOURCE_STATUS * pStatus ) = 0;
      virtual HRESULT __stdcall get_PreferredStatus (
        /*[out,retval]*/ enum UCC_APPLICATION_SHARING_RESOURCE_STATUS * pStatus ) = 0;
      virtual HRESULT __stdcall put_PreferredStatus (
        /*[in]*/ enum UCC_APPLICATION_SHARING_RESOURCE_STATUS pStatus ) = 0;
};

struct __declspec(uuid("5383c344-0684-4cdf-8899-3af3963fa1e1"))
IUccApplicationSharingChannelResourceEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldNegotiatedStatus (
        /*[out,retval]*/ enum UCC_APPLICATION_SHARING_RESOURCE_STATUS * pStatus ) = 0;
      virtual HRESULT __stdcall get_NewNegotiatedStatus (
        /*[out,retval]*/ enum UCC_APPLICATION_SHARING_RESOURCE_STATUS * pStatus ) = 0;
};

struct __declspec(uuid("74236283-2f10-422a-a99f-da2b61585a53"))
IUccApplicationSharerChannelResource : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_LocalResource (
        /*[out,retval]*/ struct IUccApplicationSharingResource * * ppResource ) = 0;
};

struct __declspec(uuid("cd0d0e0f-79d8-422d-9bd5-e439812253a4"))
IUccApplicationSharingViewer : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Window (
        /*[out,retval]*/ IUnknown * * ppOleObject ) = 0;
      virtual HRESULT __stdcall get_BoundingRectangle (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * ppCoordinates ) = 0;
};

struct __declspec(uuid("5247ac0e-a9f3-4514-baac-ca7e9f52180f"))
IUccApplicationViewerChannelResource : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Viewer (
        /*[out,retval]*/ struct IUccApplicationSharingViewer * * ppViewer ) = 0;
};

struct __declspec(uuid("b7112800-b839-4f64-8ca0-5ce22d1d6666"))
IUccApplicationSharingResourceManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Resources (
        /*[out,retval]*/ struct IUccCollection * * ppResources ) = 0;
};

struct __declspec(uuid("7d6773b5-19a5-4668-8811-5510dce5bee2"))
IUccApplicationSharingSession : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall put_EncryptionLevel (
        /*[in]*/ enum UCC_MEDIA_ENCRYPTION_LEVEL penEncryptionLevel ) = 0;
      virtual HRESULT __stdcall get_EncryptionLevel (
        /*[out,retval]*/ enum UCC_MEDIA_ENCRYPTION_LEVEL * penEncryptionLevel ) = 0;
      virtual HRESULT __stdcall get_MediaConnectivityDisabled (
        /*[out,retval]*/ VARIANT_BOOL * pvbDisableICE ) = 0;
      virtual HRESULT __stdcall put_MediaConnectivityDisabled (
        /*[in]*/ VARIANT_BOOL pvbDisableICE ) = 0;
};

struct __declspec(uuid("11ceee73-aad1-40f5-a09a-ff88e321049d"))
IUccApplicationSharingSessionParticipant : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateChannel (
        /*[in]*/ enum UCC_APPLICATION_SHARING_CHANNEL_TYPE enType,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccMediaChannel * * ppChannel ) = 0;
      virtual HRESULT __stdcall AddChannel (
        /*[in]*/ struct IUccMediaChannel * pChannel ) = 0;
      virtual HRESULT __stdcall RemoveChannel (
        /*[in]*/ struct IUccMediaChannel * pChannel ) = 0;
      virtual HRESULT __stdcall get_Channels (
        /*[out,retval]*/ struct IUccCollection * * ppChannels ) = 0;
      virtual HRESULT __stdcall UpdateChannels (
        /*[in]*/ struct IUccOperationContext * pContext ) = 0;
};

struct __declspec(uuid("3f25cb45-e64f-4a0d-865b-0822d62ac2cc"))
IUccApplicationSharingChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateResource (
        /*[in]*/ struct IUccApplicationSharingResource * pResource,
        /*[out,retval]*/ struct IUccApplicationSharingChannelResource * * ppChannelResource ) = 0;
      virtual HRESULT __stdcall get_Controllers (
        /*[out,retval]*/ struct IUccCollection * * ppControllers ) = 0;
};

struct __declspec(uuid("f6eaba3c-69d7-41cf-8710-d1300e8425d1"))
IUccApplicationSharingControlEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Controllers (
        /*[out,retval]*/ struct IUccCollection * * ppControllers ) = 0;
};

struct __declspec(uuid("882ef9b5-92e6-466d-b31b-e7f6a4c3b57f"))
IUccApplicationSharingControlRequestAllowedEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ControlRequestAllowed (
        /*[out,retval]*/ VARIANT_BOOL * pvaGrantControl ) = 0;
};

struct __declspec(uuid("a5951db4-3d41-4af0-a360-40e2af383527"))
IUccApplicationViewerChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall RequestControl (
        /*[in]*/ struct IUccOperationContext * pContext ) = 0;
      virtual HRESULT __stdcall ReleaseControl (
        /*[in]*/ struct IUccOperationContext * pContext ) = 0;
      virtual HRESULT __stdcall get_Viewer (
        /*[out,retval]*/ struct IUccApplicationSharingViewer * * ppViewer ) = 0;
      virtual HRESULT __stdcall get_ControlRequestAllowed (
        /*[out,retval]*/ VARIANT_BOOL * pvaControlRequestAllowed ) = 0;
};

struct __declspec(uuid("809f1175-58df-4b0b-8e7d-12c15050e0d4"))
IUccApplicationSharingConferenceChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Resources (
        /*[out,retval]*/ struct IUccCollection * * ppResources ) = 0;
};

struct __declspec(uuid("5d87afee-c509-4400-b780-3da13f94f174"))
IUccApplicationSharingChannelResources : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall AddResource (
        /*[in]*/ struct IUccApplicationSharingChannelResource * pResource ) = 0;
      virtual HRESULT __stdcall RemoveResource (
        /*[in]*/ struct IUccApplicationSharingChannelResource * pResource ) = 0;
      virtual HRESULT __stdcall get_Resources (
        /*[out,retval]*/ struct IUccCollection * * ppResources ) = 0;
};

struct __declspec(uuid("81c623e0-0167-457b-80e6-c089c9105eeb"))
IUccMediaChannelConnectionStateChangedEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum UCC_MEDIA_CHANNEL_CONNECTION_STATE * pVal ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum UCC_MEDIA_CHANNEL_CONNECTION_STATE * pVal ) = 0;
};

struct __declspec(uuid("0b804e06-2e95-40b6-9748-a2aaf7796585"))
_IUccApplicationSharingChannelResourceEvents : IDispatch
{};

struct __declspec(uuid("3677165d-f78e-4d54-9039-5a082294584e"))
_IUccApplicationSharingViewerEvents : IDispatch
{};

struct __declspec(uuid("a7d52eb2-8ae3-47e5-97af-c96a6f639c2a"))
_IUccApplicationSharingResourceManagerEvents : IDispatch
{};

enum UCC_APPLICATION_SHARING_RESOURCE_CHANGE_REASON
{
    UASRCR_RDP_TRIGGERED_CHANGE = 0,
    UASRCR_USER_INITIATED_CHANGE = 1,
    UASRCR_RESOURCE_SUPERCEDED = 2
};

struct __declspec(uuid("74543c38-8c7f-4eef-85c2-1f1eb281b0d8"))
IUccApplicationSharingResourceChangeEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Resource (
        /*[out,retval]*/ struct IUccApplicationSharingResource * * ppResource ) = 0;
      virtual HRESULT __stdcall get_Reason (
        /*[out,retval]*/ enum UCC_APPLICATION_SHARING_RESOURCE_CHANGE_REASON * penReason ) = 0;
};

struct __declspec(uuid("e9a00f18-3f9f-4094-8347-887fca4b2b20"))
_IUccApplicationSharingSessionParticipantEvents : IDispatch
{};

struct __declspec(uuid("414c2bb7-3f99-41db-a369-b5d74c31b5e9"))
_IUccApplicationSharingChannelEvents : IDispatch
{};

struct __declspec(uuid("58ad2cbd-83cf-4e58-8041-7d5d664a5287"))
_IUccApplicationSharerChannelEvents : IDispatch
{};

struct __declspec(uuid("fd372aa6-ab45-4fe8-a425-4d9c21717147"))
_IUccApplicationViewerChannelEvents : IDispatch
{};

struct __declspec(uuid("f4ddfc28-ee50-48df-92ae-09fc427cb855"))
_IUccApplicationSharingChannelResourcesEvents : IDispatch
{};

struct __declspec(uuid("5de958ab-a06c-4539-bb11-1d43b0cdf8bc"))
_IUccApplicationSharingChannelCollectionEvents : IDispatch
{};

struct __declspec(uuid("15f5e833-3dc1-4f9a-993e-6fac96a82bbe"))
_IUccPublicationManagerEvents : IDispatch
{};

struct __declspec(uuid("15f8e9d8-097d-470b-9bad-965cbcfa1e03"))
_IUccCategoryInstanceEvents : IDispatch
{};

struct __declspec(uuid("c9077987-cb46-4f9c-b3fd-254fc8fa3f98"))
_IUccCategoryContextEvents : IDispatch
{};

struct __declspec(uuid("717efb8d-956d-467d-8ccf-2a461fd630da"))
_IUccContactEvents : IDispatch
{};

struct __declspec(uuid("cff1d60f-10f4-4ca6-abc5-063fd96e988b"))
_IUccDelegateEvents : IDispatch
{};

struct __declspec(uuid("dd825224-63e0-4a35-bd54-9d3cbbedaabf"))
_IUccGroupEvents : IDispatch
{};

struct __declspec(uuid("bc909c17-90d1-4c2c-be68-f74728866afc"))
_IUccSubscriptionEvents : IDispatch
{};

struct __declspec(uuid("5c7e7df6-0b5a-4ec2-9f04-a875329b4ec1"))
_IUccPresentityEvents : IDispatch
{};

struct __declspec(uuid("7a3ca589-6955-4831-8795-6623c78f8de3"))
_IUccPublicationEvent : IDispatch
{};

struct __declspec(uuid("507298f3-7975-4817-8e35-33447e837fff"))
_IUccContainerEvents : IDispatch
{};

struct __declspec(uuid("70eb6c1a-9445-4d34-a9a5-4ca0423af676"))
_IUccContainerMemberEvents : IDispatch
{};

struct __declspec(uuid("edb173ee-ee09-407b-b0fe-aef5a895542f"))
_IUccOperationManagerEvents : IDispatch
{};

struct __declspec(uuid("cb2b887f-d037-4309-a504-31d4b3c41563"))
_IUccMediaDeviceManagerEvents : IDispatch
{};

struct __declspec(uuid("ec10e7e8-797e-4495-a86d-3e9eada6d5bb"))
UccPlatform;
    // [ default ] interface IUccPlatform
    // interface IUccMediaDeviceManager
    // interface IUccMediaDeviceSettings
    // interface IUccTraceSettings
    // [ default, source ] dispinterface _IUccPlatformEvents
    // [ source ] dispinterface _IUccMediaDeviceManagerEvents
    // [ source ] dispinterface _IUccApplicationSharingResourceManagerEvents

struct __declspec(uuid("319beda1-82d6-4de2-8321-bdd5df64e2f8"))
UccPropertyCollection;
    // [ default ] interface IUccPropertyCollection
    // interface IUccReadOnlyPropertyCollection
    // interface IUccCollection

struct __declspec(uuid("645c90e5-f88c-41b1-bf6b-f0d79e99b62b"))
UccOperationContext;
    // [ default ] interface IUccOperationContext

struct __declspec(uuid("3af68966-5194-4de0-bdab-ae03653d914b"))
UccContext;
    // [ default ] interface IUccContext
    // interface IUccCollection
    // interface IDispatch

struct __declspec(uuid("08c9df14-8b30-4120-a939-44d2e68a501f"))
UccSignalingBodyPart;
    // [ default ] interface IUccSignalingBodyPart

struct __declspec(uuid("07b919a9-79d7-4fe0-8d0b-a8103a672e2a"))
UccSignalingRequest;
    // [ default ] interface IUccSignalingRequest
    // interface IUccSignalingMessage
    // interface IUccSignalingMessage2

struct __declspec(uuid("17fe5a0e-851e-4f26-91d6-75045d458c32"))
UccSignalingResponse;
    // [ default ] interface IUccSignalingResponse
    // interface IUccSignalingMessage
    // interface IUccSignalingMessage2

struct __declspec(uuid("201d2dcd-af4f-484e-9793-56cac5ec0c16"))
UccUriManager;
    // [ default ] interface IUccUriManager

struct __declspec(uuid("85b9cf4d-7b3d-4cdb-a82c-f5cd6d6d0814"))
UccApplicationSession;
    // [ default ] interface IUccSession
    // [ default, source ] dispinterface _IUccSessionEvents
    // [ source ] dispinterface _IUccSessionParticipantCollectionEvents

struct __declspec(uuid("54083592-55a6-4b24-85bd-17dc561e74f6"))
UccApplicationSessionParticipant;
    // [ default ] interface IUccApplicationSessionParticipant
    // interface IUccSessionParticipant
    // [ default, source ] dispinterface _IUccApplicationSessionParticipantEvents
    // [ source ] dispinterface _IUccSessionParticipantEndpointCollectionEvents
    // [ source ] dispinterface _IUccSessionParticipantEvents
    // [ source ] dispinterface _IUccSignalingChannelEvents

struct __declspec(uuid("0c020766-8433-420d-b847-d68a3cef8037"))
UccAudioMediaChannel;
    // [ default ] interface IUccAudioMediaChannel
    // interface IUccMediaChannel
    // interface IUccMediaChannelDevices
    // interface IUccAudioVideoMediaChannel
    // [ default, source ] dispinterface _IUccAudioMediaChannelEvents
    // [ source ] dispinterface _IUccMediaChannelEvents
    // [ source ] dispinterface _IUccMediaChannelDevicesEvents
    // [ source ] dispinterface _IUccAudioVideoMediaChannelEvents

struct __declspec(uuid("51931504-36b5-4b59-a3f5-d4f1573d84ea"))
UccAudioMediaDevice;
    // [ default ] interface IUccAudioMediaDevice
    // interface IUccMediaDevice

struct __declspec(uuid("dc3e6105-12a2-4d27-b9ce-2fd9ec166033"))
UccAudioVideoSession;
    // [ default ] interface IUccAudioVideoSession
    // interface IUccSessionCallControl
    // interface IUccSession
    // [ default, source ] dispinterface _IUccAudioVideoSessionEvents
    // [ source ] dispinterface _IUccSessionEvents
    // [ source ] dispinterface _IUccSessionParticipantCollectionEvents
    // [ source ] dispinterface _IUccSessionCallControlEvents

struct __declspec(uuid("d7dadc80-4285-4ab5-a96b-8a60c3c64c41"))
UccAudioVideoSessionParticipant;
    // [ default ] interface IUccAudioVideoSessionParticipant
    // interface IUccSessionParticipant
    // interface IUccSignalingChannel
    // interface IUccOperationManager
    // [ default, source ] dispinterface _IUccAudioVideoSessionParticipantEvents
    // [ source ] dispinterface _IUccSessionParticipantEvents
    // [ source ] dispinterface _IUccSessionParticipantEndpointCollectionEvents
    // [ source ] dispinterface _IUccMediaChannelCollectionEvents
    // [ source ] dispinterface _IUccSignalingChannelEvents
    // [ source ] dispinterface _IUccOperationManagerEvents

struct __declspec(uuid("b351467a-29fe-4863-9009-d024798da02a"))
UccConferenceEntityView;
    // [ default ] interface IUccConferenceEntityView
    // [ default, source ] dispinterface _IUccConferenceEntityViewEvents

struct __declspec(uuid("12d20681-68ac-4042-b93c-8a0e19f9a0f3"))
UccConferenceInformation;
    // [ default ] interface IUccConferenceInformation

struct __declspec(uuid("e5212020-d0d7-4dd4-bcd1-875d869009e1"))
UccConferenceManagerSession;
    // [ default ] interface IUccConferenceManagerSession
    // [ default, source ] dispinterface _IUccConferenceManagerSessionEvents

struct __declspec(uuid("97acd854-7422-4667-83a7-3a877523ba3b"))
UccConferenceMediaChannel;
    // [ default ] interface IUccConferenceMediaChannel
    // interface IUccMediaChannel
    // [ default, source ] dispinterface _IUccConferenceMediaChannelEvents
    // [ source ] dispinterface _IUccMediaChannelEvents

struct __declspec(uuid("d044bcbe-3323-432b-a74a-baecab49bbd2"))
UccConferenceView;
    // [ default ] interface IUccConferenceView
    // [ default, source ] dispinterface _IUccConferenceEntityViewCollectionEvents

struct __declspec(uuid("7ad3396b-6fbd-4396-8d5c-92c7f5395aa0"))
UccConferenceSession;
    // [ default ] interface IUccConferenceSession
    // interface IUccConferenceSession2
    // interface IUccSession
    // [ default, source ] dispinterface _IUccConferenceSessionEvents
    // [ source ] dispinterface _IUccConferenceSessionEvents2
    // [ source ] dispinterface _IUccSessionParticipantCollectionEvents
    // [ source ] dispinterface _IUccSessionEvents

struct __declspec(uuid("98078c74-d34b-4625-b143-cf34c2d60633"))
UccConferenceSessionParticipant;
    // [ default ] interface IUccConferenceSessionParticipant
    // interface IUccSessionParticipant
    // [ default, source ] dispinterface _IUccConferenceSessionParticipantEvents
    // [ source ] dispinterface _IUccSessionParticipantEndpointCollectionEvents
    // [ source ] dispinterface _IUccSessionParticipantEvents

struct __declspec(uuid("b777414a-4fcd-4142-bbad-1cf69cb19381"))
UccConferenceSessionParticipantEndpoint;
    // [ default ] interface IUccConferenceSessionParticipantEndpoint
    // interface IUccSessionParticipantEndpoint
    // [ default, source ] dispinterface _IUccConferenceSessionParticipantEndpointEvents
    // [ source ] dispinterface _IUccConferenceMediaChannelCollectionEvents

struct __declspec(uuid("553a1e8d-907f-490b-bc05-8c9744036bf9"))
UccCredential;
    // [ default ] interface IUccCredential

struct __declspec(uuid("35e7fe7a-1c37-48dc-bf7b-210c61fb32e5"))
UccCredentialCache;
    // [ default ] interface IUccCredentialCache

struct __declspec(uuid("5cd2f1ba-371b-4c43-8d70-e07dbd4fd96f"))
UccDiagnosticInfo;
    // [ default ] interface IUccDiagnosticInfo

struct __declspec(uuid("2d85841a-c8a1-4dab-94bc-d72ac874b195"))
UccHistoryInfoEntry;
    // [ default ] interface IUccHistoryInfoEntry

struct __declspec(uuid("e308443e-e15a-4afc-857e-fccd8b56d97b"))
UccFindServerEvent;
    // [ default ] interface IUccFindServerEvent
    // interface IUccOperationProgressEvent

struct __declspec(uuid("8f929b47-bf58-4d77-a951-005f99fea6a5"))
UccInstantMessagingSession;
    // [ default ] interface IUccInstantMessagingSession
    // interface IUccSessionCallControl
    // interface IUccSession
    // [ default, source ] dispinterface _IUccInstantMessagingSessionEvents
    // [ source ] dispinterface _IUccSessionEvents
    // [ source ] dispinterface _IUccSessionParticipantCollectionEvents
    // [ source ] dispinterface _IUccSessionCallControlEvents

struct __declspec(uuid("eff71cf7-fec4-458d-a7ce-283145796ca2"))
UccInstantMessagingSessionParticipant;
    // [ default ] interface IUccInstantMessagingSessionParticipant
    // interface IUccSignalingChannel
    // interface IUccSessionParticipant
    // [ default, source ] dispinterface _IUccInstantMessagingSessionParticipantEvents
    // [ source ] dispinterface _IUccSessionParticipantEvents
    // [ source ] dispinterface _IUccSignalingChannelEvents
    // [ source ] dispinterface _IUccSessionParticipantEndpointCollectionEvents

struct __declspec(uuid("c7817076-d059-4603-9f51-c02ba0be82bb"))
UccMediaConnectivityServer;
    // [ default ] interface IUccMediaConnectivityServer

struct __declspec(uuid("be6cb026-0e90-46db-8129-0750961111ea"))
UccMediaPlayerDevice;
    // [ default ] interface IUccMediaPlayerDevice
    // interface IUccMediaDevice

struct __declspec(uuid("6f67517b-135b-4a42-8a55-5990c36b265c"))
UccProperty;
    // [ default ] interface IUccProperty

struct __declspec(uuid("87f779b7-f044-46f4-8a7f-afee8203fab2"))
UccProvisioningPolicyInstance;
    // [ default ] interface IUccProvisioningPolicyInstance
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccCategoryInstanceEvents

struct __declspec(uuid("b9634df3-a19d-4bdb-befc-845519c5d7ff"))
UccProxyEndpoint;
    // [ default ] interface IUccEndpoint
    // interface IUccEndpoint2
    // interface IUccSessionManager
    // interface IUccProxyEndpoint
    // [ default, source ] dispinterface _IUccEndpointEvents
    // [ source ] dispinterface _IUccSessionManagerEvents

struct __declspec(uuid("a7cb24ca-9c80-40ae-9390-e104efdbdebf"))
UccRemoteCallControlSession;
    // [ default ] interface IUccSessionCallControl
    // interface IUccSession
    // [ default, source ] dispinterface _IUccSessionCallControlEvents
    // [ source ] dispinterface _IUccSessionEvents
    // [ source ] dispinterface _IUccSessionParticipantCollectionEvents

struct __declspec(uuid("c28a10be-4d70-49a7-87e6-012eb02b9bec"))
UccRemoteCallControlSessionParticipant;
    // [ default ] interface IUccAudioVideoSessionParticipant
    // interface IUccSessionParticipant
    // [ default, source ] dispinterface _IUccAudioVideoSessionParticipantEvents
    // [ source ] dispinterface _IUccSessionParticipantEvents
    // [ source ] dispinterface _IUccSessionParticipantEndpointCollectionEvents

struct __declspec(uuid("023e043b-720b-43e8-bee2-7c4f710415dc"))
UccRemoteCallControlAudioChannel;
    // [ default ] interface IUccAudioMediaChannel
    // interface IUccMediaChannel
    // [ default, source ] dispinterface _IUccMediaChannelEvents

struct __declspec(uuid("3af314a7-edb7-42cf-915b-ae50e1602f6a"))
UccServerEndpoint;
    // [ default ] interface IUccEndpoint
    // interface IUccMediaEndpointSettings
    // interface IUccServerSignalingSettings
    // interface IUccSessionManager
    // interface IUccSubscriptionManager
    // interface IUccPublicationManager
    // interface IUccSignalingChannelManager
    // interface IUccConferenceManager
    // interface IUccDiagnosticReportingSettings
    // interface IUccUserSearchManager
    // interface IUccEndpoint2
    // [ default, source ] dispinterface _IUccEndpointEvents
    // [ source ] dispinterface _IUccServerSignalingSettingsEvents
    // [ source ] dispinterface _IUccSessionManagerEvents
    // [ source ] dispinterface _IUccPublicationManagerEvents
    // [ source ] dispinterface _IUccMediaEndpointEvents

struct __declspec(uuid("af9b29ec-84a0-482b-8e8a-2bd307d1eb24"))
IUccSignalingChannelManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateSignalingChannel (
        /*[in]*/ struct IUccUri * pDestinationUri,
        /*[out,retval]*/ struct IUccSignalingChannel * * ppSignalingChannel ) = 0;
};

struct __declspec(uuid("92391a7b-9135-42fa-a478-9e76cbfb7f3c"))
UccPrincipalServerEndpoint;
    // [ default ] interface IUccEndpoint
    // interface IUccMediaEndpointSettings
    // interface IUccServerSignalingSettings
    // interface IUccSessionManager
    // interface IUccSubscriptionManager
    // interface IUccPublicationManager
    // interface IUccSignalingChannelManager
    // interface IUccConferenceManager
    // interface IUccDiagnosticReportingSettings
    // interface IUccUserSearchManager
    // interface IUccEndpoint2
    // interface IUccCategoryTransformationManager
    // [ default, source ] dispinterface _IUccEndpointEvents
    // [ source ] dispinterface _IUccEndpointEvents2
    // [ source ] dispinterface _IUccServerSignalingSettingsEvents
    // [ source ] dispinterface _IUccSessionManagerEvents
    // [ source ] dispinterface _IUccPublicationManagerEvents
    // [ source ] dispinterface _IUccMediaEndpointEvents
    // [ source ] dispinterface _IUccProxyEndpointCollectionEvents

struct __declspec(uuid("45d01e9f-6d39-4a00-a9c9-979bb9157c1b"))
UccDelegateEndpoint;
    // [ default ] interface IUccEndpoint
    // interface IUccProxyEndpoint
    // [ default, source ] dispinterface _IUccEndpointEvents
    // [ source ] dispinterface _IUccSessionManagerEvents
    // [ source ] dispinterface _IUccPublicationManagerEvents

struct __declspec(uuid("284e107f-92f4-4091-a0db-e1535bfd9441"))
UccAssociatedUriEndpoint;
    // [ default ] interface IUccEndpoint
    // interface IUccProxyEndpoint
    // [ default, source ] dispinterface _IUccEndpointEvents
    // [ source ] dispinterface _IUccSessionManagerEvents

struct __declspec(uuid("4dc102e6-e499-45f3-852b-b43bf072e409"))
UccSignalingChannel;
    // [ default ] interface IUccSignalingChannel
    // [ default, source ] dispinterface _IUccSignalingChannelEvents

struct __declspec(uuid("7ef89df3-d4ab-4e3d-b0f0-814d6503d17b"))
UccVideoMediaChannel;
    // [ default ] interface IUccVideoMediaChannel
    // interface IUccMediaChannelDevices
    // interface IUccAudioVideoMediaChannel
    // [ default, source ] dispinterface _IUccMediaChannelEvents
    // [ source ] dispinterface _IUccMediaChannelDevicesEvents
    // [ source ] dispinterface _IUccAudioVideoMediaChannelEvents

struct __declspec(uuid("4c6fc976-1b74-4775-b404-462b62807e77"))
UccVideoMediaDevice;
    // [ default ] interface IUccVideoMediaDevice
    // interface IUccMediaDevice

struct __declspec(uuid("d9a93e9b-2991-4932-8202-212293dad181"))
UccAudioVideoMediaChannelStreamsEvent;
    // [ default ] interface IUccAudioVideoMediaChannelStreamsEvent

struct __declspec(uuid("85dd8879-b33f-4439-a0e7-4be6dca813be"))
UccAudioVideoMediaChannelSourceEvent;
    // [ default ] interface IUccAudioVideoMediaChannelSourceEvent

struct __declspec(uuid("4b7e85ca-c96f-4b00-b3e0-da2728f2cbe7"))
UccAudioSignalLevelChangedEvent;
    // [ default ] interface IUccAudioSignalLevelChangedEvent

struct __declspec(uuid("76e131bc-3b48-479f-a008-734211176334"))
UccIncomingMediaRequestEvent;
    // [ default ] interface IUccIncomingMediaRequestEvent
    // interface IUccIncomingRequestEvent

struct __declspec(uuid("88c1caf3-028d-4078-ab15-dc926daeb909"))
UccIncomingInstantMessageEvent;
    // [ default ] interface IUccIncomingInstantMessageEvent
    // interface IUccIncomingRequestEvent

struct __declspec(uuid("5a764ba5-7409-4863-85fa-ac12cade8149"))
UccIncomingInvitationEvent;
    // [ default ] interface IUccIncomingInvitationEvent
    // interface IUccIncomingRequestEvent

struct __declspec(uuid("f4c31793-7a53-4e4a-8761-b193b666ac1f"))
UccIncomingSessionEvent;
    // [ default ] interface IUccIncomingSessionEvent
    // interface IUccIncomingRequestEvent

struct __declspec(uuid("de455295-8d70-469c-91f8-4b21dab19243"))
UccIncomingSessionTransferEvent;
    // [ default ] interface IUccIncomingSessionTransferEvent
    // interface IUccIncomingRequestEvent

struct __declspec(uuid("af754b4f-53eb-41a0-ade5-7c680a025c3b"))
UccInstantMessagingComposingEvent;
    // [ default ] interface IUccInstantMessagingComposingEvent

struct __declspec(uuid("6e213239-59f9-4c14-9bb2-8406985fd5ca"))
UccInvitationAcceptedEvent;
    // [ default ] interface IUccInvitationAcceptedEvent

struct __declspec(uuid("4833dc82-cfe5-4c75-8d86-591a8fd361b5"))
UccMediaChannelEvent;
    // [ default ] interface IUccMediaChannelEvent

struct __declspec(uuid("65e014d2-c007-4dc7-898b-92d99230b126"))
UccOperationProgressEvent;
    // [ default ] interface IUccOperationProgressEvent

struct __declspec(uuid("0bfeb2bd-8bc6-47ad-8705-58496a3bc49c"))
UccOutgoingInvitationEvent;
    // [ default ] interface IUccOutgoingInvitationEvent

struct __declspec(uuid("97095729-969c-4e23-9b43-54bc2eae4268"))
UccOutgoingSessionEvent;
    // [ default ] interface IUccOutgoingSessionEvent

struct __declspec(uuid("70c76318-cb44-4a52-8308-d8d00d7d6ea6"))
UccPropertyUpdateEvent;
    // [ default ] interface IUccPropertyUpdateEvent

struct __declspec(uuid("bfef2af5-ad05-48a0-8bbc-fba89717fbc9"))
UccNamedPropertyUpdateEvent;
    // [ default ] interface IUccNamedPropertyUpdateEvent

struct __declspec(uuid("368be194-d23d-4724-a741-e606714f1e2c"))
UccSessionOperationEvent;
    // [ default ] interface IUccSessionOperationEvent

struct __declspec(uuid("155aa654-f81f-4a0f-b5ae-e2131f3ad057"))
UccSessionParticipantOperationEvent;
    // [ default ] interface IUccSessionParticipantOperationEvent
    // interface IUccOperationProgressEvent

struct __declspec(uuid("dfa87481-a0d1-484a-bf25-0b294f90d75b"))
UccSessionParticipantCollectionEvent;
    // [ default ] interface IUccSessionParticipantCollectionEvent

struct __declspec(uuid("68f95e30-7ba8-487b-89cd-20b657fc4fdd"))
UccSessionParticipantEndpointCollectionEvent;
    // [ default ] interface IUccSessionParticipantEndpointCollectionEvent

struct __declspec(uuid("b81f580d-9ec1-4833-beca-b35194ef88ec"))
UccSessionParticipantStateChangedEvent;
    // [ default ] interface IUccSessionParticipantStateChangedEvent
    // interface IUccOperationProgressEvent

struct __declspec(uuid("79dcfdf6-5778-4484-bcfd-3725c0efceab"))
UccSessionTransferProgressEvent;
    // [ default ] interface IUccSessionTransferProgressEvent
    // interface IUccOperationProgressEvent

struct __declspec(uuid("645fa567-1348-4c39-8d1a-8be5a49a5bdf"))
UccConferenceEntityViewCollectionEvent;
    // [ default ] interface IUccConferenceEntityViewCollectionEvent

struct __declspec(uuid("31911030-172f-4e99-aa09-372f5b8ccd77"))
UccConferenceSessionParticipantEndpointStateChangedEvent;
    // [ default ] interface IUccConferenceSessionParticipantEndpointStateChangedEvent
    // interface IUccOperationProgressEvent

struct __declspec(uuid("e8905a59-86e6-4e1f-998c-74edb9c0c34f"))
UccMediaDeviceChangeEvent;
    // [ default ] interface IUccMediaDeviceChangeEvent

struct __declspec(uuid("abfdaa76-8dd7-474e-8102-493851c7c963"))
UccProxyEndpointCollectionEvent;
    // [ default ] interface IUccProxyEndpointCollectionEvent

struct __declspec(uuid("2b7e2ca0-4ed4-4fb9-b997-96ca570432be"))
UccPresenceStateInstance;
    // [ default ] interface IUccPresenceStateInstance
    // interface IUccPresenceStateInstance2
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccCategoryInstanceEvents

struct __declspec(uuid("def85737-89ea-41cb-a2eb-4167874d9304"))
UccCategoryContext;
    // [ default ] interface IUccCategoryContext
    // [ default, source ] dispinterface _IUccCategoryContextEvents

struct __declspec(uuid("2480d71f-dbe2-4b25-bdf7-06c1ae4ad976"))
UccCategoryContextEvent;
    // [ default ] interface IUccCategoryContextEvent

struct __declspec(uuid("53e07ea1-87d9-4cd3-b206-d93bfbc2a6c5"))
UccCategoryInstance;
    // [ default ] interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccCategoryInstanceEvents

struct __declspec(uuid("48239b4b-3195-4c31-a49f-6f4d53300971"))
UccCategoryInstanceEvent;
    // [ default ] interface IUccCategoryInstanceEvent

struct __declspec(uuid("b87bef40-a3eb-4bc7-81e2-3f9e24ed4d9b"))
UccGroupMembershipEvent;
    // [ default ] interface IUccGroupMembershipEvent

struct __declspec(uuid("c1bbecc2-c8e0-45be-9ec2-d629d6c3beb9"))
UccContainerMemberCollectionEvent;
    // [ default ] interface IUccContainerMemberCollectionEvent

struct __declspec(uuid("864272a1-5266-4e08-b6a2-6aeabe6c1f74"))
UccContainerMembershipEvent;
    // [ default ] interface IUccContainerMembershipEvent

struct __declspec(uuid("fee4f779-944a-4afa-80e3-1ec9b2a3b3a9"))
UccContactCollectionEvent;
    // [ default ] interface IUccContactCollectionEvent

struct __declspec(uuid("88a91ddd-0bf7-4536-8687-2e20726b6cb2"))
UccSubscription;
    // [ default ] interface IUccSubscription
    // [ default, source ] dispinterface _IUccSubscriptionEvents

struct __declspec(uuid("1cbd2388-9a85-4e72-8bcf-fd8cf90218a0"))
UccPresenceContactCardIdentity;
    // [ default ] interface IUccPresenceContactCardIdentity

struct __declspec(uuid("7a26be48-15b2-473c-8bfe-b80e74c1257b"))
UccLocaleString;
    // [ default ] interface IUccLocaleString

struct __declspec(uuid("690ff5c9-dbc8-4979-b4b6-c8d727ff147d"))
UccPresenceContactCardAddress;
    // [ default ] interface IUccPresenceContactCardAddress

struct __declspec(uuid("8f59329c-8167-4069-8f56-6ca149282b57"))
UccPresenceContactCardPhone;
    // [ default ] interface IUccPresenceContactCardPhone

struct __declspec(uuid("094142c4-eb57-405e-b7a8-7f8d5d35dbde"))
UccPresenceContactCardUrl;
    // [ default ] interface IUccPresenceContactCardUrl

struct __declspec(uuid("9d3857ea-1000-4cd0-a78b-ad375f68512d"))
UccPresenceContactCardInstance;
    // [ default ] interface IUccPresenceContactCardInstance
    // interface IUccPresenceContactCardInstance2
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccCategoryInstanceEvents

struct __declspec(uuid("84b09c4b-3d6a-4e0c-acd8-e4e5c6b9c35b"))
UccPresenceFreeBusyInformation;
    // [ default ] interface IUccPresenceFreeBusyInformation

struct __declspec(uuid("49de5964-434b-4f82-9d33-26ab217a0129"))
UccPresenceTimeZoneInformation;
    // [ default ] interface IUccPresenceTimeZoneInformation

struct __declspec(uuid("e904a0d9-b6a4-4266-8728-644693633d92"))
UccPresenceWorkingHoursBlock;
    // [ default ] interface IUccPresenceWorkingHoursBlock

struct __declspec(uuid("c88fac1b-8e0b-432f-b2ae-fc5124059e95"))
UccPresenceWorkingHoursInformation;
    // [ default ] interface IUccPresenceWorkingHoursInformation

struct __declspec(uuid("34e0c61a-23f7-454a-be6f-2f8c3fa725c1"))
UccPresenceActivity;
    // [ default ] interface IUccPresenceActivity

struct __declspec(uuid("0f04a7a5-3624-468c-a5d8-58a790839eb0"))
UccPresenceCapabilities;
    // [ default ] interface IUccPresenceCapabilities

struct __declspec(uuid("4262dfa6-7a9b-4eba-bcc5-ec02e9ab8842"))
UccPresenceCapabilityInfo;
    // [ default ] interface IUccPresenceCapabilityInfo

struct __declspec(uuid("967baf31-115d-43d9-be38-72780a7229ed"))
UccPresenceCalendarDataInstance;
    // [ default ] interface IUccPresenceCalendarDataInstance
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccCategoryInstanceEvents

struct __declspec(uuid("0c4e3757-5d0a-4f7a-925e-247c75fb57a4"))
UccPresenceServicesInstance;
    // [ default ] interface IUccPresenceServicesInstance
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccCategoryInstanceEvents

struct __declspec(uuid("b32e9333-4edd-4de2-b671-454544410c4a"))
UccPresenceServiceElement;
    // [ default ] interface IUccPresenceServiceElement

struct __declspec(uuid("62f1a861-25f8-4c86-aa14-7a817fdfb0fe"))
UccPresenceDeviceInstance;
    // [ default ] interface IUccPresenceDeviceInstance
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccCategoryInstanceEvents

struct __declspec(uuid("f16ed011-533d-4041-9135-ca07a52d7418"))
UccPresenceNoteData;
    // [ default ] interface IUccPresenceNoteData

struct __declspec(uuid("aa77632f-dd46-48ee-8ed3-153ef4206d7f"))
UccPresenceNoteInstance;
    // [ default ] interface IUccPresenceNoteInstance
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccCategoryInstanceEvents

struct __declspec(uuid("18b74f75-8884-405b-a2dc-da17309c36b0"))
UccPresentity;
    // [ default ] interface IUccPresentity
    // [ default, source ] dispinterface _IUccPresentityEvents

struct __declspec(uuid("9126e974-c78d-44d0-9d1c-33d691aada12"))
UccServerConfigurationCategory;
    // [ default ] interface IUccServerConfigurationCategory
    // interface IUccCategoryContext
    // [ default, source ] dispinterface _IUccCategoryInstanceEvents

struct __declspec(uuid("e69572c9-3681-442f-b4cc-4bcc307a37a9"))
UccSignalingHeaderField;
    // [ default ] interface IUccSignalingHeaderField

struct __declspec(uuid("adc2ab82-b891-46bf-a41d-68411381b3d8"))
UccSignalingHeaderFieldParameter;
    // [ default ] interface IUccSignalingHeaderFieldParameter

struct __declspec(uuid("a0ac85b9-905c-4375-b614-d5d320cd41cd"))
UccUri;
    // [ default ] interface IUccUri

struct __declspec(uuid("f5a738e2-4a28-4f3d-978e-a1f867f56485"))
UccContact;
    // [ default ] interface IUccContact
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccContactEvents

struct __declspec(uuid("b1910c3e-9305-454c-b61e-da580e708b87"))
UccGroup;
    // [ default ] interface IUccGroup
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccGroupEvents

struct __declspec(uuid("505fde43-e305-481d-8341-6168dd28617f"))
UccDelegate;
    // [ default ] interface IUccDelegate
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccDelegateEvents

struct __declspec(uuid("9ea23fc2-fe8f-4508-b8c0-73af8e185075"))
UccDelegatee;
    // [ default ] interface IUccDelegatee
    // interface IUccCategoryInstance

struct __declspec(uuid("613c60f6-e524-4123-bd2e-b3b183ad37e8"))
UccRoutingRuleInstance;
    // [ default ] interface IUccRoutingRuleInstance
    // interface IUccCategoryInstance
    // [ default, source ] dispinterface _IUccCategoryInstanceEvents

struct __declspec(uuid("50034ab7-1512-4fd6-bb2f-a566c7bd3291"))
UccRoutingPreamble;
    // [ default ] interface IUccRoutingPreamble
    // interface IUccRoutingRuleElement

struct __declspec(uuid("be62eafa-4f29-4a44-983f-7a1de8e0b92f"))
UccRoutingPreambleListElement;
    // [ default ] interface IUccRoutingPreambleListElement
    // interface IUccRoutingPreambleElement

struct __declspec(uuid("716a4b4e-4f09-4c14-b34c-3c0c0ad9566f"))
UccRoutingTarget;
    // [ default ] interface IUccRoutingTarget

struct __declspec(uuid("3e9496ad-527d-458c-b336-79d3ba773455"))
UccRoutingPreambleWaitElement;
    // [ default ] interface IUccRoutingPreambleWaitElement
    // interface IUccRoutingPreambleElement

struct __declspec(uuid("cc36abb1-b62c-402e-9a29-41395369340c"))
UccRoutingPreambleFlagsElement;
    // [ default ] interface IUccRoutingPreambleFlagsElement
    // interface IUccRoutingPreambleElement

struct __declspec(uuid("aa60bbd9-9fae-415e-84f1-ee51e89c0d09"))
UccSubscriptionEvent;
    // [ default ] interface IUccSubscriptionEvent

struct __declspec(uuid("6a4a9729-ee1c-47db-990e-5387519ec928"))
UccPresencePhoneLine;
    // [ default ] interface IUccPresencePhoneLine

struct __declspec(uuid("f01ed73c-e9cc-4edd-bbd9-abf4bba0d3d6"))
UccUserSearchQuery;
    // [ default ] interface IUccUserSearchQuery
    // [ default, source ] dispinterface _IUccUserSearchQueryEvents

struct __declspec(uuid("fb84e8ba-d076-45a3-960d-50a88b0506b0"))
UccUserSearchQueryEvent;
    // [ default ] interface IUccUserSearchQueryEvent

struct __declspec(uuid("40f200e5-e499-4130-9ee6-5a9f4139fe48"))
UccUserSearchResultRecord;
    // [ default ] interface IUccUserSearchResultRecord

struct __declspec(uuid("e1d00d06-988a-47c5-a081-3ee6e72c4739"))
UccCategoryInstanceTransformationManifest;
    // [ default ] interface IUccCategoryInstanceTransformationManifest

struct __declspec(uuid("63dda457-200b-40f1-aff4-deb5540807ac"))
UccEndpointRefreshEvent;
    // [ default ] interface IUccEndpointRefreshEvent

struct __declspec(uuid("7f14534b-5317-4aa3-9082-f9b4c38870e5"))
UccApplicationSharingSession;
    // [ default ] interface IUccApplicationSharingSession
    // interface IUccSession

struct __declspec(uuid("50e3fc54-bf78-4342-83cf-28f837e38ba4"))
UccApplicationSharingSessionParticipant;
    // [ default ] interface IUccApplicationSharingSessionParticipant
    // interface IUccSessionParticipant
    // [ default, source ] dispinterface _IUccApplicationSharingSessionParticipantEvents
    // [ source ] dispinterface _IUccApplicationSharingChannelCollectionEvents
    // [ source ] dispinterface _IUccSessionParticipantEvents
    // [ source ] dispinterface _IUccSessionParticipantEndpointCollectionEvents

struct __declspec(uuid("a81ca775-181b-4728-ab56-f307ab9a3d9c"))
UccApplicationSharingResourceChangeEvent;
    // [ default ] interface IUccApplicationSharingResourceChangeEvent

struct __declspec(uuid("ce008d78-2bec-4b7a-9b10-12b9e1ed1bed"))
UccApplicationSharingRequestControlEvent;
    // [ default ] interface IUccApplicationSharingRequestControlEvent

struct __declspec(uuid("5618d4f8-bbc8-4363-905f-bd126bfa17f6"))
UccApplicationSharingControlEvent;
    // [ default ] interface IUccApplicationSharingControlEvent

struct __declspec(uuid("5b673bde-c698-4991-a556-24f666d60553"))
UccApplicationSharingControlRequestAllowedEvent;
    // [ default ] interface IUccApplicationSharingControlRequestAllowedEvent

struct __declspec(uuid("2cff1652-10a0-46fb-a1d8-62ef685ffebd"))
UccConferenceMediaChannelSourceEvent;
    // [ default ] interface IUccConferenceMediaChannelSourceEvent

struct __declspec(uuid("f38cdd95-d1ba-4497-b84f-90c58292ed3d"))
UccMediaChannelConnectionStateChangedEvent;
    // [ default ] interface IUccMediaChannelConnectionStateChangedEvent

struct __declspec(uuid("cc4e82db-d848-4309-8c3c-473fc150b4eb"))
UccConferenceMediaChannelStateChangedEvent;
    // [ default ] interface IUccConferenceMediaChannelStateChangedEvent

struct __declspec(uuid("af404bf1-2495-4c3e-bf3b-be8359341f78"))
IUccConferenceMediaChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall SetProperty (
        /*[in]*/ enum UCC_CONFERENCE_MEDIA_CHANNEL_PROPERTY enPropertyId,
        /*[in]*/ VARIANT vPropertyValue,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * pVal ) = 0;
      virtual HRESULT __stdcall get_ParticipantEndpoint (
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * ppParticipantEndpoint ) = 0;
};

struct __declspec(uuid("16b1749a-39ce-4fd1-945c-b13d569261b9"))
IUccSessionParticipantEndpoint : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall get_Participant (
        /*[out,retval]*/ struct IUccSessionParticipant * * ppParticipant ) = 0;
      virtual HRESULT __stdcall get_IsLocal (
        /*[out,retval]*/ VARIANT_BOOL * pVal ) = 0;
};

struct __declspec(uuid("a0897f48-a177-4c66-a223-f7b8d695b6b8"))
IUccConferenceApplicationViewerChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ChannelSource (
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * ppParticipantEndpoint ) = 0;
};

struct __declspec(uuid("09717930-a1de-4b46-bae2-08f3e5c74fa2"))
IUccConferenceSessionParticipant : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * pVal ) = 0;
      virtual HRESULT __stdcall SetProperty (
        /*[in]*/ enum UCC_CONFERENCE_PARTICIPANT_PROPERTY enPropertyId,
        /*[in]*/ VARIANT vPropertyValue,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall PlayRecordedName (
        /*[in]*/ BSTR bstrEndpoint,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall CreateParticipantEndpoint (
        /*[in]*/ struct IUccUri * pUri,
        /*[in]*/ BSTR bstrEndpointId,
        /*[in]*/ enum UCC_CONFERENCE_ENTITY_TYPE enType,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccConferenceSessionParticipantEndpoint * * ppParticipantEndpoint ) = 0;
      virtual HRESULT __stdcall CopyParticipantEndpoint (
        /*[in]*/ struct IUccSessionParticipantEndpoint * pInputParticipantEndpoint,
        /*[in]*/ enum UCC_CONFERENCE_ENTITY_TYPE enType,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccConferenceSessionParticipantEndpoint * * ppParticipantEndpoint ) = 0;
};

struct __declspec(uuid("2c46dc00-1382-4c65-8b79-3cedb358e5d0"))
IUccAudioVideoMediaChannelSourceEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_PrimarySource (
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * ppEndpoint ) = 0;
      virtual HRESULT __stdcall get_PrimarySourceId (
        /*[out,retval]*/ unsigned long * pdwSourceId ) = 0;
      virtual HRESULT __stdcall get_ContributingSources (
        /*[out,retval]*/ struct IUccReadOnlyPropertyCollection * * ppContributingSources ) = 0;
      virtual HRESULT __stdcall get_Timestamp (
        /*[out,retval]*/ unsigned __int64 * pTimestamp ) = 0;
};

struct __declspec(uuid("e5837f40-2f55-458f-9adf-42c1bf973560"))
IUccConferenceMediaChannelSourceEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ChannelSource (
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * ppParticipantEndpoint ) = 0;
};

struct __declspec(uuid("ca7acf03-537d-4654-92f6-b556fde0cd80"))
IUccIncomingInstantMessageEvent : IUccIncomingRequestEvent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ParticipantEndpoint (
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * ppSenderEndpoint ) = 0;
      virtual HRESULT __stdcall get_ContentType (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall get_Content (
        /*[out,retval]*/ BSTR * pVal ) = 0;
      virtual HRESULT __stdcall IsBroadcast (
        /*[out,retval]*/ VARIANT_BOOL * pVal ) = 0;
};

struct __declspec(uuid("11ebeff7-e4c7-4bc3-8752-493b947004fa"))
IUccInstantMessagingComposingEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ParticipantEndpoint (
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * ppEndpoint ) = 0;
};

struct __declspec(uuid("683701e6-6c96-441a-8ab2-b5292dbb7104"))
IUccSessionParticipantEndpointCollectionEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Endpoint (
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * pVal ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * pVal ) = 0;
};

struct __declspec(uuid("8c37fed9-ab4a-45bc-b018-5ceefb94f34f"))
IUccSessionParticipantOperationEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ParticipantEndpoint (
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * ppParticipantEndpoint ) = 0;
      virtual HRESULT __stdcall get_Result (
        /*[out,retval]*/ struct IUccOperationProgressEvent * * ppResult ) = 0;
};

struct __declspec(uuid("1c88a56d-8f5f-4793-a990-e5333bfe8d37"))
IUccApplicationSharerChannel : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall RevokeControl (
        /*[in]*/ struct IUccSessionParticipantEndpoint * pEndpoint ) = 0;
      virtual HRESULT __stdcall GrantControl (
        /*[in]*/ struct IUccSessionParticipantEndpoint * pEndpoint ) = 0;
      virtual HRESULT __stdcall get_ControlRequestAllowed (
        /*[out,retval]*/ VARIANT_BOOL * pvaAutoGrantControl ) = 0;
      virtual HRESULT __stdcall put_ControlRequestAllowed (
        /*[in]*/ VARIANT_BOOL pvaAutoGrantControl ) = 0;
};

struct __declspec(uuid("ad1a9ef1-2584-4693-83c4-1f479d8b4fba"))
IUccApplicationSharingRequestControlEvent : IUccIncomingRequestEvent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ParticipantEndpoint (
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * ppParticipant ) = 0;
};

struct __declspec(uuid("c520c114-7ebf-4a14-8416-d17adfa86227"))
IUccSessionParticipant : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * ppContext ) = 0;
      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ enum UCC_SESSION_ENTITY_STATE * penState ) = 0;
      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall get_Session (
        /*[out,retval]*/ struct IUccSession * * ppSession ) = 0;
      virtual HRESULT __stdcall get_IsLocal (
        /*[out,retval]*/ VARIANT_BOOL * pbIsLocal ) = 0;
      virtual HRESULT __stdcall get_Endpoints (
        /*[out,retval]*/ struct IUccCollection * * ppEndpoints ) = 0;
      virtual HRESULT __stdcall CreateParticipantEndpoint (
        /*[in]*/ struct IUccUri * pUri,
        /*[in]*/ BSTR bstrEndpointId,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * ppParticipantEndpoint ) = 0;
      virtual HRESULT __stdcall CopyParticipantEndpoint (
        /*[in]*/ struct IUccSessionParticipantEndpoint * pInputParticipantEndpoint,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccSessionParticipantEndpoint * * ppParticipantEndpoint ) = 0;
      virtual HRESULT __stdcall AddParticipantEndpoint (
        /*[in]*/ struct IUccSessionParticipantEndpoint * pParticipantEndpoint,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall RemoveParticipantEndpoint (
        /*[in]*/ struct IUccSessionParticipantEndpoint * pParticipantEndpoint,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("233693fb-ad38-4152-8d4f-688fdcc3fcde"))
IUccSession : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum UCC_SESSION_TYPE * penType ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * ppContext ) = 0;
      virtual HRESULT __stdcall get_LocalParticipant (
        /*[out,retval]*/ struct IUccSessionParticipant * * ppParticipant ) = 0;
      virtual HRESULT __stdcall get_Participants (
        /*[out,retval]*/ struct IUccCollection * * ppParticipants ) = 0;
      virtual HRESULT __stdcall CreateParticipant (
        /*[in]*/ struct IUccUri * pUri,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccSessionParticipant * * ppParticipant ) = 0;
      virtual HRESULT __stdcall AddParticipant (
        /*[in]*/ struct IUccSessionParticipant * pParticipant,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall RemoveParticipant (
        /*[in]*/ struct IUccSessionParticipant * pParticipant,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall Terminate (
        /*[in]*/ enum UCC_REJECT_OR_TERMINATE_REASON enTerminateReason,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("2d237189-43ff-40ce-bdfb-1f560d56b4fa"))
IUccSessionCallControl : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall Alternate (
        /*[in]*/ struct IUccSession * pActiveSession,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall BlindTransfer (
        /*[in]*/ struct IUccUri * pUri,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall ConsultativeTransfer (
        /*[in]*/ struct IUccSession * pActiveSession,
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
};

struct __declspec(uuid("57061595-cea1-4c1d-af44-3dff48a7e0df"))
IUccSessionManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateSession (
        /*[in]*/ enum UCC_SESSION_TYPE enSessionType,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccSession * * ppSession ) = 0;
      virtual HRESULT __stdcall RegisterSessionDescriptionEvaluator (
        /*[in]*/ struct _IUccSessionDescriptionEvaluator * pSessionDescriptionEvaluator ) = 0;
      virtual HRESULT __stdcall UnregisterSessionDescriptionEvaluator ( ) = 0;
};

struct __declspec(uuid("005a2d10-dd52-4ace-9ca1-9bdd699315f0"))
IUccSessionParticipantCollectionEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Participant (
        /*[out,retval]*/ struct IUccSessionParticipant * * pVal ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * pVal ) = 0;
};

struct __declspec(uuid("37cf9d37-c22c-4c5d-9023-f78ae5325036"))
IUccIncomingSessionEvent : IUccIncomingRequestEvent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Session (
        /*[out,retval]*/ struct IUccSession * * ppIncomingSession ) = 0;
      virtual HRESULT __stdcall get_Inviter (
        /*[out,retval]*/ struct IUccSessionParticipant * * ppInvitingParticipant ) = 0;
      virtual HRESULT __stdcall get_History (
        /*[out,retval]*/ struct IUccCollection * * pVal ) = 0;
      virtual HRESULT __stdcall get_SourceType (
        /*[out,retval]*/ enum UCC_SOURCE_NETWORK_TYPE * penSourceType ) = 0;
      virtual HRESULT __stdcall Redirect (
        /*[in]*/ enum UCC_SESSION_REDIRECT_METHOD enRedirectMethod,
        /*[in]*/ struct IUccPropertyCollection * pUserNameUriPairCollection ) = 0;
      virtual HRESULT __stdcall get_TargetReplaceSession (
        /*[out,retval]*/ struct IUccSession * * ppTargetReplaceSession ) = 0;
};

struct __declspec(uuid("e11e841b-da1d-4425-bb62-44df4418cb24"))
IUccOutgoingSessionEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Session (
        /*[out,retval]*/ struct IUccSession * * pVal ) = 0;
};

struct __declspec(uuid("fb194526-2e78-4cca-927e-e967bf0101f7"))
IUccSubscriptionManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateSubscription (
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccSubscription * * ppSubscription ) = 0;
};

struct __declspec(uuid("f6e4f91e-40e2-4bf2-a808-37ec02daf67b"))
IUccSubscription : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreatePresentity (
        /*[in]*/ struct IUccUri * pPresentityUri,
        /*[in]*/ struct IUccContext * pContext,
        /*[out,retval]*/ struct IUccPresentity * * ppPresentity ) = 0;
      virtual HRESULT __stdcall AddPresentity (
        /*[in]*/ struct IUccPresentity * pPresentity ) = 0;
      virtual HRESULT __stdcall RemovePresentity (
        /*[in]*/ struct IUccPresentity * pPresentity ) = 0;
      virtual HRESULT __stdcall get_Presentities (
        /*[out,retval]*/ struct IUccCollection * * ppPresentityCollection ) = 0;
      virtual HRESULT __stdcall AddCategoryName (
        /*[in]*/ BSTR bstrName ) = 0;
      virtual HRESULT __stdcall RemoveCategoryName (
        /*[in]*/ BSTR bstrName ) = 0;
      virtual HRESULT __stdcall get_CategoryNames (
        /*[out,retval]*/ struct IUccCollection * * ppNameCollection ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * ppContext ) = 0;
      virtual HRESULT __stdcall Subscribe (
        /*[in]*/ struct IUccOperationContext * pContext ) = 0;
      virtual HRESULT __stdcall Unsubscribe (
        /*[in]*/ struct IUccOperationContext * pContext ) = 0;
      virtual HRESULT __stdcall Query (
        /*[in]*/ struct IUccOperationContext * pContext ) = 0;
};

struct __declspec(uuid("d0171dc1-7a36-494b-a8e5-cc51711a8cac"))
IUccPresentity : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_CategoryContexts (
        /*[out,retval]*/ struct IUccCollection * * ppCategoryCtxtCollection ) = 0;
      virtual HRESULT __stdcall TryFindCategoryContext (
        /*[in]*/ BSTR bstrCategoryName,
        /*[out]*/ struct IUccCategoryContext * * ppContext,
        /*[out,retval]*/ VARIANT_BOOL * pbFound ) = 0;
      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ struct IUccUri * * pUri ) = 0;
      virtual HRESULT __stdcall get_SourceType (
        /*[out,retval]*/ enum UCC_SOURCE_NETWORK_TYPE * penSourceType ) = 0;
      virtual HRESULT __stdcall get_IsSelf (
        /*[out,retval]*/ VARIANT_BOOL * pfIsSelf ) = 0;
      virtual HRESULT __stdcall get_SubscriberContext (
        /*[out,retval]*/ BSTR * pbstrSetSubscriberCtxt ) = 0;
      virtual HRESULT __stdcall put_SubscriberContext (
        /*[in]*/ BSTR pbstrSetSubscriberCtxt ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * ppContext ) = 0;
      virtual HRESULT __stdcall RefreshPresence (
        /*[in]*/ struct IUccOperationContext * pCtxt ) = 0;
};

struct __declspec(uuid("104cc179-5a41-4539-a57f-708bcbaaacd8"))
IUccSubscriptionEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetOperationInfo (
        /*[in]*/ struct IUccPresentity * pPresentity,
        /*[out,retval]*/ struct IUccOperationProgressEvent * * ppOperationData ) = 0;
      virtual HRESULT __stdcall get_Presentities (
        /*[out,retval]*/ struct IUccCollection * * ppPresentityCollection ) = 0;
      virtual HRESULT __stdcall get_OriginalOperationContext (
        /*[out,retval]*/ struct IUccOperationContext * * ppOperationContext ) = 0;
      virtual HRESULT __stdcall get_IsComplete (
        /*[out,retval]*/ VARIANT_BOOL * pfIsComplete ) = 0;
};

struct __declspec(uuid("69f9efcf-5fea-46b1-8395-bd96212a8c51"))
IUccCategoryContext : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Presentity (
        /*[out,retval]*/ struct IUccPresentity * * ppPresentity ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pbstrName ) = 0;
      virtual HRESULT __stdcall TryFindCategoryInstance (
        /*[in]*/ long lContainerID,
        /*[in]*/ long lInstanceId,
        /*[out]*/ struct IUccCategoryInstance * * ppCategoryInstance,
        /*[out,retval]*/ VARIANT_BOOL * pbFound ) = 0;
      virtual HRESULT __stdcall get_CategoryInstances (
        /*[in]*/ long lContainerID,
        /*[out,retval]*/ struct IUccCollection * * ppCategoryInstanceCollection ) = 0;
};

struct __declspec(uuid("f13220e8-dec0-4e04-ac27-3a7785e9d8c5"))
IUccCategoryContextEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_CategoryContext (
        /*[out,retval]*/ struct IUccCategoryContext * * ppVal ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * pVal ) = 0;
};

struct __declspec(uuid("72d6a6c3-ac51-4166-93f0-e941c3d784df"))
IUccCategoryInstance : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Presentity (
        /*[out,retval]*/ struct IUccPresentity * * ppPresentity ) = 0;
      virtual HRESULT __stdcall get_CategoryContext (
        /*[out,retval]*/ struct IUccCategoryContext * * pCategoryCtxt ) = 0;
      virtual HRESULT __stdcall get_InstanceId (
        /*[out,retval]*/ long * plInstanceId ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ BSTR * pbstrData ) = 0;
      virtual HRESULT __stdcall put_Value (
        /*[in]*/ BSTR pbstrData ) = 0;
      virtual HRESULT __stdcall get_Container (
        /*[out,retval]*/ struct IUccContainer * * ppContainer ) = 0;
      virtual HRESULT __stdcall get_ContainerId (
        /*[out,retval]*/ long * plContainerId ) = 0;
      virtual HRESULT __stdcall get_Version (
        /*[out,retval]*/ long * plVersion ) = 0;
      virtual HRESULT __stdcall get_ExpiryTime (
        /*[out,retval]*/ long * plExpiryTime ) = 0;
      virtual HRESULT __stdcall put_ExpiryTime (
        /*[in]*/ long plExpiryTime ) = 0;
      virtual HRESULT __stdcall get_ExpireType (
        /*[out,retval]*/ enum UCC_CATEGORY_INSTANCE_EXPIRE_TYPE * penExpireType ) = 0;
      virtual HRESULT __stdcall get_PublishedTime (
        /*[out,retval]*/ DATE * pPublishedTime ) = 0;
      virtual HRESULT __stdcall put_PublicationOperation (
        /*[in]*/ enum UCC_PUBLICATION_OPERATION_TYPE _arg1 ) = 0;
      virtual HRESULT __stdcall CreatePublishableCategoryInstance (
        /*[out,retval]*/ struct IUccCategoryInstance * * ppPublishableCategoryInstance ) = 0;
      virtual HRESULT __stdcall get_IsPublishedByLocalEndpoint (
        /*[out,retval]*/ VARIANT_BOOL * pfIsPublishedByLocalEndpoint ) = 0;
};

struct __declspec(uuid("6ed14bd9-627d-4d9d-9405-ed9ad0f07023"))
IUccCategoryInstanceEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_CategoryInstance (
        /*[out,retval]*/ struct IUccCategoryInstance * * ppVal ) = 0;
      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IUccContext * * pVal ) = 0;
};

struct __declspec(uuid("f1d73cad-6f6a-413b-bc2b-158dc2a7ccb2"))
IUccPublication : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall AddPublishableCategoryInstance (
        /*[in]*/ struct IUccCategoryInstance * pPublishableCategoryInstance ) = 0;
      virtual HRESULT __stdcall RemovePublishableCategoryInstance (
        /*[in]*/ struct IUccCategoryInstance * pPublishableCategoryInstance ) = 0;
      virtual HRESULT __stdcall Publish (
        /*[in]*/ struct IUccOperationContext * pOperationContext ) = 0;
      virtual HRESULT __stdcall get_PublishableCategoryInstances (
        /*[out,retval]*/ struct IUccCollection * * ppPublishableCategoryInstanceCollection ) = 0;
};

struct __declspec(uuid("646d6d11-c95c-49b2-a60f-77c5e91408d9"))
IUccContainer : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ long * plContainerId ) = 0;
      virtual HRESULT __stdcall AddMember (
        /*[in]*/ struct IUccContainerMember * pMember ) = 0;
      virtual HRESULT __stdcall RemoveMember (
        /*[in]*/ struct IUccContainerMember * pMember ) = 0;
      virtual HRESULT __stdcall GetMembersWithScope (
        /*[in]*/ enum UCC_CONTAINER_MEMBERSHIP_SCOPE enScope,
        /*[out,retval]*/ struct IUccCollection * * ppMemberCollection ) = 0;
      virtual HRESULT __stdcall get_Members (
        /*[out,retval]*/ struct IUccCollection * * ppMemberCollection ) = 0;
      virtual HRESULT __stdcall GetCategoryInstances (
        /*[in]*/ BSTR bstrCategoryName,
        /*[out,retval]*/ struct IUccCollection * * ppCategoryInstanceCollection ) = 0;
};

struct __declspec(uuid("0dd847ab-0f32-4449-9165-7425d11ddd89"))
IUccContainerMember : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * pbstrURI ) = 0;
      virtual HRESULT __stdcall get_Scope (
        /*[out,retval]*/ enum UCC_CONTAINER_MEMBERSHIP_SCOPE * penScope ) = 0;
      virtual HRESULT __stdcall get_Containers (
        /*[out,retval]*/ struct IUccCollection * * ppContainerCollection ) = 0;
      virtual HRESULT __stdcall IsInContainer (
        /*[in]*/ long lContainerID,
        /*[out,retval]*/ VARIANT_BOOL * pfIsInContainer ) = 0;
      virtual HRESULT __stdcall TryGetContainerWithCategory (
        /*[in]*/ BSTR bstrCategoryName,
        /*[out]*/ struct IUccContainer * * ppContainer,
        /*[out,retval]*/ VARIANT_BOOL * pbFound ) = 0;
      virtual HRESULT __stdcall CompareTo (
        /*[in]*/ struct IUccContainerMember * pOtherMember,
        /*[out,retval]*/ enum UCC_CONTAINER_MEMBER_COMPARE_RESULT * penResult ) = 0;
};

struct __declspec(uuid("8698350f-4b7e-4e38-a7e7-83038b69db7b"))
IUccPublicationManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateContainer (
        /*[in]*/ long lContainerID,
        /*[out,retval]*/ struct IUccContainer * * pContainer ) = 0;
      virtual HRESULT __stdcall CreatePublishableCategoryInstance (
        /*[in]*/ BSTR bstrCategoryName,
        /*[in]*/ long lContainerID,
        /*[in]*/ long lInstance,
        /*[in]*/ enum UCC_CATEGORY_INSTANCE_EXPIRE_TYPE enExpireType,
        /*[in]*/ long lExpireTime,
        /*[out,retval]*/ struct IUccCategoryInstance * * ppCategoryInstance ) = 0;
      virtual HRESULT __stdcall CreateContainerMember (
        /*[in]*/ BSTR bstrUri,
        /*[in]*/ enum UCC_CONTAINER_MEMBERSHIP_SCOPE enScope,
        /*[out,retval]*/ struct IUccContainerMember * * ppMember ) = 0;
      virtual HRESULT __stdcall CreatePublication (
        /*[out,retval]*/ struct IUccPublication * * ppPublication ) = 0;
      virtual HRESULT __stdcall GetPresenceCategoryName (
        /*[in]*/ enum UCC_PRESENCE_CATEGORY_TYPE enType,
        /*[out,retval]*/ BSTR * pbstrName ) = 0;
};

struct __declspec(uuid("679d5f13-90bb-40ab-98f7-6243b25b8d89"))
IUccContainerMemberCollectionEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ContainerMember (
        /*[out,retval]*/ struct IUccContainerMember * * ppVal ) = 0;
};

struct __declspec(uuid("e7122690-cbf5-4d7d-9f1b-bf28668d91a9"))
IUccContainerMembershipEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Container (
        /*[out,retval]*/ struct IUccContainer * * ppVal ) = 0;
};

struct __declspec(uuid("92700fb4-ff9b-46b7-af97-789d3d0e4ce3"))
IUccContainerMembershipManager : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall TryFindContainerMember (
        /*[in]*/ BSTR bstrUri,
        /*[in]*/ enum UCC_CONTAINER_MEMBERSHIP_SCOPE enScope,
        /*[in]*/ enum UCC_CONTAINER_MEMBER_MATCHING_TYPE enMatchingType,
        /*[out]*/ struct IUccContainerMember * * ppContainerMember,
        /*[out,retval]*/ VARIANT_BOOL * pbFound ) = 0;
      virtual HRESULT __stdcall GetCategoryInstancesForUri (
        /*[in]*/ struct IUccUri * pUri,
        /*[in]*/ BSTR bstrCategoryName,
        /*[in]*/ enum UCC_CONTAINER_MEMBERSHIP_SCOPE enBroaderScope,
        /*[out,retval]*/ struct IUccCollection * * ppCategoryInstanceCollection ) = 0;
};

struct __declspec(uuid("70c8f58f-a1ca-4614-8760-68399e111cfa"))
IUccContact : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ struct IUccUri * * ppUri ) = 0;
      virtual HRESULT __stdcall put_Uri (
        /*[in]*/ struct IUccUri * ppUri ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pbstrName ) = 0;
      virtual HRESULT __stdcall put_Name (
        /*[in]*/ BSTR pbstrName ) = 0;
      virtual HRESULT __stdcall get_ExternalUri (
        /*[out,retval]*/ BSTR * pbstrExternalUri ) = 0;
      virtual HRESULT __stdcall put_ExternalUri (
        /*[in]*/ BSTR pbstrExternalUri ) = 0;
      virtual HRESULT __stdcall get_Groups (
        /*[out,retval]*/ struct IUccCollection * * ppGroups ) = 0;
      virtual HRESULT __stdcall AddToGroup (
        /*[in]*/ struct IUccGroup * pGroup ) = 0;
      virtual HRESULT __stdcall RemoveFromGroup (
        /*[in]*/ struct IUccGroup * pGroup ) = 0;
      virtual HRESULT __stdcall get_ContactExtension (
        /*[out,retval]*/ BSTR * pbstrContactExtension ) = 0;
      virtual HRESULT __stdcall put_ContactExtension (
        /*[in]*/ BSTR pbstrContactExtension ) = 0;
};

struct __declspec(uuid("f36d99bb-5cb1-4dd6-9677-11787069e405"))
IUccGroup : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * pbstrName ) = 0;
      virtual HRESULT __stdcall put_Name (
        /*[in]*/ BSTR pbstrName ) = 0;
      virtual HRESULT __stdcall get_ExternalUri (
        /*[out,retval]*/ BSTR * pbstrExternalUri ) = 0;
      virtual HRESULT __stdcall put_ExternalUri (
        /*[in]*/ BSTR pbstrExternalUri ) = 0;
      virtual HRESULT __stdcall AddContact (
        /*[in]*/ struct IUccContact * pContact ) = 0;
      virtual HRESULT __stdcall RemoveContact (
        /*[in]*/ struct IUccContact * pContact ) = 0;
      virtual HRESULT __stdcall get_Contacts (
        /*[out,retval]*/ struct IUccCollection * * ppContacts ) = 0;
      virtual HRESULT __stdcall get_GroupType (
        /*[out,retval]*/ enum UCC_GROUP_TYPE * penGroupType ) = 0;
      virtual HRESULT __stdcall put_GroupType (
        /*[in]*/ enum UCC_GROUP_TYPE penGroupType ) = 0;
      virtual HRESULT __stdcall get_PrivateId (
        /*[out,retval]*/ BSTR * pbstrPrivateID ) = 0;
      virtual HRESULT __stdcall put_PrivateId (
        /*[in]*/ BSTR pbstrPrivateID ) = 0;
      virtual HRESULT __stdcall get_GroupExtension (
        /*[out,retval]*/ BSTR * pbstrGroupExtension ) = 0;
      virtual HRESULT __stdcall put_GroupExtension (
        /*[in]*/ BSTR pbstrGroupExtension ) = 0;
};

struct __declspec(uuid("0994e0c2-b0ac-46b5-b258-7df51fab8b2e"))
IUccContactCollectionEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Contact (
        /*[out,retval]*/ struct IUccContact * * ppVal ) = 0;
};

struct __declspec(uuid("7377729f-2fa4-4e73-a6a9-cdb8b035676d"))
IUccGroupMembershipEvent : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Group (
        /*[out,retval]*/ struct IUccGroup * * ppVal ) = 0;
};

//
// Named GUID constants initializations
//

extern "C" const GUID __declspec(selectany) LIBID_UCCAPILib =
    {0xecd5307e,0x4419,0x43cf,{0x8b,0xda,0xc9,0x94,0x6a,0xc3,0x75,0xcf}};
extern "C" const GUID __declspec(selectany) IID_IUccCollection =
    {0x0814cabe,0x959d,0x47b4,{0x84,0x1e,0x5b,0x77,0xb0,0x6d,0x60,0x46}};
extern "C" const GUID __declspec(selectany) IID_IUccProperty =
    {0xcadc8c3e,0x857e,0x4898,{0x90,0x83,0x87,0x50,0x55,0x93,0x76,0xaf}};
extern "C" const GUID __declspec(selectany) IID_IUccReadOnlyPropertyCollection =
    {0xe97197bf,0x9939,0x4ab5,{0x90,0x56,0x5a,0x1f,0xdb,0xd1,0x53,0x90}};
extern "C" const GUID __declspec(selectany) IID_IUccPropertyCollection =
    {0x73fe78fc,0xefe6,0x422b,{0xb9,0xdd,0x4a,0x33,0xd9,0xed,0x63,0x95}};
extern "C" const GUID __declspec(selectany) IID_IUccContext =
    {0xe66a1e6b,0x5c1c,0x40b6,{0xa3,0xad,0x23,0xc3,0x69,0x9b,0x96,0x58}};
extern "C" const GUID __declspec(selectany) IID_IUccOperationContext =
    {0xafe09299,0x87d8,0x4737,{0xad,0x38,0xfa,0x6d,0x78,0x53,0x7f,0x4a}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSessionParticipant =
    {0xe8909d14,0x268c,0x4cb9,{0xa9,0xbf,0x48,0x81,0x8f,0x5c,0x48,0xbc}};
extern "C" const GUID __declspec(selectany) IID_IUccArray =
    {0x307c5595,0xc51d,0x414b,{0xbc,0xcf,0x38,0xc0,0xcd,0x22,0x7f,0x52}};
extern "C" const GUID __declspec(selectany) IID_IUccAudioMediaChannel =
    {0x5c3ed734,0x2a86,0x4b9a,{0x8a,0x03,0x83,0x7d,0x43,0x42,0xf5,0xd3}};
extern "C" const GUID __declspec(selectany) IID_IUccAudioMediaDevice =
    {0xa98332d2,0x2d80,0x4759,{0xb8,0xa8,0xcc,0x5f,0x4c,0x79,0xf4,0x60}};
extern "C" const GUID __declspec(selectany) IID_IUccAudioVideoSession =
    {0xd0283fc6,0x4ab7,0x4d63,{0xbb,0x24,0x8c,0xdd,0x72,0xfd,0x89,0x7a}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaChannel =
    {0xfaff2612,0x3bf1,0x44aa,{0xac,0xb6,0x80,0xd3,0x9c,0x0d,0x99,0x1d}};
extern "C" const GUID __declspec(selectany) IID_IUccAudioVideoSessionParticipant =
    {0xc3c42eac,0x9d5c,0x4daa,{0xab,0xcc,0x84,0x01,0x9d,0x37,0x06,0xad}};
extern "C" const GUID __declspec(selectany) IID_IUccUri =
    {0x8b55d477,0xd96e,0x4670,{0xb0,0xa9,0xf1,0x65,0x3b,0x9d,0x97,0xce}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceEntityView =
    {0xc2a50530,0x5635,0x41b5,{0x8d,0x30,0xd4,0x06,0x48,0x92,0x38,0xfa}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceInformation =
    {0x65d83c13,0xb0a2,0x400c,{0xbf,0xd8,0xbc,0xc4,0x89,0x0b,0x75,0x74}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceManagerSession =
    {0x052e68e9,0x3ea7,0x4b9c,{0x97,0x2c,0x91,0x01,0x3f,0x04,0x4a,0xf7}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceManager =
    {0x9dcfc8ab,0x4b32,0x4078,{0x84,0x14,0xa5,0x9d,0xe5,0x29,0xb3,0x27}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceView =
    {0xa8c81cfe,0x853d,0x413a,{0xab,0x91,0xd5,0xb1,0xc6,0xee,0x10,0x99}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceSession =
    {0xf8c1eee0,0x4272,0x4dd6,{0x89,0x86,0x78,0x63,0x58,0xe6,0xa0,0xe3}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceSessionParticipantEndpoint =
    {0x48461087,0x92f1,0x43bc,{0x9d,0x42,0xae,0x61,0x5f,0xc3,0x97,0x71}};
extern "C" const GUID __declspec(selectany) IID_IUccCredential =
    {0xbfc14f2e,0x90fe,0x4a5e,{0x89,0x4d,0x62,0xa1,0x64,0x35,0xb4,0xd8}};
extern "C" const GUID __declspec(selectany) IID_IUccCredentialCache =
    {0xab9f6a4a,0xa930,0x4b5e,{0xbb,0x77,0xea,0x6e,0xc3,0x13,0x46,0xb3}};
extern "C" const GUID __declspec(selectany) IID_IUccDelegate =
    {0xce70dd5e,0xc89e,0x4863,{0x8a,0x59,0x39,0x22,0x17,0x73,0xc9,0xe9}};
extern "C" const GUID __declspec(selectany) IID_IUccDelegatee =
    {0xd3e9e4fa,0xe813,0x431e,{0xaa,0x53,0xb7,0x3c,0xd7,0x79,0x98,0x43}};
extern "C" const GUID __declspec(selectany) IID_IUccDelegateEvent =
    {0xfab74d22,0x5830,0x4d14,{0xa9,0x18,0x17,0x7f,0x77,0xea,0x02,0x56}};
extern "C" const GUID __declspec(selectany) IID_IUccDiagnosticInfo =
    {0x7bcb8062,0x1931,0x42ca,{0x8c,0xc6,0x38,0xea,0x95,0xd1,0x94,0x1f}};
extern "C" const GUID __declspec(selectany) IID_IUccDiagnosticReportingSettings =
    {0xce5f4b42,0x3d3f,0x4466,{0xa3,0xd5,0xb1,0x1e,0x58,0x8f,0xad,0xe8}};
extern "C" const GUID __declspec(selectany) IID_IUccEndpoint =
    {0x2dc3c92f,0x8fe2,0x4a84,{0x8f,0x76,0xbb,0x1a,0x14,0xff,0x0f,0x3a}};
extern "C" const GUID __declspec(selectany) IID_IUccEndpoint2 =
    {0x943bca5a,0xe22d,0x4c5d,{0x88,0xd0,0xd0,0x02,0x1e,0xe5,0x10,0xd7}};
extern "C" const GUID __declspec(selectany) IID_IUccOperationProgressEvent =
    {0x6cfb1827,0x9feb,0x462a,{0x97,0xc6,0xb8,0x58,0x70,0x2c,0x4f,0xe4}};
extern "C" const GUID __declspec(selectany) IID_IUccFindMediaConnectivityServersEvent =
    {0x91154bc1,0x8079,0x4bf7,{0x89,0x14,0xfc,0x5b,0x62,0xaa,0xc3,0xd6}};
extern "C" const GUID __declspec(selectany) IID_IUccFindServerEvent =
    {0x162809da,0x825d,0x4e75,{0x98,0x77,0xa8,0x90,0x98,0x97,0xe5,0x67}};
extern "C" const GUID __declspec(selectany) IID_IUccHistoryInfoEntry =
    {0x36e59b05,0x2fc3,0x4782,{0xb3,0xc9,0x7d,0x5d,0xdd,0xbe,0x25,0xb4}};
extern "C" const GUID __declspec(selectany) IID_IUccInstantMessagingSession =
    {0x2f96ff5d,0xe7d3,0x4aef,{0xa9,0x79,0x4a,0x02,0xd6,0x9b,0x9d,0xd0}};
extern "C" const GUID __declspec(selectany) IID_IUccInstantMessagingSessionParticipant =
    {0xe5e2d6bc,0x910d,0x4348,{0x9d,0x6b,0x23,0xec,0x74,0x90,0x8e,0x68}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaConnectivityServer =
    {0x780eca62,0xe839,0x4bb0,{0x84,0xc1,0x00,0x5f,0x5f,0x4b,0xc8,0x55}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaConnectivityServerConfiguration =
    {0x6a84605f,0xaf6a,0x402b,{0x9b,0x0d,0x5f,0xf8,0x81,0xf8,0x18,0x79}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaDevice =
    {0x216d2747,0x60a0,0x43be,{0xab,0xb3,0x19,0x93,0xd2,0x63,0xaa,0x68}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaChannelDevices =
    {0x8d538142,0x1fc9,0x4cd4,{0xb5,0x3b,0x56,0x01,0x36,0x0a,0x90,0x8e}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaDeviceManager =
    {0x24b14557,0x850e,0x494f,{0x9c,0xa7,0x79,0x95,0x59,0xcd,0x3e,0xeb}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaDeviceSettings =
    {0x56a60090,0x39c2,0x48fb,{0xad,0x07,0x43,0xc8,0x47,0xc5,0x5c,0x4c}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaEndpointSettings =
    {0x4d587df7,0x1adb,0x48cb,{0xb6,0xb4,0x91,0x4a,0x85,0xc2,0x6f,0x6a}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaPlayerDevice =
    {0x2df08b07,0x6f9c,0x4890,{0xa2,0xb0,0xdd,0xa6,0x8f,0x43,0x0b,0xfb}};
extern "C" const GUID __declspec(selectany) IID_IUccPlatform =
    {0x8ab54fb5,0x7656,0x4348,{0xb5,0xc5,0x32,0xca,0x2e,0xfe,0xe7,0xf4}};
extern "C" const GUID __declspec(selectany) IID_IUccProvisioningPolicyInstance =
    {0x8e17e673,0x96ec,0x4aea,{0xaa,0xec,0x52,0x80,0xed,0x16,0x24,0x10}};
extern "C" const GUID __declspec(selectany) IID_IUccProxyEndpoint =
    {0xb22ef3c5,0xc0dc,0x473b,{0xb2,0xa4,0x5e,0xe2,0x53,0xe8,0x7f,0x5e}};
extern "C" const GUID __declspec(selectany) IID_IUccEndpointRefreshEvent =
    {0x5e4739c8,0x240d,0x4678,{0x95,0x74,0x9b,0x34,0x2f,0x8e,0x60,0x94}};
extern "C" const GUID __declspec(selectany) IID_IUccSignalingServer =
    {0x20f08a08,0xe163,0x45f8,{0xb4,0x3a,0x2d,0xab,0x57,0x9c,0x51,0xae}};
extern "C" const GUID __declspec(selectany) IID_IUccServerSignalingSettings =
    {0x3405d52a,0x1498,0x4bd5,{0x92,0x48,0xe4,0x23,0xe9,0x7b,0xf1,0xfe}};
extern "C" const GUID __declspec(selectany) IID__IUccSessionDescriptionEvaluator =
    {0x83a209ca,0x296c,0x44d9,{0x93,0x31,0xf5,0xdc,0xf3,0xe3,0xc0,0xab}};
extern "C" const GUID __declspec(selectany) IID_IUccTraceSettings =
    {0x8e0d1060,0xdf8e,0x45f9,{0xa8,0x34,0x23,0x7f,0x1e,0x44,0x18,0xf0}};
//extern "C" const GUID __declspec(selectany) IID_IVideoWindow =
//    {0x56a868b4,0x0ad4,0x11ce,{0xb0,0x3a,0x00,0x20,0xaf,0x0b,0xa7,0x70}};
extern "C" const GUID __declspec(selectany) IID_IUccVideoMediaChannel =
    {0xbb782727,0x4c0a,0x42ba,{0xa8,0xf8,0xcb,0x6b,0xd7,0x14,0x0b,0xab}};
extern "C" const GUID __declspec(selectany) IID_IUccVideoMediaDevice =
    {0x230175b3,0x791c,0x4904,{0xbc,0xe9,0x80,0x48,0xcd,0xa8,0x5f,0x87}};
extern "C" const GUID __declspec(selectany) IID_IUccAudioVideoMediaChannelStreamsEvent =
    {0x90a4b7cd,0x1dac,0x498f,{0xa1,0x2b,0xfc,0xe9,0x49,0x55,0x06,0x94}};
extern "C" const GUID __declspec(selectany) IID_IUccAudioVideoMediaChannel =
    {0x63bbfd31,0xc323,0x4a38,{0xbe,0x5a,0xe9,0xaf,0x55,0xd5,0xe6,0xa7}};
extern "C" const GUID __declspec(selectany) IID_IUccUserSearchQuery =
    {0x5072148c,0xde7a,0x4826,{0x96,0x5c,0x81,0x2a,0xb6,0x76,0xe0,0xa4}};
extern "C" const GUID __declspec(selectany) IID_IUccUserSearchManager =
    {0x345420c6,0x72b7,0x4640,{0xa4,0x74,0x9e,0xb8,0x77,0xfc,0xf6,0x0a}};
extern "C" const GUID __declspec(selectany) IID_IUccUserSearchQueryEvent =
    {0x54562fbc,0x5a84,0x4461,{0x8b,0xc9,0x59,0x07,0x37,0xe5,0xde,0x13}};
extern "C" const GUID __declspec(selectany) IID_IUccUserSearchResultRecord =
    {0xac628a4c,0x7595,0x4129,{0x83,0xf8,0xb3,0x7c,0x2f,0x8a,0xf8,0x13}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceMediaChannelStateChangedEvent =
    {0x8f77e9f1,0xebe2,0x496d,{0x96,0xfd,0xa7,0xdd,0xd1,0xad,0xb3,0x29}};
extern "C" const GUID __declspec(selectany) DIID__IUccApplicationSessionParticipantEvents =
    {0xdb49e9b0,0x4c96,0x46b6,{0x82,0x27,0x15,0x94,0x81,0x61,0x0c,0x6b}};
extern "C" const GUID __declspec(selectany) IID_IUccOutgoingInvitationEvent =
    {0x6d362a12,0xa539,0x4581,{0xa2,0xec,0x26,0x47,0xca,0xff,0x6a,0x72}};
extern "C" const GUID __declspec(selectany) IID_IUccIncomingRequestEvent =
    {0xceb0d9cb,0x78b9,0x4d19,{0xb8,0xd9,0x48,0x0a,0x1b,0xbf,0x9c,0x38}};
extern "C" const GUID __declspec(selectany) IID_IUccIncomingInvitationEvent =
    {0x4896d466,0x934f,0x43a2,{0xbf,0x32,0xe6,0xcc,0xc6,0x72,0x15,0x24}};
extern "C" const GUID __declspec(selectany) IID_IUccInvitationAcceptedEvent =
    {0x06eefc22,0x08be,0x4ea4,{0x87,0x9e,0x11,0x59,0x39,0xe5,0x9e,0x96}};
extern "C" const GUID __declspec(selectany) DIID__IUccAudioVideoSessionEvents =
    {0xb6dc5e1a,0x18ca,0x47bb,{0xbc,0xe7,0x15,0x61,0x42,0xac,0x54,0x53}};
extern "C" const GUID __declspec(selectany) DIID__IUccAudioVideoSessionParticipantEvents =
    {0x80bdd0ab,0x207b,0x4180,{0x85,0x11,0xc9,0x96,0x7b,0x0b,0xf6,0x98}};
extern "C" const GUID __declspec(selectany) IID_IUccIncomingMediaRequestEvent =
    {0x46ace669,0x1ede,0x478d,{0x98,0xcd,0xf1,0xd1,0xdd,0xda,0x11,0x6f}};
extern "C" const GUID __declspec(selectany) DIID__IUccAudioMediaChannelEvents =
    {0xda1582aa,0x99e3,0x424b,{0x87,0xdf,0xd3,0xd1,0x5a,0x27,0x8e,0x75}};
extern "C" const GUID __declspec(selectany) IID_IUccAudioSignalLevelChangedEvent =
    {0xd82b8d0c,0x1e1d,0x4f13,{0x8a,0xe1,0xdd,0x54,0xf7,0x0f,0xfb,0x09}};
extern "C" const GUID __declspec(selectany) DIID__IUccMediaChannelDevicesEvents =
    {0xea82efbc,0x1da0,0x45cb,{0xb4,0xc2,0xd0,0x33,0xb5,0xd1,0x95,0x59}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaDeviceChangeEvent =
    {0xd8ba144f,0xefe7,0x4eb7,{0x85,0x7c,0x53,0x77,0x1f,0xc7,0x98,0x4e}};
extern "C" const GUID __declspec(selectany) DIID__IUccMediaChannelEvents =
    {0xd7db1218,0xa366,0x4482,{0xaf,0x15,0x41,0xd5,0xbc,0x85,0x9d,0x94}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaChannelEvent =
    {0x14689259,0x2a48,0x4eb3,{0x90,0x94,0xec,0x65,0x01,0x9b,0x40,0x76}};
extern "C" const GUID __declspec(selectany) DIID__IUccMediaEndpointEvents =
    {0xc7eddd72,0x0fd5,0x422e,{0xa2,0x06,0xcf,0x19,0xb0,0x9d,0x0e,0x40}};
extern "C" const GUID __declspec(selectany) DIID__IUccAudioVideoMediaChannelEvents =
    {0xbbfcb8d3,0x7be1,0x4525,{0x91,0x6b,0x54,0x8a,0x6a,0x61,0x7e,0xe4}};
extern "C" const GUID __declspec(selectany) DIID__IUccConferenceManagerSessionEvents =
    {0xf891ace7,0xe9fc,0x4768,{0x8a,0x50,0x61,0x48,0x90,0x20,0xbc,0x8d}};
extern "C" const GUID __declspec(selectany) DIID__IUccConferenceMediaChannelEvents =
    {0x7dbbf4e6,0xa70d,0x4889,{0x86,0xa0,0x2d,0x95,0xd3,0xef,0x04,0x7b}};
extern "C" const GUID __declspec(selectany) IID_IUccPropertyUpdateEvent =
    {0x826ef7c1,0x84b3,0x486a,{0x97,0x90,0x14,0x43,0x0f,0x64,0x2d,0x4f}};
extern "C" const GUID __declspec(selectany) DIID__IUccConferenceMediaChannelEvents2 =
    {0x386ab1af,0xdd8b,0x41e9,{0xad,0x87,0x8c,0x09,0x58,0x3e,0x80,0x5e}};
extern "C" const GUID __declspec(selectany) IID_IUccNamedPropertyUpdateEvent =
    {0x07b3bf54,0xc27f,0x4413,{0xb6,0xbf,0x84,0x1c,0xef,0x84,0xe5,0xba}};
extern "C" const GUID __declspec(selectany) DIID__IUccConferenceSessionEvents =
    {0x8ec5f518,0x6c4d,0x4c79,{0xaf,0xa6,0x47,0xb6,0xd1,0x03,0x9b,0xd0}};
extern "C" const GUID __declspec(selectany) DIID__IUccConferenceSessionEvents2 =
    {0x390a56e6,0x5868,0x433c,{0x8c,0xf3,0x40,0x1d,0x10,0x4c,0x59,0xf3}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceSession2 =
    {0x25c44946,0x2e3d,0x4bf3,{0xa5,0xbe,0x95,0x13,0xe5,0xf9,0x69,0xe9}};
extern "C" const GUID __declspec(selectany) DIID__IUccConferenceSessionParticipantEvents =
    {0xc819c292,0xf45f,0x4854,{0xae,0x66,0xc9,0xe3,0xcf,0xe4,0x51,0x95}};
extern "C" const GUID __declspec(selectany) DIID__IUccConferenceSessionParticipantEndpointEvents =
    {0xc57a388e,0xc215,0x4bd4,{0x8b,0x2d,0xda,0x5f,0x34,0x1e,0xf9,0x6b}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceSessionParticipantEndpointStateChangedEvent =
    {0xa68a5d90,0x944e,0x40c0,{0x91,0x46,0x88,0x3f,0x9c,0x85,0x56,0xbd}};
extern "C" const GUID __declspec(selectany) DIID__IUccConferenceMediaChannelCollectionEvents =
    {0x75025986,0xcdaa,0x42e7,{0x9e,0x06,0x64,0x2a,0x8e,0x9e,0x42,0x53}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaChannelCollectionEvent =
    {0x56a2af6c,0xe402,0x4fc9,{0x94,0x61,0xd8,0xa9,0x2a,0xbe,0x19,0x18}};
extern "C" const GUID __declspec(selectany) DIID__IUccConferenceEntityViewCollectionEvents =
    {0x53548bb8,0x662d,0x4117,{0x9f,0x82,0x26,0xb8,0x5a,0x30,0x74,0x2a}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceEntityViewCollectionEvent =
    {0xbe9e679a,0xd5af,0x485e,{0xb1,0xf6,0x8c,0x74,0x37,0x8a,0x53,0x59}};
extern "C" const GUID __declspec(selectany) DIID__IUccConferenceEntityViewEvents =
    {0x67e8922c,0x51d8,0x409a,{0x92,0x19,0x30,0x64,0x6c,0x50,0xba,0xea}};
extern "C" const GUID __declspec(selectany) DIID__IUccEndpointEvents =
    {0x060bd8f0,0x5a2a,0x4d56,{0xa0,0x65,0x21,0x73,0x89,0x4a,0xef,0xaf}};
extern "C" const GUID __declspec(selectany) DIID__IUccEndpointEvents2 =
    {0x1a8a9e13,0x99e3,0x4edf,{0xa2,0xf6,0x9d,0xdb,0xb0,0x41,0x02,0x3c}};
extern "C" const GUID __declspec(selectany) DIID__IUccInstantMessagingSessionEvents =
    {0xfdcb6c19,0x66cb,0x492a,{0x9f,0x4e,0x99,0x6d,0x92,0xa5,0x04,0x07}};
extern "C" const GUID __declspec(selectany) IID_IUccSessionOperationEvent =
    {0x70e0fff6,0xdd6d,0x4420,{0x8e,0x19,0x2d,0xe9,0x1d,0xa9,0x91,0x95}};
extern "C" const GUID __declspec(selectany) DIID__IUccInstantMessagingSessionParticipantEvents =
    {0x98a51831,0x7b5b,0x49ba,{0x84,0xcf,0x2a,0x4c,0xde,0x61,0xfa,0x35}};
extern "C" const GUID __declspec(selectany) DIID__IUccPlatformEvents =
    {0xa1faa7b6,0x0e6e,0x4d7e,{0x9e,0x4d,0x62,0xc1,0x62,0x92,0x92,0xb8}};
extern "C" const GUID __declspec(selectany) DIID__IUccProxyEndpointCollectionEvents =
    {0xe7a261a2,0xf96b,0x4ce9,{0xa0,0xcf,0x66,0x0a,0x4b,0x64,0xfa,0x70}};
extern "C" const GUID __declspec(selectany) IID_IUccProxyEndpointCollectionEvent =
    {0x03988bae,0x597c,0x402e,{0x8b,0x7d,0x2c,0xd4,0x1a,0xa2,0x5e,0xcb}};
extern "C" const GUID __declspec(selectany) DIID__IUccServerSignalingSettingsEvents =
    {0xacee6559,0x9e58,0x4c91,{0xbc,0x24,0xee,0xa3,0xf8,0xfc,0xe5,0xd6}};
extern "C" const GUID __declspec(selectany) DIID__IUccSessionEvents =
    {0xab81ffb3,0x80a5,0x4fe9,{0xac,0x37,0xb0,0x6d,0x5d,0x63,0x1b,0xe1}};
extern "C" const GUID __declspec(selectany) DIID__IUccSessionParticipantCollectionEvents =
    {0x3130f1a7,0x0569,0x4b49,{0xa8,0xc4,0x6b,0x90,0x21,0x94,0xe5,0x5e}};
extern "C" const GUID __declspec(selectany) DIID__IUccSessionManagerEvents =
    {0x90bbe2b8,0xfb81,0x4e70,{0xae,0x47,0x92,0xb7,0x54,0x31,0xd4,0xc1}};
extern "C" const GUID __declspec(selectany) DIID__IUccSessionParticipantEvents =
    {0x9b48f770,0x4eac,0x4967,{0xbc,0xac,0x28,0x3b,0x73,0x55,0x78,0x60}};
extern "C" const GUID __declspec(selectany) IID_IUccSessionParticipantStateChangedEvent =
    {0xd0e8a9a5,0x2f04,0x42f7,{0x96,0xea,0x89,0xab,0x2b,0x54,0x7d,0xa8}};
extern "C" const GUID __declspec(selectany) DIID__IUccSessionParticipantEndpointCollectionEvents =
    {0x8db557ba,0xe669,0x4de9,{0x99,0x22,0x8c,0x6c,0x53,0xad,0x1e,0xa3}};
extern "C" const GUID __declspec(selectany) DIID__IUccMediaChannelCollectionEvents =
    {0x00fb5f00,0xb2d5,0x4f3d,{0x9e,0x4b,0xfe,0xb4,0x18,0x0d,0x66,0xfa}};
extern "C" const GUID __declspec(selectany) DIID__IUccSessionCallControlEvents =
    {0x79712662,0x75f4,0x44a0,{0x85,0x19,0x1e,0x60,0x13,0xde,0x41,0x09}};
extern "C" const GUID __declspec(selectany) IID_IUccSessionTransferProgressEvent =
    {0xc31836bd,0xad83,0x4a63,{0xb5,0x3a,0xac,0x61,0x61,0xce,0xa9,0x7f}};
extern "C" const GUID __declspec(selectany) IID_IUccIncomingSessionTransferEvent =
    {0x7d3446dd,0xa657,0x437a,{0xb4,0x7f,0x30,0x80,0xcb,0xb0,0x2c,0xa3}};
extern "C" const GUID __declspec(selectany) DIID__IUccSignalingChannelEvents =
    {0x6a8e8235,0x9734,0x4eb0,{0x93,0x8c,0xa6,0xde,0x10,0xde,0xb4,0xea}};
extern "C" const GUID __declspec(selectany) IID_IUccSignalingHeaderFieldParameter =
    {0x5a9747c4,0x41ed,0x4812,{0x8e,0x73,0x8d,0xff,0xd9,0xf5,0xc9,0x7b}};
extern "C" const GUID __declspec(selectany) IID_IUccSignalingHeaderField =
    {0x05bf6f94,0xbb11,0x4f38,{0xa4,0x5a,0x1e,0x1c,0xbe,0x17,0x7d,0xbf}};
extern "C" const GUID __declspec(selectany) IID_IUccSignalingMessage =
    {0x9d45f185,0x46b0,0x48b4,{0x9e,0x9c,0x08,0x03,0x13,0x28,0x07,0x11}};
extern "C" const GUID __declspec(selectany) IID_IUccSignalingRequest =
    {0x9ea6b70c,0xc59e,0x4823,{0x9f,0xdc,0xe7,0xc8,0x5e,0x63,0x20,0xa0}};
extern "C" const GUID __declspec(selectany) IID_IUccSignalingChannel =
    {0x23053341,0x87a9,0x44eb,{0xbb,0x22,0x80,0x08,0x35,0xbb,0xa6,0x20}};
extern "C" const GUID __declspec(selectany) IID_IUccIncomingSignalingMessageEvent =
    {0xbc23ff63,0xd024,0x46df,{0xb9,0x80,0xf8,0xd7,0xd7,0x6c,0xd6,0xa9}};
extern "C" const GUID __declspec(selectany) DIID__IUccUserSearchQueryEvents =
    {0x94f59d79,0x583a,0x4547,{0xa6,0x20,0xea,0xd9,0x32,0xa2,0xf2,0xeb}};
extern "C" const GUID __declspec(selectany) IID_IUccOperationManager =
    {0x1713afd4,0x3ba6,0x48bc,{0x85,0x3b,0xe1,0xb7,0xe1,0x42,0x5e,0xc8}};
extern "C" const GUID __declspec(selectany) IID_IUccCategoryInstanceTransformationManifest =
    {0x84b2232b,0x9ac5,0x4a70,{0xbe,0xe3,0x9a,0xfb,0x52,0xb3,0xad,0x25}};
extern "C" const GUID __declspec(selectany) IID_IUccCategoryTransformationManager =
    {0x218ed275,0xcbe4,0x43f2,{0x91,0x08,0x24,0x2e,0x59,0x68,0xb7,0xe4}};
extern "C" const GUID __declspec(selectany) IID_IUccLocaleString =
    {0x6fdbdefe,0x8a31,0x49cd,{0x87,0xb8,0xcc,0x61,0xa4,0x99,0x25,0xfe}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceActivity =
    {0x890cf073,0xbb33,0x4dcb,{0xae,0xe8,0x08,0xc3,0x00,0x46,0x9b,0x62}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceStateInstance =
    {0x3afa976f,0x86ee,0x412e,{0x98,0x6c,0xeb,0x08,0x4a,0x3d,0x03,0x47}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceStateInstance2 =
    {0xb6608b42,0x8838,0x4ae6,{0xaf,0x83,0xd0,0xf0,0xb4,0x6c,0x7f,0x1a}};
extern "C" const GUID __declspec(selectany) IID_IUccSubscriber =
    {0xd450256a,0xadb1,0x4ea7,{0x9c,0xf1,0x3b,0xd6,0x45,0xb1,0xc6,0xd1}};
extern "C" const GUID __declspec(selectany) IID_IUccSubscriberEx =
    {0x573039b6,0xc916,0x4e9a,{0xa6,0xa2,0x4b,0x05,0x30,0x24,0x8a,0xf6}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceContactCardIdentity =
    {0xf5791faa,0x7225,0x41a7,{0x80,0x43,0x04,0x0d,0xce,0xa4,0x51,0x7a}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceContactCardAddress =
    {0x755b7ded,0xe5a2,0x456a,{0xa5,0xd3,0x43,0xc1,0xb9,0x4c,0xfb,0xcc}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceContactCardPhone =
    {0x9433ccce,0x7a69,0x4a71,{0x97,0x13,0xea,0xe8,0x39,0xea,0xf8,0x3c}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceContactCardUrl =
    {0x06e754c6,0x4bff,0x41fb,{0xba,0x4d,0xaf,0x1a,0x96,0x2b,0xe2,0x36}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceContactCardInstance =
    {0xfc3cfe6f,0x29d0,0x4efc,{0xb3,0xec,0x6c,0x69,0xbb,0xcb,0xcd,0x74}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceContactCardInstance2 =
    {0x9cce5b3c,0x7dd5,0x4452,{0x84,0xf7,0xca,0xda,0xeb,0xc1,0xce,0x53}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceFreeBusyInformation =
    {0x088fea55,0x8dc9,0x4808,{0xa0,0xf5,0xb0,0x4f,0x04,0xc5,0x13,0xaf}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceTimeZoneInformation =
    {0xf3933dc2,0x1e04,0x43c9,{0xa0,0xff,0x5d,0x57,0x8a,0x35,0x2b,0xa8}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceWorkingHoursBlock =
    {0x10659f8e,0xf79f,0x4d16,{0xb2,0x6d,0xa7,0x07,0x90,0x14,0x48,0x3b}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceWorkingHoursInformation =
    {0xd8bec519,0x3f8d,0x4fa3,{0xa9,0xc3,0x43,0xbe,0x4b,0x4e,0x7c,0x0b}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceCapabilityInfo =
    {0xa035a856,0x83b7,0x43bc,{0xb3,0xac,0x1a,0xa1,0xe2,0xc1,0x4f,0xf1}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceCapabilities =
    {0x1de93dcf,0x3bd3,0x411d,{0x82,0x2c,0x6e,0xc2,0xad,0xf9,0xff,0xb7}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceCalendarDataInstance =
    {0x0fae5c66,0x18af,0x4d5d,{0x91,0x64,0xb0,0xe4,0x1d,0xb0,0x4f,0xb7}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceServicesInstance =
    {0xeb59d99c,0x8e23,0x410a,{0x9f,0x49,0xe3,0xa3,0x58,0x5c,0x81,0x16}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceServiceElement =
    {0x21c89f28,0x862d,0x42ce,{0xb1,0x2c,0x09,0x0c,0xd3,0xad,0xea,0x82}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceDeviceInstance =
    {0x53c63418,0xb703,0x4d25,{0x95,0xd6,0xc2,0x38,0x5c,0x18,0xed,0x6a}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceDeviceInstance2 =
    {0x6a332cd1,0x594b,0x413b,{0x9d,0x42,0x55,0xef,0xc5,0xa0,0xa6,0x05}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceNoteData =
    {0xf0e4beea,0x90a7,0x48f4,{0xb8,0xf7,0xf3,0xcc,0xe0,0xd8,0xa8,0xfa}};
extern "C" const GUID __declspec(selectany) IID_IUccPresenceNoteInstance =
    {0xcca44790,0xb93f,0x4564,{0x91,0x35,0x38,0x4f,0xcd,0x1e,0xd9,0x34}};
extern "C" const GUID __declspec(selectany) IID_IUccPresencePhoneLine =
    {0x9081029f,0xbcc4,0x4cf5,{0xac,0x2d,0x44,0x0c,0x97,0xb8,0x31,0x5b}};
extern "C" const GUID __declspec(selectany) IID_IUccRoutingRuleElement =
    {0xb8bc87e3,0x1c8b,0x4f9e,{0x8a,0xb6,0x8b,0xe9,0xa2,0xc6,0x71,0xaf}};
extern "C" const GUID __declspec(selectany) IID_IUccRoutingRuleInstance =
    {0x273819b7,0x8c74,0x4300,{0x8c,0x65,0xfb,0x7e,0x7d,0xfc,0x01,0x46}};
extern "C" const GUID __declspec(selectany) IID_IUccRoutingPreambleElement =
    {0xd0467f80,0x4d40,0x4668,{0x8a,0xcd,0xae,0xef,0x4f,0x69,0xb9,0x2c}};
extern "C" const GUID __declspec(selectany) IID_IUccRoutingPreamble =
    {0xf28f34f5,0x9e9e,0x4166,{0x8d,0x6c,0xc9,0xd1,0xe6,0x26,0x24,0x87}};
extern "C" const GUID __declspec(selectany) IID_IUccRoutingTarget =
    {0xd624a32e,0x76b3,0x4d75,{0xb4,0x8e,0x87,0xa9,0x7d,0x5f,0xd7,0x86}};
extern "C" const GUID __declspec(selectany) IID_IUccRoutingPreambleListElement =
    {0x0255c189,0xdb59,0x4732,{0xbb,0x8c,0x73,0x53,0x48,0x60,0x71,0x1c}};
extern "C" const GUID __declspec(selectany) IID_IUccRoutingPreambleWaitElement =
    {0x27c32a76,0x49b9,0x4a1a,{0xb2,0x93,0x77,0x2e,0x3d,0x5a,0x90,0x20}};
extern "C" const GUID __declspec(selectany) IID_IUccRoutingPreambleFlagsElement =
    {0x1e8e354c,0x7dcf,0x4ca3,{0xa0,0x09,0xed,0x38,0x76,0x01,0x50,0x76}};
extern "C" const GUID __declspec(selectany) IID_IUccServerConfigurationCategory =
    {0x6689760d,0x24fa,0x4b3f,{0xac,0x35,0x20,0x42,0xbf,0xd7,0x0b,0x89}};
extern "C" const GUID __declspec(selectany) IID_IUccSignalingBodyPart =
    {0x3b4688bf,0x4d98,0x4d91,{0xb3,0xfd,0xc3,0x0f,0xe0,0x38,0x95,0x97}};
extern "C" const GUID __declspec(selectany) IID_IUccSignalingMessage2 =
    {0xbf78a2b4,0xcb02,0x4c00,{0x87,0x5f,0xf5,0x1e,0x3f,0x0a,0x74,0xf2}};
extern "C" const GUID __declspec(selectany) IID_IUccSignalingResponse =
    {0x2285d731,0x7269,0x43f8,{0xac,0x54,0x68,0xa6,0x1c,0xa8,0xe1,0xa8}};
extern "C" const GUID __declspec(selectany) IID_IUccUriManager =
    {0xf17f94f7,0xbf07,0x4949,{0xba,0x20,0xeb,0x37,0x4f,0x6d,0x58,0xc0}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingResource =
    {0x79f1c237,0xbd7d,0x4627,{0xbe,0xfd,0x81,0x39,0x86,0x37,0x3f,0xcc}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingChannelResource =
    {0x32d73ef2,0x18a2,0x42f1,{0x81,0x6c,0xef,0x44,0x36,0x71,0x9f,0x40}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingChannelResourceEvent =
    {0x5383c344,0x0684,0x4cdf,{0x88,0x99,0x3a,0xf3,0x96,0x3f,0xa1,0xe1}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharerChannelResource =
    {0x74236283,0x2f10,0x422a,{0xa9,0x9f,0xda,0x2b,0x61,0x58,0x5a,0x53}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingViewer =
    {0xcd0d0e0f,0x79d8,0x422d,{0x9b,0xd5,0xe4,0x39,0x81,0x22,0x53,0xa4}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationViewerChannelResource =
    {0x5247ac0e,0xa9f3,0x4514,{0xba,0xac,0xca,0x7e,0x9f,0x52,0x18,0x0f}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingResourceManager =
    {0xb7112800,0xb839,0x4f64,{0x8c,0xa0,0x5c,0xe2,0x2d,0x1d,0x66,0x66}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingSession =
    {0x7d6773b5,0x19a5,0x4668,{0x88,0x11,0x55,0x10,0xdc,0xe5,0xbe,0xe2}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingSessionParticipant =
    {0x11ceee73,0xaad1,0x40f5,{0xa0,0x9a,0xff,0x88,0xe3,0x21,0x04,0x9d}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingChannel =
    {0x3f25cb45,0xe64f,0x4a0d,{0x86,0x5b,0x08,0x22,0xd6,0x2a,0xc2,0xcc}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingControlEvent =
    {0xf6eaba3c,0x69d7,0x41cf,{0x87,0x10,0xd1,0x30,0x0e,0x84,0x25,0xd1}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingControlRequestAllowedEvent =
    {0x882ef9b5,0x92e6,0x466d,{0xb3,0x1b,0xe7,0xf6,0xa4,0xc3,0xb5,0x7f}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationViewerChannel =
    {0xa5951db4,0x3d41,0x4af0,{0xa3,0x60,0x40,0xe2,0xaf,0x38,0x35,0x27}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingConferenceChannel =
    {0x809f1175,0x58df,0x4b0b,{0x8e,0x7d,0x12,0xc1,0x50,0x50,0xe0,0xd4}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingChannelResources =
    {0x5d87afee,0xc509,0x4400,{0xb7,0x80,0x3d,0xa1,0x3f,0x94,0xf1,0x74}};
extern "C" const GUID __declspec(selectany) IID_IUccMediaChannelConnectionStateChangedEvent =
    {0x81c623e0,0x0167,0x457b,{0x80,0xe6,0xc0,0x89,0xc9,0x10,0x5e,0xeb}};
extern "C" const GUID __declspec(selectany) DIID__IUccApplicationSharingChannelResourceEvents =
    {0x0b804e06,0x2e95,0x40b6,{0x97,0x48,0xa2,0xaa,0xf7,0x79,0x65,0x85}};
extern "C" const GUID __declspec(selectany) DIID__IUccApplicationSharingViewerEvents =
    {0x3677165d,0xf78e,0x4d54,{0x90,0x39,0x5a,0x08,0x22,0x94,0x58,0x4e}};
extern "C" const GUID __declspec(selectany) DIID__IUccApplicationSharingResourceManagerEvents =
    {0xa7d52eb2,0x8ae3,0x47e5,{0x97,0xaf,0xc9,0x6a,0x6f,0x63,0x9c,0x2a}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingResourceChangeEvent =
    {0x74543c38,0x8c7f,0x4eef,{0x85,0xc2,0x1f,0x1e,0xb2,0x81,0xb0,0xd8}};
extern "C" const GUID __declspec(selectany) DIID__IUccApplicationSharingSessionParticipantEvents =
    {0xe9a00f18,0x3f9f,0x4094,{0x83,0x47,0x88,0x7f,0xca,0x4b,0x2b,0x20}};
extern "C" const GUID __declspec(selectany) DIID__IUccApplicationSharingChannelEvents =
    {0x414c2bb7,0x3f99,0x41db,{0xa3,0x69,0xb5,0xd7,0x4c,0x31,0xb5,0xe9}};
extern "C" const GUID __declspec(selectany) DIID__IUccApplicationSharerChannelEvents =
    {0x58ad2cbd,0x83cf,0x4e58,{0x80,0x41,0x7d,0x5d,0x66,0x4a,0x52,0x87}};
extern "C" const GUID __declspec(selectany) DIID__IUccApplicationViewerChannelEvents =
    {0xfd372aa6,0xab45,0x4fe8,{0xa4,0x25,0x4d,0x9c,0x21,0x71,0x71,0x47}};
extern "C" const GUID __declspec(selectany) DIID__IUccApplicationSharingChannelResourcesEvents =
    {0xf4ddfc28,0xee50,0x48df,{0x92,0xae,0x09,0xfc,0x42,0x7c,0xb8,0x55}};
extern "C" const GUID __declspec(selectany) DIID__IUccApplicationSharingChannelCollectionEvents =
    {0x5de958ab,0xa06c,0x4539,{0xbb,0x11,0x1d,0x43,0xb0,0xcd,0xf8,0xbc}};
extern "C" const GUID __declspec(selectany) DIID__IUccPublicationManagerEvents =
    {0x15f5e833,0x3dc1,0x4f9a,{0x99,0x3e,0x6f,0xac,0x96,0xa8,0x2b,0xbe}};
extern "C" const GUID __declspec(selectany) DIID__IUccCategoryInstanceEvents =
    {0x15f8e9d8,0x097d,0x470b,{0x9b,0xad,0x96,0x5c,0xbc,0xfa,0x1e,0x03}};
extern "C" const GUID __declspec(selectany) DIID__IUccCategoryContextEvents =
    {0xc9077987,0xcb46,0x4f9c,{0xb3,0xfd,0x25,0x4f,0xc8,0xfa,0x3f,0x98}};
extern "C" const GUID __declspec(selectany) DIID__IUccContactEvents =
    {0x717efb8d,0x956d,0x467d,{0x8c,0xcf,0x2a,0x46,0x1f,0xd6,0x30,0xda}};
extern "C" const GUID __declspec(selectany) DIID__IUccDelegateEvents =
    {0xcff1d60f,0x10f4,0x4ca6,{0xab,0xc5,0x06,0x3f,0xd9,0x6e,0x98,0x8b}};
extern "C" const GUID __declspec(selectany) DIID__IUccGroupEvents =
    {0xdd825224,0x63e0,0x4a35,{0xbd,0x54,0x9d,0x3c,0xbb,0xed,0xaa,0xbf}};
extern "C" const GUID __declspec(selectany) DIID__IUccSubscriptionEvents =
    {0xbc909c17,0x90d1,0x4c2c,{0xbe,0x68,0xf7,0x47,0x28,0x86,0x6a,0xfc}};
extern "C" const GUID __declspec(selectany) DIID__IUccPresentityEvents =
    {0x5c7e7df6,0x0b5a,0x4ec2,{0x9f,0x04,0xa8,0x75,0x32,0x9b,0x4e,0xc1}};
extern "C" const GUID __declspec(selectany) DIID__IUccPublicationEvent =
    {0x7a3ca589,0x6955,0x4831,{0x87,0x95,0x66,0x23,0xc7,0x8f,0x8d,0xe3}};
extern "C" const GUID __declspec(selectany) DIID__IUccContainerEvents =
    {0x507298f3,0x7975,0x4817,{0x8e,0x35,0x33,0x44,0x7e,0x83,0x7f,0xff}};
extern "C" const GUID __declspec(selectany) DIID__IUccContainerMemberEvents =
    {0x70eb6c1a,0x9445,0x4d34,{0xa9,0xa5,0x4c,0xa0,0x42,0x3a,0xf6,0x76}};
extern "C" const GUID __declspec(selectany) DIID__IUccOperationManagerEvents =
    {0xedb173ee,0xee09,0x407b,{0xb0,0xfe,0xae,0xf5,0xa8,0x95,0x54,0x2f}};
extern "C" const GUID __declspec(selectany) DIID__IUccMediaDeviceManagerEvents =
    {0xcb2b887f,0xd037,0x4309,{0xa5,0x04,0x31,0xd4,0xb3,0xc4,0x15,0x63}};
extern "C" const GUID __declspec(selectany) CLSID_UccPlatform =
    {0xec10e7e8,0x797e,0x4495,{0xa8,0x6d,0x3e,0x9e,0xad,0xa6,0xd5,0xbb}};
extern "C" const GUID __declspec(selectany) CLSID_UccPropertyCollection =
    {0x319beda1,0x82d6,0x4de2,{0x83,0x21,0xbd,0xd5,0xdf,0x64,0xe2,0xf8}};
extern "C" const GUID __declspec(selectany) CLSID_UccOperationContext =
    {0x645c90e5,0xf88c,0x41b1,{0xbf,0x6b,0xf0,0xd7,0x9e,0x99,0xb6,0x2b}};
extern "C" const GUID __declspec(selectany) CLSID_UccContext =
    {0x3af68966,0x5194,0x4de0,{0xbd,0xab,0xae,0x03,0x65,0x3d,0x91,0x4b}};
extern "C" const GUID __declspec(selectany) CLSID_UccSignalingBodyPart =
    {0x08c9df14,0x8b30,0x4120,{0xa9,0x39,0x44,0xd2,0xe6,0x8a,0x50,0x1f}};
extern "C" const GUID __declspec(selectany) CLSID_UccSignalingRequest =
    {0x07b919a9,0x79d7,0x4fe0,{0x8d,0x0b,0xa8,0x10,0x3a,0x67,0x2e,0x2a}};
extern "C" const GUID __declspec(selectany) CLSID_UccSignalingResponse =
    {0x17fe5a0e,0x851e,0x4f26,{0x91,0xd6,0x75,0x04,0x5d,0x45,0x8c,0x32}};
extern "C" const GUID __declspec(selectany) CLSID_UccUriManager =
    {0x201d2dcd,0xaf4f,0x484e,{0x97,0x93,0x56,0xca,0xc5,0xec,0x0c,0x16}};
extern "C" const GUID __declspec(selectany) CLSID_UccApplicationSession =
    {0x85b9cf4d,0x7b3d,0x4cdb,{0xa8,0x2c,0xf5,0xcd,0x6d,0x6d,0x08,0x14}};
extern "C" const GUID __declspec(selectany) CLSID_UccApplicationSessionParticipant =
    {0x54083592,0x55a6,0x4b24,{0x85,0xbd,0x17,0xdc,0x56,0x1e,0x74,0xf6}};
extern "C" const GUID __declspec(selectany) CLSID_UccAudioMediaChannel =
    {0x0c020766,0x8433,0x420d,{0xb8,0x47,0xd6,0x8a,0x3c,0xef,0x80,0x37}};
extern "C" const GUID __declspec(selectany) CLSID_UccAudioMediaDevice =
    {0x51931504,0x36b5,0x4b59,{0xa3,0xf5,0xd4,0xf1,0x57,0x3d,0x84,0xea}};
extern "C" const GUID __declspec(selectany) CLSID_UccAudioVideoSession =
    {0xdc3e6105,0x12a2,0x4d27,{0xb9,0xce,0x2f,0xd9,0xec,0x16,0x60,0x33}};
extern "C" const GUID __declspec(selectany) CLSID_UccAudioVideoSessionParticipant =
    {0xd7dadc80,0x4285,0x4ab5,{0xa9,0x6b,0x8a,0x60,0xc3,0xc6,0x4c,0x41}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceEntityView =
    {0xb351467a,0x29fe,0x4863,{0x90,0x09,0xd0,0x24,0x79,0x8d,0xa0,0x2a}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceInformation =
    {0x12d20681,0x68ac,0x4042,{0xb9,0x3c,0x8a,0x0e,0x19,0xf9,0xa0,0xf3}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceManagerSession =
    {0xe5212020,0xd0d7,0x4dd4,{0xbc,0xd1,0x87,0x5d,0x86,0x90,0x09,0xe1}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceMediaChannel =
    {0x97acd854,0x7422,0x4667,{0x83,0xa7,0x3a,0x87,0x75,0x23,0xba,0x3b}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceView =
    {0xd044bcbe,0x3323,0x432b,{0xa7,0x4a,0xba,0xec,0xab,0x49,0xbb,0xd2}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceSession =
    {0x7ad3396b,0x6fbd,0x4396,{0x8d,0x5c,0x92,0xc7,0xf5,0x39,0x5a,0xa0}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceSessionParticipant =
    {0x98078c74,0xd34b,0x4625,{0xb1,0x43,0xcf,0x34,0xc2,0xd6,0x06,0x33}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceSessionParticipantEndpoint =
    {0xb777414a,0x4fcd,0x4142,{0xbb,0xad,0x1c,0xf6,0x9c,0xb1,0x93,0x81}};
extern "C" const GUID __declspec(selectany) CLSID_UccCredential =
    {0x553a1e8d,0x907f,0x490b,{0xbc,0x05,0x8c,0x97,0x44,0x03,0x6b,0xf9}};
extern "C" const GUID __declspec(selectany) CLSID_UccCredentialCache =
    {0x35e7fe7a,0x1c37,0x48dc,{0xbf,0x7b,0x21,0x0c,0x61,0xfb,0x32,0xe5}};
extern "C" const GUID __declspec(selectany) CLSID_UccDiagnosticInfo =
    {0x5cd2f1ba,0x371b,0x4c43,{0x8d,0x70,0xe0,0x7d,0xbd,0x4f,0xd9,0x6f}};
extern "C" const GUID __declspec(selectany) CLSID_UccHistoryInfoEntry =
    {0x2d85841a,0xc8a1,0x4dab,{0x94,0xbc,0xd7,0x2a,0xc8,0x74,0xb1,0x95}};
extern "C" const GUID __declspec(selectany) CLSID_UccFindServerEvent =
    {0xe308443e,0xe15a,0x4afc,{0x85,0x7e,0xfc,0xcd,0x8b,0x56,0xd9,0x7b}};
extern "C" const GUID __declspec(selectany) CLSID_UccInstantMessagingSession =
    {0x8f929b47,0xbf58,0x4d77,{0xa9,0x51,0x00,0x5f,0x99,0xfe,0xa6,0xa5}};
extern "C" const GUID __declspec(selectany) CLSID_UccInstantMessagingSessionParticipant =
    {0xeff71cf7,0xfec4,0x458d,{0xa7,0xce,0x28,0x31,0x45,0x79,0x6c,0xa2}};
extern "C" const GUID __declspec(selectany) CLSID_UccMediaConnectivityServer =
    {0xc7817076,0xd059,0x4603,{0x9f,0x51,0xc0,0x2b,0xa0,0xbe,0x82,0xbb}};
extern "C" const GUID __declspec(selectany) CLSID_UccMediaPlayerDevice =
    {0xbe6cb026,0x0e90,0x46db,{0x81,0x29,0x07,0x50,0x96,0x11,0x11,0xea}};
extern "C" const GUID __declspec(selectany) CLSID_UccProperty =
    {0x6f67517b,0x135b,0x4a42,{0x8a,0x55,0x59,0x90,0xc3,0x6b,0x26,0x5c}};
extern "C" const GUID __declspec(selectany) CLSID_UccProvisioningPolicyInstance =
    {0x87f779b7,0xf044,0x46f4,{0x8a,0x7f,0xaf,0xee,0x82,0x03,0xfa,0xb2}};
extern "C" const GUID __declspec(selectany) CLSID_UccProxyEndpoint =
    {0xb9634df3,0xa19d,0x4bdb,{0xbe,0xfc,0x84,0x55,0x19,0xc5,0xd7,0xff}};
extern "C" const GUID __declspec(selectany) CLSID_UccRemoteCallControlSession =
    {0xa7cb24ca,0x9c80,0x40ae,{0x93,0x90,0xe1,0x04,0xef,0xdb,0xde,0xbf}};
extern "C" const GUID __declspec(selectany) CLSID_UccRemoteCallControlSessionParticipant =
    {0xc28a10be,0x4d70,0x49a7,{0x87,0xe6,0x01,0x2e,0xb0,0x2b,0x9b,0xec}};
extern "C" const GUID __declspec(selectany) CLSID_UccRemoteCallControlAudioChannel =
    {0x023e043b,0x720b,0x43e8,{0xbe,0xe2,0x7c,0x4f,0x71,0x04,0x15,0xdc}};
extern "C" const GUID __declspec(selectany) CLSID_UccServerEndpoint =
    {0x3af314a7,0xedb7,0x42cf,{0x91,0x5b,0xae,0x50,0xe1,0x60,0x2f,0x6a}};
extern "C" const GUID __declspec(selectany) IID_IUccSignalingChannelManager =
    {0xaf9b29ec,0x84a0,0x482b,{0x8e,0x8a,0x2b,0xd3,0x07,0xd1,0xeb,0x24}};
extern "C" const GUID __declspec(selectany) CLSID_UccPrincipalServerEndpoint =
    {0x92391a7b,0x9135,0x42fa,{0xa4,0x78,0x9e,0x76,0xcb,0xfb,0x7f,0x3c}};
extern "C" const GUID __declspec(selectany) CLSID_UccDelegateEndpoint =
    {0x45d01e9f,0x6d39,0x4a00,{0xa9,0xc9,0x97,0x9b,0xb9,0x15,0x7c,0x1b}};
extern "C" const GUID __declspec(selectany) CLSID_UccAssociatedUriEndpoint =
    {0x284e107f,0x92f4,0x4091,{0xa0,0xdb,0xe1,0x53,0x5b,0xfd,0x94,0x41}};
extern "C" const GUID __declspec(selectany) CLSID_UccSignalingChannel =
    {0x4dc102e6,0xe499,0x45f3,{0x85,0x2b,0xb4,0x3b,0xf0,0x72,0xe4,0x09}};
extern "C" const GUID __declspec(selectany) CLSID_UccVideoMediaChannel =
    {0x7ef89df3,0xd4ab,0x4e3d,{0xb0,0xf0,0x81,0x4d,0x65,0x03,0xd1,0x7b}};
extern "C" const GUID __declspec(selectany) CLSID_UccVideoMediaDevice =
    {0x4c6fc976,0x1b74,0x4775,{0xb4,0x04,0x46,0x2b,0x62,0x80,0x7e,0x77}};
extern "C" const GUID __declspec(selectany) CLSID_UccAudioVideoMediaChannelStreamsEvent =
    {0xd9a93e9b,0x2991,0x4932,{0x82,0x02,0x21,0x22,0x93,0xda,0xd1,0x81}};
extern "C" const GUID __declspec(selectany) CLSID_UccAudioVideoMediaChannelSourceEvent =
    {0x85dd8879,0xb33f,0x4439,{0xa0,0xe7,0x4b,0xe6,0xdc,0xa8,0x13,0xbe}};
extern "C" const GUID __declspec(selectany) CLSID_UccAudioSignalLevelChangedEvent =
    {0x4b7e85ca,0xc96f,0x4b00,{0xb3,0xe0,0xda,0x27,0x28,0xf2,0xcb,0xe7}};
extern "C" const GUID __declspec(selectany) CLSID_UccIncomingMediaRequestEvent =
    {0x76e131bc,0x3b48,0x479f,{0xa0,0x08,0x73,0x42,0x11,0x17,0x63,0x34}};
extern "C" const GUID __declspec(selectany) CLSID_UccIncomingInstantMessageEvent =
    {0x88c1caf3,0x028d,0x4078,{0xab,0x15,0xdc,0x92,0x6d,0xae,0xb9,0x09}};
extern "C" const GUID __declspec(selectany) CLSID_UccIncomingInvitationEvent =
    {0x5a764ba5,0x7409,0x4863,{0x85,0xfa,0xac,0x12,0xca,0xde,0x81,0x49}};
extern "C" const GUID __declspec(selectany) CLSID_UccIncomingSessionEvent =
    {0xf4c31793,0x7a53,0x4e4a,{0x87,0x61,0xb1,0x93,0xb6,0x66,0xac,0x1f}};
extern "C" const GUID __declspec(selectany) CLSID_UccIncomingSessionTransferEvent =
    {0xde455295,0x8d70,0x469c,{0x91,0xf8,0x4b,0x21,0xda,0xb1,0x92,0x43}};
extern "C" const GUID __declspec(selectany) CLSID_UccInstantMessagingComposingEvent =
    {0xaf754b4f,0x53eb,0x41a0,{0xad,0xe5,0x7c,0x68,0x0a,0x02,0x5c,0x3b}};
extern "C" const GUID __declspec(selectany) CLSID_UccInvitationAcceptedEvent =
    {0x6e213239,0x59f9,0x4c14,{0x9b,0xb2,0x84,0x06,0x98,0x5f,0xd5,0xca}};
extern "C" const GUID __declspec(selectany) CLSID_UccMediaChannelEvent =
    {0x4833dc82,0xcfe5,0x4c75,{0x8d,0x86,0x59,0x1a,0x8f,0xd3,0x61,0xb5}};
extern "C" const GUID __declspec(selectany) CLSID_UccOperationProgressEvent =
    {0x65e014d2,0xc007,0x4dc7,{0x89,0x8b,0x92,0xd9,0x92,0x30,0xb1,0x26}};
extern "C" const GUID __declspec(selectany) CLSID_UccOutgoingInvitationEvent =
    {0x0bfeb2bd,0x8bc6,0x47ad,{0x87,0x05,0x58,0x49,0x6a,0x3b,0xc4,0x9c}};
extern "C" const GUID __declspec(selectany) CLSID_UccOutgoingSessionEvent =
    {0x97095729,0x969c,0x4e23,{0x9b,0x43,0x54,0xbc,0x2e,0xae,0x42,0x68}};
extern "C" const GUID __declspec(selectany) CLSID_UccPropertyUpdateEvent =
    {0x70c76318,0xcb44,0x4a52,{0x83,0x08,0xd8,0xd0,0x0d,0x7d,0x6e,0xa6}};
extern "C" const GUID __declspec(selectany) CLSID_UccNamedPropertyUpdateEvent =
    {0xbfef2af5,0xad05,0x48a0,{0x8b,0xbc,0xfb,0xa8,0x97,0x17,0xfb,0xc9}};
extern "C" const GUID __declspec(selectany) CLSID_UccSessionOperationEvent =
    {0x368be194,0xd23d,0x4724,{0xa7,0x41,0xe6,0x06,0x71,0x4f,0x1e,0x2c}};
extern "C" const GUID __declspec(selectany) CLSID_UccSessionParticipantOperationEvent =
    {0x155aa654,0xf81f,0x4a0f,{0xb5,0xae,0xe2,0x13,0x1f,0x3a,0xd0,0x57}};
extern "C" const GUID __declspec(selectany) CLSID_UccSessionParticipantCollectionEvent =
    {0xdfa87481,0xa0d1,0x484a,{0xbf,0x25,0x0b,0x29,0x4f,0x90,0xd7,0x5b}};
extern "C" const GUID __declspec(selectany) CLSID_UccSessionParticipantEndpointCollectionEvent =
    {0x68f95e30,0x7ba8,0x487b,{0x89,0xcd,0x20,0xb6,0x57,0xfc,0x4f,0xdd}};
extern "C" const GUID __declspec(selectany) CLSID_UccSessionParticipantStateChangedEvent =
    {0xb81f580d,0x9ec1,0x4833,{0xbe,0xca,0xb3,0x51,0x94,0xef,0x88,0xec}};
extern "C" const GUID __declspec(selectany) CLSID_UccSessionTransferProgressEvent =
    {0x79dcfdf6,0x5778,0x4484,{0xbc,0xfd,0x37,0x25,0xc0,0xef,0xce,0xab}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceEntityViewCollectionEvent =
    {0x645fa567,0x1348,0x4c39,{0x8d,0x1a,0x8b,0xe5,0xa4,0x9a,0x5b,0xdf}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceSessionParticipantEndpointStateChangedEvent =
    {0x31911030,0x172f,0x4e99,{0xaa,0x09,0x37,0x2f,0x5b,0x8c,0xcd,0x77}};
extern "C" const GUID __declspec(selectany) CLSID_UccMediaDeviceChangeEvent =
    {0xe8905a59,0x86e6,0x4e1f,{0x99,0x8c,0x74,0xed,0xb9,0xc0,0xc3,0x4f}};
extern "C" const GUID __declspec(selectany) CLSID_UccProxyEndpointCollectionEvent =
    {0xabfdaa76,0x8dd7,0x474e,{0x81,0x02,0x49,0x38,0x51,0xc7,0xc9,0x63}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceStateInstance =
    {0x2b7e2ca0,0x4ed4,0x4fb9,{0xb9,0x97,0x96,0xca,0x57,0x04,0x32,0xbe}};
extern "C" const GUID __declspec(selectany) CLSID_UccCategoryContext =
    {0xdef85737,0x89ea,0x41cb,{0xa2,0xeb,0x41,0x67,0x87,0x4d,0x93,0x04}};
extern "C" const GUID __declspec(selectany) CLSID_UccCategoryContextEvent =
    {0x2480d71f,0xdbe2,0x4b25,{0xbd,0xf7,0x06,0xc1,0xae,0x4a,0xd9,0x76}};
extern "C" const GUID __declspec(selectany) CLSID_UccCategoryInstance =
    {0x53e07ea1,0x87d9,0x4cd3,{0xb2,0x06,0xd9,0x3b,0xfb,0xc2,0xa6,0xc5}};
extern "C" const GUID __declspec(selectany) CLSID_UccCategoryInstanceEvent =
    {0x48239b4b,0x3195,0x4c31,{0xa4,0x9f,0x6f,0x4d,0x53,0x30,0x09,0x71}};
extern "C" const GUID __declspec(selectany) CLSID_UccGroupMembershipEvent =
    {0xb87bef40,0xa3eb,0x4bc7,{0x81,0xe2,0x3f,0x9e,0x24,0xed,0x4d,0x9b}};
extern "C" const GUID __declspec(selectany) CLSID_UccContainerMemberCollectionEvent =
    {0xc1bbecc2,0xc8e0,0x45be,{0x9e,0xc2,0xd6,0x29,0xd6,0xc3,0xbe,0xb9}};
extern "C" const GUID __declspec(selectany) CLSID_UccContainerMembershipEvent =
    {0x864272a1,0x5266,0x4e08,{0xb6,0xa2,0x6a,0xea,0xbe,0x6c,0x1f,0x74}};
extern "C" const GUID __declspec(selectany) CLSID_UccContactCollectionEvent =
    {0xfee4f779,0x944a,0x4afa,{0x80,0xe3,0x1e,0xc9,0xb2,0xa3,0xb3,0xa9}};
extern "C" const GUID __declspec(selectany) CLSID_UccSubscription =
    {0x88a91ddd,0x0bf7,0x4536,{0x86,0x87,0x2e,0x20,0x72,0x6b,0x6c,0xb2}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceContactCardIdentity =
    {0x1cbd2388,0x9a85,0x4e72,{0x8b,0xcf,0xfd,0x8c,0xf9,0x02,0x18,0xa0}};
extern "C" const GUID __declspec(selectany) CLSID_UccLocaleString =
    {0x7a26be48,0x15b2,0x473c,{0x8b,0xfe,0xb8,0x0e,0x74,0xc1,0x25,0x7b}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceContactCardAddress =
    {0x690ff5c9,0xdbc8,0x4979,{0xb4,0xb6,0xc8,0xd7,0x27,0xff,0x14,0x7d}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceContactCardPhone =
    {0x8f59329c,0x8167,0x4069,{0x8f,0x56,0x6c,0xa1,0x49,0x28,0x2b,0x57}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceContactCardUrl =
    {0x094142c4,0xeb57,0x405e,{0xb7,0xa8,0x7f,0x8d,0x5d,0x35,0xdb,0xde}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceContactCardInstance =
    {0x9d3857ea,0x1000,0x4cd0,{0xa7,0x8b,0xad,0x37,0x5f,0x68,0x51,0x2d}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceFreeBusyInformation =
    {0x84b09c4b,0x3d6a,0x4e0c,{0xac,0xd8,0xe4,0xe5,0xc6,0xb9,0xc3,0x5b}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceTimeZoneInformation =
    {0x49de5964,0x434b,0x4f82,{0x9d,0x33,0x26,0xab,0x21,0x7a,0x01,0x29}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceWorkingHoursBlock =
    {0xe904a0d9,0xb6a4,0x4266,{0x87,0x28,0x64,0x46,0x93,0x63,0x3d,0x92}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceWorkingHoursInformation =
    {0xc88fac1b,0x8e0b,0x432f,{0xb2,0xae,0xfc,0x51,0x24,0x05,0x9e,0x95}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceActivity =
    {0x34e0c61a,0x23f7,0x454a,{0xbe,0x6f,0x2f,0x8c,0x3f,0xa7,0x25,0xc1}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceCapabilities =
    {0x0f04a7a5,0x3624,0x468c,{0xa5,0xd8,0x58,0xa7,0x90,0x83,0x9e,0xb0}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceCapabilityInfo =
    {0x4262dfa6,0x7a9b,0x4eba,{0xbc,0xc5,0xec,0x02,0xe9,0xab,0x88,0x42}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceCalendarDataInstance =
    {0x967baf31,0x115d,0x43d9,{0xbe,0x38,0x72,0x78,0x0a,0x72,0x29,0xed}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceServicesInstance =
    {0x0c4e3757,0x5d0a,0x4f7a,{0x92,0x5e,0x24,0x7c,0x75,0xfb,0x57,0xa4}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceServiceElement =
    {0xb32e9333,0x4edd,0x4de2,{0xb6,0x71,0x45,0x45,0x44,0x41,0x0c,0x4a}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceDeviceInstance =
    {0x62f1a861,0x25f8,0x4c86,{0xaa,0x14,0x7a,0x81,0x7f,0xdf,0xb0,0xfe}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceNoteData =
    {0xf16ed011,0x533d,0x4041,{0x91,0x35,0xca,0x07,0xa5,0x2d,0x74,0x18}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresenceNoteInstance =
    {0xaa77632f,0xdd46,0x48ee,{0x8e,0xd3,0x15,0x3e,0xf4,0x20,0x6d,0x7f}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresentity =
    {0x18b74f75,0x8884,0x405b,{0xa2,0xdc,0xda,0x17,0x30,0x9c,0x36,0xb0}};
extern "C" const GUID __declspec(selectany) CLSID_UccServerConfigurationCategory =
    {0x9126e974,0xc78d,0x44d0,{0x9d,0x1c,0x33,0xd6,0x91,0xaa,0xda,0x12}};
extern "C" const GUID __declspec(selectany) CLSID_UccSignalingHeaderField =
    {0xe69572c9,0x3681,0x442f,{0xb4,0xcc,0x4b,0xcc,0x30,0x7a,0x37,0xa9}};
extern "C" const GUID __declspec(selectany) CLSID_UccSignalingHeaderFieldParameter =
    {0xadc2ab82,0xb891,0x46bf,{0xa4,0x1d,0x68,0x41,0x13,0x81,0xb3,0xd8}};
extern "C" const GUID __declspec(selectany) CLSID_UccUri =
    {0xa0ac85b9,0x905c,0x4375,{0xb6,0x14,0xd5,0xd3,0x20,0xcd,0x41,0xcd}};
extern "C" const GUID __declspec(selectany) CLSID_UccContact =
    {0xf5a738e2,0x4a28,0x4f3d,{0x97,0x8e,0xa1,0xf8,0x67,0xf5,0x64,0x85}};
extern "C" const GUID __declspec(selectany) CLSID_UccGroup =
    {0xb1910c3e,0x9305,0x454c,{0xb6,0x1e,0xda,0x58,0x0e,0x70,0x8b,0x87}};
extern "C" const GUID __declspec(selectany) CLSID_UccDelegate =
    {0x505fde43,0xe305,0x481d,{0x83,0x41,0x61,0x68,0xdd,0x28,0x61,0x7f}};
extern "C" const GUID __declspec(selectany) CLSID_UccDelegatee =
    {0x9ea23fc2,0xfe8f,0x4508,{0xb8,0xc0,0x73,0xaf,0x8e,0x18,0x50,0x75}};
extern "C" const GUID __declspec(selectany) CLSID_UccRoutingRuleInstance =
    {0x613c60f6,0xe524,0x4123,{0xbd,0x2e,0xb3,0xb1,0x83,0xad,0x37,0xe8}};
extern "C" const GUID __declspec(selectany) CLSID_UccRoutingPreamble =
    {0x50034ab7,0x1512,0x4fd6,{0xbb,0x2f,0xa5,0x66,0xc7,0xbd,0x32,0x91}};
extern "C" const GUID __declspec(selectany) CLSID_UccRoutingPreambleListElement =
    {0xbe62eafa,0x4f29,0x4a44,{0x98,0x3f,0x7a,0x1d,0xe8,0xe0,0xb9,0x2f}};
extern "C" const GUID __declspec(selectany) CLSID_UccRoutingTarget =
    {0x716a4b4e,0x4f09,0x4c14,{0xb3,0x4c,0x3c,0x0c,0x0a,0xd9,0x56,0x6f}};
extern "C" const GUID __declspec(selectany) CLSID_UccRoutingPreambleWaitElement =
    {0x3e9496ad,0x527d,0x458c,{0xb3,0x36,0x79,0xd3,0xba,0x77,0x34,0x55}};
extern "C" const GUID __declspec(selectany) CLSID_UccRoutingPreambleFlagsElement =
    {0xcc36abb1,0xb62c,0x402e,{0x9a,0x29,0x41,0x39,0x53,0x69,0x34,0x0c}};
extern "C" const GUID __declspec(selectany) CLSID_UccSubscriptionEvent =
    {0xaa60bbd9,0x9fae,0x415e,{0x84,0xf1,0xee,0x51,0xe8,0x9c,0x0d,0x09}};
extern "C" const GUID __declspec(selectany) CLSID_UccPresencePhoneLine =
    {0x6a4a9729,0xee1c,0x47db,{0x99,0x0e,0x53,0x87,0x51,0x9e,0xc9,0x28}};
extern "C" const GUID __declspec(selectany) CLSID_UccUserSearchQuery =
    {0xf01ed73c,0xe9cc,0x4edd,{0xbb,0xd9,0xab,0xf4,0xbb,0xa0,0xd3,0xd6}};
extern "C" const GUID __declspec(selectany) CLSID_UccUserSearchQueryEvent =
    {0xfb84e8ba,0xd076,0x45a3,{0x96,0x0d,0x50,0xa8,0x8b,0x05,0x06,0xb0}};
extern "C" const GUID __declspec(selectany) CLSID_UccUserSearchResultRecord =
    {0x40f200e5,0xe499,0x4130,{0x9e,0xe6,0x5a,0x9f,0x41,0x39,0xfe,0x48}};
extern "C" const GUID __declspec(selectany) CLSID_UccCategoryInstanceTransformationManifest =
    {0xe1d00d06,0x988a,0x47c5,{0xa0,0x81,0x3e,0xe6,0xe7,0x2c,0x47,0x39}};
extern "C" const GUID __declspec(selectany) CLSID_UccEndpointRefreshEvent =
    {0x63dda457,0x200b,0x40f1,{0xaf,0xf4,0xde,0xb5,0x54,0x08,0x07,0xac}};
extern "C" const GUID __declspec(selectany) CLSID_UccApplicationSharingSession =
    {0x7f14534b,0x5317,0x4aa3,{0x90,0x82,0xf9,0xb4,0xc3,0x88,0x70,0xe5}};
extern "C" const GUID __declspec(selectany) CLSID_UccApplicationSharingSessionParticipant =
    {0x50e3fc54,0xbf78,0x4342,{0x83,0xcf,0x28,0xf8,0x37,0xe3,0x8b,0xa4}};
extern "C" const GUID __declspec(selectany) CLSID_UccApplicationSharingResourceChangeEvent =
    {0xa81ca775,0x181b,0x4728,{0xab,0x56,0xf3,0x07,0xab,0x9a,0x3d,0x9c}};
extern "C" const GUID __declspec(selectany) CLSID_UccApplicationSharingRequestControlEvent =
    {0xce008d78,0x2bec,0x4b7a,{0x9b,0x10,0x12,0xb9,0xe1,0xed,0x1b,0xed}};
extern "C" const GUID __declspec(selectany) CLSID_UccApplicationSharingControlEvent =
    {0x5618d4f8,0xbbc8,0x4363,{0x90,0x5f,0xbd,0x12,0x6b,0xfa,0x17,0xf6}};
extern "C" const GUID __declspec(selectany) CLSID_UccApplicationSharingControlRequestAllowedEvent =
    {0x5b673bde,0xc698,0x4991,{0xa5,0x56,0x24,0xf6,0x66,0xd6,0x05,0x53}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceMediaChannelSourceEvent =
    {0x2cff1652,0x10a0,0x46fb,{0xa1,0xd8,0x62,0xef,0x68,0x5f,0xfe,0xbd}};
extern "C" const GUID __declspec(selectany) CLSID_UccMediaChannelConnectionStateChangedEvent =
    {0xf38cdd95,0xd1ba,0x4497,{0xb8,0x4f,0x90,0xc5,0x82,0x92,0xed,0x3d}};
extern "C" const GUID __declspec(selectany) CLSID_UccConferenceMediaChannelStateChangedEvent =
    {0xcc4e82db,0xd848,0x4309,{0x8c,0x3c,0x47,0x3f,0xc1,0x50,0xb4,0xeb}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceMediaChannel =
    {0xaf404bf1,0x2495,0x4c3e,{0xbf,0x3b,0xbe,0x83,0x59,0x34,0x1f,0x78}};
extern "C" const GUID __declspec(selectany) IID_IUccSessionParticipantEndpoint =
    {0x16b1749a,0x39ce,0x4fd1,{0x94,0x5c,0xb1,0x3d,0x56,0x92,0x61,0xb9}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceApplicationViewerChannel =
    {0xa0897f48,0xa177,0x4c66,{0xa2,0x23,0xf7,0xb8,0xd6,0x95,0xb6,0xb8}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceSessionParticipant =
    {0x09717930,0xa1de,0x4b46,{0xba,0xe2,0x08,0xf3,0xe5,0xc7,0x4f,0xa2}};
extern "C" const GUID __declspec(selectany) IID_IUccAudioVideoMediaChannelSourceEvent =
    {0x2c46dc00,0x1382,0x4c65,{0x8b,0x79,0x3c,0xed,0xb3,0x58,0xe5,0xd0}};
extern "C" const GUID __declspec(selectany) IID_IUccConferenceMediaChannelSourceEvent =
    {0xe5837f40,0x2f55,0x458f,{0x9a,0xdf,0x42,0xc1,0xbf,0x97,0x35,0x60}};
extern "C" const GUID __declspec(selectany) IID_IUccIncomingInstantMessageEvent =
    {0xca7acf03,0x537d,0x4654,{0x92,0xf6,0xb5,0x56,0xfd,0xe0,0xcd,0x80}};
extern "C" const GUID __declspec(selectany) IID_IUccInstantMessagingComposingEvent =
    {0x11ebeff7,0xe4c7,0x4bc3,{0x87,0x52,0x49,0x3b,0x94,0x70,0x04,0xfa}};
extern "C" const GUID __declspec(selectany) IID_IUccSessionParticipantEndpointCollectionEvent =
    {0x683701e6,0x6c96,0x441a,{0x8a,0xb2,0xb5,0x29,0x2d,0xbb,0x71,0x04}};
extern "C" const GUID __declspec(selectany) IID_IUccSessionParticipantOperationEvent =
    {0x8c37fed9,0xab4a,0x45bc,{0xb0,0x18,0x5c,0xee,0xfb,0x94,0xf3,0x4f}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharerChannel =
    {0x1c88a56d,0x8f5f,0x4793,{0xa9,0x90,0xe5,0x33,0x3b,0xfe,0x8d,0x37}};
extern "C" const GUID __declspec(selectany) IID_IUccApplicationSharingRequestControlEvent =
    {0xad1a9ef1,0x2584,0x4693,{0x83,0xc4,0x1f,0x47,0x9d,0x8b,0x4f,0xba}};
extern "C" const GUID __declspec(selectany) IID_IUccSessionParticipant =
    {0xc520c114,0x7ebf,0x4a14,{0x84,0x16,0xd1,0x7a,0xdf,0xa8,0x62,0x27}};
extern "C" const GUID __declspec(selectany) IID_IUccSession =
    {0x233693fb,0xad38,0x4152,{0x8d,0x4f,0x68,0x8f,0xdc,0xc3,0xfc,0xde}};
extern "C" const GUID __declspec(selectany) IID_IUccSessionCallControl =
    {0x2d237189,0x43ff,0x40ce,{0xbd,0xfb,0x1f,0x56,0x0d,0x56,0xb4,0xfa}};
extern "C" const GUID __declspec(selectany) IID_IUccSessionManager =
    {0x57061595,0xcea1,0x4c1d,{0xaf,0x44,0x3d,0xff,0x48,0xa7,0xe0,0xdf}};
extern "C" const GUID __declspec(selectany) IID_IUccSessionParticipantCollectionEvent =
    {0x005a2d10,0xdd52,0x4ace,{0x9c,0xa1,0x9b,0xdd,0x69,0x93,0x15,0xf0}};
extern "C" const GUID __declspec(selectany) IID_IUccIncomingSessionEvent =
    {0x37cf9d37,0xc22c,0x4c5d,{0x90,0x23,0xf7,0x8a,0xe5,0x32,0x50,0x36}};
extern "C" const GUID __declspec(selectany) IID_IUccOutgoingSessionEvent =
    {0xe11e841b,0xda1d,0x4425,{0xbb,0x62,0x44,0xdf,0x44,0x18,0xcb,0x24}};
extern "C" const GUID __declspec(selectany) IID_IUccSubscriptionManager =
    {0xfb194526,0x2e78,0x4cca,{0x92,0x7e,0xe9,0x67,0xbf,0x01,0x01,0xf7}};
extern "C" const GUID __declspec(selectany) IID_IUccSubscription =
    {0xf6e4f91e,0x40e2,0x4bf2,{0xa8,0x08,0x37,0xec,0x02,0xda,0xf6,0x7b}};
extern "C" const GUID __declspec(selectany) IID_IUccPresentity =
    {0xd0171dc1,0x7a36,0x494b,{0xa8,0xe5,0xcc,0x51,0x71,0x1a,0x8c,0xac}};
extern "C" const GUID __declspec(selectany) IID_IUccSubscriptionEvent =
    {0x104cc179,0x5a41,0x4539,{0xa5,0x7f,0x70,0x8b,0xcb,0xaa,0xac,0xd8}};
extern "C" const GUID __declspec(selectany) IID_IUccCategoryContext =
    {0x69f9efcf,0x5fea,0x46b1,{0x83,0x95,0xbd,0x96,0x21,0x2a,0x8c,0x51}};
extern "C" const GUID __declspec(selectany) IID_IUccCategoryContextEvent =
    {0xf13220e8,0xdec0,0x4e04,{0xac,0x27,0x3a,0x77,0x85,0xe9,0xd8,0xc5}};
extern "C" const GUID __declspec(selectany) IID_IUccCategoryInstance =
    {0x72d6a6c3,0xac51,0x4166,{0x93,0xf0,0xe9,0x41,0xc3,0xd7,0x84,0xdf}};
extern "C" const GUID __declspec(selectany) IID_IUccCategoryInstanceEvent =
    {0x6ed14bd9,0x627d,0x4d9d,{0x94,0x05,0xed,0x9a,0xd0,0xf0,0x70,0x23}};
extern "C" const GUID __declspec(selectany) IID_IUccPublication =
    {0xf1d73cad,0x6f6a,0x413b,{0xbc,0x2b,0x15,0x8d,0xc2,0xa7,0xcc,0xb2}};
extern "C" const GUID __declspec(selectany) IID_IUccContainer =
    {0x646d6d11,0xc95c,0x49b2,{0xa6,0x0f,0x77,0xc5,0xe9,0x14,0x08,0xd9}};
extern "C" const GUID __declspec(selectany) IID_IUccContainerMember =
    {0x0dd847ab,0x0f32,0x4449,{0x91,0x65,0x74,0x25,0xd1,0x1d,0xdd,0x89}};
extern "C" const GUID __declspec(selectany) IID_IUccPublicationManager =
    {0x8698350f,0x4b7e,0x4e38,{0xa7,0xe7,0x83,0x03,0x8b,0x69,0xdb,0x7b}};
extern "C" const GUID __declspec(selectany) IID_IUccContainerMemberCollectionEvent =
    {0x679d5f13,0x90bb,0x40ab,{0x98,0xf7,0x62,0x43,0xb2,0x5b,0x8d,0x89}};
extern "C" const GUID __declspec(selectany) IID_IUccContainerMembershipEvent =
    {0xe7122690,0xcbf5,0x4d7d,{0x9f,0x1b,0xbf,0x28,0x66,0x8d,0x91,0xa9}};
extern "C" const GUID __declspec(selectany) IID_IUccContainerMembershipManager =
    {0x92700fb4,0xff9b,0x46b7,{0xaf,0x97,0x78,0x9d,0x3d,0x0e,0x4c,0xe3}};
extern "C" const GUID __declspec(selectany) IID_IUccContact =
    {0x70c8f58f,0xa1ca,0x4614,{0x87,0x60,0x68,0x39,0x9e,0x11,0x1c,0xfa}};
extern "C" const GUID __declspec(selectany) IID_IUccGroup =
    {0xf36d99bb,0x5cb1,0x4dd6,{0x96,0x77,0x11,0x78,0x70,0x69,0xe4,0x05}};
extern "C" const GUID __declspec(selectany) IID_IUccContactCollectionEvent =
    {0x0994e0c2,0xb0ac,0x46b5,{0xb2,0x58,0x7d,0xf5,0x1f,0xab,0x8b,0x2e}};
extern "C" const GUID __declspec(selectany) IID_IUccGroupMembershipEvent =
    {0x7377729f,0x2fa4,0x4e73,{0xa6,0xa9,0xcd,0xb8,0xb0,0x35,0x67,0x6d}};

} // namespace UCCAPILib

#pragma pack(pop)
