
/********************************************************************************
*					Tangram Library - version 10.0.0							*
*********************************************************************************
* Copyright (C) 2006-2013 by Tangram Team.   All Rights Reserved.				*
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
* mailto:sunhuizlz@yeah.net
* http://www.cloudaddin.com
*
********************************************************************************/
#pragma once
#pragma pack(push, 8)

#include <comdef.h>

namespace UCCollaborationLib {
//
// Forward references and typedefs
//

struct __declspec(uuid("b9aa1f11-f480-4054-a84e-b5d9277e40a8"))
/* LIBID */ __UCCollaborationLib;
enum PropertyMask;
enum PropertyMaskIndex;
enum PropertyType;
enum PropertyClass;
enum ConversationWindowExtensionSize;
enum ModalityTypes;
struct __declspec(uuid("f7ee956b-a713-48d3-afef-0a28f37740b3"))
/* dual interface */ IAsynchronousOperation;
struct __declspec(uuid("c12f0de5-9a7d-425c-b391-8be004eaa2f6"))
/* dual interface */ ILocaleString;
struct __declspec(uuid("dc63d621-11c8-4490-b017-b8c57ee1be25"))
/* dual interface */ IAddress;
struct __declspec(uuid("1a8a9402-e89f-40ae-88a8-b328391684b6"))
/* dual interface */ INamedPropertyList;
struct /* coclass */ AsynchronousOperation;
struct /* coclass */ LocaleString;
struct /* coclass */ Address;
struct /* coclass */ NamedPropertyList;
enum ContactSetting;
enum GroupAction;
enum GroupType;
enum ContactInformationType;
enum UnifiedCommunicationType;
enum ContactSubscriptionRefreshRate;
enum ContactSourceTypes;
enum SearchProviders;
enum SearchFields;
enum SearchOptions;
enum SearchProviderStatusType;
enum ContactType;
enum SourceNetworkType;
enum ContactCalendarState;
enum DaysOfWeek;
enum DefaultNoteType;
enum ContactCapabilities;
enum ContactEndpointType;
enum OrganizationStructureTypes;
struct __declspec(uuid("ec9773ad-7d22-4b03-b949-5e9055af0f19"))
/* dual interface */ IGroupCollectionChangedEventData;
struct __declspec(uuid("f97973d7-9436-4a3e-a802-f6ea172a0425"))
/* dual interface */ IGroup;
struct __declspec(uuid("57c3cb4d-c794-456f-846d-0e1eca78bd00"))
/* dual interface */ IContactCollection;
struct __declspec(uuid("85109845-6ae2-4b6b-9524-65ba203e5b7d"))
/* dual interface */ IContact;
struct __declspec(uuid("f0ce810f-e36f-4cd8-a836-515b9a6e09fc"))
/* dual interface */ IContactInformationDictionary;
struct __declspec(uuid("fd9000b3-479f-4b16-9d63-70a49e078946"))
/* dual interface */ IContactSettingDictionary;
struct __declspec(uuid("e6b3f5fa-4208-4c37-ae07-f6435b68d693"))
/* dual interface */ IContactManager;
struct __declspec(uuid("0c81ec90-9469-4ed4-acad-adb86e11039c"))
/* dual interface */ IGroupCollection;
struct __declspec(uuid("95b8d3e8-f3d5-4dc0-bcfe-ab80c835dea0"))
/* dual interface */ IContactSubscription;
struct __declspec(uuid("340c54a6-35cf-4971-b540-72d7f040ac24"))
/* dual interface */ IDistributionGroup;
struct __declspec(uuid("1ff0e6f1-2fe4-4e29-a123-557af0db6927"))
/* dual interface */ IContactEndpoint;
struct __declspec(uuid("efec2816-f16d-48d8-9306-26c810f0ea55"))
/* dual interface */ IContactInformationChangedEventData;
struct __declspec(uuid("4b3a2412-c149-4e52-a713-6cebcf47503e"))
/* dual interface */ IGroupMemberChangedEventData;
struct __declspec(uuid("d2bfc7d7-a64f-48ef-bf90-c3c7e684d89c"))
/* dual interface */ IDistributionGroupExpansionStateChangedEventData;
struct __declspec(uuid("0bdb9057-28ae-4bf0-aff0-12a148e51637"))
/* dual interface */ IUriChangedEventData;
struct __declspec(uuid("283c2089-b760-4d65-9199-716627174f7a"))
/* dual interface */ IGroupNameChangedEventData;
struct __declspec(uuid("378cd767-1b64-4a0e-a941-8f52869fc9df"))
/* dual interface */ ISearchResult;
struct __declspec(uuid("7356d7bb-fd71-4554-84a1-3bbe28726551"))
/* dual interface */ ISearchResults;
struct __declspec(uuid("995f992c-9def-47b9-bf11-81813c0c0e28"))
/* dual interface */ ISearchProviderStateChangedEventData;
struct __declspec(uuid("d804ba85-db0f-49ee-b723-dcdf0425bd9f"))
/* dual interface */ ICustomGroup;
struct __declspec(uuid("dcdea425-b5f2-4719-a3b4-69ffb9770be6"))
/* dual interface */ IRecentContacts;
struct __declspec(uuid("18a9a79a-cbc7-4f8b-8b2b-9b40ca75ec3c"))
/* dual interface */ IFrequentContacts;
struct __declspec(uuid("35ce3a44-c636-4d31-bc95-b9370c2251b9"))
/* dual interface */ IFavoriteContacts;
struct __declspec(uuid("339d28c6-3d1a-45d4-ba14-a56742749a43"))
/* dual interface */ IWorkingPeriod;
struct __declspec(uuid("53d014c1-54db-42b3-9dfd-8e231ef2c356"))
/* dual interface */ IWorkingHours;
struct __declspec(uuid("10c3acc7-6c0c-4f35-adfb-ae0f200a8583"))
/* dual interface */ IPresenceCapability;
enum PresenceCapabilityType;
struct __declspec(uuid("8affb8ce-5404-4280-bca3-e2e4388e6d73"))
/* dual interface */ IDescriptionChangedEventData;
struct __declspec(uuid("10fdd9ba-0cba-4958-b6c8-d0912bf2703f"))
/* dual interface */ IContact2;
struct __declspec(uuid("7cd92461-ca4f-4beb-a636-8bbcc192e60f"))
/* dual interface */ ISearchResults2;
struct __declspec(uuid("abf012a2-2068-4b4d-80d7-6d386ba7c8a2"))
/* dispinterface */ _IContactsAndGroupsCallback;
struct __declspec(uuid("c6605772-3d9d-44a7-b43a-4590e6b43ee3"))
/* dispinterface */ _IGroupCallback;
struct __declspec(uuid("311dde48-ed7d-46fa-9e0a-2e314b7fef7c"))
/* dispinterface */ _IContactCallback;
struct __declspec(uuid("50d86ee2-288f-44f5-8144-69f6e3b24b90"))
/* dispinterface */ _IContactManagerEvents;
struct __declspec(uuid("b1c6ecc3-1ea7-495d-9337-02fe4fb723f0"))
/* dispinterface */ _IContactEvents;
struct __declspec(uuid("12480bd4-072f-4763-b9fb-41b2ca54f9cd"))
/* dual interface */ IContactSettingChangedEventData;
struct __declspec(uuid("9918673d-4d74-41cf-8840-d4996395a34c"))
/* dispinterface */ _IGroupEvents;
struct __declspec(uuid("f719ed75-b7bd-45d3-9097-254509226f20"))
/* dispinterface */ _ICustomGroupEvents;
struct __declspec(uuid("011b320c-9ac1-4dde-b595-3d0014472e2f"))
/* dispinterface */ _IRecentContactsEvents;
struct __declspec(uuid("082ee280-726e-417f-99cb-81a0ccbff883"))
/* dispinterface */ _IDistributionGroupEvents;
struct __declspec(uuid("aa186c92-181e-417f-b150-fca0f367e0fc"))
/* dispinterface */ _IFrequentContactsEvents;
struct __declspec(uuid("62b932e9-7847-4d78-beee-5d58bf815b03"))
/* dual interface */ IContactPositionChangedEventData;
struct __declspec(uuid("dab42e13-500a-409a-ac34-c1bf4719ee2e"))
/* dispinterface */ _IFavoriteContactsEvents;
struct /* coclass */ ContactManager;
struct /* coclass */ Contact;
struct /* coclass */ Group;
struct /* coclass */ ContactCollection;
struct /* coclass */ CustomGroup;
struct /* coclass */ DistributionGroup;
struct /* coclass */ RecentContacts;
struct /* coclass */ FrequentContacts;
struct /* coclass */ FavoriteContacts;
struct /* coclass */ GroupMemberChangedEventData;
struct /* coclass */ DistributionGroupExpansionStateChangedEventData;
struct /* coclass */ UriChangedEventData;
struct /* coclass */ GroupNameChangedEventData;
struct /* coclass */ GroupCollectionChangedEventData;
struct /* coclass */ SearchResult;
struct /* coclass */ SearchResults;
struct /* coclass */ SearchProviderStateChangedEventData;
struct /* coclass */ GroupCollection;
struct /* coclass */ ContactInformationChangedEventData;
struct /* coclass */ ContactSettingChangedEventData;
struct /* coclass */ ContactPositionChangedEventData;
struct /* coclass */ WorkingHours;
struct /* coclass */ ContactInformationDictionary;
struct /* coclass */ ContactSettingDictionary;
struct /* coclass */ WorkingPeriod;
struct /* coclass */ ContactSubscription;
struct /* coclass */ PresenceCapability;
struct /* coclass */ ContactEndpoint;
struct /* coclass */ DescriptionChangedEventData;
enum ConversationSummarizerStates;
enum ConversationState;
enum ConversationAcceptanceState;
enum ModalityState;
enum ModalityAction;
enum ParticipantAction;
enum ConversationAction;
enum ConversationHistoryAction;
enum ConversationImportance;
enum ConversationDirection;
enum ConferenceEscalationProgress;
enum ConferenceAccessType;
enum ConferenceDisclaimerState;
enum ModalityDisconnectReason;
enum TransferOptions;
enum ModalityConnectOptions;
enum ChannelType;
enum ChannelState;
enum ChannelAction;
enum MediaStreamDirection;
enum ConversationContextType;
enum RepresentationType;
enum HistoryInfoEntryType;
enum ConversationContextUpdateOptions;
enum PreviousConversationHistoryRetrievalState;
enum ParticipantEventStateChangeAlsoRingingOthers;
enum ModalityActionProperty;
enum ModalityProperty;
enum ConversationProperty;
enum ParticipantProperty;
enum ConversationStateChangeProperty;
enum ModalityStateChangeProperty;
enum RedirectionTypes;
enum InstantMessageCapabilities;
enum SpeakerMode;
enum HoldFeature;
enum MediaQuality;
enum PreviousConversationsType;
enum PreviousConversationsErrorType;
enum PreviousConversationsConnectionState;
enum CallUnparkReason;
enum SoftMuteChangeInitiator;
enum InstantMessageContentType;
enum ContextType;
enum SharingResourceType;
enum ParticipationState;
enum CanShareDetail;
struct __declspec(uuid("a4e1d1e4-7839-473b-95fa-5d33ddc0d2dd"))
/* dual interface */ IConversationCollection;
struct __declspec(uuid("533891f1-a460-43e6-8bda-6d0d2c13adde"))
/* dual interface */ IConversation;
struct __declspec(uuid("4875c050-ba9c-4a05-891e-e7b0a9463664"))
/* dual interface */ IParticipant;
struct __declspec(uuid("ed5040fa-8f8f-4876-bf6d-9c62f5effd5b"))
/* dual interface */ IModalityDictionary;
struct __declspec(uuid("d4b240b9-059b-4d13-8997-31b282b3bd43"))
/* dual interface */ IModality;
struct __declspec(uuid("62074904-8d06-43fe-a531-e63df7fdc2e7"))
/* dual interface */ IModalityPropertyDictionary;
struct __declspec(uuid("8665fd95-7720-4f9e-b605-6abcbd7edff4"))
/* dual interface */ IParticipantPropertyDictionary;
struct __declspec(uuid("fd02e775-cb9d-4ce3-9243-115c81e5f11c"))
/* dual interface */ IParticipantCollection;
struct __declspec(uuid("855ad949-e2eb-473d-aaf2-9800008ea303"))
/* dual interface */ IConversationContextList;
struct __declspec(uuid("8a88c016-aa5d-4eae-9c79-7aa5c360ed76"))
/* dual interface */ IConversationContextCollection;
struct __declspec(uuid("04378c72-e58f-47c0-8621-901b85ca2ed4"))
/* dual interface */ IConversationContext;
struct __declspec(uuid("b3ff95bd-b743-498c-9973-7f8a2a78576a"))
/* dual interface */ IParticipantInfo;
struct __declspec(uuid("f5996c71-77ab-4020-9776-c2c270fef2ae"))
/* dual interface */ IConversationHistory;
struct __declspec(uuid("00e22cbb-3170-453a-ae62-eafbc75a9f8d"))
/* dual interface */ IConversationPropertyDictionary;
struct __declspec(uuid("c31fb52d-9e9e-45ad-a102-5218e1685b78"))
/* dual interface */ IConversationManager;
struct __declspec(uuid("e24c2484-c70d-44ab-8a4e-f1585a17a3e1"))
/* dual interface */ IVoiceMailsManager;
struct __declspec(uuid("15d4461d-c7eb-45cb-b214-f1c45fed1db2"))
/* dual interface */ IPreviousConversationsManagerDictionary;
struct __declspec(uuid("acabfa1c-0738-4c4d-af7e-326873cf4ace"))
/* dual interface */ IPreviousConversationsManager;
struct __declspec(uuid("209da899-15f1-499a-b8b5-a847eaa899ae"))
/* dual interface */ IRepresentationInfo;
struct __declspec(uuid("62a75516-c79b-42d7-8b49-3ba492c2b385"))
/* dual interface */ IInstantMessageModality;
struct __declspec(uuid("e4e1d0af-5b70-45af-9a0e-f548d8fbd17c"))
/* dual interface */ IAVModality;
struct __declspec(uuid("8e839af9-9508-475a-aa47-3ee8a946b75e"))
/* dual interface */ IAudioChannel;
struct __declspec(uuid("61c747f6-157c-4cae-a37c-ec4352175860"))
/* dual interface */ IChannel;
struct __declspec(uuid("e66ca1cc-9db0-467e-9c60-6a832acd6780"))
/* dual interface */ IVideoChannel;
struct __declspec(uuid("56a868b4-0ad4-11ce-b03a-0020af0ba770"))
/* dual interface */ IVideoWindow;
struct __declspec(uuid("ac10d6e9-4a8c-4484-b8f4-ca1e36347aae"))
/* dual interface */ IApplicationSharingModality;
struct __declspec(uuid("6a7ff112-42ce-4d03-8922-d02ac6c5759f"))
/* dual interface */ ISharingResourceList;
struct __declspec(uuid("3690d374-3779-4efc-b8e8-20642c01529a"))
/* dual interface */ ISharingResource;
struct __declspec(uuid("7b39cb77-4aeb-42d7-b351-cb5472c1c6ad"))
/* dual interface */ IChannel2;
struct __declspec(uuid("3cd1db65-d85b-4912-8536-6bd8e19f1738"))
/* dual interface */ IVideoChannelCollection;
struct __declspec(uuid("478dda59-64c0-4429-ac8a-f10a940b52cc"))
/* dual interface */ IModalityStateChangedEventData;
struct __declspec(uuid("c79e4169-535f-4a41-8de2-c65a90d9503d"))
/* dual interface */ IModalityStateChangePropertyDictionary;
struct __declspec(uuid("deddd7bd-4763-41d2-9aaa-b2c143457ca4"))
/* dual interface */ IModalityActionAvailabilityChangedEventData;
struct __declspec(uuid("2a9385aa-cf61-4e47-b64c-214de4a4ad11"))
/* dual interface */ IConversationStateChangedEventData;
struct __declspec(uuid("56f7dec0-59ca-47c6-9f35-d5066a702b39"))
/* dual interface */ IConversationStateChangePropertyDictionary;
struct __declspec(uuid("88c688b0-3908-4c56-a2e8-f90ab705c536"))
/* dual interface */ IConversationActionAvailabilityEventData;
struct __declspec(uuid("8057f99d-292c-4371-926b-5312e61a3a40"))
/* dual interface */ ISummarizerStatesChangedEventData;
struct __declspec(uuid("ab2d1a67-2b3d-4e25-b21c-03f4bfc4c2be"))
/* dual interface */ IConversationManagerEventData;
struct __declspec(uuid("064e6072-5592-4d13-afdb-55f5a494e137"))
/* dual interface */ IParticipantCollectionChangedEventData;
struct __declspec(uuid("905b727d-fd8f-46ec-8006-ce245f6835ee"))
/* dual interface */ IMessageSentEventData;
struct __declspec(uuid("dba05c15-1c07-4a76-8248-08d8416a24e3"))
/* dual interface */ IIsTypingChangedEventData;
struct __declspec(uuid("d5c1c8fb-dc8a-45a3-aaa3-53b89d43cde4"))
/* dual interface */ IIsContributingChangedEventData;
struct __declspec(uuid("2495b94c-38af-439b-bba0-0ad38d959ce4"))
/* dual interface */ IChannelStateChangedEventData;
struct __declspec(uuid("ca7eeb7a-7dc3-4ffe-a174-23db5a003c04"))
/* dual interface */ IStreamStateChangedEventData;
struct __declspec(uuid("9d4d3d10-debc-40c6-acf9-527986cb42bc"))
/* dual interface */ IChannelActionAvailabilityEventData;
struct __declspec(uuid("3debdc7e-811f-4304-bdd7-7d98c732026c"))
/* dual interface */ IHistoryInfo;
struct __declspec(uuid("7cfe77cd-731d-48b2-82b1-eca3414d62e3"))
/* dual interface */ IConversationPropertyChangedEventData;
struct __declspec(uuid("70dcce38-90bf-4428-a34f-3a6082f29e50"))
/* dual interface */ IParticipantPropertyChangedEventData;
struct __declspec(uuid("4290152c-75c9-47c8-aa12-814818360b3f"))
/* dual interface */ IModalityPropertyChangedEventData;
struct __declspec(uuid("25a47938-5c31-4f43-9108-da01ee5869c0"))
/* dual interface */ IVoiceMail;
struct __declspec(uuid("f644c610-a146-4a56-8338-a69c45c71cef"))
/* dual interface */ IVoiceMailCollection;
struct __declspec(uuid("df414a68-5051-4465-aae2-4f301315734e"))
/* dual interface */ IPreviousConversation;
struct __declspec(uuid("81c9d13f-a4f9-4e13-92d3-bb271e8df3d2"))
/* dual interface */ IPreviousConversationCollection;
struct __declspec(uuid("f87afd9a-9bf8-4ac4-953e-33c4d2035d33"))
/* dual interface */ IPreviousConversationsManagerConnectionStateChangedEventData;
struct __declspec(uuid("2c59a886-acd3-46bb-8bac-7c1c59fb2e0d"))
/* dual interface */ IPreviousConversationsManagerNewItemCountChangedEventData;
struct __declspec(uuid("ebd19735-0cb5-4af3-8e05-b0140cf9d545"))
/* dual interface */ IPreviousConversationBatchEventData;
struct __declspec(uuid("18ba13c7-a64e-4301-ba51-d1bfb3c1c9d5"))
/* dual interface */ IVoiceMailBatchEventData;
struct __declspec(uuid("2f8c3e58-436b-42db-8924-6c394b37dca2"))
/* dual interface */ IConferenceAccessInformation;
struct __declspec(uuid("1124aa59-b6f3-4eb1-b79a-e7f53b388b8d"))
/* dual interface */ IParticipantActionAvailabilityChangedEventData;
struct __declspec(uuid("37c7f935-e215-46d2-82a3-a451840ec733"))
/* dual interface */ IConversation2;
struct __declspec(uuid("d60b1701-766d-401e-8586-83e0c9106be0"))
/* dual interface */ IParticipant2;
struct __declspec(uuid("25025b1b-0084-44d1-b383-9ffc9a99abcd"))
/* dispinterface */ _IConversationManagerEvents;
struct __declspec(uuid("c63f58cb-ef11-4934-aa3b-2d69244f0780"))
/* dispinterface */ _IConversationEvents;
struct __declspec(uuid("ef5c00e4-75e2-4d39-b358-0ed9cb14f3cb"))
/* dual interface */ IConversationContextCollectionEventData;
struct __declspec(uuid("92d7e7a8-48d8-4166-8911-630ae02b2b93"))
/* dual interface */ IInitialContextEventData;
struct __declspec(uuid("21629ab6-9557-4985-9237-49177f618692"))
/* dual interface */ IContextEventData;
struct __declspec(uuid("6a986b5d-e5ec-4395-ad2d-0e62b72da776"))
/* dispinterface */ _IParticipantEvents;
struct __declspec(uuid("5418e2ab-eb9a-4659-b4dc-28de633b2b8f"))
/* dual interface */ IMutedChangedEventData;
struct __declspec(uuid("10d66b39-5e4b-4c98-aac6-d0458feb1bdb"))
/* dispinterface */ _IModalityEvents;
struct __declspec(uuid("4fe44049-4e44-4109-b234-4e4efc135a86"))
/* dispinterface */ _IInstantMessageModalityEvents;
struct __declspec(uuid("acd2fdb9-98ed-4283-aae8-3cd989ebd9ae"))
/* dispinterface */ _IAVModalityEvents;
struct __declspec(uuid("be026cd2-7e82-4f7c-8762-f6b02f496174"))
/* dispinterface */ _IApplicationSharingModalityEvents;
struct __declspec(uuid("b22edbea-9e61-4703-82be-01c05619b6d6"))
/* dual interface */ ILocalSharedResourcesChangedEventData;
struct __declspec(uuid("8fe9a23d-72bc-4ea3-a5af-e37f75ead8a2"))
/* dual interface */ IParticipationStateChangedEventData;
struct __declspec(uuid("32430741-330b-48df-96f6-0cfe00484d40"))
/* dual interface */ IControlRequestReceivedEventData;
struct __declspec(uuid("4e8dc7e0-04b8-470b-bdfa-f520099b975f"))
/* dual interface */ IControllerChangedEventData;
struct __declspec(uuid("d0dbb17d-565e-486c-a47b-ba32dc1fad0a"))
/* dual interface */ ISharerChangedEventData;
struct __declspec(uuid("fa7971ef-1e41-429f-8c79-dac5f88d444b"))
/* dispinterface */ _IConversationHistoryEvents;
struct __declspec(uuid("4c675532-6d55-49dd-8c8e-3c1cc5624cf2"))
/* dual interface */ IConversationHistoryItemAddedEventData;
struct __declspec(uuid("464d5228-9f68-4b1f-b430-156a104e2b85"))
/* dual interface */ IPreviousHistoryRetrievalStateChangedEventData;
struct __declspec(uuid("9397ff55-ee06-4f02-8f2a-be3ae249d4bb"))
/* dual interface */ IConversationHistoryActionAvailabilityEventData;
struct __declspec(uuid("d992371e-5161-453b-97e6-6e7c67bc075e"))
/* dispinterface */ _IPreviousConversationsManagerEvents;
struct __declspec(uuid("b0c10f94-cab6-4821-9643-d781885a46af"))
/* dispinterface */ _IVoiceMailsManagerEvents;
struct __declspec(uuid("0ca9bae8-dd28-4929-a821-4d00f11203cc"))
/* dispinterface */ _IChannelEvents;
struct __declspec(uuid("57cc2235-96c5-49fa-b92b-350486c6cf52"))
/* dispinterface */ _IConversationCallback;
struct __declspec(uuid("26c8d96d-7d22-4e9e-948a-edccb4ca9c64"))
/* dispinterface */ _IParticipantCallback;
struct __declspec(uuid("3f60cff7-be3f-4404-8395-c02d0cdd318e"))
/* dispinterface */ _IModalityCallback;
struct __declspec(uuid("8684d3ad-7061-4056-a894-ebf1270adb9b"))
/* dispinterface */ _IApplicationSharingModalityCallback;
struct __declspec(uuid("5fa92ea7-6e6e-4a82-8f0d-107feaf5a75d"))
/* dispinterface */ _IConversationHistoryCallback;
struct __declspec(uuid("95b9a382-ea50-45e4-965e-ba61e7213dba"))
/* dispinterface */ _IAudioChannelCallback;
struct __declspec(uuid("08e1e9cf-a7e4-43be-ae8e-b223ce71e46e"))
/* dispinterface */ _IChannelCallback;
struct __declspec(uuid("68b1fc26-b48f-4d5f-b041-8e246676d836"))
/* dispinterface */ _IVoiceMailCallback;
struct __declspec(uuid("d9d04b55-d820-4d5b-a690-658a49368478"))
/* dispinterface */ _IVoiceMailsManagerCallback;
struct __declspec(uuid("4bacb4b8-1236-42b8-bda1-d1533148780d"))
/* dispinterface */ _IPreviousConversationsManagerCallback;
struct /* coclass */ ConversationManager;
struct /* coclass */ Conversation;
struct /* coclass */ Participant;
struct /* coclass */ InstantMessageModality;
struct /* coclass */ AVModality;
struct /* coclass */ Channel;
struct /* coclass */ Modality;
struct /* coclass */ ApplicationSharingModality;
struct /* coclass */ AudioChannel;
struct /* coclass */ IsTypingChangedEventData;
struct /* coclass */ IsContributingChangedEventData;
struct /* coclass */ ContextEventData;
struct /* coclass */ InitialContextEventData;
struct /* coclass */ ChannelStateChangedEventData;
struct /* coclass */ ChannelActionAvailabilityEventData;
struct /* coclass */ StreamStateChangedEventData;
struct /* coclass */ ConversationActionAvailabilityEventData;
struct /* coclass */ ModalityActionAvailabilityChangedEventData;
struct /* coclass */ ConversationContextCollectionEventData;
struct /* coclass */ SummarizerStatesChangedEventData;
struct /* coclass */ ConversationManagerEventData;
struct /* coclass */ ParticipantCollectionChangedEventData;
struct /* coclass */ MessageSentEventData;
struct /* coclass */ ConversationStateChangedEventData;
struct /* coclass */ ModalityStateChangedEventData;
struct /* coclass */ ConversationPropertyChangedEventData;
struct /* coclass */ ParticipantPropertyChangedEventData;
struct /* coclass */ MutedChangedEventData;
struct /* coclass */ ParticipantActionAvailabilityChangedEventData;
struct /* coclass */ ModalityPropertyChangedEventData;
struct /* coclass */ ConversationCollection;
struct /* coclass */ ConversationContext;
struct /* coclass */ ConversationContextList;
struct /* coclass */ ConversationHistory;
struct /* coclass */ ParticipantCollection;
struct /* coclass */ VideoChannel;
struct /* coclass */ VideoChannelCollection;
struct /* coclass */ VideoWindow;
struct /* coclass */ RepresentationInfo;
struct /* coclass */ ParticipantInfo;
struct /* coclass */ HistoryInfo;
struct /* coclass */ CallParkOrbit;
struct __declspec(uuid("7cc306e9-2546-4101-82ab-29613ad6f7a7"))
/* dual interface */ ICallParkOrbit;
struct /* coclass */ ConversationPropertyDictionary;
struct /* coclass */ ModalityPropertyDictionary;
struct /* coclass */ ParticipantPropertyDictionary;
struct /* coclass */ ModalityDictionary;
struct /* coclass */ ModalityStateChangePropertyDictionary;
struct /* coclass */ ConversationStateChangePropertyDictionary;
struct /* coclass */ VoiceMail;
struct /* coclass */ VoiceMailCollection;
struct /* coclass */ VoiceMailsManager;
struct /* coclass */ PreviousConversation;
struct /* coclass */ PreviousConversationCollection;
struct /* coclass */ PreviousConversationsManager;
struct /* coclass */ PreviousConversationsManagerDictionary;
struct /* coclass */ PreviousConversationsManagerConnectionStateChangedEventData;
struct /* coclass */ PreviousConversationsManagerNewItemCountChangedEventData;
struct /* coclass */ PreviousConversationBatchEventData;
struct /* coclass */ VoiceMailBatchEventData;
struct /* coclass */ ConferenceAccessInformation;
struct /* coclass */ ConversationHistoryActionAvailabilityEventData;
struct /* coclass */ ConversationHistoryItemAddedEventData;
struct /* coclass */ PreviousHistoryRetrievalStateChangedEventData;
struct /* coclass */ ConversationContextCollection;
struct /* coclass */ SharingResource;
struct /* coclass */ SharingResourceList;
struct /* coclass */ ControlRequestReceivedEventData;
struct /* coclass */ LocalSharedResourcesChangedEventData;
struct /* coclass */ ParticipationStateChangedEventData;
struct /* coclass */ ControllerChangedEventData;
struct /* coclass */ SharerChangedEventData;
enum AccessLevel;
enum AccessEntryScope;
enum PhotoSetting;
enum PublishableContactInformationType;
enum SettingItemTypeInternal;
struct __declspec(uuid("4da899d3-fac7-4f43-8e4a-1987988fbc9b"))
/* dual interface */ ISelf;
struct __declspec(uuid("7693221d-7e9f-47d3-8782-6437cf3d4c0c"))
/* dual interface */ IPhone;
struct __declspec(uuid("d0984cef-e323-4574-b07b-5c970c9caee0"))
/* dual interface */ ICustomAvailabilityState;
enum ContactAvailability;
struct __declspec(uuid("dd181e97-eee0-4d65-bad1-16866d0c7953"))
/* dual interface */ IAccessPermission;
struct __declspec(uuid("65b71572-5338-43c8-9e4c-f1dfc6711ad1"))
/* dual interface */ IAccessEntry;
struct __declspec(uuid("daa5fde3-a81c-4f23-80b6-c47b52c649bf"))
/* dual interface */ IContactEndpointCollection;
struct __declspec(uuid("c0034194-f7b6-43eb-b0e0-7852fc8e7bfa"))
/* dual interface */ IPhonesChangedEventData;
struct __declspec(uuid("dd31376e-e529-43f2-b3b4-e379ebd28e2a"))
/* dual interface */ IAccessEntryCollectionChangedEventData;
struct __declspec(uuid("dac6d77b-12f7-405c-bb33-002d6d4007f3"))
/* dual interface */ IResiliencyModeChangedEventData;
struct __declspec(uuid("3c8d7146-35ea-4133-b2f6-c1fc2401a091"))
/* dual interface */ ISelf2;
enum NotificationTypes;
enum NotificationUrgencyType;
struct __declspec(uuid("fb67275c-4d9e-4c32-9793-a96749884fc1"))
/* dual interface */ IAlertLevel;
enum AlertPrivacyType;
enum AlertModeTypes;
struct __declspec(uuid("a9aab6a0-54b9-4419-aaaf-6b26dfac1585"))
/* dispinterface */ _ISelfCallback;
struct __declspec(uuid("06437abb-c419-4b11-a474-1a2b02fbd646"))
/* dispinterface */ _ISelfEvents;
struct __declspec(uuid("dc4bc923-6a52-4f02-9f8a-547b606955e7"))
/* dispinterface */ _IAccessPermissionEvents;
struct __declspec(uuid("3c43a2be-ba4a-41e8-8f93-a9dc5f091892"))
/* dispinterface */ _ISelf2Events;
struct __declspec(uuid("90b49a60-72af-4360-9f9b-916db7504a3c"))
/* dual interface */ IAlertLevelChangedEventData;
struct /* coclass */ Self;
struct /* coclass */ AccessEntry;
struct /* coclass */ Phone;
struct /* coclass */ AccessPermission;
struct /* coclass */ ContactEndpointCollection;
struct /* coclass */ PhonesChangedEventData;
struct /* coclass */ AccessEntryCollectionChangedEventData;
struct /* coclass */ ResiliencyModeChangedEventData;
enum AudioPlayBackModes;
enum DeviceType;
struct __declspec(uuid("9274dbdc-43ce-45aa-a817-414a4494ad28"))
/* dual interface */ IDeviceManager;
struct __declspec(uuid("1f8ebe6f-7993-42ce-980a-a2f03793be71"))
/* dual interface */ IDeviceCollection;
struct __declspec(uuid("ce16d540-b9e8-4742-8659-9a7a318afb92"))
/* dual interface */ IDevice;
struct __declspec(uuid("b0b0e8f8-32da-4cb0-a124-d1e35891d53a"))
/* dual interface */ IAudioDevice;
struct __declspec(uuid("da7cf784-67df-429c-82fb-9afd99f6cccb"))
/* dual interface */ IVideoDevice;
struct __declspec(uuid("08ddf5c4-fa12-4978-b26e-c6d23c453413"))
/* dispinterface */ _IDeviceManagerCallback;
struct /* coclass */ DeviceManager;
struct /* coclass */ DeviceCollection;
struct /* coclass */ Device;
struct /* coclass */ AudioDevice;
struct /* coclass */ VideoDevice;
enum ConferenceInfoProperty;
enum ConferenceSetting;
enum ConferenceAdmissionType;
enum ConferenceAutoPromoteType;
struct __declspec(uuid("b3c5872e-591c-4a74-9f84-b00ef970a838"))
/* dual interface */ IConferenceInfo;
struct __declspec(uuid("2d3d62ca-e9f6-4f94-8ebd-5fabab29e6a5"))
/* dual interface */ IConferenceInfoPropertyDictionary;
struct __declspec(uuid("bd891697-c93e-4a0b-9b6c-004bd81b6ee8"))
/* dual interface */ IConferenceSettingDictionary;
struct __declspec(uuid("c7acb102-b692-49cc-92da-5824822c7b96"))
/* dual interface */ IConferenceScheduler;
struct __declspec(uuid("1ff8aa01-d8af-4b2b-b254-befefabaf2d6"))
/* dual interface */ IConferenceInfoCollection;
struct __declspec(uuid("e0e8b413-bb57-4b58-b1fb-288f3a269e35"))
/* dual interface */ IConferenceCapabilities;
struct __declspec(uuid("66d1ffc0-c234-4f3d-abaa-98222a69e525"))
/* dual interface */ IConferenceTelephoneAccessInformation;
struct __declspec(uuid("a0984f01-5d2c-4302-87a0-69be7b015143"))
/* dual interface */ IConferenceRegionalAccessInformationCollection;
struct __declspec(uuid("420a24e2-5c31-4262-9bd5-058682300ed6"))
/* dual interface */ IConferenceRegionalAccessInformation;
struct __declspec(uuid("4cb70415-ba71-491e-b12d-155d85cbdff1"))
/* dual interface */ IConferenceRegionalAccessNumberCollection;
struct __declspec(uuid("9244d573-914f-4c1f-93f6-31609a95cbed"))
/* dual interface */ IConferenceRegionalAccessNumber;
struct __declspec(uuid("e5cff4a5-83a8-4fd8-8daa-77a3e2288d4a"))
/* dual interface */ IConferenceCapabilities2;
struct __declspec(uuid("c22ce6bf-3a47-46b7-8544-44305badfef9"))
/* dispinterface */ _IConferenceSchedulerCallback;
struct /* coclass */ ConferenceScheduler;
struct /* coclass */ ConferenceInfo;
struct /* coclass */ ConferenceInfoPropertyDictionary;
struct /* coclass */ ConferenceApplicationData;
struct __declspec(uuid("be6086a1-d436-4834-89ee-3ca4f8596a58"))
/* dual interface */ IConferenceApplicationData;
struct __declspec(uuid("64b7e9ad-d920-4eaa-95c7-9d95de9ace24"))
/* dual interface */ IConferenceApplicationProperties;
struct /* coclass */ ConferenceApplicationProperties;
struct /* coclass */ ConferenceSettingDictionary;
struct /* coclass */ ConferenceCapabilities;
struct /* coclass */ ConferenceInfoCollection;
struct /* coclass */ ConferenceTelephoneAccessInformation;
struct /* coclass */ ConferenceRegionalAccessInformation;
struct /* coclass */ ConferenceRegionalAccessInformationCollection;
struct /* coclass */ ConferenceRegionalAccessNumberCollection;
struct /* coclass */ ConferenceRegionalAccessNumber;
enum UtilitiesAction;
struct __declspec(uuid("c4d95336-5949-4d13-aac8-82039e5cbcfc"))
/* dispinterface */ _IUtilitiesCallback;
struct __declspec(uuid("87084018-7e74-4f92-ab39-1cf03188580e"))
/* dual interface */ IUtilities;
struct __declspec(uuid("9112ce9b-3704-48c0-86dc-7a7b0f34d7a3"))
/* dual interface */ IUtilities2;
struct __declspec(uuid("26f15083-86db-4663-89ed-a0abc896b9a1"))
/* dual interface */ IScheduledConferenceManager;
struct __declspec(uuid("3417d640-3976-4cb6-bc82-54e5e2e10784"))
/* dispinterface */ _IUtilities2Events;
struct __declspec(uuid("7818a3f7-674b-477e-99e8-54774579a8e7"))
/* dual interface */ IExchangeECPUrlChangedEventData;
//struct /* coclass */ Utilities;
struct __declspec(uuid("71a054b3-4d4b-4643-b14f-4572494b19e7"))
/* dual interface */ IScheduledConference;
struct __declspec(uuid("f423726d-0e9b-4b55-9569-e79865210f69"))
/* dual interface */ IScheduledConferenceCollection;
struct __declspec(uuid("219dc126-9a48-483f-80c2-3f22b3b47829"))
/* dispinterface */ _IScheduledConferenceManagerCallback;
struct /* coclass */ Utilities2;
struct /* coclass */ ExchangeECPUrlChangedEventData;
struct /* coclass */ ScheduledConference;
struct /* coclass */ ScheduledConferenceCollection;
struct /* coclass */ ScheduledConferenceManager;
enum RoomMessageType;
enum RoomMessageFormat;
enum RoomMessageFilteringAction;
enum RoomSearchModeType;
enum RoomType;
enum RoomProperty;
enum RoomJoinState;
enum RoomManagerState;
struct __declspec(uuid("a0c3e9ba-c622-45a2-9fed-2c662f3f489a"))
/* dual interface */ IRoomMessage;
struct __declspec(uuid("603e2454-725a-4679-a6a8-363b21633cb2"))
/* dual interface */ IRoomMessageDictionary;
struct __declspec(uuid("21c41ede-7ed8-4d97-9b0d-dd785781af9b"))
/* dual interface */ IRoomMessageCollection;
struct __declspec(uuid("68da8699-6c6e-4390-bd82-709ff1dc1702"))
/* dual interface */ IRoomUser;
struct __declspec(uuid("a766520d-f700-4bf5-8428-56bd914c1c40"))
/* dual interface */ IRoom;
struct __declspec(uuid("7bf20b14-58d1-494b-b301-9b16bacc9610"))
/* dual interface */ IRoomPropertyDictionary;
struct __declspec(uuid("05789509-afa1-4e90-b4fa-1a7820db9796"))
/* dual interface */ IRoomCollection;
struct __declspec(uuid("76de1cd8-f278-48a1-a391-e5beb8f428e6"))
/* dual interface */ IRoomManager;
struct __declspec(uuid("b03614ae-8d68-4386-9e1c-939cabcf1232"))
/* dual interface */ IRoomPropertyChangedEventData;
struct __declspec(uuid("b8fd4a53-e7e6-4995-a5b5-1306c7584964"))
/* dual interface */ IUnreadMessageCountChangedEventData;
struct __declspec(uuid("1566fa3c-fd24-41db-a2b8-6232f89783de"))
/* dual interface */ IRoomMessageEventData;
struct __declspec(uuid("081793da-7221-45c3-8afa-3a2e6e464670"))
/* dual interface */ IRoomMessagesEventData;
struct __declspec(uuid("bc9f5f40-d315-40c1-b1c6-cea0646e18e2"))
/* dual interface */ IRoomUserChangedEventData;
struct __declspec(uuid("7452bd0f-65cb-4a5e-ac37-e2bea1f43dd9"))
/* dual interface */ IRoomParticipantsEventData;
struct __declspec(uuid("4d120020-ce64-43c5-9f84-7a7b2360388f"))
/* dual interface */ IRoomJoinStateChangedEventData;
struct __declspec(uuid("300e56a3-ce08-4eb3-9f48-505aa162c9da"))
/* dispinterface */ _IRoomEvents;
struct __declspec(uuid("e2c6d6e0-52ef-47a2-bda9-9d6a7a29a391"))
/* dispinterface */ _IRoomManagerEvents;
struct __declspec(uuid("8f04e92d-ee26-4b64-b9c1-56dd90816810"))
/* dual interface */ IFollowedRoomsChangedEventData;
struct __declspec(uuid("89c7eb8e-42bc-4c9f-bb34-88cde83ccb37"))
/* dual interface */ IRoomManagerStateChangedEventData;
struct __declspec(uuid("8e73e7e1-90ce-4d3b-b5b7-71e87b2e8ba5"))
/* dispinterface */ _IRoomUserEvents;
struct __declspec(uuid("628d7bdd-45ab-404d-8505-b3adc15f454d"))
/* dispinterface */ _IRoomCallback;
struct __declspec(uuid("0fdb5b86-dd44-41bf-a443-317fd19bbf3d"))
/* dispinterface */ _IRoomManagerCallback;
struct /* coclass */ RoomMessage;
struct /* coclass */ RoomUser;
struct /* coclass */ Room;
struct /* coclass */ RoomManager;
struct /* coclass */ RoomMessageCollection;
struct /* coclass */ RoomCollection;
struct /* coclass */ RoomPropertyChangedEventData;
struct /* coclass */ UnreadMessageCountChangedEventData;
struct /* coclass */ RoomMessageEventData;
struct /* coclass */ RoomMessagesEventData;
struct /* coclass */ FollowedRoomsChangedEventData;
struct /* coclass */ RoomManagerStateChangedEventData;
struct /* coclass */ RoomUserChangedEventData;
struct /* coclass */ RoomParticipantsEventData;
struct /* coclass */ RoomJoinStateChangedEventData;
struct /* coclass */ RoomPropertyDictionary;
struct /* coclass */ RoomMessageDictionary;
enum AutomationModalities;
enum AutomationModalitySettings;
enum ConversationWindowState;
enum ConversationWindowInformationType;
enum ConversationWindowAction;
struct __declspec(uuid("2edaf003-14a6-4faf-8a72-99526d048d63"))
/* dual interface */ IAutomation;
struct __declspec(uuid("b1f5bcd8-69a6-4fc3-9ef0-9bd4ad999061"))
/* dual interface */ IConversationWindow;
struct __declspec(uuid("ca62e6f3-3e6d-451b-aeaa-6a1b7aacf8d4"))
/* dual interface */ IConversationWindowNeedsSizeChangeEventData;
struct __declspec(uuid("f97bed54-e434-4020-a197-f15aea9d9c95"))
/* dual interface */ IConversationWindowNeedsAttentionEventData;
struct __declspec(uuid("806d3227-4cb8-47c4-9864-7d4df4f44069"))
/* dispinterface */ _IAutomationCallback;
struct __declspec(uuid("9e0f8fca-d9c9-47cd-87f4-5554217dffcc"))
/* dispinterface */ _IConversationWindowEvents;
struct __declspec(uuid("76428270-36fa-4236-8bdf-aadb39fd1371"))
/* dispinterface */ _IConversationWindowCallback;
struct __declspec(uuid("02dd541f-c863-4295-92ab-a4cf29f5d841"))
/* dual interface */ IConversationWindow2;
struct __declspec(uuid("109d32ce-e09f-46f5-9ecb-01e1d273b7da"))
/* dual interface */ IConversationWindowInformationDictionary;
struct __declspec(uuid("663c61f4-a9da-4a6f-a606-f142f665df61"))
/* dual interface */ IConversationWindowActions;
enum AnnotationTool;
struct __declspec(uuid("55a21cf3-a2b3-484e-be2a-14280f501289"))
/* dual interface */ IVideoView;
struct __declspec(uuid("b00f2520-029c-47d0-b4e8-8fbef47caa7e"))
/* dual interface */ IVideoViewInformationDictionary;
enum VideoViewInformationType;
enum VideoViewMode;
struct __declspec(uuid("af2aba54-f591-42ba-8004-20944c9fb4d7"))
/* dispinterface */ _IConversationWindow2Events;
struct __declspec(uuid("2209e091-5b90-4caf-8582-cc5979883376"))
/* dual interface */ IConversationWindowStateChangedEventData;
struct __declspec(uuid("ff037274-455a-4e34-b5d1-d42db866f9b7"))
/* dual interface */ IConversationWindowInformationChangedEventData;
struct __declspec(uuid("0678c83e-f580-4d99-902f-930699b28be8"))
/* dual interface */ IConversationWindowActionAvailabilityChangedEventData;
struct /* coclass */ ConversationWindow;
struct /* coclass */ Automation;
struct /* coclass */ ConversationWindowNeedsSizeChangeEventData;
struct /* coclass */ ConversationWindowNeedsAttentionEventData;
enum ShareableContentAction;
enum ShareableContentProperty;
enum ShareableContentState;
enum ShareableContentType;
struct __declspec(uuid("38e5ec21-81a1-45a7-94d6-b812976231a4"))
/* dual interface */ IContentSharingModality;
struct __declspec(uuid("a58f54d2-9786-4309-964d-96549aec7611"))
/* dual interface */ IShareableContentCollection;
struct __declspec(uuid("b335ae5e-e4ea-49d3-b03b-646a96fe66d3"))
/* dual interface */ IShareableContent;
struct __declspec(uuid("ca478b44-4fce-4082-8700-8cec69fb50e1"))
/* dual interface */ IShareableContentPropertyDictionary;
enum ContentSavingFileType;
struct __declspec(uuid("ae3a7c39-9c30-4e0f-bae4-d8344ef377ea"))
/* dispinterface */ _IContentSharingModalityEvents;
struct __declspec(uuid("17635563-2fa5-46ad-8f47-3f0ff958ece2"))
/* dual interface */ IContentCollectionChangedEventData;
struct __declspec(uuid("48b3099d-841d-4caa-9202-5787596e2bd2"))
/* dual interface */ IActivePresenterChangedEventData;
struct __declspec(uuid("899b846a-8887-45d6-b8f5-8aaa0bfdf722"))
/* dual interface */ IActiveContentChangedEventData;
struct __declspec(uuid("6b551c9e-de81-41df-a0ae-39f0af11d508"))
/* dispinterface */ _IShareableContentEvents;
struct __declspec(uuid("94093301-8854-408e-8138-c7c3d2daca1f"))
/* dual interface */ IShareableContentStateChangedEventData;
struct __declspec(uuid("3c8210c8-8578-47c6-87a9-fa1ad2ba9873"))
/* dual interface */ IShareableContentPropertyChangedEventData;
struct __declspec(uuid("4023052e-e3b7-48a3-bcb3-ca2234d9baee"))
/* dual interface */ IShareableContentActionAvailabilityChangedEventData;
struct __declspec(uuid("6964b116-8fe6-4c19-9052-0eaa74460a58"))
/* dispinterface */ _IConversation2Callback;
struct __declspec(uuid("84340964-7820-4ebc-bcd3-702926de23e8"))
/* dispinterface */ _IContentSharingModalityCallback;
struct __declspec(uuid("3187885c-a8af-4d4b-8e80-f4b9d447a7b2"))
/* dispinterface */ _IShareableContentCallback;
struct __declspec(uuid("36cdc166-4f21-46ad-a60e-8551f26c1d41"))
/* dispinterface */ _IParticipant2Callback;
struct __declspec(uuid("70986dce-4aa8-4dfb-9a45-4bfcf7a206f5"))
/* dispinterface */ _IChannel2Callback;
struct /* coclass */ Conversation2;
struct /* coclass */ Participant2;
struct /* coclass */ Channel2;
struct /* coclass */ ConferenceAccessInformation2;
struct __declspec(uuid("5e7ab04c-e4b8-49fb-85ff-9e2bca2899d3"))
/* dual interface */ IConferenceAccessInformation2;
struct /* coclass */ ContentSharingModality;
struct /* coclass */ ShareableContentPropertyDictionary;
struct /* coclass */ ShareableContent;
struct /* coclass */ PowerPointContent;
struct __declspec(uuid("aafe380b-b079-4b55-b81b-a9401d77b776"))
/* dual interface */ IPowerPointContent;
struct /* coclass */ ShareableContentCollection;
struct /* coclass */ ContentCollectionChangedEventData;
struct /* coclass */ ActiveContentChangedEventData;
struct /* coclass */ ActivePresenterChangedEventData;
struct /* coclass */ ShareableContentPropertyChangedEventData;
struct /* coclass */ ShareableContentStateChangedEventData;
struct /* coclass */ ShareableContentActionAvailabilityChangedEventData;
struct /* coclass */ Self2;
struct /* coclass */ AlertLevel;
struct /* coclass */ AlertLevelChangedEventData;
struct __declspec(uuid("086f0e9d-a416-40f9-877c-f1c82dc1a6ad"))
/* dual interface */ IVideoViewInformationChangedEventData;
struct __declspec(uuid("30f974fb-1f7b-49d3-8da4-01694ba474d2"))
/* dispinterface */ _IVideoViewEvents;
struct /* coclass */ ConversationWindow2;
struct /* coclass */ ConversationWindowInformationDictionary;
struct /* coclass */ ConversationWindowStateChangedEventData;
struct /* coclass */ ConversationWindowInformationChangedEventData;
struct /* coclass */ ConversationWindowActionAvailabilityChangedEventData;
struct /* coclass */ ConversationWindowActions;
struct /* coclass */ VideoView;
struct /* coclass */ VideoViewInformationChangedEventData;
struct /* coclass */ VideoViewInformationDictionary;
struct /* coclass */ Contact2;
struct /* coclass */ SearchResults2;
enum ClientSettingsType;
struct __declspec(uuid("86225e5c-7595-4d1c-985d-8a0458d714c0"))
/* dual interface */ IClient2;
struct __declspec(uuid("1b04d929-11e4-46c2-9162-ea4f2758c3a7"))
/* dual interface */ IClientSettings;
struct __declspec(uuid("25d64aea-0e65-49cb-8d6d-65db0ac1af65"))
/* dual interface */ IClientSettingsChangedEventData;
struct __declspec(uuid("74b6caf0-755e-454d-a819-d0e439ef5296"))
/* dual interface */ IPublicCloudProviderInfo;
struct __declspec(uuid("569e04f2-d286-4903-9642-44d850310956"))
/* dispinterface */ _IClientSettingsEvents;
struct /* coclass */ Client2;
struct /* coclass */ ClientSettings;
struct /* coclass */ ClientSettingsChangedEventData;
struct /* coclass */ PublicCloudProviderInfo;
struct /* coclass */ ConferenceCapabilities2;
enum ClientState;
enum CredentialRequestedType;
enum LyncClientNotification;
enum MapiStatus;
enum LyncClientConfigurationMode;
enum TransportMode;
enum SignInConfigurationType;
enum SupportedFeatures;
enum ClientType;
enum LyncClientCapabilityTypes;
enum OIFeature;
enum OIInterface;
enum OIHelperAction;
struct __declspec(uuid("6a222195-f65e-467f-8f77-eb180bd85288"))
/* dual interface */ IUCOfficeIntegration;
struct __declspec(uuid("ee9f3e74-ac61-469e-80d6-e22bf4eeff5c"))
/* dual interface */ IClient;
struct __declspec(uuid("c413a27f-601e-4261-b9e3-0eb966f4e7d1"))
/* dual interface */ ILyncClient;
struct __declspec(uuid("f8df0b4b-a4ce-476b-81d3-9fba98461b62"))
/* dual interface */ ISignInConfiguration;
struct __declspec(uuid("80edbf3a-812e-42b5-a67f-6cc6d9a19a6c"))
/* dual interface */ IDelegatorClient;
struct __declspec(uuid("aa41ee75-f9c5-42f7-8d17-1d92d6ba4d92"))
/* dual interface */ IApplicationRegistration;
struct __declspec(uuid("330639f2-e399-48cb-863f-56a50a27a138"))
/* dual interface */ IWebAppExtension;
struct __declspec(uuid("9b3a702d-5be1-4fce-adb3-ebbd23e078c3"))
/* dual interface */ IClientStateChangedEventData;
struct __declspec(uuid("9ab20314-b258-48f8-b659-ad250dfa20f5"))
/* dual interface */ ILyncClientNotificationReceivedEventData;
struct __declspec(uuid("f0b791fa-434b-49f2-b84e-05619ebe6beb"))
/* dual interface */ ICredentialRequestedEventData;
enum CredentialRequestResult;
struct __declspec(uuid("e664682c-cba5-4c7c-a65b-9cd046f24cbd"))
/* dual interface */ ISignInDelayedEventData;
struct __declspec(uuid("b28f2007-3926-49cf-9101-4ccb9574049b"))
/* dual interface */ IPreferredCapabilitiesChangedEventData;
struct __declspec(uuid("744bab86-b15f-4bd5-b7dd-476b0ce0162a"))
/* dual interface */ IDelegatorClientCollectionEventData;
struct __declspec(uuid("09028936-cd20-43ef-98cf-56726addba34"))
/* interface */ IUCOfficeIntegrationHelper;
struct __declspec(uuid("b26f6403-761e-48d4-84e2-619c17eaeb13"))
/* dispinterface */ _ILyncClientCallback;
struct __declspec(uuid("64561371-77a2-4bd4-a909-e6deb1b5c6cf"))
/* dispinterface */ _IClientEvents;
struct __declspec(uuid("953ae732-f53a-4116-ac1b-0321b3fb3dba"))
/* dispinterface */ _ILyncClientEvents;
struct __declspec(uuid("92a9eb9b-85ee-444e-a302-00c3c945baa4"))
/* dispinterface */ _IDelegatorClientEvents;
struct __declspec(uuid("8c9b6953-33cc-4a04-9da4-f71ae79da0c0"))
/* dispinterface */ _IUCOfficeIntegrationEvents;
struct /* coclass */ Client;
struct /* coclass */ LyncClient;
struct /* coclass */ SignInConfiguration;
struct /* coclass */ ApplicationRegistration;
struct /* coclass */ CustomAvailabilityState;
struct /* coclass */ DelegatorClient;
struct /* coclass */ ClientStateChangedEventData;
struct /* coclass */ LyncClientNotificationReceivedEventData;
struct /* coclass */ CredentialRequestedEventData;
struct /* coclass */ SignInDelayedEventData;
struct /* coclass */ PreferredCapabilitiesChangedEventData;
struct /* coclass */ DelegatorClientCollectionEventData;
struct /* coclass */ UCOfficeIntegration;

//
// Smart pointer typedef declarations
//

_COM_SMARTPTR_TYPEDEF(IAsynchronousOperation, __uuidof(IAsynchronousOperation));
_COM_SMARTPTR_TYPEDEF(ILocaleString, __uuidof(ILocaleString));
_COM_SMARTPTR_TYPEDEF(IAddress, __uuidof(IAddress));
_COM_SMARTPTR_TYPEDEF(INamedPropertyList, __uuidof(INamedPropertyList));
_COM_SMARTPTR_TYPEDEF(IContactInformationDictionary, __uuidof(IContactInformationDictionary));
_COM_SMARTPTR_TYPEDEF(IContactSettingDictionary, __uuidof(IContactSettingDictionary));
_COM_SMARTPTR_TYPEDEF(IContactEndpoint, __uuidof(IContactEndpoint));
_COM_SMARTPTR_TYPEDEF(IContactInformationChangedEventData, __uuidof(IContactInformationChangedEventData));
_COM_SMARTPTR_TYPEDEF(IDistributionGroupExpansionStateChangedEventData, __uuidof(IDistributionGroupExpansionStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IUriChangedEventData, __uuidof(IUriChangedEventData));
_COM_SMARTPTR_TYPEDEF(IGroupNameChangedEventData, __uuidof(IGroupNameChangedEventData));
_COM_SMARTPTR_TYPEDEF(ISearchResult, __uuidof(ISearchResult));
_COM_SMARTPTR_TYPEDEF(ISearchResults, __uuidof(ISearchResults));
_COM_SMARTPTR_TYPEDEF(ISearchProviderStateChangedEventData, __uuidof(ISearchProviderStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IWorkingPeriod, __uuidof(IWorkingPeriod));
_COM_SMARTPTR_TYPEDEF(IWorkingHours, __uuidof(IWorkingHours));
_COM_SMARTPTR_TYPEDEF(IPresenceCapability, __uuidof(IPresenceCapability));
_COM_SMARTPTR_TYPEDEF(IDescriptionChangedEventData, __uuidof(IDescriptionChangedEventData));
_COM_SMARTPTR_TYPEDEF(ISearchResults2, __uuidof(ISearchResults2));
_COM_SMARTPTR_TYPEDEF(_IContactsAndGroupsCallback, __uuidof(_IContactsAndGroupsCallback));
_COM_SMARTPTR_TYPEDEF(_IGroupCallback, __uuidof(_IGroupCallback));
_COM_SMARTPTR_TYPEDEF(_IContactCallback, __uuidof(_IContactCallback));
_COM_SMARTPTR_TYPEDEF(_IContactManagerEvents, __uuidof(_IContactManagerEvents));
_COM_SMARTPTR_TYPEDEF(_IContactEvents, __uuidof(_IContactEvents));
_COM_SMARTPTR_TYPEDEF(IContactSettingChangedEventData, __uuidof(IContactSettingChangedEventData));
_COM_SMARTPTR_TYPEDEF(_IGroupEvents, __uuidof(_IGroupEvents));
_COM_SMARTPTR_TYPEDEF(_ICustomGroupEvents, __uuidof(_ICustomGroupEvents));
_COM_SMARTPTR_TYPEDEF(_IRecentContactsEvents, __uuidof(_IRecentContactsEvents));
_COM_SMARTPTR_TYPEDEF(_IDistributionGroupEvents, __uuidof(_IDistributionGroupEvents));
_COM_SMARTPTR_TYPEDEF(_IFrequentContactsEvents, __uuidof(_IFrequentContactsEvents));
_COM_SMARTPTR_TYPEDEF(IContactPositionChangedEventData, __uuidof(IContactPositionChangedEventData));
_COM_SMARTPTR_TYPEDEF(_IFavoriteContactsEvents, __uuidof(_IFavoriteContactsEvents));
_COM_SMARTPTR_TYPEDEF(IModalityPropertyDictionary, __uuidof(IModalityPropertyDictionary));
_COM_SMARTPTR_TYPEDEF(IParticipantPropertyDictionary, __uuidof(IParticipantPropertyDictionary));
_COM_SMARTPTR_TYPEDEF(IParticipantInfo, __uuidof(IParticipantInfo));
_COM_SMARTPTR_TYPEDEF(IConversationPropertyDictionary, __uuidof(IConversationPropertyDictionary));
_COM_SMARTPTR_TYPEDEF(IVoiceMailsManager, __uuidof(IVoiceMailsManager));
_COM_SMARTPTR_TYPEDEF(IPreviousConversationsManager, __uuidof(IPreviousConversationsManager));
_COM_SMARTPTR_TYPEDEF(IPreviousConversationsManagerDictionary, __uuidof(IPreviousConversationsManagerDictionary));
_COM_SMARTPTR_TYPEDEF(IRepresentationInfo, __uuidof(IRepresentationInfo));
_COM_SMARTPTR_TYPEDEF(IVideoWindow, __uuidof(IVideoWindow));
_COM_SMARTPTR_TYPEDEF(ISharingResource, __uuidof(ISharingResource));
_COM_SMARTPTR_TYPEDEF(ISharingResourceList, __uuidof(ISharingResourceList));
_COM_SMARTPTR_TYPEDEF(IChannel2, __uuidof(IChannel2));
_COM_SMARTPTR_TYPEDEF(IModalityStateChangePropertyDictionary, __uuidof(IModalityStateChangePropertyDictionary));
_COM_SMARTPTR_TYPEDEF(IModalityStateChangedEventData, __uuidof(IModalityStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IModalityActionAvailabilityChangedEventData, __uuidof(IModalityActionAvailabilityChangedEventData));
_COM_SMARTPTR_TYPEDEF(IConversationStateChangePropertyDictionary, __uuidof(IConversationStateChangePropertyDictionary));
_COM_SMARTPTR_TYPEDEF(IConversationStateChangedEventData, __uuidof(IConversationStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IConversationActionAvailabilityEventData, __uuidof(IConversationActionAvailabilityEventData));
_COM_SMARTPTR_TYPEDEF(IMessageSentEventData, __uuidof(IMessageSentEventData));
_COM_SMARTPTR_TYPEDEF(IIsTypingChangedEventData, __uuidof(IIsTypingChangedEventData));
_COM_SMARTPTR_TYPEDEF(IIsContributingChangedEventData, __uuidof(IIsContributingChangedEventData));
_COM_SMARTPTR_TYPEDEF(IChannelStateChangedEventData, __uuidof(IChannelStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IChannelActionAvailabilityEventData, __uuidof(IChannelActionAvailabilityEventData));
_COM_SMARTPTR_TYPEDEF(IHistoryInfo, __uuidof(IHistoryInfo));
_COM_SMARTPTR_TYPEDEF(IConversationPropertyChangedEventData, __uuidof(IConversationPropertyChangedEventData));
_COM_SMARTPTR_TYPEDEF(IParticipantPropertyChangedEventData, __uuidof(IParticipantPropertyChangedEventData));
_COM_SMARTPTR_TYPEDEF(IModalityPropertyChangedEventData, __uuidof(IModalityPropertyChangedEventData));
_COM_SMARTPTR_TYPEDEF(IPreviousConversation, __uuidof(IPreviousConversation));
_COM_SMARTPTR_TYPEDEF(IPreviousConversationCollection, __uuidof(IPreviousConversationCollection));
_COM_SMARTPTR_TYPEDEF(IPreviousConversationsManagerConnectionStateChangedEventData, __uuidof(IPreviousConversationsManagerConnectionStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IPreviousConversationsManagerNewItemCountChangedEventData, __uuidof(IPreviousConversationsManagerNewItemCountChangedEventData));
_COM_SMARTPTR_TYPEDEF(IPreviousConversationBatchEventData, __uuidof(IPreviousConversationBatchEventData));
_COM_SMARTPTR_TYPEDEF(IConferenceAccessInformation, __uuidof(IConferenceAccessInformation));
_COM_SMARTPTR_TYPEDEF(IParticipantActionAvailabilityChangedEventData, __uuidof(IParticipantActionAvailabilityChangedEventData));
_COM_SMARTPTR_TYPEDEF(IConversation2, __uuidof(IConversation2));
_COM_SMARTPTR_TYPEDEF(IParticipant2, __uuidof(IParticipant2));
_COM_SMARTPTR_TYPEDEF(_IConversationManagerEvents, __uuidof(_IConversationManagerEvents));
_COM_SMARTPTR_TYPEDEF(_IConversationEvents, __uuidof(_IConversationEvents));
_COM_SMARTPTR_TYPEDEF(IInitialContextEventData, __uuidof(IInitialContextEventData));
_COM_SMARTPTR_TYPEDEF(IContextEventData, __uuidof(IContextEventData));
_COM_SMARTPTR_TYPEDEF(_IParticipantEvents, __uuidof(_IParticipantEvents));
_COM_SMARTPTR_TYPEDEF(IMutedChangedEventData, __uuidof(IMutedChangedEventData));
_COM_SMARTPTR_TYPEDEF(_IModalityEvents, __uuidof(_IModalityEvents));
_COM_SMARTPTR_TYPEDEF(_IInstantMessageModalityEvents, __uuidof(_IInstantMessageModalityEvents));
_COM_SMARTPTR_TYPEDEF(_IAVModalityEvents, __uuidof(_IAVModalityEvents));
_COM_SMARTPTR_TYPEDEF(_IApplicationSharingModalityEvents, __uuidof(_IApplicationSharingModalityEvents));
_COM_SMARTPTR_TYPEDEF(ILocalSharedResourcesChangedEventData, __uuidof(ILocalSharedResourcesChangedEventData));
_COM_SMARTPTR_TYPEDEF(IParticipationStateChangedEventData, __uuidof(IParticipationStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IControllerChangedEventData, __uuidof(IControllerChangedEventData));
_COM_SMARTPTR_TYPEDEF(ISharerChangedEventData, __uuidof(ISharerChangedEventData));
_COM_SMARTPTR_TYPEDEF(_IConversationHistoryEvents, __uuidof(_IConversationHistoryEvents));
_COM_SMARTPTR_TYPEDEF(IConversationHistoryItemAddedEventData, __uuidof(IConversationHistoryItemAddedEventData));
_COM_SMARTPTR_TYPEDEF(IPreviousHistoryRetrievalStateChangedEventData, __uuidof(IPreviousHistoryRetrievalStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IConversationHistoryActionAvailabilityEventData, __uuidof(IConversationHistoryActionAvailabilityEventData));
_COM_SMARTPTR_TYPEDEF(_IPreviousConversationsManagerEvents, __uuidof(_IPreviousConversationsManagerEvents));
_COM_SMARTPTR_TYPEDEF(_IVoiceMailsManagerEvents, __uuidof(_IVoiceMailsManagerEvents));
_COM_SMARTPTR_TYPEDEF(_IChannelEvents, __uuidof(_IChannelEvents));
_COM_SMARTPTR_TYPEDEF(_IConversationCallback, __uuidof(_IConversationCallback));
_COM_SMARTPTR_TYPEDEF(_IParticipantCallback, __uuidof(_IParticipantCallback));
_COM_SMARTPTR_TYPEDEF(_IModalityCallback, __uuidof(_IModalityCallback));
_COM_SMARTPTR_TYPEDEF(_IApplicationSharingModalityCallback, __uuidof(_IApplicationSharingModalityCallback));
_COM_SMARTPTR_TYPEDEF(_IConversationHistoryCallback, __uuidof(_IConversationHistoryCallback));
_COM_SMARTPTR_TYPEDEF(_IAudioChannelCallback, __uuidof(_IAudioChannelCallback));
_COM_SMARTPTR_TYPEDEF(_IChannelCallback, __uuidof(_IChannelCallback));
_COM_SMARTPTR_TYPEDEF(_IVoiceMailCallback, __uuidof(_IVoiceMailCallback));
_COM_SMARTPTR_TYPEDEF(_IVoiceMailsManagerCallback, __uuidof(_IVoiceMailsManagerCallback));
_COM_SMARTPTR_TYPEDEF(_IPreviousConversationsManagerCallback, __uuidof(_IPreviousConversationsManagerCallback));
_COM_SMARTPTR_TYPEDEF(ICallParkOrbit, __uuidof(ICallParkOrbit));
_COM_SMARTPTR_TYPEDEF(IPhone, __uuidof(IPhone));
_COM_SMARTPTR_TYPEDEF(ICustomAvailabilityState, __uuidof(ICustomAvailabilityState));
_COM_SMARTPTR_TYPEDEF(IContactEndpointCollection, __uuidof(IContactEndpointCollection));
_COM_SMARTPTR_TYPEDEF(IPhonesChangedEventData, __uuidof(IPhonesChangedEventData));
_COM_SMARTPTR_TYPEDEF(IResiliencyModeChangedEventData, __uuidof(IResiliencyModeChangedEventData));
_COM_SMARTPTR_TYPEDEF(IAlertLevel, __uuidof(IAlertLevel));
_COM_SMARTPTR_TYPEDEF(ISelf2, __uuidof(ISelf2));
_COM_SMARTPTR_TYPEDEF(_ISelfCallback, __uuidof(_ISelfCallback));
_COM_SMARTPTR_TYPEDEF(_ISelfEvents, __uuidof(_ISelfEvents));
_COM_SMARTPTR_TYPEDEF(_IAccessPermissionEvents, __uuidof(_IAccessPermissionEvents));
_COM_SMARTPTR_TYPEDEF(_ISelf2Events, __uuidof(_ISelf2Events));
_COM_SMARTPTR_TYPEDEF(IAlertLevelChangedEventData, __uuidof(IAlertLevelChangedEventData));
_COM_SMARTPTR_TYPEDEF(IDevice, __uuidof(IDevice));
_COM_SMARTPTR_TYPEDEF(IDeviceCollection, __uuidof(IDeviceCollection));
_COM_SMARTPTR_TYPEDEF(IAudioDevice, __uuidof(IAudioDevice));
_COM_SMARTPTR_TYPEDEF(IVideoDevice, __uuidof(IVideoDevice));
_COM_SMARTPTR_TYPEDEF(IDeviceManager, __uuidof(IDeviceManager));
_COM_SMARTPTR_TYPEDEF(_IDeviceManagerCallback, __uuidof(_IDeviceManagerCallback));
_COM_SMARTPTR_TYPEDEF(IConferenceInfoPropertyDictionary, __uuidof(IConferenceInfoPropertyDictionary));
_COM_SMARTPTR_TYPEDEF(IConferenceInfo, __uuidof(IConferenceInfo));
_COM_SMARTPTR_TYPEDEF(IConferenceSettingDictionary, __uuidof(IConferenceSettingDictionary));
_COM_SMARTPTR_TYPEDEF(IConferenceScheduler, __uuidof(IConferenceScheduler));
_COM_SMARTPTR_TYPEDEF(IConferenceInfoCollection, __uuidof(IConferenceInfoCollection));
_COM_SMARTPTR_TYPEDEF(IConferenceRegionalAccessNumber, __uuidof(IConferenceRegionalAccessNumber));
_COM_SMARTPTR_TYPEDEF(IConferenceRegionalAccessNumberCollection, __uuidof(IConferenceRegionalAccessNumberCollection));
_COM_SMARTPTR_TYPEDEF(IConferenceRegionalAccessInformation, __uuidof(IConferenceRegionalAccessInformation));
_COM_SMARTPTR_TYPEDEF(IConferenceRegionalAccessInformationCollection, __uuidof(IConferenceRegionalAccessInformationCollection));
_COM_SMARTPTR_TYPEDEF(IConferenceTelephoneAccessInformation, __uuidof(IConferenceTelephoneAccessInformation));
_COM_SMARTPTR_TYPEDEF(IConferenceCapabilities, __uuidof(IConferenceCapabilities));
_COM_SMARTPTR_TYPEDEF(IConferenceCapabilities2, __uuidof(IConferenceCapabilities2));
_COM_SMARTPTR_TYPEDEF(_IConferenceSchedulerCallback, __uuidof(_IConferenceSchedulerCallback));
_COM_SMARTPTR_TYPEDEF(IConferenceApplicationProperties, __uuidof(IConferenceApplicationProperties));
_COM_SMARTPTR_TYPEDEF(IConferenceApplicationData, __uuidof(IConferenceApplicationData));
_COM_SMARTPTR_TYPEDEF(_IUtilitiesCallback, __uuidof(_IUtilitiesCallback));
_COM_SMARTPTR_TYPEDEF(IScheduledConferenceManager, __uuidof(IScheduledConferenceManager));
_COM_SMARTPTR_TYPEDEF(IUtilities2, __uuidof(IUtilities2));
_COM_SMARTPTR_TYPEDEF(_IUtilities2Events, __uuidof(_IUtilities2Events));
_COM_SMARTPTR_TYPEDEF(IExchangeECPUrlChangedEventData, __uuidof(IExchangeECPUrlChangedEventData));
_COM_SMARTPTR_TYPEDEF(IScheduledConference, __uuidof(IScheduledConference));
_COM_SMARTPTR_TYPEDEF(IScheduledConferenceCollection, __uuidof(IScheduledConferenceCollection));
_COM_SMARTPTR_TYPEDEF(_IScheduledConferenceManagerCallback, __uuidof(_IScheduledConferenceManagerCallback));
_COM_SMARTPTR_TYPEDEF(IRoomMessageDictionary, __uuidof(IRoomMessageDictionary));
_COM_SMARTPTR_TYPEDEF(IRoomMessage, __uuidof(IRoomMessage));
_COM_SMARTPTR_TYPEDEF(IRoomMessageCollection, __uuidof(IRoomMessageCollection));
_COM_SMARTPTR_TYPEDEF(IRoomUser, __uuidof(IRoomUser));
_COM_SMARTPTR_TYPEDEF(IRoomPropertyDictionary, __uuidof(IRoomPropertyDictionary));
_COM_SMARTPTR_TYPEDEF(IRoom, __uuidof(IRoom));
_COM_SMARTPTR_TYPEDEF(IRoomCollection, __uuidof(IRoomCollection));
_COM_SMARTPTR_TYPEDEF(IRoomManager, __uuidof(IRoomManager));
_COM_SMARTPTR_TYPEDEF(IRoomPropertyChangedEventData, __uuidof(IRoomPropertyChangedEventData));
_COM_SMARTPTR_TYPEDEF(IUnreadMessageCountChangedEventData, __uuidof(IUnreadMessageCountChangedEventData));
_COM_SMARTPTR_TYPEDEF(IRoomMessageEventData, __uuidof(IRoomMessageEventData));
_COM_SMARTPTR_TYPEDEF(IRoomMessagesEventData, __uuidof(IRoomMessagesEventData));
_COM_SMARTPTR_TYPEDEF(IRoomUserChangedEventData, __uuidof(IRoomUserChangedEventData));
_COM_SMARTPTR_TYPEDEF(IRoomParticipantsEventData, __uuidof(IRoomParticipantsEventData));
_COM_SMARTPTR_TYPEDEF(IRoomJoinStateChangedEventData, __uuidof(IRoomJoinStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(_IRoomEvents, __uuidof(_IRoomEvents));
_COM_SMARTPTR_TYPEDEF(_IRoomManagerEvents, __uuidof(_IRoomManagerEvents));
_COM_SMARTPTR_TYPEDEF(IFollowedRoomsChangedEventData, __uuidof(IFollowedRoomsChangedEventData));
_COM_SMARTPTR_TYPEDEF(IRoomManagerStateChangedEventData, __uuidof(IRoomManagerStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(_IRoomUserEvents, __uuidof(_IRoomUserEvents));
_COM_SMARTPTR_TYPEDEF(_IRoomCallback, __uuidof(_IRoomCallback));
_COM_SMARTPTR_TYPEDEF(_IRoomManagerCallback, __uuidof(_IRoomManagerCallback));
_COM_SMARTPTR_TYPEDEF(IConversationWindowNeedsSizeChangeEventData, __uuidof(IConversationWindowNeedsSizeChangeEventData));
_COM_SMARTPTR_TYPEDEF(IConversationWindowNeedsAttentionEventData, __uuidof(IConversationWindowNeedsAttentionEventData));
_COM_SMARTPTR_TYPEDEF(_IAutomationCallback, __uuidof(_IAutomationCallback));
_COM_SMARTPTR_TYPEDEF(_IConversationWindowEvents, __uuidof(_IConversationWindowEvents));
_COM_SMARTPTR_TYPEDEF(_IConversationWindowCallback, __uuidof(_IConversationWindowCallback));
_COM_SMARTPTR_TYPEDEF(IConversationWindowInformationDictionary, __uuidof(IConversationWindowInformationDictionary));
_COM_SMARTPTR_TYPEDEF(IConversationWindowActions, __uuidof(IConversationWindowActions));
_COM_SMARTPTR_TYPEDEF(IVideoViewInformationDictionary, __uuidof(IVideoViewInformationDictionary));
_COM_SMARTPTR_TYPEDEF(IVideoView, __uuidof(IVideoView));
_COM_SMARTPTR_TYPEDEF(IConversationWindow2, __uuidof(IConversationWindow2));
_COM_SMARTPTR_TYPEDEF(_IConversationWindow2Events, __uuidof(_IConversationWindow2Events));
_COM_SMARTPTR_TYPEDEF(IConversationWindowStateChangedEventData, __uuidof(IConversationWindowStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IConversationWindowInformationChangedEventData, __uuidof(IConversationWindowInformationChangedEventData));
_COM_SMARTPTR_TYPEDEF(IConversationWindowActionAvailabilityChangedEventData, __uuidof(IConversationWindowActionAvailabilityChangedEventData));
_COM_SMARTPTR_TYPEDEF(IShareableContentPropertyDictionary, __uuidof(IShareableContentPropertyDictionary));
_COM_SMARTPTR_TYPEDEF(_IContentSharingModalityEvents, __uuidof(_IContentSharingModalityEvents));
_COM_SMARTPTR_TYPEDEF(_IShareableContentEvents, __uuidof(_IShareableContentEvents));
_COM_SMARTPTR_TYPEDEF(IShareableContentStateChangedEventData, __uuidof(IShareableContentStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IShareableContentPropertyChangedEventData, __uuidof(IShareableContentPropertyChangedEventData));
_COM_SMARTPTR_TYPEDEF(IShareableContentActionAvailabilityChangedEventData, __uuidof(IShareableContentActionAvailabilityChangedEventData));
_COM_SMARTPTR_TYPEDEF(_IConversation2Callback, __uuidof(_IConversation2Callback));
_COM_SMARTPTR_TYPEDEF(_IContentSharingModalityCallback, __uuidof(_IContentSharingModalityCallback));
_COM_SMARTPTR_TYPEDEF(_IShareableContentCallback, __uuidof(_IShareableContentCallback));
_COM_SMARTPTR_TYPEDEF(_IParticipant2Callback, __uuidof(_IParticipant2Callback));
_COM_SMARTPTR_TYPEDEF(_IChannel2Callback, __uuidof(_IChannel2Callback));
_COM_SMARTPTR_TYPEDEF(IConferenceAccessInformation2, __uuidof(IConferenceAccessInformation2));
_COM_SMARTPTR_TYPEDEF(IVideoViewInformationChangedEventData, __uuidof(IVideoViewInformationChangedEventData));
_COM_SMARTPTR_TYPEDEF(_IVideoViewEvents, __uuidof(_IVideoViewEvents));
_COM_SMARTPTR_TYPEDEF(IClientSettings, __uuidof(IClientSettings));
_COM_SMARTPTR_TYPEDEF(IClientSettingsChangedEventData, __uuidof(IClientSettingsChangedEventData));
_COM_SMARTPTR_TYPEDEF(IPublicCloudProviderInfo, __uuidof(IPublicCloudProviderInfo));
_COM_SMARTPTR_TYPEDEF(_IClientSettingsEvents, __uuidof(_IClientSettingsEvents));
_COM_SMARTPTR_TYPEDEF(IUCOfficeIntegration, __uuidof(IUCOfficeIntegration));
_COM_SMARTPTR_TYPEDEF(ISignInConfiguration, __uuidof(ISignInConfiguration));
_COM_SMARTPTR_TYPEDEF(IApplicationRegistration, __uuidof(IApplicationRegistration));
_COM_SMARTPTR_TYPEDEF(IWebAppExtension, __uuidof(IWebAppExtension));
_COM_SMARTPTR_TYPEDEF(IClientStateChangedEventData, __uuidof(IClientStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(ILyncClientNotificationReceivedEventData, __uuidof(ILyncClientNotificationReceivedEventData));
_COM_SMARTPTR_TYPEDEF(ICredentialRequestedEventData, __uuidof(ICredentialRequestedEventData));
_COM_SMARTPTR_TYPEDEF(ISignInDelayedEventData, __uuidof(ISignInDelayedEventData));
_COM_SMARTPTR_TYPEDEF(IPreferredCapabilitiesChangedEventData, __uuidof(IPreferredCapabilitiesChangedEventData));
_COM_SMARTPTR_TYPEDEF(_ILyncClientCallback, __uuidof(_ILyncClientCallback));
_COM_SMARTPTR_TYPEDEF(_IClientEvents, __uuidof(_IClientEvents));
_COM_SMARTPTR_TYPEDEF(_ILyncClientEvents, __uuidof(_ILyncClientEvents));
_COM_SMARTPTR_TYPEDEF(_IDelegatorClientEvents, __uuidof(_IDelegatorClientEvents));
_COM_SMARTPTR_TYPEDEF(_IUCOfficeIntegrationEvents, __uuidof(_IUCOfficeIntegrationEvents));
_COM_SMARTPTR_TYPEDEF(IGroupCollectionChangedEventData, __uuidof(IGroupCollectionChangedEventData));
_COM_SMARTPTR_TYPEDEF(IContactCollection, __uuidof(IContactCollection));
_COM_SMARTPTR_TYPEDEF(IGroup, __uuidof(IGroup));
_COM_SMARTPTR_TYPEDEF(IGroupCollection, __uuidof(IGroupCollection));
_COM_SMARTPTR_TYPEDEF(IDistributionGroup, __uuidof(IDistributionGroup));
_COM_SMARTPTR_TYPEDEF(ICustomGroup, __uuidof(ICustomGroup));
_COM_SMARTPTR_TYPEDEF(IFavoriteContacts, __uuidof(IFavoriteContacts));
_COM_SMARTPTR_TYPEDEF(IContact, __uuidof(IContact));
_COM_SMARTPTR_TYPEDEF(IContactSubscription, __uuidof(IContactSubscription));
_COM_SMARTPTR_TYPEDEF(IContactManager, __uuidof(IContactManager));
_COM_SMARTPTR_TYPEDEF(IGroupMemberChangedEventData, __uuidof(IGroupMemberChangedEventData));
_COM_SMARTPTR_TYPEDEF(IRecentContacts, __uuidof(IRecentContacts));
_COM_SMARTPTR_TYPEDEF(IFrequentContacts, __uuidof(IFrequentContacts));
_COM_SMARTPTR_TYPEDEF(IContact2, __uuidof(IContact2));
_COM_SMARTPTR_TYPEDEF(ISelf, __uuidof(ISelf));
_COM_SMARTPTR_TYPEDEF(IUtilities, __uuidof(IUtilities));
_COM_SMARTPTR_TYPEDEF(IClient2, __uuidof(IClient2));
_COM_SMARTPTR_TYPEDEF(IConversationCollection, __uuidof(IConversationCollection));
_COM_SMARTPTR_TYPEDEF(IConversation, __uuidof(IConversation));
_COM_SMARTPTR_TYPEDEF(IConversationHistory, __uuidof(IConversationHistory));
_COM_SMARTPTR_TYPEDEF(IConversationManager, __uuidof(IConversationManager));
_COM_SMARTPTR_TYPEDEF(ISummarizerStatesChangedEventData, __uuidof(ISummarizerStatesChangedEventData));
_COM_SMARTPTR_TYPEDEF(IConversationManagerEventData, __uuidof(IConversationManagerEventData));
_COM_SMARTPTR_TYPEDEF(IVoiceMail, __uuidof(IVoiceMail));
_COM_SMARTPTR_TYPEDEF(IVoiceMailCollection, __uuidof(IVoiceMailCollection));
_COM_SMARTPTR_TYPEDEF(IVoiceMailBatchEventData, __uuidof(IVoiceMailBatchEventData));
_COM_SMARTPTR_TYPEDEF(IConversationWindow, __uuidof(IConversationWindow));
_COM_SMARTPTR_TYPEDEF(IAutomation, __uuidof(IAutomation));
_COM_SMARTPTR_TYPEDEF(IClient, __uuidof(IClient));
_COM_SMARTPTR_TYPEDEF(ILyncClient, __uuidof(ILyncClient));
_COM_SMARTPTR_TYPEDEF(IDelegatorClient, __uuidof(IDelegatorClient));
_COM_SMARTPTR_TYPEDEF(IDelegatorClientCollectionEventData, __uuidof(IDelegatorClientCollectionEventData));
_COM_SMARTPTR_TYPEDEF(IUCOfficeIntegrationHelper, __uuidof(IUCOfficeIntegrationHelper));
_COM_SMARTPTR_TYPEDEF(IParticipant, __uuidof(IParticipant));
_COM_SMARTPTR_TYPEDEF(IModality, __uuidof(IModality));
_COM_SMARTPTR_TYPEDEF(IModalityDictionary, __uuidof(IModalityDictionary));
_COM_SMARTPTR_TYPEDEF(IParticipantCollection, __uuidof(IParticipantCollection));
_COM_SMARTPTR_TYPEDEF(IInstantMessageModality, __uuidof(IInstantMessageModality));
_COM_SMARTPTR_TYPEDEF(IChannel, __uuidof(IChannel));
_COM_SMARTPTR_TYPEDEF(IAudioChannel, __uuidof(IAudioChannel));
_COM_SMARTPTR_TYPEDEF(IVideoChannel, __uuidof(IVideoChannel));
_COM_SMARTPTR_TYPEDEF(IAVModality, __uuidof(IAVModality));
_COM_SMARTPTR_TYPEDEF(IApplicationSharingModality, __uuidof(IApplicationSharingModality));
_COM_SMARTPTR_TYPEDEF(IVideoChannelCollection, __uuidof(IVideoChannelCollection));
_COM_SMARTPTR_TYPEDEF(IParticipantCollectionChangedEventData, __uuidof(IParticipantCollectionChangedEventData));
_COM_SMARTPTR_TYPEDEF(IStreamStateChangedEventData, __uuidof(IStreamStateChangedEventData));
_COM_SMARTPTR_TYPEDEF(IControlRequestReceivedEventData, __uuidof(IControlRequestReceivedEventData));
_COM_SMARTPTR_TYPEDEF(IShareableContent, __uuidof(IShareableContent));
_COM_SMARTPTR_TYPEDEF(IShareableContentCollection, __uuidof(IShareableContentCollection));
_COM_SMARTPTR_TYPEDEF(IContentSharingModality, __uuidof(IContentSharingModality));
_COM_SMARTPTR_TYPEDEF(IContentCollectionChangedEventData, __uuidof(IContentCollectionChangedEventData));
_COM_SMARTPTR_TYPEDEF(IActivePresenterChangedEventData, __uuidof(IActivePresenterChangedEventData));
_COM_SMARTPTR_TYPEDEF(IActiveContentChangedEventData, __uuidof(IActiveContentChangedEventData));
_COM_SMARTPTR_TYPEDEF(IPowerPointContent, __uuidof(IPowerPointContent));
_COM_SMARTPTR_TYPEDEF(IConversationContextCollection, __uuidof(IConversationContextCollection));
_COM_SMARTPTR_TYPEDEF(IConversationContext, __uuidof(IConversationContext));
_COM_SMARTPTR_TYPEDEF(IConversationContextList, __uuidof(IConversationContextList));
_COM_SMARTPTR_TYPEDEF(IConversationContextCollectionEventData, __uuidof(IConversationContextCollectionEventData));
_COM_SMARTPTR_TYPEDEF(IAccessPermission, __uuidof(IAccessPermission));
_COM_SMARTPTR_TYPEDEF(IAccessEntry, __uuidof(IAccessEntry));
_COM_SMARTPTR_TYPEDEF(IAccessEntryCollectionChangedEventData, __uuidof(IAccessEntryCollectionChangedEventData));

//
// Type library items
//
    const long ucMinAudioVolume = 0;
    const long ucMaxAudioVolume = 65535;
    const long uiaConversationModeAll = 31;
    const unsigned long ucDefaultMaxSearchResults = 30;

enum PropertyMask
{
    ucPropertyMaskType = -268435456,
    ucPropertyMaskClass = 268369920,
    ucPropertyMaskId = 65535
};

enum PropertyMaskIndex
{
    ucPropertyMaskIndexType = 28,
    ucPropertyMaskIndexClass = 16,
    ucPropertyMaskIndexId = 0
};

enum PropertyType
{
    ucPropertyString = 0,
    ucPropertyBoolean = 268435456,
    ucPropertyNumeric = 536870912,
    ucPropertyEnumeration = 805306368,
    ucPropertyObject = 1073741824,
    ucPropertyDate = 1342177280,
    ucPropertySafeArray = 1610612736
};

enum PropertyClass
{
    ucPropertyClassContact = 0,
    ucPropertyClassGroup = 65536,
    ucPropertyClassOperation = 131072,
    ucPropertyClassUCClientAddtionalInfo = 196608,
    ucPropertyClassSearchProvider = 262144,
    ucPropertyClassOption = 327680,
    ucPropertyClassConferenceInfo = 393216,
    ucPropertyClassConferenceSetting = 458752,
    ucPropertyClassConversation = 524288,
    ucPropertyClassConversationStateChange = 589824,
    ucPropertyClassParticipant = 655360,
    ucPropertyClassModality = 720896,
    ucPropertyClassModalityStateChange = 786432,
    ucPropertyClassModalityAction = 851968,
    ucPropertyClassConversationWindow = 917504,
    ucPropertyClassRoom = 983040,
    ucPropertyClassContent = 1048576,
    ucPropertyClassVideoView = 1114112,
    ucPropertyClassClientSettings = 1179648,
    ucPropertyClassReserved1 = 327680000,
    ucPropertyClassReserved2 = 327745536,
    ucPropertyClassReserved3 = 327811072,
    ucPropertyClassReserved4 = 327876608
};

enum ConversationWindowExtensionSize
{
    ucExtensibilityWindowSmall = 0,
    ucExtensibilityWindowMedium = 1,
    ucExtensibilityWindowLarge = 2,
    ucConversationWindowExtensionSizeMax = 3
};

enum ModalityTypes
{
    ucModalityNone = 0,
    ucModalityInstantMessage = 1,
    ucModalityAudioVideo = 2,
    ucModalityApplicationSharing = 4,
    ucModalityContentSharing = 8,
    ucModalityAppSharing = 16,
    ucModalityContent = 32
};

struct __declspec(uuid("f7ee956b-a713-48d3-afef-0a28f37740b3"))
IAsynchronousOperation : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_AsyncState (
        /*[out,retval]*/ VARIANT * _userState ) = 0;
      virtual HRESULT __stdcall get_StatusCode (
        /*[out,retval]*/ long * _HResultStatusCode ) = 0;
      virtual HRESULT __stdcall get_IsCompleted (
        /*[out,retval]*/ VARIANT_BOOL * _isCompleted ) = 0;
      virtual HRESULT __stdcall get_IsSucceeded (
        /*[out,retval]*/ VARIANT_BOOL * _isSucceeded ) = 0;
      virtual HRESULT __stdcall get_IsCancelable (
        /*[out,retval]*/ VARIANT_BOOL * _isCancelable ) = 0;
      virtual HRESULT __stdcall get_DiagnosticCode (
        /*[out,retval]*/ long * _HResultDiagnosticCode ) = 0;
      virtual HRESULT __stdcall CancelOperation ( ) = 0;
};

struct __declspec(uuid("c12f0de5-9a7d-425c-b391-8be004eaa2f6"))
ILocaleString : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_LocaleId (
        /*[out,retval]*/ long * _localeId ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ BSTR * _stringValue ) = 0;
};

struct __declspec(uuid("dc63d621-11c8-4490-b017-b8c57ee1be25"))
IAddress : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_StreetAddress (
        /*[out,retval]*/ BSTR * _streetAddress ) = 0;
      virtual HRESULT __stdcall get_City (
        /*[out,retval]*/ BSTR * _city ) = 0;
      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ BSTR * _state ) = 0;
      virtual HRESULT __stdcall get_PostalCode (
        /*[out,retval]*/ BSTR * _postalCode ) = 0;
      virtual HRESULT __stdcall get_CountryCode (
        /*[out,retval]*/ BSTR * _countryCode ) = 0;
};

struct __declspec(uuid("1a8a9402-e89f-40ae-88a8-b328391684b6"))
INamedPropertyList : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall Add (
        /*[in]*/ BSTR _propertyID,
        /*[in]*/ VARIANT _propertyValue ) = 0;
      virtual HRESULT __stdcall Clear ( ) = 0;
      virtual HRESULT __stdcall RemoveAt (
        /*[in]*/ long _index ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall GetIDAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ BSTR * _propertyID ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _propertyValue ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ BSTR _propertyID,
        /*[out]*/ VARIANT * _propertyValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("e6a9d57e-838c-43d1-9f12-e47c844531e0"))
AsynchronousOperation;
    // [ default ] interface IAsynchronousOperation

struct __declspec(uuid("8cd4531b-d249-489c-b55a-37e4a0badb25"))
LocaleString;
    // [ default ] interface ILocaleString

struct __declspec(uuid("99880cd9-73b0-4b84-93ff-95bcd18c9ef8"))
Address;
    // [ default ] interface IAddress

struct __declspec(uuid("66b28b84-9d92-47ff-9152-e89eafbaf3c2"))
NamedPropertyList;
    // [ default ] interface INamedPropertyList

enum ContactSetting
{
    ucContactTagged = 268435456,
    ucContactAccessLevel = 805306369,
    ucContactExchangeServiceEntryId = 2,
    ucContactSource = 805306371,
    ucContactDefaultContactEndpoint = 1073741828
};

enum GroupAction
{
    ucGroupAddContact = 1,
    ucGroupRemoveContact = 2,
    ucGroupDeleteGroup = 3,
    ucGroupRenameGroup = 4
};

enum GroupType
{
    ucCustomGroup = 1,
    ucRecentContacts = 2,
    ucDistributionGroup = 3,
    ucFrequentContacts = 4,
    ucFavoriteContacts = 5,
    ucDelegateGroup = 6,
    ucResponseServiceGroup = 7
};

enum ContactInformationType
{
    ucPresenceAvailability = 0,
    ucPresenceActivityId = 1,
    ucPresenceLocationName = 2,
    ucPresenceTimeZone = 3,
    ucPresenceTimeZoneBias = 4,
    ucPresenceMeetingSubject = 5,
    ucPresenceMeetingLocation = 6,
    ucPresenceActivity = 7,
    ucPresenceCustomActivity = 8,
    ucPresenceIdleStartTime = 9,
    ucPresenceDisplayName = 10,
    ucPresenceLegacyInstantMessageAddresses = 11,
    ucPresencePrimaryEmailAddress = 12,
    ucPresenceEmailAddresses = 13,
    ucPresenceTitle = 14,
    ucPresenceCompany = 15,
    ucPresenceDepartment = 16,
    ucPresenceOffice = 17,
    ucPresenceHomePageUrl = 18,
    ucPresencePhoto = 19,
    ucPresenceDefaultNote = 20,
    ucPresenceDefaultNoteType = 21,
    ucPresencePersonalNote = 22,
    ucPresenceOutOfficeNote = 23,
    ucPresenceSourceNetwork = 24,
    ucPresenceIconUrl = 25,
    ucPresenceIconStream = 26,
    ucPresenceContactEndpoints = 27,
    ucPresenceCurrentLegacyCalendarState = 28,
    ucPresenceNextLegacyCalendarState = 29,
    ucPresenceNextCalendarStateStartTime = 30,
    ucPresenceWorkingHours = 31,
    ucPresenceCapabilityString = 32,
    ucPresenceCapabilities = 33,
    ucPresenceContactType = 34,
    ucPresenceDescription = 35,
    ucPresenceWorkAddress = 36,
    ucPresenceFirstName = 37,
    ucPresenceLastName = 38,
    ucPresenceVoiceMemo = 39,
    ucPresenceWorkPhoneNumber = 40,
    ucPresenceHomePhoneNumber = 41,
    ucPresenceCellPhoneNumber = 42,
    ucPresenceOtherPhoneNumber = 43,
    ucContactInformationTypeMax_W14 = 44,
    ucPresenceCapabilityDetails = 45,
    ucPresenceDefaultNotePublishedTime = 46,
    ucPresenceCurrentCalendarState = 47,
    ucPresenceNextCalendarState = 48,
    ucPresenceAttributionString = 49,
    ucPresenceHiResolutionPhoto = 50,
    ucPresencePersonId = 51,
    ucPresenceDisableAggregation = 52,
    ucPresencePhotoFromPresence = 53,
    ucPresencePhotoFromActiveDirectory = 54,
    ucPresenceInstantMessageAddresses = 55,
    ucPresenceIsOutOfOffice = 56,
    ucContactInformationTypeMax = 57
};

enum UnifiedCommunicationType
{
    ucUnifiedCommunicationUnknown = 0,
    ucUnifiedCommunicationEnabled = 1,
    ucUnifiedCommunicationNotEnabled = 2
};

enum ContactSubscriptionRefreshRate
{
    ucSubscriptionFreshnessNone = 0,
    ucSubscriptionFreshnessLow = 1,
    ucSubscriptionFreshnessHigh = 2
};

enum ContactSourceTypes
{
    ucContactSourceUnknown = 0,
    ucContactSourceGlobalAddressList = 1,
    ucContactSourceWindowsAddressBook = 2,
    ucContactSourceExchangeService = 4
};

enum SearchProviders
{
    ucSearchProviderDefault = 0,
    ucSearchProviderExchangeService = 1,
    ucSearchProviderGlobalAddressList = 2,
    ucSearchProviderWindowsAddressBook = 4,
    ucSearchProviderOtherContacts = 8,
    ucSearchProviderPersonalContacts = 5,
    ucSearchProviderExpert = 65535,
    ucSearchProviderLync = 65534
};

enum SearchFields
{
    ucSearchFirstName = 1,
    ucSearchLastName = 2,
    ucSearchDisplayName = 4,
    ucSearchCompany = 8,
    ucSearchPrimaryEmailAddress = 16,
    ucSearchEmailAddresses = 32,
    ucSearchAlias = 64,
    ucSearchPhoneNumbers = 128,
    ucSearchPhoneExtention = 256,
    ucSearchAllFields = -1
};

enum SearchOptions
{
    ucSearchDefault = 0,
    ucSearchMatchWholeWord = 1,
    ucSearchContactsOnly = 2,
    ucSearchIncludeContactsWithoutSipOrTelUri = 4,
    ucSearchAllowProgressiveCallback = 268435456
};

enum SearchProviderStatusType
{
    ucProviderStatusSyncNotStarted = 0,
    ucProviderStatusSyncInProgress = 1,
    ucProviderStatusSyncSucceeded = 2,
    ucProviderStatusSyncSucceededForInternalOnly = 3,
    ucProviderStatusSyncSucceededForExternalOnly = 4,
    ucProviderStatusOtherFailure = 1000,
    ucProviderStatusLocalDatabaseFailure = 1001,
    ucProviderStatusFileNotFound = 1002,
    ucProviderStatusFileCorrupted = 1003,
    ucProviderStatusCredentialsNotEntered = 1004,
    ucProviderStatusNotConfigured = 1005,
    ucProviderStatusInitializationFailed = 1006
};

enum ContactType
{
    ucContactPresentityUnknown = 0,
    ucContactPresentityPerson = 1,
    ucContactPresentityHuntgroup = 2,
    ucContactPresentityAutoAttendant = 3,
    ucContactPresentityAutomatedBot = 4
};

enum SourceNetworkType
{
    ucSourceNetworkUnknown = 0,
    ucSourceNetworkEnterprise = 1,
    ucSourceNetworkFederatedEnterprise = 2,
    ucSourceNetworkFederatedPublic = 3
};

enum ContactCalendarState
{
    ucCalendarStateAvailable = 0,
    ucCalendarStateTentative = 1,
    ucCalendarStateOutsideWorkPeriod = 2,
    ucCalendarStateBusy = 3,
    ucCalendarStateOutOfOffice = 4,
    ucCalendarStateNotAvailable = 5,
    ucCalendarStateUnknown = 6,
    ucCalendarStateWorkingElsewhere = 7
};

enum DaysOfWeek
{
    ucSunday = 1,
    ucMonday = 2,
    ucTuesday = 4,
    ucWednesday = 8,
    ucThursday = 16,
    ucFriday = 32,
    ucSaturday = 64
};

enum DefaultNoteType
{
    ucNoteTypeUnknown = 0,
    ucNoteTypePersonal = 1,
    ucNoteTypeOutOfOffice = 2
};

enum ContactCapabilities
{
    ucContactCanShowPresence = 1,
    ucContactCanCaptureInstantMessage = 2,
    ucContactCanRenderInstantMessage = 4,
    ucContactCanCaptureAudio = 8,
    ucContactCanRenderAudio = 16,
    ucContactCanCaptureVideo = 32,
    ucContactCanRenderVideo = 64,
    ucContactCanCaptureAppShare = 128,
    ucContactCanRenderAppShare = 256
};

enum ContactEndpointType
{
    ucContactEndpointTypeWorkPhone = 0,
    ucContactEndpointTypeMobilePhone = 1,
    ucContactEndpointTypeHomePhone = 2,
    ucContactEndpointTypeOtherPhone = 3,
    ucContactEndpointTypeLync = 4,
    ucContactEndpointTypeVoiceMail = 5,
    ucContactEndpointTypeMax = 6
};

enum OrganizationStructureTypes
{
    ucOrganizationInfoManagers = 1,
    ucOrganizationInfoPeers = 2,
    ucOrganizationInfoDirectReports = 4
};

struct __declspec(uuid("f0ce810f-e36f-4cd8-a836-515b9a6e09fc"))
IContactInformationDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ContactInformationType _key,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ContactInformationType _key,
        /*[out]*/ VARIANT * _value,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ContactInformationType * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ContactInformationType _key,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("fd9000b3-479f-4b16-9d63-70a49e078946"))
IContactSettingDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ContactSetting _key,
        /*[out]*/ VARIANT * _value,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ContactSetting _key,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ContactSetting * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ContactSetting _key,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
};

struct __declspec(uuid("1ff0e6f1-2fe4-4e29-a123-557af0db6927"))
IContactEndpoint : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum ContactEndpointType * _type ) = 0;
      virtual HRESULT __stdcall get_DisplayName (
        /*[out,retval]*/ BSTR * _displayString ) = 0;
      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * _uri ) = 0;
      virtual HRESULT __stdcall CanStart (
        /*[in]*/ enum ModalityTypes _modalityTypes,
        /*[out,retval]*/ VARIANT_BOOL * _canStart ) = 0;
};

struct __declspec(uuid("efec2816-f16d-48d8-9306-26c810f0ea55"))
IContactInformationChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ChangedContactInformation (
        /*[out,retval]*/ SAFEARRAY * * _contactInformationTypes ) = 0;
};

struct __declspec(uuid("d2bfc7d7-a64f-48ef-bf90-c3c7e684d89c"))
IDistributionGroupExpansionStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_IsExpanded (
        /*[out,retval]*/ VARIANT_BOOL * _isExpanded ) = 0;
};

struct __declspec(uuid("0bdb9057-28ae-4bf0-aff0-12a148e51637"))
IUriChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_NewUri (
        /*[out,retval]*/ BSTR * _newUri ) = 0;
      virtual HRESULT __stdcall get_OldUri (
        /*[out,retval]*/ BSTR * _oldUri ) = 0;
};

struct __declspec(uuid("283c2089-b760-4d65-9199-716627174f7a"))
IGroupNameChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_NewName (
        /*[out,retval]*/ BSTR * _newGroupName ) = 0;
};

struct __declspec(uuid("378cd767-1b64-4a0e-a941-8f52869fc9df"))
ISearchResult : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Result (
        /*[out,retval]*/ IDispatch * * _searchResult ) = 0;
      virtual HRESULT __stdcall get_HitHighlightSummary (
        /*[out,retval]*/ BSTR * _highlightSummary ) = 0;
};

struct __declspec(uuid("7356d7bb-fd71-4554-84a1-3bbe28726551"))
ISearchResults : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_IsMoreAvailable (
        /*[out,retval]*/ VARIANT_BOOL * _isMoreAvailable ) = 0;
      virtual HRESULT __stdcall get_AllContactsAndGroups (
        /*[out,retval]*/ SAFEARRAY * * _allResults ) = 0;
      virtual HRESULT __stdcall get_Contacts (
        /*[out,retval]*/ SAFEARRAY * * _contacts ) = 0;
      virtual HRESULT __stdcall get_Groups (
        /*[out,retval]*/ SAFEARRAY * * _groups ) = 0;
      virtual HRESULT __stdcall get_AllResults (
        /*[out,retval]*/ SAFEARRAY * * _resultsWithAdditionalInfo ) = 0;
};

struct __declspec(uuid("995f992c-9def-47b9-bf11-81813c0c0e28"))
ISearchProviderStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Provider (
        /*[out,retval]*/ enum SearchProviders * _provider ) = 0;
      virtual HRESULT __stdcall get_NewStatus (
        /*[out,retval]*/ enum SearchProviderStatusType * _status ) = 0;
};

struct __declspec(uuid("339d28c6-3d1a-45d4-ba14-a56742749a43"))
IWorkingPeriod : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_StartTime (
        /*[out,retval]*/ DATE * _startTime ) = 0;
      virtual HRESULT __stdcall get_EndTime (
        /*[out,retval]*/ DATE * _endTime ) = 0;
      virtual HRESULT __stdcall get_Days (
        /*[out,retval]*/ enum DaysOfWeek * _daysOfweek ) = 0;
};

struct __declspec(uuid("53d014c1-54db-42b3-9dfd-8e231ef2c356"))
IWorkingHours : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_EmailAddress (
        /*[out,retval]*/ BSTR * _emailAddress ) = 0;
      virtual HRESULT __stdcall get_WorkingPeriods (
        /*[out,retval]*/ SAFEARRAY * * _workingPeriods ) = 0;
};

enum PresenceCapabilityType
{
    ucPresenceCapabilityInstantMessaging = 0,
    ucPresenceCapabilityAudio = 1,
    ucPresenceCapabilityVideo = 2,
    ucPresenceCapabilityAppShare = 3
};

struct __declspec(uuid("10c3acc7-6c0c-4f35-adfb-ae0f200a8583"))
IPresenceCapability : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum PresenceCapabilityType * _type ) = 0;
      virtual HRESULT __stdcall get_CanRender (
        /*[out,retval]*/ VARIANT_BOOL * _canRender ) = 0;
      virtual HRESULT __stdcall get_CanCapture (
        /*[out,retval]*/ VARIANT_BOOL * _canCapture ) = 0;
      virtual HRESULT __stdcall get_Availability (
        /*[out,retval]*/ long * _availability ) = 0;
};

struct __declspec(uuid("8affb8ce-5404-4280-bca3-e2e4388e6d73"))
IDescriptionChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Description (
        /*[out,retval]*/ BSTR * _description ) = 0;
};

struct __declspec(uuid("7cd92461-ca4f-4beb-a636-8bbcc192e60f"))
ISearchResults2 : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_FastResults (
        /*[out,retval]*/ SAFEARRAY * * _fastResults ) = 0;
      virtual HRESULT __stdcall GetPartialResults (
        /*[out]*/ SAFEARRAY * * _results,
        /*[out]*/ VARIANT_BOOL * insert,
        /*[out]*/ unsigned long * index ) = 0;
};

struct __declspec(uuid("abf012a2-2068-4b4d-80d7-6d386ba7c8a2"))
_IContactsAndGroupsCallback : IDispatch
{};

struct __declspec(uuid("c6605772-3d9d-44a7-b43a-4590e6b43ee3"))
_IGroupCallback : IDispatch
{};

struct __declspec(uuid("311dde48-ed7d-46fa-9e0a-2e314b7fef7c"))
_IContactCallback : IDispatch
{};

struct __declspec(uuid("50d86ee2-288f-44f5-8144-69f6e3b24b90"))
_IContactManagerEvents : IDispatch
{};

struct __declspec(uuid("b1c6ecc3-1ea7-495d-9337-02fe4fb723f0"))
_IContactEvents : IDispatch
{};

struct __declspec(uuid("12480bd4-072f-4763-b9fb-41b2ca54f9cd"))
IContactSettingChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Setting (
        /*[out,retval]*/ enum ContactSetting * _propertyType ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ VARIANT * _value ) = 0;
};

struct __declspec(uuid("9918673d-4d74-41cf-8840-d4996395a34c"))
_IGroupEvents : IDispatch
{};

struct __declspec(uuid("f719ed75-b7bd-45d3-9097-254509226f20"))
_ICustomGroupEvents : IDispatch
{};

struct __declspec(uuid("011b320c-9ac1-4dde-b595-3d0014472e2f"))
_IRecentContactsEvents : IDispatch
{};

struct __declspec(uuid("082ee280-726e-417f-99cb-81a0ccbff883"))
_IDistributionGroupEvents : IDispatch
{};

struct __declspec(uuid("aa186c92-181e-417f-b150-fca0f367e0fc"))
_IFrequentContactsEvents : IDispatch
{};

struct __declspec(uuid("62b932e9-7847-4d78-beee-5d58bf815b03"))
IContactPositionChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_StartPosition (
        /*[out,retval]*/ long * _position ) = 0;
      virtual HRESULT __stdcall get_EndPosition (
        /*[out,retval]*/ long * _position ) = 0;
};

struct __declspec(uuid("dab42e13-500a-409a-ac34-c1bf4719ee2e"))
_IFavoriteContactsEvents : IDispatch
{};

struct __declspec(uuid("43eaf09a-5dd2-4ce7-9b03-fceed187bf48"))
ContactManager;
    // [ default ] interface IContactManager
    // [ default, source ] dispinterface _IContactManagerEvents

struct __declspec(uuid("d3ef7c83-13f3-4bf2-8216-540fb95bbd00"))
Contact;
    // [ default ] interface IContact
    // interface IContact2
    // [ default, source ] dispinterface _IContactEvents

struct __declspec(uuid("95ee29d5-2b51-46e4-8f7c-406fc3827a80"))
Group;
    // [ default ] interface IGroup

struct __declspec(uuid("9bb71fc6-1bba-407a-a086-87d8e525bfef"))
ContactCollection;
    // [ default ] interface IContactCollection

struct __declspec(uuid("a2e0476a-945f-4c93-98fa-87aed29e9467"))
CustomGroup;
    // [ default ] interface ICustomGroup
    // [ default, source ] dispinterface _ICustomGroupEvents

struct __declspec(uuid("caa28f4e-98dd-4822-8462-955aaddf2364"))
DistributionGroup;
    // [ default ] interface IDistributionGroup
    // [ default, source ] dispinterface _IDistributionGroupEvents

struct __declspec(uuid("e6162417-23be-437b-b9d2-7bf6cac5e829"))
RecentContacts;
    // [ default ] interface IRecentContacts
    // [ default, source ] dispinterface _IRecentContactsEvents

struct __declspec(uuid("b395d986-2542-4c59-883e-6612a34e6d75"))
FrequentContacts;
    // [ default ] interface IFrequentContacts
    // [ default, source ] dispinterface _IFrequentContactsEvents

struct __declspec(uuid("c66de384-1f97-4f1a-af30-0005990000f5"))
FavoriteContacts;
    // [ default ] interface IFavoriteContacts
    // [ default, source ] dispinterface _IFavoriteContactsEvents

struct __declspec(uuid("bbc00952-0dc4-4abc-89f7-5336c525b8ae"))
GroupMemberChangedEventData;
    // [ default ] interface IGroupMemberChangedEventData

struct __declspec(uuid("dd734d35-7a6b-42d8-be0b-af68c76a487a"))
DistributionGroupExpansionStateChangedEventData;
    // [ default ] interface IDistributionGroupExpansionStateChangedEventData

struct __declspec(uuid("b89f18d4-5041-488a-8d81-de8ebb3a362c"))
UriChangedEventData;
    // [ default ] interface IUriChangedEventData

struct __declspec(uuid("4eb781a9-a304-4bc5-a399-2c79b20d4079"))
GroupNameChangedEventData;
    // [ default ] interface IGroupNameChangedEventData

struct __declspec(uuid("ab12b906-de1a-4fb7-8d1a-774a04c7f3dc"))
GroupCollectionChangedEventData;
    // [ default ] interface IGroupCollectionChangedEventData

struct __declspec(uuid("facab35b-efe6-4228-82db-5aaabc2db239"))
SearchResult;
    // [ default ] interface ISearchResult

struct __declspec(uuid("e73ddb55-acb3-4f33-a0ed-7b22d68bf311"))
SearchResults;
    // [ default ] interface ISearchResults
    // interface ISearchResults2

struct __declspec(uuid("b53c508e-e967-43d2-a5c6-85523859581f"))
SearchProviderStateChangedEventData;
    // [ default ] interface ISearchProviderStateChangedEventData

struct __declspec(uuid("da937c78-8d00-41ec-874f-156d50ee97ec"))
GroupCollection;
    // [ default ] interface IGroupCollection

struct __declspec(uuid("9c49c870-9104-41f6-ae37-109434f8ba8d"))
ContactInformationChangedEventData;
    // [ default ] interface IContactInformationChangedEventData

struct __declspec(uuid("1496ae1a-5ba4-484d-8b2b-57c44fe810cd"))
ContactSettingChangedEventData;
    // [ default ] interface IContactSettingChangedEventData

struct __declspec(uuid("9367ca1b-8317-448a-bf18-8d88d809d4df"))
ContactPositionChangedEventData;
    // [ default ] interface IContactPositionChangedEventData

struct __declspec(uuid("ff235ffe-edef-4d1b-9205-519888a7ad8d"))
WorkingHours;
    // [ default ] interface IWorkingHours

struct __declspec(uuid("0858ab97-0e11-49bd-8f88-4e2ae0232e33"))
ContactInformationDictionary;
    // [ default ] interface IContactInformationDictionary

struct __declspec(uuid("60a8f660-c321-4761-8aff-41072dcfdd03"))
ContactSettingDictionary;
    // [ default ] interface IContactSettingDictionary

struct __declspec(uuid("b482e1e6-afb7-4e58-91ee-3e20f6139220"))
WorkingPeriod;
    // [ default ] interface IWorkingPeriod

struct __declspec(uuid("feaeffd6-c800-46a0-b634-fdc415fec211"))
ContactSubscription;
    // [ default ] interface IContactSubscription

struct __declspec(uuid("8d418098-61cd-4970-a0e0-2e9bd6420976"))
PresenceCapability;
    // [ default ] interface IPresenceCapability

struct __declspec(uuid("d2f90870-11a2-4396-806b-d79ddfbe9f50"))
ContactEndpoint;
    // [ default ] interface IContactEndpoint

struct __declspec(uuid("9bf7a83a-92f2-4c1d-91b6-7be048efd8a5"))
DescriptionChangedEventData;
    // [ default ] interface IDescriptionChangedEventData

enum ConversationSummarizerStates
{
    ucConversationSummarizerNoSummarizedState = 0,
    ucConversationSummarizerLocalOutgoing = 1,
    ucConversationSummarizerLocalRinging = 2,
    ucConversationSummarizerRemoteRinging = 4,
    ucConversationSummarizerLocalJoining = 8,
    ucConversationSummarizerLocalTransferring = 16,
    ucConversationSummarizerRemoteTransferring = 32,
    ucConversationSummarizerInConversationSip = 64,
    ucConversationSummarizerInConversationTel = 128,
    ucConversationSummarizerInConferenceCall = 256,
    ucConversationSummarizerLocalP2PSipOnHold = 512,
    ucConversationSummarizerLocalP2PTelOnHold = 1024,
    ucConversationSummarizerLocalConfOnHold = 2048,
    ucConversationSummarizerAllRemoteOnHold = 4096,
    ucConversationSummarizerRemoteJoining = 8192,
    ucConversationSummarizerRemoteForwarding = 16384,
    ucConversationSummarizerEarlyMediaSip = 32768,
    ucConversationSummarizerEarlyMediaTel = 65536,
    ucConversationSummarizerPairedCall = 131072,
    ucConversationSummarizerConnectingToCaller = 262144,
    ucConversationSummarizerInConversationUSB = 524288,
    ucConversationSummarizerRemoteP2POnHoldWithoutMusic = 1048576,
    ucConversationSummarizerLocalScreenSharing = 2097152,
    ucConversationSummarizerLocalP2POnHold = 1536,
    ucConversationSummarizerLocalOnHold = 3584,
    ucConversationSummarizerInConversation = 98752,
    ucConversationSummarizerAllStates = -1
};

enum ConversationState
{
    ucConversationInactive = 0,
    ucConversationActive = 1,
    ucConversationParked = 2,
    ucConversationTerminated = 3
};

enum ConversationAcceptanceState
{
    ucConversationAcceptanceNone = 0,
    ucConversationAcceptanceAccepted = 1,
    ucConversationAcceptanceRejected = 2,
    ucConversationAcceptanceMissed = 3,
    ucConversationAcceptanceDeflected = 4
};

enum ModalityState
{
    ucModalityDisconnected = 0,
    ucModalityConnecting = 1,
    ucModalityNotified = 2,
    ucModalityJoining = 3,
    ucModalityConnectingToCaller = 4,
    ucModalityConnected = 5,
    ucModalitySuspended = 6,
    ucModalityOnHold = 7,
    ucModalityForwarding = 8,
    ucModalityTransferring = 9,
    ucModalityDisconnecting = 10
};

enum ModalityAction
{
    ucModalityActionConnect = 0,
    ucModalityActionDisconnect = 1,
    ucModalityActionSetProperty = 2,
    ucModalityActionHold = 3,
    ucModalityActionRetrieve = 4,
    ucModalityActionForward = 5,
    ucModalityActionRemoteTransfer = 6,
    ucModalityActionConsultAndTransfer = 7,
    ucModalityActionSendInstantMessage = 8,
    ucModalityActionSetIsTyping = 9,
    ucModalityActionSetAudioEndpoint = 10,
    ucModalityActionAccept = 11,
    ucModalityActionReject = 12,
    ucModalityActionLocalTransfer = 13,
    ucModalityActionRequestSharingControl = 14,
    ucModalityActionReleaseSharingControl = 15,
    ucModalityActionGrantSharingControl = 16,
    ucModalityActionRevokeSharingControl = 17,
    ucModalityActionAcceptSharingControlRequest = 18,
    ucModalityActionDeclineSharingControlRequest = 19,
    ucModalityActionCreateShareablePowerPointContent = 20,
    ucModalityActionCreateShareableWhiteboardContent = 21,
    ucModalityActionCreateShareableNativeFileOnlyContent = 22,
    ucModalityActionCreateShareablePollContent = 23,
    ucModalityActionCreateShareableNotesContent = 24
};

enum ParticipantAction
{
    ucParticipantActionSetProperty = 0,
    ucParticipantActionAdmit = 1,
    ucParticipantActionDeny = 2,
    ucParticipantActionPin = 3,
    ucParticipantActionLock = 4,
    ucParticipantActionSetMute = 5,
    ucParticipantEjectAction = 5000,
    ucParticipantRemoveAction = 5001
};

enum ConversationAction
{
    ucConversationActionMerge = 0,
    ucConversationActionPark = 1,
    ucConversationActionAddParticipant = 2,
    ucConversationActionRemoveParticipant = 3
};

enum ConversationHistoryAction
{
    ucConversationHistorySaveAllHistory = 0,
    ucConversationHistoryRetrievePast = 1,
    ucConversationHistoryRetrievePastByEntryId = 2
};

enum ConversationImportance
{
    ucConversationImportanceLow = 0,
    ucConversationImportanceNormal = 1,
    ucConversationImportanceHigh = 2,
    ucConversationImportanceEmergency = 3
};

enum ConversationDirection
{
    ucConversationDirectionNone = 0,
    ucConversationDirectionIncoming = 1,
    ucConversationDirectionOutgoing = 2
};

enum ConferenceEscalationProgress
{
    ConferenceEscalationNotStarted = 0,
    ConferenceEscalationSchedulingConference = 1,
    ConferenceEscalationConnectingToConference = 2,
    ConferenceEscalationConnectedToLobby = 3,
    ConferenceEscalationJoiningLocalMedia = 4,
    ConferenceEscalationWaitingForPair = 5,
    ConferenceEscalationWaitingForPeer = 6,
    ConferenceEscalationCompleted = 7,
    ConferenceEscalationFailed = 8,
    ConferenceEscalationAwaitingDisclaimerResponse = 9,
    ConferenceEscalationAwaitingJoinDialogResponse = 10
};

enum ConferenceAccessType
{
    ucConferenceAccessTypeUnknown = 0,
    ucConferenceAccessTypeOpen = 1,
    ucConferenceAccessTypeClosed = 16,
    ucConferenceAccessTypeAnonymous = 256,
    ucConferenceAccessTypeLocked = 4096
};

enum ConferenceDisclaimerState
{
    ucConferenceDisclaimerNone = 0,
    ucConferenceDisclaimerRejected = 1,
    ucConferenceDisclaimerAccepted = 2
};

enum ModalityDisconnectReason
{
    ucDisconnectReasonNone = 0,
    ucDisconnectReasonTimeout = 1,
    ucDisconnectReasonBusy = 2,
    ucDisconnectReasonNotAcceptableHere = 3,
    ucDisconnectReasonDecline = 4,
    ucDisconnectReasonDeclineEverywhere = 5,
    ucDisconnectReasonReplyOther = 6
};

enum TransferOptions
{
    ucTransferOptionsNone = 0,
    ucTransferOptionsDisallowRedirection = 1,
    ucTransferOptionsIncludeActiveContext = 2,
    ucTransferOptionsIncludeAllContexts = 4
};

enum ModalityConnectOptions
{
    ucConnectOptionsNone = 0,
    ucConnectOptionsIncludeActiveContext = 1,
    ucConnectOptionsIncludeAllContexts = 2
};

enum ChannelType
{
    ucChannelAudio = 0,
    ucChannelVideo = 1,
    ucChannelVideoPano = 2
};

enum ChannelState
{
    ucChannelStateNone = 0,
    ucChannelStateConnecting = 1,
    ucChannelStateNotified = 2,
    ucChannelStateSend = 3,
    ucChannelStateReceive = 4,
    ucChannelStateSendReceive = 5,
    ucChannelStateInactive = 6
};

enum ChannelAction
{
    ucChannelAction_Start = 0,
    ucChannelAction_Stop = 1,
    ucChannelAction_SendDtmf = 2,
    ucChannelAction_PauseSend = 5001,
    ucChannelAction_PauseReceive = 5002,
    ucChannelAction_PlaySend = 5003,
    ucChannelAction_PlayReceive = 5004,
    ucChannelAction_Pause = 5005,
    ucChannelAction_Resume = 5006
};

enum MediaStreamDirection
{
    MediaStreamDirection_None = 0,
    MediaStreamDirection_Send = 1,
    MediaStreamDirection_Receive = 2
};

enum ConversationContextType
{
    ucConversationContextSimple = 0,
    ucConversationContextMultiple = 1,
    ucConversationContextAlternate = 2
};

enum RepresentationType
{
    ucRepresentationTypeDelegate = 0,
    ucRepresentationTypeTeam = 1,
    ucRepresentationTypeAutomaticCallDistributor = 2
};

enum HistoryInfoEntryType
{
    ucHistoryInfoTypeNone = 0,
    ucHistoryInfoTypeRedirectStatic = 1,
    ucHistoryInfoTypeRedirectHuntgroup = 2,
    ucHistoryInfoTypeRetargetDelegate = 3,
    ucHistoryInfoTypeRetargetTeam = 4,
    ucHistoryInfoTypeRetargetTransferer = 5,
    ucHistoryInfoTypeRetargetTransferee = 6,
    ucHistoryInfoTypeUnknown = 7
};

enum ConversationContextUpdateOptions
{
    ucConversationContextUpdateNone = 0,
    ucConversationContextUpdateAssociatedInfo = 1
};

enum PreviousConversationHistoryRetrievalState
{
    ucPreviousConversationHistoryRetrievalNotStarted = 0,
    ucPreviousConversationHistoryRetrievalInProgress = 1,
    ucPreviousConversationHistoryRetrievalComplete = 2
};

enum ParticipantEventStateChangeAlsoRingingOthers
{
    ucAlsoRingingOthersAdmins = 0,
    ucAlsoRingingOthersTeam = 1
};

enum ModalityActionProperty
{
    ucModalityTransferContextFailures = 1611464704,
    ucModalityConnectContextFailures = 1611464705
};

enum ModalityProperty
{
    ucInstantMessageModalitySharedCapabilities = 537591808,
    ucInstantMessageModalityInviteMessage = 720897,
    ucInstantMessageModalityNoAutoAccept = 269156354,
    ucInstantMessageModalityInviteForked = 269156355,
    ucAVModalityAudioRenderVolume = 537591812,
    ucAVModalityAudioRenderMute = 269156357,
    ucAVModalityAudioRenderIntensity = 537591814,
    ucAVModalityAudioRenderIntensityMin = 537591815,
    ucAVModalityAudioRenderIntensityMax = 537591816,
    ucAVModalityAudioCaptureVolume = 537591817,
    ucAVModalityAudioCaptureMute = 269156362,
    ucAVModalityAudioCaptureIntensity = 537591819,
    ucAVModalityAudioCaptureIntensityMin = 537591820,
    ucAVModalityAudioCaptureIntensityMax = 537591821,
    ucAVModalityAudioCanCallComputer = 269156366,
    ucAVModalityAudioCanCallTelephone = 269156367,
    ucAVModalityAudioRenderDevice = 720912,
    ucAVModalityAudioCaptureDevice = 720913,
    ucAVModalityAudioSpeakerMode = 806027282,
    ucAVModalityAudioHistoryInfo = 1074462739,
    ucAVModalityAudioOriginalTargetUri = 720916,
    ucAVModalityAudioReferredByUri = 720917,
    ucAVModalityAudioSupportsAdvancedTransfer = 269156374,
    ucAVModalityAudioQuality = 806027287,
    ucAVModalityAudioOnHoldSource = 720920,
    ucAVModalityAudioOnHoldRepeated = 269156377,
    ucAVModalityAudioRemoteHoldFeature = 806027290,
    ucAVModalityIsAutoAccepted = 269156379,
    ucAVModalityAudioSupportsSafeTransfer = 269156380,
    ucAVModalityAudioNetworkQuality = 806027293,
    ucAVModalityAudioBandwidthQuality = 806027294,
    ucAVModalityAudioDeviceQuality = 806027295,
    ucAVModalityVoicemailItemId = 720928,
    ucAVModalityVideoNetworkQuality = 806027297,
    ucAVModalityVideoBandwidthQuality = 806027298,
    ucAVModalityVideoDeviceQuality = 806027299,
    ucAVModalityAudioIsEmergencyCall = 806027300,
    ucAVModalityUsbOffHookLedStatus = 269156389,
    ucAVModalityUsbDisplayName = 720934,
    ucAVModalityAudioLocationSent = 1074495488,
    ucAVModalityAudioPrivateCall = 269189121,
    ucAVModalityAudioIsPersistentSession = 269189122,
    ucAVModalityAudioCannotDisconnectSession = 269189123,
    ucAVModalityAudioIsAudioMCUModeAllowed = 269189124,
    ucAVModalityAudioMCUSession = 269189125,
    ucAVModalityAudioRemoteSupportsReferExtensions = 269189126,
    ucAVModalityAudioRemoteSupportsSafeTransfer = 269189127,
    ucAVModalityAudioStartTime = 537624584,
    ucAVModalityAudioSupportsReplaces = 269189129,
    ucAVModalityExtendedAudioQuality = 1074495498,
    ucAVModalityVideoCapturePaused = 269189131,
    ucAVModalityVideoRenderPaused = 269189132,
    ucAVModalityVideoIsAVMCUModeAllowed = 269189133,
    ucAVModalityVideoStartTime = 537624590,
    ucAVModalityExtendedVideoQuality = 1074495503,
    ucAVModalityDTMFDigits = 753680,
    ucAVModalityUsbConvId = 753681,
    ucContentModalityPresentedContentIdAtDisconnect = 537624594
};

enum ConversationProperty
{
    ucConversationId = 524288,
    ucConversationSubject = 524289,
    ucConversationImportance = 805830658,
    ucConversationTransferredBy = 524291,
    ucConversationReplaced = 524292,
    ucConversationConferencingUri = 524293,
    ucConversationRepresentedBy = 1074266118,
    ucConversationConferenceInviterRepresentationInfo = 1074266119,
    ucConversationFollowUp = 268959752,
    ucConversationDirection = 805830665,
    ucConversationConferenceAcceptingParticipant = 1074266122,
    ucConversationPreviousConversationId = 524299,
    ucConversationAcceptanceState = 805830668,
    ucConversationIsUsbConversation = 268959757,
    ucConversationAutoTerminateOnIdle = 268959758,
    ucConversationConferenceEscalationProgress = 805830671,
    ucConversationConferenceEscalationResult = 537395216,
    ucConversationConferencingInvitedModes = 537395217,
    ucConversationInviter = 1074266130,
    ucConversationConferencingLocked = 268959763,
    ucConversationConferencingFirstInstantMessage = 524308,
    ucConversationConferenceAccessInformation = 1074266133,
    ucConversationConferencingAccessType = 537395222,
    ucConversationCallParkOrbit = 1074266136,
    ucConversationConferenceDisclaimer = 1074266137,
    ucConversationConferenceDisclaimerAccepted = 537395226,
    ucConversationConferenceTerminateOnLeave = 268959771,
    ucConversationNumberOfParticipantsRecording = 537395228,
    ucConversationConferenceJoinDialogCompleted = 268959773,
    ucConversationLastActivityTimeStamp = 1342701598,
    ucConversationConferenceDialogId = 524319,
    ucConversationConferenceDialogFromTag = 524320,
    ucConversationConferenceDialogToTag = 524321,
    ucConversationConferenceEndorseEnabled = 268959778,
    ucConversationConferenceHardMute = 268959779,
    ucConversationConferenceAutoPromoteLevel = 537427969,
    ucConversationConferencePermittedAutoPromoteLevels = 537427970,
    ucConversationConferencePSTNBypassEnabled = 268992515,
    ucConversationConferencePermissionManager = 1074298884,
    ucConversationPresentedItem = 1074298885,
    ucConversationActivePresenter = 1074298886,
    ucConversationConferenceGlobalAnnouncements = 268992519,
    ucConversationViewedItem = 1074298888,
    ucConversationPresentationState = 1074298889,
    ucConversationConferenceIsRosterLimited = 268992522
};

enum ParticipantProperty
{
    ucParticipantRepresentedBy = 1074397184,
    ucParticipantReferredByUri = 1074397185,
    ucParticipantName = 655362,
    ucParticipantIsAnonymous = 269090819,
    ucParticipantIsAuthenticated = 269090820,
    ucParticipantIsInLobby = 269090821,
    ucParticipantIsLeader = 269090822,
    ucParticipantIsCallParkService = 269090823,
    ucParticipantParkedCall = 655368,
    ucParticipantSoftMuteChangeInitiator = 805961737,
    ucParticipantIsRecording = 269090826,
    ucParticipantIsPresenter = 269090827,
    ucParticipantEndorserName = 655372,
    ucParticipantEndorserUri = 655373,
    ucParticipantIsPinned = 269090830,
    ucParticipantIsLocked = 269090831,
    ucParticipantAnonymizedResponseGroupAgent = 269123585,
    ucParticipantAsOrganizer = 269123586
};

enum ConversationStateChangeProperty
{
    ucConversationStateChangeUnparkReason = 805896192,
    ucConversationStateChangeUnparkTarget = 589825
};

enum ModalityStateChangeProperty
{
    ucModalityStateChangeForwardHistoryInfo = 1074528256,
    ucModalityStateChangeAcceptingParticipant = 1074528257,
    ucModalityStateChangeAcceptingEndpoint = 1074528258,
    ucModalityStateChangeTransferTargetParticipant = 1074528259,
    ucModalityStateChangeTransferTargetEndpoint = 1074528260,
    ucModalityStateChangeAlsoRingingOthers = 806092805,
    ucModalityStateChangeIsAutoAccepted = 269221894,
    ucModalityStateChangeAllowedRedirection = 537657351,
    ucModalityStateChangeInvitingDevice = 1074528264,
    ucModalityStateChangeInvitingParticipant = 1074561025,
    ucModalityStateChangeInvitedParticipant = 1074561027,
    ucModalityStateChangeInvitedDevice = 1074561028,
    ucModalityStateChangeTransferFromParticipant = 1074561029,
    ucModalityStateChangeTransferFromDevice = 1074561030,
    ucModalityStateChangeTransferToParticipant = 1074561031,
    ucModalityStateChangeTransferedParticipant = 1074561032,
    ucModalityStateChangeTransferedDevice = 1074561033,
    ucModalityStateChangeForwardFromParticipant = 1074561034,
    ucModalityStateChangeForwardFromDevice = 1074561035,
    ucModalityStateChangeForwardToParticipant = 1074561036,
    ucModalityStateChangeForwardToDevice = 1074561037,
    ucModalityStateChangeRemoteParticipantInP2PSession = 1074561038,
    ucModalityStateChangeRemoteDeviceInP2PSession = 1074561039,
    ucModalityStateChangeInviteMessage = 819216,
    ucModalityStateChangeInviteMessageMode = 537690129,
    ucModalityStateChangeAllModes = 537690130,
    ucModalityStateChangeInitiator = 269254675,
    ucModalityStateChangeMovedToNewConversation = 269254676,
    ucModalityStateChangeSessionReplaced = 269254677,
    ucModalityStateChangeCallInitiatedFromDevice = 269254678,
    ucModalityStateChangeFailedToConnectToFocus = 269254679,
    ucModalityStateChangePrevState = 537690136,
    ucModalityStateChangePersonalTransfer = 269254681,
    ucModalityStateChangeReplacesSession = 537690138,
    ucModalityStateChangeTransferContext = 1074561051,
    ucModalityStateChangeTransferSession = 537690140,
    ucModalityStateChangeTransferToDeviceExact = 1074561053,
    ucModalityStateChangeTransferToConvParticipant = 1074561054,
    ucModalityStateChangeDevice = 1074561055
};

enum RedirectionTypes
{
    ucRedirectionTypesNone = 0,
    ucRedirectionTypesReplyWithInstantMessage = 1,
    ucRedirectionTypesReplyWithAudioVideo = 2,
    ucRedirectionTypesForwardToVoiceMail = 4,
    ucRedirectionTypesForwardToContact = 8
};

enum InstantMessageCapabilities
{
    ucInstantMessageCapabilityCanRenderIsf = 1,
    ucInstantMessageCapabilityCanRenderGif = 2,
    ucInstantMessageCapabilityCanRenderRtf = 4,
    ucInstantMessageCapabilitySupportMime = 8,
    ucInstantMessageCapabilityCanRenderHtml = 16
};

enum SpeakerMode
{
    ucSpeakerOff = 0,
    ucSpeakerOn = 1
};

enum HoldFeature
{
    ucHoldFeatureNone = 0,
    ucHoldFeatureMusicOnHold = 1,
    ucHoldFeatureOther = 2
};

enum MediaQuality
{
    ucMediaQualityNotSet = 0,
    ucMediaQualityUnknown = 1,
    ucMediaQualityGood = 2,
    ucMediaQualityFair = 3,
    ucMediaQualityBad = 4
};

enum PreviousConversationsType
{
    ucPreviousConversationsMissedConversations = 0,
    ucPreviousConversationsConversationHistory = 1,
    ucPreviousConversationsAudioCallLogs = 2,
    ucPreviousConversationsFollowUpConversations = 3,
    ucPreviousConversationsMissedAudioCalls = 4,
    ucPreviousConversationsTypeMax = 5
};

enum PreviousConversationsErrorType
{
    ucPreviousConversationsErrorCritical = 0,
    ucPreviousConversationsErrorTransient = 1,
    ucPreviousConversationsErrorNotConfigured = 2
};

enum PreviousConversationsConnectionState
{
    ucPreviousConversationsDisconnected = 0,
    ucPreviousConversationsConnected = 1
};

enum CallUnparkReason
{
    ucUnparkReasonNone = 0,
    ucUnparkReasonAbandoned = 1,
    ucUnparkReasonDisconnected = 2,
    ucUnparkReasonFallback = 3,
    ucUnparkReasonRetrieved = 4,
    ucUnparkReasonRingback = 5,
    ucUnparkReasonUnknown = 6
};

enum SoftMuteChangeInitiator
{
    ucMuteInitiatorLocal = 0,
    ucMuteInitiatorRemote = 1,
    ucMuteInitiatorJoin = 2
};

enum InstantMessageContentType
{
    ucInstantMessageContentPlainText = 0,
    ucInstantMessageContentHtml = 1,
    ucInstantMessageContentRichText = 2,
    ucInstantMessageContentGif = 3,
    ucInstantMessageContentInk = 4,
    ucInstantMessageContentUnknown = 5
};

enum ContextType
{
    ucContextualConversationApplicationId = 0,
    ucContextualConversationApplicationData = 1,
    ucContextualConversationHyperLink = 2
};

enum SharingResourceType
{
    ucSharingResourceTypeDesktop = 0,
    ucSharingResourceTypeMonitor = 1,
    ucSharingResourceTypeProcess = 2,
    ucSharingResourceTypeWindow = 3,
    ucSharingResourceTypeOther = 4
};

enum ParticipationState
{
    ucParticipationStateNone = 0,
    ucParticipationStateViewing = 1,
    ucParticipationStateRequestingControl = 2,
    ucParticipationStateControlling = 3,
    ucParticipationStateSharing = 4
};

enum CanShareDetail
{
    ucCanShareDetailAllowed = 0,
    ucCanShareDetailDisabledByPolicy = 1,
    ucCanShareDetailDisabledByOrganizerPolicy = 2,
    ucCanShareDetailDisabledByRole = 3,
    ucCanShareDetailCannotConnect = 4,
    ucCanShareDetailOther = 5
};

struct __declspec(uuid("62074904-8d06-43fe-a531-e63df7fdc2e7"))
IModalityPropertyDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ModalityProperty _propertyType,
        /*[out,retval]*/ VARIANT * _itemValue ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ModalityProperty _propertyType,
        /*[out]*/ VARIANT * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ModalityProperty * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ModalityProperty _propertyType,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("8665fd95-7720-4f9e-b605-6abcbd7edff4"))
IParticipantPropertyDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ParticipantProperty _propertyType,
        /*[out,retval]*/ VARIANT * _itemValue ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ParticipantProperty _propertyType,
        /*[out]*/ VARIANT * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ParticipantProperty * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ParticipantProperty _propertyType,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("b3ff95bd-b743-498c-9973-7f8a2a78576a"))
IParticipantInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * _uri ) = 0;
      virtual HRESULT __stdcall get_IsRepresenting (
        /*[out,retval]*/ VARIANT_BOOL * _isRequired ) = 0;
      virtual HRESULT __stdcall get_RepresentedUri (
        /*[out,retval]*/ BSTR * _representedUriString ) = 0;
      virtual HRESULT __stdcall get_DisplayName (
        /*[out,retval]*/ BSTR * _name ) = 0;
      virtual HRESULT __stdcall put_DisplayName (
        /*[in]*/ BSTR _name ) = 0;
      virtual HRESULT __stdcall get_RepresentedDisplayName (
        /*[out,retval]*/ BSTR * _name ) = 0;
      virtual HRESULT __stdcall put_RepresentedDisplayName (
        /*[in]*/ BSTR _name ) = 0;
};

struct __declspec(uuid("00e22cbb-3170-453a-ae62-eafbc75a9f8d"))
IConversationPropertyDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ConversationProperty _propertyType,
        /*[out,retval]*/ VARIANT * _itemValue ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ConversationProperty _propertyType,
        /*[out]*/ VARIANT * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ConversationProperty * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ConversationProperty _propertyType,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("e24c2484-c70d-44ab-8a4e-f1585a17a3e1"))
IVoiceMailsManager : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ConnectionState (
        /*[out,retval]*/ enum PreviousConversationsConnectionState * _connectionState ) = 0;
      virtual HRESULT __stdcall Connect ( ) = 0;
      virtual HRESULT __stdcall Reset ( ) = 0;
      virtual HRESULT __stdcall get_NewVoiceMailCountSinceLastReset (
        /*[out,retval]*/ unsigned long * _count ) = 0;
      virtual HRESULT __stdcall get_MaxItems (
        /*[out,retval]*/ unsigned long * _maxItems ) = 0;
      virtual HRESULT __stdcall put_MaxItems (
        /*[in]*/ unsigned long _maxItems ) = 0;
      virtual HRESULT __stdcall GetUnreadVoiceMails (
        /*[in]*/ IUnknown * _voiceMailsManagerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetAllVoiceMails (
        /*[in]*/ IUnknown * _voiceMailsManagerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("acabfa1c-0738-4c4d-af7e-326873cf4ace"))
IPreviousConversationsManager : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum PreviousConversationsType * _type ) = 0;
      virtual HRESULT __stdcall get_ConnectionState (
        /*[out,retval]*/ enum PreviousConversationsConnectionState * _connectionState ) = 0;
      virtual HRESULT __stdcall Connect ( ) = 0;
      virtual HRESULT __stdcall Reset ( ) = 0;
      virtual HRESULT __stdcall get_NewItemCountSinceLastReset (
        /*[out,retval]*/ unsigned long * _count ) = 0;
      virtual HRESULT __stdcall get_AccessTime (
        /*[out,retval]*/ DATE * AccessTime ) = 0;
      virtual HRESULT __stdcall get_MaxItems (
        /*[out,retval]*/ unsigned long * _maxItems ) = 0;
      virtual HRESULT __stdcall put_MaxItems (
        /*[in]*/ unsigned long _maxItems ) = 0;
      virtual HRESULT __stdcall GetUnreadItems (
        /*[in]*/ IUnknown * _previousConversationsManagerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetAllItems (
        /*[in]*/ IUnknown * _previousConversationsManagerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("15d4461d-c7eb-45cb-b214-f1c45fed1db2"))
IPreviousConversationsManagerDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum PreviousConversationsType _key,
        /*[out,retval]*/ struct IPreviousConversationsManager * * _pastConversations ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum PreviousConversationsType * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IPreviousConversationsManager * * _value ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum PreviousConversationsType _type,
        /*[out]*/ struct IPreviousConversationsManager * * _pastConversations,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum PreviousConversationsType _type,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("209da899-15f1-499a-b8b5-a847eaa899ae"))
IRepresentationInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum RepresentationType * _type ) = 0;
      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * _uri ) = 0;
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
        /*[in]*/ __int64 Owner ) = 0;
      virtual HRESULT __stdcall get_Owner (
        /*[out,retval]*/ __int64 * Owner ) = 0;
      virtual HRESULT __stdcall put_MessageDrain (
        /*[in]*/ __int64 Drain ) = 0;
      virtual HRESULT __stdcall get_MessageDrain (
        /*[out,retval]*/ __int64 * Drain ) = 0;
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
        /*[in]*/ __int64 hwnd,
        /*[in]*/ long uMsg,
        /*[in]*/ __int64 wParam,
        /*[in]*/ __int64 lParam ) = 0;
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

struct __declspec(uuid("3690d374-3779-4efc-b8e8-20642c01529a"))
ISharingResource : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum SharingResourceType * _val ) = 0;
      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ int * _id ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * _name ) = 0;
};

struct __declspec(uuid("6a7ff112-42ce-4d03-8922-d02ac6c5759f"))
ISharingResourceList : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long index,
        /*[out,retval]*/ struct ISharingResource * * _value ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall Add (
        /*[in]*/ struct ISharingResource * _resource ) = 0;
      virtual HRESULT __stdcall Remove (
        /*[in]*/ struct ISharingResource * _resource ) = 0;
      virtual HRESULT __stdcall get_IsShareable (
        /*[out,retval]*/ VARIANT_BOOL * _isShareable ) = 0;
};

struct __declspec(uuid("7b39cb77-4aeb-42d7-b351-cb5472c1c6ad"))
IChannel2 : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall Pause (
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Resume (
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("c79e4169-535f-4a41-8de2-c65a90d9503d"))
IModalityStateChangePropertyDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ModalityStateChangeProperty _modalityType,
        /*[out,retval]*/ VARIANT * _itemValue ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ModalityStateChangeProperty _modalityType,
        /*[out]*/ VARIANT * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ModalityStateChangeProperty * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ModalityStateChangeProperty _modalityType,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("478dda59-64c0-4429-ac8a-f10a940b52cc"))
IModalityStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum ModalityState * _oldState ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum ModalityState * _newState ) = 0;
      virtual HRESULT __stdcall get_StatusCode (
        /*[out,retval]*/ long * _statusCode ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IModalityStateChangePropertyDictionary * * _properties ) = 0;
      virtual HRESULT __stdcall get_ExtendedStatusCode (
        /*[out,retval]*/ long * _extendedStatusCode ) = 0;
};

struct __declspec(uuid("deddd7bd-4763-41d2-9aaa-b2c143457ca4"))
IModalityActionAvailabilityChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Action (
        /*[out,retval]*/ enum ModalityAction * _action ) = 0;
      virtual HRESULT __stdcall get_IsAvailable (
        /*[out,retval]*/ VARIANT_BOOL * _isAvailable ) = 0;
};

struct __declspec(uuid("56f7dec0-59ca-47c6-9f35-d5066a702b39"))
IConversationStateChangePropertyDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ConversationStateChangeProperty _modalityType,
        /*[out,retval]*/ VARIANT * _itemValue ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ConversationStateChangeProperty _propertyType,
        /*[out]*/ VARIANT * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ConversationStateChangeProperty * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ConversationStateChangeProperty _propertyType,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("2a9385aa-cf61-4e47-b64c-214de4a4ad11"))
IConversationStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum ConversationState * _oldState ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum ConversationState * _newState ) = 0;
      virtual HRESULT __stdcall get_StatusCode (
        /*[out,retval]*/ long * _statusCode ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IConversationStateChangePropertyDictionary * * _properties ) = 0;
};

struct __declspec(uuid("88c688b0-3908-4c56-a2e8-f90ab705c536"))
IConversationActionAvailabilityEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Action (
        /*[out,retval]*/ enum ConversationAction * _action ) = 0;
      virtual HRESULT __stdcall get_IsAvailable (
        /*[out,retval]*/ VARIANT_BOOL * _isAvailable ) = 0;
};

struct __declspec(uuid("905b727d-fd8f-46ec-8006-ce245f6835ee"))
IMessageSentEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Text (
        /*[out,retval]*/ BSTR * _text ) = 0;
      virtual HRESULT __stdcall get_Formats (
        /*[out,retval]*/ struct INamedPropertyList * * _formats ) = 0;
};

struct __declspec(uuid("dba05c15-1c07-4a76-8248-08d8416a24e3"))
IIsTypingChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_IsTyping (
        /*[out,retval]*/ VARIANT_BOOL * _isTyping ) = 0;
};

struct __declspec(uuid("d5c1c8fb-dc8a-45a3-aaa3-53b89d43cde4"))
IIsContributingChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_IsContributing (
        /*[out,retval]*/ VARIANT_BOOL * _isContributing ) = 0;
};

struct __declspec(uuid("2495b94c-38af-439b-bba0-0ad38d959ce4"))
IChannelStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum ChannelState * _oldState ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum ChannelState * _newState ) = 0;
};

struct __declspec(uuid("9d4d3d10-debc-40c6-acf9-527986cb42bc"))
IChannelActionAvailabilityEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Action (
        /*[out,retval]*/ enum ChannelAction * _action ) = 0;
      virtual HRESULT __stdcall get_IsAvailable (
        /*[out,retval]*/ VARIANT_BOOL * _isAvailable ) = 0;
};

struct __declspec(uuid("3debdc7e-811f-4304-bdd7-7d98c732026c"))
IHistoryInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum HistoryInfoEntryType * _type ) = 0;
      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * _uri ) = 0;
      virtual HRESULT __stdcall get_Status (
        /*[out,retval]*/ long * _statusCode ) = 0;
      virtual HRESULT __stdcall get_NestedHistory (
        /*[out,retval]*/ SAFEARRAY * * _histories ) = 0;
};

struct __declspec(uuid("7cfe77cd-731d-48b2-82b1-eca3414d62e3"))
IConversationPropertyChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Property (
        /*[out,retval]*/ enum ConversationProperty * _propertyType ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ VARIANT * _propertyValue ) = 0;
};

struct __declspec(uuid("70dcce38-90bf-4428-a34f-3a6082f29e50"))
IParticipantPropertyChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Property (
        /*[out,retval]*/ enum ParticipantProperty * _propertyType ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ VARIANT * _propertyValue ) = 0;
};

struct __declspec(uuid("4290152c-75c9-47c8-aa12-814818360b3f"))
IModalityPropertyChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Property (
        /*[out,retval]*/ enum ModalityProperty * _propertyType ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ VARIANT * _propertyValue ) = 0;
};

struct __declspec(uuid("df414a68-5051-4465-aae2-4f301315734e"))
IPreviousConversation : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * _entryId ) = 0;
      virtual HRESULT __stdcall get_From (
        /*[out,retval]*/ BSTR * _remoteParty ) = 0;
      virtual HRESULT __stdcall get_FromUri (
        /*[out,retval]*/ BSTR * _uri ) = 0;
      virtual HRESULT __stdcall get_ConferenceUri (
        /*[out,retval]*/ BSTR * _uri ) = 0;
      virtual HRESULT __stdcall get_ModalityTypes (
        /*[out,retval]*/ enum ModalityTypes * _modalities ) = 0;
      virtual HRESULT __stdcall get_Direction (
        /*[out,retval]*/ enum ConversationDirection * _direction ) = 0;
      virtual HRESULT __stdcall get_Subject (
        /*[out,retval]*/ BSTR * _subject ) = 0;
      virtual HRESULT __stdcall get_ReceiveTime (
        /*[out,retval]*/ DATE * _received ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct INamedPropertyList * * _properties ) = 0;
      virtual HRESULT __stdcall get_IsRead (
        /*[out,retval]*/ VARIANT_BOOL * _isRead ) = 0;
};

struct __declspec(uuid("81c9d13f-a4f9-4e13-92d3-bb271e8df3d2"))
IPreviousConversationCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IPreviousConversation * * _previousConversation ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
};

struct __declspec(uuid("f87afd9a-9bf8-4ac4-953e-33c4d2035d33"))
IPreviousConversationsManagerConnectionStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ConnectionState (
        /*[out,retval]*/ enum PreviousConversationsConnectionState * _connectionState ) = 0;
      virtual HRESULT __stdcall get_Errors (
        /*[out,retval]*/ SAFEARRAY * * _errors ) = 0;
};

struct __declspec(uuid("2c59a886-acd3-46bb-8bac-7c1c59fb2e0d"))
IPreviousConversationsManagerNewItemCountChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_NewItemCountSinceLastReset (
        /*[out,retval]*/ unsigned long * _count ) = 0;
};

struct __declspec(uuid("ebd19735-0cb5-4af3-8e05-b0140cf9d545"))
IPreviousConversationBatchEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Items (
        /*[out,retval]*/ struct IPreviousConversationCollection * * _value ) = 0;
};

struct __declspec(uuid("2f8c3e58-436b-42db-8924-6c394b37dca2"))
IConferenceAccessInformation : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * _id ) = 0;
      virtual HRESULT __stdcall get_AutoAttendantNumbers (
        /*[out,retval]*/ SAFEARRAY * * _autoAttendantNumbers ) = 0;
      virtual HRESULT __stdcall get_AdmissionKey (
        /*[out,retval]*/ BSTR * _admissionKey ) = 0;
      virtual HRESULT __stdcall get_InternalUrl (
        /*[out,retval]*/ BSTR * _cwaInternalUrl ) = 0;
      virtual HRESULT __stdcall get_ExternalUrl (
        /*[out,retval]*/ BSTR * _cwaExternalUrl ) = 0;
};

struct __declspec(uuid("1124aa59-b6f3-4eb1-b79a-e7f53b388b8d"))
IParticipantActionAvailabilityChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Action (
        /*[out,retval]*/ enum ParticipantAction * _action ) = 0;
      virtual HRESULT __stdcall get_IsAvailable (
        /*[out,retval]*/ VARIANT_BOOL * _isAvailable ) = 0;
};

struct __declspec(uuid("37c7f935-e215-46d2-82a3-a451840ec733"))
IConversation2 : IDispatch
{};

struct __declspec(uuid("d60b1701-766d-401e-8586-83e0c9106be0"))
IParticipant2 : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall PinVideo (
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall UnPinVideo (
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall LockVideo (
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall UnLockVideo (
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("25025b1b-0084-44d1-b383-9ffc9a99abcd"))
_IConversationManagerEvents : IDispatch
{};

struct __declspec(uuid("c63f58cb-ef11-4934-aa3b-2d69244f0780"))
_IConversationEvents : IDispatch
{};

struct __declspec(uuid("92d7e7a8-48d8-4166-8911-630ae02b2b93"))
IInitialContextEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ApplicationId (
        /*[out,retval]*/ BSTR * _appGuid ) = 0;
      virtual HRESULT __stdcall get_ApplicationData (
        /*[out,retval]*/ BSTR * _appData ) = 0;
};

struct __declspec(uuid("21629ab6-9557-4985-9237-49177f618692"))
IContextEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ApplicationId (
        /*[out,retval]*/ BSTR * _appGuid ) = 0;
      virtual HRESULT __stdcall get_ContextDataType (
        /*[out,retval]*/ BSTR * _dataType ) = 0;
      virtual HRESULT __stdcall get_ContextData (
        /*[out,retval]*/ BSTR * _data ) = 0;
};

struct __declspec(uuid("6a986b5d-e5ec-4395-ad2d-0e62b72da776"))
_IParticipantEvents : IDispatch
{};

struct __declspec(uuid("5418e2ab-eb9a-4659-b4dc-28de633b2b8f"))
IMutedChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_IsMuted (
        /*[out,retval]*/ VARIANT_BOOL * _isMuted ) = 0;
};

struct __declspec(uuid("10d66b39-5e4b-4c98-aac6-d0458feb1bdb"))
_IModalityEvents : IDispatch
{};

struct __declspec(uuid("4fe44049-4e44-4109-b234-4e4efc135a86"))
_IInstantMessageModalityEvents : IDispatch
{};

struct __declspec(uuid("acd2fdb9-98ed-4283-aae8-3cd989ebd9ae"))
_IAVModalityEvents : IDispatch
{};

struct __declspec(uuid("be026cd2-7e82-4f7c-8762-f6b02f496174"))
_IApplicationSharingModalityEvents : IDispatch
{};

struct __declspec(uuid("b22edbea-9e61-4703-82be-01c05619b6d6"))
ILocalSharedResourcesChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ResourceList (
        /*[out,retval]*/ struct ISharingResourceList * * _resources ) = 0;
};

struct __declspec(uuid("8fe9a23d-72bc-4ea3-a5af-e37f75ead8a2"))
IParticipationStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum ParticipationState * _oldState ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum ParticipationState * _newState ) = 0;
};

struct __declspec(uuid("4e8dc7e0-04b8-470b-bdfa-f520099b975f"))
IControllerChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ControllerName (
        /*[out,retval]*/ BSTR * _displayName ) = 0;
};

struct __declspec(uuid("d0dbb17d-565e-486c-a47b-ba32dc1fad0a"))
ISharerChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_SharerName (
        /*[out,retval]*/ BSTR * _displayName ) = 0;
};

struct __declspec(uuid("fa7971ef-1e41-429f-8c79-dac5f88d444b"))
_IConversationHistoryEvents : IDispatch
{};

struct __declspec(uuid("4c675532-6d55-49dd-8c8e-3c1cc5624cf2"))
IConversationHistoryItemAddedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ItemIndex (
        /*[out,retval]*/ long * _itemIndex ) = 0;
};

struct __declspec(uuid("464d5228-9f68-4b1f-b430-156a104e2b85"))
IPreviousHistoryRetrievalStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum PreviousConversationHistoryRetrievalState * _oldState ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum PreviousConversationHistoryRetrievalState * _newState ) = 0;
      virtual HRESULT __stdcall get_StatusCode (
        /*[out,retval]*/ long * _statusCode ) = 0;
};

struct __declspec(uuid("9397ff55-ee06-4f02-8f2a-be3ae249d4bb"))
IConversationHistoryActionAvailabilityEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Action (
        /*[out,retval]*/ enum ConversationHistoryAction * _action ) = 0;
      virtual HRESULT __stdcall get_IsAvailable (
        /*[out,retval]*/ VARIANT_BOOL * _isAvailable ) = 0;
};

struct __declspec(uuid("d992371e-5161-453b-97e6-6e7c67bc075e"))
_IPreviousConversationsManagerEvents : IDispatch
{};

struct __declspec(uuid("b0c10f94-cab6-4821-9643-d781885a46af"))
_IVoiceMailsManagerEvents : IDispatch
{};

struct __declspec(uuid("0ca9bae8-dd28-4929-a821-4d00f11203cc"))
_IChannelEvents : IDispatch
{};

struct __declspec(uuid("57cc2235-96c5-49fa-b92b-350486c6cf52"))
_IConversationCallback : IDispatch
{};

struct __declspec(uuid("26c8d96d-7d22-4e9e-948a-edccb4ca9c64"))
_IParticipantCallback : IDispatch
{};

struct __declspec(uuid("3f60cff7-be3f-4404-8395-c02d0cdd318e"))
_IModalityCallback : IDispatch
{};

struct __declspec(uuid("8684d3ad-7061-4056-a894-ebf1270adb9b"))
_IApplicationSharingModalityCallback : IDispatch
{};

struct __declspec(uuid("5fa92ea7-6e6e-4a82-8f0d-107feaf5a75d"))
_IConversationHistoryCallback : IDispatch
{};

struct __declspec(uuid("95b9a382-ea50-45e4-965e-ba61e7213dba"))
_IAudioChannelCallback : IDispatch
{};

struct __declspec(uuid("08e1e9cf-a7e4-43be-ae8e-b223ce71e46e"))
_IChannelCallback : IDispatch
{};

struct __declspec(uuid("68b1fc26-b48f-4d5f-b041-8e246676d836"))
_IVoiceMailCallback : IDispatch
{};

struct __declspec(uuid("d9d04b55-d820-4d5b-a690-658a49368478"))
_IVoiceMailsManagerCallback : IDispatch
{};

struct __declspec(uuid("4bacb4b8-1236-42b8-bda1-d1533148780d"))
_IPreviousConversationsManagerCallback : IDispatch
{};

struct __declspec(uuid("c87287d3-0a25-414d-9a92-16a11a2099f0"))
ConversationManager;
    // [ default ] interface IConversationManager
    // [ default, source ] dispinterface _IConversationManagerEvents

struct __declspec(uuid("457d962f-5f66-4429-bd94-6f9f366fae0d"))
Conversation;
    // [ default ] interface IConversation
    // interface IConversation2
    // [ default, source ] dispinterface _IConversationEvents

struct __declspec(uuid("de5598bc-7188-415a-a3d9-7e86863e7a09"))
Participant;
    // [ default ] interface IParticipant
    // interface IParticipant2
    // [ default, source ] dispinterface _IParticipantEvents

struct __declspec(uuid("9582d285-7e6e-4bfa-bcbc-7de3ea3e6d74"))
InstantMessageModality;
    // [ default ] interface IInstantMessageModality
    // [ default, source ] dispinterface _IInstantMessageModalityEvents

struct __declspec(uuid("ba2bd6f3-7676-42e3-89c6-10ceb3f7e106"))
AVModality;
    // [ default ] interface IAVModality
    // [ default, source ] dispinterface _IAVModalityEvents

struct __declspec(uuid("556a0d7a-05fb-4a8b-b15a-ad3eacde3ad2"))
Channel;
    // [ default ] interface IChannel
    // interface IChannel2

struct __declspec(uuid("61419951-5d5a-4da3-be6a-076f9645229e"))
Modality;
    // [ default ] interface IModality

struct __declspec(uuid("f8f4a9b9-7e52-44e0-89e0-30a3b4da4d23"))
ApplicationSharingModality;
    // [ default ] interface IApplicationSharingModality
    // [ default, source ] dispinterface _IApplicationSharingModalityEvents

struct __declspec(uuid("d5975289-8d98-481a-92a7-f8536085d621"))
AudioChannel;
    // [ default ] interface IAudioChannel
    // interface IChannel2
    // [ default, source ] dispinterface _IChannelEvents

struct __declspec(uuid("56805856-acad-4d8c-aeaf-5576404bf72b"))
IsTypingChangedEventData;
    // [ default ] interface IIsTypingChangedEventData

struct __declspec(uuid("0af70da9-90e5-4128-8d67-a002cc94c53b"))
IsContributingChangedEventData;
    // [ default ] interface IIsContributingChangedEventData

struct __declspec(uuid("288926af-3015-4da0-b41c-9381e545eecf"))
ContextEventData;
    // [ default ] interface IContextEventData

struct __declspec(uuid("cf6796fe-c1d0-400f-85c1-3e8a0f5c247c"))
InitialContextEventData;
    // [ default ] interface IInitialContextEventData

struct __declspec(uuid("00068095-fa0a-4142-9f08-e0e5a9599621"))
ChannelStateChangedEventData;
    // [ default ] interface IChannelStateChangedEventData

struct __declspec(uuid("61e9f245-ce9a-4403-aad9-012b52ce151d"))
ChannelActionAvailabilityEventData;
    // [ default ] interface IChannelActionAvailabilityEventData

struct __declspec(uuid("6ec2948c-51d0-4806-baf1-218eecb20d08"))
StreamStateChangedEventData;
    // [ default ] interface IStreamStateChangedEventData

struct __declspec(uuid("f8946ac5-5eb8-4237-a1f5-1b03c6eaaac1"))
ConversationActionAvailabilityEventData;
    // [ default ] interface IConversationActionAvailabilityEventData

struct __declspec(uuid("36d057b0-0f7a-4c96-9e99-285073d50cc7"))
ModalityActionAvailabilityChangedEventData;
    // [ default ] interface IModalityActionAvailabilityChangedEventData

struct __declspec(uuid("2ede3b04-2f52-4ef0-b793-77397b6418de"))
ConversationContextCollectionEventData;
    // [ default ] interface IConversationContextCollectionEventData

struct __declspec(uuid("c42e7aba-dea0-43ec-96aa-6d8211cb2e14"))
SummarizerStatesChangedEventData;
    // [ default ] interface ISummarizerStatesChangedEventData

struct __declspec(uuid("cc1991bf-a8f3-48cb-94c2-ede7aa5545f3"))
ConversationManagerEventData;
    // [ default ] interface IConversationManagerEventData

struct __declspec(uuid("2d9dd8cd-381f-42b9-aa02-65df5e7d1fb1"))
ParticipantCollectionChangedEventData;
    // [ default ] interface IParticipantCollectionChangedEventData

struct __declspec(uuid("ee1abe27-10e4-4a70-b4dd-9a55149fd486"))
MessageSentEventData;
    // [ default ] interface IMessageSentEventData

struct __declspec(uuid("45076174-6cbd-471c-a8a7-19391a2c252c"))
ConversationStateChangedEventData;
    // [ default ] interface IConversationStateChangedEventData

struct __declspec(uuid("458858d8-9ac1-4431-b516-e44362fac7ad"))
ModalityStateChangedEventData;
    // [ default ] interface IModalityStateChangedEventData

struct __declspec(uuid("0d2b5648-d0cf-4069-b505-cd121319fdd7"))
ConversationPropertyChangedEventData;
    // [ default ] interface IConversationPropertyChangedEventData

struct __declspec(uuid("fe8d6343-3bdf-4ae7-a2bc-702f5f331118"))
ParticipantPropertyChangedEventData;
    // [ default ] interface IParticipantPropertyChangedEventData

struct __declspec(uuid("cc671df2-ecb2-4cf2-876e-ac14c25aa4fb"))
MutedChangedEventData;
    // [ default ] interface IMutedChangedEventData

struct __declspec(uuid("30e3bdff-0457-4695-a7a5-c63c36db069e"))
ParticipantActionAvailabilityChangedEventData;
    // [ default ] interface IParticipantActionAvailabilityChangedEventData

struct __declspec(uuid("91d5058a-77c3-40da-930c-f6207bd691c3"))
ModalityPropertyChangedEventData;
    // [ default ] interface IModalityPropertyChangedEventData

struct __declspec(uuid("373b822d-c7ee-4763-93a7-86bb2ea7e031"))
ConversationCollection;
    // [ default ] interface IConversationCollection

struct __declspec(uuid("3afa5d5f-c3d5-453f-96e0-30c3dbc34951"))
ConversationContext;
    // [ default ] interface IConversationContext

struct __declspec(uuid("931a066a-af82-4278-9084-4f8b517f9105"))
ConversationContextList;
    // [ default ] interface IConversationContextList

struct __declspec(uuid("00b1f05c-924f-4a74-aed2-8ede3d08de8b"))
ConversationHistory;
    // [ default ] interface IConversationHistory
    // [ default, source ] dispinterface _IConversationHistoryEvents

struct __declspec(uuid("bee2a581-7f14-4c61-89a6-47430ec138ea"))
ParticipantCollection;
    // [ default ] interface IParticipantCollection

struct __declspec(uuid("90b18225-0bc7-4120-84af-20056f7e7da4"))
VideoChannel;
    // [ default ] interface IVideoChannel
    // interface IChannel2
    // [ default, source ] dispinterface _IChannelEvents

struct __declspec(uuid("67c2d762-4e8d-45bf-94f0-68d99103567a"))
VideoChannelCollection;
    // [ default ] interface IVideoChannelCollection

struct __declspec(uuid("6c97f743-778b-460f-b14e-073290c4e7e4"))
VideoWindow;
    // [ default ] interface IVideoWindow

struct __declspec(uuid("1eabd900-9351-4306-936b-47dfb3170bc4"))
RepresentationInfo;
    // [ default ] interface IRepresentationInfo

struct __declspec(uuid("df54c8f4-06a1-426f-87d9-e2df86ff5e19"))
ParticipantInfo;
    // [ default ] interface IParticipantInfo

struct __declspec(uuid("14a62ae0-9d65-49f2-8331-ef85bff91244"))
HistoryInfo;
    // [ default ] interface IHistoryInfo

struct __declspec(uuid("378bf356-0ea3-4025-84fc-f64430adc784"))
CallParkOrbit;
    // [ default ] interface ICallParkOrbit

struct __declspec(uuid("7cc306e9-2546-4101-82ab-29613ad6f7a7"))
ICallParkOrbit : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_FriendlyOrbit (
        /*[out,retval]*/ BSTR * _orbit ) = 0;
      virtual HRESULT __stdcall get_SafeRetrieveUri (
        /*[out,retval]*/ BSTR * _uri ) = 0;
      virtual HRESULT __stdcall get_MsParkedCall (
        /*[out,retval]*/ BSTR * _msParkedCall ) = 0;
};

struct __declspec(uuid("ce124f88-3665-4a4a-b2e0-76355f832adf"))
ConversationPropertyDictionary;
    // [ default ] interface IConversationPropertyDictionary

struct __declspec(uuid("f251612d-4456-4acc-bf24-b0d79f004f0b"))
ModalityPropertyDictionary;
    // [ default ] interface IModalityPropertyDictionary

struct __declspec(uuid("939c4502-83de-4561-bd85-b56680c82428"))
ParticipantPropertyDictionary;
    // [ default ] interface IParticipantPropertyDictionary

struct __declspec(uuid("d56d0f3e-24e3-421a-8ae8-5bfb8d97fe90"))
ModalityDictionary;
    // [ default ] interface IModalityDictionary

struct __declspec(uuid("4a53d656-346e-4ddb-adab-7f504ce6c02c"))
ModalityStateChangePropertyDictionary;
    // [ default ] interface IModalityStateChangePropertyDictionary

struct __declspec(uuid("3e40c2d1-0a01-4d22-a5b8-3da4e7fd0170"))
ConversationStateChangePropertyDictionary;
    // [ default ] interface IConversationStateChangePropertyDictionary

struct __declspec(uuid("56d29009-c2f0-48b5-89c1-0f725199ecea"))
VoiceMail;
    // [ default ] interface IVoiceMail

struct __declspec(uuid("e5fbeb06-f18e-4351-b197-67cd6aa22c01"))
VoiceMailCollection;
    // [ default ] interface IVoiceMailCollection

struct __declspec(uuid("60a46701-3e4c-4115-83a3-b64eee5b4c0d"))
VoiceMailsManager;
    // [ default ] interface IVoiceMailsManager
    // [ default, source ] dispinterface _IVoiceMailsManagerEvents

struct __declspec(uuid("ca757729-da0f-452d-b001-936556bfc1d8"))
PreviousConversation;
    // [ default ] interface IPreviousConversation

struct __declspec(uuid("563b1f6b-f388-448c-b679-af1fa3a4329a"))
PreviousConversationCollection;
    // [ default ] interface IPreviousConversationCollection

struct __declspec(uuid("781a3b60-3cc1-42b2-8840-84e1a981800f"))
PreviousConversationsManager;
    // [ default ] interface IPreviousConversationsManager
    // [ default, source ] dispinterface _IPreviousConversationsManagerEvents

struct __declspec(uuid("4a7391d7-1a17-4996-9412-6828849395e8"))
PreviousConversationsManagerDictionary;
    // [ default ] interface IPreviousConversationsManagerDictionary

struct __declspec(uuid("d82fc397-451e-4e91-8557-d096d76792fe"))
PreviousConversationsManagerConnectionStateChangedEventData;
    // [ default ] interface IPreviousConversationsManagerConnectionStateChangedEventData

struct __declspec(uuid("d85b4d22-7e0a-40f6-ac97-fb09df0ff9d0"))
PreviousConversationsManagerNewItemCountChangedEventData;
    // [ default ] interface IPreviousConversationsManagerNewItemCountChangedEventData

struct __declspec(uuid("31d221f7-400e-4f04-8e3c-0bf253591d31"))
PreviousConversationBatchEventData;
    // [ default ] interface IPreviousConversationBatchEventData

struct __declspec(uuid("d5980e2d-0886-4ecd-9d2b-0057a6c061f3"))
VoiceMailBatchEventData;
    // [ default ] interface IVoiceMailBatchEventData

struct __declspec(uuid("19258088-aea6-4a73-b3d0-e4b1d9e11d9e"))
ConferenceAccessInformation;
    // [ default ] interface IConferenceAccessInformation

struct __declspec(uuid("c608ce35-98f7-41a6-951b-91c37d2d3851"))
ConversationHistoryActionAvailabilityEventData;
    // [ default ] interface IConversationHistoryActionAvailabilityEventData

struct __declspec(uuid("7eb20639-5a99-484e-99b3-c778d2f6699e"))
ConversationHistoryItemAddedEventData;
    // [ default ] interface IConversationHistoryItemAddedEventData

struct __declspec(uuid("8da3daf8-e6d0-4a19-9552-ef293f01d8a4"))
PreviousHistoryRetrievalStateChangedEventData;
    // [ default ] interface IPreviousHistoryRetrievalStateChangedEventData

struct __declspec(uuid("653a42da-8411-4902-8013-574f1aeec476"))
ConversationContextCollection;
    // [ default ] interface IConversationContextCollection

struct __declspec(uuid("ebd24819-d80a-44de-9c7b-606c512a1d71"))
SharingResource;
    // [ default ] interface ISharingResource

struct __declspec(uuid("f77d5a8d-8d57-40f0-ab49-85baa3699238"))
SharingResourceList;
    // [ default ] interface ISharingResourceList

struct __declspec(uuid("797d542e-5beb-4926-8145-653cb060366f"))
ControlRequestReceivedEventData;
    // [ default ] interface IControlRequestReceivedEventData

struct __declspec(uuid("ad0dfab8-d464-43c6-a8b4-f495c4cb8afe"))
LocalSharedResourcesChangedEventData;
    // [ default ] interface ILocalSharedResourcesChangedEventData

struct __declspec(uuid("29c7c658-9ca5-40aa-ac38-a25cb6f2ad74"))
ParticipationStateChangedEventData;
    // [ default ] interface IParticipationStateChangedEventData

struct __declspec(uuid("271bbe62-8678-48d2-807d-9bc9478c63d7"))
ControllerChangedEventData;
    // [ default ] interface IControllerChangedEventData

struct __declspec(uuid("26fdcf1d-abe4-46fc-a2da-474a47da8237"))
SharerChangedEventData;
    // [ default ] interface ISharerChangedEventData

enum AccessLevel
{
    ucAccessLevelDefault = 0,
    ucAccessLevelExternal = 100,
    ucAccessLevelColleague = 200,
    ucAccessLevelWorkgroup = 300,
    ucAccessLevelFriends = 400,
    ucAccessLevelBlocked = 32000
};

enum AccessEntryScope
{
    ucAccessEntryScopeUser = 0,
    ucAccessEntryScopeDomain = 1,
    ucAccessEntryScopeEnterprise = 2,
    ucAccessEntryScopeFederated = 3,
    ucAccessEntryScopePublic = 4,
    ucAccessEntryScopeEveryone = 5
};

enum PhotoSetting
{
    ucADPhotoDisplay = 0,
    ucADPhotoDoNotDisplay = 1,
    ucADPhotoDefault = 2
};

enum PublishableContactInformationType
{
    ucPresenceUpdateAvailability = 0,
    ucPresenceUpdateActivityId = 1,
    ucPresenceUpdateCustomActivityId = 2,
    ucPresenceUpdateLocationName = 3,
    ucPresenceUpdatePersonalNote = 4,
    ucPresenceUpdateDisplayPhoto = 5,
    ucPresenceUpdatePhotoUrl = 6,
    ucPublishablePresenceItemTypeMax = 7
};

enum SettingItemTypeInternal
{
    ucPC2PCVideoEnabled = 0,
    ucAVConferencingEnable = 1,
    ucAudioSetting = 2,
    ucAllowedAppSharingLevel = 3,
    ucEnableAppSharing = 4,
    ucPhotoBase64 = 5,
    ucDefaultLCID = 6,
    ucCurrentLocation = 7
};

struct __declspec(uuid("7693221d-7e9f-47d3-8782-6437cf3d4c0c"))
IPhone : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Endpoint (
        /*[out,retval]*/ struct IContactEndpoint * * _phoneEndpoint ) = 0;
      virtual HRESULT __stdcall get_Published (
        /*[out,retval]*/ VARIANT_BOOL * _isPublished ) = 0;
      virtual HRESULT __stdcall put_Published (
        /*[in]*/ VARIANT_BOOL _isPublished ) = 0;
};

enum ContactAvailability
{
    ucAvailabilityNone = 0,
    ucAvailabilityFree = 3500,
    ucAvailabilityFreeIdle = 5000,
    ucAvailabilityBusy = 6500,
    ucAvailabilityBusyIdle = 7500,
    ucAvailabilityDoNotDisturb = 9500,
    ucAvailabilityTemporarilyAway = 12500,
    ucAvailabilityAway = 15500,
    ucAvailabilityOffline = 18500
};

struct __declspec(uuid("d0984cef-e323-4574-b07b-5c970c9caee0"))
ICustomAvailabilityState : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Availability (
        /*[out,retval]*/ enum ContactAvailability * _availability ) = 0;
      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ long * _customPresenceActivityId ) = 0;
      virtual HRESULT __stdcall get_Activity (
        /*[out,retval]*/ BSTR * _value ) = 0;
};

struct __declspec(uuid("daa5fde3-a81c-4f23-80b6-c47b52c649bf"))
IContactEndpointCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall Add (
        /*[in]*/ struct IContactEndpoint * _endpoint ) = 0;
      virtual HRESULT __stdcall Clear ( ) = 0;
      virtual HRESULT __stdcall RemoveAt (
        /*[in]*/ long _index ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IContactEndpoint * * _endpoint ) = 0;
      virtual HRESULT __stdcall put_Item (
        /*[in]*/ long _index,
        /*[in]*/ struct IContactEndpoint * _endpoint ) = 0;
      virtual HRESULT __stdcall AddByUri (
        /*[in]*/ enum ContactEndpointType _endpointType,
        /*[in]*/ BSTR _endpointUri ) = 0;
      virtual HRESULT __stdcall IndexOf (
        /*[in]*/ struct IContactEndpoint * _contactEndpoint,
        /*[out,retval]*/ long * _index ) = 0;
};

struct __declspec(uuid("c0034194-f7b6-43eb-b0e0-7852fc8e7bfa"))
IPhonesChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Phones (
        /*[out,retval]*/ SAFEARRAY * * _phones ) = 0;
};

struct __declspec(uuid("dac6d77b-12f7-405c-bb33-002d6d4007f3"))
IResiliencyModeChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_IsInResiliencyMode (
        /*[out,retval]*/ VARIANT_BOOL * _isResiliencyMode ) = 0;
};

enum NotificationTypes
{
    ucNotificationTypeUnknown = 0,
    ucNotificationTypeAudio = 1,
    ucNotificationTypeVideo = 2,
    ucNotificationTypeInstantMessage = 4,
    ucNotificationTypeSharing = 8,
    ucNotificationTypeFileTransfer = 16,
    ucNotificationTypeAll = -1
};

enum NotificationUrgencyType
{
    ucNotificationUrgencyLow = 20,
    ucNotificationUrgencyNeutral = 50,
    ucNotificationUrgencyHigh = 80
};

enum AlertPrivacyType
{
    ucAlertPrivacyLow = 20,
    ucAlertPrivacyNeutral = 50,
    ucAlertPrivacyHigh = 80
};

enum AlertModeTypes
{
    ucAlertModeVisual = 1,
    ucAlertModeAudio = 2,
    ucAlertModeAny = -1
};

struct __declspec(uuid("fb67275c-4d9e-4c32-9793-a96749884fc1"))
IAlertLevel : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_NotificationEnabled (
        /*[out,retval]*/ VARIANT_BOOL * _value ) = 0;
      virtual HRESULT __stdcall get_LevelOfPrivacy (
        /*[out,retval]*/ enum AlertPrivacyType * _value ) = 0;
      virtual HRESULT __stdcall get_AlertModes (
        /*[out,retval]*/ enum AlertModeTypes * _value ) = 0;
};

struct __declspec(uuid("3c8d7146-35ea-4133-b2f6-c1fc2401a091"))
ISelf2 : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetAlertLevelForNotification (
        /*[in]*/ BSTR _senderIdentity,
        /*[in]*/ enum NotificationTypes _notificationTypes,
        /*[in]*/ enum NotificationUrgencyType _notificationUrgency,
        /*[out,retval]*/ struct IAlertLevel * * _alertLevel ) = 0;
};

struct __declspec(uuid("a9aab6a0-54b9-4419-aaaf-6b26dfac1585"))
_ISelfCallback : IDispatch
{};

struct __declspec(uuid("06437abb-c419-4b11-a474-1a2b02fbd646"))
_ISelfEvents : IDispatch
{};

struct __declspec(uuid("dc4bc923-6a52-4f02-9f8a-547b606955e7"))
_IAccessPermissionEvents : IDispatch
{};

struct __declspec(uuid("3c43a2be-ba4a-41e8-8f93-a9dc5f091892"))
_ISelf2Events : IDispatch
{};

struct __declspec(uuid("90b49a60-72af-4360-9f9b-916db7504a3c"))
IAlertLevelChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ChangedNotificationTypes (
        /*[out,retval]*/ enum NotificationTypes * _value ) = 0;
};

struct __declspec(uuid("cdcf9caf-ec72-4ec4-a7ca-05cbdda1959d"))
Self;
    // [ default ] interface ISelf
    // interface ISelf2
    // [ default, source ] dispinterface _ISelfEvents
    // [ source ] dispinterface _ISelf2Events

struct __declspec(uuid("b37552f7-36b6-4ef3-80c2-47f17957e683"))
AccessEntry;
    // [ default ] interface IAccessEntry

struct __declspec(uuid("260e70c7-05bd-4c59-9ea3-5ebb9efe5d93"))
Phone;
    // [ default ] interface IPhone

struct __declspec(uuid("19cc2b4b-2489-450c-9708-a6c28282aa76"))
AccessPermission;
    // [ default ] interface IAccessPermission
    // [ default, source ] dispinterface _IAccessPermissionEvents

struct __declspec(uuid("4cd03696-f8ba-4068-ac3c-a611f8a4bf3e"))
ContactEndpointCollection;
    // [ default ] interface IContactEndpointCollection

struct __declspec(uuid("e769d3b9-f1ac-42f9-9702-43aeb7b3d3c2"))
PhonesChangedEventData;
    // [ default ] interface IPhonesChangedEventData

struct __declspec(uuid("aca62ad7-05bf-4f21-81ef-f905170dcf23"))
AccessEntryCollectionChangedEventData;
    // [ default ] interface IAccessEntryCollectionChangedEventData

struct __declspec(uuid("9fad5936-57c6-4e92-b4a2-106011a3ae25"))
ResiliencyModeChangedEventData;
    // [ default ] interface IResiliencyModeChangedEventData

enum AudioPlayBackModes
{
    ucAudioFeedbackNone = 0,
    ucAudioFeedbackAlert = 1,
    ucAudioFeedbackCommunication = 2,
    ucAudioFeedbackHandset = 4,
    ucAudioFeedbackRinging = 8,
    ucAudioFeedbackAlertAndCommunication = 3,
    ucAudioFeedbackSuppressOnDeskphones = 0x80000000
};

enum DeviceType
{
    ucAudioDeviceUnknown = 0,
    ucAudioDeviceHumanInterfaceDevice = 1,
    ucAudioDeviceHeadSet = 2,
    ucAudioDeviceCustom = 3
};

struct __declspec(uuid("ce16d540-b9e8-4742-8659-9a7a318afb92"))
IDevice : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * _name ) = 0;
      virtual HRESULT __stdcall get_IsActive (
        /*[out,retval]*/ VARIANT_BOOL * _active ) = 0;
      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * _id ) = 0;
};

struct __declspec(uuid("1f8ebe6f-7993-42ce-980a-a2f03793be71"))
IDeviceCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IDevice * * _device ) = 0;
      virtual HRESULT __stdcall IndexOf (
        /*[in]*/ struct IDevice * _device,
        /*[out,retval]*/ long * _index ) = 0;
};

struct __declspec(uuid("b0b0e8f8-32da-4cb0-a124-d1e35891d53a"))
IAudioDevice : IDevice
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum DeviceType * _deviceType ) = 0;
      virtual HRESULT __stdcall get_Priority (
        /*[out,retval]*/ long * _priority ) = 0;
      virtual HRESULT __stdcall get_IsCertified (
        /*[out,retval]*/ VARIANT_BOOL * _certified ) = 0;
};

struct __declspec(uuid("da7cf784-67df-429c-82fb-9afd99f6cccb"))
IVideoDevice : IDevice
{};

struct __declspec(uuid("9274dbdc-43ce-45aa-a817-414a4494ad28"))
IDeviceManager : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall PlayAudioFile (
        /*[in]*/ BSTR _audioFileName,
        /*[in]*/ enum AudioPlayBackModes _playbackDevices,
        /*[in]*/ VARIANT_BOOL _loop,
        /*[in]*/ IUnknown * _deviceManagerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall StopPlayingAllAudioFiles ( ) = 0;
      virtual HRESULT __stdcall get_AudioDevices (
        /*[out,retval]*/ struct IDeviceCollection * * _devices ) = 0;
      virtual HRESULT __stdcall get_VideoDevices (
        /*[out,retval]*/ struct IDeviceCollection * * _devices ) = 0;
      virtual HRESULT __stdcall get_ActiveAudioDevice (
        /*[out,retval]*/ struct IAudioDevice * * _audioDevice ) = 0;
      virtual HRESULT __stdcall put_ActiveAudioDevice (
        /*[in]*/ struct IAudioDevice * _audioDevice ) = 0;
      virtual HRESULT __stdcall get_ActiveVideoDevice (
        /*[out,retval]*/ struct IVideoDevice * * _videoDevice ) = 0;
      virtual HRESULT __stdcall put_ActiveVideoDevice (
        /*[in]*/ struct IVideoDevice * _videoDevice ) = 0;
};

struct __declspec(uuid("08ddf5c4-fa12-4978-b26e-c6d23c453413"))
_IDeviceManagerCallback : IDispatch
{};

struct __declspec(uuid("fda610fc-4565-4e70-b088-4b72a16e1967"))
DeviceManager;
    // [ default ] interface IDeviceManager

struct __declspec(uuid("a0ba7d36-93cd-48c0-a168-3ef5933cee8a"))
DeviceCollection;
    // [ default ] interface IDeviceCollection

struct __declspec(uuid("81cb2ae5-6c07-4bc6-bc2b-f1c04545ce54"))
Device;
    // [ default ] interface IDevice

struct __declspec(uuid("ed1ce184-17aa-497f-abf2-c69b2144df2a"))
AudioDevice;
    // [ default ] interface IAudioDevice

struct __declspec(uuid("baccb218-6318-4ad9-b0a0-5b578c05f84e"))
VideoDevice;
    // [ default ] interface IVideoDevice

enum ConferenceInfoProperty
{
    ucConfInfoSubject = 393217,
    ucConfInfoExpiryTime = 1342570498,
    ucConfInfoOrganizerRoamingData = 393219,
    ucConfInfoNotificationData = 393220,
    ucConfInfoAttendees = 1074135045,
    ucConfInfoLeaders = 1074135046,
    ucConfInfoModalities = 537264135,
    ucConfInfoAdmissionKey = 393224,
    ucConfInfoPresentationModeEnabled = 268828681,
    ucConfInfoCollaborationModeEnabled = 268828682,
    ucConfInfoAutoAttendentEnabled = 268828683,
    ucConfInfoTelephonyAccessId = 393228,
    ucConfInfoAdmissionKeyRequired = 268828685,
    ucConfInfoInternalWebLocation = 393230,
    ucConfInfoExternalWebLocation = 393231,
    ucConfInfoIsStatic = 268828688,
    ucConfInfoAutopromote = 805699601,
    ucPSTNLobbyBypass = 268828690,
    ucACPTollNumber = 393235,
    ucACPTollFreeNumber = 393236,
    ucACPDomain = 393237,
    ucACPLeaderPasscode = 393238,
    ucACPParticipantPasscode = 393239,
    ucACPShowTollNumber = 268828696,
    ucACPShowTollFreeNumber = 268828697,
    ucACPEnableCallMe = 268828698,
    ucIsACPEnable = 268828699,
    ucConfInfoApplicationData = 1074135068,
    ucConfInfoVideoPresentationModeEnabled = 268828701,
    ucConfInfoGlobalAnnouncements = 268861441
};

enum ConferenceSetting
{
    ucConfFocusFactory = 458753,
    ucConfEnableAudioConferencing = 268894210,
    ucConfEnableAVConferencing = 268894211,
    ucConfEnableDataConferencing = 268894212,
    ucConfMeetingSize = 537329669,
    ucConfAllowAnonymousParticipants = 268894214,
    ucConfEnablePSTNConferencing = 268894215,
    ucConfACPsInfo = 458760,
    ucConfEnableEnterpriseCustomizedHelp = 268894217,
    ucConfCustomizedHelpUrl = 458762
};

enum ConferenceAdmissionType
{
    ucOpenAuthenticated = 1,
    ucClosedAuthenticated = 2,
    ucAnonymous = 3,
    ucLocked = 4
};

enum ConferenceAutoPromoteType
{
    ucAutoPromoteUnknown = 0,
    ucAutoPromoteNone = 1,
    ucAutoPromoteAuthenticated = 2,
    ucAutoPromoteAll = 4
};

struct __declspec(uuid("2d3d62ca-e9f6-4f94-8ebd-5fabab29e6a5"))
IConferenceInfoPropertyDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ConferenceInfoProperty _key,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall put_Item (
        /*[in]*/ enum ConferenceInfoProperty _key,
        /*[in]*/ VARIANT _value ) = 0;
      virtual HRESULT __stdcall Add (
        /*[in]*/ enum ConferenceInfoProperty _key,
        /*[in]*/ VARIANT _value ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ConferenceInfoProperty _key,
        /*[out]*/ VARIANT * _value,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ConferenceInfoProperty * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ConferenceInfoProperty _key,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall RemoveAt (
        /*[in]*/ long _index ) = 0;
      virtual HRESULT __stdcall Remove (
        /*[in]*/ enum ConferenceInfoProperty _key ) = 0;
      virtual HRESULT __stdcall RemoveAll ( ) = 0;
};

struct __declspec(uuid("b3c5872e-591c-4a74-9f84-b00ef970a838"))
IConferenceInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * _conferenceId ) = 0;
      virtual HRESULT __stdcall get_AlreadyScheduled (
        /*[out,retval]*/ VARIANT_BOOL * _value ) = 0;
      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * _uri ) = 0;
      virtual HRESULT __stdcall get_AdmissionType (
        /*[out,retval]*/ enum ConferenceAdmissionType * _admissionType ) = 0;
      virtual HRESULT __stdcall put_AdmissionType (
        /*[in]*/ enum ConferenceAdmissionType _admissionType ) = 0;
      virtual HRESULT __stdcall get_AdmissionKey (
        /*[out,retval]*/ BSTR * _admissionKey ) = 0;
      virtual HRESULT __stdcall put_AdmissionKey (
        /*[in]*/ BSTR _admissionKey ) = 0;
      virtual HRESULT __stdcall get_ConferenceInfoProperties (
        /*[out,retval]*/ struct IConferenceInfoPropertyDictionary * * _properties ) = 0;
};

struct __declspec(uuid("bd891697-c93e-4a0b-9b6c-004bd81b6ee8"))
IConferenceSettingDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ConferenceSetting _key,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall put_Item (
        /*[in]*/ enum ConferenceSetting _key,
        /*[in]*/ VARIANT _value ) = 0;
      virtual HRESULT __stdcall Add (
        /*[in]*/ enum ConferenceSetting _key,
        /*[in]*/ VARIANT _value ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ConferenceSetting _key,
        /*[out]*/ VARIANT * _value,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ConferenceSetting * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ConferenceSetting _key,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall RemoveAt (
        /*[in]*/ long _index ) = 0;
      virtual HRESULT __stdcall Remove (
        /*[in]*/ enum ConferenceSetting _key ) = 0;
      virtual HRESULT __stdcall RemoveAll ( ) = 0;
};

struct __declspec(uuid("c7acb102-b692-49cc-92da-5824822c7b96"))
IConferenceScheduler : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateConferenceInfo (
        /*[out,retval]*/ struct IConferenceInfo * * _confInfo ) = 0;
      virtual HRESULT __stdcall ScheduleConference (
        /*[in]*/ struct IConferenceInfo * _confInfo,
        /*[in]*/ IUnknown * _conferenceSchedulerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetConferenceInfo (
        /*[in]*/ BSTR _conferenceId,
        /*[in]*/ IUnknown * _conferenceSchedulerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetAllConferences (
        /*[in]*/ VARIANT_BOOL _isPublic,
        /*[in]*/ IUnknown * _conferenceSchedulerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall ModifyConference (
        /*[in]*/ struct IConferenceInfo * _conferenceInfo,
        /*[in]*/ IUnknown * _conferenceSchedulerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall DeleteConference (
        /*[in]*/ BSTR _confId,
        /*[in]*/ IUnknown * _conferenceSchedulerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetAvailableModalities (
        /*[in]*/ IUnknown * _conferenceSchedulerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetCapabilities (
        /*[in]*/ IUnknown * _conferenceSchedulerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetSettings (
        /*[out,retval]*/ struct IConferenceSettingDictionary * * _settings ) = 0;
};

struct __declspec(uuid("1ff8aa01-d8af-4b2b-b254-befefabaf2d6"))
IConferenceInfoCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IConferenceInfo * * _item ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
};

struct __declspec(uuid("9244d573-914f-4c1f-93f6-31609a95cbed"))
IConferenceRegionalAccessNumber : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_AccessNumber (
        /*[out,retval]*/ BSTR * _accessNumber ) = 0;
      virtual HRESULT __stdcall get_LanguageId (
        /*[out,retval]*/ unsigned long * _langId ) = 0;
};

struct __declspec(uuid("4cb70415-ba71-491e-b12d-155d85cbdff1"))
IConferenceRegionalAccessNumberCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IConferenceRegionalAccessNumber * * _accessNumber ) = 0;
};

struct __declspec(uuid("420a24e2-5c31-4262-9bd5-058682300ed6"))
IConferenceRegionalAccessInformation : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_RegionName (
        /*[out,retval]*/ BSTR * _regionName ) = 0;
      virtual HRESULT __stdcall get_DefaultAccessNumber (
        /*[out,retval]*/ BSTR * _accessNumber ) = 0;
      virtual HRESULT __stdcall get_DefaultLanguageId (
        /*[out,retval]*/ unsigned long * _accessNumber ) = 0;
      virtual HRESULT __stdcall get_AccessNumbers (
        /*[out,retval]*/ struct IConferenceRegionalAccessNumberCollection * * _ppAccessNumbers ) = 0;
};

struct __declspec(uuid("a0984f01-5d2c-4302-87a0-69be7b015143"))
IConferenceRegionalAccessInformationCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IConferenceRegionalAccessInformation * * _accessInformation ) = 0;
};

struct __declspec(uuid("66d1ffc0-c234-4f3d-abaa-98222a69e525"))
IConferenceTelephoneAccessInformation : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ExternalWebLink (
        /*[out,retval]*/ BSTR * _webLink ) = 0;
      virtual HRESULT __stdcall get_InternalWebLink (
        /*[out,retval]*/ BSTR * _webLink ) = 0;
      virtual HRESULT __stdcall get_RegionalAccessInformation (
        /*[out,retval]*/ struct IConferenceRegionalAccessInformationCollection * * _webLink ) = 0;
};

struct __declspec(uuid("e0e8b413-bb57-4b58-b1fb-288f3a269e35"))
IConferenceCapabilities : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Modalities (
        /*[out,retval]*/ unsigned long * _modalities ) = 0;
      virtual HRESULT __stdcall get_DefaultAdmissionType (
        /*[out,retval]*/ enum ConferenceAdmissionType * _type ) = 0;
      virtual HRESULT __stdcall get_IsPstnBridgingEnabled (
        /*[out,retval]*/ VARIANT_BOOL * _isEnabled ) = 0;
      virtual HRESULT __stdcall get_Version (
        /*[out,retval]*/ unsigned long * _version ) = 0;
      virtual HRESULT __stdcall get_IsConferenceKeyOptional (
        /*[out,retval]*/ VARIANT_BOOL * _isConferenceKeyOptional ) = 0;
      virtual HRESULT __stdcall get_IsAnonymousAllowed (
        /*[out,retval]*/ VARIANT_BOOL * _isAnonymousAllowed ) = 0;
      virtual HRESULT __stdcall get_PstnBridgingAccessNumbers (
        /*[out,retval]*/ BSTR * _PstnBridgingAccessNumbers ) = 0;
      virtual HRESULT __stdcall get_StaticMeetingLimit (
        /*[out,retval]*/ unsigned long * _StaticMeetingLimit ) = 0;
      virtual HRESULT __stdcall get_AutoPromoteAllow (
        /*[out,retval]*/ enum ConferenceAutoPromoteType * _AutoPromoteAllow ) = 0;
      virtual HRESULT __stdcall get_IsDefaultMeetingStatic (
        /*[out,retval]*/ VARIANT_BOOL * _isDefaultMeetingStatic ) = 0;
      virtual HRESULT __stdcall get_TelephoneAccessNumbers (
        /*[out,retval]*/ struct IConferenceTelephoneAccessInformation * * _telephoneAccessNumber ) = 0;
      virtual HRESULT __stdcall get_IsPSTNLobbyByPassAllowed (
        /*[out,retval]*/ VARIANT_BOOL * _isPSTNLobbyByPassAllowed ) = 0;
      virtual HRESULT __stdcall get_AreGlobalAnnouncementsAllowed (
        /*[out,retval]*/ VARIANT_BOOL * _areGlobalAnnouncementsAllowed ) = 0;
      virtual HRESULT __stdcall get_DefaultAutoPromote (
        /*[out,retval]*/ enum ConferenceAutoPromoteType * _DefaultAutoPromote ) = 0;
      virtual HRESULT __stdcall get_IsACPEnabled (
        /*[out,retval]*/ VARIANT_BOOL * _isACPEnabled ) = 0;
      virtual HRESULT __stdcall get_DefaultGlobalAnnouncements (
        /*[out,retval]*/ VARIANT_BOOL * _DefaultGlobalAnnouncements ) = 0;
};

struct __declspec(uuid("e5cff4a5-83a8-4fd8-8daa-77a3e2288d4a"))
IConferenceCapabilities2 : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_IsAudioPresentationModeAllowed (
        /*[out,retval]*/ VARIANT_BOOL * _isPresentationModeAllowed ) = 0;
      virtual HRESULT __stdcall get_CustomLogoUrl (
        /*[out,retval]*/ BSTR * _CustomLogoUrl ) = 0;
      virtual HRESULT __stdcall get_CustomHelpUrl (
        /*[out,retval]*/ BSTR * _CustomHelpUrl ) = 0;
      virtual HRESULT __stdcall get_CustomLegalUrl (
        /*[out,retval]*/ BSTR * _CustomLegalUrl ) = 0;
      virtual HRESULT __stdcall get_CustomFooterText (
        /*[out,retval]*/ BSTR * _CustomFooterText ) = 0;
      virtual HRESULT __stdcall get_IsMainVideoPresentationModeAllowed (
        /*[out,retval]*/ VARIANT_BOOL * _isPresentationModeAllowed ) = 0;
      virtual HRESULT __stdcall get_IsPanoVideoPresentationModeAllowed (
        /*[out,retval]*/ VARIANT_BOOL * _isPresentationModeAllowed ) = 0;
};

struct __declspec(uuid("c22ce6bf-3a47-46b7-8544-44305badfef9"))
_IConferenceSchedulerCallback : IDispatch
{};

struct __declspec(uuid("57dd2d13-f18d-4a4e-a022-2f8c269a4afd"))
ConferenceScheduler;
    // [ default ] interface IConferenceScheduler

struct __declspec(uuid("2ae6db87-2dfd-427d-b8e3-b4d738e2684e"))
ConferenceInfo;
    // [ default ] interface IConferenceInfo

struct __declspec(uuid("40368b76-a6ad-44c8-a088-cdea0214cb10"))
ConferenceInfoPropertyDictionary;
    // [ default ] interface IConferenceInfoPropertyDictionary

struct __declspec(uuid("960390d4-c70d-43a1-b272-22e73bc4d4b6"))
ConferenceApplicationData;
    // [ default ] interface IConferenceApplicationData

struct __declspec(uuid("64b7e9ad-d920-4eaa-95c7-9d95de9ace24"))
IConferenceApplicationProperties : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall GetItem (
        /*[in]*/ long index,
        /*[out]*/ BSTR * _key,
        /*[out]*/ BSTR * _value ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ BSTR _key,
        /*[out,retval]*/ BSTR * _value ) = 0;
      virtual HRESULT __stdcall Add (
        /*[in]*/ BSTR _key,
        /*[in]*/ BSTR _value ) = 0;
      virtual HRESULT __stdcall Remove (
        /*[in]*/ BSTR _key ) = 0;
      virtual HRESULT __stdcall RemoveAll ( ) = 0;
};

struct __declspec(uuid("be6086a1-d436-4834-89ee-3ca4f8596a58"))
IConferenceApplicationData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ApplicationProperties (
        /*[in]*/ BSTR _applicationId,
        /*[out,retval]*/ struct IConferenceApplicationProperties * * _properties ) = 0;
};

struct __declspec(uuid("541fe2a9-95be-4098-960a-64464b251405"))
ConferenceApplicationProperties;
    // [ default ] interface IConferenceApplicationProperties

struct __declspec(uuid("7f10fab7-2a38-444c-8f43-31ddf0fe79ae"))
ConferenceSettingDictionary;
    // [ default ] interface IConferenceSettingDictionary

struct __declspec(uuid("488f9ce1-5468-48ec-b4db-fc2f7f1ab247"))
ConferenceCapabilities;
    // [ default ] interface IConferenceCapabilities
    // interface IConferenceCapabilities2

struct __declspec(uuid("8b8cbacf-0aa1-419e-94d2-60f844a6fb9e"))
ConferenceInfoCollection;
    // [ default ] interface IConferenceInfoCollection

struct __declspec(uuid("bf527f65-83f6-4fa4-ab1b-17a8585f0d8a"))
ConferenceTelephoneAccessInformation;
    // [ default ] interface IConferenceTelephoneAccessInformation

struct __declspec(uuid("36815016-484e-41f5-9e54-b952f4b44400"))
ConferenceRegionalAccessInformation;
    // [ default ] interface IConferenceRegionalAccessInformation

struct __declspec(uuid("f080136d-d09a-42de-b9e4-2e926d1a430f"))
ConferenceRegionalAccessInformationCollection;
    // [ default ] interface IConferenceRegionalAccessInformationCollection

struct __declspec(uuid("ba9b1f14-ce79-417e-9320-704c9e391116"))
ConferenceRegionalAccessNumberCollection;
    // [ default ] interface IConferenceRegionalAccessNumberCollection

struct __declspec(uuid("b82a9c56-654d-4a24-b308-aba49f65183a"))
ConferenceRegionalAccessNumber;
    // [ default ] interface IConferenceRegionalAccessNumber

enum UtilitiesAction
{
    ucUtilitiesActionShowPropertiesDialog = 0,
    ucUtilitiesActionAddToExternalContacts = 1,
    ucUtilitiesActionFindPreviousConverstations = 2,
    ucUtilitiesActionComposeEmail = 3,
    ucUtilitiesActionScheduleMeeting = 4
};

struct __declspec(uuid("c4d95336-5949-4d13-aac8-82039e5cbcfc"))
_IUtilitiesCallback : IDispatch
{};

struct __declspec(uuid("26f15083-86db-4663-89ed-a0abc896b9a1"))
IScheduledConferenceManager : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetScheduledConferenceItems (
        /*[in]*/ IUnknown * _scheduledConferenceManagerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("9112ce9b-3704-48c0-86dc-7a7b0f34d7a3"))
IUtilities2 : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ExchangeECPUrl (
        /*[out,retval]*/ BSTR * _url ) = 0;
      virtual HRESULT __stdcall get_ScheduledConferenceManager (
        /*[out,retval]*/ struct IScheduledConferenceManager * * scheduleConferenceManager ) = 0;
};

struct __declspec(uuid("3417d640-3976-4cb6-bc82-54e5e2e10784"))
_IUtilities2Events : IDispatch
{};

struct __declspec(uuid("7818a3f7-674b-477e-99e8-54774579a8e7"))
IExchangeECPUrlChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_NewUrl (
        /*[out,retval]*/ BSTR * _url ) = 0;
};

//struct __declspec(uuid("b5b1ff8f-3f61-4ea5-a339-d94fdfaf15e1"))
//Utilities;
//    // [ default ] interface IUtilities
//    // interface IUtilities2
//    // [ default, source ] dispinterface _IUtilities2Events

struct __declspec(uuid("71a054b3-4d4b-4643-b14f-4572494b19e7"))
IScheduledConference : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ConferenceId (
        /*[out,retval]*/ BSTR * _entryId ) = 0;
      virtual HRESULT __stdcall get_Organizer (
        /*[out,retval]*/ BSTR * _sender ) = 0;
      virtual HRESULT __stdcall get_OrganizerUri (
        /*[out,retval]*/ BSTR * _senderUri ) = 0;
      virtual HRESULT __stdcall get_StartTime (
        /*[out,retval]*/ DATE * _received ) = 0;
      virtual HRESULT __stdcall get_EndTime (
        /*[out,retval]*/ DATE * _received ) = 0;
      virtual HRESULT __stdcall get_Location (
        /*[out,retval]*/ BSTR * Location ) = 0;
      virtual HRESULT __stdcall get_Subject (
        /*[out,retval]*/ BSTR * _subject ) = 0;
      virtual HRESULT __stdcall get_ConferenceURL (
        /*[out,retval]*/ BSTR * _conferenceUrl ) = 0;
      virtual HRESULT __stdcall get_ConferenceExternalURL (
        /*[out,retval]*/ BSTR * _conferenceExternalURL ) = 0;
};

struct __declspec(uuid("f423726d-0e9b-4b55-9569-e79865210f69"))
IScheduledConferenceCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IScheduledConference * * _scheduledConference ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
};

struct __declspec(uuid("219dc126-9a48-483f-80c2-3f22b3b47829"))
_IScheduledConferenceManagerCallback : IDispatch
{};

struct __declspec(uuid("ead92b6d-a7db-4ab8-a494-c5a2771fa987"))
Utilities2;
    // [ default ] interface IUtilities2
    // [ default, source ] dispinterface _IUtilities2Events

struct __declspec(uuid("d40faf18-96f8-407e-ba32-48daba521e18"))
ExchangeECPUrlChangedEventData;
    // [ default ] interface IExchangeECPUrlChangedEventData

struct __declspec(uuid("bb99483d-2df9-424f-96a8-6718cd7fbb9f"))
ScheduledConference;
    // [ default ] interface IScheduledConference

struct __declspec(uuid("584dc0a4-f664-4a1d-b781-662c082d6b78"))
ScheduledConferenceCollection;
    // [ default ] interface IScheduledConferenceCollection

struct __declspec(uuid("75927b32-8f46-40eb-a6c9-7598f0db38c0"))
ScheduledConferenceManager;
    // [ default ] interface IScheduledConferenceManager

enum RoomMessageType
{
    ucRoomMessageAlert = 0,
    ucRoomMessageRegular = 1,
    ucRoomMessageStory = 2,
    ucRoomMessageTypeMAX = 3
};

enum RoomMessageFormat
{
    ucRoomMessagePlainText = 0,
    ucRoomMessageRtf = 1,
    ucRoomMessageMAX = 2
};

enum RoomMessageFilteringAction
{
    ucRoomMessageFilterCanceled = 0,
    ucRoomMessageFilterReplaced = 1,
    ucRoomMessageFilterPassed = 2
};

enum RoomSearchModeType
{
    ucRoomSearchModeRegular = 0,
    ucRoomSearchModeExtended = 1
};

enum RoomType
{
    ucRoomUnset = 0,
    ucRoomNormal = 1,
    ucRoomAuditorium = 2,
    ucRoomAddhoc = 3
};

enum RoomProperty
{
    ucRoomUri = 983040,
    ucRoomDomainUrl = 983041,
    ucRoomDescription = 983042,
    ucRoomTitle = 983043,
    ucRoomType = 806289412
};

enum RoomJoinState
{
    ucNotJoined = 0,
    ucPending = 1,
    ucSuccess = 2,
    ucFailed = 3,
    ucFailedUnauthorized = 4
};

enum RoomManagerState
{
    ucRoomManagerDisabled = 0,
    ucRoomManagerEnabled = 1
};

struct __declspec(uuid("603e2454-725a-4679-a6a8-363b21633cb2"))
IRoomMessageDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum RoomMessageFormat _key,
        /*[out,retval]*/ VARIANT * _itemValue ) = 0;
      virtual HRESULT __stdcall put_Item (
        /*[in]*/ enum RoomMessageFormat _key,
        /*[in]*/ VARIANT _itemValue ) = 0;
      virtual HRESULT __stdcall Add (
        /*[in]*/ enum RoomMessageFormat _key,
        /*[in]*/ VARIANT _value ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum RoomMessageFormat _key,
        /*[out]*/ VARIANT * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum RoomMessageFormat * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum RoomMessageFormat _key,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall RemoveAt (
        /*[in]*/ long _index ) = 0;
      virtual HRESULT __stdcall Remove (
        /*[in]*/ enum RoomMessageFormat _key ) = 0;
      virtual HRESULT __stdcall RemoveAll ( ) = 0;
      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum RoomMessageType * _type ) = 0;
      virtual HRESULT __stdcall put_Type (
        /*[in]*/ enum RoomMessageType _type ) = 0;
      virtual HRESULT __stdcall get_StoryTitle (
        /*[out,retval]*/ BSTR * _storyTitle ) = 0;
      virtual HRESULT __stdcall put_StoryTitle (
        /*[in]*/ BSTR _storyTitle ) = 0;
};

struct __declspec(uuid("a0c3e9ba-c622-45a2-9fed-2c662f3f489a"))
IRoomMessage : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ unsigned long * _id ) = 0;
      virtual HRESULT __stdcall get_MessageDictionary (
        /*[out,retval]*/ struct IRoomMessageDictionary * * _messageDictionary ) = 0;
      virtual HRESULT __stdcall get_SentTime (
        /*[out,retval]*/ DATE * _sentTime ) = 0;
      virtual HRESULT __stdcall get_SenderUri (
        /*[out,retval]*/ BSTR * _senderUri ) = 0;
      virtual HRESULT __stdcall get_RoomUri (
        /*[out,retval]*/ BSTR * _roomUri ) = 0;
};

struct __declspec(uuid("21c41ede-7ed8-4d97-9b0d-dd785781af9b"))
IRoomMessageCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IRoomMessage * * _roomMessages ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall Clear ( ) = 0;
      virtual HRESULT __stdcall IndexOf (
        /*[in]*/ struct IRoomMessage * _roomMesssage,
        /*[out,retval]*/ long * _index ) = 0;
};

struct __declspec(uuid("68da8699-6c6e-4390-bd82-709ff1dc1702"))
IRoomUser : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * _uri ) = 0;
      virtual HRESULT __stdcall get_CanSendMessage (
        /*[out,retval]*/ VARIANT_BOOL * _canSendMessage ) = 0;
};

struct __declspec(uuid("7bf20b14-58d1-494b-b301-9b16bacc9610"))
IRoomPropertyDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum RoomProperty _propertyType,
        /*[out,retval]*/ VARIANT * _itemValue ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum RoomProperty _propertyType,
        /*[out]*/ VARIANT * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum RoomProperty * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum RoomProperty _propertyType,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("a766520d-f700-4bf5-8428-56bd914c1c40"))
IRoom : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Participants (
        /*[out,retval]*/ SAFEARRAY * * _participants ) = 0;
      virtual HRESULT __stdcall get_JoinedState (
        /*[out,retval]*/ enum RoomJoinState * _state ) = 0;
      virtual HRESULT __stdcall get_UnreadRoomMessageCount (
        /*[out,retval]*/ unsigned long * _count ) = 0;
      virtual HRESULT __stdcall get_IsOutgoingMessageFilterEnabled (
        /*[out,retval]*/ VARIANT_BOOL * isEnabled ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IRoomPropertyDictionary * * _properties ) = 0;
      virtual HRESULT __stdcall CreateRoomMessageDictionary (
        /*[out,retval]*/ struct IRoomMessageDictionary * * _messageDictionary ) = 0;
      virtual HRESULT __stdcall EnableOutgoingMessageFilter (
        /*[in]*/ VARIANT_BOOL _noWarning ) = 0;
      virtual HRESULT __stdcall DisableOutgoingMessageFilter ( ) = 0;
      virtual HRESULT __stdcall SendMessage (
        /*[in]*/ struct IRoomMessageDictionary * _messageInfo,
        /*[in]*/ IUnknown * _roomCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall SendFilteredMessage (
        /*[in]*/ struct IRoomMessage * _message,
        /*[in]*/ enum RoomMessageFilteringAction _action ) = 0;
      virtual HRESULT __stdcall RetrieveAdditionalMessages (
        /*[in]*/ unsigned long _lastMessageID,
        /*[in]*/ unsigned long _count,
        /*[in]*/ IUnknown * _roomCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Join (
        /*[in]*/ IUnknown * _roomCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Leave (
        /*[in]*/ IUnknown * _roomCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("05789509-afa1-4e90-b4fa-1a7820db9796"))
IRoomCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IRoom * * _room ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall IndexOf (
        /*[in]*/ struct IRoom * _room,
        /*[out,retval]*/ long * _index ) = 0;
};

struct __declspec(uuid("76de1cd8-f278-48a1-a391-e5beb8f428e6"))
IRoomManager : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ enum RoomManagerState * _state ) = 0;
      virtual HRESULT __stdcall get_FollowedRooms (
        /*[out,retval]*/ SAFEARRAY * * _followedRooms ) = 0;
      virtual HRESULT __stdcall QueryRooms (
        /*[in]*/ BSTR _keyword,
        /*[in]*/ enum RoomSearchModeType _searchType,
        /*[in]*/ IUnknown * _roomManagerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall QueryRoomByUri (
        /*[in]*/ BSTR _roomUri,
        /*[in]*/ IUnknown * _roomManagerCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall TryGetRoomByUri (
        /*[in]*/ BSTR _roomUri,
        /*[out]*/ struct IRoom * * _room,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("b03614ae-8d68-4386-9e1c-939cabcf1232"))
IRoomPropertyChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Property (
        /*[out,retval]*/ enum RoomProperty * _propertyType ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ VARIANT * _propertyValue ) = 0;
};

struct __declspec(uuid("b8fd4a53-e7e6-4995-a5b5-1306c7584964"))
IUnreadMessageCountChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_NewUnreadMessageCount (
        /*[out,retval]*/ unsigned long * _count ) = 0;
};

struct __declspec(uuid("1566fa3c-fd24-41db-a2b8-6232f89783de"))
IRoomMessageEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Message (
        /*[out,retval]*/ struct IRoomMessage * * _roomMessage ) = 0;
};

struct __declspec(uuid("081793da-7221-45c3-8afa-3a2e6e464670"))
IRoomMessagesEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Messages (
        /*[out,retval]*/ struct IRoomMessageCollection * * _roomMessages ) = 0;
};

struct __declspec(uuid("bc9f5f40-d315-40c1-b1c6-cea0646e18e2"))
IRoomUserChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_CanSendMessage (
        /*[out,retval]*/ VARIANT_BOOL * _canSendMessage ) = 0;
};

struct __declspec(uuid("7452bd0f-65cb-4a5e-ac37-e2bea1f43dd9"))
IRoomParticipantsEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Participants (
        /*[out,retval]*/ SAFEARRAY * * _roomUsers ) = 0;
};

struct __declspec(uuid("4d120020-ce64-43c5-9f84-7a7b2360388f"))
IRoomJoinStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_JoinState (
        /*[out,retval]*/ enum RoomJoinState * _state ) = 0;
};

struct __declspec(uuid("300e56a3-ce08-4eb3-9f48-505aa162c9da"))
_IRoomEvents : IDispatch
{};

struct __declspec(uuid("e2c6d6e0-52ef-47a2-bda9-9d6a7a29a391"))
_IRoomManagerEvents : IDispatch
{};

struct __declspec(uuid("8f04e92d-ee26-4b64-b9c1-56dd90816810"))
IFollowedRoomsChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Room (
        /*[out,retval]*/ struct IRoom * * _room ) = 0;
};

struct __declspec(uuid("89c7eb8e-42bc-4c9f-bb34-88cde83ccb37"))
IRoomManagerStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum RoomManagerState * _state ) = 0;
};

struct __declspec(uuid("8e73e7e1-90ce-4d3b-b5b7-71e87b2e8ba5"))
_IRoomUserEvents : IDispatch
{};

struct __declspec(uuid("628d7bdd-45ab-404d-8505-b3adc15f454d"))
_IRoomCallback : IDispatch
{};

struct __declspec(uuid("0fdb5b86-dd44-41bf-a443-317fd19bbf3d"))
_IRoomManagerCallback : IDispatch
{};

struct __declspec(uuid("44e7ad71-d7d5-45b7-9faa-6c346c6da105"))
RoomMessage;
    // [ default ] interface IRoomMessage

struct __declspec(uuid("ad75a6fb-14e3-4830-adf1-c7067d6550a7"))
RoomUser;
    // [ default ] interface IRoomUser

struct __declspec(uuid("4fac92c8-e226-44b6-a530-aed7aeb4cef4"))
Room;
    // [ default ] interface IRoom
    // [ default, source ] dispinterface _IRoomEvents

struct __declspec(uuid("0909f1e4-5f02-4e09-857d-bd951e956fc1"))
RoomManager;
    // [ default ] interface IRoomManager
    // [ default, source ] dispinterface _IRoomManagerEvents

struct __declspec(uuid("3068d48b-3c56-4b15-aca8-97a9c8b43c13"))
RoomMessageCollection;
    // [ default ] interface IRoomMessageCollection

struct __declspec(uuid("3df72257-ea41-45a7-bd01-dcf2104b250c"))
RoomCollection;
    // [ default ] interface IRoomCollection

struct __declspec(uuid("0a52c7b8-6123-413b-ad6a-05faff69c519"))
RoomPropertyChangedEventData;
    // [ default ] interface IRoomPropertyChangedEventData

struct __declspec(uuid("96e4b115-4c68-4038-a0db-3777c4ae4aeb"))
UnreadMessageCountChangedEventData;
    // [ default ] interface IUnreadMessageCountChangedEventData

struct __declspec(uuid("8d69edcf-b7c8-4bd1-bfd2-08c8c5b1e030"))
RoomMessageEventData;
    // [ default ] interface IRoomMessageEventData

struct __declspec(uuid("09ec6931-ac8d-4b45-88b1-ded520f562dc"))
RoomMessagesEventData;
    // [ default ] interface IRoomMessagesEventData

struct __declspec(uuid("fcb650ff-e3a0-47bc-acd6-c1c3c8473ae4"))
FollowedRoomsChangedEventData;
    // [ default ] interface IFollowedRoomsChangedEventData

struct __declspec(uuid("e269b211-d1f7-45fc-9202-437f71e9f0b3"))
RoomManagerStateChangedEventData;
    // [ default ] interface IRoomManagerStateChangedEventData

struct __declspec(uuid("31b10e61-fccb-42d3-873d-936ca52160e8"))
RoomUserChangedEventData;
    // [ default ] interface IRoomUserChangedEventData

struct __declspec(uuid("3e207f26-c1de-43f6-bd2b-8a46615325bf"))
RoomParticipantsEventData;
    // [ default ] interface IRoomParticipantsEventData

struct __declspec(uuid("cdfc37e9-bd76-420f-881f-4a6aab4e51bf"))
RoomJoinStateChangedEventData;
    // [ default ] interface IRoomJoinStateChangedEventData

struct __declspec(uuid("b4f20beb-831c-4bb8-9065-3964e4330e79"))
RoomPropertyDictionary;
    // [ default ] interface IRoomPropertyDictionary

struct __declspec(uuid("6c4db7f3-3a3b-410c-944c-027289e79804"))
RoomMessageDictionary;
    // [ default ] interface IRoomMessageDictionary

enum AutomationModalities
{
    uiaConversationModeInstantMessage = 1,
    uiaConversationModeAudio = 2,
    uiaConversationModeVideo = 4,
    uiaConversationModeFileTransfer = 8,
    uiaConversationModeApplicationSharing = 16,
    uiaConversationModeNone = 0
};

enum AutomationModalitySettings
{
    uiaConversationContextSubject = 1,
    uiaConversationContextPreviousConversation = 2,
    uiaConversationContextOutlookEntryId = 3,
    uiaConversationContextFilePathToTransfer = 4,
    uiaConversationContextFileIsShared = 5,
    uiaConversationContextFirstInstantMessage = 6,
    uiaConversationContextSendFirstInstantMessageImmediately = 7,
    uiaConversationContextFileHistoryLink = 8,
    uiaConversationContextStartConferenceByCallingMeAt = 9,
    uiaConversationContextSharedProcess = 10,
    uiaConversationContextSharedWindow = 11,
    uiaConversationContextSharedDesktop = 12,
    uiaConversationContextSharedMonitor = 13,
    uiaConversationContextDataObjectForFileTransfer = 14,
    uiaConversationContextApplicationId = 15,
    uiaConversationContextApplicationData = 16,
    uiaConversationContextHyperLink = 17,
    uiaConversationContextParentWindow = 18,
    uiaConversationContextMeetNow = 19,
    uiaConversationContextRemoteEndpointUri = 20,
    uiaConversationContextMax = 21
};

enum ConversationWindowState
{
    uiaConversationWindowNotInitialized = 0,
    uiaConversationWindowInitialized = 1,
    uiaConversationWindowDestroyed = 2
};

enum ConversationWindowInformationType
{
    uiaConversationWindowWidthMin = 537788416,
    uiaConversationWindowHeightMin = 537788417,
    uiaConversationWindowIsDocked = 269352962,
    uiaConversationWindowHasVideo = 269352963,
    uiaConversationWindowHasContentStage = 269352964,
    uiaConversationWindowHasExtensionPane = 269352965,
    uiaConversationWindowIsFullScreen = 269352966
};

enum ConversationWindowAction
{
    ucConversationWindowActionAddOfficePowerPoint = 0,
    ucConversationWindowActionAddOfficeOneNote = 1,
    ucConversationWindowFullScreen = 2,
    ucConversationWindowActionMAX = 3
};

struct __declspec(uuid("ca62e6f3-3e6d-451b-aeaa-6a1b7aacf8d4"))
IConversationWindowNeedsSizeChangeEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_RecommendedWindowWidth (
        /*[out,retval]*/ long * _width ) = 0;
      virtual HRESULT __stdcall get_RecommendedWindowHeight (
        /*[out,retval]*/ long * _height ) = 0;
      virtual HRESULT __stdcall get_MinimumWindowWidth (
        /*[out,retval]*/ long * _width ) = 0;
      virtual HRESULT __stdcall get_MinimumWindowHeight (
        /*[out,retval]*/ long * _height ) = 0;
      virtual HRESULT __stdcall get_MaximumWindowWidth (
        /*[out,retval]*/ long * _width ) = 0;
      virtual HRESULT __stdcall get_MaximumWindowHeight (
        /*[out,retval]*/ long * _height ) = 0;
};

struct __declspec(uuid("f97bed54-e434-4020-a197-f15aea9d9c95"))
IConversationWindowNeedsAttentionEventData : IDispatch
{};

struct __declspec(uuid("806d3227-4cb8-47c4-9864-7d4df4f44069"))
_IAutomationCallback : IDispatch
{};

struct __declspec(uuid("9e0f8fca-d9c9-47cd-87f4-5554217dffcc"))
_IConversationWindowEvents : IDispatch
{};

struct __declspec(uuid("76428270-36fa-4236-8bdf-aadb39fd1371"))
_IConversationWindowCallback : IDispatch
{};

struct __declspec(uuid("109d32ce-e09f-46f5-9ecb-01e1d273b7da"))
IConversationWindowInformationDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ConversationWindowInformationType _key,
        /*[out,retval]*/ VARIANT * _itemValue ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ConversationWindowInformationType _key,
        /*[out]*/ VARIANT * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ConversationWindowInformationType * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ConversationWindowInformationType _key,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("663c61f4-a9da-4a6f-a606-f142f665df61"))
IConversationWindowActions : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall AddOfficePowerPointToConversation (
        /*[in]*/ BSTR _PowerPointFilePath,
        /*[in]*/ BSTR _OriginalPowerPointUrl ) = 0;
      virtual HRESULT __stdcall AddOfficeOneNoteToConversation (
        /*[in]*/ BSTR _pageId ) = 0;
      virtual HRESULT __stdcall CanInvoke (
        /*[in]*/ enum ConversationWindowAction _action,
        /*[out,retval]*/ VARIANT_BOOL * _enabled ) = 0;
};

enum AnnotationTool
{
    ucAnnotationToolTelepointer = 0,
    ucAnnotationToolDrawLine = 1,
    ucAnnotationToolDrawArrowLine = 2,
    ucAnnotationToolDrawDoubleArrowLine = 3,
    ucAnnotationToolDrawRectangle = 4,
    ucAnnotationToolDrawEllipse = 5,
    ucAnnotationToolDrawPen = 6,
    ucAnnotationToolDrawHighlighter = 7,
    ucAnnotationToolPutCheckStamp = 8,
    ucAnnotationToolPutArrowStamp = 9,
    ucAnnotationToolPutXStamp = 10,
    ucAnnotationToolMAX = 11
};

enum VideoViewInformationType
{
    uiaVideoViewIsDocked = 269549568,
    uiaVideoViewIsFullScreen = 269549569,
    uiaVideoViewIsHasPano = 269549570,
    uiaVideoViewHasGallery = 269549571,
    uiaVideoViewHasSelfVideo = 269549572,
    uiaVideoViewIsExtended = 269549573,
    uiaVideoViewIsActiveSpeakerMode = 269549574,
    uiaVideoViewMultipleVideoEnabled = 269549575
};

struct __declspec(uuid("b00f2520-029c-47d0-b4e8-8fbef47caa7e"))
IVideoViewInformationDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum VideoViewInformationType _key,
        /*[out,retval]*/ VARIANT * _itemValue ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum VideoViewInformationType _key,
        /*[out]*/ VARIANT * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum VideoViewInformationType * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum VideoViewInformationType _key,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

enum VideoViewMode
{
    ucVideoViewModeGallery = 0,
    uiaVideoViewModeActiveSpeakerOnly = 1
};

struct __declspec(uuid("55a21cf3-a2b3-484e-be2a-14280f501289"))
IVideoView : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IVideoViewInformationDictionary * * _properties ) = 0;
      virtual HRESULT __stdcall SetVideoViewMode (
        /*[in]*/ enum VideoViewMode _mode ) = 0;
      virtual HRESULT __stdcall EnterFullScreen (
        /*[in]*/ SAFEARRAY * _monitorIds ) = 0;
      virtual HRESULT __stdcall LeaveFullScreen ( ) = 0;
      virtual HRESULT __stdcall DockVideo ( ) = 0;
      virtual HRESULT __stdcall UndockVideo ( ) = 0;
      virtual HRESULT __stdcall ShowPano ( ) = 0;
      virtual HRESULT __stdcall HidePano ( ) = 0;
      virtual HRESULT __stdcall ShowGallery ( ) = 0;
      virtual HRESULT __stdcall HideGallery ( ) = 0;
      virtual HRESULT __stdcall ShowSelfVideo ( ) = 0;
      virtual HRESULT __stdcall HideSelfVideo ( ) = 0;
};

struct __declspec(uuid("02dd541f-c863-4295-92ab-a4cf29f5d841"))
IConversationWindow2 : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ enum ConversationWindowState * _state ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IConversationWindowInformationDictionary * * _properties ) = 0;
      virtual HRESULT __stdcall GetConversationWindowActions (
        /*[out,retval]*/ struct IConversationWindowActions * * _convWindowActions ) = 0;
      virtual HRESULT __stdcall SetPreferredAnnotationTool (
        /*[in]*/ enum AnnotationTool _annotationTool,
        /*[in]*/ unsigned long _annotationColor ) = 0;
      virtual HRESULT __stdcall ShowFullScreen (
        /*[in]*/ int _monitorId ) = 0;
      virtual HRESULT __stdcall ExitFullScreen ( ) = 0;
      virtual HRESULT __stdcall ShowContent ( ) = 0;
      virtual HRESULT __stdcall HideContent ( ) = 0;
      virtual HRESULT __stdcall get_VideoView (
        /*[out,retval]*/ struct IVideoView * * _videoView ) = 0;
      virtual HRESULT __stdcall CanInvoke (
        /*[in]*/ enum ConversationWindowAction _action,
        /*[out,retval]*/ VARIANT_BOOL * _enabled ) = 0;
};

struct __declspec(uuid("af2aba54-f591-42ba-8004-20944c9fb4d7"))
_IConversationWindow2Events : IDispatch
{};

struct __declspec(uuid("2209e091-5b90-4caf-8582-cc5979883376"))
IConversationWindowStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum ConversationWindowState * _oldState ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum ConversationWindowState * _newState ) = 0;
};

struct __declspec(uuid("ff037274-455a-4e34-b5d1-d42db866f9b7"))
IConversationWindowInformationChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ChangedProperties (
        /*[out,retval]*/ struct IConversationWindowInformationDictionary * * _changedProperties ) = 0;
};

struct __declspec(uuid("0678c83e-f580-4d99-902f-930699b28be8"))
IConversationWindowActionAvailabilityChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Action (
        /*[out,retval]*/ enum ConversationWindowAction * _action ) = 0;
      virtual HRESULT __stdcall get_IsAvailable (
        /*[out,retval]*/ VARIANT_BOOL * _isAvailable ) = 0;
};

struct __declspec(uuid("440f28e7-13b4-4e6a-8483-65ad69b83070"))
ConversationWindow;
    // [ default ] interface IConversationWindow
    // interface IConversationWindow2
    // [ default, source ] dispinterface _IConversationWindowEvents
    // [ source ] dispinterface _IConversationWindow2Events

struct __declspec(uuid("e45fb8a0-6716-4f0f-a4f0-220c3deb7337"))
Automation;
    // [ default ] interface IAutomation

struct __declspec(uuid("3dc07b52-702c-47fe-8893-b96e2a798163"))
ConversationWindowNeedsSizeChangeEventData;
    // [ default ] interface IConversationWindowNeedsSizeChangeEventData

struct __declspec(uuid("ac4f1849-52c8-4dcf-a21f-93206b320af0"))
ConversationWindowNeedsAttentionEventData;
    // [ default ] interface IConversationWindowNeedsAttentionEventData

enum ShareableContentAction
{
    ucShareableContentActionRemove = 0,
    ucShareableContentActionPresent = 1,
    ucShareableContentActionStopPresenting = 2,
    ucShareableContentActionSyncWithPresenter = 3,
    ucShareableContentActionDownloadFile = 4,
    ucShareableContentActionTakeOverAsPresenter = 5,
    ucShareableContentActionSaveAnnotation = 6,
    ucShareableContentActionClearAllAnnotations = 7,
    ucShareableContentActionUpload = 8
};

enum ShareableContentProperty
{
    ucShareableContentId = 1048576,
    ucShareableContentType = 806354945,
    ucShareableContentTitle = 1048578,
    ucShareableContentOwner = 1074790403,
    ucShareableContentCreationTime = 1343225860,
    ucShareableContentLastUsedTime = 1343225861,
    ucShareableContentPresenter = 1074790420,
    ucShareableContentIsInSyncWithPresenter = 269484062,
    ucShareableContentOriginalFileUrl = 1048607,
    ucShareableContentPresentationViewingURL = 1048608,
    ucShareableContentVisibility = 806354977
};

enum ShareableContentState
{
    ucShareableContentStateInitializing = 0,
    ucShareableContentStateOffline = 1,
    ucShareableContentStateConnecting = 2,
    ucShareableContentStateOnline = 3,
    ucShareableContentStateDisconnecting = 4,
    ucShareableContentStateActive = 5,
    ucShareableContentStateUnusable = 6
};

enum ShareableContentType
{
    ucShareableContentTypeUnsupported = 0,
    ucShareableContentTypePowerPoint = 1,
    ucShareableContentTypeWhiteboard = 2,
    ucShareableContentTypeNativeFile = 3,
    ucShareableContentTypePoll = 4,
    ucShareableContentTypeSharedNotes = 5
};

struct __declspec(uuid("ca478b44-4fce-4082-8700-8cec69fb50e1"))
IShareableContentPropertyDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ShareableContentProperty _propertyType,
        /*[out,retval]*/ VARIANT * _itemValue ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ShareableContentProperty _propertyType,
        /*[out]*/ VARIANT * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ShareableContentProperty * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ VARIANT * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ShareableContentProperty _propertyType,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

enum ContentSavingFileType
{
    ucContentSavingFileTypeXPS = 806354944,
    ucContentSavingFileTypePNG = 806354945
};

struct __declspec(uuid("ae3a7c39-9c30-4e0f-bae4-d8344ef377ea"))
_IContentSharingModalityEvents : IDispatch
{};

struct __declspec(uuid("6b551c9e-de81-41df-a0ae-39f0af11d508"))
_IShareableContentEvents : IDispatch
{};

struct __declspec(uuid("94093301-8854-408e-8138-c7c3d2daca1f"))
IShareableContentStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum ShareableContentState * _oldState ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum ShareableContentState * _newState ) = 0;
};

struct __declspec(uuid("3c8210c8-8578-47c6-87a9-fa1ad2ba9873"))
IShareableContentPropertyChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Property (
        /*[out,retval]*/ enum ShareableContentProperty * _propertyType ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ VARIANT * _propertyValue ) = 0;
};

struct __declspec(uuid("4023052e-e3b7-48a3-bcb3-ca2234d9baee"))
IShareableContentActionAvailabilityChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Action (
        /*[out,retval]*/ enum ShareableContentAction * _action ) = 0;
      virtual HRESULT __stdcall get_IsAvailable (
        /*[out,retval]*/ VARIANT_BOOL * _isAvailable ) = 0;
};

struct __declspec(uuid("6964b116-8fe6-4c19-9052-0eaa74460a58"))
_IConversation2Callback : IDispatch
{};

struct __declspec(uuid("84340964-7820-4ebc-bcd3-702926de23e8"))
_IContentSharingModalityCallback : IDispatch
{};

struct __declspec(uuid("3187885c-a8af-4d4b-8e80-f4b9d447a7b2"))
_IShareableContentCallback : IDispatch
{};

struct __declspec(uuid("36cdc166-4f21-46ad-a60e-8551f26c1d41"))
_IParticipant2Callback : IDispatch
{};

struct __declspec(uuid("70986dce-4aa8-4dfb-9a45-4bfcf7a206f5"))
_IChannel2Callback : IDispatch
{};

struct __declspec(uuid("46b5e644-a89e-4c63-86e5-34666807ee75"))
Conversation2;
    // [ default ] interface IConversation2

struct __declspec(uuid("6fb3b002-987c-460c-9e9f-b45f2ed3b1e9"))
Participant2;
    // [ default ] interface IParticipant2

struct __declspec(uuid("a351a6d6-9b47-41a3-a6d3-07ddf2203e7d"))
Channel2;
    // [ default ] interface IChannel2

struct __declspec(uuid("ed094732-ddca-4a31-8095-e9774e3dbc0d"))
ConferenceAccessInformation2;
    // [ default ] interface IConferenceAccessInformation2

struct __declspec(uuid("5e7ab04c-e4b8-49fb-85ff-9e2bca2899d3"))
IConferenceAccessInformation2 : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_IsScheduled (
        /*[out,retval]*/ VARIANT_BOOL * _isScheduled ) = 0;
      virtual HRESULT __stdcall get_OrganizerDisplayName (
        /*[out,retval]*/ BSTR * _organizerDisplayName ) = 0;
      virtual HRESULT __stdcall get_OrganizerUri (
        /*[out,retval]*/ BSTR * _organizerUri ) = 0;
      virtual HRESULT __stdcall get_IsCustomized (
        /*[out,retval]*/ VARIANT_BOOL * _isCustomized ) = 0;
};

struct __declspec(uuid("509879fa-3ee5-4328-8625-a0dfb9211b63"))
ContentSharingModality;
    // [ default ] interface IContentSharingModality
    // [ default, source ] dispinterface _IContentSharingModalityEvents

struct __declspec(uuid("235b711c-6fa3-4f91-b50e-36ed53b3801f"))
ShareableContentPropertyDictionary;
    // [ default ] interface IShareableContentPropertyDictionary

struct __declspec(uuid("dff47f2b-ad75-44dc-81aa-e23ca7c6a66f"))
ShareableContent;
    // [ default ] interface IShareableContent
    // [ default, source ] dispinterface _IShareableContentEvents

struct __declspec(uuid("2d1a91b0-a593-4327-a7b2-f7dd01574bf3"))
PowerPointContent;
    // [ default ] interface IPowerPointContent
    // [ default, source ] dispinterface _IShareableContentEvents

struct __declspec(uuid("90f87439-acd6-4802-a740-59051166e57d"))
ShareableContentCollection;
    // [ default ] interface IShareableContentCollection

struct __declspec(uuid("34c459ac-5125-42d4-bf92-b14988f0289e"))
ContentCollectionChangedEventData;
    // [ default ] interface IContentCollectionChangedEventData

struct __declspec(uuid("411768f9-a6e5-4857-8a10-7923c9fcbb1e"))
ActiveContentChangedEventData;
    // [ default ] interface IActiveContentChangedEventData

struct __declspec(uuid("12063e4a-c16f-4e0f-8886-61ddf52f4090"))
ActivePresenterChangedEventData;
    // [ default ] interface IActivePresenterChangedEventData

struct __declspec(uuid("71a27f3d-167b-42de-a8e6-5cd8658bc5d8"))
ShareableContentPropertyChangedEventData;
    // [ default ] interface IShareableContentPropertyChangedEventData

struct __declspec(uuid("3ba1b3e8-f66f-4eec-a84e-8256ddad606e"))
ShareableContentStateChangedEventData;
    // [ default ] interface IShareableContentStateChangedEventData

struct __declspec(uuid("a7b1e1de-cf73-492f-95c2-dfdc02d4ee2b"))
ShareableContentActionAvailabilityChangedEventData;
    // [ default ] interface IShareableContentActionAvailabilityChangedEventData

struct __declspec(uuid("8b9f0b59-7c29-46e2-b6a6-804043a346db"))
Self2;
    // [ default ] interface ISelf2
    // [ default, source ] dispinterface _ISelf2Events

struct __declspec(uuid("3bc642ef-64be-45ae-982d-babdaabb355a"))
AlertLevel;
    // [ default ] interface IAlertLevel

struct __declspec(uuid("17374941-9877-429f-a88c-8ecd0bfc49d4"))
AlertLevelChangedEventData;
    // [ default ] interface IAlertLevelChangedEventData

struct __declspec(uuid("086f0e9d-a416-40f9-877c-f1c82dc1a6ad"))
IVideoViewInformationChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum VideoViewInformationType * PropertyType ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ VARIANT * Value ) = 0;
};

struct __declspec(uuid("30f974fb-1f7b-49d3-8da4-01694ba474d2"))
_IVideoViewEvents : IDispatch
{};

struct __declspec(uuid("b58839e6-eb13-402d-bb2c-83ebae14e6aa"))
ConversationWindow2;
    // [ default ] interface IConversationWindow2
    // [ default, source ] dispinterface _IConversationWindow2Events

struct __declspec(uuid("360c1678-114e-42c0-a8ff-ea8ba10ca2d8"))
ConversationWindowInformationDictionary;
    // [ default ] interface IConversationWindowInformationDictionary

struct __declspec(uuid("0e32c65c-c4d9-4280-9142-ec102e19a05b"))
ConversationWindowStateChangedEventData;
    // [ default ] interface IConversationWindowStateChangedEventData

struct __declspec(uuid("f10b6480-558f-49f3-a515-ab6d05fb949f"))
ConversationWindowInformationChangedEventData;
    // [ default ] interface IConversationWindowInformationChangedEventData

struct __declspec(uuid("7c4db3cb-796e-4993-8d49-414317d014da"))
ConversationWindowActionAvailabilityChangedEventData;
    // [ default ] interface IConversationWindowActionAvailabilityChangedEventData

struct __declspec(uuid("85147550-4a7b-4b36-8fb5-144b5087a299"))
ConversationWindowActions;
    // [ default ] interface IConversationWindowActions

struct __declspec(uuid("24efe0bc-5661-4f96-adc4-351dbe47a4cc"))
VideoView;
    // [ default ] interface IVideoView
    // [ default, source ] dispinterface _IVideoViewEvents

struct __declspec(uuid("a13f4ab3-c0fc-466a-9457-2df0ad8c74ed"))
VideoViewInformationChangedEventData;
    // [ default ] interface IVideoViewInformationChangedEventData

struct __declspec(uuid("4c7d404c-e8a6-416c-8f5f-b5b95f634b47"))
VideoViewInformationDictionary;
    // [ default ] interface IVideoViewInformationDictionary

struct __declspec(uuid("fd28f0a8-d83a-40fb-8c30-6908a6117d5b"))
Contact2;
    // [ default ] interface IContact2

struct __declspec(uuid("24290b3e-7dc8-485b-85a1-956bcd8d2c5c"))
SearchResults2;
    // [ default ] interface ISearchResults2

enum ClientSettingsType
{
    ucClientSettingsEnableP2PVideo = 269615105,
    ucClientSettingsPublicCloudProviderInfo = 1611792386,
    ucClientSettingsMax = 3
};

struct __declspec(uuid("1b04d929-11e4-46c2-9162-ea4f2758c3a7"))
IClientSettings : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ClientSettingsType _settingType,
        /*[out,retval]*/ VARIANT * _settingValue ) = 0;
      virtual HRESULT __stdcall put_Item (
        /*[in]*/ enum ClientSettingsType _settingType,
        /*[in]*/ VARIANT _settingValue ) = 0;
      virtual HRESULT __stdcall CanSetValue (
        /*[in]*/ enum ClientSettingsType _settingType,
        /*[out,retval]*/ VARIANT_BOOL * _canSetValue ) = 0;
};

struct __declspec(uuid("25d64aea-0e65-49cb-8d6d-65db0ac1af65"))
IClientSettingsChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum ClientSettingsType * _settingType ) = 0;
      virtual HRESULT __stdcall get_Value (
        /*[out,retval]*/ VARIANT * _settingValue ) = 0;
};

struct __declspec(uuid("74b6caf0-755e-454d-a819-d0e439ef5296"))
IPublicCloudProviderInfo : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * _name ) = 0;
      virtual HRESULT __stdcall get_IconUrl (
        /*[out,retval]*/ BSTR * _iconUrl ) = 0;
      virtual HRESULT __stdcall get_ProxyFqdn (
        /*[out,retval]*/ BSTR * _proxyFqdn ) = 0;
      virtual HRESULT __stdcall get_NameDecorationDomain (
        /*[out,retval]*/ BSTR * _nameDecorationDomain ) = 0;
};

struct __declspec(uuid("569e04f2-d286-4903-9642-44d850310956"))
_IClientSettingsEvents : IDispatch
{};

struct __declspec(uuid("66d123ba-11e2-4c2c-8056-5aca97245cd2"))
Client2;
    // [ default ] interface IClient2

struct __declspec(uuid("8eb30566-8e31-4dd9-b1e7-477756160076"))
ClientSettings;
    // [ default ] interface IClientSettings
    // [ default, source ] dispinterface _IClientSettingsEvents

struct __declspec(uuid("f19e9036-6db9-4d48-9f46-26995f9d5163"))
ClientSettingsChangedEventData;
    // [ default ] interface IClientSettingsChangedEventData

struct __declspec(uuid("8e99de73-6a8b-47a9-bae6-6b03e53622e8"))
PublicCloudProviderInfo;
    // [ default ] interface IPublicCloudProviderInfo

struct __declspec(uuid("e8451131-7352-4906-9e57-5cb164fee062"))
ConferenceCapabilities2;
    // [ default ] interface IConferenceCapabilities2

enum ClientState
{
    ucClientStateUninitialized = 0,
    ucClientStateSignedOut = 1,
    ucClientStateSigningIn = 2,
    ucClientStateSignedIn = 3,
    ucClientStateSigningOut = 4,
    ucClientStateShuttingDown = 5,
    ucClientStateInitializing = 6
};

enum CredentialRequestedType
{
    ucCredentialRequestSignIn = 0,
    ucCredentialRequestGlobalAddressListImport = 1,
    ucCredentialRequestDGExpansion = 2,
    ucCredentialRequestCalendarData = 3,
    ucCredentialRequestAutoUpdate = 4,
    ucCredentialRequestWebSearch = 5,
    ucCredentialRequestPhotoProvider = 6,
    ucCredentialRequestSharePointManagement = 7,
    ucCredentialRequestCertificateProvider = 8,
    ucCredentialRequestLocationProvider = 9,
    ucCredentialRequestResponseGroupServer = 10,
    ucCredentialRequestAnonymousSignIn = 11,
    ucCredentialRequestBitsDownloadManager = 12,
    ucCredentialRequestSipAuthBroker = 13,
    ucCredentialRequestLyncAutodiscover = 14,
    ucCredentialRequestUnknown = 15,
    ucCredentialRequestQueryConferenceInfo = 16
};

enum LyncClientNotification
{
    ucClientNotificationMapiStatus = 0,
    ucClientNotificationAutodiscovery = 1,
    ucClientNotificationCalendarQueryStatus = 2,
    ucClientNotificationOutlookContactProviderStatus = 3,
    ucClientNotificationWabContactProviderStatus = 4,
    ucClientNotificationFolderManagerStatus = 5
};

enum MapiStatus
{
    ucMapiStatusOk = 0,
    ucMapiStatusInstallationMisconfigured = 1,
    ucMapiStatusProfileMisconfigured = 2,
    ucMapiStatusEmailAddressesDoNotMatch = 3,
    ucMapiStatusAuthenticationFailedOrServerDown = 4,
    ucMapiStatusOutlookContactProviderAuthenticationFailedOrServerDown = 5,
    ucMapiStatusOutlookContactProviderUnknownError = 6,
    ucMapiStatusAutodiscoveryPending = 7,
    ucMapiStatusAutodiscoveryFailedTemporarily = 8,
    ucMapiStatusAutodiscoveryFailedPermanently = 9,
    ucMapiStatusWebServiceAuthenticationFailedOrServerDown = 10,
    ucMapiStatusCalendarUnknownError = 11,
    ucMapiStatusCredentialDialogCancel = 12,
    ucMapiStatusOffline = 13,
    ucMapiStatusNotReady = 14,
    ucMapiStatusMisconfiguration = 15,
    ucMapiStatusWabContactProviderUnknownError = 16,
    ucMapiStatusSpoolerMalfunctioning = 17,
    ucMapiStatusSpoolerDiscardedOldFiles = 18,
    ucMapiStatusOtherFailure = 19,
    ucMapiStatusMapiFolderCriticalError = 20,
    ucMapiStatusEwsFolderCriticalError = 21,
    ucMapiStatusEwsFolderConnectionError = 22,
    ucMapiStatusEwsFolderSecurityError = 23,
    ucMapiStatusFolderManagerUnknownError = 24
};

enum LyncClientConfigurationMode
{
    ucConfigurationModeAuto = 0,
    ucConfigurationModeManual = 1
};

enum TransportMode
{
    ucTcpTransport = 0,
    ucTlsTransport = 1
};

enum SignInConfigurationType
{
    ucCanSetMode = 0,
    ucCanSetTransport = 1,
    ucCanSetInternalServer = 2,
    ucCanSetExternalServer = 3,
    ucCanSetSavePassword = 4
};

enum SupportedFeatures
{
    ucNone = 0,
    ucDelegatorContexts = 1,
    ucInstantMessageGif = 2,
    ucInstantMessageInk = 4,
    ucInstantMessageRtf = 8,
    ucInstantMessageHtml = 16,
    ucPhonePairing = 32,
    ucApplicationSharing = 64,
    ucLegacyApplicationSharing = 128,
    ucApplicationInvite = 256,
    ucAudio = 512,
    ucFileTransfer = 1024,
    ucVideo = 2048,
    ucInstantMessage = 4096,
    ucBranchOfficeResiliency = 8192,
    ucRegistrarClustering = 16384,
    ucDhcpServerDiscovery = 32768,
    ucPinBasedBootstrapping = 65536,
    ucTrustModelPrompt = 131072,
    ucDataConferencing = 262144,
    ucRecording = 524288,
    ucCanShowLobby = 1048576,
    ucProcessLogonEvents = 2097152,
    ucFreeBusyIntegration = 4194304,
    ucAllFeatures = -1
};

enum ClientType
{
    ucClientTypeUC = 1,
    ucClientTypeDelegator = 2,
    ucClientTypeResponseGroup = 3,
    ucClientTypeAnonymous = 4
};

enum LyncClientCapabilityTypes
{
    ucCapabilityUserActivity = 1,
    ucCapabilityText = 2,
    ucCapabilityGifInk = 4,
    ucCapabilityIsfInk = 8,
    ucCapabilityAudio = 16,
    ucCapabilityVideo = 32,
    ucCapabilityCccp = 64,
    ucCapabilityCalendar = 128,
    ucCapabilityRemoteCallControl = 256,
    ucCapabilityBreakthrough = 512,
    ucCapabilityApplicationSharing = 1024
};

enum OIFeature
{
    oiFeatureFastSearch = 1,
    oiFeaturePictures = 2,
    oiFeatureNonBuddyPresence = 4,
    oiFeatureContactTagging = 8,
    oiFeatureSinglePartyFileTransfer = 16,
    oiFeatureMultipartyFileTransfer = 32,
    oiFeatureQuickContacts = 64,
    oiFeatureResolveContact = 128,
    oiFeatureAddPowerPointToConversation = 256,
    oiFeatureAddOneNoteToConversation = 512
};

enum OIInterface
{
    oiInterfaceILyncClient = 1,
    oiInterfaceIAutomation = 2,
    oiInterfaceIInteropHelper = 3,
    oiInterfaceILyncClientRaw = 256,
    oiInterfaceIAutomationRaw = 257
};

enum OIHelperAction
{
    oiActionStartup = 0,
    oiActionShutdown = 1
};

struct __declspec(uuid("6a222195-f65e-467f-8f77-eb180bd85288"))
IUCOfficeIntegration : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetSupportedFeatures (
        /*[in]*/ BSTR _version,
        /*[out,retval]*/ enum OIFeature * _features ) = 0;
      virtual HRESULT __stdcall GetAuthenticationInfo (
        /*[in]*/ BSTR _version,
        /*[out,retval]*/ BSTR * _info ) = 0;
      virtual HRESULT __stdcall GetInterface (
        /*[in]*/ BSTR _version,
        /*[in]*/ enum OIInterface _interface,
        /*[out,retval]*/ IDispatch * * _value ) = 0;
};

struct __declspec(uuid("f8df0b4b-a4ce-476b-81d3-9fba98461b62"))
ISignInConfiguration : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CanSet (
        /*[in]*/ enum SignInConfigurationType _configurationType,
        /*[out,retval]*/ VARIANT_BOOL * _value ) = 0;
      virtual HRESULT __stdcall get_Mode (
        /*[out,retval]*/ enum LyncClientConfigurationMode * _value ) = 0;
      virtual HRESULT __stdcall put_Mode (
        /*[in]*/ enum LyncClientConfigurationMode _value ) = 0;
      virtual HRESULT __stdcall get_TransportMode (
        /*[out,retval]*/ enum TransportMode * _value ) = 0;
      virtual HRESULT __stdcall put_TransportMode (
        /*[in]*/ enum TransportMode _value ) = 0;
      virtual HRESULT __stdcall get_InternalServerUrl (
        /*[out,retval]*/ BSTR * _server ) = 0;
      virtual HRESULT __stdcall put_InternalServerUrl (
        /*[in]*/ BSTR _server ) = 0;
      virtual HRESULT __stdcall get_ExternalServerUrl (
        /*[out,retval]*/ BSTR * _server ) = 0;
      virtual HRESULT __stdcall put_ExternalServerUrl (
        /*[in]*/ BSTR _server ) = 0;
      virtual HRESULT __stdcall get_UserName (
        /*[out,retval]*/ BSTR * _userName ) = 0;
      virtual HRESULT __stdcall get_IsPasswordSaved (
        /*[out,retval]*/ VARIANT_BOOL * _savePassword ) = 0;
      virtual HRESULT __stdcall put_IsPasswordSaved (
        /*[in]*/ VARIANT_BOOL _savePassword ) = 0;
      virtual HRESULT __stdcall get_SignInAutoRetry (
        /*[out,retval]*/ VARIANT_BOOL * _autoRetry ) = 0;
      virtual HRESULT __stdcall put_SignInAutoRetry (
        /*[in]*/ VARIANT_BOOL _autoRetry ) = 0;
      virtual HRESULT __stdcall get_SignInAsAvailability (
        /*[out,retval]*/ enum ContactAvailability * _availability ) = 0;
      virtual HRESULT __stdcall put_SignInAsAvailability (
        /*[in]*/ enum ContactAvailability _availability ) = 0;
      virtual HRESULT __stdcall get_SignedInFromIntranet (
        /*[out,retval]*/ VARIANT_BOOL * _value ) = 0;
      virtual HRESULT __stdcall put_DisplayName (
        /*[in]*/ BSTR _arg1 ) = 0;
};

struct __declspec(uuid("aa41ee75-f9c5-42f7-8d17-1d92d6ba4d92"))
IApplicationRegistration : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall AddRegistration ( ) = 0;
      virtual HRESULT __stdcall RemoveRegistration ( ) = 0;
      virtual HRESULT __stdcall SetExtensibilityWindowProperties (
        /*[in]*/ BSTR _intrnalUrl,
        /*[in]*/ BSTR _externalUrl,
        /*[in]*/ enum ConversationWindowExtensionSize _windowSize ) = 0;
};

struct __declspec(uuid("330639f2-e399-48cb-863f-56a50a27a138"))
IWebAppExtension : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Parameters (
        /*[out,retval]*/ BSTR * _value ) = 0;
};

struct __declspec(uuid("9b3a702d-5be1-4fce-adb3-ebbd23e078c3"))
IClientStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum ClientState * _newState ) = 0;
      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum ClientState * _oldState ) = 0;
      virtual HRESULT __stdcall get_StatusCode (
        /*[out,retval]*/ long * _HResultStatusCode ) = 0;
};

struct __declspec(uuid("9ab20314-b258-48f8-b659-ad250dfa20f5"))
ILyncClientNotificationReceivedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum LyncClientNotification * _notification ) = 0;
      virtual HRESULT __stdcall get_MapiStatus (
        /*[out,retval]*/ enum MapiStatus * _status ) = 0;
};

enum CredentialRequestResult
{
    ucCredsResultUnknown = 0,
    ucCredsResultSuccess = 1,
    ucCredsResultFailure = 2
};

struct __declspec(uuid("f0b791fa-434b-49f2-b84e-05619ebe6beb"))
ICredentialRequestedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum CredentialRequestedType * _type ) = 0;
      virtual HRESULT __stdcall get_UserName (
        /*[out,retval]*/ BSTR * _userName ) = 0;
      virtual HRESULT __stdcall get_CanSavePassword (
        /*[out,retval]*/ VARIANT_BOOL * _value ) = 0;
      virtual HRESULT __stdcall get_IsPasswordSaved (
        /*[out,retval]*/ VARIANT_BOOL * _value ) = 0;
      virtual HRESULT __stdcall Submit (
        /*[in]*/ BSTR _domainAndUsername,
        /*[in]*/ BSTR _password,
        /*[in]*/ VARIANT_BOOL _savePassword ) = 0;
      virtual HRESULT __stdcall Cancel (
        /*[in]*/ VARIANT_BOOL _cancelForWholeSession ) = 0;
      virtual HRESULT __stdcall get_CredentialRequestResult (
        /*[out,retval]*/ enum CredentialRequestResult * _credsResult ) = 0;
};

struct __declspec(uuid("e664682c-cba5-4c7c-a65b-9cd046f24cbd"))
ISignInDelayedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_EstimatedStartDelay (
        /*[out,retval]*/ unsigned long * _startDelay ) = 0;
      virtual HRESULT __stdcall get_StatusCode (
        /*[out,retval]*/ long * _HResultStatusCode ) = 0;
};

struct __declspec(uuid("b28f2007-3926-49cf-9101-4ccb9574049b"))
IPreferredCapabilitiesChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_PreferredCapabilities (
        /*[out,retval]*/ enum LyncClientCapabilityTypes * _capabilities ) = 0;
};

struct __declspec(uuid("b26f6403-761e-48d4-84e2-619c17eaeb13"))
_ILyncClientCallback : IDispatch
{};

struct __declspec(uuid("64561371-77a2-4bd4-a909-e6deb1b5c6cf"))
_IClientEvents : IDispatch
{};

struct __declspec(uuid("953ae732-f53a-4116-ac1b-0321b3fb3dba"))
_ILyncClientEvents : IDispatch
{};

struct __declspec(uuid("92a9eb9b-85ee-444e-a302-00c3c945baa4"))
_IDelegatorClientEvents : IDispatch
{};

struct __declspec(uuid("8c9b6953-33cc-4a04-9da4-f71ae79da0c0"))
_IUCOfficeIntegrationEvents : IDispatch
{};

struct __declspec(uuid("3f927576-996e-4a91-9cf9-d9fa46d2aaf5"))
Client;
    // [ default ] interface IClient
    // interface IClient2

struct __declspec(uuid("a9f3c4a7-6de0-48e4-8184-ba7d6d985908"))
LyncClient;
    // [ default ] interface ILyncClient
    // interface IClient2
    // [ default, source ] dispinterface _ILyncClientEvents

struct __declspec(uuid("6a69dc0b-f5ab-47e7-8500-00ed31486c01"))
SignInConfiguration;
    // [ default ] interface ISignInConfiguration

struct __declspec(uuid("b7a53e76-8630-4a93-9771-940012a1371b"))
ApplicationRegistration;
    // [ default ] interface IApplicationRegistration

struct __declspec(uuid("859d1502-64e6-4977-a9b3-d131d68c723b"))
CustomAvailabilityState;
    // [ default ] interface ICustomAvailabilityState

struct __declspec(uuid("90099fb2-b732-4cb7-81de-c35bca5e15dc"))
DelegatorClient;
    // [ default ] interface IDelegatorClient
    // interface IClient2
    // [ default, source ] dispinterface _IDelegatorClientEvents

struct __declspec(uuid("36522520-67a7-48ae-aa63-b5a7bea21b7c"))
ClientStateChangedEventData;
    // [ default ] interface IClientStateChangedEventData

struct __declspec(uuid("861bfb12-94ce-4c1d-9b3a-33b20c9d48ac"))
LyncClientNotificationReceivedEventData;
    // [ default ] interface ILyncClientNotificationReceivedEventData

struct __declspec(uuid("cbd9b3de-2ed2-47cc-ae82-54daeb26ebb7"))
CredentialRequestedEventData;
    // [ default ] interface ICredentialRequestedEventData

struct __declspec(uuid("4883d56b-83f3-4cab-8161-0b1554113037"))
SignInDelayedEventData;
    // [ default ] interface ISignInDelayedEventData

struct __declspec(uuid("cec8bfbe-3e0e-4732-b3fe-5babf6170f11"))
PreferredCapabilitiesChangedEventData;
    // [ default ] interface IPreferredCapabilitiesChangedEventData

struct __declspec(uuid("37e5d2a9-55bd-42f8-8b46-3c02496853b0"))
DelegatorClientCollectionEventData;
    // [ default ] interface IDelegatorClientCollectionEventData

struct __declspec(uuid("a0651028-ba7a-4d71-877f-12e0175a5806"))
UCOfficeIntegration;
    // [ default ] interface IUCOfficeIntegration
    // [ default, source ] dispinterface _IUCOfficeIntegrationEvents

struct __declspec(uuid("ec9773ad-7d22-4b03-b949-5e9055af0f19"))
IGroupCollectionChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Group (
        /*[out,retval]*/ struct IGroup * * _group ) = 0;
};

struct __declspec(uuid("57c3cb4d-c794-456f-846d-0e1eca78bd00"))
IContactCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall TryGetContact (
        /*[in]*/ BSTR _uri,
        /*[out]*/ struct IContact * * _value,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("f97973d7-9436-4a3e-a802-f6ea172a0425"))
IGroup : IContactCollection
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum GroupType * _groupType ) = 0;
      virtual HRESULT __stdcall get_Name (
        /*[out,retval]*/ BSTR * _groupName ) = 0;
      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * _groupId ) = 0;
      virtual HRESULT __stdcall AddContact (
        /*[in]*/ struct IContact * _contact,
        /*[in]*/ IUnknown * _groupCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall RemoveContact (
        /*[in]*/ struct IContact * pContact,
        /*[in]*/ IUnknown * _groupCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall get_ContactManager (
        /*[out,retval]*/ struct IContactManager * * _contactsAndGroupsManager ) = 0;
      virtual HRESULT __stdcall CanInvoke (
        /*[in]*/ enum GroupAction _action,
        /*[in]*/ struct IContact * _contact,
        /*[out,retval]*/ VARIANT_BOOL * _canInvoke ) = 0;
};

struct __declspec(uuid("0c81ec90-9469-4ed4-acad-adb86e11039c"))
IGroupCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IGroup * * _group ) = 0;
      virtual HRESULT __stdcall TryGetGroup (
        /*[in]*/ BSTR _groupName,
        /*[out]*/ struct IGroup * * _value,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetGroupsByType (
        /*[in]*/ enum GroupType _groupType,
        /*[out,retval]*/ struct IGroupCollection * * _groups ) = 0;
      virtual HRESULT __stdcall IndexOf (
        /*[in]*/ struct IGroup * _group,
        /*[out,retval]*/ long * _index ) = 0;
};

struct __declspec(uuid("340c54a6-35cf-4971-b540-72d7f040ac24"))
IDistributionGroup : IGroup
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_EmailAddress (
        /*[out,retval]*/ BSTR * _emailAddress ) = 0;
      virtual HRESULT __stdcall get_NestedGroups (
        /*[out,retval]*/ struct IGroupCollection * * _nestedGroups ) = 0;
      virtual HRESULT __stdcall Expand (
        /*[in]*/ IUnknown * _groupCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall get_IsExpanded (
        /*[out,retval]*/ VARIANT_BOOL * _isExpanded ) = 0;
      virtual HRESULT __stdcall GetAllMembers (
        /*[in]*/ IUnknown * _groupCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall get_Description (
        /*[out,retval]*/ BSTR * _description ) = 0;
      virtual HRESULT __stdcall GetOwner (
        /*[in]*/ IUnknown * _groupCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("d804ba85-db0f-49ee-b723-dcdf0425bd9f"))
ICustomGroup : IGroup
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall Rename (
        /*[in]*/ BSTR _newName,
        /*[in]*/ IUnknown * _groupCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("35ce3a44-c636-4d31-bc95-b9370c2251b9"))
IFavoriteContacts : IGroup
{};

struct __declspec(uuid("85109845-6ae2-4b6b-9524-65ba203e5b7d"))
IContact : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * _contactUri ) = 0;
      virtual HRESULT __stdcall GetContactInformation (
        /*[in]*/ enum ContactInformationType _contactInformationType,
        /*[out,retval]*/ VARIANT * _contactInformationValue ) = 0;
      virtual HRESULT __stdcall BatchGetContactInformation (
        /*[in]*/ SAFEARRAY * _contactInformationTypes,
        /*[out,retval]*/ struct IContactInformationDictionary * * _contactInformation ) = 0;
      virtual HRESULT __stdcall get_Settings (
        /*[out,retval]*/ struct IContactSettingDictionary * * _contactSettings ) = 0;
      virtual HRESULT __stdcall CanStart (
        /*[in]*/ enum ModalityTypes _modalityTypes,
        /*[out,retval]*/ VARIANT_BOOL * _canStart ) = 0;
      virtual HRESULT __stdcall ChangeSetting (
        /*[in]*/ enum ContactSetting _contactSettingType,
        /*[in]*/ VARIANT _contactSettingValue,
        /*[in]*/ IUnknown * _contactCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CanChangeSetting (
        /*[in]*/ enum ContactSetting _contactSetting,
        /*[out,retval]*/ VARIANT_BOOL * _canSet ) = 0;
      virtual HRESULT __stdcall get_ContactManager (
        /*[out,retval]*/ struct IContactManager * * _contactsAndGroupsManager ) = 0;
      virtual HRESULT __stdcall get_CustomGroups (
        /*[out,retval]*/ struct IGroupCollection * * _groups ) = 0;
      virtual HRESULT __stdcall MoveToGroup (
        /*[in]*/ struct IGroup * _targetGroup,
        /*[in]*/ struct IGroup * _sourceGroup,
        /*[in]*/ IUnknown * _contactCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CanMoveToGroup (
        /*[in]*/ struct IGroup * _targetGroup,
        /*[in]*/ struct IGroup * _sourceGroup,
        /*[out,retval]*/ VARIANT_BOOL * _canMoveToGroup ) = 0;
      virtual HRESULT __stdcall CreateContactEndpoint (
        /*[in]*/ BSTR _telephoneUri,
        /*[out,retval]*/ struct IContactEndpoint * * _collaborationEndpoint ) = 0;
      virtual HRESULT __stdcall get_UnifiedCommunicationType (
        /*[out,retval]*/ enum UnifiedCommunicationType * _unifiedCommunicationType ) = 0;
      virtual HRESULT __stdcall GetOrganizationInformation (
        /*[in]*/ enum OrganizationStructureTypes _orgInfoTypes,
        /*[in]*/ IUnknown * _contactCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetMultipleContactInformation (
        /*[in]*/ SAFEARRAY * _contactInformationTypes,
        /*[out,retval]*/ struct IContactInformationDictionary * * _contactInformation ) = 0;
};

struct __declspec(uuid("95b8d3e8-f3d5-4dc0-bcfe-ab80c835dea0"))
IContactSubscription : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall Subscribe (
        /*[in]*/ enum ContactSubscriptionRefreshRate _subscriptionFreshness,
        /*[in]*/ SAFEARRAY * _contactInformationTypes ) = 0;
      virtual HRESULT __stdcall Unsubscribe ( ) = 0;
      virtual HRESULT __stdcall AddContact (
        /*[in]*/ struct IContact * _contact ) = 0;
      virtual HRESULT __stdcall RemoveContact (
        /*[in]*/ struct IContact * _contact ) = 0;
      virtual HRESULT __stdcall get_LastSubscribedRefreshRate (
        /*[out,retval]*/ enum ContactSubscriptionRefreshRate * _subscriptionFreshness ) = 0;
      virtual HRESULT __stdcall get_LastSubscribedContactInformation (
        /*[out,retval]*/ SAFEARRAY * * _contactInformationTypes ) = 0;
      virtual HRESULT __stdcall get_Contacts (
        /*[out,retval]*/ SAFEARRAY * * _contacts ) = 0;
      virtual HRESULT __stdcall AddContactByUri (
        /*[in]*/ BSTR _contactUri ) = 0;
      virtual HRESULT __stdcall AddContacts (
        /*[in]*/ SAFEARRAY * _contacts ) = 0;
};

struct __declspec(uuid("e6b3f5fa-4208-4c37-ae07-f6435b68d693"))
IContactManager : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Groups (
        /*[out,retval]*/ struct IGroupCollection * * _groups ) = 0;
      virtual HRESULT __stdcall GetContactByUri (
        /*[in]*/ BSTR _contactUri,
        /*[out,retval]*/ struct IContact * * _contact ) = 0;
      virtual HRESULT __stdcall Lookup (
        /*[in]*/ BSTR _lookupString,
        /*[in]*/ IUnknown * _contactsAndGroupsCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CreateSubscription (
        /*[out,retval]*/ struct IContactSubscription * * _subscription ) = 0;
      virtual HRESULT __stdcall Search (
        /*[in]*/ BSTR _searchString,
        /*[in]*/ enum SearchProviders _providers,
        /*[in]*/ enum SearchFields _searchFields,
        /*[in]*/ enum SearchOptions _searchOptions,
        /*[in]*/ unsigned long _maxResults,
        /*[in]*/ IUnknown * _contactsAndGroupsCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall RemoveContactFromAllGroups (
        /*[in]*/ struct IContact * _contact,
        /*[in]*/ IUnknown * _contactsAndGroupsCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetSearchProviderID (
        /*[in]*/ enum SearchProviders _provider,
        /*[out,retval]*/ BSTR * _providerId ) = 0;
      virtual HRESULT __stdcall AddCustomGroup (
        /*[in]*/ BSTR _customGroupName,
        /*[in]*/ IUnknown * _contactsAndGroupsCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall AddDistributionGroup (
        /*[in]*/ struct IDistributionGroup * _distributionGroup,
        /*[in]*/ IUnknown * _contactsAndGroupsCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall RemoveGroup (
        /*[in]*/ struct IGroup * _group,
        /*[in]*/ IUnknown * _contactsAndGroupsCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall SuspendSubscriptions ( ) = 0;
      virtual HRESULT __stdcall ResumeSubscriptions ( ) = 0;
      virtual HRESULT __stdcall GetSearchFields (
        /*[out,retval]*/ enum SearchFields * _searchFields ) = 0;
      virtual HRESULT __stdcall GetSearchProviderStatus (
        /*[in]*/ enum SearchProviders _provider,
        /*[out,retval]*/ enum SearchProviderStatusType * _status ) = 0;
      virtual HRESULT __stdcall GetExpertSearchQueryString (
        /*[in]*/ BSTR _searchString,
        /*[out,retval]*/ BSTR * _queryString ) = 0;
      virtual HRESULT __stdcall GetSelfContact (
        /*[out,retval]*/ struct IContact * * _selfContact ) = 0;
      virtual HRESULT __stdcall GetContactByTel (
        /*[in]*/ BSTR _telUri,
        /*[out,retval]*/ struct IContact * * _contact ) = 0;
};

struct __declspec(uuid("4b3a2412-c149-4e52-a713-6cebcf47503e"))
IGroupMemberChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Contact (
        /*[out,retval]*/ struct IContact * * _contact ) = 0;
};

struct __declspec(uuid("dcdea425-b5f2-4719-a3b4-69ffb9770be6"))
IRecentContacts : IGroup
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IContact * * _contact ) = 0;
};

struct __declspec(uuid("18a9a79a-cbc7-4f8b-8b2b-9b40ca75ec3c"))
IFrequentContacts : IGroup
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IContact * * _contact ) = 0;
};

struct __declspec(uuid("10fdd9ba-0cba-4958-b6c8-d0912bf2703f"))
IContact2 : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall SetAdditionalUris (
        /*[in]*/ SAFEARRAY * _contactUris ) = 0;
      virtual HRESULT __stdcall GetRawContactbyUri (
        /*[in]*/ BSTR _contactUri,
        /*[out,retval]*/ struct IContact * * _contact ) = 0;
      virtual HRESULT __stdcall IsAnonymous (
        /*[out,retval]*/ VARIANT_BOOL * _isAnonymous ) = 0;
};

struct __declspec(uuid("4da899d3-fac7-4f43-8e4a-1987988fbc9b"))
ISelf : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall PublishContactInformation (
        /*[in]*/ SAFEARRAY * _publishablePresenceItemTypes,
        /*[in]*/ SAFEARRAY * _publishablePresenceItemValues,
        /*[in]*/ IUnknown * _selfCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall get_Contact (
        /*[out,retval]*/ struct IContact * * _selfContact ) = 0;
      virtual HRESULT __stdcall CreatePhone (
        /*[in]*/ enum ContactEndpointType _phoneType,
        /*[in]*/ BSTR _phoneUri,
        /*[in]*/ VARIANT_BOOL _toBePublished,
        /*[out,retval]*/ struct IPhone * * _phone ) = 0;
      virtual HRESULT __stdcall CanSetPhone (
        /*[in]*/ enum ContactEndpointType _phoneType,
        /*[out,retval]*/ VARIANT_BOOL * _canSet ) = 0;
      virtual HRESULT __stdcall get_Phones (
        /*[out,retval]*/ SAFEARRAY * * _phones ) = 0;
      virtual HRESULT __stdcall SetPhones (
        /*[in]*/ SAFEARRAY * _phones,
        /*[in]*/ IUnknown * _selfCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetPublishableCustomAvailabilityStates (
        /*[in]*/ long _localeId,
        /*[out,retval]*/ SAFEARRAY * * _states ) = 0;
      virtual HRESULT __stdcall get_Permissions (
        /*[out,retval]*/ SAFEARRAY * * _permissions ) = 0;
      virtual HRESULT __stdcall get_PhotoDisplayed (
        /*[out,retval]*/ VARIANT_BOOL * _displayed ) = 0;
      virtual HRESULT __stdcall get_TestCallEndpoint (
        /*[out,retval]*/ struct IContactEndpoint * * _testCallEndpoint ) = 0;
      virtual HRESULT __stdcall SetPhone (
        /*[in]*/ enum ContactEndpointType _phoneType,
        /*[in]*/ BSTR _phoneUri,
        /*[in]*/ VARIANT_BOOL _toBePublished,
        /*[in]*/ IUnknown * _selfCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall RemovePhone (
        /*[in]*/ enum ContactEndpointType _phoneType,
        /*[in]*/ IUnknown * _selfCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetPhone (
        /*[in]*/ enum ContactEndpointType _phoneType,
        /*[out,retval]*/ struct IPhone * * _phone ) = 0;
      virtual HRESULT __stdcall get_IsInResiliencyMode (
        /*[out,retval]*/ VARIANT_BOOL * _isResiliencyMode ) = 0;
      virtual HRESULT __stdcall GetSettingInternal (
        /*[in]*/ enum SettingItemTypeInternal _settingItem,
        /*[out,retval]*/ VARIANT * _setting ) = 0;
};

struct __declspec(uuid("87084018-7e74-4f92-ab39-1cf03188580e"))
IUtilities : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall CreateNamedPropertyList (
        /*[out,retval]*/ struct INamedPropertyList * * _namedProperties ) = 0;
      virtual HRESULT __stdcall get_IsDefaultProvider (
        /*[out,retval]*/ VARIANT_BOOL * _isDefaultProvider ) = 0;
      virtual HRESULT __stdcall SetAsDefaultProvider ( ) = 0;
      virtual HRESULT __stdcall CanInvoke (
        /*[in]*/ enum UtilitiesAction _utilitiesAction,
        /*[in]*/ VARIANT _actionContext,
        /*[out,retval]*/ VARIANT_BOOL * _canInvoke ) = 0;
      virtual HRESULT __stdcall ShowPropertiesDialog (
        /*[in]*/ VARIANT _context ) = 0;
      virtual HRESULT __stdcall AddToExternalContacts (
        /*[in]*/ struct IContact * _contactContext,
        /*[in]*/ IUnknown * _utilitiesCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall FindPreviousConversations (
        /*[in]*/ VARIANT _context,
        /*[in]*/ IUnknown * _utilitiesCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall ComposeEmail (
        /*[in]*/ VARIANT _to,
        /*[in]*/ BSTR _subject,
        /*[in]*/ IUnknown * _utilitiesCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall ScheduleMeeting (
        /*[in]*/ VARIANT _paticipants,
        /*[in]*/ BSTR _subject,
        /*[in]*/ IUnknown * _utilitiesCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("86225e5c-7595-4d1c-985d-8a0458d714c0"))
IClient2 : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_PrivateContactManager (
        /*[out,retval]*/ struct IContactManager * * _value ) = 0;
      virtual HRESULT __stdcall get_RoomManager (
        /*[out,retval]*/ struct IRoomManager * * _value ) = 0;
      virtual HRESULT __stdcall get_Settings (
        /*[out,retval]*/ struct IClientSettings * * _value ) = 0;
};

struct __declspec(uuid("a4e1d1e4-7839-473b-95fa-5d33ddc0d2dd"))
IConversationCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IConversation * * _conversation ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall IndexOf (
        /*[in]*/ struct IConversation * _conversation,
        /*[out,retval]*/ long * _index ) = 0;
};

struct __declspec(uuid("533891f1-a460-43e6-8bda-6d0d2c13adde"))
IConversation : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ enum ConversationState * _state ) = 0;
      virtual HRESULT __stdcall get_SelfParticipant (
        /*[out,retval]*/ struct IParticipant * * _participant ) = 0;
      virtual HRESULT __stdcall get_Modalities (
        /*[out,retval]*/ struct IModalityDictionary * * _modalities ) = 0;
      virtual HRESULT __stdcall get_Participants (
        /*[out,retval]*/ struct IParticipantCollection * * _participants ) = 0;
      virtual HRESULT __stdcall AddParticipant (
        /*[in]*/ IDispatch * _contactOrCollaborationEndpoint,
        /*[out,retval]*/ struct IParticipant * * _participant ) = 0;
      virtual HRESULT __stdcall RemoveParticipant (
        /*[in]*/ struct IParticipant * _participant ) = 0;
      virtual HRESULT __stdcall get_Contexts (
        /*[out,retval]*/ struct IConversationContextList * * _contexts ) = 0;
      virtual HRESULT __stdcall get_History (
        /*[out,retval]*/ struct IConversationHistory * * _history ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IConversationPropertyDictionary * * _properties ) = 0;
      virtual HRESULT __stdcall SetProperty (
        /*[in]*/ enum ConversationProperty _propertyType,
        /*[in]*/ VARIANT _propertyValue,
        /*[in]*/ IUnknown * _conversationCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall SendInitialContext (
        /*[in]*/ SAFEARRAY * _contextTypes,
        /*[in]*/ SAFEARRAY * _contextDatas,
        /*[in]*/ IUnknown * _conversationCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall SendContextData (
        /*[in]*/ BSTR _applicationId,
        /*[in]*/ BSTR _dataType,
        /*[in]*/ BSTR _data,
        /*[in]*/ IUnknown * _conversationCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetApplicationData (
        /*[in]*/ BSTR _appGuid,
        /*[out,retval]*/ BSTR * _appData ) = 0;
      virtual HRESULT __stdcall get_ConversationManager (
        /*[out,retval]*/ struct IConversationManager * * _conversationsManager ) = 0;
      virtual HRESULT __stdcall Merge (
        /*[in]*/ struct IConversation * _conversation,
        /*[in]*/ enum ModalityTypes _modalityTypes,
        /*[in]*/ IUnknown * _conversationCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall AdmitParticipants (
        /*[in]*/ SAFEARRAY * _participants,
        /*[in]*/ IUnknown * _conversationCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall DenyParticipants (
        /*[in]*/ SAFEARRAY * _participants,
        /*[in]*/ IUnknown * _conversationCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Park (
        /*[in]*/ IUnknown * _conversationCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall End ( ) = 0;
      virtual HRESULT __stdcall CanInvoke (
        /*[in]*/ enum ConversationAction _action,
        /*[out,retval]*/ VARIANT_BOOL * _enabled ) = 0;
      virtual HRESULT __stdcall CanSetProperty (
        /*[in]*/ enum ConversationProperty _convProperty,
        /*[out,retval]*/ VARIANT_BOOL * _canSet ) = 0;
      virtual HRESULT __stdcall get_Direction (
        /*[out,retval]*/ enum ConversationDirection * _direction ) = 0;
      virtual HRESULT __stdcall MuteParticipants (
        /*[in]*/ SAFEARRAY * _participants,
        /*[in]*/ IUnknown * _conversationCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall UnmuteParticipants (
        /*[in]*/ SAFEARRAY * _participants,
        /*[in]*/ IUnknown * _conversationCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("f5996c71-77ab-4020-9776-c2c270fef2ae"))
IConversationHistory : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Conversation (
        /*[out,retval]*/ struct IConversation * * _conversation ) = 0;
      virtual HRESULT __stdcall get_IsPastHistoryRetrieved (
        /*[out,retval]*/ VARIANT_BOOL * _isActive ) = 0;
      virtual HRESULT __stdcall get_CurrentHistory (
        /*[out,retval]*/ SAFEARRAY * * _currentHistory ) = 0;
      virtual HRESULT __stdcall get_AllHistory (
        /*[out,retval]*/ SAFEARRAY * * _allHistories ) = 0;
      virtual HRESULT __stdcall get_PastHistory (
        /*[out,retval]*/ SAFEARRAY * * _pastHistories ) = 0;
      virtual HRESULT __stdcall RetrievePastHistory (
        /*[in]*/ IUnknown * _conversationHistoryCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall RetrievePastHistoryByEntryId (
        /*[in]*/ BSTR _entryId,
        /*[in]*/ IUnknown * _conversationHistoryCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall SaveAllHistory (
        /*[in]*/ IUnknown * _conversationHistoryCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CanInvoke (
        /*[in]*/ enum ConversationHistoryAction _action,
        /*[out,retval]*/ VARIANT_BOOL * _enabled ) = 0;
};

struct __declspec(uuid("c31fb52d-9e9e-45ad-a102-5218e1685b78"))
IConversationManager : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_SummarizerStates (
        /*[out,retval]*/ enum ConversationSummarizerStates * _state ) = 0;
      virtual HRESULT __stdcall get_Conversations (
        /*[out,retval]*/ struct IConversationCollection * * _conversations ) = 0;
      virtual HRESULT __stdcall AddConversation (
        /*[out,retval]*/ struct IConversation * * _conversation ) = 0;
      virtual HRESULT __stdcall JoinConference (
        /*[in]*/ BSTR _conferenceUrl,
        /*[out,retval]*/ struct IConversation * * _conversation ) = 0;
      virtual HRESULT __stdcall get_VoiceMailsManager (
        /*[out,retval]*/ struct IVoiceMailsManager * * _voiceMails ) = 0;
      virtual HRESULT __stdcall get_PreviousConversationsManagers (
        /*[out,retval]*/ struct IPreviousConversationsManagerDictionary * * _pastConversationsManagers ) = 0;
};

struct __declspec(uuid("8057f99d-292c-4371-926b-5312e61a3a40"))
ISummarizerStatesChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_TriggeringConversation (
        /*[out,retval]*/ struct IConversation * * _triggeringConversation ) = 0;
      virtual HRESULT __stdcall get_OldState (
        /*[out,retval]*/ enum ConversationSummarizerStates * _oldState ) = 0;
      virtual HRESULT __stdcall get_NewState (
        /*[out,retval]*/ enum ConversationSummarizerStates * _newState ) = 0;
};

struct __declspec(uuid("ab2d1a67-2b3d-4e25-b21c-03f4bfc4c2be"))
IConversationManagerEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Conversation (
        /*[out,retval]*/ struct IConversation * * _conversation ) = 0;
};

struct __declspec(uuid("25a47938-5c31-4f43-9108-da01ee5869c0"))
IVoiceMail : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * _entryId ) = 0;
      virtual HRESULT __stdcall get_EntryId (
        /*[out,retval]*/ BSTR * _entryId ) = 0;
      virtual HRESULT __stdcall get_Sender (
        /*[out,retval]*/ BSTR * _sender ) = 0;
      virtual HRESULT __stdcall get_SenderUri (
        /*[out,retval]*/ BSTR * _senderUri ) = 0;
      virtual HRESULT __stdcall get_ReceiveTime (
        /*[out,retval]*/ DATE * _received ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct INamedPropertyList * * _properties ) = 0;
      virtual HRESULT __stdcall get_IsRead (
        /*[out,retval]*/ VARIANT_BOOL * _isRead ) = 0;
      virtual HRESULT __stdcall CallVoiceMail (
        /*[out,retval]*/ struct IConversation * * _conversation ) = 0;
      virtual HRESULT __stdcall Delete (
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("f644c610-a146-4a56-8338-a69c45c71cef"))
IVoiceMailCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IVoiceMail * * _voiceMail ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
};

struct __declspec(uuid("18ba13c7-a64e-4301-ba51-d1bfb3c1c9d5"))
IVoiceMailBatchEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_VoiceMails (
        /*[out,retval]*/ struct IVoiceMailCollection * * _value ) = 0;
};

struct __declspec(uuid("b1f5bcd8-69a6-4fc3-9ef0-9bd4ad999061"))
IConversationWindow : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Conversation (
        /*[out,retval]*/ struct IConversation * * _conversation ) = 0;
      virtual HRESULT __stdcall get_Left (
        /*[out,retval]*/ long * _left ) = 0;
      virtual HRESULT __stdcall get_Top (
        /*[out,retval]*/ long * _top ) = 0;
      virtual HRESULT __stdcall get_Width (
        /*[out,retval]*/ long * _width ) = 0;
      virtual HRESULT __stdcall get_Height (
        /*[out,retval]*/ long * _height ) = 0;
      virtual HRESULT __stdcall Move (
        /*[in]*/ long _left,
        /*[in]*/ long _top ) = 0;
      virtual HRESULT __stdcall Resize (
        /*[in]*/ long _width,
        /*[in]*/ long _height ) = 0;
      virtual HRESULT __stdcall MoveAndResize (
        /*[in]*/ long _left,
        /*[in]*/ long _top,
        /*[in]*/ long _width,
        /*[in]*/ long _height ) = 0;
      virtual HRESULT __stdcall Close ( ) = 0;
      virtual HRESULT __stdcall Dock (
        /*[in]*/ __int64 _parentHWND ) = 0;
      virtual HRESULT __stdcall Undock ( ) = 0;
      virtual HRESULT __stdcall get_IsDocked (
        /*[out,retval]*/ VARIANT_BOOL * _isDocked ) = 0;
      virtual HRESULT __stdcall get_IsFullScreen (
        /*[out,retval]*/ VARIANT_BOOL * _isFullScreen ) = 0;
      virtual HRESULT __stdcall get_Handle (
        /*[out,retval]*/ __int64 * _hwnd ) = 0;
      virtual HRESULT __stdcall get_ParentWindow (
        /*[out,retval]*/ __int64 * _parentWindow ) = 0;
      virtual HRESULT __stdcall OpenExtensibilityWindow (
        /*[in]*/ BSTR _applicationId,
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CloseExtensibilityWindow (
        /*[in]*/ BSTR _applicationId ) = 0;
      virtual HRESULT __stdcall FlashExtensibilityWindow (
        /*[in]*/ BSTR _applicationId,
        /*[in]*/ VARIANT_BOOL _flash ) = 0;
};

struct __declspec(uuid("2edaf003-14a6-4faf-8a72-99526d048d63"))
IAutomation : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall LaunchAddContactWizard (
        /*[in]*/ BSTR _contactEmail ) = 0;
      virtual HRESULT __stdcall StartConversation (
        /*[in]*/ enum AutomationModalities _conversationModes,
        /*[in]*/ SAFEARRAY * _participantUris,
        /*[in]*/ SAFEARRAY * _contextTypes,
        /*[in]*/ SAFEARRAY * _contextDatas,
        /*[out,retval]*/ struct IConversationWindow * * _window ) = 0;
      virtual HRESULT __stdcall JoinConference (
        /*[in]*/ BSTR _conferenceUrl,
        /*[out,retval]*/ struct IConversationWindow * * _window ) = 0;
      virtual HRESULT __stdcall StartConversationEx (
        /*[in]*/ enum AutomationModalities _conversationModes,
        /*[in]*/ SAFEARRAY * _participantUris,
        /*[in]*/ SAFEARRAY * _contextTypes,
        /*[in]*/ SAFEARRAY * _contextDatas,
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall JoinConferenceEx (
        /*[in]*/ BSTR _conferenceUrl,
        /*[in]*/ __int64 _parentHWND,
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GetConversationWindow (
        /*[in]*/ struct IConversation * _conversation,
        /*[out,retval]*/ struct IConversationWindow * * _window ) = 0;
};

struct __declspec(uuid("ee9f3e74-ac61-469e-80d6-e22bf4eeff5c"))
IClient : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum ClientType * _type ) = 0;
      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ enum ClientState * _state ) = 0;
      virtual HRESULT __stdcall get_Uri (
        /*[out,retval]*/ BSTR * _uri ) = 0;
      virtual HRESULT __stdcall get_Self (
        /*[out,retval]*/ struct ISelf * * _value ) = 0;
      virtual HRESULT __stdcall get_ContactManager (
        /*[out,retval]*/ struct IContactManager * * _value ) = 0;
      virtual HRESULT __stdcall get_ConversationManager (
        /*[out,retval]*/ struct IConversationManager * * _conversationManager ) = 0;
      virtual HRESULT __stdcall get_ConferenceScheduler (
        /*[out,retval]*/ struct IConferenceScheduler * * _value ) = 0;
};

struct __declspec(uuid("c413a27f-601e-4261-b9e3-0eb966f4e7d1"))
ILyncClient : IClient
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall SignIn (
        /*[in]*/ BSTR _userUri,
        /*[in]*/ BSTR _domainAndUsername,
        /*[in]*/ BSTR _password,
        /*[in]*/ IUnknown * _CommunicatorClientCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall SignOut (
        /*[in]*/ IUnknown * _CommunicatorClientCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall get_SignInConfiguration (
        /*[out,retval]*/ struct ISignInConfiguration * * _configuration ) = 0;
      virtual HRESULT __stdcall get_Capabilities (
        /*[out,retval]*/ enum LyncClientCapabilityTypes * _capabilities ) = 0;
      virtual HRESULT __stdcall get_DeviceManager (
        /*[out,retval]*/ struct IDeviceManager * * _deviceManager ) = 0;
      virtual HRESULT __stdcall get_DelegatorClients (
        /*[out,retval]*/ SAFEARRAY * * _delegateClients ) = 0;
      virtual HRESULT __stdcall get_Utilities (
        /*[out,retval]*/ struct IUtilities * * _utilities ) = 0;
      virtual HRESULT __stdcall CreateApplicationRegistration (
        /*[in]*/ BSTR _appGuid,
        /*[in]*/ BSTR _appName,
        /*[out,retval]*/ struct IApplicationRegistration * * _registration ) = 0;
      virtual HRESULT __stdcall get_InSuppressedMode (
        /*[out,retval]*/ VARIANT_BOOL * _value ) = 0;
      virtual HRESULT __stdcall Initialize (
        /*[in]*/ BSTR _clientName,
        /*[in]*/ BSTR _version,
        /*[in]*/ BSTR _clientShortName,
        /*[in]*/ BSTR _clientNameAbbreviation,
        /*[in]*/ BSTR _clientLongName,
        /*[in]*/ enum SupportedFeatures _supportedFeatures,
        /*[in]*/ IUnknown * _CommunicatorClientCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Shutdown (
        /*[in]*/ IUnknown * _CommunicatorClientCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("80edbf3a-812e-42b5-a67f-6cc6d9a19a6c"))
IDelegatorClient : IClient
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_LyncClient (
        /*[out,retval]*/ struct ILyncClient * * _client ) = 0;
};

struct __declspec(uuid("744bab86-b15f-4bd5-b7dd-476b0ce0162a"))
IDelegatorClientCollectionEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_DelegatorClient (
        /*[out,retval]*/ struct IDelegatorClient * * _delegator ) = 0;
};

struct __declspec(uuid("09028936-cd20-43ef-98cf-56726addba34"))
IUCOfficeIntegrationHelper : IUnknown
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GetCommunicatorClient (
        /*[out,retval]*/ struct ILyncClient * * _value ) = 0;
      virtual HRESULT __stdcall GetUIAutomation (
        /*[out,retval]*/ struct IAutomation * * _value ) = 0;
      virtual HRESULT __stdcall GetInteropHelper (
        /*[out,retval]*/ IDispatch * * _value ) = 0;
      virtual HRESULT __stdcall GetAuthenticationInfo (
        /*[out,retval]*/ BSTR * _value ) = 0;
      virtual HRESULT __stdcall Cleanup ( ) = 0;
      virtual HRESULT __stdcall Notify (
        /*[in]*/ enum OIHelperAction Action,
        /*[in]*/ HRESULT hrRes ) = 0;
};

struct __declspec(uuid("4875c050-ba9c-4a05-891e-e7b0a9463664"))
IParticipant : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Modalities (
        /*[out,retval]*/ struct IModalityDictionary * * _modalities ) = 0;
      virtual HRESULT __stdcall get_IsSelf (
        /*[out,retval]*/ VARIANT_BOOL * _isSelf ) = 0;
      virtual HRESULT __stdcall get_Contact (
        /*[out,retval]*/ struct IContact * * _contact ) = 0;
      virtual HRESULT __stdcall get_Conversation (
        /*[out,retval]*/ struct IConversation * * _conversation ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IParticipantPropertyDictionary * * _properties ) = 0;
      virtual HRESULT __stdcall SetProperty (
        /*[in]*/ enum ParticipantProperty _propertyType,
        /*[in]*/ VARIANT _propertyValue,
        /*[in]*/ IUnknown * _participantCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall SetMute (
        /*[in]*/ VARIANT_BOOL _mute,
        /*[in]*/ IUnknown * _participantCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall get_IsMuted (
        /*[out,retval]*/ VARIANT_BOOL * _isMuted ) = 0;
      virtual HRESULT __stdcall CanSetProperty (
        /*[in]*/ enum ParticipantProperty _participantProperty,
        /*[out,retval]*/ VARIANT_BOOL * _canSet ) = 0;
      virtual HRESULT __stdcall Admit (
        /*[in]*/ IUnknown * _participantCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Deny (
        /*[in]*/ IUnknown * _participantCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CanAdmit (
        /*[out,retval]*/ VARIANT_BOOL * _canAdmit ) = 0;
      virtual HRESULT __stdcall CanDeny (
        /*[out,retval]*/ VARIANT_BOOL * _canDeny ) = 0;
      virtual HRESULT __stdcall CanBeMuted (
        /*[out,retval]*/ VARIANT_BOOL * _canBeMuted ) = 0;
      virtual HRESULT __stdcall CanInvoke (
        /*[in]*/ enum ParticipantAction _action,
        /*[out,retval]*/ VARIANT_BOOL * _enabled ) = 0;
};

struct __declspec(uuid("d4b240b9-059b-4d13-8997-31b282b3bd43"))
IModality : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum ModalityTypes * _type ) = 0;
      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ enum ModalityState * _state ) = 0;
      virtual HRESULT __stdcall get_Conversation (
        /*[out,retval]*/ struct IConversation * * _conversation ) = 0;
      virtual HRESULT __stdcall get_Participant (
        /*[out,retval]*/ struct IParticipant * * _participant ) = 0;
      virtual HRESULT __stdcall get_Endpoint (
        /*[out,retval]*/ struct IContactEndpoint * * _collaborationEndpoint ) = 0;
      virtual HRESULT __stdcall put_Endpoint (
        /*[in]*/ struct IContactEndpoint * _collaborationEndpoint ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IModalityPropertyDictionary * * _properties ) = 0;
      virtual HRESULT __stdcall SetProperty (
        /*[in]*/ enum ModalityProperty _propertyType,
        /*[in]*/ VARIANT _propertyValue,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Connect (
        /*[in]*/ enum ModalityConnectOptions _options,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Disconnect (
        /*[in]*/ enum ModalityDisconnectReason _reason,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Reject (
        /*[in]*/ enum ModalityDisconnectReason _reason ) = 0;
      virtual HRESULT __stdcall Accept ( ) = 0;
      virtual HRESULT __stdcall Hold (
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Retrieve (
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Forward (
        /*[in]*/ IDispatch * _contactOrCollaborationEndpoint,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Transfer (
        /*[in]*/ IDispatch * _contactOrCollaborationEndpoint,
        /*[in]*/ enum TransferOptions _options,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall ConsultativeTransfer (
        /*[in]*/ struct IConversation * _conversation,
        /*[in]*/ enum TransferOptions _options,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CanInvoke (
        /*[in]*/ enum ModalityAction _action,
        /*[out,retval]*/ VARIANT_BOOL * _enabled ) = 0;
      virtual HRESULT __stdcall CanSetProperty (
        /*[in]*/ enum ModalityProperty _modalityProperty,
        /*[out,retval]*/ VARIANT_BOOL * _canSet ) = 0;
};

struct __declspec(uuid("ed5040fa-8f8f-4876-bf6d-9c62f5effd5b"))
IModalityDictionary : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ enum ModalityTypes _modalityType,
        /*[out,retval]*/ struct IModality * * _itemValue ) = 0;
      virtual HRESULT __stdcall get_Keys (
        /*[out,retval]*/ SAFEARRAY * * _keys ) = 0;
      virtual HRESULT __stdcall get_Values (
        /*[out,retval]*/ SAFEARRAY * * _values ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum ModalityTypes _modalityType,
        /*[out]*/ struct IModality * * _itemValue,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetKeyAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ enum ModalityTypes * _key ) = 0;
      virtual HRESULT __stdcall GetValueAt (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IModality * * _value ) = 0;
      virtual HRESULT __stdcall ContainsKey (
        /*[in]*/ enum ModalityTypes _modalityType,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
};

struct __declspec(uuid("fd02e775-cb9d-4ce3-9243-115c81e5f11c"))
IParticipantCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IParticipant * * _participant ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall IndexOf (
        /*[in]*/ struct IParticipant * _participant,
        /*[out,retval]*/ long * _index ) = 0;
};

struct __declspec(uuid("62a75516-c79b-42d7-8b49-3ba492c2b385"))
IInstantMessageModality : IModality
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall SendMessage (
        /*[in]*/ BSTR _message,
        /*[in]*/ struct INamedPropertyList * _additionalFormats,
        /*[in]*/ enum ModalityConnectOptions _options,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall SetComposing (
        /*[in]*/ VARIANT_BOOL _isComposing,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall SetCapabilities (
        /*[in]*/ enum InstantMessageCapabilities _capabilities ) = 0;
      virtual HRESULT __stdcall get_IsTyping (
        /*[out,retval]*/ VARIANT_BOOL * _isTyping ) = 0;
      virtual HRESULT __stdcall get_Capabilities (
        /*[out,retval]*/ long * _capabilities ) = 0;
};

struct __declspec(uuid("61c747f6-157c-4cae-a37c-ec4352175860"))
IChannel : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum ChannelType * _type ) = 0;
      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ enum ChannelState * _state ) = 0;
      virtual HRESULT __stdcall get_Modality (
        /*[out,retval]*/ struct IModality * * _modality ) = 0;
      virtual HRESULT __stdcall get_IsContributing (
        /*[out,retval]*/ VARIANT_BOOL * _isContributing ) = 0;
      virtual HRESULT __stdcall Start (
        /*[in]*/ IUnknown * _mediaChannelCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall Stop (
        /*[in]*/ IUnknown * _mediaChannelCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CanInvoke (
        /*[in]*/ enum ChannelAction _action,
        /*[out,retval]*/ VARIANT_BOOL * _enabled ) = 0;
};

struct __declspec(uuid("8e839af9-9508-475a-aa47-3ee8a946b75e"))
IAudioChannel : IChannel
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall SendDtmf (
        /*[in]*/ BSTR _tones,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("e66ca1cc-9db0-467e-9c60-6a832acd6780"))
IVideoChannel : IChannel
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_RenderVideoWindow (
        /*[out,retval]*/ struct IVideoWindow * * _renderVideoWindow ) = 0;
      virtual HRESULT __stdcall get_CaptureVideoWindow (
        /*[out,retval]*/ struct IVideoWindow * * _captureVideoWindow ) = 0;
};

struct __declspec(uuid("e4e1d0af-5b70-45af-9a0e-f548d8fbd17c"))
IAVModality : IModality
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_AudioChannel (
        /*[out,retval]*/ struct IAudioChannel * * _channel ) = 0;
      virtual HRESULT __stdcall get_VideoChannel (
        /*[out,retval]*/ struct IVideoChannel * * _channel ) = 0;
};

struct __declspec(uuid("ac10d6e9-4a8c-4484-b8f4-ca1e36347aae"))
IApplicationSharingModality : IModality
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ShareableResources (
        /*[out,retval]*/ struct ISharingResourceList * * _resources ) = 0;
      virtual HRESULT __stdcall CreateEmptyShareableResourceList (
        /*[out,retval]*/ struct ISharingResourceList * * _resources ) = 0;
      virtual HRESULT __stdcall ShareResources (
        /*[in]*/ struct ISharingResourceList * _resources,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall ShareDesktop (
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall ShareMonitor (
        /*[in]*/ int _monitorId,
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall get_LocalSharedResources (
        /*[out,retval]*/ struct ISharingResourceList * * _resources ) = 0;
      virtual HRESULT __stdcall get_Controller (
        /*[out,retval]*/ struct IParticipant * * _participant ) = 0;
      virtual HRESULT __stdcall get_ParticipationState (
        /*[out,retval]*/ enum ParticipationState * _state ) = 0;
      virtual HRESULT __stdcall RequestControl (
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall ReleaseControl (
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall GrantControl (
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall RevokeControl (
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall AcceptControlRequest (
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall DeclineControlRequest (
        /*[in]*/ IUnknown * _modalityCallback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CanShare (
        /*[in]*/ enum SharingResourceType _resourceType,
        /*[out,retval]*/ enum CanShareDetail * _canShareDetail ) = 0;
      virtual HRESULT __stdcall get_Sharer (
        /*[out,retval]*/ struct IParticipant * * _participant ) = 0;
};

struct __declspec(uuid("3cd1db65-d85b-4912-8536-6bd8e19f1738"))
IVideoChannelCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IVideoChannel * * _channel ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
      virtual HRESULT __stdcall IndexOf (
        /*[in]*/ struct IVideoChannel * _videoChannel,
        /*[out,retval]*/ long * _index ) = 0;
};

struct __declspec(uuid("064e6072-5592-4d13-afdb-55f5a494e137"))
IParticipantCollectionChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Participant (
        /*[out,retval]*/ struct IParticipant * * _participant ) = 0;
      virtual HRESULT __stdcall get_StatusCode (
        /*[out,retval]*/ long * _statusCode ) = 0;
};

struct __declspec(uuid("ca7eeb7a-7dc3-4ffe-a174-23db5a003c04"))
IStreamStateChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Channel (
        /*[out,retval]*/ struct IChannel * * _channel ) = 0;
      virtual HRESULT __stdcall get_Direction (
        /*[out,retval]*/ enum MediaStreamDirection * _streamDirection ) = 0;
      virtual HRESULT __stdcall get_IsActive (
        /*[out,retval]*/ VARIANT_BOOL * _isStreaming ) = 0;
};

struct __declspec(uuid("32430741-330b-48df-96f6-0cfe00484d40"))
IControlRequestReceivedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Requester (
        /*[out,retval]*/ struct IParticipant * * _requester ) = 0;
};

struct __declspec(uuid("b335ae5e-e4ea-49d3-b03b-646a96fe66d3"))
IShareableContent : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum ShareableContentType * _type ) = 0;
      virtual HRESULT __stdcall get_State (
        /*[out,retval]*/ enum ShareableContentState * _state ) = 0;
      virtual HRESULT __stdcall get_Title (
        /*[out,retval]*/ BSTR * _title ) = 0;
      virtual HRESULT __stdcall get_Owner (
        /*[out,retval]*/ struct IContact * * _owner ) = 0;
      virtual HRESULT __stdcall get_Presenter (
        /*[out,retval]*/ struct IParticipant * * _presenter ) = 0;
      virtual HRESULT __stdcall get_CreationTime (
        /*[out,retval]*/ DATE * _creation ) = 0;
      virtual HRESULT __stdcall get_IsActive (
        /*[out,retval]*/ VARIANT_BOOL * _isActive ) = 0;
      virtual HRESULT __stdcall get_Properties (
        /*[out,retval]*/ struct IShareableContentPropertyDictionary * * _properties ) = 0;
      virtual HRESULT __stdcall get_Modality (
        /*[out,retval]*/ struct IModality * * _modality ) = 0;
      virtual HRESULT __stdcall Upload ( ) = 0;
      virtual HRESULT __stdcall Remove ( ) = 0;
      virtual HRESULT __stdcall Present ( ) = 0;
      virtual HRESULT __stdcall StopPresenting ( ) = 0;
      virtual HRESULT __stdcall SaveAnnotation (
        /*[in]*/ BSTR _fileName,
        /*[in]*/ enum ContentSavingFileType _fileType ) = 0;
      virtual HRESULT __stdcall DownloadFile (
        /*[in]*/ BSTR _localPath,
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CanInvoke (
        /*[in]*/ enum ShareableContentAction _action,
        /*[out]*/ HRESULT * _hrReason,
        /*[out,retval]*/ VARIANT_BOOL * _enabled ) = 0;
      virtual HRESULT __stdcall TakeOverAsPresenter ( ) = 0;
      virtual HRESULT __stdcall ClearAllAnnotations ( ) = 0;
};

struct __declspec(uuid("a58f54d2-9786-4309-964d-96549aec7611"))
IShareableContentCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long index,
        /*[out,retval]*/ struct IShareableContent * * _value ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
};

struct __declspec(uuid("38e5ec21-81a1-45a7-94d6-b812976231a4"))
IContentSharingModality : IModality
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ContentCollection (
        /*[out,retval]*/ struct IShareableContentCollection * * _contents ) = 0;
      virtual HRESULT __stdcall get_ActiveContent (
        /*[out,retval]*/ struct IShareableContent * * _activeContent ) = 0;
      virtual HRESULT __stdcall get_ActivePresenter (
        /*[out,retval]*/ struct IParticipant * * _activePresenter ) = 0;
      virtual HRESULT __stdcall CreateContent (
        /*[in]*/ enum ShareableContentType _contentType,
        /*[in]*/ BSTR _title,
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
      virtual HRESULT __stdcall CreateContentFromFile (
        /*[in]*/ enum ShareableContentType _contentType,
        /*[in]*/ BSTR _title,
        /*[in]*/ BSTR _fileName,
        /*[in]*/ VARIANT_BOOL _includeNativeFile,
        /*[in]*/ IUnknown * _callback,
        /*[in]*/ VARIANT _state,
        /*[out,retval]*/ struct IAsynchronousOperation * * _asyncOperation ) = 0;
};

struct __declspec(uuid("17635563-2fa5-46ad-8f47-3f0ff958ece2"))
IContentCollectionChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Item (
        /*[out,retval]*/ struct IShareableContent * * _content ) = 0;
};

struct __declspec(uuid("48b3099d-841d-4caa-9202-5787596e2bd2"))
IActivePresenterChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ActivePresenter (
        /*[out,retval]*/ struct IParticipant * * _presenter ) = 0;
};

struct __declspec(uuid("899b846a-8887-45d6-b8f5-8aaa0bfdf722"))
IActiveContentChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_ActiveContent (
        /*[out,retval]*/ struct IShareableContent * * _content ) = 0;
};

struct __declspec(uuid("aafe380b-b079-4b55-b81b-a9401d77b776"))
IPowerPointContent : IShareableContent
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall GoForward ( ) = 0;
      virtual HRESULT __stdcall GoBackward ( ) = 0;
      virtual HRESULT __stdcall SyncWithPresenter ( ) = 0;
      virtual HRESULT __stdcall get_IsInSyncWithPresenter (
        /*[out,retval]*/ VARIANT_BOOL * _isInSync ) = 0;
};

struct __declspec(uuid("8a88c016-aa5d-4eae-9c79-7aa5c360ed76"))
IConversationContextCollection : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Count (
        /*[out,retval]*/ long * _count ) = 0;
      virtual HRESULT __stdcall get_Item (
        /*[in]*/ long _index,
        /*[out,retval]*/ struct IConversationContext * * _context ) = 0;
      virtual HRESULT __stdcall get__NewEnum (
        /*[out,retval]*/ IUnknown * * _variantEnumerator ) = 0;
};

struct __declspec(uuid("04378c72-e58f-47c0-8621-901b85ca2ed4"))
IConversationContext : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * _id ) = 0;
      virtual HRESULT __stdcall get_Type (
        /*[out,retval]*/ enum ConversationContextType * _type ) = 0;
      virtual HRESULT __stdcall get_IsActive (
        /*[out,retval]*/ VARIANT_BOOL * _isActive ) = 0;
      virtual HRESULT __stdcall get_IsRequired (
        /*[out,retval]*/ VARIANT_BOOL * _isRequired ) = 0;
      virtual HRESULT __stdcall put_IsRequired (
        /*[in]*/ VARIANT_BOOL _isRequired ) = 0;
      virtual HRESULT __stdcall get_From (
        /*[out,retval]*/ struct IParticipantInfo * * _from ) = 0;
      virtual HRESULT __stdcall put_From (
        /*[in]*/ struct IParticipantInfo * _from ) = 0;
      virtual HRESULT __stdcall get_To (
        /*[out,retval]*/ struct IParticipantInfo * * _to ) = 0;
      virtual HRESULT __stdcall put_To (
        /*[in]*/ struct IParticipantInfo * _to ) = 0;
      virtual HRESULT __stdcall get_ConversationId (
        /*[out,retval]*/ BSTR * _conversationId ) = 0;
      virtual HRESULT __stdcall put_ConversationId (
        /*[in]*/ BSTR _conversationId ) = 0;
      virtual HRESULT __stdcall get_Date (
        /*[out,retval]*/ BSTR * _date ) = 0;
      virtual HRESULT __stdcall put_Date (
        /*[in]*/ BSTR _date ) = 0;
      virtual HRESULT __stdcall get_Participants (
        /*[out,retval]*/ SAFEARRAY * * _participants ) = 0;
      virtual HRESULT __stdcall put_Participants (
        /*[in]*/ SAFEARRAY * _participants ) = 0;
      virtual HRESULT __stdcall get_ModalityTypes (
        /*[out,retval]*/ enum ModalityTypes * _modalityTypes ) = 0;
      virtual HRESULT __stdcall put_ModalityTypes (
        /*[in]*/ enum ModalityTypes _modalityTypes ) = 0;
      virtual HRESULT __stdcall get_DataFormat (
        /*[out,retval]*/ BSTR * _dataFormat ) = 0;
      virtual HRESULT __stdcall put_DataFormat (
        /*[in]*/ BSTR _dataFormat ) = 0;
      virtual HRESULT __stdcall get_Data (
        /*[out,retval]*/ BSTR * _data ) = 0;
      virtual HRESULT __stdcall put_Data (
        /*[in]*/ BSTR _data ) = 0;
      virtual HRESULT __stdcall get_SubContexts (
        /*[out,retval]*/ struct IConversationContextCollection * * _contexts ) = 0;
      virtual HRESULT __stdcall CreateParticipantInfo (
        /*[in]*/ BSTR _uri,
        /*[in]*/ BSTR _representedUri,
        /*[out,retval]*/ struct IParticipantInfo * * _participantInfo ) = 0;
};

struct __declspec(uuid("855ad949-e2eb-473d-aaf2-9800008ea303"))
IConversationContextList : IConversationContextCollection
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall Add (
        /*[in]*/ struct IConversationContext * _context ) = 0;
      virtual HRESULT __stdcall Remove (
        /*[in]*/ struct IConversationContext * _context ) = 0;
      virtual HRESULT __stdcall CreateContext (
        /*[in]*/ BSTR _id,
        /*[in]*/ enum ConversationContextType _contextType,
        /*[in]*/ VARIANT_BOOL _doMakeActive,
        /*[out,retval]*/ struct IConversationContext * * _context ) = 0;
      virtual HRESULT __stdcall ModifyActiveContext (
        /*[in]*/ BSTR _contextId,
        /*[in]*/ BSTR _contextData,
        /*[in]*/ enum ConversationContextUpdateOptions _contextUpdateOptions ) = 0;
};

struct __declspec(uuid("ef5c00e4-75e2-4d39-b358-0ed9cb14f3cb"))
IConversationContextCollectionEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Context (
        /*[out,retval]*/ struct IConversationContext * * _conversationContext ) = 0;
};

struct __declspec(uuid("dd181e97-eee0-4d65-bad1-16866d0c7953"))
IAccessPermission : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_AccessLevel (
        /*[out,retval]*/ enum AccessLevel * _accessLevel ) = 0;
      virtual HRESULT __stdcall get_AccessEntries (
        /*[out,retval]*/ SAFEARRAY * * _accessEntries ) = 0;
      virtual HRESULT __stdcall TryGetValue (
        /*[in]*/ enum AccessEntryScope _scope,
        /*[in]*/ BSTR _entryId,
        /*[out]*/ struct IAccessEntry * * _accessEntry,
        /*[out,retval]*/ VARIANT_BOOL * _found ) = 0;
      virtual HRESULT __stdcall GetAccessEntry (
        /*[in]*/ enum AccessEntryScope _scope,
        /*[in]*/ BSTR _entryId,
        /*[out]*/ struct IAccessEntry * * _accessEntry ) = 0;
};

struct __declspec(uuid("65b71572-5338-43c8-9e4c-f1dfc6711ad1"))
IAccessEntry : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_Scope (
        /*[out,retval]*/ enum AccessEntryScope * _scope ) = 0;
      virtual HRESULT __stdcall get_Id (
        /*[out,retval]*/ BSTR * _accessEntryId ) = 0;
      virtual HRESULT __stdcall get_Permission (
        /*[out,retval]*/ struct IAccessPermission * * _permission ) = 0;
      virtual HRESULT __stdcall SetPermission (
        /*[in]*/ struct IAccessPermission * _permission ) = 0;
      virtual HRESULT __stdcall ClearPermission ( ) = 0;
      virtual HRESULT __stdcall CanSetPermission (
        /*[in]*/ struct IAccessPermission * _permission,
        /*[out,retval]*/ VARIANT_BOOL * _canSet ) = 0;
      virtual HRESULT __stdcall CanClearPermission (
        /*[out,retval]*/ VARIANT_BOOL * _canClear ) = 0;
};

struct __declspec(uuid("dd31376e-e529-43f2-b3b4-e379ebd28e2a"))
IAccessEntryCollectionChangedEventData : IDispatch
{
    //
    // Raw methods provided by interface
    //

      virtual HRESULT __stdcall get_AccessEntry (
        /*[out,retval]*/ struct IAccessEntry * * _accessEntry ) = 0;
};

//
// Named GUID constants initializations
//

extern "C" const GUID __declspec(selectany) LIBID_UCCollaborationLib =
    {0xb9aa1f11,0xf480,0x4054,{0xa8,0x4e,0xb5,0xd9,0x27,0x7e,0x40,0xa8}};
extern "C" const GUID __declspec(selectany) IID_IAsynchronousOperation =
    {0xf7ee956b,0xa713,0x48d3,{0xaf,0xef,0x0a,0x28,0xf3,0x77,0x40,0xb3}};
extern "C" const GUID __declspec(selectany) IID_ILocaleString =
    {0xc12f0de5,0x9a7d,0x425c,{0xb3,0x91,0x8b,0xe0,0x04,0xea,0xa2,0xf6}};
extern "C" const GUID __declspec(selectany) IID_IAddress =
    {0xdc63d621,0x11c8,0x4490,{0xb0,0x17,0xb8,0xc5,0x7e,0xe1,0xbe,0x25}};
extern "C" const GUID __declspec(selectany) IID_INamedPropertyList =
    {0x1a8a9402,0xe89f,0x40ae,{0x88,0xa8,0xb3,0x28,0x39,0x16,0x84,0xb6}};
extern "C" const GUID __declspec(selectany) CLSID_AsynchronousOperation =
    {0xe6a9d57e,0x838c,0x43d1,{0x9f,0x12,0xe4,0x7c,0x84,0x45,0x31,0xe0}};
extern "C" const GUID __declspec(selectany) CLSID_LocaleString =
    {0x8cd4531b,0xd249,0x489c,{0xb5,0x5a,0x37,0xe4,0xa0,0xba,0xdb,0x25}};
extern "C" const GUID __declspec(selectany) CLSID_Address =
    {0x99880cd9,0x73b0,0x4b84,{0x93,0xff,0x95,0xbc,0xd1,0x8c,0x9e,0xf8}};
extern "C" const GUID __declspec(selectany) CLSID_NamedPropertyList =
    {0x66b28b84,0x9d92,0x47ff,{0x91,0x52,0xe8,0x9e,0xaf,0xba,0xf3,0xc2}};
extern "C" const GUID __declspec(selectany) IID_IContactInformationDictionary =
    {0xf0ce810f,0xe36f,0x4cd8,{0xa8,0x36,0x51,0x5b,0x9a,0x6e,0x09,0xfc}};
extern "C" const GUID __declspec(selectany) IID_IContactSettingDictionary =
    {0xfd9000b3,0x479f,0x4b16,{0x9d,0x63,0x70,0xa4,0x9e,0x07,0x89,0x46}};
extern "C" const GUID __declspec(selectany) IID_IContactEndpoint =
    {0x1ff0e6f1,0x2fe4,0x4e29,{0xa1,0x23,0x55,0x7a,0xf0,0xdb,0x69,0x27}};
extern "C" const GUID __declspec(selectany) IID_IContactInformationChangedEventData =
    {0xefec2816,0xf16d,0x48d8,{0x93,0x06,0x26,0xc8,0x10,0xf0,0xea,0x55}};
extern "C" const GUID __declspec(selectany) IID_IDistributionGroupExpansionStateChangedEventData =
    {0xd2bfc7d7,0xa64f,0x48ef,{0xbf,0x90,0xc3,0xc7,0xe6,0x84,0xd8,0x9c}};
extern "C" const GUID __declspec(selectany) IID_IUriChangedEventData =
    {0x0bdb9057,0x28ae,0x4bf0,{0xaf,0xf0,0x12,0xa1,0x48,0xe5,0x16,0x37}};
extern "C" const GUID __declspec(selectany) IID_IGroupNameChangedEventData =
    {0x283c2089,0xb760,0x4d65,{0x91,0x99,0x71,0x66,0x27,0x17,0x4f,0x7a}};
extern "C" const GUID __declspec(selectany) IID_ISearchResult =
    {0x378cd767,0x1b64,0x4a0e,{0xa9,0x41,0x8f,0x52,0x86,0x9f,0xc9,0xdf}};
extern "C" const GUID __declspec(selectany) IID_ISearchResults =
    {0x7356d7bb,0xfd71,0x4554,{0x84,0xa1,0x3b,0xbe,0x28,0x72,0x65,0x51}};
extern "C" const GUID __declspec(selectany) IID_ISearchProviderStateChangedEventData =
    {0x995f992c,0x9def,0x47b9,{0xbf,0x11,0x81,0x81,0x3c,0x0c,0x0e,0x28}};
extern "C" const GUID __declspec(selectany) IID_IWorkingPeriod =
    {0x339d28c6,0x3d1a,0x45d4,{0xba,0x14,0xa5,0x67,0x42,0x74,0x9a,0x43}};
extern "C" const GUID __declspec(selectany) IID_IWorkingHours =
    {0x53d014c1,0x54db,0x42b3,{0x9d,0xfd,0x8e,0x23,0x1e,0xf2,0xc3,0x56}};
extern "C" const GUID __declspec(selectany) IID_IPresenceCapability =
    {0x10c3acc7,0x6c0c,0x4f35,{0xad,0xfb,0xae,0x0f,0x20,0x0a,0x85,0x83}};
extern "C" const GUID __declspec(selectany) IID_IDescriptionChangedEventData =
    {0x8affb8ce,0x5404,0x4280,{0xbc,0xa3,0xe2,0xe4,0x38,0x8e,0x6d,0x73}};
extern "C" const GUID __declspec(selectany) IID_ISearchResults2 =
    {0x7cd92461,0xca4f,0x4beb,{0xa6,0x36,0x8b,0xbc,0xc1,0x92,0xe6,0x0f}};
extern "C" const GUID __declspec(selectany) DIID__IContactsAndGroupsCallback =
    {0xabf012a2,0x2068,0x4b4d,{0x80,0xd7,0x6d,0x38,0x6b,0xa7,0xc8,0xa2}};
extern "C" const GUID __declspec(selectany) DIID__IGroupCallback =
    {0xc6605772,0x3d9d,0x44a7,{0xb4,0x3a,0x45,0x90,0xe6,0xb4,0x3e,0xe3}};
extern "C" const GUID __declspec(selectany) DIID__IContactCallback =
    {0x311dde48,0xed7d,0x46fa,{0x9e,0x0a,0x2e,0x31,0x4b,0x7f,0xef,0x7c}};
extern "C" const GUID __declspec(selectany) DIID__IContactManagerEvents =
    {0x50d86ee2,0x288f,0x44f5,{0x81,0x44,0x69,0xf6,0xe3,0xb2,0x4b,0x90}};
extern "C" const GUID __declspec(selectany) DIID__IContactEvents =
    {0xb1c6ecc3,0x1ea7,0x495d,{0x93,0x37,0x02,0xfe,0x4f,0xb7,0x23,0xf0}};
extern "C" const GUID __declspec(selectany) IID_IContactSettingChangedEventData =
    {0x12480bd4,0x072f,0x4763,{0xb9,0xfb,0x41,0xb2,0xca,0x54,0xf9,0xcd}};
extern "C" const GUID __declspec(selectany) DIID__IGroupEvents =
    {0x9918673d,0x4d74,0x41cf,{0x88,0x40,0xd4,0x99,0x63,0x95,0xa3,0x4c}};
extern "C" const GUID __declspec(selectany) DIID__ICustomGroupEvents =
    {0xf719ed75,0xb7bd,0x45d3,{0x90,0x97,0x25,0x45,0x09,0x22,0x6f,0x20}};
extern "C" const GUID __declspec(selectany) DIID__IRecentContactsEvents =
    {0x011b320c,0x9ac1,0x4dde,{0xb5,0x95,0x3d,0x00,0x14,0x47,0x2e,0x2f}};
extern "C" const GUID __declspec(selectany) DIID__IDistributionGroupEvents =
    {0x082ee280,0x726e,0x417f,{0x99,0xcb,0x81,0xa0,0xcc,0xbf,0xf8,0x83}};
extern "C" const GUID __declspec(selectany) DIID__IFrequentContactsEvents =
    {0xaa186c92,0x181e,0x417f,{0xb1,0x50,0xfc,0xa0,0xf3,0x67,0xe0,0xfc}};
extern "C" const GUID __declspec(selectany) IID_IContactPositionChangedEventData =
    {0x62b932e9,0x7847,0x4d78,{0xbe,0xee,0x5d,0x58,0xbf,0x81,0x5b,0x03}};
extern "C" const GUID __declspec(selectany) DIID__IFavoriteContactsEvents =
    {0xdab42e13,0x500a,0x409a,{0xac,0x34,0xc1,0xbf,0x47,0x19,0xee,0x2e}};
extern "C" const GUID __declspec(selectany) CLSID_ContactManager =
    {0x43eaf09a,0x5dd2,0x4ce7,{0x9b,0x03,0xfc,0xee,0xd1,0x87,0xbf,0x48}};
extern "C" const GUID __declspec(selectany) CLSID_Contact =
    {0xd3ef7c83,0x13f3,0x4bf2,{0x82,0x16,0x54,0x0f,0xb9,0x5b,0xbd,0x00}};
extern "C" const GUID __declspec(selectany) CLSID_Group =
    {0x95ee29d5,0x2b51,0x46e4,{0x8f,0x7c,0x40,0x6f,0xc3,0x82,0x7a,0x80}};
extern "C" const GUID __declspec(selectany) CLSID_ContactCollection =
    {0x9bb71fc6,0x1bba,0x407a,{0xa0,0x86,0x87,0xd8,0xe5,0x25,0xbf,0xef}};
extern "C" const GUID __declspec(selectany) CLSID_CustomGroup =
    {0xa2e0476a,0x945f,0x4c93,{0x98,0xfa,0x87,0xae,0xd2,0x9e,0x94,0x67}};
extern "C" const GUID __declspec(selectany) CLSID_DistributionGroup =
    {0xcaa28f4e,0x98dd,0x4822,{0x84,0x62,0x95,0x5a,0xad,0xdf,0x23,0x64}};
extern "C" const GUID __declspec(selectany) CLSID_RecentContacts =
    {0xe6162417,0x23be,0x437b,{0xb9,0xd2,0x7b,0xf6,0xca,0xc5,0xe8,0x29}};
extern "C" const GUID __declspec(selectany) CLSID_FrequentContacts =
    {0xb395d986,0x2542,0x4c59,{0x88,0x3e,0x66,0x12,0xa3,0x4e,0x6d,0x75}};
extern "C" const GUID __declspec(selectany) CLSID_FavoriteContacts =
    {0xc66de384,0x1f97,0x4f1a,{0xaf,0x30,0x00,0x05,0x99,0x00,0x00,0xf5}};
extern "C" const GUID __declspec(selectany) CLSID_GroupMemberChangedEventData =
    {0xbbc00952,0x0dc4,0x4abc,{0x89,0xf7,0x53,0x36,0xc5,0x25,0xb8,0xae}};
extern "C" const GUID __declspec(selectany) CLSID_DistributionGroupExpansionStateChangedEventData =
    {0xdd734d35,0x7a6b,0x42d8,{0xbe,0x0b,0xaf,0x68,0xc7,0x6a,0x48,0x7a}};
extern "C" const GUID __declspec(selectany) CLSID_UriChangedEventData =
    {0xb89f18d4,0x5041,0x488a,{0x8d,0x81,0xde,0x8e,0xbb,0x3a,0x36,0x2c}};
extern "C" const GUID __declspec(selectany) CLSID_GroupNameChangedEventData =
    {0x4eb781a9,0xa304,0x4bc5,{0xa3,0x99,0x2c,0x79,0xb2,0x0d,0x40,0x79}};
extern "C" const GUID __declspec(selectany) CLSID_GroupCollectionChangedEventData =
    {0xab12b906,0xde1a,0x4fb7,{0x8d,0x1a,0x77,0x4a,0x04,0xc7,0xf3,0xdc}};
extern "C" const GUID __declspec(selectany) CLSID_SearchResult =
    {0xfacab35b,0xefe6,0x4228,{0x82,0xdb,0x5a,0xaa,0xbc,0x2d,0xb2,0x39}};
extern "C" const GUID __declspec(selectany) CLSID_SearchResults =
    {0xe73ddb55,0xacb3,0x4f33,{0xa0,0xed,0x7b,0x22,0xd6,0x8b,0xf3,0x11}};
extern "C" const GUID __declspec(selectany) CLSID_SearchProviderStateChangedEventData =
    {0xb53c508e,0xe967,0x43d2,{0xa5,0xc6,0x85,0x52,0x38,0x59,0x58,0x1f}};
extern "C" const GUID __declspec(selectany) CLSID_GroupCollection =
    {0xda937c78,0x8d00,0x41ec,{0x87,0x4f,0x15,0x6d,0x50,0xee,0x97,0xec}};
extern "C" const GUID __declspec(selectany) CLSID_ContactInformationChangedEventData =
    {0x9c49c870,0x9104,0x41f6,{0xae,0x37,0x10,0x94,0x34,0xf8,0xba,0x8d}};
extern "C" const GUID __declspec(selectany) CLSID_ContactSettingChangedEventData =
    {0x1496ae1a,0x5ba4,0x484d,{0x8b,0x2b,0x57,0xc4,0x4f,0xe8,0x10,0xcd}};
extern "C" const GUID __declspec(selectany) CLSID_ContactPositionChangedEventData =
    {0x9367ca1b,0x8317,0x448a,{0xbf,0x18,0x8d,0x88,0xd8,0x09,0xd4,0xdf}};
extern "C" const GUID __declspec(selectany) CLSID_WorkingHours =
    {0xff235ffe,0xedef,0x4d1b,{0x92,0x05,0x51,0x98,0x88,0xa7,0xad,0x8d}};
extern "C" const GUID __declspec(selectany) CLSID_ContactInformationDictionary =
    {0x0858ab97,0x0e11,0x49bd,{0x8f,0x88,0x4e,0x2a,0xe0,0x23,0x2e,0x33}};
extern "C" const GUID __declspec(selectany) CLSID_ContactSettingDictionary =
    {0x60a8f660,0xc321,0x4761,{0x8a,0xff,0x41,0x07,0x2d,0xcf,0xdd,0x03}};
extern "C" const GUID __declspec(selectany) CLSID_WorkingPeriod =
    {0xb482e1e6,0xafb7,0x4e58,{0x91,0xee,0x3e,0x20,0xf6,0x13,0x92,0x20}};
extern "C" const GUID __declspec(selectany) CLSID_ContactSubscription =
    {0xfeaeffd6,0xc800,0x46a0,{0xb6,0x34,0xfd,0xc4,0x15,0xfe,0xc2,0x11}};
extern "C" const GUID __declspec(selectany) CLSID_PresenceCapability =
    {0x8d418098,0x61cd,0x4970,{0xa0,0xe0,0x2e,0x9b,0xd6,0x42,0x09,0x76}};
extern "C" const GUID __declspec(selectany) CLSID_ContactEndpoint =
    {0xd2f90870,0x11a2,0x4396,{0x80,0x6b,0xd7,0x9d,0xdf,0xbe,0x9f,0x50}};
extern "C" const GUID __declspec(selectany) CLSID_DescriptionChangedEventData =
    {0x9bf7a83a,0x92f2,0x4c1d,{0x91,0xb6,0x7b,0xe0,0x48,0xef,0xd8,0xa5}};
extern "C" const GUID __declspec(selectany) IID_IModalityPropertyDictionary =
    {0x62074904,0x8d06,0x43fe,{0xa5,0x31,0xe6,0x3d,0xf7,0xfd,0xc2,0xe7}};
extern "C" const GUID __declspec(selectany) IID_IParticipantPropertyDictionary =
    {0x8665fd95,0x7720,0x4f9e,{0xb6,0x05,0x6a,0xbc,0xbd,0x7e,0xdf,0xf4}};
extern "C" const GUID __declspec(selectany) IID_IParticipantInfo =
    {0xb3ff95bd,0xb743,0x498c,{0x99,0x73,0x7f,0x8a,0x2a,0x78,0x57,0x6a}};
extern "C" const GUID __declspec(selectany) IID_IConversationPropertyDictionary =
    {0x00e22cbb,0x3170,0x453a,{0xae,0x62,0xea,0xfb,0xc7,0x5a,0x9f,0x8d}};
extern "C" const GUID __declspec(selectany) IID_IVoiceMailsManager =
    {0xe24c2484,0xc70d,0x44ab,{0x8a,0x4e,0xf1,0x58,0x5a,0x17,0xa3,0xe1}};
extern "C" const GUID __declspec(selectany) IID_IPreviousConversationsManager =
    {0xacabfa1c,0x0738,0x4c4d,{0xaf,0x7e,0x32,0x68,0x73,0xcf,0x4a,0xce}};
extern "C" const GUID __declspec(selectany) IID_IPreviousConversationsManagerDictionary =
    {0x15d4461d,0xc7eb,0x45cb,{0xb2,0x14,0xf1,0xc4,0x5f,0xed,0x1d,0xb2}};
extern "C" const GUID __declspec(selectany) IID_IRepresentationInfo =
    {0x209da899,0x15f1,0x499a,{0xb8,0xb5,0xa8,0x47,0xea,0xa8,0x99,0xae}};
extern "C" const GUID __declspec(selectany) IID_IVideoWindow =
    {0x56a868b4,0x0ad4,0x11ce,{0xb0,0x3a,0x00,0x20,0xaf,0x0b,0xa7,0x70}};
extern "C" const GUID __declspec(selectany) IID_ISharingResource =
    {0x3690d374,0x3779,0x4efc,{0xb8,0xe8,0x20,0x64,0x2c,0x01,0x52,0x9a}};
extern "C" const GUID __declspec(selectany) IID_ISharingResourceList =
    {0x6a7ff112,0x42ce,0x4d03,{0x89,0x22,0xd0,0x2a,0xc6,0xc5,0x75,0x9f}};
extern "C" const GUID __declspec(selectany) IID_IChannel2 =
    {0x7b39cb77,0x4aeb,0x42d7,{0xb3,0x51,0xcb,0x54,0x72,0xc1,0xc6,0xad}};
extern "C" const GUID __declspec(selectany) IID_IModalityStateChangePropertyDictionary =
    {0xc79e4169,0x535f,0x4a41,{0x8d,0xe2,0xc6,0x5a,0x90,0xd9,0x50,0x3d}};
extern "C" const GUID __declspec(selectany) IID_IModalityStateChangedEventData =
    {0x478dda59,0x64c0,0x4429,{0xac,0x8a,0xf1,0x0a,0x94,0x0b,0x52,0xcc}};
extern "C" const GUID __declspec(selectany) IID_IModalityActionAvailabilityChangedEventData =
    {0xdeddd7bd,0x4763,0x41d2,{0x9a,0xaa,0xb2,0xc1,0x43,0x45,0x7c,0xa4}};
extern "C" const GUID __declspec(selectany) IID_IConversationStateChangePropertyDictionary =
    {0x56f7dec0,0x59ca,0x47c6,{0x9f,0x35,0xd5,0x06,0x6a,0x70,0x2b,0x39}};
extern "C" const GUID __declspec(selectany) IID_IConversationStateChangedEventData =
    {0x2a9385aa,0xcf61,0x4e47,{0xb6,0x4c,0x21,0x4d,0xe4,0xa4,0xad,0x11}};
extern "C" const GUID __declspec(selectany) IID_IConversationActionAvailabilityEventData =
    {0x88c688b0,0x3908,0x4c56,{0xa2,0xe8,0xf9,0x0a,0xb7,0x05,0xc5,0x36}};
extern "C" const GUID __declspec(selectany) IID_IMessageSentEventData =
    {0x905b727d,0xfd8f,0x46ec,{0x80,0x06,0xce,0x24,0x5f,0x68,0x35,0xee}};
extern "C" const GUID __declspec(selectany) IID_IIsTypingChangedEventData =
    {0xdba05c15,0x1c07,0x4a76,{0x82,0x48,0x08,0xd8,0x41,0x6a,0x24,0xe3}};
extern "C" const GUID __declspec(selectany) IID_IIsContributingChangedEventData =
    {0xd5c1c8fb,0xdc8a,0x45a3,{0xaa,0xa3,0x53,0xb8,0x9d,0x43,0xcd,0xe4}};
extern "C" const GUID __declspec(selectany) IID_IChannelStateChangedEventData =
    {0x2495b94c,0x38af,0x439b,{0xbb,0xa0,0x0a,0xd3,0x8d,0x95,0x9c,0xe4}};
extern "C" const GUID __declspec(selectany) IID_IChannelActionAvailabilityEventData =
    {0x9d4d3d10,0xdebc,0x40c6,{0xac,0xf9,0x52,0x79,0x86,0xcb,0x42,0xbc}};
extern "C" const GUID __declspec(selectany) IID_IHistoryInfo =
    {0x3debdc7e,0x811f,0x4304,{0xbd,0xd7,0x7d,0x98,0xc7,0x32,0x02,0x6c}};
extern "C" const GUID __declspec(selectany) IID_IConversationPropertyChangedEventData =
    {0x7cfe77cd,0x731d,0x48b2,{0x82,0xb1,0xec,0xa3,0x41,0x4d,0x62,0xe3}};
extern "C" const GUID __declspec(selectany) IID_IParticipantPropertyChangedEventData =
    {0x70dcce38,0x90bf,0x4428,{0xa3,0x4f,0x3a,0x60,0x82,0xf2,0x9e,0x50}};
extern "C" const GUID __declspec(selectany) IID_IModalityPropertyChangedEventData =
    {0x4290152c,0x75c9,0x47c8,{0xaa,0x12,0x81,0x48,0x18,0x36,0x0b,0x3f}};
extern "C" const GUID __declspec(selectany) IID_IPreviousConversation =
    {0xdf414a68,0x5051,0x4465,{0xaa,0xe2,0x4f,0x30,0x13,0x15,0x73,0x4e}};
extern "C" const GUID __declspec(selectany) IID_IPreviousConversationCollection =
    {0x81c9d13f,0xa4f9,0x4e13,{0x92,0xd3,0xbb,0x27,0x1e,0x8d,0xf3,0xd2}};
extern "C" const GUID __declspec(selectany) IID_IPreviousConversationsManagerConnectionStateChangedEventData =
    {0xf87afd9a,0x9bf8,0x4ac4,{0x95,0x3e,0x33,0xc4,0xd2,0x03,0x5d,0x33}};
extern "C" const GUID __declspec(selectany) IID_IPreviousConversationsManagerNewItemCountChangedEventData =
    {0x2c59a886,0xacd3,0x46bb,{0x8b,0xac,0x7c,0x1c,0x59,0xfb,0x2e,0x0d}};
extern "C" const GUID __declspec(selectany) IID_IPreviousConversationBatchEventData =
    {0xebd19735,0x0cb5,0x4af3,{0x8e,0x05,0xb0,0x14,0x0c,0xf9,0xd5,0x45}};
extern "C" const GUID __declspec(selectany) IID_IConferenceAccessInformation =
    {0x2f8c3e58,0x436b,0x42db,{0x89,0x24,0x6c,0x39,0x4b,0x37,0xdc,0xa2}};
extern "C" const GUID __declspec(selectany) IID_IParticipantActionAvailabilityChangedEventData =
    {0x1124aa59,0xb6f3,0x4eb1,{0xb7,0x9a,0xe7,0xf5,0x3b,0x38,0x8b,0x8d}};
extern "C" const GUID __declspec(selectany) IID_IConversation2 =
    {0x37c7f935,0xe215,0x46d2,{0x82,0xa3,0xa4,0x51,0x84,0x0e,0xc7,0x33}};
extern "C" const GUID __declspec(selectany) IID_IParticipant2 =
    {0xd60b1701,0x766d,0x401e,{0x85,0x86,0x83,0xe0,0xc9,0x10,0x6b,0xe0}};
extern "C" const GUID __declspec(selectany) DIID__IConversationManagerEvents =
    {0x25025b1b,0x0084,0x44d1,{0xb3,0x83,0x9f,0xfc,0x9a,0x99,0xab,0xcd}};
extern "C" const GUID __declspec(selectany) DIID__IConversationEvents =
    {0xc63f58cb,0xef11,0x4934,{0xaa,0x3b,0x2d,0x69,0x24,0x4f,0x07,0x80}};
extern "C" const GUID __declspec(selectany) IID_IInitialContextEventData =
    {0x92d7e7a8,0x48d8,0x4166,{0x89,0x11,0x63,0x0a,0xe0,0x2b,0x2b,0x93}};
extern "C" const GUID __declspec(selectany) IID_IContextEventData =
    {0x21629ab6,0x9557,0x4985,{0x92,0x37,0x49,0x17,0x7f,0x61,0x86,0x92}};
extern "C" const GUID __declspec(selectany) DIID__IParticipantEvents =
    {0x6a986b5d,0xe5ec,0x4395,{0xad,0x2d,0x0e,0x62,0xb7,0x2d,0xa7,0x76}};
extern "C" const GUID __declspec(selectany) IID_IMutedChangedEventData =
    {0x5418e2ab,0xeb9a,0x4659,{0xb4,0xdc,0x28,0xde,0x63,0x3b,0x2b,0x8f}};
extern "C" const GUID __declspec(selectany) DIID__IModalityEvents =
    {0x10d66b39,0x5e4b,0x4c98,{0xaa,0xc6,0xd0,0x45,0x8f,0xeb,0x1b,0xdb}};
extern "C" const GUID __declspec(selectany) DIID__IInstantMessageModalityEvents =
    {0x4fe44049,0x4e44,0x4109,{0xb2,0x34,0x4e,0x4e,0xfc,0x13,0x5a,0x86}};
extern "C" const GUID __declspec(selectany) DIID__IAVModalityEvents =
    {0xacd2fdb9,0x98ed,0x4283,{0xaa,0xe8,0x3c,0xd9,0x89,0xeb,0xd9,0xae}};
extern "C" const GUID __declspec(selectany) DIID__IApplicationSharingModalityEvents =
    {0xbe026cd2,0x7e82,0x4f7c,{0x87,0x62,0xf6,0xb0,0x2f,0x49,0x61,0x74}};
extern "C" const GUID __declspec(selectany) IID_ILocalSharedResourcesChangedEventData =
    {0xb22edbea,0x9e61,0x4703,{0x82,0xbe,0x01,0xc0,0x56,0x19,0xb6,0xd6}};
extern "C" const GUID __declspec(selectany) IID_IParticipationStateChangedEventData =
    {0x8fe9a23d,0x72bc,0x4ea3,{0xa5,0xaf,0xe3,0x7f,0x75,0xea,0xd8,0xa2}};
extern "C" const GUID __declspec(selectany) IID_IControllerChangedEventData =
    {0x4e8dc7e0,0x04b8,0x470b,{0xbd,0xfa,0xf5,0x20,0x09,0x9b,0x97,0x5f}};
extern "C" const GUID __declspec(selectany) IID_ISharerChangedEventData =
    {0xd0dbb17d,0x565e,0x486c,{0xa4,0x7b,0xba,0x32,0xdc,0x1f,0xad,0x0a}};
extern "C" const GUID __declspec(selectany) DIID__IConversationHistoryEvents =
    {0xfa7971ef,0x1e41,0x429f,{0x8c,0x79,0xda,0xc5,0xf8,0x8d,0x44,0x4b}};
extern "C" const GUID __declspec(selectany) IID_IConversationHistoryItemAddedEventData =
    {0x4c675532,0x6d55,0x49dd,{0x8c,0x8e,0x3c,0x1c,0xc5,0x62,0x4c,0xf2}};
extern "C" const GUID __declspec(selectany) IID_IPreviousHistoryRetrievalStateChangedEventData =
    {0x464d5228,0x9f68,0x4b1f,{0xb4,0x30,0x15,0x6a,0x10,0x4e,0x2b,0x85}};
extern "C" const GUID __declspec(selectany) IID_IConversationHistoryActionAvailabilityEventData =
    {0x9397ff55,0xee06,0x4f02,{0x8f,0x2a,0xbe,0x3a,0xe2,0x49,0xd4,0xbb}};
extern "C" const GUID __declspec(selectany) DIID__IPreviousConversationsManagerEvents =
    {0xd992371e,0x5161,0x453b,{0x97,0xe6,0x6e,0x7c,0x67,0xbc,0x07,0x5e}};
extern "C" const GUID __declspec(selectany) DIID__IVoiceMailsManagerEvents =
    {0xb0c10f94,0xcab6,0x4821,{0x96,0x43,0xd7,0x81,0x88,0x5a,0x46,0xaf}};
extern "C" const GUID __declspec(selectany) DIID__IChannelEvents =
    {0x0ca9bae8,0xdd28,0x4929,{0xa8,0x21,0x4d,0x00,0xf1,0x12,0x03,0xcc}};
extern "C" const GUID __declspec(selectany) DIID__IConversationCallback =
    {0x57cc2235,0x96c5,0x49fa,{0xb9,0x2b,0x35,0x04,0x86,0xc6,0xcf,0x52}};
extern "C" const GUID __declspec(selectany) DIID__IParticipantCallback =
    {0x26c8d96d,0x7d22,0x4e9e,{0x94,0x8a,0xed,0xcc,0xb4,0xca,0x9c,0x64}};
extern "C" const GUID __declspec(selectany) DIID__IModalityCallback =
    {0x3f60cff7,0xbe3f,0x4404,{0x83,0x95,0xc0,0x2d,0x0c,0xdd,0x31,0x8e}};
extern "C" const GUID __declspec(selectany) DIID__IApplicationSharingModalityCallback =
    {0x8684d3ad,0x7061,0x4056,{0xa8,0x94,0xeb,0xf1,0x27,0x0a,0xdb,0x9b}};
extern "C" const GUID __declspec(selectany) DIID__IConversationHistoryCallback =
    {0x5fa92ea7,0x6e6e,0x4a82,{0x8f,0x0d,0x10,0x7f,0xea,0xf5,0xa7,0x5d}};
extern "C" const GUID __declspec(selectany) DIID__IAudioChannelCallback =
    {0x95b9a382,0xea50,0x45e4,{0x96,0x5e,0xba,0x61,0xe7,0x21,0x3d,0xba}};
extern "C" const GUID __declspec(selectany) DIID__IChannelCallback =
    {0x08e1e9cf,0xa7e4,0x43be,{0xae,0x8e,0xb2,0x23,0xce,0x71,0xe4,0x6e}};
extern "C" const GUID __declspec(selectany) DIID__IVoiceMailCallback =
    {0x68b1fc26,0xb48f,0x4d5f,{0xb0,0x41,0x8e,0x24,0x66,0x76,0xd8,0x36}};
extern "C" const GUID __declspec(selectany) DIID__IVoiceMailsManagerCallback =
    {0xd9d04b55,0xd820,0x4d5b,{0xa6,0x90,0x65,0x8a,0x49,0x36,0x84,0x78}};
extern "C" const GUID __declspec(selectany) DIID__IPreviousConversationsManagerCallback =
    {0x4bacb4b8,0x1236,0x42b8,{0xbd,0xa1,0xd1,0x53,0x31,0x48,0x78,0x0d}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationManager =
    {0xc87287d3,0x0a25,0x414d,{0x9a,0x92,0x16,0xa1,0x1a,0x20,0x99,0xf0}};
extern "C" const GUID __declspec(selectany) CLSID_Conversation =
    {0x457d962f,0x5f66,0x4429,{0xbd,0x94,0x6f,0x9f,0x36,0x6f,0xae,0x0d}};
extern "C" const GUID __declspec(selectany) CLSID_Participant =
    {0xde5598bc,0x7188,0x415a,{0xa3,0xd9,0x7e,0x86,0x86,0x3e,0x7a,0x09}};
extern "C" const GUID __declspec(selectany) CLSID_InstantMessageModality =
    {0x9582d285,0x7e6e,0x4bfa,{0xbc,0xbc,0x7d,0xe3,0xea,0x3e,0x6d,0x74}};
extern "C" const GUID __declspec(selectany) CLSID_AVModality =
    {0xba2bd6f3,0x7676,0x42e3,{0x89,0xc6,0x10,0xce,0xb3,0xf7,0xe1,0x06}};
extern "C" const GUID __declspec(selectany) CLSID_Channel =
    {0x556a0d7a,0x05fb,0x4a8b,{0xb1,0x5a,0xad,0x3e,0xac,0xde,0x3a,0xd2}};
extern "C" const GUID __declspec(selectany) CLSID_Modality =
    {0x61419951,0x5d5a,0x4da3,{0xbe,0x6a,0x07,0x6f,0x96,0x45,0x22,0x9e}};
extern "C" const GUID __declspec(selectany) CLSID_ApplicationSharingModality =
    {0xf8f4a9b9,0x7e52,0x44e0,{0x89,0xe0,0x30,0xa3,0xb4,0xda,0x4d,0x23}};
extern "C" const GUID __declspec(selectany) CLSID_AudioChannel =
    {0xd5975289,0x8d98,0x481a,{0x92,0xa7,0xf8,0x53,0x60,0x85,0xd6,0x21}};
extern "C" const GUID __declspec(selectany) CLSID_IsTypingChangedEventData =
    {0x56805856,0xacad,0x4d8c,{0xae,0xaf,0x55,0x76,0x40,0x4b,0xf7,0x2b}};
extern "C" const GUID __declspec(selectany) CLSID_IsContributingChangedEventData =
    {0x0af70da9,0x90e5,0x4128,{0x8d,0x67,0xa0,0x02,0xcc,0x94,0xc5,0x3b}};
extern "C" const GUID __declspec(selectany) CLSID_ContextEventData =
    {0x288926af,0x3015,0x4da0,{0xb4,0x1c,0x93,0x81,0xe5,0x45,0xee,0xcf}};
extern "C" const GUID __declspec(selectany) CLSID_InitialContextEventData =
    {0xcf6796fe,0xc1d0,0x400f,{0x85,0xc1,0x3e,0x8a,0x0f,0x5c,0x24,0x7c}};
extern "C" const GUID __declspec(selectany) CLSID_ChannelStateChangedEventData =
    {0x00068095,0xfa0a,0x4142,{0x9f,0x08,0xe0,0xe5,0xa9,0x59,0x96,0x21}};
extern "C" const GUID __declspec(selectany) CLSID_ChannelActionAvailabilityEventData =
    {0x61e9f245,0xce9a,0x4403,{0xaa,0xd9,0x01,0x2b,0x52,0xce,0x15,0x1d}};
extern "C" const GUID __declspec(selectany) CLSID_StreamStateChangedEventData =
    {0x6ec2948c,0x51d0,0x4806,{0xba,0xf1,0x21,0x8e,0xec,0xb2,0x0d,0x08}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationActionAvailabilityEventData =
    {0xf8946ac5,0x5eb8,0x4237,{0xa1,0xf5,0x1b,0x03,0xc6,0xea,0xaa,0xc1}};
extern "C" const GUID __declspec(selectany) CLSID_ModalityActionAvailabilityChangedEventData =
    {0x36d057b0,0x0f7a,0x4c96,{0x9e,0x99,0x28,0x50,0x73,0xd5,0x0c,0xc7}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationContextCollectionEventData =
    {0x2ede3b04,0x2f52,0x4ef0,{0xb7,0x93,0x77,0x39,0x7b,0x64,0x18,0xde}};
extern "C" const GUID __declspec(selectany) CLSID_SummarizerStatesChangedEventData =
    {0xc42e7aba,0xdea0,0x43ec,{0x96,0xaa,0x6d,0x82,0x11,0xcb,0x2e,0x14}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationManagerEventData =
    {0xcc1991bf,0xa8f3,0x48cb,{0x94,0xc2,0xed,0xe7,0xaa,0x55,0x45,0xf3}};
extern "C" const GUID __declspec(selectany) CLSID_ParticipantCollectionChangedEventData =
    {0x2d9dd8cd,0x381f,0x42b9,{0xaa,0x02,0x65,0xdf,0x5e,0x7d,0x1f,0xb1}};
extern "C" const GUID __declspec(selectany) CLSID_MessageSentEventData =
    {0xee1abe27,0x10e4,0x4a70,{0xb4,0xdd,0x9a,0x55,0x14,0x9f,0xd4,0x86}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationStateChangedEventData =
    {0x45076174,0x6cbd,0x471c,{0xa8,0xa7,0x19,0x39,0x1a,0x2c,0x25,0x2c}};
extern "C" const GUID __declspec(selectany) CLSID_ModalityStateChangedEventData =
    {0x458858d8,0x9ac1,0x4431,{0xb5,0x16,0xe4,0x43,0x62,0xfa,0xc7,0xad}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationPropertyChangedEventData =
    {0x0d2b5648,0xd0cf,0x4069,{0xb5,0x05,0xcd,0x12,0x13,0x19,0xfd,0xd7}};
extern "C" const GUID __declspec(selectany) CLSID_ParticipantPropertyChangedEventData =
    {0xfe8d6343,0x3bdf,0x4ae7,{0xa2,0xbc,0x70,0x2f,0x5f,0x33,0x11,0x18}};
extern "C" const GUID __declspec(selectany) CLSID_MutedChangedEventData =
    {0xcc671df2,0xecb2,0x4cf2,{0x87,0x6e,0xac,0x14,0xc2,0x5a,0xa4,0xfb}};
extern "C" const GUID __declspec(selectany) CLSID_ParticipantActionAvailabilityChangedEventData =
    {0x30e3bdff,0x0457,0x4695,{0xa7,0xa5,0xc6,0x3c,0x36,0xdb,0x06,0x9e}};
extern "C" const GUID __declspec(selectany) CLSID_ModalityPropertyChangedEventData =
    {0x91d5058a,0x77c3,0x40da,{0x93,0x0c,0xf6,0x20,0x7b,0xd6,0x91,0xc3}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationCollection =
    {0x373b822d,0xc7ee,0x4763,{0x93,0xa7,0x86,0xbb,0x2e,0xa7,0xe0,0x31}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationContext =
    {0x3afa5d5f,0xc3d5,0x453f,{0x96,0xe0,0x30,0xc3,0xdb,0xc3,0x49,0x51}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationContextList =
    {0x931a066a,0xaf82,0x4278,{0x90,0x84,0x4f,0x8b,0x51,0x7f,0x91,0x05}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationHistory =
    {0x00b1f05c,0x924f,0x4a74,{0xae,0xd2,0x8e,0xde,0x3d,0x08,0xde,0x8b}};
extern "C" const GUID __declspec(selectany) CLSID_ParticipantCollection =
    {0xbee2a581,0x7f14,0x4c61,{0x89,0xa6,0x47,0x43,0x0e,0xc1,0x38,0xea}};
extern "C" const GUID __declspec(selectany) CLSID_VideoChannel =
    {0x90b18225,0x0bc7,0x4120,{0x84,0xaf,0x20,0x05,0x6f,0x7e,0x7d,0xa4}};
extern "C" const GUID __declspec(selectany) CLSID_VideoChannelCollection =
    {0x67c2d762,0x4e8d,0x45bf,{0x94,0xf0,0x68,0xd9,0x91,0x03,0x56,0x7a}};
extern "C" const GUID __declspec(selectany) CLSID_VideoWindow =
    {0x6c97f743,0x778b,0x460f,{0xb1,0x4e,0x07,0x32,0x90,0xc4,0xe7,0xe4}};
extern "C" const GUID __declspec(selectany) CLSID_RepresentationInfo =
    {0x1eabd900,0x9351,0x4306,{0x93,0x6b,0x47,0xdf,0xb3,0x17,0x0b,0xc4}};
extern "C" const GUID __declspec(selectany) CLSID_ParticipantInfo =
    {0xdf54c8f4,0x06a1,0x426f,{0x87,0xd9,0xe2,0xdf,0x86,0xff,0x5e,0x19}};
extern "C" const GUID __declspec(selectany) CLSID_HistoryInfo =
    {0x14a62ae0,0x9d65,0x49f2,{0x83,0x31,0xef,0x85,0xbf,0xf9,0x12,0x44}};
extern "C" const GUID __declspec(selectany) CLSID_CallParkOrbit =
    {0x378bf356,0x0ea3,0x4025,{0x84,0xfc,0xf6,0x44,0x30,0xad,0xc7,0x84}};
extern "C" const GUID __declspec(selectany) IID_ICallParkOrbit =
    {0x7cc306e9,0x2546,0x4101,{0x82,0xab,0x29,0x61,0x3a,0xd6,0xf7,0xa7}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationPropertyDictionary =
    {0xce124f88,0x3665,0x4a4a,{0xb2,0xe0,0x76,0x35,0x5f,0x83,0x2a,0xdf}};
extern "C" const GUID __declspec(selectany) CLSID_ModalityPropertyDictionary =
    {0xf251612d,0x4456,0x4acc,{0xbf,0x24,0xb0,0xd7,0x9f,0x00,0x4f,0x0b}};
extern "C" const GUID __declspec(selectany) CLSID_ParticipantPropertyDictionary =
    {0x939c4502,0x83de,0x4561,{0xbd,0x85,0xb5,0x66,0x80,0xc8,0x24,0x28}};
extern "C" const GUID __declspec(selectany) CLSID_ModalityDictionary =
    {0xd56d0f3e,0x24e3,0x421a,{0x8a,0xe8,0x5b,0xfb,0x8d,0x97,0xfe,0x90}};
extern "C" const GUID __declspec(selectany) CLSID_ModalityStateChangePropertyDictionary =
    {0x4a53d656,0x346e,0x4ddb,{0xad,0xab,0x7f,0x50,0x4c,0xe6,0xc0,0x2c}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationStateChangePropertyDictionary =
    {0x3e40c2d1,0x0a01,0x4d22,{0xa5,0xb8,0x3d,0xa4,0xe7,0xfd,0x01,0x70}};
extern "C" const GUID __declspec(selectany) CLSID_VoiceMail =
    {0x56d29009,0xc2f0,0x48b5,{0x89,0xc1,0x0f,0x72,0x51,0x99,0xec,0xea}};
extern "C" const GUID __declspec(selectany) CLSID_VoiceMailCollection =
    {0xe5fbeb06,0xf18e,0x4351,{0xb1,0x97,0x67,0xcd,0x6a,0xa2,0x2c,0x01}};
extern "C" const GUID __declspec(selectany) CLSID_VoiceMailsManager =
    {0x60a46701,0x3e4c,0x4115,{0x83,0xa3,0xb6,0x4e,0xee,0x5b,0x4c,0x0d}};
extern "C" const GUID __declspec(selectany) CLSID_PreviousConversation =
    {0xca757729,0xda0f,0x452d,{0xb0,0x01,0x93,0x65,0x56,0xbf,0xc1,0xd8}};
extern "C" const GUID __declspec(selectany) CLSID_PreviousConversationCollection =
    {0x563b1f6b,0xf388,0x448c,{0xb6,0x79,0xaf,0x1f,0xa3,0xa4,0x32,0x9a}};
extern "C" const GUID __declspec(selectany) CLSID_PreviousConversationsManager =
    {0x781a3b60,0x3cc1,0x42b2,{0x88,0x40,0x84,0xe1,0xa9,0x81,0x80,0x0f}};
extern "C" const GUID __declspec(selectany) CLSID_PreviousConversationsManagerDictionary =
    {0x4a7391d7,0x1a17,0x4996,{0x94,0x12,0x68,0x28,0x84,0x93,0x95,0xe8}};
extern "C" const GUID __declspec(selectany) CLSID_PreviousConversationsManagerConnectionStateChangedEventData =
    {0xd82fc397,0x451e,0x4e91,{0x85,0x57,0xd0,0x96,0xd7,0x67,0x92,0xfe}};
extern "C" const GUID __declspec(selectany) CLSID_PreviousConversationsManagerNewItemCountChangedEventData =
    {0xd85b4d22,0x7e0a,0x40f6,{0xac,0x97,0xfb,0x09,0xdf,0x0f,0xf9,0xd0}};
extern "C" const GUID __declspec(selectany) CLSID_PreviousConversationBatchEventData =
    {0x31d221f7,0x400e,0x4f04,{0x8e,0x3c,0x0b,0xf2,0x53,0x59,0x1d,0x31}};
extern "C" const GUID __declspec(selectany) CLSID_VoiceMailBatchEventData =
    {0xd5980e2d,0x0886,0x4ecd,{0x9d,0x2b,0x00,0x57,0xa6,0xc0,0x61,0xf3}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceAccessInformation =
    {0x19258088,0xaea6,0x4a73,{0xb3,0xd0,0xe4,0xb1,0xd9,0xe1,0x1d,0x9e}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationHistoryActionAvailabilityEventData =
    {0xc608ce35,0x98f7,0x41a6,{0x95,0x1b,0x91,0xc3,0x7d,0x2d,0x38,0x51}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationHistoryItemAddedEventData =
    {0x7eb20639,0x5a99,0x484e,{0x99,0xb3,0xc7,0x78,0xd2,0xf6,0x69,0x9e}};
extern "C" const GUID __declspec(selectany) CLSID_PreviousHistoryRetrievalStateChangedEventData =
    {0x8da3daf8,0xe6d0,0x4a19,{0x95,0x52,0xef,0x29,0x3f,0x01,0xd8,0xa4}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationContextCollection =
    {0x653a42da,0x8411,0x4902,{0x80,0x13,0x57,0x4f,0x1a,0xee,0xc4,0x76}};
extern "C" const GUID __declspec(selectany) CLSID_SharingResource =
    {0xebd24819,0xd80a,0x44de,{0x9c,0x7b,0x60,0x6c,0x51,0x2a,0x1d,0x71}};
extern "C" const GUID __declspec(selectany) CLSID_SharingResourceList =
    {0xf77d5a8d,0x8d57,0x40f0,{0xab,0x49,0x85,0xba,0xa3,0x69,0x92,0x38}};
extern "C" const GUID __declspec(selectany) CLSID_ControlRequestReceivedEventData =
    {0x797d542e,0x5beb,0x4926,{0x81,0x45,0x65,0x3c,0xb0,0x60,0x36,0x6f}};
extern "C" const GUID __declspec(selectany) CLSID_LocalSharedResourcesChangedEventData =
    {0xad0dfab8,0xd464,0x43c6,{0xa8,0xb4,0xf4,0x95,0xc4,0xcb,0x8a,0xfe}};
extern "C" const GUID __declspec(selectany) CLSID_ParticipationStateChangedEventData =
    {0x29c7c658,0x9ca5,0x40aa,{0xac,0x38,0xa2,0x5c,0xb6,0xf2,0xad,0x74}};
extern "C" const GUID __declspec(selectany) CLSID_ControllerChangedEventData =
    {0x271bbe62,0x8678,0x48d2,{0x80,0x7d,0x9b,0xc9,0x47,0x8c,0x63,0xd7}};
extern "C" const GUID __declspec(selectany) CLSID_SharerChangedEventData =
    {0x26fdcf1d,0xabe4,0x46fc,{0xa2,0xda,0x47,0x4a,0x47,0xda,0x82,0x37}};
extern "C" const GUID __declspec(selectany) IID_IPhone =
    {0x7693221d,0x7e9f,0x47d3,{0x87,0x82,0x64,0x37,0xcf,0x3d,0x4c,0x0c}};
extern "C" const GUID __declspec(selectany) IID_ICustomAvailabilityState =
    {0xd0984cef,0xe323,0x4574,{0xb0,0x7b,0x5c,0x97,0x0c,0x9c,0xae,0xe0}};
extern "C" const GUID __declspec(selectany) IID_IContactEndpointCollection =
    {0xdaa5fde3,0xa81c,0x4f23,{0x80,0xb6,0xc4,0x7b,0x52,0xc6,0x49,0xbf}};
extern "C" const GUID __declspec(selectany) IID_IPhonesChangedEventData =
    {0xc0034194,0xf7b6,0x43eb,{0xb0,0xe0,0x78,0x52,0xfc,0x8e,0x7b,0xfa}};
extern "C" const GUID __declspec(selectany) IID_IResiliencyModeChangedEventData =
    {0xdac6d77b,0x12f7,0x405c,{0xbb,0x33,0x00,0x2d,0x6d,0x40,0x07,0xf3}};
extern "C" const GUID __declspec(selectany) IID_IAlertLevel =
    {0xfb67275c,0x4d9e,0x4c32,{0x97,0x93,0xa9,0x67,0x49,0x88,0x4f,0xc1}};
extern "C" const GUID __declspec(selectany) IID_ISelf2 =
    {0x3c8d7146,0x35ea,0x4133,{0xb2,0xf6,0xc1,0xfc,0x24,0x01,0xa0,0x91}};
extern "C" const GUID __declspec(selectany) DIID__ISelfCallback =
    {0xa9aab6a0,0x54b9,0x4419,{0xaa,0xaf,0x6b,0x26,0xdf,0xac,0x15,0x85}};
extern "C" const GUID __declspec(selectany) DIID__ISelfEvents =
    {0x06437abb,0xc419,0x4b11,{0xa4,0x74,0x1a,0x2b,0x02,0xfb,0xd6,0x46}};
extern "C" const GUID __declspec(selectany) DIID__IAccessPermissionEvents =
    {0xdc4bc923,0x6a52,0x4f02,{0x9f,0x8a,0x54,0x7b,0x60,0x69,0x55,0xe7}};
extern "C" const GUID __declspec(selectany) DIID__ISelf2Events =
    {0x3c43a2be,0xba4a,0x41e8,{0x8f,0x93,0xa9,0xdc,0x5f,0x09,0x18,0x92}};
extern "C" const GUID __declspec(selectany) IID_IAlertLevelChangedEventData =
    {0x90b49a60,0x72af,0x4360,{0x9f,0x9b,0x91,0x6d,0xb7,0x50,0x4a,0x3c}};
extern "C" const GUID __declspec(selectany) CLSID_Self =
    {0xcdcf9caf,0xec72,0x4ec4,{0xa7,0xca,0x05,0xcb,0xdd,0xa1,0x95,0x9d}};
extern "C" const GUID __declspec(selectany) CLSID_AccessEntry =
    {0xb37552f7,0x36b6,0x4ef3,{0x80,0xc2,0x47,0xf1,0x79,0x57,0xe6,0x83}};
extern "C" const GUID __declspec(selectany) CLSID_Phone =
    {0x260e70c7,0x05bd,0x4c59,{0x9e,0xa3,0x5e,0xbb,0x9e,0xfe,0x5d,0x93}};
extern "C" const GUID __declspec(selectany) CLSID_AccessPermission =
    {0x19cc2b4b,0x2489,0x450c,{0x97,0x08,0xa6,0xc2,0x82,0x82,0xaa,0x76}};
extern "C" const GUID __declspec(selectany) CLSID_ContactEndpointCollection =
    {0x4cd03696,0xf8ba,0x4068,{0xac,0x3c,0xa6,0x11,0xf8,0xa4,0xbf,0x3e}};
extern "C" const GUID __declspec(selectany) CLSID_PhonesChangedEventData =
    {0xe769d3b9,0xf1ac,0x42f9,{0x97,0x02,0x43,0xae,0xb7,0xb3,0xd3,0xc2}};
extern "C" const GUID __declspec(selectany) CLSID_AccessEntryCollectionChangedEventData =
    {0xaca62ad7,0x05bf,0x4f21,{0x81,0xef,0xf9,0x05,0x17,0x0d,0xcf,0x23}};
extern "C" const GUID __declspec(selectany) CLSID_ResiliencyModeChangedEventData =
    {0x9fad5936,0x57c6,0x4e92,{0xb4,0xa2,0x10,0x60,0x11,0xa3,0xae,0x25}};
extern "C" const GUID __declspec(selectany) IID_IDevice =
    {0xce16d540,0xb9e8,0x4742,{0x86,0x59,0x9a,0x7a,0x31,0x8a,0xfb,0x92}};
extern "C" const GUID __declspec(selectany) IID_IDeviceCollection =
    {0x1f8ebe6f,0x7993,0x42ce,{0x98,0x0a,0xa2,0xf0,0x37,0x93,0xbe,0x71}};
extern "C" const GUID __declspec(selectany) IID_IAudioDevice =
    {0xb0b0e8f8,0x32da,0x4cb0,{0xa1,0x24,0xd1,0xe3,0x58,0x91,0xd5,0x3a}};
extern "C" const GUID __declspec(selectany) IID_IVideoDevice =
    {0xda7cf784,0x67df,0x429c,{0x82,0xfb,0x9a,0xfd,0x99,0xf6,0xcc,0xcb}};
extern "C" const GUID __declspec(selectany) IID_IDeviceManager =
    {0x9274dbdc,0x43ce,0x45aa,{0xa8,0x17,0x41,0x4a,0x44,0x94,0xad,0x28}};
extern "C" const GUID __declspec(selectany) DIID__IDeviceManagerCallback =
    {0x08ddf5c4,0xfa12,0x4978,{0xb2,0x6e,0xc6,0xd2,0x3c,0x45,0x34,0x13}};
extern "C" const GUID __declspec(selectany) CLSID_DeviceManager =
    {0xfda610fc,0x4565,0x4e70,{0xb0,0x88,0x4b,0x72,0xa1,0x6e,0x19,0x67}};
extern "C" const GUID __declspec(selectany) CLSID_DeviceCollection =
    {0xa0ba7d36,0x93cd,0x48c0,{0xa1,0x68,0x3e,0xf5,0x93,0x3c,0xee,0x8a}};
extern "C" const GUID __declspec(selectany) CLSID_Device =
    {0x81cb2ae5,0x6c07,0x4bc6,{0xbc,0x2b,0xf1,0xc0,0x45,0x45,0xce,0x54}};
extern "C" const GUID __declspec(selectany) CLSID_AudioDevice =
    {0xed1ce184,0x17aa,0x497f,{0xab,0xf2,0xc6,0x9b,0x21,0x44,0xdf,0x2a}};
extern "C" const GUID __declspec(selectany) CLSID_VideoDevice =
    {0xbaccb218,0x6318,0x4ad9,{0xb0,0xa0,0x5b,0x57,0x8c,0x05,0xf8,0x4e}};
extern "C" const GUID __declspec(selectany) IID_IConferenceInfoPropertyDictionary =
    {0x2d3d62ca,0xe9f6,0x4f94,{0x8e,0xbd,0x5f,0xab,0xab,0x29,0xe6,0xa5}};
extern "C" const GUID __declspec(selectany) IID_IConferenceInfo =
    {0xb3c5872e,0x591c,0x4a74,{0x9f,0x84,0xb0,0x0e,0xf9,0x70,0xa8,0x38}};
extern "C" const GUID __declspec(selectany) IID_IConferenceSettingDictionary =
    {0xbd891697,0xc93e,0x4a0b,{0x9b,0x6c,0x00,0x4b,0xd8,0x1b,0x6e,0xe8}};
extern "C" const GUID __declspec(selectany) IID_IConferenceScheduler =
    {0xc7acb102,0xb692,0x49cc,{0x92,0xda,0x58,0x24,0x82,0x2c,0x7b,0x96}};
extern "C" const GUID __declspec(selectany) IID_IConferenceInfoCollection =
    {0x1ff8aa01,0xd8af,0x4b2b,{0xb2,0x54,0xbe,0xfe,0xfa,0xba,0xf2,0xd6}};
extern "C" const GUID __declspec(selectany) IID_IConferenceRegionalAccessNumber =
    {0x9244d573,0x914f,0x4c1f,{0x93,0xf6,0x31,0x60,0x9a,0x95,0xcb,0xed}};
extern "C" const GUID __declspec(selectany) IID_IConferenceRegionalAccessNumberCollection =
    {0x4cb70415,0xba71,0x491e,{0xb1,0x2d,0x15,0x5d,0x85,0xcb,0xdf,0xf1}};
extern "C" const GUID __declspec(selectany) IID_IConferenceRegionalAccessInformation =
    {0x420a24e2,0x5c31,0x4262,{0x9b,0xd5,0x05,0x86,0x82,0x30,0x0e,0xd6}};
extern "C" const GUID __declspec(selectany) IID_IConferenceRegionalAccessInformationCollection =
    {0xa0984f01,0x5d2c,0x4302,{0x87,0xa0,0x69,0xbe,0x7b,0x01,0x51,0x43}};
extern "C" const GUID __declspec(selectany) IID_IConferenceTelephoneAccessInformation =
    {0x66d1ffc0,0xc234,0x4f3d,{0xab,0xaa,0x98,0x22,0x2a,0x69,0xe5,0x25}};
extern "C" const GUID __declspec(selectany) IID_IConferenceCapabilities =
    {0xe0e8b413,0xbb57,0x4b58,{0xb1,0xfb,0x28,0x8f,0x3a,0x26,0x9e,0x35}};
extern "C" const GUID __declspec(selectany) IID_IConferenceCapabilities2 =
    {0xe5cff4a5,0x83a8,0x4fd8,{0x8d,0xaa,0x77,0xa3,0xe2,0x28,0x8d,0x4a}};
extern "C" const GUID __declspec(selectany) DIID__IConferenceSchedulerCallback =
    {0xc22ce6bf,0x3a47,0x46b7,{0x85,0x44,0x44,0x30,0x5b,0xad,0xfe,0xf9}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceScheduler =
    {0x57dd2d13,0xf18d,0x4a4e,{0xa0,0x22,0x2f,0x8c,0x26,0x9a,0x4a,0xfd}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceInfo =
    {0x2ae6db87,0x2dfd,0x427d,{0xb8,0xe3,0xb4,0xd7,0x38,0xe2,0x68,0x4e}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceInfoPropertyDictionary =
    {0x40368b76,0xa6ad,0x44c8,{0xa0,0x88,0xcd,0xea,0x02,0x14,0xcb,0x10}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceApplicationData =
    {0x960390d4,0xc70d,0x43a1,{0xb2,0x72,0x22,0xe7,0x3b,0xc4,0xd4,0xb6}};
extern "C" const GUID __declspec(selectany) IID_IConferenceApplicationProperties =
    {0x64b7e9ad,0xd920,0x4eaa,{0x95,0xc7,0x9d,0x95,0xde,0x9a,0xce,0x24}};
extern "C" const GUID __declspec(selectany) IID_IConferenceApplicationData =
    {0xbe6086a1,0xd436,0x4834,{0x89,0xee,0x3c,0xa4,0xf8,0x59,0x6a,0x58}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceApplicationProperties =
    {0x541fe2a9,0x95be,0x4098,{0x96,0x0a,0x64,0x46,0x4b,0x25,0x14,0x05}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceSettingDictionary =
    {0x7f10fab7,0x2a38,0x444c,{0x8f,0x43,0x31,0xdd,0xf0,0xfe,0x79,0xae}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceCapabilities =
    {0x488f9ce1,0x5468,0x48ec,{0xb4,0xdb,0xfc,0x2f,0x7f,0x1a,0xb2,0x47}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceInfoCollection =
    {0x8b8cbacf,0x0aa1,0x419e,{0x94,0xd2,0x60,0xf8,0x44,0xa6,0xfb,0x9e}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceTelephoneAccessInformation =
    {0xbf527f65,0x83f6,0x4fa4,{0xab,0x1b,0x17,0xa8,0x58,0x5f,0x0d,0x8a}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceRegionalAccessInformation =
    {0x36815016,0x484e,0x41f5,{0x9e,0x54,0xb9,0x52,0xf4,0xb4,0x44,0x00}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceRegionalAccessInformationCollection =
    {0xf080136d,0xd09a,0x42de,{0xb9,0xe4,0x2e,0x92,0x6d,0x1a,0x43,0x0f}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceRegionalAccessNumberCollection =
    {0xba9b1f14,0xce79,0x417e,{0x93,0x20,0x70,0x4c,0x9e,0x39,0x11,0x16}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceRegionalAccessNumber =
    {0xb82a9c56,0x654d,0x4a24,{0xb3,0x08,0xab,0xa4,0x9f,0x65,0x18,0x3a}};
extern "C" const GUID __declspec(selectany) DIID__IUtilitiesCallback =
    {0xc4d95336,0x5949,0x4d13,{0xaa,0xc8,0x82,0x03,0x9e,0x5c,0xbc,0xfc}};
extern "C" const GUID __declspec(selectany) IID_IScheduledConferenceManager =
    {0x26f15083,0x86db,0x4663,{0x89,0xed,0xa0,0xab,0xc8,0x96,0xb9,0xa1}};
extern "C" const GUID __declspec(selectany) IID_IUtilities2 =
    {0x9112ce9b,0x3704,0x48c0,{0x86,0xdc,0x7a,0x7b,0x0f,0x34,0xd7,0xa3}};
extern "C" const GUID __declspec(selectany) DIID__IUtilities2Events =
    {0x3417d640,0x3976,0x4cb6,{0xbc,0x82,0x54,0xe5,0xe2,0xe1,0x07,0x84}};
extern "C" const GUID __declspec(selectany) IID_IExchangeECPUrlChangedEventData =
    {0x7818a3f7,0x674b,0x477e,{0x99,0xe8,0x54,0x77,0x45,0x79,0xa8,0xe7}};
extern "C" const GUID __declspec(selectany) CLSID_Utilities =
    {0xb5b1ff8f,0x3f61,0x4ea5,{0xa3,0x39,0xd9,0x4f,0xdf,0xaf,0x15,0xe1}};
extern "C" const GUID __declspec(selectany) IID_IScheduledConference =
    {0x71a054b3,0x4d4b,0x4643,{0xb1,0x4f,0x45,0x72,0x49,0x4b,0x19,0xe7}};
extern "C" const GUID __declspec(selectany) IID_IScheduledConferenceCollection =
    {0xf423726d,0x0e9b,0x4b55,{0x95,0x69,0xe7,0x98,0x65,0x21,0x0f,0x69}};
extern "C" const GUID __declspec(selectany) DIID__IScheduledConferenceManagerCallback =
    {0x219dc126,0x9a48,0x483f,{0x80,0xc2,0x3f,0x22,0xb3,0xb4,0x78,0x29}};
extern "C" const GUID __declspec(selectany) CLSID_Utilities2 =
    {0xead92b6d,0xa7db,0x4ab8,{0xa4,0x94,0xc5,0xa2,0x77,0x1f,0xa9,0x87}};
extern "C" const GUID __declspec(selectany) CLSID_ExchangeECPUrlChangedEventData =
    {0xd40faf18,0x96f8,0x407e,{0xba,0x32,0x48,0xda,0xba,0x52,0x1e,0x18}};
extern "C" const GUID __declspec(selectany) CLSID_ScheduledConference =
    {0xbb99483d,0x2df9,0x424f,{0x96,0xa8,0x67,0x18,0xcd,0x7f,0xbb,0x9f}};
extern "C" const GUID __declspec(selectany) CLSID_ScheduledConferenceCollection =
    {0x584dc0a4,0xf664,0x4a1d,{0xb7,0x81,0x66,0x2c,0x08,0x2d,0x6b,0x78}};
extern "C" const GUID __declspec(selectany) CLSID_ScheduledConferenceManager =
    {0x75927b32,0x8f46,0x40eb,{0xa6,0xc9,0x75,0x98,0xf0,0xdb,0x38,0xc0}};
extern "C" const GUID __declspec(selectany) IID_IRoomMessageDictionary =
    {0x603e2454,0x725a,0x4679,{0xa6,0xa8,0x36,0x3b,0x21,0x63,0x3c,0xb2}};
extern "C" const GUID __declspec(selectany) IID_IRoomMessage =
    {0xa0c3e9ba,0xc622,0x45a2,{0x9f,0xed,0x2c,0x66,0x2f,0x3f,0x48,0x9a}};
extern "C" const GUID __declspec(selectany) IID_IRoomMessageCollection =
    {0x21c41ede,0x7ed8,0x4d97,{0x9b,0x0d,0xdd,0x78,0x57,0x81,0xaf,0x9b}};
extern "C" const GUID __declspec(selectany) IID_IRoomUser =
    {0x68da8699,0x6c6e,0x4390,{0xbd,0x82,0x70,0x9f,0xf1,0xdc,0x17,0x02}};
extern "C" const GUID __declspec(selectany) IID_IRoomPropertyDictionary =
    {0x7bf20b14,0x58d1,0x494b,{0xb3,0x01,0x9b,0x16,0xba,0xcc,0x96,0x10}};
extern "C" const GUID __declspec(selectany) IID_IRoom =
    {0xa766520d,0xf700,0x4bf5,{0x84,0x28,0x56,0xbd,0x91,0x4c,0x1c,0x40}};
extern "C" const GUID __declspec(selectany) IID_IRoomCollection =
    {0x05789509,0xafa1,0x4e90,{0xb4,0xfa,0x1a,0x78,0x20,0xdb,0x97,0x96}};
extern "C" const GUID __declspec(selectany) IID_IRoomManager =
    {0x76de1cd8,0xf278,0x48a1,{0xa3,0x91,0xe5,0xbe,0xb8,0xf4,0x28,0xe6}};
extern "C" const GUID __declspec(selectany) IID_IRoomPropertyChangedEventData =
    {0xb03614ae,0x8d68,0x4386,{0x9e,0x1c,0x93,0x9c,0xab,0xcf,0x12,0x32}};
extern "C" const GUID __declspec(selectany) IID_IUnreadMessageCountChangedEventData =
    {0xb8fd4a53,0xe7e6,0x4995,{0xa5,0xb5,0x13,0x06,0xc7,0x58,0x49,0x64}};
extern "C" const GUID __declspec(selectany) IID_IRoomMessageEventData =
    {0x1566fa3c,0xfd24,0x41db,{0xa2,0xb8,0x62,0x32,0xf8,0x97,0x83,0xde}};
extern "C" const GUID __declspec(selectany) IID_IRoomMessagesEventData =
    {0x081793da,0x7221,0x45c3,{0x8a,0xfa,0x3a,0x2e,0x6e,0x46,0x46,0x70}};
extern "C" const GUID __declspec(selectany) IID_IRoomUserChangedEventData =
    {0xbc9f5f40,0xd315,0x40c1,{0xb1,0xc6,0xce,0xa0,0x64,0x6e,0x18,0xe2}};
extern "C" const GUID __declspec(selectany) IID_IRoomParticipantsEventData =
    {0x7452bd0f,0x65cb,0x4a5e,{0xac,0x37,0xe2,0xbe,0xa1,0xf4,0x3d,0xd9}};
extern "C" const GUID __declspec(selectany) IID_IRoomJoinStateChangedEventData =
    {0x4d120020,0xce64,0x43c5,{0x9f,0x84,0x7a,0x7b,0x23,0x60,0x38,0x8f}};
extern "C" const GUID __declspec(selectany) DIID__IRoomEvents =
    {0x300e56a3,0xce08,0x4eb3,{0x9f,0x48,0x50,0x5a,0xa1,0x62,0xc9,0xda}};
extern "C" const GUID __declspec(selectany) DIID__IRoomManagerEvents =
    {0xe2c6d6e0,0x52ef,0x47a2,{0xbd,0xa9,0x9d,0x6a,0x7a,0x29,0xa3,0x91}};
extern "C" const GUID __declspec(selectany) IID_IFollowedRoomsChangedEventData =
    {0x8f04e92d,0xee26,0x4b64,{0xb9,0xc1,0x56,0xdd,0x90,0x81,0x68,0x10}};
extern "C" const GUID __declspec(selectany) IID_IRoomManagerStateChangedEventData =
    {0x89c7eb8e,0x42bc,0x4c9f,{0xbb,0x34,0x88,0xcd,0xe8,0x3c,0xcb,0x37}};
extern "C" const GUID __declspec(selectany) DIID__IRoomUserEvents =
    {0x8e73e7e1,0x90ce,0x4d3b,{0xb5,0xb7,0x71,0xe8,0x7b,0x2e,0x8b,0xa5}};
extern "C" const GUID __declspec(selectany) DIID__IRoomCallback =
    {0x628d7bdd,0x45ab,0x404d,{0x85,0x05,0xb3,0xad,0xc1,0x5f,0x45,0x4d}};
extern "C" const GUID __declspec(selectany) DIID__IRoomManagerCallback =
    {0x0fdb5b86,0xdd44,0x41bf,{0xa4,0x43,0x31,0x7f,0xd1,0x9b,0xbf,0x3d}};
extern "C" const GUID __declspec(selectany) CLSID_RoomMessage =
    {0x44e7ad71,0xd7d5,0x45b7,{0x9f,0xaa,0x6c,0x34,0x6c,0x6d,0xa1,0x05}};
extern "C" const GUID __declspec(selectany) CLSID_RoomUser =
    {0xad75a6fb,0x14e3,0x4830,{0xad,0xf1,0xc7,0x06,0x7d,0x65,0x50,0xa7}};
extern "C" const GUID __declspec(selectany) CLSID_Room =
    {0x4fac92c8,0xe226,0x44b6,{0xa5,0x30,0xae,0xd7,0xae,0xb4,0xce,0xf4}};
extern "C" const GUID __declspec(selectany) CLSID_RoomManager =
    {0x0909f1e4,0x5f02,0x4e09,{0x85,0x7d,0xbd,0x95,0x1e,0x95,0x6f,0xc1}};
extern "C" const GUID __declspec(selectany) CLSID_RoomMessageCollection =
    {0x3068d48b,0x3c56,0x4b15,{0xac,0xa8,0x97,0xa9,0xc8,0xb4,0x3c,0x13}};
extern "C" const GUID __declspec(selectany) CLSID_RoomCollection =
    {0x3df72257,0xea41,0x45a7,{0xbd,0x01,0xdc,0xf2,0x10,0x4b,0x25,0x0c}};
extern "C" const GUID __declspec(selectany) CLSID_RoomPropertyChangedEventData =
    {0x0a52c7b8,0x6123,0x413b,{0xad,0x6a,0x05,0xfa,0xff,0x69,0xc5,0x19}};
extern "C" const GUID __declspec(selectany) CLSID_UnreadMessageCountChangedEventData =
    {0x96e4b115,0x4c68,0x4038,{0xa0,0xdb,0x37,0x77,0xc4,0xae,0x4a,0xeb}};
extern "C" const GUID __declspec(selectany) CLSID_RoomMessageEventData =
    {0x8d69edcf,0xb7c8,0x4bd1,{0xbf,0xd2,0x08,0xc8,0xc5,0xb1,0xe0,0x30}};
extern "C" const GUID __declspec(selectany) CLSID_RoomMessagesEventData =
    {0x09ec6931,0xac8d,0x4b45,{0x88,0xb1,0xde,0xd5,0x20,0xf5,0x62,0xdc}};
extern "C" const GUID __declspec(selectany) CLSID_FollowedRoomsChangedEventData =
    {0xfcb650ff,0xe3a0,0x47bc,{0xac,0xd6,0xc1,0xc3,0xc8,0x47,0x3a,0xe4}};
extern "C" const GUID __declspec(selectany) CLSID_RoomManagerStateChangedEventData =
    {0xe269b211,0xd1f7,0x45fc,{0x92,0x02,0x43,0x7f,0x71,0xe9,0xf0,0xb3}};
extern "C" const GUID __declspec(selectany) CLSID_RoomUserChangedEventData =
    {0x31b10e61,0xfccb,0x42d3,{0x87,0x3d,0x93,0x6c,0xa5,0x21,0x60,0xe8}};
extern "C" const GUID __declspec(selectany) CLSID_RoomParticipantsEventData =
    {0x3e207f26,0xc1de,0x43f6,{0xbd,0x2b,0x8a,0x46,0x61,0x53,0x25,0xbf}};
extern "C" const GUID __declspec(selectany) CLSID_RoomJoinStateChangedEventData =
    {0xcdfc37e9,0xbd76,0x420f,{0x88,0x1f,0x4a,0x6a,0xab,0x4e,0x51,0xbf}};
extern "C" const GUID __declspec(selectany) CLSID_RoomPropertyDictionary =
    {0xb4f20beb,0x831c,0x4bb8,{0x90,0x65,0x39,0x64,0xe4,0x33,0x0e,0x79}};
extern "C" const GUID __declspec(selectany) CLSID_RoomMessageDictionary =
    {0x6c4db7f3,0x3a3b,0x410c,{0x94,0x4c,0x02,0x72,0x89,0xe7,0x98,0x04}};
extern "C" const GUID __declspec(selectany) IID_IConversationWindowNeedsSizeChangeEventData =
    {0xca62e6f3,0x3e6d,0x451b,{0xae,0xaa,0x6a,0x1b,0x7a,0xac,0xf8,0xd4}};
extern "C" const GUID __declspec(selectany) IID_IConversationWindowNeedsAttentionEventData =
    {0xf97bed54,0xe434,0x4020,{0xa1,0x97,0xf1,0x5a,0xea,0x9d,0x9c,0x95}};
extern "C" const GUID __declspec(selectany) DIID__IAutomationCallback =
    {0x806d3227,0x4cb8,0x47c4,{0x98,0x64,0x7d,0x4d,0xf4,0xf4,0x40,0x69}};
extern "C" const GUID __declspec(selectany) DIID__IConversationWindowEvents =
    {0x9e0f8fca,0xd9c9,0x47cd,{0x87,0xf4,0x55,0x54,0x21,0x7d,0xff,0xcc}};
extern "C" const GUID __declspec(selectany) DIID__IConversationWindowCallback =
    {0x76428270,0x36fa,0x4236,{0x8b,0xdf,0xaa,0xdb,0x39,0xfd,0x13,0x71}};
extern "C" const GUID __declspec(selectany) IID_IConversationWindowInformationDictionary =
    {0x109d32ce,0xe09f,0x46f5,{0x9e,0xcb,0x01,0xe1,0xd2,0x73,0xb7,0xda}};
extern "C" const GUID __declspec(selectany) IID_IConversationWindowActions =
    {0x663c61f4,0xa9da,0x4a6f,{0xa6,0x06,0xf1,0x42,0xf6,0x65,0xdf,0x61}};
extern "C" const GUID __declspec(selectany) IID_IVideoViewInformationDictionary =
    {0xb00f2520,0x029c,0x47d0,{0xb4,0xe8,0x8f,0xbe,0xf4,0x7c,0xaa,0x7e}};
extern "C" const GUID __declspec(selectany) IID_IVideoView =
    {0x55a21cf3,0xa2b3,0x484e,{0xbe,0x2a,0x14,0x28,0x0f,0x50,0x12,0x89}};
extern "C" const GUID __declspec(selectany) IID_IConversationWindow2 =
    {0x02dd541f,0xc863,0x4295,{0x92,0xab,0xa4,0xcf,0x29,0xf5,0xd8,0x41}};
extern "C" const GUID __declspec(selectany) DIID__IConversationWindow2Events =
    {0xaf2aba54,0xf591,0x42ba,{0x80,0x04,0x20,0x94,0x4c,0x9f,0xb4,0xd7}};
extern "C" const GUID __declspec(selectany) IID_IConversationWindowStateChangedEventData =
    {0x2209e091,0x5b90,0x4caf,{0x85,0x82,0xcc,0x59,0x79,0x88,0x33,0x76}};
extern "C" const GUID __declspec(selectany) IID_IConversationWindowInformationChangedEventData =
    {0xff037274,0x455a,0x4e34,{0xb5,0xd1,0xd4,0x2d,0xb8,0x66,0xf9,0xb7}};
extern "C" const GUID __declspec(selectany) IID_IConversationWindowActionAvailabilityChangedEventData =
    {0x0678c83e,0xf580,0x4d99,{0x90,0x2f,0x93,0x06,0x99,0xb2,0x8b,0xe8}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationWindow =
    {0x440f28e7,0x13b4,0x4e6a,{0x84,0x83,0x65,0xad,0x69,0xb8,0x30,0x70}};
extern "C" const GUID __declspec(selectany) CLSID_Automation =
    {0xe45fb8a0,0x6716,0x4f0f,{0xa4,0xf0,0x22,0x0c,0x3d,0xeb,0x73,0x37}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationWindowNeedsSizeChangeEventData =
    {0x3dc07b52,0x702c,0x47fe,{0x88,0x93,0xb9,0x6e,0x2a,0x79,0x81,0x63}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationWindowNeedsAttentionEventData =
    {0xac4f1849,0x52c8,0x4dcf,{0xa2,0x1f,0x93,0x20,0x6b,0x32,0x0a,0xf0}};
extern "C" const GUID __declspec(selectany) IID_IShareableContentPropertyDictionary =
    {0xca478b44,0x4fce,0x4082,{0x87,0x00,0x8c,0xec,0x69,0xfb,0x50,0xe1}};
extern "C" const GUID __declspec(selectany) DIID__IContentSharingModalityEvents =
    {0xae3a7c39,0x9c30,0x4e0f,{0xba,0xe4,0xd8,0x34,0x4e,0xf3,0x77,0xea}};
extern "C" const GUID __declspec(selectany) DIID__IShareableContentEvents =
    {0x6b551c9e,0xde81,0x41df,{0xa0,0xae,0x39,0xf0,0xaf,0x11,0xd5,0x08}};
extern "C" const GUID __declspec(selectany) IID_IShareableContentStateChangedEventData =
    {0x94093301,0x8854,0x408e,{0x81,0x38,0xc7,0xc3,0xd2,0xda,0xca,0x1f}};
extern "C" const GUID __declspec(selectany) IID_IShareableContentPropertyChangedEventData =
    {0x3c8210c8,0x8578,0x47c6,{0x87,0xa9,0xfa,0x1a,0xd2,0xba,0x98,0x73}};
extern "C" const GUID __declspec(selectany) IID_IShareableContentActionAvailabilityChangedEventData =
    {0x4023052e,0xe3b7,0x48a3,{0xbc,0xb3,0xca,0x22,0x34,0xd9,0xba,0xee}};
extern "C" const GUID __declspec(selectany) DIID__IConversation2Callback =
    {0x6964b116,0x8fe6,0x4c19,{0x90,0x52,0x0e,0xaa,0x74,0x46,0x0a,0x58}};
extern "C" const GUID __declspec(selectany) DIID__IContentSharingModalityCallback =
    {0x84340964,0x7820,0x4ebc,{0xbc,0xd3,0x70,0x29,0x26,0xde,0x23,0xe8}};
extern "C" const GUID __declspec(selectany) DIID__IShareableContentCallback =
    {0x3187885c,0xa8af,0x4d4b,{0x8e,0x80,0xf4,0xb9,0xd4,0x47,0xa7,0xb2}};
extern "C" const GUID __declspec(selectany) DIID__IParticipant2Callback =
    {0x36cdc166,0x4f21,0x46ad,{0xa6,0x0e,0x85,0x51,0xf2,0x6c,0x1d,0x41}};
extern "C" const GUID __declspec(selectany) DIID__IChannel2Callback =
    {0x70986dce,0x4aa8,0x4dfb,{0x9a,0x45,0x4b,0xfc,0xf7,0xa2,0x06,0xf5}};
extern "C" const GUID __declspec(selectany) CLSID_Conversation2 =
    {0x46b5e644,0xa89e,0x4c63,{0x86,0xe5,0x34,0x66,0x68,0x07,0xee,0x75}};
extern "C" const GUID __declspec(selectany) CLSID_Participant2 =
    {0x6fb3b002,0x987c,0x460c,{0x9e,0x9f,0xb4,0x5f,0x2e,0xd3,0xb1,0xe9}};
extern "C" const GUID __declspec(selectany) CLSID_Channel2 =
    {0xa351a6d6,0x9b47,0x41a3,{0xa6,0xd3,0x07,0xdd,0xf2,0x20,0x3e,0x7d}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceAccessInformation2 =
    {0xed094732,0xddca,0x4a31,{0x80,0x95,0xe9,0x77,0x4e,0x3d,0xbc,0x0d}};
extern "C" const GUID __declspec(selectany) IID_IConferenceAccessInformation2 =
    {0x5e7ab04c,0xe4b8,0x49fb,{0x85,0xff,0x9e,0x2b,0xca,0x28,0x99,0xd3}};
extern "C" const GUID __declspec(selectany) CLSID_ContentSharingModality =
    {0x509879fa,0x3ee5,0x4328,{0x86,0x25,0xa0,0xdf,0xb9,0x21,0x1b,0x63}};
extern "C" const GUID __declspec(selectany) CLSID_ShareableContentPropertyDictionary =
    {0x235b711c,0x6fa3,0x4f91,{0xb5,0x0e,0x36,0xed,0x53,0xb3,0x80,0x1f}};
extern "C" const GUID __declspec(selectany) CLSID_ShareableContent =
    {0xdff47f2b,0xad75,0x44dc,{0x81,0xaa,0xe2,0x3c,0xa7,0xc6,0xa6,0x6f}};
extern "C" const GUID __declspec(selectany) CLSID_PowerPointContent =
    {0x2d1a91b0,0xa593,0x4327,{0xa7,0xb2,0xf7,0xdd,0x01,0x57,0x4b,0xf3}};
extern "C" const GUID __declspec(selectany) CLSID_ShareableContentCollection =
    {0x90f87439,0xacd6,0x4802,{0xa7,0x40,0x59,0x05,0x11,0x66,0xe5,0x7d}};
extern "C" const GUID __declspec(selectany) CLSID_ContentCollectionChangedEventData =
    {0x34c459ac,0x5125,0x42d4,{0xbf,0x92,0xb1,0x49,0x88,0xf0,0x28,0x9e}};
extern "C" const GUID __declspec(selectany) CLSID_ActiveContentChangedEventData =
    {0x411768f9,0xa6e5,0x4857,{0x8a,0x10,0x79,0x23,0xc9,0xfc,0xbb,0x1e}};
extern "C" const GUID __declspec(selectany) CLSID_ActivePresenterChangedEventData =
    {0x12063e4a,0xc16f,0x4e0f,{0x88,0x86,0x61,0xdd,0xf5,0x2f,0x40,0x90}};
extern "C" const GUID __declspec(selectany) CLSID_ShareableContentPropertyChangedEventData =
    {0x71a27f3d,0x167b,0x42de,{0xa8,0xe6,0x5c,0xd8,0x65,0x8b,0xc5,0xd8}};
extern "C" const GUID __declspec(selectany) CLSID_ShareableContentStateChangedEventData =
    {0x3ba1b3e8,0xf66f,0x4eec,{0xa8,0x4e,0x82,0x56,0xdd,0xad,0x60,0x6e}};
extern "C" const GUID __declspec(selectany) CLSID_ShareableContentActionAvailabilityChangedEventData =
    {0xa7b1e1de,0xcf73,0x492f,{0x95,0xc2,0xdf,0xdc,0x02,0xd4,0xee,0x2b}};
extern "C" const GUID __declspec(selectany) CLSID_Self2 =
    {0x8b9f0b59,0x7c29,0x46e2,{0xb6,0xa6,0x80,0x40,0x43,0xa3,0x46,0xdb}};
extern "C" const GUID __declspec(selectany) CLSID_AlertLevel =
    {0x3bc642ef,0x64be,0x45ae,{0x98,0x2d,0xba,0xbd,0xaa,0xbb,0x35,0x5a}};
extern "C" const GUID __declspec(selectany) CLSID_AlertLevelChangedEventData =
    {0x17374941,0x9877,0x429f,{0xa8,0x8c,0x8e,0xcd,0x0b,0xfc,0x49,0xd4}};
extern "C" const GUID __declspec(selectany) IID_IVideoViewInformationChangedEventData =
    {0x086f0e9d,0xa416,0x40f9,{0x87,0x7c,0xf1,0xc8,0x2d,0xc1,0xa6,0xad}};
extern "C" const GUID __declspec(selectany) DIID__IVideoViewEvents =
    {0x30f974fb,0x1f7b,0x49d3,{0x8d,0xa4,0x01,0x69,0x4b,0xa4,0x74,0xd2}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationWindow2 =
    {0xb58839e6,0xeb13,0x402d,{0xbb,0x2c,0x83,0xeb,0xae,0x14,0xe6,0xaa}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationWindowInformationDictionary =
    {0x360c1678,0x114e,0x42c0,{0xa8,0xff,0xea,0x8b,0xa1,0x0c,0xa2,0xd8}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationWindowStateChangedEventData =
    {0x0e32c65c,0xc4d9,0x4280,{0x91,0x42,0xec,0x10,0x2e,0x19,0xa0,0x5b}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationWindowInformationChangedEventData =
    {0xf10b6480,0x558f,0x49f3,{0xa5,0x15,0xab,0x6d,0x05,0xfb,0x94,0x9f}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationWindowActionAvailabilityChangedEventData =
    {0x7c4db3cb,0x796e,0x4993,{0x8d,0x49,0x41,0x43,0x17,0xd0,0x14,0xda}};
extern "C" const GUID __declspec(selectany) CLSID_ConversationWindowActions =
    {0x85147550,0x4a7b,0x4b36,{0x8f,0xb5,0x14,0x4b,0x50,0x87,0xa2,0x99}};
extern "C" const GUID __declspec(selectany) CLSID_VideoView =
    {0x24efe0bc,0x5661,0x4f96,{0xad,0xc4,0x35,0x1d,0xbe,0x47,0xa4,0xcc}};
extern "C" const GUID __declspec(selectany) CLSID_VideoViewInformationChangedEventData =
    {0xa13f4ab3,0xc0fc,0x466a,{0x94,0x57,0x2d,0xf0,0xad,0x8c,0x74,0xed}};
extern "C" const GUID __declspec(selectany) CLSID_VideoViewInformationDictionary =
    {0x4c7d404c,0xe8a6,0x416c,{0x8f,0x5f,0xb5,0xb9,0x5f,0x63,0x4b,0x47}};
extern "C" const GUID __declspec(selectany) CLSID_Contact2 =
    {0xfd28f0a8,0xd83a,0x40fb,{0x8c,0x30,0x69,0x08,0xa6,0x11,0x7d,0x5b}};
extern "C" const GUID __declspec(selectany) CLSID_SearchResults2 =
    {0x24290b3e,0x7dc8,0x485b,{0x85,0xa1,0x95,0x6b,0xcd,0x8d,0x2c,0x5c}};
extern "C" const GUID __declspec(selectany) IID_IClientSettings =
    {0x1b04d929,0x11e4,0x46c2,{0x91,0x62,0xea,0x4f,0x27,0x58,0xc3,0xa7}};
extern "C" const GUID __declspec(selectany) IID_IClientSettingsChangedEventData =
    {0x25d64aea,0x0e65,0x49cb,{0x8d,0x6d,0x65,0xdb,0x0a,0xc1,0xaf,0x65}};
extern "C" const GUID __declspec(selectany) IID_IPublicCloudProviderInfo =
    {0x74b6caf0,0x755e,0x454d,{0xa8,0x19,0xd0,0xe4,0x39,0xef,0x52,0x96}};
extern "C" const GUID __declspec(selectany) DIID__IClientSettingsEvents =
    {0x569e04f2,0xd286,0x4903,{0x96,0x42,0x44,0xd8,0x50,0x31,0x09,0x56}};
extern "C" const GUID __declspec(selectany) CLSID_Client2 =
    {0x66d123ba,0x11e2,0x4c2c,{0x80,0x56,0x5a,0xca,0x97,0x24,0x5c,0xd2}};
extern "C" const GUID __declspec(selectany) CLSID_ClientSettings =
    {0x8eb30566,0x8e31,0x4dd9,{0xb1,0xe7,0x47,0x77,0x56,0x16,0x00,0x76}};
extern "C" const GUID __declspec(selectany) CLSID_ClientSettingsChangedEventData =
    {0xf19e9036,0x6db9,0x4d48,{0x9f,0x46,0x26,0x99,0x5f,0x9d,0x51,0x63}};
extern "C" const GUID __declspec(selectany) CLSID_PublicCloudProviderInfo =
    {0x8e99de73,0x6a8b,0x47a9,{0xba,0xe6,0x6b,0x03,0xe5,0x36,0x22,0xe8}};
extern "C" const GUID __declspec(selectany) CLSID_ConferenceCapabilities2 =
    {0xe8451131,0x7352,0x4906,{0x9e,0x57,0x5c,0xb1,0x64,0xfe,0xe0,0x62}};
extern "C" const GUID __declspec(selectany) IID_IUCOfficeIntegration =
    {0x6a222195,0xf65e,0x467f,{0x8f,0x77,0xeb,0x18,0x0b,0xd8,0x52,0x88}};
extern "C" const GUID __declspec(selectany) IID_ISignInConfiguration =
    {0xf8df0b4b,0xa4ce,0x476b,{0x81,0xd3,0x9f,0xba,0x98,0x46,0x1b,0x62}};
extern "C" const GUID __declspec(selectany) IID_IApplicationRegistration =
    {0xaa41ee75,0xf9c5,0x42f7,{0x8d,0x17,0x1d,0x92,0xd6,0xba,0x4d,0x92}};
extern "C" const GUID __declspec(selectany) IID_IWebAppExtension =
    {0x330639f2,0xe399,0x48cb,{0x86,0x3f,0x56,0xa5,0x0a,0x27,0xa1,0x38}};
extern "C" const GUID __declspec(selectany) IID_IClientStateChangedEventData =
    {0x9b3a702d,0x5be1,0x4fce,{0xad,0xb3,0xeb,0xbd,0x23,0xe0,0x78,0xc3}};
extern "C" const GUID __declspec(selectany) IID_ILyncClientNotificationReceivedEventData =
    {0x9ab20314,0xb258,0x48f8,{0xb6,0x59,0xad,0x25,0x0d,0xfa,0x20,0xf5}};
extern "C" const GUID __declspec(selectany) IID_ICredentialRequestedEventData =
    {0xf0b791fa,0x434b,0x49f2,{0xb8,0x4e,0x05,0x61,0x9e,0xbe,0x6b,0xeb}};
extern "C" const GUID __declspec(selectany) IID_ISignInDelayedEventData =
    {0xe664682c,0xcba5,0x4c7c,{0xa6,0x5b,0x9c,0xd0,0x46,0xf2,0x4c,0xbd}};
extern "C" const GUID __declspec(selectany) IID_IPreferredCapabilitiesChangedEventData =
    {0xb28f2007,0x3926,0x49cf,{0x91,0x01,0x4c,0xcb,0x95,0x74,0x04,0x9b}};
extern "C" const GUID __declspec(selectany) DIID__ILyncClientCallback =
    {0xb26f6403,0x761e,0x48d4,{0x84,0xe2,0x61,0x9c,0x17,0xea,0xeb,0x13}};
extern "C" const GUID __declspec(selectany) DIID__IClientEvents =
    {0x64561371,0x77a2,0x4bd4,{0xa9,0x09,0xe6,0xde,0xb1,0xb5,0xc6,0xcf}};
extern "C" const GUID __declspec(selectany) DIID__ILyncClientEvents =
    {0x953ae732,0xf53a,0x4116,{0xac,0x1b,0x03,0x21,0xb3,0xfb,0x3d,0xba}};
extern "C" const GUID __declspec(selectany) DIID__IDelegatorClientEvents =
    {0x92a9eb9b,0x85ee,0x444e,{0xa3,0x02,0x00,0xc3,0xc9,0x45,0xba,0xa4}};
extern "C" const GUID __declspec(selectany) DIID__IUCOfficeIntegrationEvents =
    {0x8c9b6953,0x33cc,0x4a04,{0x9d,0xa4,0xf7,0x1a,0xe7,0x9d,0xa0,0xc0}};
extern "C" const GUID __declspec(selectany) CLSID_Client =
    {0x3f927576,0x996e,0x4a91,{0x9c,0xf9,0xd9,0xfa,0x46,0xd2,0xaa,0xf5}};
extern "C" const GUID __declspec(selectany) CLSID_LyncClient =
    {0xa9f3c4a7,0x6de0,0x48e4,{0x81,0x84,0xba,0x7d,0x6d,0x98,0x59,0x08}};
extern "C" const GUID __declspec(selectany) CLSID_SignInConfiguration =
    {0x6a69dc0b,0xf5ab,0x47e7,{0x85,0x00,0x00,0xed,0x31,0x48,0x6c,0x01}};
extern "C" const GUID __declspec(selectany) CLSID_ApplicationRegistration =
    {0xb7a53e76,0x8630,0x4a93,{0x97,0x71,0x94,0x00,0x12,0xa1,0x37,0x1b}};
extern "C" const GUID __declspec(selectany) CLSID_CustomAvailabilityState =
    {0x859d1502,0x64e6,0x4977,{0xa9,0xb3,0xd1,0x31,0xd6,0x8c,0x72,0x3b}};
extern "C" const GUID __declspec(selectany) CLSID_DelegatorClient =
    {0x90099fb2,0xb732,0x4cb7,{0x81,0xde,0xc3,0x5b,0xca,0x5e,0x15,0xdc}};
extern "C" const GUID __declspec(selectany) CLSID_ClientStateChangedEventData =
    {0x36522520,0x67a7,0x48ae,{0xaa,0x63,0xb5,0xa7,0xbe,0xa2,0x1b,0x7c}};
extern "C" const GUID __declspec(selectany) CLSID_LyncClientNotificationReceivedEventData =
    {0x861bfb12,0x94ce,0x4c1d,{0x9b,0x3a,0x33,0xb2,0x0c,0x9d,0x48,0xac}};
extern "C" const GUID __declspec(selectany) CLSID_CredentialRequestedEventData =
    {0xcbd9b3de,0x2ed2,0x47cc,{0xae,0x82,0x54,0xda,0xeb,0x26,0xeb,0xb7}};
extern "C" const GUID __declspec(selectany) CLSID_SignInDelayedEventData =
    {0x4883d56b,0x83f3,0x4cab,{0x81,0x61,0x0b,0x15,0x54,0x11,0x30,0x37}};
extern "C" const GUID __declspec(selectany) CLSID_PreferredCapabilitiesChangedEventData =
    {0xcec8bfbe,0x3e0e,0x4732,{0xb3,0xfe,0x5b,0xab,0xf6,0x17,0x0f,0x11}};
extern "C" const GUID __declspec(selectany) CLSID_DelegatorClientCollectionEventData =
    {0x37e5d2a9,0x55bd,0x42f8,{0x8b,0x46,0x3c,0x02,0x49,0x68,0x53,0xb0}};
extern "C" const GUID __declspec(selectany) IID_IGroupCollectionChangedEventData =
    {0xec9773ad,0x7d22,0x4b03,{0xb9,0x49,0x5e,0x90,0x55,0xaf,0x0f,0x19}};
extern "C" const GUID __declspec(selectany) IID_IContactCollection =
    {0x57c3cb4d,0xc794,0x456f,{0x84,0x6d,0x0e,0x1e,0xca,0x78,0xbd,0x00}};
extern "C" const GUID __declspec(selectany) IID_IGroup =
    {0xf97973d7,0x9436,0x4a3e,{0xa8,0x02,0xf6,0xea,0x17,0x2a,0x04,0x25}};
extern "C" const GUID __declspec(selectany) IID_IGroupCollection =
    {0x0c81ec90,0x9469,0x4ed4,{0xac,0xad,0xad,0xb8,0x6e,0x11,0x03,0x9c}};
extern "C" const GUID __declspec(selectany) IID_IDistributionGroup =
    {0x340c54a6,0x35cf,0x4971,{0xb5,0x40,0x72,0xd7,0xf0,0x40,0xac,0x24}};
extern "C" const GUID __declspec(selectany) IID_ICustomGroup =
    {0xd804ba85,0xdb0f,0x49ee,{0xb7,0x23,0xdc,0xdf,0x04,0x25,0xbd,0x9f}};
extern "C" const GUID __declspec(selectany) IID_IFavoriteContacts =
    {0x35ce3a44,0xc636,0x4d31,{0xbc,0x95,0xb9,0x37,0x0c,0x22,0x51,0xb9}};
extern "C" const GUID __declspec(selectany) IID_IContact =
    {0x85109845,0x6ae2,0x4b6b,{0x95,0x24,0x65,0xba,0x20,0x3e,0x5b,0x7d}};
extern "C" const GUID __declspec(selectany) IID_IContactSubscription =
    {0x95b8d3e8,0xf3d5,0x4dc0,{0xbc,0xfe,0xab,0x80,0xc8,0x35,0xde,0xa0}};
extern "C" const GUID __declspec(selectany) IID_IContactManager =
    {0xe6b3f5fa,0x4208,0x4c37,{0xae,0x07,0xf6,0x43,0x5b,0x68,0xd6,0x93}};
extern "C" const GUID __declspec(selectany) IID_IGroupMemberChangedEventData =
    {0x4b3a2412,0xc149,0x4e52,{0xa7,0x13,0x6c,0xeb,0xcf,0x47,0x50,0x3e}};
extern "C" const GUID __declspec(selectany) IID_IRecentContacts =
    {0xdcdea425,0xb5f2,0x4719,{0xa3,0xb4,0x69,0xff,0xb9,0x77,0x0b,0xe6}};
extern "C" const GUID __declspec(selectany) IID_IFrequentContacts =
    {0x18a9a79a,0xcbc7,0x4f8b,{0x8b,0x2b,0x9b,0x40,0xca,0x75,0xec,0x3c}};
extern "C" const GUID __declspec(selectany) IID_IContact2 =
    {0x10fdd9ba,0x0cba,0x4958,{0xb6,0xc8,0xd0,0x91,0x2b,0xf2,0x70,0x3f}};
extern "C" const GUID __declspec(selectany) IID_ISelf =
    {0x4da899d3,0xfac7,0x4f43,{0x8e,0x4a,0x19,0x87,0x98,0x8f,0xbc,0x9b}};
extern "C" const GUID __declspec(selectany) IID_IUtilities =
    {0x87084018,0x7e74,0x4f92,{0xab,0x39,0x1c,0xf0,0x31,0x88,0x58,0x0e}};
extern "C" const GUID __declspec(selectany) IID_IClient2 =
    {0x86225e5c,0x7595,0x4d1c,{0x98,0x5d,0x8a,0x04,0x58,0xd7,0x14,0xc0}};
extern "C" const GUID __declspec(selectany) IID_IConversationCollection =
    {0xa4e1d1e4,0x7839,0x473b,{0x95,0xfa,0x5d,0x33,0xdd,0xc0,0xd2,0xdd}};
extern "C" const GUID __declspec(selectany) IID_IConversation =
    {0x533891f1,0xa460,0x43e6,{0x8b,0xda,0x6d,0x0d,0x2c,0x13,0xad,0xde}};
extern "C" const GUID __declspec(selectany) IID_IConversationHistory =
    {0xf5996c71,0x77ab,0x4020,{0x97,0x76,0xc2,0xc2,0x70,0xfe,0xf2,0xae}};
extern "C" const GUID __declspec(selectany) IID_IConversationManager =
    {0xc31fb52d,0x9e9e,0x45ad,{0xa1,0x02,0x52,0x18,0xe1,0x68,0x5b,0x78}};
extern "C" const GUID __declspec(selectany) IID_ISummarizerStatesChangedEventData =
    {0x8057f99d,0x292c,0x4371,{0x92,0x6b,0x53,0x12,0xe6,0x1a,0x3a,0x40}};
extern "C" const GUID __declspec(selectany) IID_IConversationManagerEventData =
    {0xab2d1a67,0x2b3d,0x4e25,{0xb2,0x1c,0x03,0xf4,0xbf,0xc4,0xc2,0xbe}};
extern "C" const GUID __declspec(selectany) IID_IVoiceMail =
    {0x25a47938,0x5c31,0x4f43,{0x91,0x08,0xda,0x01,0xee,0x58,0x69,0xc0}};
extern "C" const GUID __declspec(selectany) IID_IVoiceMailCollection =
    {0xf644c610,0xa146,0x4a56,{0x83,0x38,0xa6,0x9c,0x45,0xc7,0x1c,0xef}};
extern "C" const GUID __declspec(selectany) IID_IVoiceMailBatchEventData =
    {0x18ba13c7,0xa64e,0x4301,{0xba,0x51,0xd1,0xbf,0xb3,0xc1,0xc9,0xd5}};
extern "C" const GUID __declspec(selectany) IID_IConversationWindow =
    {0xb1f5bcd8,0x69a6,0x4fc3,{0x9e,0xf0,0x9b,0xd4,0xad,0x99,0x90,0x61}};
extern "C" const GUID __declspec(selectany) IID_IAutomation =
    {0x2edaf003,0x14a6,0x4faf,{0x8a,0x72,0x99,0x52,0x6d,0x04,0x8d,0x63}};
extern "C" const GUID __declspec(selectany) IID_IClient =
    {0xee9f3e74,0xac61,0x469e,{0x80,0xd6,0xe2,0x2b,0xf4,0xee,0xff,0x5c}};
extern "C" const GUID __declspec(selectany) IID_ILyncClient =
    {0xc413a27f,0x601e,0x4261,{0xb9,0xe3,0x0e,0xb9,0x66,0xf4,0xe7,0xd1}};
extern "C" const GUID __declspec(selectany) IID_IDelegatorClient =
    {0x80edbf3a,0x812e,0x42b5,{0xa6,0x7f,0x6c,0xc6,0xd9,0xa1,0x9a,0x6c}};
extern "C" const GUID __declspec(selectany) IID_IDelegatorClientCollectionEventData =
    {0x744bab86,0xb15f,0x4bd5,{0xb7,0xdd,0x47,0x6b,0x0c,0xe0,0x16,0x2a}};
extern "C" const GUID __declspec(selectany) IID_IUCOfficeIntegrationHelper =
    {0x09028936,0xcd20,0x43ef,{0x98,0xcf,0x56,0x72,0x6a,0xdd,0xba,0x34}};
extern "C" const GUID __declspec(selectany) IID_IParticipant =
    {0x4875c050,0xba9c,0x4a05,{0x89,0x1e,0xe7,0xb0,0xa9,0x46,0x36,0x64}};
extern "C" const GUID __declspec(selectany) IID_IModality =
    {0xd4b240b9,0x059b,0x4d13,{0x89,0x97,0x31,0xb2,0x82,0xb3,0xbd,0x43}};
extern "C" const GUID __declspec(selectany) IID_IModalityDictionary =
    {0xed5040fa,0x8f8f,0x4876,{0xbf,0x6d,0x9c,0x62,0xf5,0xef,0xfd,0x5b}};
extern "C" const GUID __declspec(selectany) IID_IParticipantCollection =
    {0xfd02e775,0xcb9d,0x4ce3,{0x92,0x43,0x11,0x5c,0x81,0xe5,0xf1,0x1c}};
extern "C" const GUID __declspec(selectany) IID_IInstantMessageModality =
    {0x62a75516,0xc79b,0x42d7,{0x8b,0x49,0x3b,0xa4,0x92,0xc2,0xb3,0x85}};
extern "C" const GUID __declspec(selectany) IID_IChannel =
    {0x61c747f6,0x157c,0x4cae,{0xa3,0x7c,0xec,0x43,0x52,0x17,0x58,0x60}};
extern "C" const GUID __declspec(selectany) IID_IAudioChannel =
    {0x8e839af9,0x9508,0x475a,{0xaa,0x47,0x3e,0xe8,0xa9,0x46,0xb7,0x5e}};
extern "C" const GUID __declspec(selectany) IID_IVideoChannel =
    {0xe66ca1cc,0x9db0,0x467e,{0x9c,0x60,0x6a,0x83,0x2a,0xcd,0x67,0x80}};
extern "C" const GUID __declspec(selectany) IID_IAVModality =
    {0xe4e1d0af,0x5b70,0x45af,{0x9a,0x0e,0xf5,0x48,0xd8,0xfb,0xd1,0x7c}};
extern "C" const GUID __declspec(selectany) IID_IApplicationSharingModality =
    {0xac10d6e9,0x4a8c,0x4484,{0xb8,0xf4,0xca,0x1e,0x36,0x34,0x7a,0xae}};
extern "C" const GUID __declspec(selectany) IID_IVideoChannelCollection =
    {0x3cd1db65,0xd85b,0x4912,{0x85,0x36,0x6b,0xd8,0xe1,0x9f,0x17,0x38}};
extern "C" const GUID __declspec(selectany) IID_IParticipantCollectionChangedEventData =
    {0x064e6072,0x5592,0x4d13,{0xaf,0xdb,0x55,0xf5,0xa4,0x94,0xe1,0x37}};
extern "C" const GUID __declspec(selectany) IID_IStreamStateChangedEventData =
    {0xca7eeb7a,0x7dc3,0x4ffe,{0xa1,0x74,0x23,0xdb,0x5a,0x00,0x3c,0x04}};
extern "C" const GUID __declspec(selectany) IID_IControlRequestReceivedEventData =
    {0x32430741,0x330b,0x48df,{0x96,0xf6,0x0c,0xfe,0x00,0x48,0x4d,0x40}};
extern "C" const GUID __declspec(selectany) IID_IShareableContent =
    {0xb335ae5e,0xe4ea,0x49d3,{0xb0,0x3b,0x64,0x6a,0x96,0xfe,0x66,0xd3}};
extern "C" const GUID __declspec(selectany) IID_IShareableContentCollection =
    {0xa58f54d2,0x9786,0x4309,{0x96,0x4d,0x96,0x54,0x9a,0xec,0x76,0x11}};
extern "C" const GUID __declspec(selectany) IID_IContentSharingModality =
    {0x38e5ec21,0x81a1,0x45a7,{0x94,0xd6,0xb8,0x12,0x97,0x62,0x31,0xa4}};
extern "C" const GUID __declspec(selectany) IID_IContentCollectionChangedEventData =
    {0x17635563,0x2fa5,0x46ad,{0x8f,0x47,0x3f,0x0f,0xf9,0x58,0xec,0xe2}};
extern "C" const GUID __declspec(selectany) IID_IActivePresenterChangedEventData =
    {0x48b3099d,0x841d,0x4caa,{0x92,0x02,0x57,0x87,0x59,0x6e,0x2b,0xd2}};
extern "C" const GUID __declspec(selectany) IID_IActiveContentChangedEventData =
    {0x899b846a,0x8887,0x45d6,{0xb8,0xf5,0x8a,0xaa,0x0b,0xfd,0xf7,0x22}};
extern "C" const GUID __declspec(selectany) IID_IPowerPointContent =
    {0xaafe380b,0xb079,0x4b55,{0xb8,0x1b,0xa9,0x40,0x1d,0x77,0xb7,0x76}};
extern "C" const GUID __declspec(selectany) IID_IConversationContextCollection =
    {0x8a88c016,0xaa5d,0x4eae,{0x9c,0x79,0x7a,0xa5,0xc3,0x60,0xed,0x76}};
extern "C" const GUID __declspec(selectany) IID_IConversationContext =
    {0x04378c72,0xe58f,0x47c0,{0x86,0x21,0x90,0x1b,0x85,0xca,0x2e,0xd4}};
extern "C" const GUID __declspec(selectany) IID_IConversationContextList =
    {0x855ad949,0xe2eb,0x473d,{0xaa,0xf2,0x98,0x00,0x00,0x8e,0xa3,0x03}};
extern "C" const GUID __declspec(selectany) IID_IConversationContextCollectionEventData =
    {0xef5c00e4,0x75e2,0x4d39,{0xb3,0x58,0x0e,0xd9,0xcb,0x14,0xf3,0xcb}};
extern "C" const GUID __declspec(selectany) IID_IAccessPermission =
    {0xdd181e97,0xeee0,0x4d65,{0xba,0xd1,0x16,0x86,0x6d,0x0c,0x79,0x53}};
extern "C" const GUID __declspec(selectany) IID_IAccessEntry =
    {0x65b71572,0x5338,0x43c8,{0x9e,0x4c,0xf1,0xdf,0xc6,0x71,0x1a,0xd1}};
extern "C" const GUID __declspec(selectany) IID_IAccessEntryCollectionChangedEventData =
    {0xdd31376e,0xe529,0x43f2,{0xb3,0xb4,0xe3,0x79,0xeb,0xd2,0x8e,0x2a}};
extern "C" const GUID __declspec(selectany) CLSID_TangramUCOfficeIntegration =
    {0x424be3cd,0x34ab,0x4f51,{0x9c,0x57,0x43,0x41,0x16,0x6d,0xc8,0xfa}};
extern "C" const GUID __declspec(selectany) CLSID_UCOfficeIntegration =
	{0xa0651028,0xba7a,0x4d71,{0x87,0x7f,0x12,0xe0,0x17,0x5a,0x58,0x06 } };
} // namespace UCCollaborationLib
#pragma pack(pop)
