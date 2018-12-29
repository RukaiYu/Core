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

#pragma once
#include "Tangram.h"
#include "TangramClrProxy.h"
#include "../Tangram/chromium/ChromeProxy.h"
#include <map>

using namespace std;
using namespace System;
using namespace System::Windows;
using namespace System::Reflection;
using namespace System::Windows::Forms;
using namespace System::Runtime::InteropServices;
using namespace System::ComponentModel;
using namespace System::Collections;
using namespace System::Collections::Generic;
using namespace System::IO;
using namespace System::Threading;
using namespace System::Threading::Tasks;

using System::Runtime::InteropServices::Marshal;

extern CTangramCLRProxy theAppProxy;;
class CWndNodeCLREvent;
class CTangramNodeEvent;


namespace TangramCLR 
{
	//[Guid("54499D5E-AC2F-4F8B-9782-C00A9BB2F4E2")]
	//[ComVisibleAttribute(true)]
	//[InterfaceTypeAttribute(ComInterfaceType::InterfaceIsIDispatch)]
	//public interface class IManagerExtender
	//{
	//	[DispId(0x000001)]
	//	virtual void CloseForm(System::Object^ pForm);
	//	[DispId(0x000002)]
	//	virtual void OnCloseForm(long long hFormWnd);
	//};

	/// <summary>
	/// Summary for Tangram
	/// </summary>
	interface class ITangramApp;
	ref class WndNode;
	ref class WorkBenchWindow;
	ref class ChromeWebBrowser;

	public enum class WndNodeType
	{
		TNT_Blank			= 0x00000001,
		TNT_ActiveX			= 0x00000002,
		TNT_Splitter		= 0x00000004,
		TNT_Tabbed			= 0x00000008,

		TNT_CLR_Control		= 0x00000010,
		TNT_CLR_Form		= 0x00000020,
		TNT_CLR_Window		= 0x00000040,

		TNT_View			= 0x00000080		
	};

	template<typename T> 
	private ref class TangramBaseEnumerator : public System::Collections::IEnumerator
	{
	public:
		TangramBaseEnumerator(T^ v, int iCount)
		{
			m_pT = v;
			nCount = iCount;
			Reset();
		}

		virtual property Object^ Current
		{
			Object^ get()
			{
				if(index < 0 || index >= nCount)
				{
					throw gcnew InvalidOperationException();
				}
				return m_pT->default[index];
			}
		}

		virtual bool MoveNext()
		{
			index++;
			return (index < nCount);
		} 

		virtual void Reset()
		{
			index = -1;
		}

	private:
		int index;		
		long nCount;
		T^ m_pT;
	};

	public ref class TangramAppProxy
	{
	public:
		TangramAppProxy();
		~TangramAppProxy();

		void Init();

		CTangramAppProxy* m_pTangramAppProxy;
	};

	ref class WndPage;
	ref class WndNode;
	
	public ref class WndNodeCollection : public System::Collections::IEnumerable
	{
	public:
		WndNodeCollection(IWndNodeCollection* pNodes)
		{		
			SetNewNodeCollection(pNodes);
		};

		~WndNodeCollection()
		{
			SetNewNodeCollection(NULL);
		}
		
	private:
		IWndNodeCollection* m_pNodeCollection;

		void SetNewNodeCollection(IWndNodeCollection* pNodes)
		{
			if (m_pNodeCollection != NULL)
			{
				m_pNodeCollection->Release();
				m_pNodeCollection = NULL;
			}
			if (pNodes != NULL)
			{
				m_pNodeCollection = pNodes;
				m_pNodeCollection->AddRef();					
			}
		}
	public:
		virtual System::Collections::IEnumerator^ GetEnumerator()
		{
			return gcnew TangramBaseEnumerator<WndNodeCollection>(this,NodeCount);
		}
		
		property WndNode^ default[int]
		{
			WndNode^ get(int iIndex);
		}

		property int NodeCount
		{
			int get()
			{
				long n = 0;
				m_pNodeCollection->get_NodeCount(&n);
				return n;
			}
		}		
	};

	ref class WndFrame;
	public ref class WndNode
	{
	public:
		WndNode(IWndNode* pNode);
		~WndNode();

		IWndNode* m_pWndNode;
		Object^ m_pHostObj = nullptr;
		IMessageObj* m_pCurMsgObj;
		ChromePlus::CChromeBrowserProxy* m_pChromeBrowserProxy;
		CTangramNodeEvent* m_pTangramNodeEvent;
		CWndNodeCLREvent* m_pTangramNodeCLREvent;

		//virtual void CloseForm(Object^ pForm){};
		//virtual void OnCloseForm(long long hFormWnd){};

		delegate void NodeAddInCreated(WndNode^ sender, Object^ pAddIndisp, String^ bstrAddInID, String^ bstrAddInXml);
		event NodeAddInCreated^ OnNodeAddInCreated;
		void Fire_NodeAddInCreated(WndNode^ sender, Object^ pAddIndisp, String^ bstrAddInID, String^ bstrAddInXml)
		{
			OnNodeAddInCreated(sender, pAddIndisp, bstrAddInID, bstrAddInXml);
		}

		delegate void NodeAddInsCreated(WndNode^ sender);
		event NodeAddInsCreated^ OnNodeAddInsCreated;
		void Fire_NodeAddInsCreated(WndNode^ sender)
		{
			OnNodeAddInsCreated(sender);
		}

		delegate void ExtendComplete(WndNode^ sender);
		event ExtendComplete^ OnExtendComplete;
		void Fire_ExtendComplete(WndNode^ sender)
		{
			OnExtendComplete(sender);
		}

		delegate void Destroy(WndNode^ sender);
		event Destroy^ OnDestroy;
		void Fire_OnDestroy(WndNode^ sender)
		{
			OnDestroy(sender);
		}

		delegate void DocumentComplete(WndNode^ sender, Object^ pHtmlDoc, String^ strURL);
		event DocumentComplete^ OnDocumentComplete;
		void Fire_OnDocumentComplete(WndNode^ sender, Object^ pHtmlDoc, String^ strURL)
		{
			OnDocumentComplete(sender, pHtmlDoc, strURL);
		}

		delegate void TabChange(int nActivePage, int nOldActivePage);
		event TabChange^ OnTabChange;
		void Fire_OnTabChange(int nActivePage, int nOldActivePage)
		{
			OnTabChange(nActivePage, nOldActivePage);
		}

		delegate void MessageHandle(String^ arg1, String^ arg2);
		event MessageHandle^ OnMessageReceived;
		void Fire_OnMessageReceived(String^ arg1, String^ arg2)
		{
      OnMessageReceived(arg1, arg2);
		}
	private:
		HWND m_hWnd;
		Dictionary<String^, MethodInfo^>^	m_pTangramCLRMethodDic = nullptr;
		Dictionary<String^, Object^>^	m_pTangramPlugInDic = nullptr;
		void SetNewNode(IWndNode* pNode)
		{
			if (m_pWndNode != NULL)
			{
				m_pWndNode = NULL;
			}

			if (pNode != NULL)
			{
				m_pWndNode = pNode;
			}
		}

	public:
		void Init();
		//void SaveTangramDoc(String^ m_strName);
		Object^ ActiveMethod(String^ strMethod, cli::array<Object^, 1>^ p);
		//[DispId(0x000001)]
		//[ComVisibleAttribute(true)]
		//Object^ ExecScript(String^ strScript);
		WndNode^ Extend(String^ layerName, String^ layerXML);
		WndNode^ ExtendChild(int rowNum, int colName, String^ layerName, String^ layerXML);

		void OpenURL(String^ strURL, Disposition nDisposition, String^ strKey, String^ strXml);

		property String^ Caption
		{
			String^ get();
			void set(String^ strCaption)
			{
				m_pWndNode->put_Caption(STRING2BSTR(strCaption));
			}
		}

		property bool SaveToConfigFile
		{
			void set(bool value)
			{
				m_pWndNode->put_SaveToConfigFile(value);
			}
		}

		[BrowsableAttribute(false)]
		property WndPage^ WndPage
		{
			TangramCLR::WndPage^ get();
		}

		property String^ Name
		{
			String^ get()
			{
				if (m_pWndNode)
				{
					CComBSTR bstrCap("");
					m_pWndNode->get_Name(&bstrCap);
					return BSTR2STRING(bstrCap);
				}
				return "";
			}

			void set(String^ newVal)
			{
				if (m_pWndNode)
				{
					m_pWndNode->put_Name(STRING2BSTR(newVal));
				}
			}
		}

		property String^ URL
		{
			String^ get();
			void set(String^ newVal);
		}

		property String^ ExternalObjID
		{
			String^ get()
			{
				if (m_pWndNode)
				{
					CComBSTR bstrKey("");
					m_pWndNode->get_Attribute(CComBSTR("externalobjid"),&bstrKey);
					return BSTR2STRING(bstrKey);
				}
				return "";
			}

			void set(String^ newVal)
			{
				if (m_pWndNode)
				{
					ViewType nType;
					m_pWndNode->get_NodeType(&nType);
					if(nType==BlankView)
						m_pWndNode->put_Attribute(CComBSTR("externalobjid"), STRING2BSTR(newVal));
				}
			}
		}

		[BrowsableAttribute(false)]
		property String^ SenderSip
		{
			String^ get()
			{
				if (m_pCurMsgObj)
				{
					CComBSTR bstrCap("");
					m_pCurMsgObj->get_SenderSip(&bstrCap);
					return BSTR2STRING(bstrCap);
				}
				return "";
			}
		}

		[BrowsableAttribute(false)]
		property String^ TargetSip
		{
			String^ get()
			{
				if (m_pCurMsgObj)
				{
					CComBSTR bstrCap("");
					m_pCurMsgObj->get_TargetSip(&bstrCap);
					return BSTR2STRING(bstrCap);
				}
				return "";
			}
		}

		[BrowsableAttribute(false)]
		property String^ ReceivedMsg
		{
			String^ get()
			{
				if (m_pCurMsgObj)
				{
					CComBSTR bstrCap("");
					m_pCurMsgObj->get_ReceivedMsg(&bstrCap);
					return BSTR2STRING(bstrCap);
				}
				return "";
			}
		}

		[BrowsableAttribute(false)]
		property Object^ MsgSender
		{
			Object^ get()
			{
				if(m_pCurMsgObj)
				{
					CComBSTR bstrCap("");
					IDispatch* pDisp = nullptr;
					m_pCurMsgObj->get_Sender(&pDisp);
					CComVariant var(pDisp);
					return Marshal::GetObjectForNativeVariant((System::IntPtr)&var);;
				}
				return nullptr;
			}
		}

		[BrowsableAttribute(false)]
		property Object^ XObject
		{
			Object^ get()
			{
				if (m_pHostObj != nullptr)
					return m_pHostObj;

				VARIANT var;
				m_pWndNode->get_XObject(&var);
				
				try
				{
					m_pHostObj = Marshal::GetObjectForNativeVariant((System::IntPtr)&var);
				}
				catch (InvalidOleVariantTypeException^ e)
				{
					e->Message;
				}
				catch (...)
				{

				}
				return m_pHostObj;
			}
		}

		property System::Drawing::Color rgbMiddle
		{
			System::Drawing::Color get()
			{
				OLE_COLOR rgb;
				m_pWndNode->get_rgbMiddle(&rgb);
				return System::Drawing::ColorTranslator::FromOle(rgb);
			}
			void set(System::Drawing::Color newVal)
			{
				m_pWndNode->put_rgbMiddle(System::Drawing::ColorTranslator::ToOle(newVal));
			}
		}

		property System::Drawing::Color rgbLeftTop
		{
			System::Drawing::Color get()
			{
				OLE_COLOR rgb;
				m_pWndNode->get_rgbLeftTop(&rgb);
				return System::Drawing::ColorTranslator::FromOle(rgb);
			}
			void set(System::Drawing::Color newVal)
			{
				m_pWndNode->put_rgbLeftTop(System::Drawing::ColorTranslator::ToOle(newVal));
			}
		}

		property System::Drawing::Color rgbRightBottom
		{
			System::Drawing::Color get()
			{
				OLE_COLOR rgb;
				m_pWndNode->get_rgbRightBottom(&rgb);
				return System::Drawing::ColorTranslator::FromOle(rgb);
			}
			void set(System::Drawing::Color newVal)
			{
				m_pWndNode->put_rgbRightBottom(System::Drawing::ColorTranslator::ToOle(newVal));
			}
		}

		[BrowsableAttribute(false)]
		property Object^ Extender
		{
			Object^ get()
			{
				Object^ pRetObject = nullptr;
				IDispatch* pDisp = NULL;
				m_pWndNode->get_Extender(&pDisp);
				if (pDisp)
					pRetObject = Marshal::GetObjectForIUnknown((IntPtr)pDisp); 
				return pRetObject;
			}
			void set(Object^ obj)
			{
				IDispatch* pDisp = (IDispatch*)Marshal::GetIUnknownForObject(obj).ToPointer();
				m_pWndNode->put_Extender(pDisp);
			}
		}

		property Object^ Tag
		{
			Object^ get()
			{
				Object^ pRetObject = nullptr;
				VARIANT var;
				m_pWndNode->get_Tag(&var);	
				try
				{
					pRetObject = Marshal::GetObjectForNativeVariant((System::IntPtr)&var);
				}
				catch (InvalidOleVariantTypeException^ )
				{

				}
				catch (NotSupportedException^ )
				{

				}
				return pRetObject;
			}

			void set(Object^ obj)
			{
				try
				{
					VARIANT var;
					Marshal::GetNativeVariantForObject(obj,(System::IntPtr)&var);
					m_pWndNode->put_Tag(var);
				}
				catch (ArgumentException^ e)
				{
					e->Data->ToString();
				}
			}
		}

		[BrowsableAttribute(false)]
		property int ActivePage
		{
			int get()
			{
				int nPage = 0;
				m_pWndNode->get_ActivePage(&nPage);	
				
				return nPage;
			}

			void set(int nPage)
			{
				m_pWndNode->put_ActivePage(nPage);
			}
		}

		[BrowsableAttribute(false)]
		property WndFrame^ Frame
		{
			WndFrame^ get()
			{
				CComPtr<IWndFrame> pTangramFrame;
				m_pWndNode->get_Frame(&pTangramFrame);

				WndFrame^ pFrame = theAppProxy._createObject<IWndFrame, WndFrame>(pTangramFrame);
				return pFrame;
			}
		}

		[BrowsableAttribute(false)]
		property WndFrame^ HostFrame
		{
			WndFrame^ get()
			{
				CComPtr<IWndFrame> pTangramFrame;
				m_pWndNode->get_HostFrame(&pTangramFrame);
				if (pTangramFrame)
				{
					WndFrame^ pFrame = theAppProxy._createObject<IWndFrame, WndFrame>(pTangramFrame);
					return pFrame;
				}
				return nullptr;
			}
		}

		[BrowsableAttribute(false)]
		property Object^ PlugIn[String^]
		{
			Object^ get(String^ strName);
		}

		[BrowsableAttribute(false)]
		property WndNodeCollection^ ChildNodes
		{
			WndNodeCollection^ get();
		}

		[BrowsableAttribute(false)]
		property WndNodeCollection^ Objects[WndNodeType]
		{
			WndNodeCollection^ get(WndNodeType nValue)
			{
				CComPtr<IWndNodeCollection> pNodes = NULL;
				m_pWndNode->get_Objects((long)nValue,&pNodes);
				return gcnew WndNodeCollection(pNodes);				
			}
		}

		[BrowsableAttribute(false)]
		property WndNode^ RootNode
		{
			WndNode^ get()
			{
				CComPtr<IWndNode> pRootNode;
				m_pWndNode->get_RootNode(&pRootNode);

				return theAppProxy._createObject<IWndNode, WndNode>(pRootNode);
			}
		}

		[BrowsableAttribute(false)]
		property WndNode^ ParentNode
		{
			WndNode^ get()
			{
				CComPtr<IWndNode> pRootNode;
				m_pWndNode->get_ParentNode(&pRootNode);

				return theAppProxy._createObject<IWndNode, WndNode>(pRootNode);
			}
		}

		[BrowsableAttribute(false)]
		property WndNode^ HostNode
		{
			WndNode^ get()
			{
				CComPtr<IWndNode> pNode;
				m_pWndNode->get_HostNode(&pNode);

				return theAppProxy._createObject<IWndNode, WndNode>(pNode);
			}
		}

		[BrowsableAttribute(false)]
		property int Col
		{
			int get()
			{
				long nValue = 0;
				m_pWndNode->get_Col(&nValue);
				return nValue;
			}
		}

		[BrowsableAttribute(false)]
		property int Row
		{
			int get()
			{
				long nValue = 0;
				m_pWndNode->get_Row(&nValue);
				return nValue;
			}
		}

		property int Hmin
		{
			int get()
			{
				int nValue = 0;
				m_pWndNode->get_Hmin(&nValue);
				return nValue;
			}
			void set(int newVal)
			{
				m_pWndNode->put_Hmin(newVal);
			}
		}

		property int Hmax
		{
			int get()
			{
				int nValue = 0;
				m_pWndNode->get_Hmax(&nValue);
				return nValue;
			}
			void set(int newVal)
			{
				m_pWndNode->put_Hmax(newVal);
			}
		}

		property int Vmin
		{
			int get()
			{
				int nValue = 0;
				m_pWndNode->get_Vmin(&nValue);
				return nValue;
			}
			void set(int newVal)
			{
				m_pWndNode->put_Vmin(newVal);
			}
		}

		property int Vmax
		{
			int get()
			{
				int nValue = 0;
				m_pWndNode->get_Vmax(&nValue);
				return nValue;
			}
			void set(int newVal)
			{
				m_pWndNode->put_Vmax(newVal);
			}
		}

		property int Rows
		{
			int get()
			{
				long nValue = 0;
				m_pWndNode->get_Rows(&nValue);
				return nValue;
			}
		}

		property int Cols
		{
			int get()
			{
				long nValue = 0;
				m_pWndNode->get_Cols(&nValue);
				return nValue;
			}
		}

		property WndNodeType NodeType
		{
			WndNodeType get()
			{
				ViewType type;
				m_pWndNode->get_NodeType(&type);
				return (TangramCLR::WndNodeType)type;
			}
		}

		property __int64 Handle
		{
			__int64 get()
			{
				if (m_hWnd)
					return (__int64)m_hWnd;
				__int64 h = 0;
				m_pWndNode->get_Handle(&h);
				m_hWnd = (HWND)h;
				return h;
			}
		}

		[BrowsableAttribute(false)]
		property String^ Attribute[String^]
		{
			String^ get(String^ strKey)
			{
				BSTR bstrVal;
				m_pWndNode->get_Attribute(STRING2BSTR(strKey),&bstrVal);

				return BSTR2STRING(bstrVal);
			}

			void set(String^ strKey, String^ strVal)
			{
				m_pWndNode->put_Attribute(
					STRING2BSTR(strKey),
					STRING2BSTR(strVal));
			}
		}

public:
	WndNode^ GetNode(int nRow, int nCol)
	{
		IWndNode* pNode;		
		m_pWndNode->GetNode(nRow,nCol,&pNode);

		WndNode^ tNode = theAppProxy._createObject<IWndNode, WndNode>(pNode);

		return tNode;
	}

	int GetNodes(String^ strName, WndNode^% pNode, WndNodeCollection^% pNodeCollection)
	{
		IWndNodeCollection* pNodes;
		IWndNode* pTNode;

		pNode = nullptr;
		pNodeCollection = nullptr;

		long nCount;
		m_pWndNode->GetNodes(STRING2BSTR(strName),
			&pTNode,&pNodes,&nCount);

		pNode = theAppProxy._createObject<IWndNode, WndNode>(pTNode);

		pNodeCollection = gcnew WndNodeCollection(pNodes);
		
		pNodes->Release();

		return nCount;
	}

	void SendUCMAMessage(String^ strSipFrom, String^ strSipTo, String^ strMsg);
	//void Unload()
	//{
	//	m_pWndNode->Unload();
	//}

	private:
		/// <summary>
		/// Required designer variable.
		/// </summary>

		WndNodeCollection^ m_pChildNodes;
	};

	public ref class WndFrame : public Dictionary<String^, WndNode^>
	{
	public:
		WndFrame(IWndFrame* pFrame)
		{
			m_pWndFrame = pFrame;
		}
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		~WndFrame()
		{
			LONGLONG nValue = (LONGLONG)m_pWndFrame;
			theAppProxy._removeObject(nValue);
		}

		IWndFrame* m_pWndFrame;

	public:
		void SendToHostNode(String^ arg1, String^ arg2);
		WndNode^ Extend(String^  layerName, String^ layerXML);
		void Attach(bool bAttach)
		{
			if(m_pWndFrame)
			{
				if(bAttach)
				{
					m_pWndFrame->Attach();
				}
				else
				{
					m_pWndFrame->Detach();
				}
			}
		}

		property Object^ FrameData[String^]
		{
			Object^ get(String^ iIndex);
			void set(String^ iIndex, Object^ newVal);
		}
			
		property String^ Name
		{
			String^ get()
			{
				CComBSTR bstrName("");
				m_pWndFrame->get_Name(&bstrName);
				return BSTR2STRING(bstrName);
			}
		}
			
		property WndNode^ Node[Object^]
		{
			WndNode^ get(Object^ index)
			{
				VARIANT var;
				Marshal::GetNativeVariantForObject(index,(System::IntPtr)&var);
				IWndNode* pNode = nullptr;
				m_pWndFrame->get_WndNode(var, &pNode);
				return theAppProxy._createObject<IWndNode, WndNode>(pNode);
			}
		}
			
		property WndNodeCollection^ RootNodes
		{
			WndNodeCollection^ get()
			{
				CComPtr<IWndNodeCollection> pNodes = NULL;
				m_pWndFrame->get_RootNodes(&pNodes);

				return gcnew WndNodeCollection(pNodes.p);
			}
		}
			
		property WndPage^ Page
		{
			WndPage^ get()
			{
				CComPtr<IWndPage> pPage = NULL;
				m_pWndFrame->get_WndPage(&pPage);

				return theAppProxy._createObject<IWndPage, WndPage>(pPage);
			}
		}
	};

	//[ComSourceInterfacesAttribute(TangramCLR::IManagerExtender::typeid)]	 map<gcroot<>, IWndNode*>	m_mapFrameworkElement;

	public ref class Tangram
	{
	public:
		Tangram();
		Tangram(ITangram* pTangram);
		~Tangram();
	private:
		ITangram*										m_pHostTangram;
		static Tangram	^								m_pManager;
		static Dictionary<String^, MethodInfo^>^		m_pTangramCLRMethodDic = gcnew Dictionary<String^, MethodInfo^>();
		static Dictionary<String^, Type^>^				m_pTangramCLRTypeDic = gcnew Dictionary<String^, Type^>();
		static Dictionary<String^, Type^>^				m_pAppFormTypeDic = nullptr;
		static Dictionary<String^, Type^>^				m_pAppMDIFormTypeDic = nullptr;
		static Dictionary<Control^, String^>^			m_pControlRelationDic = nullptr;
	public:
		static Dictionary<String^, Object^>^			m_pTangramCLRObjDic = gcnew Dictionary<String^, Object^>();
		static Dictionary<String^, TangramAppProxy^>^	m_pTangramAppProxyDic = gcnew Dictionary<String^, TangramAppProxy^>();
		static Dictionary<Object^, WndNode^>^			m_pFrameworkElementDic = gcnew Dictionary<Object^, WndNode^>();
		static Func<Object^, String^>^					m_pTangramInitGlobalTask;
		static String^ ComputeHash(String^ source);
		static String^ TangramGlobalFunc(Object^);
		static AutoResetEvent^	m_pAppQuitEvent = nullptr;
		//static String^ TangramInitFunc(Object^);
		//static AutoResetEvent^	m_pAppFormTypeDicEvent = nullptr;
		static AutoResetEvent^	m_pPlatformStartupCompletedEvent = nullptr;
		static WndPage^ CreateWndPage(Control^ ctrl,Object^ ExternalObj);
		static Object^ CreateObject(String^ ObjID);
		static Type^ GetType(String^ ObjID);
		static Object^ ActiveMethod(String^ strObjID, String^ strMethod, cli::array<Object^, 1>^ p);
		static Object^ ActiveObjectMethod(Object^ pObj, String^ strMethod, cli::array<Object^, 1>^ p);
		static void StartApplication(String^ appID, String^ strXML);
		static WorkBenchWindow^ RemoteActiveWorkBenchWindow(String^ appID);
		static Control^ GetMDIClient(Form^ pForm);
		//static Dictionary<String^, Type^>^ InitAppFormTypeDic();
		static ChromeWebBrowser^ ActiveBrowser();
		
		TangramCLR::ITangramApp^		m_pTangramAppProxy;

		static void SetControlRelation(Control^ ctrl, String^ strTypes);
		static Tangram^ GetTangram()
		{
			 if(m_pManager==nullptr)
				 m_pManager = gcnew Tangram();
			 return m_pManager;
		}

		//static Tangram^ GetRemoteTangram(String^ strAppID)
		//{
		//	if (m_pManager == nullptr)
		//		m_pManager = gcnew Tangram();
		//	return m_pManager;
		//}

		static void InitEclipse()
		{
			 if(m_pManager&&m_pManager->m_pHostTangram)
				 m_pManager->m_pHostTangram->InitEclipseApp();
		}

		static WndNode^ GetWndNodeFromObj(Object^ obj)
		{
			WndNode^ pNode = nullptr;
			if (m_pFrameworkElementDic->TryGetValue(obj, pNode))
			{
				return pNode;
			}
			return nullptr;
		}

		static String^ GetObjAssemblyName(Object^ obj)
		{
			 if(obj==nullptr)
				 return L"";
			 Assembly^ a = nullptr;
			 String^ strName = L"";
			 try
			 {
				a = Assembly::GetAssembly(obj->GetType());
			 }
			 catch (System::ArgumentNullException^)
			 {

			 }
			 finally
			 {
				 if (a != nullptr)
				 {
					 strName = a->FullName;
					 strName = strName->Substring(0, strName->IndexOf(","));
				 }
			 }
			 return strName;
		}

		static String^ LoadCLRResource(Object^ destObj, String^ resourceName)
		{
			 if(destObj ==nullptr)
				 return L"";
			 if (resourceName == L"" || resourceName == nullptr)
         resourceName = L"tangramresource.xml";
			 Assembly^ a = nullptr;
			 String^ strName = L"";
			 try
			 {
				a = Assembly::GetAssembly(destObj->GetType());
			 }
			 catch (ArgumentNullException^)
			 {

			 }
			 finally
			 {
				 if (a != nullptr)
				 {
					 strName = a->FullName;
					 strName = strName->Substring(0, strName->IndexOf(","));
				 }
			 }

			 System::IO::Stream^ sm = nullptr;
			 try
			 {
				 sm = a->GetManifestResourceStream(strName + "." + resourceName);
			 }
			 catch (...)
			 {

			 }
			 finally
			 {
				 if (sm != nullptr)
				 {
					 cli::array<byte,1>^ bs = gcnew cli::array<byte,1>(sm->Length);
					 sm->Read(bs, 0, (int)sm->Length);
					 sm->Close();

					 System::Text::UTF8Encoding^ con = gcnew System::Text::UTF8Encoding();

					 strName = con->GetString(bs);
				 }
			 }
			 return strName;
		}

		static property Object^ Application 
		{
			Object^ get();
			void set(Object^ obj);
		}

		static property String^ CurrentDesigningTangramXml
		{
			String^ get();
		}

		static property WndNode^ CreatingNode 
		{
			WndNode^ get();
		}

		static property WndFrame^ ActiveEclipseTopFrame 
		{
			WndFrame^ get();
		}

		//static property Dictionary<String^, Type^>^ FormTypeDictionary
		//{
		//	Dictionary<String^, Type^>^ get()
		//	{
		//		if (TangramCLR::Tangram::m_pAppFormTypeDicEvent != nullptr)
		//			TangramCLR::Tangram::m_pAppFormTypeDicEvent->WaitOne();
		//		return m_pAppFormTypeDic;
		//	}
		//}

		//static property Dictionary<String^, Type^>^ MDIFormTypeDictionary
		//{
		//	Dictionary<String^, Type^>^ get()
		//	{
		//		if (TangramCLR::Tangram::m_pAppFormTypeDicEvent != nullptr)
		//			TangramCLR::Tangram::m_pAppFormTypeDicEvent->WaitOne();
		//		return m_pAppMDIFormTypeDic;
		//	}
		//}

		static property Dictionary<String^, TangramAppProxy^>^ TangramAppProxyDictionary
		{
			Dictionary<String^, TangramAppProxy^>^ get()
			{
				return m_pTangramAppProxyDic;
			}
		}

		static property bool RemoteAppConnected[String^]
		{
			bool get(String^ strAppID);
		}

		static property TangramAppProxy^ AppProxy[String^]
		{
			TangramAppProxy^ get(String^ strAppID);
		}

		static property WorkBenchWindow^ ActiveWorkBenchWindow
		{
			WorkBenchWindow^ get();
		}

		delegate void Close();
		event Close^ OnClose;
		void Fire_OnClose();

		delegate void ExtendComplete(IntPtr hWnd, String^ bstrUrl, WndNode^ pRootNode);
		event ExtendComplete^ OnExtendComplete;
		void Fire_OnExtendComplete(IntPtr hWnd, String^ bstrUrl, WndNode^ pRootNode)
		{
			OnExtendComplete(hWnd, bstrUrl, pRootNode);
		}

		//delegate void DocumentComplete(Object^ Doc, String^ bstrUrl);
		//event DocumentComplete^ OnDocumentComplete;
		//void Fire_OnDocumentComplete(Object^ Doc, String^ bstrUrl)
		//{
		//	OnDocumentComplete(Doc, bstrUrl);
		//}

		property String^ AppKeyValue[String^]
		{
			String^ get(String^ iIndex);
			void set(String^ iIndex, String^ newVal);
		}

#ifdef _WIN64
#ifdef TANGRAMCOLLABORATION
		static String^ ExportData(String^ strPathName, bool bRecursive);
		static void ImportData(String^ strPathName, String^ strData);
		static void AddOrganizationUnit(String^ strPathName, String^ strOrgName);
		static void AddUser(String^ strPathName, String^ strDisplayName, String^ strUsrName);
#endif
#endif
};

	/// <summary>
	/// WndPage 
	/// </summary>
	//[ComSourceInterfacesAttribute(TangramCLR::IManagerExtender::typeid)]
	public ref class WndPage : public IWin32Window, public Dictionary<String^, WndFrame^>
	{
	public:
		WndPage(void);
		WndPage(IWndPage* pPage);

		IWndPage*			m_pPage;
		property IntPtr Handle
		{
			virtual IntPtr get()
			{
				LONGLONG h = 0;
				m_pPage->get_Handle(&h);
				return (IntPtr)h;
			};
		}

		property String^ URL
		{
			String^ get()
			{
				String^ strRet = nullptr;
				CComBSTR bstr(L"");
				m_pPage->get_URL(&bstr);
				strRet = BSTR2STRING(bstr); 
				return strRet;
			}
			void set(String^ newVal)
			{
				CComBSTR bstr = STRING2BSTR(newVal);
				m_pPage->put_URL(bstr);
			}
		}

		property Object^ External
		{
			void set(Object^ newVal)
			{
				if (newVal != nullptr)
				{
					IDispatch* pDisp = (IDispatch*)Marshal::GetIUnknownForObject(newVal).ToPointer();
					if (pDisp)
						m_pPage->put_External(pDisp);
				}
			}
		}

		property String^ FrameName[Control^]
		{
			String^ get(Control^ ctrl)
			{
				if (ctrl != nullptr)
				{
					LONGLONG hWnd = ctrl->Handle.ToInt64();
					BSTR bstrName = ::SysAllocString(L"");
					m_pPage->get_FrameName(hWnd, &bstrName);
					String^ strRet = BSTR2STRING(bstrName);
					::SysFreeString(bstrName);
					return strRet;
				}
				return String::Empty;
			}
		}

		property String^ FrameName[IntPtr]
		{
			String^ get(IntPtr handle)
			{
				if (::IsWindow((HWND)handle.ToInt64()))
				{
					BSTR bstrName = ::SysAllocString(L"");
					m_pPage->get_FrameName(handle.ToInt32(), &bstrName);
					String^ strRet = BSTR2STRING(bstrName);
					::SysFreeString(bstrName);
					return strRet;
				}
				return String::Empty;
			}
		}

		property Object^ Extender[String^]
		{
			Object^ get(String^ strName)
			{
				BSTR bstrName = STRING2BSTR(strName);
				CComPtr<IDispatch> pDisp;
				m_pPage->get_Extender(bstrName, &pDisp);
				::SysFreeString(bstrName);
				return Marshal::GetObjectForIUnknown((IntPtr)pDisp.p);
			}

			void set(String^ strName, Object^ newObj)
			{
				IDispatch* pDisp = (IDispatch*)Marshal::GetIDispatchForObject(newObj).ToPointer();
				m_pPage->put_Extender(STRING2BSTR(strName), pDisp);
			}
		}

		property WndFrame^ Frames[Object^]
		{
			WndFrame^ get(Object^ obj)
			{
				CComVariant m_v;
				IntPtr handle = (IntPtr)&m_v;
				Marshal::GetNativeVariantForObject(obj, handle);
				CComPtr<IWndFrame> pFrame;
				m_pPage->get_Frame(m_v, &pFrame);
				if (pFrame)
				{
					return theAppProxy._createObject<IWndFrame, WndFrame>(pFrame);
				}
				return nullptr;
			}
		}

		WndNode^ GetWndNode(String^ bstrXml, String^ bstrFrameName);

		void AddObjToHtml(String^ strObjName, bool bConnectEvent, Object^ pObj);

		WndFrame^ CreateFrame(Control^ ctrl, String^ strName)
		{
			WndFrame^ pRetFrame = nullptr;
			if (ctrl != nullptr&&String::IsNullOrEmpty(strName) == false)
			{
				if (TryGetValue(strName, pRetFrame))
				{
					return pRetFrame;
				}
				CComPtr<IWndFrame> pFrame;
				m_pPage->CreateFrame(CComVariant(0), CComVariant((LONGLONG)ctrl->Handle.ToInt64()), STRING2BSTR(strName), &pFrame);
				if (pFrame)
				{
					pRetFrame = theAppProxy._createObject<IWndFrame, WndFrame>(pFrame);
					if (pRetFrame != nullptr)
						Add(strName, pRetFrame);
				}
			}
			return pRetFrame;
		}

		WndFrame^ CreateFrame(IntPtr handle,String^ strName)
		{
			WndFrame^ pRetFrame = nullptr;
			if (::IsWindow((HWND)handle.ToInt64()) && String::IsNullOrEmpty(strName) == false)
			{
				if (TryGetValue(strName, pRetFrame))
				{
					return pRetFrame;
				}
				CComPtr<IWndFrame> pFrame;
				m_pPage->CreateFrame(CComVariant(0), CComVariant((LONGLONG)handle.ToInt64()), STRING2BSTR(strName), &pFrame);
				if (pFrame)
				{
					pRetFrame = theAppProxy._createObject<IWndFrame, WndFrame>(pFrame);
					if (pRetFrame != nullptr)
						Add(strName, pRetFrame);
				}
			}
			return pRetFrame;
		}
		
		void ExtendFrames(String^ strFrames, String^ strKey, String^ bstrXml, bool bSaveToConfigFile);

		delegate void DocumentComplete(WndPage^ sender, Object^ pHtmlDoc, String^ strURL);
		event DocumentComplete^ OnDocumentComplete;
		void Fire_OnDocumentComplete(WndPage^ sender, Object^ pHtmlDoc, String^ strURL)
		{
			OnDocumentComplete(sender, pHtmlDoc, strURL);
		}

		delegate void IPCMsg(WndFrame^ sender, String^ strType, String^ strContent, String^ strFeature);
		event IPCMsg^ OnIPCMsg;
		void Fire_OnIPCMsg(WndFrame^ sender, String^ strType, String^ strContent, String^ strFeature)
		{
			OnIPCMsg(sender, strType, strContent, strFeature);
		}

		delegate void Destroy(WndNode^ sender);
		event Destroy^ OnDestroy;
		void Fire_OnDestroy(WndNode^ sender)
		{
			OnDestroy(sender);
		}

		delegate void TabChange(WndNode^ sender, int nActivePage, int nOldActivePage);
		event TabChange^ OnTabChange;
		void Fire_OnTabChange(WndNode^ sender, int nActivePage, int nOldActivePage)
		{
			OnTabChange(sender, nActivePage, nOldActivePage);
		}

		delegate void PageLoad(WndPage^ sender);
		event PageLoad^ OnPageLoad;
		void Fire_OnPageLoad(WndPage^ sender)
		{
			OnPageLoad(sender);
		}

	protected:
		CWndPageClrEvent*	m_pTangramClrEvent;

		~WndPage();
	}; 

	public ref class EclipseView
	{
	public:
		IEclipseCtrl* m_pCtrl;
		EclipseView(IEclipseCtrl* pCtrl)
		{
			m_pCtrl = pCtrl;
		};

		property IntPtr Handle
		{
			virtual IntPtr get()
			{
				LONGLONG h = 0;
				m_pCtrl->get_HWND(&h);
				return (IntPtr)h;
			};
		}

		WndNode^ Extend(String^ strFrameName, String^ strKey, String^ strXml)
		{
			if (m_pCtrl)
			{
				IWndNode* pNode = nullptr;
				m_pCtrl->Extend(STRING2BSTR(strFrameName), STRING2BSTR(strKey), STRING2BSTR(strXml), &pNode);
				if (pNode)
				{
					return theAppProxy._createObject<IWndNode, WndNode>(pNode);
				}
			}
			return nullptr;
		}
			//property String^ FrameName[Control^]
		//{
		//	String^ get(Control^ ctrl)
		//	{
		//		if (ctrl != nullptr)
		//		{
		//			long hWnd = ctrl->Handle.ToInt64();
		//			BSTR bstrName = ::SysAllocString(L"");
		//			m_pPage->get_FrameName(hWnd, &bstrName);
		//			String^ strRet = BSTR2STRING(bstrName);
		//			::SysFreeString(bstrName);
		//			return strRet;
		//		}
		//		return String::Empty;
		//	}
		//}

		//property String^ FrameName[IntPtr]
		//{
		//	String^ get(IntPtr handle)
		//	{
		//		if (::IsWindow((HWND)handle.ToInt64()))
		//		{
		//			BSTR bstrName = ::SysAllocString(L"");
		//			m_pPage->get_FrameName(handle.ToInt32(), &bstrName);
		//			String^ strRet = BSTR2STRING(bstrName);
		//			::SysFreeString(bstrName);
		//			return strRet;
		//		}
		//		return String::Empty;
		//	}
		//}


		//property WndFrame^ Frames[Object^]
		//{
		//	WndFrame^ get(Object^ obj)
		//	{
		//		CComVariant m_v;
		//		IntPtr handle = (IntPtr)&m_v;
		//		Marshal::GetNativeVariantForObject(obj, handle);
		//		CComPtr<IWndFrame> pFrame;
		//		m_pPage->get_Frame(m_v, &pFrame);
		//		if (pFrame)
		//		{
		//			return theAppProxy._createObject<IWndFrame, WndFrame>(pFrame);
		//		}
		//		return nullptr;
		//	}
		//}

		//WndNode^ GetWndNode(String^ bstrXml, String^ bstrFrameName);

		//void AddObjToHtml(String^ strObjName, bool bConnectEvent, Object^ pObj);

		//WndFrame^ CreateFrame(Control^ ctrl, String^ strName)
		//{
		//	if (ctrl != nullptr&&String::IsNullOrEmpty(strName) == false)
		//	{
		//		CComPtr<IWndFrame> pFrame;
		//		m_pPage->CreateFrame(CComVariant(0), CComVariant(ctrl->Handle.ToInt64()), STRING2BSTR(strName), &pFrame);
		//		if (pFrame)
		//		{
		//			return theAppProxy._createObject<IWndFrame, WndFrame>(pFrame);
		//		}
		//	}
		//	return nullptr;
		//}

		//WndFrame^ CreateFrame(IntPtr handle,String^ strName)
		//{
		//	if (::IsWindow((HWND)handle.ToInt64()) && String::IsNullOrEmpty(strName) == false)
		//	{
		//		CComPtr<IWndFrame> pFrame;
		//		m_pPage->CreateFrame(CComVariant(0), CComVariant(handle.ToInt64()), STRING2BSTR(strName), &pFrame);
		//		if (pFrame)
		//		{
		//			return theAppProxy._createObject<IWndFrame, WndFrame>(pFrame);
		//		}
		//	}
		//	return nullptr;
		//}

		//delegate void DocumentComplete(WndPage^ sender, Object^ pHtmlDoc, String^ strURL);
		//event DocumentComplete^ OnDocumentComplete;
		//void Fire_OnDocumentComplete(WndPage^ sender, Object^ pHtmlDoc, String^ strURL)
		//{
		//	OnDocumentComplete(sender, pHtmlDoc, strURL);
		//}
		//delegate void Destroy(WndNode^ sender);
		//event Destroy^ OnDestroy;
		//void Fire_OnDestroy(WndNode^ sender)
		//{
		//	OnDestroy(sender);
		//}
	//protected:
	//	IWndPage*			m_pPage;
	//	CWndPageClrEvent*	m_pTangramClrEvent;

	//	~WndPage();
	}; 

	public ref class WorkBenchWindow : public System::Collections::IEnumerable
	{
	public:
		WorkBenchWindow(IWorkBenchWindow* pWorkBenchWindow)
		{
			m_pWorkBenchWindow = pWorkBenchWindow;
			//SetNewNodeCollection(pNodes);
		};

		~WorkBenchWindow()
		{
			//SetNewNodeCollection(NULL);
		}

		void Active()
		{
			if (m_pWorkBenchWindow)
			{
				LONGLONG h = 0;
				m_pWorkBenchWindow->get_Handle(&h);
				HWND hwnd = (HWND)h;
				hwnd = ::GetAncestor(hwnd, GA_ROOT);
				::SetFocus(hwnd);
			}
		}

		WndNode^ SWTExtend(String^ strKey, String^ strXml)
		{
			if (m_pWorkBenchWindow)
			{
				IWndNode* pNode = nullptr;
				m_pWorkBenchWindow->Extend(STRING2BSTR(strKey), STRING2BSTR(strXml), &pNode);
				if(pNode)
					return theAppProxy._createObject<IWndNode, WndNode>(pNode);
			}
			return nullptr;
		}

		WndNode^ SWTExtendInView(int nView, String^ strKey, String^ strXml)
		{
			if (m_pWorkBenchWindow)
			{
				IEclipseCtrl* pCtrl = nullptr;
				m_pWorkBenchWindow->get_Ctrl(CComVariant((int)nView), &pCtrl);
				if (pCtrl)
				{
					IWndNode* pNode = nullptr;
					pCtrl->Extend(CComBSTR(L"EclipseView"), STRING2BSTR(strKey), STRING2BSTR(strXml), &pNode);
					if(pNode)
						return theAppProxy._createObject<IWndNode, WndNode>(pNode);
				}

			}
			return nullptr;
		}

	private:
		IWorkBenchWindow* m_pWorkBenchWindow;

	public:
		virtual System::Collections::IEnumerator^ GetEnumerator()
		{
			return gcnew TangramBaseEnumerator<WorkBenchWindow>(this, NodeCount);
		}

		void CloseTangramUI()
		{
			if (m_pWorkBenchWindow)
				m_pWorkBenchWindow->CloseTangramUI();
		}

		property EclipseView^ default[int]
		{
			EclipseView^ get(int iIndex)
			{
				if (m_pWorkBenchWindow)
				{
					long nCount = 0;
					m_pWorkBenchWindow->get_Count(&nCount);
					if (iIndex >= nCount)
						return nullptr;
					else
					{
						IEclipseCtrl* pCtrl = nullptr;
						m_pWorkBenchWindow->get_Ctrl(CComVariant((long)iIndex), &pCtrl);
						if (pCtrl)
						{
							return theAppProxy._createObject<IEclipseCtrl, EclipseView>(pCtrl);
						}
					}
				}
				return nullptr;
			};
		}

		property int NodeCount
		{
			int get()
			{
				long n = 0;
				m_pWorkBenchWindow->get_Count(&n);
				return n;
			}
		}
	};

	public ref class TangramChromeObjProxy
	{
	public:
		TangramChromeObjProxy() 
		{
			m_pWndNode = nullptr;
		};

		virtual ~TangramChromeObjProxy() 
		{
		};

		WndNode^	m_pWndNode;
		virtual void Create(IntPtr nHandle) {};
		virtual void OnTangramFunction(cli::array<Object^, 1>^ arguments) {};
		virtual void OnJSExtInvokeHandler(String^ strName, cli::array<Object^, 1>^ arguments) {};
	};

	public interface class ITangramEventObjProxy
	{
	};

	public interface class ITangramWndNodeProxy
	{
		virtual void OnExtendComplete();
		virtual void OnDestroy();
		virtual void OnNodeAddInCreated(Object^ pAddIndisp, String^ bstrAddInID, String^ bstrAddInXml);
		virtual void OnNodeAddInsCreated();
		virtual void OnNodeDocumentComplete(Object^ ExtenderDisp, String^ bstrURL);
		virtual void OnControlNotify(WndNode^ sender, LONG NotifyCode, LONG CtrlID, IntPtr CtrlHandle, String^ CtrlClassName);
		virtual void OnTabChange(LONG ActivePage, LONG OldPage);
	};

	public interface class ITangramWndPageProxy
	{
		virtual void OnPageLoaded(WndPage^ sender, String^ url);
		virtual void OnNodeCreated(WndNode^ pNodeCreated);
		virtual void OnAddInCreated(WndNode^ pRootNode, Object^ pAddIn, String^ strID, String^ strAddInXml);
		virtual void OnBeforeExtendXml(String^ bstrXml, IntPtr hWnd);
		virtual void OnExtendXmlComplete(String^ bstrXml, IntPtr hWnd, WndNode^ pRetRootNode);
		virtual void OnDestroy();
		virtual void OnNodeMouseActivate(WndNode^ pActiveNode);
		virtual void OnClrControlCreated(WndNode^ Node, Control^ Ctrl, String^ CtrlName, IntPtr CtrlHandle);
		virtual void OnTabChange(WndNode^ sender, LONG ActivePage, LONG OldPage);
		virtual void OnEvent(Object^ sender, Object^ EventArg);
		virtual void OnControlNotify(WndNode^ sender, LONG NotifyCode, LONG CtrlID, IntPtr CtrlHandle, String^ CtrlClassName);
		virtual void OnTangramEvent(ITangramEventObjProxy^ NotifyObj);
	};

	public interface class ITangramWndFrameProxy
	{
		virtual void OnExtend(WndNode^ pRetNode, String^ bstrKey, String^ bstrXml);
	};

	public interface class ITangramApp
	{
		[DispId(0x000001)]
		virtual void TangramClose();
		[DispId(0x000002)]
		virtual ITangramWndNodeProxy^ OnTangramNodeInit(WndNode^ pNewNode);
		[DispId(0x000003)]
		virtual ITangramWndFrameProxy^ OnWndFrameCreated(WndFrame^ pNewFrame);
		[DispId(0x000004)]
		virtual ITangramWndPageProxy^ OnWndPageCreated(WndPage^ pNewPage);
		[DispId(0x000005)]
		virtual void OnExtendComplete(IntPtr hWnd, String^ bstrUrl, WndNode^ pRootNode);
		[DispId(0x000006)]
		virtual void OnTangramEvent(ITangramEventObjProxy^ NotifyObj);
	};
}
