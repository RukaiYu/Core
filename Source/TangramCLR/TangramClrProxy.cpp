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
********************************************************************************/

// HostExtender.cpp : Implementation of CTangramNavigator

#include "stdafx.h"
#include "dteinternal.h"
#include "dllmain.h" 
#include "TangramCLRHost.h"
#include "TangramNodeCLREvent.h"
#include "ErrorCtrl.h"
#include "TangramClrProxy.h"
#ifdef TANGRAMCOLLABORATION
#include "TangramUcma.h"
#endif

//#include <shellapi.h>
//#include <shlobj.h>

using namespace TangramCLR;

typedef HRESULT (__stdcall *TangramCLRCreateInstance)(REFCLSID clsid, REFIID riid, /*iid_is(riid)*/ LPVOID *ppInterface);

CTangramCLRProxy theAppProxy;

CTangramCLRProxy::CTangramCLRProxy():CApplicationCLRProxyImpl()
{
	//if (theApp.m_bRegisterServer)
	//	return;
	//_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF | _CRTDBG_LEAK_CHECK_DF);
	//_CrtSetBreakAlloc(308);
	//_CrtSetBreakAlloc(306);System.Windows.Forms.GroupBox
	//_CrtSetBreakAlloc(298);
	//m_strExtendableTypes = L"|System.Windows.Forms.WebBrowser|System.Windows.Forms.Panel|System.Windows.Forms.TreeView|System.Windows.Forms.ListView|System.Windows.Forms.MonthCalendar|System.Windows.Forms.TabPage|System.Windows.Forms.TabControl|System.Windows.Forms.GroupBox|System.Windows.Forms.FlowLayoutPanel|System.Windows.Forms.TableLayoutPanel|System.Windows.Forms.SplitContainer|";
	m_strExtendableTypes = L"|WebBrowser|Panel|TreeView|ListView|MonthCalendar|TabPage|TabControl|GroupBox|FlowLayoutPanel|TableLayoutPanel|SplitContainer|";
	m_bHostApp = false;
	System::Windows::Forms::Application::EnableVisualStyles();
	m_pPropertyGrid = nullptr;
	m_pSystemAssembly = nullptr;
	m_pOnLoad = nullptr;
	m_pOnMdiChildActivate = nullptr;
	m_pOnCtrlVisible = nullptr;
	m_pDefaultPageType = nullptr;
	m_htObjects = gcnew Hashtable();
	m_pTangramProxy = gcnew TangramProxy();
	System::Windows::Forms::Application::ApplicationExit += gcnew EventHandler(&OnApplicationExit);

	if (::GetModuleHandle(_T("TangramCore.dll")) == nullptr)
	{
		m_bHostApp = true;
		CComPtr<ITangram> pTangram;
		pTangram.CoCreateInstance(CComBSTR("Tangram.Tangram.1"));
		theApp.m_pTangram = pTangram.Detach();
		if (theApp.m_pTangram)
		{
			theApp.m_pTangram->put_AppKeyValue(CComBSTR(L"CLRProxy"), CComVariant((LONGLONG)this));
			theApp.m_pTangram->put_AppKeyValue(CComBSTR(L"CLRAppProxy"), CComVariant((LONGLONG)static_cast<CTangramAppProxy*>(&theApp)));
			ITangramExtender* pExtender = nullptr;
			theApp.m_pTangram->get_Extender(&pExtender);
			if (pExtender)
			{
				CComQIPtr<IVSExtender> pVSExtender(pExtender);
				if (pVSExtender)
					theApp.m_pVSExtender = pVSExtender.Detach();
			}

	#ifdef TANGRAMCOLLABORATION
			CComVariant m_v;
			if (theApp.m_pTangram->get_AppKeyValue(CComBSTR(L"collaborationscript"), &m_v) == S_OK)
			{
				m_strCollaborationScript = OLE2T(m_v.bstrVal);
				if (m_strCollaborationScript!=_T(""))
					Collaboration::TangramUcmaApp::TangramVerb(BSTR2STRING(m_strCollaborationScript), L"");
			}
	#endif
		}
	}
}

CTangramCLRProxy::~CTangramCLRProxy()
{
	if(TangramCLR::Tangram::m_pAppQuitEvent !=nullptr)
		TangramCLR::Tangram::m_pAppQuitEvent->WaitOne();
	if (m_strCollaborationScript!=_T(""))
	{
#ifdef TANGRAMCOLLABORATION
		Collaboration::TangramUcmaApp::TangramVerb(L"<tangram  Init='StopUcma'/>", L"");
#endif
	}
	for (auto it : m_mapFrameInfo)
	{
		delete it.second;
	}
	TangramCLR::Tangram^ pManager = TangramCLR::Tangram::GetTangram();
	delete pManager;
	if (m_pTangramProxy)
	{
		delete m_pTangramProxy;
	}

	if (m_pProxy)
	{
		m_pProxy->m_pCLRProxy = nullptr;
	}

	ATLTRACE(_T("Release CTangramCLRProxy :%p\n"), this);
	if(m_bHostApp&&::GetModuleHandle(_T("comdlg32.dll")))
		::FreeLibrary(::GetModuleHandle(_T("TangramCore.dll")));
}

void CTangramCLRProxy::AttachVSPropertyWnd(HWND hVSPropertyGrid)
{ 
	m_pPropertyGrid = (PropertyGrid^)(Control^)Control::FromHandle((IntPtr)hVSPropertyGrid);
	if (m_pPropertyGrid)
	{
		if (m_pPropertyGrid->SelectedObject != nullptr)
		{
			OnSelectedObjectsChanged(nullptr, nullptr);
		}
		m_pPropertyGrid->SelectedObjectsChanged += gcnew System::EventHandler(&OnSelectedObjectsChanged);
	}
}

void CTangramCLRProxy::AttachCLRObjEvent(IDispatch* Sender, WindowEventType nType, HWND hNotifyWnd, VARIANT_BOOL bAttachEvent)
{
	Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)Sender);
	if (pCtrl != nullptr)
	{
		HWND hWnd = (HWND)pCtrl->Handle.ToInt64();
		if (bAttachEvent)
		{
		}
		else
		{

		}
	}
}

BSTR CTangramCLRProxy::AttachObjEvent(IDispatch* EventObj, IDispatch* SourceObj, WindowEventType EventType, IDispatch* HTMLWindow)
{
	BSTR bstrRet = L"";
	Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)SourceObj);
	if (pCtrl != nullptr)
	{
		bstrRet = STRING2BSTR(pCtrl->Name);
		LONGLONG nIndex = (LONGLONG)HTMLWindow;
		HWND hWnd = (HWND)pCtrl->Handle.ToInt64();
		ObjectEventInfo* pInfo = (ObjectEventInfo*)::GetWindowLongPtr(hWnd, GWLP_USERDATA);
		if (pInfo == nullptr)
		{
			m_pTangramProxy->AttachHandleDestroyedEvent(pCtrl);
			pInfo = new ObjectEventInfo();
			::SetWindowLongPtr(hWnd, GWLP_USERDATA, (LONG_PTR)pInfo);
		}
		CComQIPtr<IEventProxy> pTangramEvent(EventObj);
		if (pTangramEvent)
		{
			if (pInfo)
			{
				ObjectEventInfo2* pObjectEventInfo2 = nullptr;
				auto it2 = pInfo->m_mapEventObj2.find(nIndex);
				if (it2 == pInfo->m_mapEventObj2.end())
				{
					pObjectEventInfo2 = new ObjectEventInfo2();
					pInfo->m_mapEventObj2[nIndex] = pObjectEventInfo2;
				}
				else
					pObjectEventInfo2 = it2->second;
				auto it = pObjectEventInfo2->m_mapEventObj.find(EventType);
				if (it == pObjectEventInfo2->m_mapEventObj.end())
				{
					pObjectEventInfo2->m_mapEventObj[EventType] = pTangramEvent.p;
					m_pTangramProxy->AttachEvent(pCtrl, EventType);
				}
			}
		}
	}
	return bstrRet;
}

void CTangramCLRProxy::OnDestroyChromeBrowser(IChromeWebBrowser* pChromeWebBrowser) 
{
	auto it = m_mapChromeWebBrowser.find(pChromeWebBrowser);
	if (it != theAppProxy.m_mapChromeWebBrowser.end())
	{
		it->second->m_pChromeWebBrowserHost = nullptr;
		theAppProxy.m_mapChromeWebBrowser.erase(it);
		//delete it->second;
	}
};

CTangramWPFObj* CTangramCLRProxy::CreateWPFControl(IWndNode* pNode, HWND hPWnd, UINT nID)
{
	if (pNode)
	{
		CComBSTR bstrCnnID(L"");
		pNode->get_Attribute(CComBSTR(L"cnnid"), &bstrCnnID);
		CTangramWPFObjWrapper* pWpfControlWrapper = nullptr;
		Type^ pType = TangramCLR::Tangram::GetType(BSTR2STRING(bstrCnnID));
		::SysFreeString(bstrCnnID);
		if (pType == nullptr)
		{
			pType = m_pDefaultPageType;
			if (pType == nullptr)
			{
				m_pDefaultPageType = TangramCLR::Tangram::GetType(L"tangramclrhelper.tangramwpfctrl, tangramclrhelper, version=1.0.0.0, culture=neutral, publickeytoken=83d96735ac82c5df");
				pType = m_pDefaultPageType;
			}
		}

		try
		{
			pWpfControlWrapper = new CTangramWPFObjWrapper();
			if (pWpfControlWrapper->CreateControl(pType, hPWnd, WS_CHILD | WS_VISIBLE, 0, 0, 0, 0))
			{
				WndNode^ _pNode = (WndNode^)theAppProxy._createObject<IWndNode, TangramCLR::WndNode>(pNode);
				TangramCLR::Tangram::m_pFrameworkElementDic[pWpfControlWrapper->m_pUIElement] = _pNode;
				::SetWindowLongPtr(pWpfControlWrapper->m_hwndWPF, GWLP_ID, nID);
				m_mapWpfControlWrapper[pWpfControlWrapper->m_hwndWPF] = pWpfControlWrapper;
				try {
					pWpfControlWrapper->m_pSource->RootVisual = pWpfControlWrapper->m_pUIElement;
				}
				catch (System::Windows::Markup::XamlParseException^ e)
				{
					Debug::WriteLine(L"Tangram WPFControlWrapper Exception 1: " + e->Message);
					Debug::WriteLine(L"Tangram WPFControlWrapper Exception 1: " + e->InnerException->Message);
				}
			}
		}
		catch (System::Exception^ ex)
		{
			Debug::WriteLine(L"Tangram WPFControlWrapper Exception 1: " + ex->Message);
			Debug::WriteLine(L"Tangram WPFControlWrapper Exception 1: " + ex->InnerException->Message);
		}

		if(pWpfControlWrapper!=nullptr)
			return pWpfControlWrapper;
	}
	return nullptr; 
}

HRESULT CTangramCLRProxy::NavigateURL(IWndNode* pNode, CString strURL, IDispatch* dispObjforScript)
{ 
	return S_FALSE; 
}

void CTangramCLRProxy::WindowCreated(LPCTSTR strClassName, LPCTSTR strName, HWND hPWnd, HWND hWnd)
{
	auto it = m_mapForm.find(hPWnd);
	if (it == m_mapForm.end())
	{
		Control^  pPForm = Form::FromHandle((IntPtr)hPWnd);
		if (pPForm!=nullptr)
		{
			if (IsWinForm(hPWnd))
			{
				Form^ _pForm = static_cast<Form^>(pPForm);
				auto it = m_mapForm.find(hPWnd);
				if (it == m_mapForm.end())
				{
					m_mapForm[hPWnd] = _pForm;
					if(m_pOnLoad)
					{ 
					}
					else
						m_pOnLoad = gcnew EventHandler(CTangramCLRProxy::OnLoad);
					_pForm->Load += m_pOnLoad;
				}
			}
		}
	}
	CString strClsName = _T("Tangram Window Class");
}

void CTangramCLRProxy::WindowDestroy(HWND hWnd) 
{
	auto it = m_mapForm.find(hWnd);
	if (it != m_mapForm.end())
	{
		m_mapForm.erase(it);
		if (hWnd == m_hCLRMainWnd)
		{
			if (::IsWindow(m_hMsgWnd))
				::DestroyWindow(m_hMsgWnd);
			System::Windows::Forms::Application::Exit();
			PostQuitMessage(0);
		}
	}
	auto it2 = this->m_mapWpfControlWrapper.find(hWnd);
	if (it2 != m_mapWpfControlWrapper.end())
	{
		delete it2->second;
		m_mapWpfControlWrapper.erase(it2);
	}
	auto it3 = m_mapFrameInfo.find(hWnd);
	if (it3 != m_mapFrameInfo.end())
	{
		delete it3->second;
		m_mapFrameInfo.erase(it3);
	}
}

void CTangramCLRProxy::OnVisibleChanged(System::Object ^sender, System::EventArgs ^e)
{
	Control^ pCtrl = (Control^)sender;
	if (pCtrl->Visible)
	{
		BSTR bstrName = STRING2BSTR(pCtrl->Name);
		theAppProxy.m_pProxy->ExtendFrame((HWND)pCtrl->Handle.ToInt64(), OLE2T(bstrName), _T("default"));
		::SysFreeString(bstrName);
	}
}

void CTangramCLRProxy::OnItemSelectionChanged(System::Object ^sender, System::Windows::Forms::ListViewItemSelectionChangedEventArgs ^e)
{
	if (e->Item->Tag != nullptr)
	{
		String^ strTag = e->Item->Tag->ToString();
		if (String::IsNullOrEmpty(strTag) == false)
		{
			if (strTag->IndexOf(L"|TangramNode|") != -1)
			{
				String^ strIndex = strTag->Substring(strTag->IndexOf(L":") + 1);
				if (String::IsNullOrEmpty(strIndex) == false)
				{
					Control^ pCtrl = (Control^)sender;
					Control^ pTop = pCtrl->TopLevelControl;
					Type^ pType = pTop->GetType();
					if (pType->IsSubclassOf(Form::typeid))
					{
						String^ name = pType->Name + pCtrl->Name;
						theAppProxy.m_pProxy->ExtendCtrl(pCtrl->Handle.ToInt64(), name, strIndex);
					}
					pCtrl->Select();
				}
			}
		}
	}
}

Object^ CTangramCLRProxy::InitTangramCtrl(Form^ pForm, Control^ pCtrl, bool bSave)
{
	WndPage^ pWndPage = nullptr;
	if (pCtrl&&pForm)
	{
		BSTR bstrIndex = STRING2BSTR(pForm->GetType()->FullName);
		CString strIndex = OLE2T(bstrIndex);
		for each (Control^ pChild in pCtrl->Controls)
		{
			Type^ pType = pChild->GetType();
			String^ strType = pType->FullName;
			if (strType->IndexOf(L"System.Drawing") == 0)
				break;
			String^ strType2 = strType->Replace(L"System.Windows.Forms.", L"");
			if ((m_strExtendableTypes->IndexOf(L"|" + strType2 + L"|") != -1 && pChild->Dock == DockStyle::None) || pChild->Dock == DockStyle::Fill)
			{
				bool bExtendable = false;// (pChild->Tag == nullptr);
				if (pChild->Tag != nullptr)
				{
					String^ strTag = pChild->Tag->ToString();
					bExtendable = (strTag->IndexOf(L"|Extendable|") >= 0);
					//bExtendable = String::IsNullOrEmpty(strTag);
					//if (bExtendable == false)
				}
				if (bExtendable)
				{
					if (pWndPage == nullptr)
					{
						auto it = theAppProxy.m_mapConfigPage.find(strIndex);
						if (it == theAppProxy.m_mapConfigPage.end())
						{
							pWndPage = TangramCLR::Tangram::CreateWndPage(pForm, nullptr);
							pWndPage->m_pPage->put_ConfigName(strIndex.AllocSysString());
							theAppProxy.m_mapConfigPage[strIndex] = pWndPage;
						}
						else
						{
							Object^ pObj = it->second;
							pWndPage = static_cast<WndPage^>(pObj);
						}
					}
					if (m_pOnCtrlVisible)
					{
					}
					else
					{
						m_pOnCtrlVisible = gcnew EventHandler(CTangramCLRProxy::OnVisibleChanged);
					}
					pChild->VisibleChanged += m_pOnCtrlVisible;
					if (strType == L"System.Windows.Forms.MdiClient")
					{
						if (m_pOnMdiChildActivate)
						{
						}
						else
						{
							m_pOnMdiChildActivate = gcnew EventHandler(CTangramCLRProxy::OnMdiChildActivate);
						}
						pForm->MdiChildActivate += m_pOnMdiChildActivate;
					}
					else if (strType == L"System.Windows.Forms.TreeView")
					{
						TreeView^ pTreeView = (TreeView^)pChild;
						pTreeView->AfterSelect += gcnew System::Windows::Forms::TreeViewEventHandler(&OnAfterSelect);
					}
					else if (strType == L"System.Windows.Forms.ListView")
					{
						ListView^ pListView = (ListView^)pChild;
						pListView->ItemSelectionChanged += gcnew System::Windows::Forms::ListViewItemSelectionChangedEventHandler(&OnItemSelectionChanged);
					}

					String^ name = pChild->Name;
					if (strType == L"System.Windows.Forms.MdiClient")
					{
						name = "MdiClient";
					}
					else if (String::IsNullOrEmpty(name))
						name = strType;
					BSTR strName = STRING2BSTR(name);
					WndFrameInfo* pInfo = new WndFrameInfo;
					pInfo->m_hCtrlHandle = (HWND)pChild->Handle.ToInt64();
					m_mapFrameInfo[pInfo->m_hCtrlHandle] = pInfo;
					pInfo->m_strCtrlName = pChild->Name;
					pInfo->m_strParentCtrlName = pCtrl->Name;
					IWndFrame* _pFrame = m_pProxy->ConnectPage((HWND)pChild->Handle.ToInt64(), OLE2T(strName), pWndPage->m_pPage, pInfo);
					::SysFreeString(strName);
				}
			}
			if (pType->IsSubclassOf(UserControl::typeid) == false)
				InitTangramCtrl(pForm, pChild,bSave);
		}
		::SysFreeString(bstrIndex);
	}
	return pWndPage;
}

void CTangramCLRProxy::OnAfterSelect(System::Object ^sender, System::Windows::Forms::TreeViewEventArgs ^e)
{
	if (e->Node->Tag != nullptr)
	{
		String^ strTag = e->Node->Tag->ToString();
		if (String::IsNullOrEmpty(strTag) == false)
		{
			if (strTag->IndexOf(L"|TangramNode|") != -1)
			{
				String^ strIndex = strTag->Substring(strTag->IndexOf(L":") + 1);
				if (String::IsNullOrEmpty(strIndex) == false)
				{
					Control^ pCtrl = (Control^)sender;
					Control^ pTop = pCtrl->TopLevelControl;
					Type^ pType = pTop->GetType();
					if (pType->IsSubclassOf(Form::typeid))
					{
						String^ name = pType->Name + pCtrl->Name;
						theAppProxy.m_pProxy->ExtendCtrl(pCtrl->Handle.ToInt64(), name, strIndex);
					}
					pCtrl->Select();
				}
			}
		}
	}
}

Object^ CTangramCLRProxy::InitTangramNode(IWndNode* _pNode, Control^ pCtrl, bool bSave)
{
	if (::IsChild(m_pProxy->m_hHostWnd, (HWND)pCtrl->Handle.ToInt64()))
		return nullptr;
	WndPage^ pWndPage = nullptr;
	WndNode^ pNode = (WndNode^)theAppProxy._createObject<IWndNode, TangramCLR::WndNode>(_pNode);
	if (pNode)
	{
		pWndPage = pNode->WndPage;
		IWndPage* pPage = pWndPage->m_pPage;
		for each (Control^ pChild in pCtrl->Controls)
		{
			Type^ pType = pChild->GetType();
			String^ strType = pType->FullName;
			if (strType->IndexOf(L"System.Drawing") == 0)
				break;
			String^ strType2 = strType->Replace(L"System.Windows.Forms.", L"");
			if ((m_strExtendableTypes->IndexOf(L"|" + strType2 + L"|") != -1 && pChild->Dock == DockStyle::None) || pChild->Dock == DockStyle::Fill)
			{
				bool bExtendable = false;// (pChild->Tag == nullptr);
				if (pChild->Tag != nullptr)
				{
					String^ strTag = pChild->Tag->ToString();
					bExtendable = (strTag->IndexOf(L"|Extendable|") >= 0);
					//bExtendable = String::IsNullOrEmpty(strTag);
					//if (bExtendable == false)
				}
				if (bExtendable)
				{
					IWndFrame* pFrame = nullptr;
					_pNode->get_Frame(&pFrame);
					CComBSTR bstrName("");
					pFrame->get_Name(&bstrName);
					String^ name = pNode->Name + L".";
					if (String::IsNullOrEmpty(pChild->Name))
						name += strType;
					else
						name += pChild->Name;
					name += L"." + BSTR2STRING(bstrName);
					BSTR strName = STRING2BSTR(name);
					WndFrameInfo* pInfo = new WndFrameInfo;
					pInfo->m_hCtrlHandle = (HWND)pChild->Handle.ToInt64();
					m_mapFrameInfo[pInfo->m_hCtrlHandle] = pInfo;
					pInfo->m_strCtrlName = pChild->Name;
					pInfo->m_strParentCtrlName = pCtrl->Name;
					IWndFrame* _pFrame = m_pProxy->ConnectPage((HWND)pChild->Handle.ToInt64(), OLE2T(strName), pWndPage->m_pPage, pInfo);
					if (m_pOnCtrlVisible)
					{
					}
					else
					{
						m_pOnCtrlVisible = gcnew EventHandler(CTangramCLRProxy::OnVisibleChanged);
					}
					pChild->VisibleChanged += m_pOnCtrlVisible;
					if (strType == L"System.Windows.Forms.TreeView")
					{
						TreeView^ pTreeView = (TreeView^)pChild;
						pTreeView->AfterSelect += gcnew TreeViewEventHandler(&OnAfterSelect);
					}
					else if (strType == L"System.Windows.Forms.ListView")
					{
						ListView^ pListView = (ListView^)pChild;
						pListView->ItemSelectionChanged += gcnew ListViewItemSelectionChangedEventHandler(&OnItemSelectionChanged);
					}

					::SysFreeString(strName);
				}
			}
			InitTangramNode(_pNode, pChild, bSave);
		}
	}

	return pWndPage;
}

void CTangramCLRProxy::OnMdiChildActivate(System::Object ^sender, System::EventArgs ^e)
{
	Form^ pForm = static_cast<Form^>(sender);
	String^ strKey = L"";
	if (pForm->ActiveMdiChild != nullptr)
	{
		strKey = pForm->ActiveMdiChild->GetType()->FullName;
		Object^ objTag = pForm->Tag;
		if (objTag != nullptr)
		{
			String^ strTag = objTag->ToString();
			if (String::IsNullOrEmpty(strTag)==false)
			{
				int nIndex = strTag->IndexOf("|");
				if (nIndex != -1)
				{
					String^ strKey2 = strTag->Substring(0, nIndex);
					if (String::IsNullOrEmpty(strKey2) == false)
					{
						strKey += L"_";
						strKey += strKey2;
					}
				}
			}
		}
	}
	theApp.m_pTangram->ExtendFrames(pForm->Handle.ToInt64(), CComBSTR(L""), STRING2BSTR(strKey), CComBSTR(L""), true);
}

void CTangramCLRProxy::OnLoad(System::Object ^sender, System::EventArgs ^e)
{
	Form^ pForm = static_cast<Form^>(sender);
	WndPage^ pWndPage = static_cast<WndPage^>(theAppProxy.InitTangramCtrl(pForm, pForm, true));
	if (pWndPage)
		pWndPage->Fire_OnPageLoad(pWndPage);
	Control^ ctrl = TangramCLR::Tangram::GetMDIClient(pForm);
	if (ctrl != nullptr)
	{
		Form^ pForm2 = pForm->ActiveMdiChild;
		if (pForm2 != nullptr)
		{
			String^ strKey = pForm2->GetType()->FullName;
			Object^ objTag = pForm2->Tag;
			if (objTag != nullptr)
			{
				String^ strTag = objTag->ToString();
				if (String::IsNullOrEmpty(strTag) == false)
				{
					int nIndex = strTag->IndexOf("|");
					if (nIndex != -1)
					{
						String^ strKey2 = strTag->Substring(0, nIndex);
						if (String::IsNullOrEmpty(strKey2) == false)
						{
							strKey += L"_";
							strKey += strKey2;
						}
					}
				}
			}
			theApp.m_pTangram->ExtendFrames(pForm->Handle.ToInt64(), CComBSTR(L""), STRING2BSTR(strKey), CComBSTR(L""), true);
		}
	}
	theAppProxy.m_mapConfigPage.clear();
	pForm->Load -= theAppProxy.m_pOnLoad;
}

void CTangramCLRProxy::OnCLRHostExit() 
{
	System::Windows::Forms::Application::Exit();
}

HRESULT CTangramCLRProxy::ActiveCLRMethod(BSTR bstrObjID, BSTR bstrMethod, BSTR bstrParam, BSTR bstrData)
{
	String^ strObjID = BSTR2STRING(bstrObjID);
	String^ strMethod = BSTR2STRING(bstrMethod);
	String^ strData = BSTR2STRING(bstrData);
#ifdef TANGRAMCOLLABORATION	
	if (m_strCollaborationScript!=_T("")&&String::Compare(L"ucmamsg", strObjID, true) == 0)
	{
		if (strMethod == L"SendMessage")
		{
			CString strSips = OLE2T(bstrParam);
			int nPos = strSips.Find(_T("|"));
			if (nPos != -1)
			{
				CString strSipFrom = strSips.Left(nPos);
				strSipFrom.MakeLower();
				CString strSipTo = strSips.Mid(nPos + 1);
				String^ _strSipFrom = BSTR2STRING(strSipFrom);
				String^ _strSipTo = BSTR2STRING(strSipTo);
				if (String::IsNullOrEmpty(_strSipTo) == false)
				{
					Collaboration::TangramEndpoint^ pTangramEndpoint = Collaboration::TangramUcmaApp::GetEndpoint(_strSipFrom);
					if (pTangramEndpoint != nullptr)
					{
						pTangramEndpoint->SendIMMessage(_strSipTo, 0, strData);
					}
				}
			}
		}
		return S_OK;
	}
#endif
	cli::array<Object^, 1>^ pObjs = { BSTR2STRING(bstrParam), BSTR2STRING(bstrData) };
	TangramCLR::Tangram::ActiveMethod(strObjID, strMethod, pObjs);
	return S_OK;
}
//
//Type^ CTangramCLRProxy::GetTypeFromAssemblyQualifiedName(CString strAssemblyQualifiedName)
//{
//	//wpfcontrollibrary1.page1, wpfcontrollibrary1, version=1.0.0.0, culture=neutral, publickeytoken=null
//	//"tangramclrhelper.tangramwpfctrl, tangramclrhelper, version=1.0.0.0, culture=neutral, publickeytoken=9f80b1630127f641"	
//	strAssemblyQualifiedName.MakeLower();
//	int nPos = strAssemblyQualifiedName.Find(_T(","));
//	if (nPos != -1)
//	{
//		CString strName = strAssemblyQualifiedName.Left(nPos);
//		CString strName2 = strAssemblyQualifiedName.Mid(nPos + 1);
//		strName2.Trim();
//		CString strLib = _T("");
//		nPos = strName2.Find(_T(","));
//		if (nPos == -1)
//		{
//			strLib = strName2;
//			strLib += _T(".dll");
//		}
//		else
//		{
//			TCHAR m_szBuffer[MAX_PATH];
//			CString strPath = _T("");
//			HRESULT hr = SHGetFolderPath(NULL, CSIDL_WINDOWS, NULL, 0, m_szBuffer);
//			strPath.Format(_T("%s\\Microsoft.NET\\assembly\\GAC_MSIL\\TangramCLR\\v4.0_1.0.1992.1963__1bcc94f26a4807a7\\TangramCLR.dll"), m_szBuffer);
//#ifdef _WIN64
//			strPath.Format(_T("%s\\Microsoft.NET\\assembly\\GAC_%d\\TangramCLR\\v4.0_1.0.1992.1963__1bcc94f26a4807a7\\TangramCLR.dll"), m_szBuffer, 64);
//#else
//			strPath.Format(_T("%s\\Microsoft.NET\\assembly\\GAC_%d\\TangramCLR\\v4.0_1.0.1992.1963__1bcc94f26a4807a7\\TangramCLR.dll"), m_szBuffer, 32);
//#endif
//		}
//	}
//	return nullptr;
//}

IDispatch* CTangramCLRProxy::CreateCLRObj(BSTR bstrObjID)
{
	Object^ pObj = TangramCLR::Tangram::CreateObject(BSTR2STRING(bstrObjID));
	::SysFreeString(bstrObjID);
	if (pObj != nullptr)
	{
		if (pObj->GetType()->IsSubclassOf(Form::typeid))
		{
			Form^ thisForm = (Form^)pObj;
			if (m_hCLRMainWnd == (HWND)100100)
			{
				m_hCLRMainWnd = (HWND)thisForm->Handle.ToInt64();
			}
			thisForm->Show();
		}
		return (IDispatch*)Marshal::GetIUnknownForObject(pObj).ToPointer();
	}
	return nullptr;
}

Control^ CTangramCLRProxy::GetCanSelect(Control^ ctrl, bool direct)
{
	int nCount = ctrl->Controls->Count;
	Control^ pCtrl = nullptr;
	if (nCount)
	{
		for (int i = direct ? (nCount - 1):0; direct ? i >= 0 : i <= nCount - 1; direct ? i-- : i++)
		{
			pCtrl = ctrl->Controls[i];
			if (direct && pCtrl->TabStop && pCtrl->Visible && pCtrl->Enabled)
				return pCtrl;
			pCtrl = GetCanSelect(pCtrl, direct);
			if (pCtrl != nullptr)
				return pCtrl;
		}
	}
	else if ((ctrl->CanSelect||ctrl->TabStop)&&ctrl->Visible&&ctrl->Enabled)
		return ctrl;
	return nullptr;
}

HRESULT CTangramCLRProxy::ProcessCtrlMsg(HWND hCtrl,bool bShiftKey)
{
	Control^ pCtrl = (Control^)Control::FromHandle((IntPtr)hCtrl);
	if (pCtrl == nullptr)
		return S_FALSE;
	Control^ pChildCtrl = GetCanSelect(pCtrl, !bShiftKey);
	
	if (pChildCtrl)
		pChildCtrl->Select();
	return S_OK;
}

BOOL CTangramCLRProxy::ProcessUCMAMsg(IWndNode* pObj, IMessageObj* pMsgObj)
{
	WndNode^ pNode = (WndNode^)theAppProxy._createObject<IWndNode, TangramCLR::WndNode>(pObj);
	if (pNode != nullptr)
	{
		pNode->m_pCurMsgObj = pMsgObj;
		pNode->ActiveMethod(L"ProcessUCMAMsg", nullptr);
	}
	return false; 
}

BOOL CTangramCLRProxy::ProcessFormMsg(HWND hFormWnd, LPMSG lpMSG, int nMouseButton)
{
	Control^ Ctrl = Form::FromHandle((IntPtr)hFormWnd);
	//LPMSG lpMsg = (LPMSG)lpMSG;
	if (Ctrl == nullptr)
		return false;
	System::Windows::Forms::Message Msg;
	Msg.HWnd = (IntPtr)lpMSG->hwnd;
	Msg.Msg = lpMSG->message;
	Msg.WParam = (IntPtr)((LONG)lpMSG->wParam);
	Msg.LParam = (IntPtr)lpMSG->lParam;
	Form^ pForm = static_cast<Form^>(Ctrl);
	if (pForm == nullptr)
		return Ctrl->PreProcessMessage(Msg);
	if (pForm->IsMdiContainer)
	{
		Ctrl = pForm->ActiveControl;
		if (Ctrl!=nullptr)
			return Ctrl->PreProcessMessage(Msg);
	}
	return pForm->PreProcessMessage(Msg);
}

HWND CTangramCLRProxy::GetHwnd(HWND parent, int x, int y, int width, int height)
{
	System::Windows::Interop::HwndSourceParameters hwsPars;
	hwsPars.ParentWindow = System::IntPtr(parent);
	hwsPars.WindowStyle = WS_CHILD | WS_VISIBLE;
	hwsPars.PositionX = x;
	hwsPars.PositionY = y;
	hwsPars.Width = width;
	hwsPars.Height = height;
	System::Windows::Interop::HwndSource^ hws= gcnew System::Windows::Interop::HwndSource(hwsPars);
	return nullptr;
}

void CTangramCLRProxy::SelectNode(IWndNode* pNode)
{
	if (pNode == nullptr)
	{
		return;
	}
	Object^ pObj = nullptr;
	try
	{
		if(pNode)
			pObj = theAppProxy._createObject<IWndNode, TangramCLR::WndNode>(pNode);
	}
	catch (...)
	{

	}
	finally
	{
		if (pObj != nullptr)
		{
			try
			{
				m_pPropertyGrid->SelectedObject = pObj;
			}
			catch (...)
			{

			}
		}
		else
		{
			m_pPropertyGrid->SelectedObject = nullptr;
		}
	}
}

IDispatch* CTangramCLRProxy::TangramCreateObject(BSTR bstrObjID, long hParent, IWndNode* pHostNode)
{
	String^ strID = BSTR2STRING(bstrObjID);

	Control^ pObj = static_cast<Control^>(TangramCLR::Tangram::CreateObject(strID));
	WndNode^ _pNode = (WndNode^)theAppProxy._createObject<IWndNode, TangramCLR::WndNode>(pHostNode);
	if (pObj != nullptr&&pHostNode)
	{
		m_strObjTypeName = pObj->GetType()->Name;
		__int64 h = 0;
		pHostNode->get_Handle(&h);
		::SendMessage((HWND)h, WM_TANGRAMMSG, 0, 19920612);
		IWndNode* pRootNode = NULL;
		pHostNode->get_RootNode(&pRootNode);
		m_pTangramProxy->DelegateEvent(pObj, pHostNode);
		HWND hWnd = (HWND)pObj->Handle.ToInt64();
		CComBSTR bstrName(L"");
		pHostNode->get_Name(&bstrName);
		CString strName = OLE2T(bstrName);
		if (strName.CompareNoCase(_T("TangramPropertyGrid")) == 0)
		{
			m_pPropertyGrid = (PropertyGrid^)pObj;
			m_pPropertyGrid->ToolbarVisible = false;
			m_pPropertyGrid->PropertySort = PropertySort::Alphabetical;
		}
		TangramCLR::Tangram::m_pFrameworkElementDic[pObj] = _pNode;
		IDispatch* pDisp = (IDispatch*)(Marshal::GetIUnknownForObject(pObj).ToInt64());
		_pNode->m_pHostObj = pObj;
		if( m_pProxy->IsMDIFrameNode(pHostNode)==false)
			InitTangramNode(pHostNode, pObj, true);
		return pDisp;
	}
	if (pObj == nullptr)
	{
		pObj = gcnew TangramCLR::ErrorCtrl();
		CString strInfo = _T("");
		strInfo.Format(_T("Error Information: the Component \"%s\" don't exists,Please install it correctly."), OLE2T(bstrObjID));
		((TangramCLR::ErrorCtrl^)pObj)->ErrorInfoText = BSTR2STRING(strInfo);
		TangramCLR::Tangram::m_pFrameworkElementDic[pObj] = _pNode;
		return (IDispatch*)(Marshal::GetIUnknownForObject(pObj).ToInt64());
	}
	return nullptr;
}

int CTangramCLRProxy::IsWinForm(HWND hWnd)
{
	if (hWnd == 0)
		return false;
	IntPtr handle = (IntPtr)hWnd;
	Control^  pControl = Control::FromHandle(handle);
	if (pControl != nullptr)
	{
		if (pControl->GetType()->IsSubclassOf(Form::typeid))
		{
			Form^ pForm = (Form^)pControl;
			if (pForm->IsMdiContainer)
				return 2;
			else
				return 1;
		}
		else if (::GetWindowLong(hWnd, GWL_EXSTYLE)&WS_EX_APPWINDOW)
		{
			int nCount = pControl->Controls->Count;
			String^ strName = L"";
			for (int i = nCount - 1; i >= 0; i--)
			{
				Control^ pCtrl = pControl->Controls[i];
				strName = pCtrl->GetType()->Name->ToLower();
				if (strName == L"mdiclient")
				{
					return 2;
				}
			}
			return 1;
		}
	}
	return 0;
}

IDispatch* CTangramCLRProxy::GetCLRControl(IDispatch* CtrlDisp, BSTR bstrNames)
{
	CString strNames = OLE2T(bstrNames);
	if (strNames != _T(""))
	{
		int nPos = strNames.Find(_T(","));
		if (nPos == -1)
		{
			Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)CtrlDisp);
			if (pCtrl != nullptr)
			{
				Control::ControlCollection^ Ctrls = pCtrl->Controls;
				pCtrl = Ctrls[BSTR2STRING(bstrNames)];
				if (pCtrl == nullptr)
				{
					int nIndex = _wtoi(OLE2T(bstrNames));
					pCtrl = Ctrls[nIndex];
				}
				if (pCtrl != nullptr)
					return (IDispatch*)Marshal::GetIDispatchForObject(pCtrl).ToPointer();
			}
			return S_OK;
		}

		Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)CtrlDisp);
		while (nPos != -1)
		{
			CString strName = strNames.Left(nPos);
			if (strName != _T(""))
			{
				if (pCtrl != nullptr)
				{
					Control^ pCtrl2 = pCtrl->Controls[BSTR2STRING(strName)];
					if (pCtrl2 == nullptr)
					{
						if (pCtrl->Controls->Count > 0)
							pCtrl2 = pCtrl->Controls[_wtoi(strName)];
					}
					if (pCtrl2 != nullptr)
						pCtrl = pCtrl2;
					else
						return S_OK;
				}
				else
					return S_OK;
			}
			strNames = strNames.Mid(nPos + 1);
			nPos = strNames.Find(_T(","));
			if (nPos == -1)
			{
				if (pCtrl != nullptr)
				{
					Control^ pCtrl2 = pCtrl->Controls[BSTR2STRING(strNames)];
					if (pCtrl2 == nullptr)
					{
						if (pCtrl->Controls->Count > 0)
							pCtrl2 = pCtrl->Controls[_wtoi(strName)];
					}
					if (pCtrl2 == nullptr)
						return S_OK;
					if (pCtrl2 != nullptr)
						return (IDispatch*)Marshal::GetIDispatchForObject(pCtrl2).ToPointer();
				}
			}
		}
	}
	return NULL;
}

BSTR CTangramCLRProxy::GetCtrlName(IDispatch* _pCtrl)
{
	Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)_pCtrl);
	if (pCtrl != nullptr)
		return STRING2BSTR(pCtrl->Name);
	return L"";
}

void CTangramCLRProxy::ReleaseTangramObj(IDispatch* pDisp)
{
	LONGLONG nValue = (LONGLONG)pDisp;
	_removeObject(nValue);
}

HWND CTangramCLRProxy::GetMDIClientHandle(IDispatch* pMDICtrl)
{
	Form^ pCtrl = (Form^)Marshal::GetObjectForIUnknown((IntPtr)pMDICtrl);
	if (pCtrl != nullptr)
	{
		Control^ ctrl = TangramCLR::Tangram::GetMDIClient(pCtrl);
		if (ctrl != nullptr)
			return (HWND)ctrl->Handle.ToInt64();
	}
	return NULL;
}

IDispatch* CTangramCLRProxy::GetCtrlByName(IDispatch* CtrlDisp, BSTR bstrName, bool bFindInChild)
{
	try
	{
		Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)CtrlDisp);
		if (pCtrl != nullptr)
		{
			cli::array<Control^, 1>^ pArray = pCtrl->Controls->Find(BSTR2STRING(bstrName), bFindInChild);
			if (pArray != nullptr&&pArray->Length)
				return (IDispatch*)Marshal::GetIDispatchForObject(pArray[0]).ToPointer();
		}

	}
	catch (System::Exception^)
	{

	}
	return NULL;
}

int CTangramCLRProxy::IsSpecifiedType(IUnknown* pUnknown, BSTR bstrName)
{
	Object^ pObj = Marshal::GetObjectForIUnknown((IntPtr)pUnknown);
	if (pObj&&pObj->GetType()->FullName->Equals(BSTR2STRING(bstrName)))
	{
		return 1;
	}
	return 0;
}

void CTangramCLRProxy::SelectObj(IDispatch* CtrlDisp)
{
	try
	{
		Object^ pCtrl = (Object^)Marshal::GetObjectForIUnknown((IntPtr)CtrlDisp);
		if (pCtrl != nullptr)
		{
			m_pPropertyGrid->SelectedObject = pCtrl;
		}

	}
	catch (System::Exception^ e)
	{
		String^ strInfo = e->Message;
	}
}

BSTR CTangramCLRProxy::GetCtrlValueByName(IDispatch* CtrlDisp, BSTR bstrName, bool bFindInChild)
{
	try
	{
		Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)CtrlDisp);
		if (pCtrl != nullptr)
		{
			cli::array<Control^, 1>^ pArray = pCtrl->Controls->Find(BSTR2STRING(bstrName), bFindInChild);
			if (pArray != nullptr&&pArray->Length)
			{
				return STRING2BSTR(pArray[0]->Text);
			}
		}
	}
	catch (System::Exception^)
	{

	}
	return NULL;
}

void CTangramCLRProxy::SetCtrlValueByName(IDispatch* CtrlDisp, BSTR bstrName, bool bFindInChild, BSTR strVal)
{
	try
	{
		Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)CtrlDisp);
		if (pCtrl != nullptr)
		{
			cli::array<Control^, 1>^ pArray = pCtrl->Controls->Find(BSTR2STRING(bstrName), bFindInChild);
			if (pArray != nullptr&&pArray->Length)
			{
				pArray[0]->Text = BSTR2STRING(strVal);
				return;
			}
		}
	}
	catch (System::Exception^)
	{

	}
}

HWND CTangramCLRProxy::GetCtrlHandle(IDispatch* _pCtrl)
{
	Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)_pCtrl);
	if (pCtrl != nullptr)
		return (HWND)pCtrl->Handle.ToInt64();
	return 0;
}

IDispatch* CTangramCLRProxy::GetCtrlFromHandle(HWND hWnd)
{
	Control^ pCtrl = Control::FromHandle((IntPtr)hWnd);
	if (pCtrl != nullptr) {
		return (IDispatch*)Marshal::GetIUnknownForObject(pCtrl).ToPointer();
	}
	return nullptr;
}

HWND CTangramCLRProxy::IsCtrlCanNavigate(IDispatch* ctrl)
{
	Control^ pCtrl = (Control^)Marshal::GetObjectForIUnknown((IntPtr)ctrl);
	if (pCtrl != nullptr)
	{
		if (pCtrl->Dock == DockStyle::Fill)
			return (HWND)pCtrl->Handle.ToInt64();
	}
	return 0;
}

void CTangramCLRProxy::TangramAction(BSTR bstrXml, IWndNode* pNode)
{
	CString strXml = OLE2T(bstrXml);
	if (strXml != _T(""))
	{
		CTangramXmlParse m_Parse;
		if (m_Parse.LoadXml(strXml))
		{
			if (pNode == nullptr)
			{
				CString strInit = m_Parse.attr(_T("Init"), _T(""));
				if (m_strCollaborationScript != _T("")&&strInit.CompareNoCase(_T("StopUcma")) == 0)
				{
#ifdef TANGRAMCOLLABORATION
					m_strCollaborationScript = _T("");
					Collaboration::TangramUcmaApp::TangramVerb(L"<tangram  Init='StopUcma'/>", L"");
#endif					
				}
			}
			else
			{
				WndNode^ pWindowNode = (WndNode^)theAppProxy._createObject<IWndNode, WndNode>(pNode);
				if (pWindowNode)
				{
					int nType = m_Parse.attrInt(_T("Type"), 0);
					switch (nType)
					{
					case 5:
						if (pWindowNode != nullptr)
						{
						}
						break;
					default:
						{
							CString strID = m_Parse.attr(_T("ObjID"), _T(""));
							CString strMethod = m_Parse.attr(_T("Method"), _T(""));
							if (strID != _T("") && strMethod != _T(""))
							{
								cli::array<Object^, 1>^ pObjs = { BSTR2STRING(strXml), pWindowNode };
								TangramCLR::Tangram::ActiveMethod(BSTR2STRING(strID), BSTR2STRING(strMethod), pObjs);
							}
						}
						break;
					}
				}
			}
		}
	}
}

bool CTangramCLRProxy::_insertObject(Object^ key, Object^ val)
{
	Hashtable^ htObjects = (Hashtable^)m_htObjects;
	htObjects[key] = val;
	return true;
}

Object^ CTangramCLRProxy::_getObject(Object^ key)
{
	Hashtable^ htObjects = (Hashtable^)m_htObjects;
	return htObjects[key];
}

bool CTangramCLRProxy::_removeObject(Object^ key)
{
	Hashtable^ htObjects = (Hashtable^)m_htObjects;

	if (htObjects->ContainsKey(key))
	{
		htObjects->Remove(key);
		return true;
	}
	return false;
}

void CTangramNodeEvent::OnExtendComplete()
{
	if (m_pTangramNodeCLREvent)
		m_pTangramNodeCLREvent->OnExtendComplete(NULL);
}

void CTangramNodeEvent::OnTabChange(int nActivePage, int nOldPage)
{
	if (m_pWndNode != nullptr)
		m_pTangramNodeCLREvent->OnTabChange(nActivePage, nOldPage);
}

void CTangramNodeEvent::OnMessageReceived(BSTR barg1, BSTR barg2)
{
	if (m_pWndNode != nullptr)
		m_pTangramNodeCLREvent->OnMessageReceived(barg1, barg2);
}

void CTangramNodeEvent::OnDestroy()
{
	LONGLONG nValue = (LONGLONG)m_pWndNode;
	theAppProxy._removeObject(nValue);
	if (m_pTangramNodeCLREvent)
	{
		m_pTangramNodeCLREvent->OnDestroy();
		delete m_pTangramNodeCLREvent;
		m_pTangramNodeCLREvent = nullptr;
	}
	this->DispEventUnadvise(m_pWndNode);
}

void CTangramNodeEvent::OnDocumentComplete(IDispatch* pDocdisp, BSTR bstrUrl)
{
	if (m_pWndNode!=nullptr)
		m_pTangramNodeCLREvent->OnDocumentComplete(pDocdisp, bstrUrl);
}

void CTangramNodeEvent::OnNodeAddInCreated(IDispatch* pAddIndisp, BSTR bstrAddInID, BSTR bstrAddInXml)
{
	if (m_pWndNode != nullptr)
		m_pTangramNodeCLREvent->OnNodeAddInCreated(pAddIndisp, bstrAddInID, bstrAddInXml);
}

void CTangramCLRApp::OnTangramClose()
{
	if (theApp.m_pTangram)
	{
		theApp.m_pTangram->put_AppKeyValue(CComBSTR(L"CLRProxy"), CComVariant((LONGLONG)0));
	}
	AtlTrace(_T("*************Begin CTangramCLRApp::OnClose:  ****************\n"));
	TangramCLR::Tangram::GetTangram()->Fire_OnClose();
	FormCollection^ pCollection = System::Windows::Forms::Application::OpenForms;
	int nCount = pCollection->Count;
	while (pCollection->Count>0)
	{
		Form^ pForm = pCollection[0];
		pForm->Close();
		delete pForm;
	}

	Object^ pPro = (Object^)theAppProxy.m_pPropertyGrid;
	if (pPro&&theAppProxy.m_pPropertyGrid->SelectedObject)
		theAppProxy.m_pPropertyGrid->SelectedObject = nullptr;
	EnterCriticalSection(&theApp.m_csTaskRecycleCriticalSection);
	LeaveCriticalSection(&theApp.m_csTaskRecycleCriticalSection);
	AtlTrace(_T("*************End CTangramCLRApp::OnClose:  ****************\n"));
}

void CTangramCLRApp::OnExtendComplete(long hWnd, BSTR bstrUrl, IWndNode* pRootNode)
{
	TangramCLR::Tangram^ pManager = TangramCLR::Tangram::GetTangram();
	WndNode^ _pRootNode = theAppProxy._createObject<IWndNode, WndNode>(pRootNode);
	IntPtr nHandle = (IntPtr)hWnd;
	pManager->Fire_OnExtendComplete(nHandle, BSTR2STRING(bstrUrl), _pRootNode);
}

CWndPageClrEvent::CWndPageClrEvent()
{

}

CWndPageClrEvent::~CWndPageClrEvent()
{
}

void __stdcall  CWndPageClrEvent::OnDestroy()
{
	if (m_pPage)
		delete m_pPage;
}

void __stdcall  CWndPageClrEvent::OnTabChange(IWndNode* sender, int nActivePage, int nOldPage)
{
	Object^ pObj = m_pPage;
	TangramCLR::WndPage^ pPage = static_cast<TangramCLR::WndPage^>(pObj);
	WndNode^ pWindowNode = (WndNode^)theAppProxy._createObject<IWndNode, WndNode>(sender);
	pPage->Fire_OnTabChange(pWindowNode, nActivePage, nOldPage);
}

void CWndPageClrEvent::OnInitialize(IDispatch* pHtmlWnd, BSTR bstrUrl)
{
	Object^ pObj = m_pPage;
	TangramCLR::WndPage^ pPage = static_cast<TangramCLR::WndPage^>(pObj);
	pPage->Fire_OnDocumentComplete(pPage, Marshal::GetObjectForIUnknown((IntPtr)pHtmlWnd), BSTR2STRING(bstrUrl));
}

void CWndPageClrEvent::OnIPCMsg(IWndFrame* sender, BSTR bstrType, BSTR bstrContent, BSTR bstrFeature)
{
	Object^ pObj = m_pPage;
	TangramCLR::WndPage^ pPage = static_cast<TangramCLR::WndPage^>(pObj);
	WndFrame^ pWndFrame = (WndFrame^)theAppProxy._createObject<IWndFrame, WndFrame>(sender);
	pPage->Fire_OnIPCMsg(pWndFrame, BSTR2STRING(bstrType), BSTR2STRING(bstrContent), BSTR2STRING(bstrFeature));
}

void CTangramCLRProxy::OnApplicationExit(System::Object ^sender, System::EventArgs ^e)
{
	if (theAppProxy.m_strCollaborationScript != _T(""))
	{
#ifdef TANGRAMCOLLABORATION
		Collaboration::TangramUcmaApp::TangramVerb(L"<tangram  Init='StopUcma'/>", L"");
		theAppProxy.m_strCollaborationScript = _T("");
#endif
	}
	for each (KeyValuePair<String^, TangramAppProxy^>^ obj in TangramCLR::Tangram::m_pTangramAppProxyDic)
	{
		if (obj->Value != nullptr)
		{
			TangramAppProxy^ proxy = obj->Value;
			if (proxy->m_pTangramAppProxy&&::IsWindow(proxy->m_pTangramAppProxy->m_hMainWnd))
				::DestroyWindow(proxy->m_pTangramAppProxy->m_hMainWnd);
		}
	}
}


CTangramWPFObjWrapper::~CTangramWPFObjWrapper(void)
{
	WndNode^ pNode = nullptr;
	if (TangramCLR::Tangram::m_pFrameworkElementDic->TryGetValue(m_pUIElement, pNode))
	{
		TangramCLR::Tangram::m_pFrameworkElementDic->Remove(m_pUIElement);
	}
}

void CTangramWPFObjWrapper::ShowVisual(BOOL bShow) 
{
	if (bShow)
	{
		m_pUIElement->Visibility = System::Windows::Visibility::Visible;
	}
	else
	{
		m_pUIElement->Visibility = System::Windows::Visibility::Hidden;
	}
}

void CTangramWPFObjWrapper::Focusable(BOOL bFocus)
{
	m_pUIElement->Focusable = bFocus;
}

void CTangramWPFObjWrapper::InvalidateVisual() 
{
	if (m_pUIElement)
	{
		m_pUIElement->InvalidateVisual();
	}
}

BOOL CTangramWPFObjWrapper::IsVisible()
{
	return m_pUIElement->IsVisible;
}

CTangramWPFObj* CTangramWPFObjWrapper::CreateControl(Type^ type, HWND parent, DWORD style, int x, int y, int width, int height)
{
	m_pUIElement = (FrameworkElement^)Activator::CreateInstance(type);
	m_pDisp = (IDispatch*)(System::Runtime::InteropServices::Marshal::GetIUnknownForObject(m_pUIElement).ToInt64());
	if (m_pDisp)
	{
		Interop::HwndSourceParameters^ sourceParams = gcnew Interop::HwndSourceParameters("Tangram WpfControlWrapper");
		sourceParams->PositionX = x;
		sourceParams->PositionY = y;
		sourceParams->Height = height;
		sourceParams->Width = width;
		sourceParams->WindowStyle = style;
		sourceParams->ParentWindow = IntPtr(parent);

		m_pSource = gcnew Interop::HwndSource(*sourceParams);
		m_pSource->AddHook(gcnew Interop::HwndSourceHook(&ChildHwndSourceHook));
		m_hwndWPF = (HWND)m_pSource->Handle.ToPointer();
	}

	return m_hwndWPF == NULL ? nullptr : this;
}


void CTangramCLRProxy::OnSelectedObjectsChanged(System::Object ^sender, System::EventArgs ^e)
{
	if (theAppProxy.m_pPropertyGrid->SelectedObject != nullptr)
	{
		int nType = 100;
		IDispatch* pDisp = (IDispatch*)(Marshal::GetIUnknownForObject(theAppProxy.m_pPropertyGrid->SelectedObject).ToInt64());
		if (pDisp)
		{
			HWND hWnd = nullptr;
			String^ s = theAppProxy.m_pPropertyGrid->SelectedObject->GetType()->ToString();
			if (s == L"System.Windows.Forms.Form")
			{
				Form^ pForm = (Form^)theAppProxy.m_pPropertyGrid->SelectedObject;
				hWnd = (HWND)pForm->Handle.ToInt64();
				auto it = theAppProxy.m_mapDesigningForm.find(hWnd);
				if (it == theAppProxy.m_mapDesigningForm.end())
				{
					pForm->ControlAdded += gcnew System::Windows::Forms::ControlEventHandler(&OnControlAdded);
					pForm->ControlRemoved += gcnew System::Windows::Forms::ControlEventHandler(&OnControlRemoved);
					pForm->HandleDestroyed += gcnew System::EventHandler(&OnHandleDestroyed);
				}
				nType = 2;
				if (pForm->IsMdiContainer)
				{
					nType = 3;
					theAppProxy.GetMDIClientHandle(pDisp);
					Control^ ctrl = TangramCLR::Tangram::GetMDIClient(pForm);
					__int64 nHandle = ctrl->Handle.ToInt64();
					::SetWindowLongPtr(hWnd, GWLP_USERDATA, nHandle);
				}
			}
			else if(s == L"System.Windows.Forms.UserControl")
			{
				nType = 1;
				UserControl^ pCtrl = (UserControl^)theAppProxy.m_pPropertyGrid->SelectedObject;
				hWnd = (HWND)pCtrl->Handle.ToInt64();
			}
			else if (theAppProxy.m_pPropertyGrid->SelectedObject->GetType()->IsSubclassOf(Control::typeid))
			{
				nType = 4;
				Control^ ctrl = (Control^)theAppProxy.m_pPropertyGrid->SelectedObject;
				hWnd = (HWND)ctrl->Handle.ToInt64();
			}
			else
			{
				CComQIPtr<VxDTE::CodeElement> pCodeElement(pDisp);
				if (pCodeElement)
				{
					BSTR bstrName = ::SysAllocString(L"");
					pCodeElement->get_FullName(&bstrName);
					theAppProxy.m_pProxy->m_pTangramPackageProxy->OnSelectedObjectsChanged(pDisp, OLE2T(bstrName), 1, 0);
					::SysFreeString(bstrName);
					return;
				}
			}
			BSTR strType = STRING2BSTR(s);
			theAppProxy.m_pProxy->m_pTangramPackageProxy->OnSelectedObjectsChanged(pDisp, OLE2T(strType), (LPARAM)hWnd, nType);
			::SysFreeString(strType);
		}
	}
}

void CTangramCLRProxy::OnControlAdded(System::Object ^sender, System::Windows::Forms::ControlEventArgs ^e)
{
	String^ strType = e->Control->GetType()->ToString();//System.Windows.Forms.MdiClient
	if (strType == L"System.Windows.Forms.MdiClient")
	{
		__int64 nHandle = e->Control->Handle.ToInt64();
		HWND hWnd = (HWND)((Form^)sender)->Handle.ToInt64();
		::SetWindowLongPtr(hWnd, GWLP_USERDATA, nHandle);
		::SendMessage(hWnd, WM_TANGRAMMSG, nHandle, 20170907);
	}
}

void CTangramCLRProxy::OnControlRemoved(System::Object ^sender, System::Windows::Forms::ControlEventArgs ^e)
{
	String^ strType = e->Control->GetType()->ToString();
	if (strType == L"System.Windows.Forms.MdiClient")
	{
		__int64 nHandle = e->Control->Handle.ToInt64();
		::SetWindowLongPtr((HWND)((Form^)sender)->Handle.ToInt64(), GWLP_USERDATA, 0);
	}
}

void CTangramCLRProxy::OnHandleDestroyed(System::Object ^sender, System::EventArgs ^e)
{
	Form^ pForm = (Form^)sender;
	HWND hWnd = (HWND)pForm->Handle.ToInt64();
	auto it = theAppProxy.m_mapDesigningForm.find(hWnd);
	if (it != theAppProxy.m_mapDesigningForm.end())
	{
		theAppProxy.m_mapDesigningForm.erase(it);
	}
}
