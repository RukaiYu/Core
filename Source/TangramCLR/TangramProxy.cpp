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

#include "StdAfx.h"
#include "dllmain.h"
#include "TangramObj.h"
#include "TangramProxy.h"

namespace TangramCLR
{
	TangramProxy::TangramProxy(void)
	{
		m_pCurrentNode = nullptr;
		m_CtrlEventDic = nullptr;
		m_pHandleCreated = nullptr;
		m_pHandleDestroyed = nullptr;
		m_pOnClick = nullptr;
		m_pDoubleClick = nullptr;
		m_pEnter = nullptr;
		m_pOnLeave =nullptr;
		m_pOnEnabledChanged = nullptr;
		m_pOnLostFocus = nullptr;
		m_pOnGotFocus = nullptr;
		m_pOnKeyUp = nullptr;
		m_pOnKeyDown = nullptr;
		m_pOnKeyPress = nullptr;
		m_pOnMouseClick = nullptr;
		m_pOnMouseDoubleClick = nullptr;
		m_pOnMouseDown = nullptr;
		m_pOnMouseEnter = nullptr;
		m_pOnMouseHover = nullptr;
		m_pOnMouseLeave = nullptr;
		m_pOnMouseMove = nullptr;
		m_pOnMouseUp = nullptr;
		m_pOnMouseWheel = nullptr;
		m_pOnTextChanged = nullptr;
		m_pOnVisibleChanged = nullptr;
		m_pOnSizeChanged = nullptr;
		m_pOnClientSizeChanged = nullptr;
		m_pOnParentChanged = nullptr;
		m_pOnResize = nullptr;
		theAppProxy.m_pTangramProxy = this;
	}

	int TangramProxy::TangramInit(String^ strInit)
	{
		CString strInfo = strInit;
		if (strInfo != _T(""))
		{
			__int64 nPointer = _wtoi64(strInfo);
			if (nPointer)
			{
				theAppProxy.m_pProxy = (CTangramProxyBase*)nPointer;
				theAppProxy.m_pProxy->m_pCLRProxy = &theAppProxy;
				theAppProxy.m_pProxy->m_pTangramCLRAppProxy = static_cast<CTangramAppProxy*>(&theApp);
				//20180817
				theApp.m_pTangramProxyBase = theAppProxy.m_pProxy;
				if (theApp.m_pTangram == nullptr)
				{
					CComPtr<ITangram> pTangram;
					pTangram.CoCreateInstance(CComBSTR("Tangram.Tangram.1"));
					theApp.m_pTangram = pTangram.Detach();
					if (theApp.m_pTangram)
					{
	#ifdef TANGRAMCOLLABORATION
						CComVariant m_v;
						if (theApp.m_pTangram->get_AppKeyValue(CComBSTR(L"collaborationscript"), &m_v) == S_OK)
						{
							theAppProxy.m_strCollaborationScript = OLE2T(m_v.bstrVal);
							if (theAppProxy.m_strCollaborationScript != _T(""))
								Collaboration::TangramUcmaApp::TangramVerb(BSTR2STRING(theAppProxy.m_strCollaborationScript), L"");
						}
	#endif
					}
				}
			}
		}
		return 0;
	};

	void TangramProxy::AttachHandleDestroyedEvent(Control^ pCtrl)
	{
		if(m_pHandleDestroyed==nullptr)
			m_pHandleDestroyed = gcnew EventHandler(theAppProxy.m_pTangramProxy, &TangramProxy::OnHandleDestroyed);
		pCtrl->HandleDestroyed += m_pHandleDestroyed;
	}

	void TangramProxy::AttachEvent(Control^ pCtrl, WindowEventType EventType)
	{
		if (m_CtrlEventDic == nullptr)
			m_CtrlEventDic = gcnew Dictionary < Control^, String^ > ;
		String^ strVal = L"";
		int nPos = -1;
		CString str = _T("");
		str.Format(_T(",%d,"),EventType);
		String^ _strKey = BSTR2STRING(str);
		if (m_CtrlEventDic->TryGetValue(pCtrl, strVal))
		{
			nPos = strVal->IndexOf(_strKey);
			if (nPos != -1)
				return;
			else
			{
				strVal += _strKey;
				m_CtrlEventDic[pCtrl] = strVal;
			}
		}
		else
		{
			m_CtrlEventDic->Add(pCtrl, _strKey);
		}
		switch (EventType)
		{
		case TangramClick:
			{
				if(m_pOnClick==nullptr)
					m_pOnClick = gcnew EventHandler(this, &TangramProxy::OnClick);
				pCtrl->Click += m_pOnClick;
			}
			break;
		case TangramDoubleClick:
			{
				if(m_pDoubleClick==nullptr)
					m_pDoubleClick = gcnew EventHandler(this, &TangramProxy::OnDoubleClick);
				pCtrl->DoubleClick += m_pDoubleClick;
			}
			break;
		case TangramEnter:
			{
				if(m_pEnter==nullptr)
					m_pEnter = gcnew EventHandler(this, &TangramProxy::OnEnter);
				pCtrl->Enter += m_pEnter;
			}
			break;
		case TangramLeave:
			{
				if(m_pOnLeave ==nullptr)
					m_pOnLeave = gcnew EventHandler(this, &TangramProxy::OnLeave);

				pCtrl->Leave += m_pOnLeave;
			}
			break;
		case TangramEnabledChanged:
			{
				if(m_pOnEnabledChanged ==nullptr)
					m_pOnEnabledChanged = gcnew EventHandler(this, &TangramProxy::OnEnabledChanged);
				pCtrl->EnabledChanged += m_pOnEnabledChanged;
			}
			break;
		case TangramLostFocus:
			{
				if(m_pOnLostFocus ==nullptr)
					m_pOnLostFocus = gcnew EventHandler(this, &TangramProxy::OnLostFocus);
				pCtrl->LostFocus += m_pOnLostFocus;
			}
			break;
		case TangramGotFocus:
			{
				if(m_pOnGotFocus ==nullptr)
					m_pOnGotFocus = gcnew EventHandler(this, &TangramProxy::OnGotFocus);
				pCtrl->GotFocus += m_pOnGotFocus;
			}
			break;
		case TangramKeyUp:
			{
				if(m_pOnKeyUp ==nullptr)
					m_pOnKeyUp = gcnew KeyEventHandler(this, &TangramProxy::OnKeyUp);
				pCtrl->KeyUp += m_pOnKeyUp;
			}
			break;
		case TangramKeyDown:
			{
				if(m_pOnKeyDown ==nullptr)
					m_pOnKeyDown = gcnew KeyEventHandler(this, &TangramProxy::OnKeyDown);
				pCtrl->KeyDown += m_pOnKeyDown;
			}
			break;
		case TangramKeyPress:
			{
				if (m_pOnKeyPress == nullptr)
					m_pOnKeyPress = gcnew KeyPressEventHandler(this, &TangramProxy::OnKeyPress);
				pCtrl->KeyPress += m_pOnKeyPress;
			}
			break;
		case TangramMouseClick:
			{
				if (m_pOnMouseClick == nullptr)
					m_pOnMouseClick = gcnew MouseEventHandler(this, &TangramProxy::OnMouseClick);
				pCtrl->MouseClick += m_pOnMouseClick;
			}
			break;// = 0x0000000a,
		case TangramMouseDoubleClick:
			{
				if (m_pOnMouseDoubleClick == nullptr)
					m_pOnMouseDoubleClick = gcnew MouseEventHandler(this, &TangramProxy::OnMouseDoubleClick);
				pCtrl->MouseDoubleClick += m_pOnMouseDoubleClick;
			}
			break;// = 0x0000000b,
		case TangramMouseDown:
			{
				if (m_pOnMouseDown == nullptr)
					m_pOnMouseDown = gcnew MouseEventHandler(this, &TangramProxy::OnMouseDown);
				pCtrl->MouseDown += m_pOnMouseDown;
			}
			break;// = 0x0000000c,
		case TangramMouseEnter:
			{
				if (m_pOnMouseEnter == nullptr)
					m_pOnMouseEnter = gcnew EventHandler(this, &TangramProxy::OnMouseEnter);
				pCtrl->MouseEnter += m_pOnMouseEnter;
			}
			break;// = 0x0000000d,
		case TangramMouseHover:
			{
				if (m_pOnMouseHover == nullptr)
					m_pOnMouseHover = gcnew EventHandler(this, &TangramProxy::OnMouseHover);
				pCtrl->MouseHover += m_pOnMouseHover;
			}
			break;// = 0x0000000e,
		case TangramMouseLeave:
			{
				if (m_pOnMouseLeave == nullptr)
					m_pOnMouseLeave = gcnew EventHandler(this, &TangramProxy::OnMouseLeave);
				pCtrl->MouseLeave += m_pOnMouseLeave;
			}
			break;// = 0x0000000f,
		case TangramMouseMove:
			{
				if (m_pOnMouseMove == nullptr)
					m_pOnMouseMove = gcnew MouseEventHandler(this, &TangramProxy::OnMouseMove);
				pCtrl->MouseMove += m_pOnMouseMove;
			}
			break;// = 0x00000010,
		case TangramMouseUp:
			{
				if (m_pOnMouseUp == nullptr)
					m_pOnMouseUp = gcnew MouseEventHandler(this, &TangramProxy::OnMouseUp);
				pCtrl->MouseUp += m_pOnMouseUp;
			}
			break;// = 0x00000011,
		case TangramMouseWheel:
			{
				if (m_pOnMouseWheel == nullptr)
					m_pOnMouseWheel = gcnew MouseEventHandler(this, &TangramProxy::OnMouseWheel);
				pCtrl->MouseWheel += m_pOnMouseWheel;
			}
			break;// = 0x00000012,
		case TangramTextChanged:
			{
				if(m_pOnTextChanged==nullptr)
					m_pOnTextChanged = gcnew EventHandler(this, &TangramProxy::OnTextChanged);
				pCtrl->TextChanged += m_pOnTextChanged;
			}
			break;// = 0x00000013,
		case TangramVisibleChanged:
			{
				if(m_pOnVisibleChanged ==nullptr)
					m_pOnVisibleChanged = gcnew EventHandler(this, &TangramProxy::OnVisibleChanged);
				pCtrl->VisibleChanged += m_pOnVisibleChanged;
			}
			break;// = 0x00000014,
		case TangramClientSizeChanged:
			{
				if(m_pOnClientSizeChanged ==nullptr)
					m_pOnClientSizeChanged = gcnew EventHandler(this, &TangramProxy::OnClientSizeChanged);
				pCtrl->ClientSizeChanged += m_pOnClientSizeChanged;
			}
			break;// = 0x00000015,
		case TangramSizeChanged:
			{
				if(m_pOnSizeChanged ==nullptr)
					m_pOnSizeChanged = gcnew EventHandler(this, &TangramProxy::OnSizeChanged);
				pCtrl->SizeChanged += m_pOnSizeChanged;
			}
			break;// = 0x00000016,
		case TangramParentChanged:
			{
				if(m_pOnParentChanged ==nullptr)
					m_pOnParentChanged = gcnew EventHandler(this, &TangramProxy::OnParentChanged);
				pCtrl->ParentChanged += m_pOnParentChanged;
			}
			break;// = 0x00000017,
		case TangramResize:
			{
				if(m_pOnResize ==nullptr)
					m_pOnResize = gcnew EventHandler(this, &TangramProxy::OnResize);
				pCtrl->Resize += m_pOnResize;
			}
			break;// = 0x00000018
		}
	}

	void TangramProxy::DelegateEvent(Control^ pObj, IWndNode* pNode)
	{
		m_pCurrentNode = pNode;
		if (pNode)
		{
			IWndPage* pPage = nullptr;
			pNode->get_WndPage(&pPage);
			if (pPage)
			{
				VARIANT_BOOL b = false;
				pPage->get_EnableSinkCLRCtrlCreatedEvent(&b);
				if (b)
				{
					if(m_pHandleCreated==nullptr)
						m_pHandleCreated = gcnew EventHandler(this,&TangramProxy::Obj_HandleCreated);
					pObj->HandleCreated += m_pHandleCreated;
				}
			}
		}
	}

	void TangramProxy::Obj_HandleCreated(Object^ sender, EventArgs^ e)
	{
		Control^ pCtrl = (Control^)sender;
		HWND hWnd = (HWND)pCtrl->Handle.ToInt64();
		if (m_pCurrentNode)
		{
			LONGLONG h = 0;
			m_pCurrentNode->get_Handle(&h);
			IWndFrame* pFrame = nullptr;
			m_pCurrentNode->get_Frame(&pFrame);
			pFrame->get_HWND(&h);
			if (::IsWindow((HWND)h))
			{
				CtrlInfo m_CtrlInfo;
				m_CtrlInfo.m_pCtrlDisp = nullptr;
				//IntPtr nDisp;
				try
				{
					IUnknown* pUnk = (IUnknown*)Marshal::GetIDispatchForObject(pCtrl).ToPointer();
					CComQIPtr<IDispatch> pDisp(pUnk);
					if (pDisp)
						m_CtrlInfo.m_pCtrlDisp = pDisp.p;
				}
				catch (InvalidCastException^ e)
				{
					String^ s = e->Message;
				}
				finally
				{
					if (m_CtrlInfo.m_pCtrlDisp)
					{
						m_CtrlInfo.m_hWnd = (HWND)pCtrl->Handle.ToInt64();
						m_CtrlInfo.m_strName = STRING2BSTR(pCtrl->Name);
						IWndPage* pPage = nullptr;
						m_pCurrentNode->get_WndPage(&pPage);
						m_CtrlInfo.m_pPage = pPage;
						m_CtrlInfo.m_pNode = m_pCurrentNode;
						::SendMessage((HWND)h, WM_TANGRAMMSG, (WPARAM)&m_CtrlInfo, 2048);
					}
				}
			}
		}

		Control^ pParentCtrl = pCtrl->Parent;
		if (pParentCtrl != nullptr)
		{
			if (pCtrl->Dock == DockStyle::Fill)
			{
				m_pCurrentNode->put_DockObj(STRING2BSTR(pCtrl->Name), (LONGLONG)pCtrl->Handle.ToInt64());
			}
		}

		Control::ControlCollection^ pControls = pCtrl->Controls;
		for each (Control^ _pCtrl in pControls)
		{
			Obj_HandleCreated(_pCtrl, e);
		}
	}

	void TangramProxy::OnEvent(Object ^sender, EventArgs ^e, WindowEventType nType)
	{
		if (theAppProxy.m_pProxy)
		{
			if (e == nullptr)
				return;
			Control^ pCtrl = static_cast<Control^>(sender);
			HWND hWnd = (HWND)pCtrl->Handle.ToInt64();
			ObjectEventInfo* pObjectEventInfo = (ObjectEventInfo*)::GetWindowLongPtr(hWnd, GWLP_USERDATA);
			if (pObjectEventInfo)
			{
				try
				{
					IDispatch* pdisp = (IDispatch*)Marshal::GetIDispatchForObject(e).ToPointer();
					IDispatch* pDisp = (IDispatch*)Marshal::GetIDispatchForObject(pCtrl).ToPointer();
					for (auto it2 : pObjectEventInfo->m_mapEventObj2)
					{
						auto it = it2.second->m_mapEventObj.find(nType);
						if (it != it2.second->m_mapEventObj.end())
							theAppProxy.m_pProxy->OnEvent(it->second, pDisp, pdisp);
					}
				}
				catch (InvalidCastException^ e)
				{
					String^ strMsg = e->Message;
				}
				finally
				{

				}
			}
		}
	}

	void TangramProxy::OnClick(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramClick);
	}

	void TangramProxy::OnDoubleClick(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramDoubleClick);
	}

	void TangramProxy::OnEnter(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramEnter);
	}

	void TangramProxy::OnLeave(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramLeave);
	}

	void TangramProxy::OnEnabledChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramEnabledChanged);
	}

	void TangramProxy::OnLostFocus(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramLostFocus);
	}

	void TangramProxy::OnGotFocus(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramGotFocus);
	}

	void TangramProxy::OnKeyUp(Object ^sender, KeyEventArgs ^e)
	{
		OnEvent(sender, e, TangramKeyUp);
	}

	void TangramProxy::OnKeyDown(Object ^sender, KeyEventArgs ^e)
	{
		OnEvent(sender, e, TangramKeyDown);
	}

	void TangramProxy::OnKeyPress(Object ^sender, KeyPressEventArgs ^e)
	{
		OnEvent(sender, e, TangramKeyPress);
	}

	void TangramProxy::OnMouseClick(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseClick);
	}

	void TangramProxy::OnMouseDoubleClick(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseDoubleClick);
	}

	void TangramProxy::OnMouseDown(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseDown);
	}

	void TangramProxy::OnMouseEnter(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseEnter);
	}

	void TangramProxy::OnMouseHover(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseHover);
	}


	void TangramProxy::OnMouseMove(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseMove);
	}

	void TangramProxy::OnMouseLeave(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseLeave);
	}

	void TangramProxy::OnMouseUp(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseUp);
	}


	void TangramProxy::OnMouseWheel(Object ^sender, MouseEventArgs ^e)
	{
		OnEvent(sender, e, TangramMouseWheel);
	}

	void TangramProxy::OnTextChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramTextChanged);
	}

	void TangramProxy::OnVisibleChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramVisibleChanged);
	}

	void TangramProxy::OnSizeChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramSizeChanged);
	}

	void TangramProxy::OnClientSizeChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramClientSizeChanged);
	}

	void TangramProxy::OnParentChanged(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramParentChanged);
	}

	void TangramProxy::OnResize(Object ^sender, EventArgs ^e)
	{
		OnEvent(sender, e, TangramResize);
	}

	void TangramProxy::OnHandleDestroyed(Object ^sender, EventArgs ^e)
	{
		Control^ pCtrl = ((Control^)sender);
		HWND hWnd = (HWND)pCtrl->Handle.ToInt64();
		ObjectEventInfo* pInfo = (ObjectEventInfo*)::GetWindowLongPtr(hWnd, GWLP_USERDATA);
		if (pInfo)
		{
			for (auto it : pInfo->m_mapEventObj2)
			{
				it.second->m_mapEventObj.clear();
				delete it.second;
			}
			pInfo->m_mapEventObj2.clear();
			delete pInfo;
		}
	}
}

