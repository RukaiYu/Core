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

using namespace System;
using namespace System::ComponentModel;
using namespace System::Collections;
using namespace System::Collections::Generic;
using namespace System::Diagnostics;
using namespace System::Windows::Forms;

using namespace System::IO;

namespace TangramCLR {
	/// <summary>
	/// Summary for TangramProxy
	/// </summary>
	private ref class TangramProxy
	{
	public:
		TangramProxy(void);

		IWndNode*					m_pCurrentNode;
		void DelegateEvent(Control^ pObj,IWndNode* pNode);
		void AttachEvent(Control^ pCtrl, WindowEventType EventType);
		void AttachHandleDestroyedEvent(Control^ pCtrl);
		static int TangramInit(String^ strInit);

	private:
		Dictionary<Control^, String^>^	m_CtrlEventDic;
		EventHandler^			m_pOnClick;
		EventHandler^			m_pDoubleClick;
		EventHandler^			m_pEnter;
		EventHandler^			m_pOnLeave;
		EventHandler^			m_pOnEnabledChanged;
		EventHandler^			m_pOnLostFocus;
		EventHandler^			m_pOnGotFocus;
		EventHandler^			m_pOnMouseEnter ;
		EventHandler^			m_pOnMouseHover ;
		EventHandler^			m_pOnMouseLeave ;
		EventHandler^			m_pOnTextChanged;
		EventHandler^			m_pOnVisibleChanged;
		EventHandler^			m_pOnSizeChanged;
		EventHandler^			m_pOnClientSizeChanged;
		EventHandler^			m_pOnParentChanged ;
		EventHandler^			m_pOnResize;
		EventHandler^			m_pHandleCreated;
		EventHandler^			m_pHandleDestroyed;
		KeyEventHandler^		m_pOnKeyUp;
		KeyEventHandler^		m_pOnKeyDown;
		KeyPressEventHandler^	m_pOnKeyPress;
		MouseEventHandler^		m_pOnMouseClick ;
		MouseEventHandler^		m_pOnMouseDoubleClick;
		MouseEventHandler^		m_pOnMouseDown ;
		MouseEventHandler^		m_pOnMouseMove;
		MouseEventHandler^		m_pOnMouseUp;
		MouseEventHandler^		m_pOnMouseWheel;

		void OnEvent(Object ^sender, EventArgs ^e, WindowEventType nType);
		void Obj_HandleCreated(Object^ sender, EventArgs^ e);
		void OnKeyPress(Object ^sender, KeyPressEventArgs ^e);
		void OnClick(Object ^sender, EventArgs ^e);
		void OnDoubleClick(Object ^sender, EventArgs ^e);
		void OnEnter(Object ^sender, EventArgs ^e);
		void OnLeave(Object ^sender, EventArgs ^e);
		void OnEnabledChanged(Object ^sender, EventArgs ^e);
		void OnLostFocus(Object ^sender, EventArgs ^e);
		void OnGotFocus(Object ^sender, EventArgs ^e);
		void OnKeyUp(Object ^sender, KeyEventArgs ^e);
		void OnKeyDown(Object ^sender, KeyEventArgs ^e);
		void OnMouseClick(Object ^sender, MouseEventArgs ^e);
		void OnMouseDoubleClick(Object ^sender, MouseEventArgs ^e);
		void OnMouseDown(System::Object ^sender, MouseEventArgs ^e);
		void OnMouseEnter(Object ^sender, EventArgs ^e);
		void OnMouseHover(Object ^sender, EventArgs ^e);
		void OnMouseMove(Object ^sender, MouseEventArgs ^e);
		void OnMouseLeave(Object ^sender, EventArgs ^e);
		void OnMouseUp(Object ^sender, MouseEventArgs ^e);
		void OnMouseWheel(Object ^sender, MouseEventArgs ^e);
		void OnTextChanged(Object ^sender, EventArgs ^e);
		void OnVisibleChanged(Object ^sender, EventArgs ^e);
		void OnSizeChanged(Object ^sender, EventArgs ^e);
		void OnClientSizeChanged(Object ^sender, EventArgs ^e);
		void OnParentChanged(Object ^sender, EventArgs ^e);
		void OnResize(System::Object ^sender, EventArgs ^e);
		void OnHandleDestroyed(Object ^sender, EventArgs ^e);
	};
}
