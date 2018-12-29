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

// TangramCLR2COM.cpp : Implementation of CTangramCLRHost

#include "stdafx.h"
#include "dllmain.h"
#include "TangramCLRHost.h"
#include "TangramProxy.h"
#include "Tangram.c"

using namespace System::Reflection;
using namespace TangramCLR;
//
//class ATL_NO_VTABLE CTangramCLRHostHelper :
//	public CComObjectRootBase,
//	public CComCoClass<CTangramCLRHostHelper>
//{
//public:
//	static HRESULT WINAPI UpdateRegistry(BOOL bRegister)
//	{
//		return theApp.UpdateRegistryFromResource(IDR_TANGRAMCLR2COM, bRegister);
//	}
//
//	static HRESULT WINAPI CreateInstance(void* pv, REFIID riid, LPVOID* ppv)
//	{
//		if (theApp.m_pTangramCLRHost == nullptr)
//			theApp.m_pTangramCLRHost = new CComObject<CTangramCLRHost>;
//		return theApp.m_pTangramCLRHost->QueryInterface(riid, ppv);
//	}
//};
//
//TANGRAM_OBJECT_ENTRY_AUTO(__uuidof(TangramCLRHost), CTangramCLRHostHelper)
//
//// CTangramCLRHost
//CTangramCLRHost::CTangramCLRHost()
//{
//}
//
//ULONG CTangramCLRHost::InternalRelease() 
//{
//	return 1; 
//}
