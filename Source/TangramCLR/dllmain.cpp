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

// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "dllmain.h"
#include "TangramCoreEvents.cpp"
#include "TangramCLRHost.h"
#include <shellapi.h>
#include <shlobj.h>

CTangramCLRApp theApp;

#ifdef _WIN64
vector<CString> adsPropertys = {
	CString("DN"),
	CString("objectClass"),
	CString("distinguishedName"),
	CString("instanceType"),
	CString("whenCreated"),
	CString("whenChanged"),
	CString("subRefs"),
	CString("uSNCreated"),
	CString("uSNChanged"),
	CString("name"),
	CString("objectGUID"),
	CString("creationTime"),
	CString("forceLogoff"),
	CString("lockoutDuration"),
	CString("lockOutObservationWindow"),
	CString("lockoutThreshold"),
	CString("maxPwdAge"),
	CString("minPwdAge"),
	CString("minPwdLength"),
	CString("modifiedCountAtLastProm"),
	CString("nextRid"),
	CString("pwdProperties"),
	CString("pwdHistoryLength"),
	CString("objectSid"),
	CString("serverState"),
	CString("uASCompat"),
	CString("modifiedCount"),
	CString("auditingPolicy"),
	CString("nTMixedDomain"),
	CString("rIDManagerReference"),
	CString("fSMORoleOwner"),
	CString("systemFlags"),
	CString("wellKnownObjects"),
	CString("objectCategory"),
	CString("isCriticalSystemObject"),
	CString("gPLink"),
	CString("dSCorePropagationData"),
	CString("otherWellKnownObjects"),
	CString("masteredBy"),
	CString("ms-DS-MachineAccountQuota"),
	CString("msDS-Behavior-Version"),
	CString("msDS-PerUserTrustQuota"),
	CString("msDS-AllUsersTrustQuota"),
	CString("msDS-PerUserTrustTombstonesQuota"),
	CString("msDs-masteredBy"),
	CString("msDS-IsDomainFor"),
	CString("msDS-NcType"),
	CString("dc"),
	CString("cn"),
	CString("description"),
	CString("showInAdvancedViewOnly"),
	CString("ou"),
	CString("msDS-TombstoneQuotaFactor"),
	CString("displayName"),
	CString("flags"),
	CString("versionNumber"),
	CString("gPCFunctionalityVersion"),
	CString("gPCFileSysPath"),
	CString("gPCMachineExtensionNames"),
	CString("ipsecName"),
	CString("ipsecID"),
	CString("ipsecDataType"),
	CString("ipsecData"),
	CString("ipsecISAKMPReference"),
	CString("ipsecNFAReference"),
	CString("ipsecOwnersReference"),
	CString("ipsecNegotiationPolicyReference"),
	CString("ipsecFilterReference"),
	CString("iPSECNegotiationPolicyType"),
	CString("iPSECNegotiationPolicyAction"),
	CString("revision"),
	CString("memberOf"),
	CString("userAccountControl"),
	CString("badPwdCount"),
	CString("codePage"),
	CString("countryCode"),
	CString("badPasswordTime"),
	CString("lastLogoff"),
	CString("lastLogon"),
	CString("logonHours"),
	CString("pwdLastSet"),
	CString("primaryGroupID"),
	CString("adminCount"),
	CString("accountExpires"),
	CString("logonCount"),
	CString("sAMAccountName"),
	CString("sAMAccountType"),
	CString("lastLogonTimestamp"),
	CString("groupType"),
	CString("member"),
	CString("samDomainUpdates"),
	CString("localPolicyFlags"),
	CString("operatingSystem"),
	CString("operatingSystemVersion"),
	CString("serverReferenceBL"),
	CString("dNSHostName"),
	CString("rIDSetReferences"),
	CString("servicePrincipalName"),
	CString("msDS-SupportedEncryptionTypes"),
	CString("msDFSR-ComputerReferenceBL"),
	CString("rIDAvailablePool"),
	CString("rIDAllocationPool"),
	CString("rIDPreviousAllocationPool"),
	CString("rIDUsedPool"),
	CString("rIDNextRID"),
	CString("dnsRecord"),
	CString("msDFSR-Flags"),
	CString("msDFSR-ReplicationGroupType"),
	CString("msDFSR-FileFilter"),
	CString("msDFSR-DirectoryFilter"),
	CString("serverReference"),
	CString("msDFSR-ComputerReference"),
	CString("msDFSR-MemberReferenceBL"),
	CString("msDFSR-Version"),
	CString("msDFSR-ReplicationGroupGuid"),
	CString("msDFSR-MemberReference"),
	CString("msDFSR-RootPath"),
	CString("msDFSR-StagingPath"),
	CString("msDFSR-Enabled"),
	CString("msDFSR-Options"),
	CString("msDFSR-ContentSetGuid"),
	CString("msDFSR-ReadOnly"),
	CString("lastSetTime"),
	CString("priorSetTime"),
	CString("sn"),
	CString("givenName"),
	CString("userPrincipalName"),
};
#endif

CTangramCLRApp::CTangramCLRApp()
{
	ATLTRACE(_T("Loading CTangramCLRApp :%p\n"), this);
	m_nAppEndPointCount = 0;
	m_pTangram = nullptr;
	m_pVSExtender = nullptr;
	m_strComponents = _T("");
	m_strAppEndpointsScript = _T("");
	InitializeCriticalSectionAndSpinCount(&m_csTaskRecycleCriticalSection, 0x00000400);
	InitializeCriticalSectionAndSpinCount(&m_csTaskListCriticalSection, 0x00000400);
	m_dwThreadID = ::GetCurrentThreadId();
	TCHAR file[MAX_PATH];
	GetModuleFileName(::GetModuleHandle(nullptr), file, MAX_PATH * sizeof(TCHAR));
	m_strAppPath = CString(file);
	int nPos = m_strAppPath.ReverseFind('\\');
	m_strAppPath = m_strAppPath.Left(nPos + 1);
}

CTangramCLRApp::~CTangramCLRApp()
{
	DeleteCriticalSection(&m_csTaskRecycleCriticalSection);
	DeleteCriticalSection(&m_csTaskListCriticalSection);
	ATLTRACE(_T("Release CTangramCLRApp :%p\n"), this);
}

CString CTangramCLRApp::GetLibPathFromAssemblyQualifiedName(CString strAssemblyQualifiedName)
{
	BOOL bLocalAssembly = false;
	strAssemblyQualifiedName.MakeLower();
	CString strPath = _T("");
	CString strLib = _T("");
	CString strObjName = _T("");
	CString strVersion = _T("");
	CString strPublickeytoken = _T("");
	int nPos = strAssemblyQualifiedName.Find(_T("publickeytoken"));
	if (nPos == -1)
	{
		bLocalAssembly = true;
		nPos = strAssemblyQualifiedName.Find(_T(","));
		if (nPos != -1)
		{
			strObjName = strAssemblyQualifiedName.Left(nPos);
			strLib = strAssemblyQualifiedName.Mid(nPos + 1);
			strLib.Trim();
			strObjName.Trim();
			if (strLib == _T("tangramclr")|| strLib == _T("tangram"))
			{
				return strObjName + _T("|") + strLib + _T("|");
			}
		}
	}
	else
	{
		strPublickeytoken = strAssemblyQualifiedName.Mid(nPos + 15);
		if (strPublickeytoken == _T("null"))
		{
			bLocalAssembly = true;
			nPos = strAssemblyQualifiedName.Find(_T("version"));
			if (nPos != -1)
			{
				strLib = strAssemblyQualifiedName.Left(nPos);
				nPos = strLib.ReverseFind(',');
				strLib = strLib.Left(nPos);
				nPos = strLib.Find(',');
				strObjName = strLib.Left(nPos);
				strLib = strLib.Mid(nPos + 1);
				strLib.Trim();
			}
		}
		else
		{
			nPos = strAssemblyQualifiedName.Find(_T("version"));
			if (nPos != -1)
			{
				strVersion = strAssemblyQualifiedName.Mid(nPos + 8);
				strLib = strAssemblyQualifiedName.Left(nPos);
				nPos = strLib.ReverseFind(',');
				strLib = strLib.Left(nPos);
				nPos = strVersion.Find(',');
				strVersion = strVersion.Left(nPos);
				nPos = strLib.Find(',');
				strObjName = strLib.Left(nPos);
				strLib = strLib.Mid(nPos + 1);
				strLib.Trim();
				TCHAR m_szBuffer[MAX_PATH];
				HRESULT hr = SHGetFolderPath(NULL, CSIDL_WINDOWS, NULL, 0, m_szBuffer);
				strPath.Format(_T("%s\\Microsoft.NET\\assembly\\GAC_MSIL\\%s\\v4.0_%s__%s\\%s.dll"), m_szBuffer, strLib, strVersion, strPublickeytoken,strLib);
				if (::PathFileExists(strPath))
					return strObjName + _T("|") + strLib + _T("|") + strPath;
				else
				{
#ifdef _WIN64
					strPath.Format(_T("%s\\Microsoft.NET\\assembly\\GAC_%d\\%s\\v4.0_%s__%s\\%s.dll"), m_szBuffer, 64, strLib, strVersion, strPublickeytoken, strLib);
#else
					strPath.Format(_T("%s\\Microsoft.NET\\assembly\\GAC_%d\\%s\\v4.0_%s__%s\\%s.dll"), m_szBuffer, 32, strLib, strVersion, strPublickeytoken, strLib);
#endif
					if (::PathFileExists(strPath))
						return strObjName + _T("|") + strLib + _T("|") + strPath;
				}
			}
		}
	}
	if (strLib != _T(""))
	{
		strPath = m_strAppPath + strLib + _T(".dll");
		if (::PathFileExists(strPath))
			return strObjName + _T("|") + strLib + _T("|") + strPath;
		else
		{
			HANDLE hFind; // file handle
			WIN32_FIND_DATA FindFileData;

			hFind = FindFirstFile(m_strAppPath + _T("*.*"), &FindFileData); // find the first file
			if (hFind == INVALID_HANDLE_VALUE)
			{
				return false;
			}

			bool bSearch = true;
			while (bSearch) // until we finds an entry
			{
				if (FindNextFile(hFind, &FindFileData))
				{
					// Don't care about . and ..
					//if(IsDots(FindFileData.cFileName))
					if ((_tcscmp(FindFileData.cFileName, _T(".")) == 0) ||
						(_tcscmp(FindFileData.cFileName, _T("..")) == 0))
						continue;

					// We have found a directory
					if ((FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
					{
						CString strPath = m_strAppPath + FindFileData.cFileName + _T("\\");
						CString strPath2 = strPath + strLib + _T(".dll");
						if (::PathFileExists(strPath2))
							return strObjName + _T("|") + strLib + _T("|") + strPath2;
						CString strRet = _GetLibPathFromAssemblyQualifiedName(strPath, strLib + _T(".dll"));
						if (strRet != _T(""))
							return strObjName + _T("|") + strLib + _T("|") + strRet;
					}

				}//FindNextFile
				else
				{
					if (GetLastError() == ERROR_NO_MORE_FILES) // no more files there
						bSearch = false;
					else {
						// some error occured, close the handle and return false
						FindClose(hFind);
						return _T("");
					}
				}
			}//while

			FindClose(hFind); // closing file handle
		}
	}

	return _T("");
}

CString CTangramCLRApp::_GetLibPathFromAssemblyQualifiedName(CString strDir, CString strLibName)
{
	CString strPath = strDir + strLibName;
	if (::PathFileExists(strPath))
		return strPath;
	HANDLE hFind; // file handle
	WIN32_FIND_DATA FindFileData;

	hFind = FindFirstFile(strDir + _T("*.*"), &FindFileData); // find the first file
	if (hFind == INVALID_HANDLE_VALUE)
	{
		return false;
	}

	bool bSearch = true;
	while (bSearch) // until we finds an entry
	{
		if (FindNextFile(hFind, &FindFileData))
		{
			// Don't care about . and ..
			//if(IsDots(FindFileData.cFileName))
			if ((_tcscmp(FindFileData.cFileName, _T(".")) == 0) ||
				(_tcscmp(FindFileData.cFileName, _T("..")) == 0))
				continue;

			// We have found a directory
			if ((FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			{
				CString strPath = strDir + FindFileData.cFileName + _T("\\");
				CString strRet = _GetLibPathFromAssemblyQualifiedName(strPath, strLibName);
				if (strRet != _T(""))
					return strRet;
			}

		}//FindNextFile
		else
		{
			if (GetLastError() == ERROR_NO_MORE_FILES) // no more files there
				bSearch = false;
			else {
				// some error occured, close the handle and return false
				FindClose(hFind);
				return false;
			}
		}
	}//while

	FindClose(hFind); // closing file handle
	return _T("");
}

#ifdef TANGRAMCOLLABORATION
#ifdef _WIN64
CString CTangramCLRApp::ExportAllObjects(IADsContainer* pContainer, VARIANT_BOOL bRecursive)
{
	HRESULT hr;
	if (pContainer == NULL)
		return _T("");

	IEnumVARIANT *pEnum = NULL;

	hr = ADsBuildEnumerator(pContainer, &pEnum);

	if (!SUCCEEDED(hr))
		return _T("");

	VARIANT var;
	VariantInit(&var);
	ULONG ulFetch = 0L;

	CString strNode;
	CString strResult;
	CString strFormatNode;
	BSTR	bstrEmpty = ::SysAllocString(L"");

	while (SUCCEEDED(ADsEnumerateNext(pEnum, 1, &var, &ulFetch)) && ulFetch > 0)
	{
		IADs* pADs;
		CComPtr<IADsContainer> pChildContainer;
		VARIANT _var;

		if (SUCCEEDED(var.pdispVal->QueryInterface(IID_IADs, (void**)&pADs)))
		{
			BSTR _bstr = NULL;
			pADs->get_Name(&_bstr);
			CString strName = OLE2T(_bstr);
			pADs->get_Class(&_bstr);
			CString strClass = OLE2T(_bstr);
			pADs->get_ADsPath(&_bstr);
			CString strPath = OLE2T(_bstr);


			VariantInit(&_var);
			CString strDisplayName = NULL;
			CString strUserAddress = NULL;
			CString strMail = NULL;
			if (strClass.CompareNoCase(_T("organizationalUnit")) == 0)
			{
				hr = pADs->Get(_T("ou"), &_var);
				if (_var.bstrVal == bstrEmpty)
					hr = pADs->Get(_T("displayName"), &_var);
			}
			else
			{
				hr = pADs->Get(_T("displayName"), &_var);
			}
			if (SUCCEEDED(hr))
				strDisplayName = OLE2T(_var.bstrVal);
			hr = pADs->Get(_T("msRTCSIP-PrimaryUserAddress"), &_var);
			if (SUCCEEDED(hr))
				strUserAddress = OLE2T(_var.bstrVal);
			hr = pADs->Get(_T("mail"), &_var);
			if (SUCCEEDED(hr))
				strMail = OLE2T(_var.bstrVal);
			strFormatNode = _T("<%s name=\"%s\" displayName=\"%s\" path=\"%s\" msRTCSIP-PrimaryUserAddress=\"%s\" mail=\"%s\">");
			strNode.Format(strFormatNode, strClass, strName, strDisplayName, strPath, strUserAddress, strMail);
			strResult += strNode;
			VariantClear(&_var);

			if (bRecursive && SUCCEEDED(pADs->QueryInterface(IID_IADsContainer, (void**)&pChildContainer)))
			{
				strResult += ExportAllObjects(pChildContainer, bRecursive);
			}

			strFormatNode = _T("</%s>");
			strNode.Format(strFormatNode, strClass);
			strResult += strNode;
			::SysFreeString(_bstr);
		}

		pADs->Release();
	}

	::SysFreeString(bstrEmpty);

	ADsFreeEnumerator(pEnum);

	VariantClear(&var);

	return strResult;
}

CString CTangramCLRApp::AddOrganizationUnit(CString strPathName, CString strOrgName)
{
	HRESULT hr;

	IADsContainer* pContainer = NULL;

	hr = ADsGetObject(strPathName, IID_IADsContainer, (void**)&pContainer);

	if (!SUCCEEDED(hr))
		return _T("");

	IDispatch* pDisp = NULL;

	CComBSTR bstrOrgName(L"ou=");
	bstrOrgName += strOrgName;

	pContainer->Create(L"organizationalUnit", bstrOrgName, &pDisp);

	IADs* pADs = NULL;

	hr = pDisp->QueryInterface(IID_IADs, (void**)&pADs);

	if (SUCCEEDED(hr))
		pADs->SetInfo();

	BSTR bstrResult;
	pADs->get_ADsPath(&bstrResult);
	CString strResult = OLE2T(bstrResult);

	pADs->Release();
	pDisp->Release();
	pContainer->Release();

	return strResult;
}
void CTangramCLRApp::ImportAllObjects(CString strPathName, CTangramXmlParse* pXmlParse)
{
	int nCount = pXmlParse->GetCount();

	if (nCount > 0)
	{
		CTangramXmlParse* pChildNode = NULL;
		for (int i = 0; i < nCount; i++)
		{
			pChildNode = pXmlParse->GetChild(i);
			CString strClass = pChildNode->name();
			CString strName = pChildNode->attr(_T("name"), _T(""));
			CString strDisplayName = pChildNode->attr(_T("displayName"), _T(""));
			CString strUserName = pChildNode->attr(_T("username"), _T(""));
			if (strClass.CompareNoCase(_T("organizationalUnit")) == 0)
			{
				CString _strPathName = AddOrganizationUnit(strPathName, strName);
				ImportAllObjects(_strPathName, pChildNode);
			}
			else if (strClass.CompareNoCase(_T("user")) == 0)
			{
				//AddUser(strPathName, strName, strDisplayName, strUserName);
				AddUser(strPathName, strDisplayName, pChildNode);
			}
		}
	}

}

CString CTangramCLRApp::AddUser(CString strPathName, CString strName, CTangramXmlParse* pXmlParse)
{
	HRESULT hr;

	IADsContainer* pContainer = NULL;

	hr = ADsGetObject(strPathName, IID_IADsContainer, (void**)&pContainer);

	if (!SUCCEEDED(hr))
		return _T("");

	IDispatch* pDisp = NULL;

	CComBSTR bstrName(L"cn=");
	bstrName += strName;

	pContainer->Create(L"user", bstrName, &pDisp);

	IADs* pADs = NULL;

	hr = pDisp->QueryInterface(IID_IADs, (void**)&pADs);

	if (SUCCEEDED(hr))
	{
		VARIANT var;
		VariantInit(&var);
		CString _propertyValue;
		for (auto _property : adsPropertys)
		{
			_propertyValue = pXmlParse->attr(_property, _T(""));
			if (_propertyValue.CompareNoCase(_T("")) != 0)
			{
				var.bstrVal = _propertyValue.AllocSysString();
				var.vt = VT_BSTR;
				pADs->Put(CComBSTR(_property), var);
				VariantClear(&var);
			}
		}

		pADs->SetInfo();
	}

	BSTR bstrResult;
	pADs->get_ADsPath(&bstrResult);
	CString strResult = OLE2T(bstrResult);

	pADs->Release();
	pDisp->Release();
	pContainer->Release();

	return strResult;
}

CString CTangramCLRApp::AddUser(CString strPathName, CString strName, CString strDisplayName, CString strUsrName)
{
	HRESULT hr;

	if (strName.CompareNoCase(_T("")) == 0)
		strName = strDisplayName;

	IADsContainer* pContainer = NULL;

	hr = ADsGetObject(strPathName, IID_IADsContainer, (void**)&pContainer);

	if (!SUCCEEDED(hr))
		return _T("");

	IDispatch* pDisp = NULL;

	CComBSTR bstrName(L"cn=");
	bstrName += strName;

	pContainer->Create(L"user", bstrName, &pDisp);

	IADs* pADs = NULL;

	hr = pDisp->QueryInterface(IID_IADs, (void**)&pADs);

	if (SUCCEEDED(hr))
	{
		VARIANT var;
		VariantInit(&var);
		var.bstrVal = strDisplayName.AllocSysString();
		var.vt = VT_BSTR;
		pADs->Put(CComBSTR("displayName"), var);
		VariantClear(&var);
		var.bstrVal = strUsrName.AllocSysString();
		var.vt = VT_BSTR;
		pADs->Put(CComBSTR("sAMAccountName"), var);
		VariantClear(&var);

		pADs->SetInfo();
	}

	BSTR bstrResult;
	pADs->get_ADsPath(&bstrResult);
	CString strResult = OLE2T(bstrResult);

	pADs->Release();
	pDisp->Release();
	pContainer->Release();

	return strResult;
}
#endif
#endif

#include <wincrypt.h>

int CTangramCLRApp::CalculateByteMD5(BYTE* pBuffer, int BufferSize, CString &MD5)
{
	HCRYPTPROV hProv;
	// Acquire a cryptographic provider context handle.
	if (CryptAcquireContext(&hProv, NULL, NULL, PROV_RSA_FULL, 0))
	{
		HCRYPTHASH hHash;
		// Create the hash object.
		if (CryptCreateHash(hProv, CALG_MD5, 0, 0, &hHash))
		{
			// Compute the cryptographic hash of the buffer.
			if (CryptHashData(hHash, pBuffer, BufferSize, 0))
			{
				DWORD dwCount = 16;
				unsigned char digest[16];
				CryptGetHashParam(hHash, HP_HASHVAL, digest, &dwCount, 0);

				if (hHash)
					CryptDestroyHash(hHash);

				// Release the provider handle.
				if (hProv)
					CryptReleaseContext(hProv, 0);

				unsigned char b;
				char c;
				char *Value = new char[1024];
				int k = 0;
				for (int i = 0; i < 16; i++)
				{
					b = digest[i];
					for (int j = 4; j >= 0; j -= 4)
					{
						c = ((char)(b >> j) & 0x0F);
						if (c < 10) c += '0';
						else c = ('a' + (c - 10));
						Value[k] = c;
						k++;
					}
				}
				Value[k] = '\0';
				MD5 = CString(Value);
				delete Value;
			}
		}
	}

	return 1;
}

CTangramNodeEvent::CTangramNodeEvent()
{
	m_pWndNode				= nullptr;
	m_pTangramNodeCLREvent	= nullptr;
}

CTangramNodeEvent::~CTangramNodeEvent()
{
	if (m_pTangramNodeCLREvent)
	{
		//LONGLONG nValue = (LONGLONG)m_pWndNode;
		DispEventUnadvise(m_pWndNode);
	}
}

using namespace ATL;
