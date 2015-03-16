// ConsoleApplication1.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"


#define   _WIN32_DCOM   
#include   <Wbemidl.h>   
#pragma   comment(lib,"Wbemuuid.lib")  
#include   <atlbase.h>   
#include   "objbase.h"  
#include   <windows.h>   
#include   <comutil.h>   
#pragma   comment(lib,"comsupp.lib")   
#include   <iostream>   

/*
Desp:演示取显卡的显示内存(Video memory)大小(Capacity)
Author:Kagula
Date:2011-11-16
测试环境：WinXP、Win7
编译环境：VS2008+SP1
*/
int main()
{
	CoInitialize(NULL);

	CComPtr <IWbemLocator>   spLoc;
	CComBSTR   bstrNamespace("root\\cimv2");
	CComPtr <IWbemServices>   spServices;
	CComPtr <IWbemClassObject>   spInstance;
	CComPtr <IEnumWbemClassObject>   spEnumInst;
	_variant_t   varProperty;
	CComBSTR   bstrPropName;

	ULONG   uNumOfInstances;
	HRESULT   hrNext;

	if (FAILED(CoInitializeSecurity(NULL, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_PKT, RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, 0)))
		goto _END;

	if (FAILED(spLoc.CoCreateInstance(CLSID_WbemAdministrativeLocator, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER)))
		goto _END;

	if (FAILED(spLoc->ConnectServer(bstrNamespace, NULL, NULL, 0, NULL, 0, 0, &spServices)))
		goto _FAILED;

	if (FAILED(CoSetProxyBlanket(spServices, RPC_C_AUTHN_DEFAULT, RPC_C_AUTHZ_NONE, NULL, RPC_C_AUTHN_LEVEL_CALL,
		RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE)))
		goto _FAILED2;

	if (FAILED(spServices->CreateInstanceEnum(CComBSTR("Win32_VideoController"), WBEM_FLAG_SHALLOW, NULL, &spEnumInst)))
		goto _FAILED2;

	while (true)
	{
		uNumOfInstances = 0;
		hrNext = spEnumInst->Next(10000, 1, &spInstance, &uNumOfInstances);

		if (hrNext == WBEM_S_FALSE)
			break;
		else   if (hrNext == WBEM_S_NO_ERROR)
		{
			/*
			bstrPropName   =   CComBSTR( "PNPDeviceID");
			spInstance-> Get(bstrPropName,0,&varProperty,NULL,0);
			*/
			bstrPropName = CComBSTR("AdapterRAM");
			spInstance->Get(bstrPropName, 0, &varProperty, NULL, 0);
			int nV = varProperty;
			printf("\tVideo Ram:%d MB \n", nV / (1024 * 1024));

			bstrPropName = CComBSTR("Description");
			spInstance->Get(bstrPropName, 0, &varProperty, NULL, 0);
			std::wstring wsT = (wchar_t *)(varProperty.pbstrVal);
			wprintf(L"\tDescription:%s \n", wsT.c_str());
		}
		spInstance.Release();
	}

	spEnumInst.Release();
_FAILED2:
	spServices.Release();
_FAILED:
	spLoc.Release();
_END:
	std::cout << "finished   " << std::endl;
	CoUninitialize();
	system("pause ");
	return   0;
}

