/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

#include <displib.h>

static LONG globalUsage;
static HMODULE globalModule;
static OLECHAR globalModuleFileName[260];

typedef struct CHelloWorldData
{
	IHelloWorldVtbl* lpVtbl;
	LONG lUsage;
	IUnknown* lpOuter;
	ITypeInfo* piTypeInfo;
} CHelloWorldData;

static STDMETHODIMP CHelloWorld_IHelloWorld_QueryInterface(IHelloWorld* pThis, REFIID riid, void** ppvObject)
{
	HRESULT hr = E_NOINTERFACE;
	CHelloWorldData* pData = (void*)pThis;

	*ppvObject = NULL;

	if (IsEqualIID(riid, &IID_IDispatch) || IsEqualIID(riid, &IID_IHelloWorld))
	{
		InterlockedIncrement(&pData->lUsage);

		*ppvObject = pThis;

		hr = S_OK;
	}
	else
	{
		if (pData->lpOuter)
		{
			hr = IUnknown_QueryInterface(pData->lpOuter, riid, ppvObject);
		}
		else
		{
			if (IsEqualIID(riid, &IID_IUnknown))
			{
				InterlockedIncrement(&pData->lUsage);

				*ppvObject = pThis;

				hr = S_OK;
			}
		}
	}

	return hr;
}

static STDMETHODIMP_(ULONG) CHelloWorld_IHelloWorld_AddRef(IHelloWorld* pThis)
{
	CHelloWorldData* pData = (void*)pThis;
	return InterlockedIncrement(&pData->lUsage);
}

static STDMETHODIMP_(ULONG) CHelloWorld_IHelloWorld_Release(IHelloWorld* pThis)
{
	CHelloWorldData* pData = (void*)pThis;
	LONG res = InterlockedDecrement(&pData->lUsage);

	if (!res)
	{
		if (pData->piTypeInfo) IUnknown_Release(pData->piTypeInfo);
		if (pData->lpOuter) IUnknown_Release(pData->lpOuter);
		LocalFree(pData);
		InterlockedDecrement(&globalUsage);
	}

	return res;
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetTypeInfoCount(IHelloWorld* pThis, UINT* pctinfo)
{
	*pctinfo = 1;
	return S_OK;
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetTypeInfo(IHelloWorld* pThis, UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	CHelloWorldData* pData = (void*)pThis;
	if (iTInfo) return DISP_E_BADINDEX;
	ITypeInfo_AddRef(pData->piTypeInfo);
	*ppTInfo = pData->piTypeInfo;

	return S_OK;
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetIDsOfNames(IHelloWorld* pThis, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	CHelloWorldData* pData = (void*)pThis;

	if (!IsEqualIID(riid, &IID_NULL))
	{
		return DISP_E_UNKNOWNINTERFACE;
	}

	return DispGetIDsOfNames(pData->piTypeInfo, rgszNames, cNames, rgDispId);
}

static STDMETHODIMP CHelloWorld_IHelloWorld_Invoke(IHelloWorld* pThis, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	CHelloWorldData* pData = (void*)pThis;

	if (!IsEqualIID(riid, &IID_NULL))
	{
		return DISP_E_UNKNOWNINTERFACE;
	}

	return DispInvoke(pThis, pData->piTypeInfo, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetMessage(IHelloWorld* pThis, int Hint, BSTR* lpMessage)
{
	*lpMessage = SysAllocString(Hint < 0 ? L"Goodbye, Cruel World" :  L"Hello World");

	return S_OK;
}

static IHelloWorldVtbl CHelloWorld_IHelloWorldVtbl =
{
	CHelloWorld_IHelloWorld_QueryInterface,
	CHelloWorld_IHelloWorld_AddRef,
	CHelloWorld_IHelloWorld_Release,
	CHelloWorld_IHelloWorld_GetTypeInfoCount,
	CHelloWorld_IHelloWorld_GetTypeInfo,
	CHelloWorld_IHelloWorld_GetIDsOfNames,
	CHelloWorld_IHelloWorld_Invoke,
	CHelloWorld_IHelloWorld_GetMessage
};

static STDMETHODIMP CClassObject_CHelloWorld_IClassFactory_QueryInterface(IClassFactory* pThis, REFIID riid, void** ppvObject)
{
	*ppvObject = NULL;

	if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IClassFactory))
	{
		InterlockedIncrement(&globalUsage);

		*ppvObject = pThis;

		return S_OK;
	}

	return E_NOINTERFACE;
}

static STDMETHODIMP_(ULONG) CClassObject_CHelloWorld_IClassFactory_AddRef(IClassFactory* pThis)
{
	return InterlockedIncrement(&globalUsage);
}

static STDMETHODIMP_(ULONG) CClassObject_CHelloWorld_IClassFactory_Release(IClassFactory* pThis)
{
	return InterlockedDecrement(&globalUsage);
}

static STDMETHODIMP CClassObject_CHelloWorld_IClassFactory_CreateInstance(IClassFactory* pThis, LPVOID punk, REFIID riid, void** ppvObject)
{
	HRESULT hr = E_FAIL;

	if (punk == NULL || IsEqualIID(riid, &IID_IUnknown))
	{
		ITypeLib* piTypeLib = NULL;

		hr = LoadTypeLibEx(globalModuleFileName, REGKIND_NONE, &piTypeLib);

		if (hr >= 0)
		{
			CHelloWorldData* pData = LocalAlloc(LMEM_ZEROINIT, sizeof(*pData));

			if (pData)
			{
				IUnknown* p = (void*)pData;

				InterlockedIncrement(&globalUsage);

				pData->lpVtbl = &CHelloWorld_IHelloWorldVtbl;

				pData->lUsage = 1;
				pData->lpOuter = punk;

				if (pData->lpOuter)
				{
					IUnknown_AddRef(pData->lpOuter);
				}

				hr = ITypeLib_GetTypeInfoOfGuid(piTypeLib, &IID_IHelloWorld, &pData->piTypeInfo);

				if (hr >= 0)
				{
					hr = IUnknown_QueryInterface(p, riid, ppvObject);
				}

				IUnknown_Release(p);
			}

			if (piTypeLib)
			{
				ITypeLib_Release(piTypeLib);
			}
		}
	}

	return hr;
}

static STDMETHODIMP CClassObject_CHelloWorld_IClassFactory_LockServer(IClassFactory* pThis, BOOL bLock)
{
	if (bLock)
	{
		InterlockedIncrement(&globalUsage);
	}
	else
	{
		InterlockedDecrement(&globalUsage);
	}

	return S_OK;
}

static IClassFactoryVtbl CClassObject_CHelloWorld_IClassFactoryVtbl =
{
	CClassObject_CHelloWorld_IClassFactory_QueryInterface,
	CClassObject_CHelloWorld_IClassFactory_AddRef,
	CClassObject_CHelloWorld_IClassFactory_Release,
	CClassObject_CHelloWorld_IClassFactory_CreateInstance,
	CClassObject_CHelloWorld_IClassFactory_LockServer
};

static struct CClassObject {
	IClassFactoryVtbl* lpVtbl;
} CClassObject_CHelloWorld = { &CClassObject_CHelloWorld_IClassFactoryVtbl };

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppvObject)
{
	if (IsEqualCLSID(rclsid, &CLSID_CHelloWorld))
	{
		return IUnknown_QueryInterface((IUnknown*)&CClassObject_CHelloWorld, riid, ppvObject);
	}

	return CLASS_E_CLASSNOTAVAILABLE;
}

STDAPI DllCanUnloadNow(void)
{
	return globalUsage ? S_FALSE : S_OK;
}

BOOL APIENTRY DllMain(HMODULE hModule,
	DWORD  ul_reason_for_call,
	LPVOID lpReserved
)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
		globalModule = hModule;
		DisableThreadLibraryCalls(hModule);

		if (!GetModuleFileNameW(globalModule, globalModuleFileName, sizeof(globalModuleFileName) / sizeof(globalModuleFileName[0])))
		{
			return FALSE;
		}

		break;
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}
