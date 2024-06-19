/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

#include<windows.h>
#undef GetMessage
#include <displib.h>
#include <stdio.h>

int main(int argc, char** argv)
{
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	if (SUCCEEDED(hr))
	{
#ifdef _DEBUG
		REFCLSID clsid = CLSID_CHelloWorld;
		BSTR libName = SysAllocString(L"DISPLIB.DLL");
		HINSTANCE hModule = CoLoadLibrary(libName, TRUE);
		LPFNGETCLASSOBJECT pfn = (LPFNGETCLASSOBJECT)GetProcAddress(hModule, "DllGetClassObject");
		IUnknown* classObject = NULL;
		DWORD dwRegister = 0;

		SysFreeString(libName);

		hr = pfn(clsid, IID_IUnknown, (void**)&classObject);

		if (SUCCEEDED(hr))
		{
			hr = CoRegisterClassObject(clsid, classObject, CLSCTX_INPROC_SERVER, REGCLS_MULTIPLEUSE, &dwRegister);
			classObject->Release();
		}
#else
		BSTR app = SysAllocString(L"RhubarbGeekNz.OLESelfRegister");
		CLSID clsid;

		hr = CLSIDFromProgID(app, &clsid);

		SysFreeString(app);
#endif

		if (SUCCEEDED(hr))
		{
			IHelloWorld* helloWorld = NULL;

			hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IHelloWorld, (void**)&helloWorld);

			if (SUCCEEDED(hr))
			{
				BSTR bstr = NULL;

				hr = helloWorld->GetMessage(1, &bstr);

				if (SUCCEEDED(hr))
				{
					printf("%S\n", bstr);

					SysFreeString(bstr);
				}

				helloWorld->Release();
			}

#ifdef _DEBUG
			CoRevokeClassObject(dwRegister);
#endif
		}

		CoUninitialize();
	}

	if (FAILED(hr))
	{
		fprintf(stderr, "0x%lx\n", (long)hr);
	}

	return hr < 0;
}
