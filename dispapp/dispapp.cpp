/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

#include<windows.h>
#undef GetMessage
#include <displib_h.h>
#include <stdio.h>

int main(int argc, char** argv)
{
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

	if (SUCCEEDED(hr))
	{
		ITypeLib* typeLib;
		hr = LoadRegTypeLib(LIBID_RhubarbGeekNzOLESelfRegister, 1, 0, 0, &typeLib);
		if (SUCCEEDED(hr))
		{
			typeLib->Release();
		}
	}

	if (SUCCEEDED(hr))
	{
		BSTR app = SysAllocString(L"RhubarbGeekNz.OLESelfRegister");
		CLSID clsid;

		hr = CLSIDFromProgID(app, &clsid);

		SysFreeString(app);

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
		}

		CoUninitialize();
	}

	if (FAILED(hr))
	{
		fprintf(stderr, "0x%lx\n", (long)hr);
	}

	return FAILED(hr);
}
