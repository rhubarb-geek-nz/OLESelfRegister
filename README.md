# rhubarb-geek-nz/OLESelfRegister

Demonstration of OLESelfRegister COM object.

```
Windows Registry Editor Version 5.00

[HKEY_CLASSES_ROOT\RhubarbGeekNz.OLESelfRegister\CLSID]
@="{74067A54-5BAF-4CD0-BF21-DB6A9EFA6CA7}"

[HKEY_CLASSES_ROOT\CLSID\{74067A54-5BAF-4CD0-BF21-DB6A9EFA6CA7}\InprocServer32]
@="C:\\wherever\\you\\put\\it\\RhubarbGeekNzOLESelfRegister.dll"
"ThreadingModel"="Both"
```

[displib.idl](displib/displib.idl) defines the dual-interface for a simple inprocess server.

[displib.c](displib/displib.c) implements the interface.

[dispapp.cpp](dispapp/dispapp.cpp) creates an instance with [CoCreateInstance](https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance) and uses it to get a message to display.

[dispnet.cs](dispnet/dispnet.cs) demonstrates using import library.

[package.ps1](package.ps1) is used to automate the building of multiple architectures.

[dispvbn.vb](dispvbn/dispvbn.vb) shows Hello World usage in Visual Basic

[dispps1.ps1](dispps1/dispps1.ps1) PowerShell creating and invoking Hello World
