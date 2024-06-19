REM Copyright (c) 2024 Roger Brown.
REM Licensed under the MIT License.

Imports RhubarbGeekNz.OLESelfRegister

Module Program
    Sub Main(args As String())
        Dim helloWorld As IHelloWorld = CreateObject("RhubarbGeekNz.OLESelfRegister")
        Console.WriteLine(helloWorld.GetMessage(1))
    End Sub
End Module
