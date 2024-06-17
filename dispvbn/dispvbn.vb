REM Copyright (c) 2024 Roger Brown.
REM Licensed under the MIT License.

Imports RhubarbGeekNzDispLib

Module Program
    Sub Main(args As String())
        Dim helloWorld As IHelloWorld = CreateObject("RhubarbGeekNz.HelloWorld")
        Console.WriteLine(helloWorld.GetMessage(1))
    End Sub
End Module
