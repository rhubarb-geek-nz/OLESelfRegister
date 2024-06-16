Imports System
Imports RhubarbGeekNzDispLib

Module Program

    Sub Main(args As String())
        Dim helloWorld As IHelloWorld = Activator.CreateInstance(Type.GetTypeFromProgID("RhubarbGeekNz.HelloWorld"))
        Console.WriteLine(helloWorld.GetMessage(1))
    End Sub
End Module
