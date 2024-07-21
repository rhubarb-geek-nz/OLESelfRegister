/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

using System;

namespace RhubarbGeekNz.OLESelfRegister
{
    internal class Program
    {
        static void Main(string[] args)
        {
            IHelloWorld helloWorld = Activator.CreateInstance(Type.GetTypeFromProgID("RhubarbGeekNz.OLESelfRegister")) as IHelloWorld;

            Console.WriteLine($"{helloWorld.GetMessage(1)}");
        }
    }
}
