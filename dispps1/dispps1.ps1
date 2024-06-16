# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

$ErrorActionPreference = 'Stop'

[System.Activator]::CreateInstance([Type]::GetTypeFromProgID('RhubarbGeekNz.HelloWorld')).GetMessage(1)
