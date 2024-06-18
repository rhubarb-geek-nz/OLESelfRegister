# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

param($ProgID = 'RhubarbGeekNz.OLESelfRegister', $Method = 'GetMessage', $Hint = 1)

[System.Activator]::CreateInstance([Type]::GetTypeFromProgID($ProgID)).$Method($Hint)
