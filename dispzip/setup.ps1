# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

param([switch]$UnregServer)

trap
{
	throw $PSItem
}

$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

$ProcessArchitecture = $Env:PROCESSOR_ARCHITECTURE.ToLower()

switch ($ProcessArchitecture)
{
	'amd64' { $ProcessArchitecture = 'x64' }
}

$ProgramFiles = $Env:ProgramFiles

$CompanyDir = Join-Path -Path $ProgramFiles -ChildPath 'rhubarb-geek-nz'
$ProductDir = Join-Path -Path $CompanyDir -ChildPath 'OLESelfRegister'
$InstallDir = Join-Path -Path $ProductDir -ChildPath $ProcessArchitecture
$DllName = 'RhubarbGeekNzOLESelfRegister.dll'
$DllPath = Join-Path -Path $InstallDir -ChildPath $DllName

Add-Type -TypeDefinition @"
	using System;
	using System.ComponentModel;
	using System.Runtime.InteropServices;

	namespace RhubarbGeekNz.OLESelfRegister
	{
		public class InterOp
		{
			[UnmanagedFunctionPointer(CallingConvention.StdCall)]
			public delegate int DllRegisterServerDelegate();

			[DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
			private static extern IntPtr LoadLibraryW(string lpszLibName);

			[DllImport("kernel32.dll", SetLastError = true)]
			private static extern int FreeLibrary(IntPtr hModule);

			[DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
			private static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

			public static void InvokeExportedProc(string dllname, string procname)
			{
				IntPtr hModule = LoadLibraryW(dllname);

				if (hModule == IntPtr.Zero)
				{
					throw new Win32Exception(Marshal.GetLastWin32Error());
				}

				try
				{
					IntPtr intPtr = GetProcAddress(hModule, procname);

					if (intPtr == IntPtr.Zero)
					{
						throw new Win32Exception(Marshal.GetLastWin32Error());
					}

					DllRegisterServerDelegate dllRegisterServerDelegate = Marshal.GetDelegateForFunctionPointer(intPtr, typeof(DllRegisterServerDelegate)) as DllRegisterServerDelegate;

					Marshal.ThrowExceptionForHR(dllRegisterServerDelegate());
				}
				finally
				{
					FreeLibrary(hModule);
				}
			}
		}
	}
"@

if ($UnregServer)
{
	if (Test-Path $DllPath)
	{
		[RhubarbGeekNz.OLESelfRegister.Interop]::InvokeExportedProc($DllPath, 'DllUnregisterServer')
	}

	$DllPath, $InstallDir, $ProductDir | ForEach-Object {
		$FilePath = $_
		if (Test-Path $FilePath)
		{
			Remove-Item -LiteralPath $FilePath
		}
	}

	if (Test-Path $CompanyDir)
	{
		$children = Get-ChildItem -LiteralPath $CompanyDir

		if (-not $children)
		{
			Remove-Item -LiteralPath $CompanyDir
		}
	}
}
else
{
	if (Test-Path $DllPath)
	{
		Write-Warning "$DllPath is already installed"
	}
	else
	{
		$SourceDir = Join-Path -Path $PSScriptRoot -ChildPath $ProcessArchitecture
		$SourceFile = Join-Path -Path $SourceDir -ChildPath $DllName

		if (-not (Test-Path $SourceFile))
		{
			Write-Error "$SourceFile does not exist"
		}
		else
		{
			$CompanyDir, $ProductDir, $InstallDir | ForEach-Object {
				$FilePath = $_
				if (-not (Test-Path $FilePath))
				{
					$Null = New-Item -Path $FilePath -ItemType 'Directory'
				}
			}

			Copy-Item $SourceFile $DllPath

			[RhubarbGeekNz.OLESelfRegister.Interop]::InvokeExportedProc($DllPath, 'DllRegisterServer')
		}
	}
}
