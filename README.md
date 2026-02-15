Archiver (COM 7-Zip wrapper)
============================

[![CI](https://github.com/BasicNullification/7z-Archiver/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/BasicNullification/7z-Archiver/actions/workflows/ci.yml)
[![Release](https://github.com/BasicNullification/7z-Archiver/actions/workflows/release.yml/badge.svg)](https://github.com/BasicNullification/7z-Archiver/actions/workflows/release.yml)


A minimal COM-visible wrapper around the 7-Zip console (`7z.exe`) focused on creating and updating archives from scripting languages (VBScript, WSH, classic ASP) or VBA. Both x86 and x64 builds are provided. The wrapper delegates all real work to the 7-Zip console and simply builds command strings, writes file lists, and runs the process.

Features
--------
- COM dual interfaces for late binding (ProgID `Archiver.7z`) and early binding (`Archiver.SevenZip` class).
- Supports add (`a`) and update (`u`) archive operations with the same option surface.
- Per-operation settings: compression level, analysis level, memory block (solid) settings, timestamp/attribute storage, and output log level.
- Configurable 7-Zip install directory; defaults to `C:\Program Files\7-Zip` but can be overridden per instance.
- Ships as installer plus standalone DLLs (x86 and x64) for manual registration scenarios.

Requirements
------------
- Windows with 7-Zip installed (uses `7z.exe`). Default path is `C:\Program Files\7-Zip`; set `ProgDirectory` if 7-Zip lives elsewhere.
- .NET Framework 4.8.1 (library target).
- Registration via `regasm` (see below) for manual installs.

Install and registration
------------------------
- Releases: GitHub releases contain an installer plus individual `Archiver.dll` binaries for x86 and x64.
- Manual registration: the release does **not** ship a `.tlb`. If you prefer manual registration, generate the type library yourself and register with the proper regasm for your bitness:

	- x64: `%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe Archiver.dll /codebase /tlb`
	- x86: `%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe Archiver.dll /codebase /tlb`

Usage overview
--------------
- ProgIDs
	- `Archiver.7z` (late binding) / `Archiver.SevenZip` (early binding class name)
	- `Archiver.Add` and `Archiver.Update` are returned from `Add()` / `Update()`
- Typical flow
	1. Create `SevenZip`.
	2. Call `Add(archivePath)` or `Update(archivePath)` to get an operation object.
	3. Queue files with `AddFile` or `AddFiles(Array(...), optionalBaseDir)`; wildcards are allowed with `AddFile`.
	4. Adjust options (compression level, analysis level, logging, memory block, timestamp and attribute storage).
	5. Call `Execute()`.
- Settings
	- `ProgDirectory`: override path to `7z.exe`.
	- `CompressionLevel`: 0-9.
	- `AnalysisLevel`: 0-9.
	- `OutputLogLevel`: `Disabled`, `Files`, `InternalFiles`, `Operations`.
	- `StoreLastModifiedTS`, `StoreCreationTS`, `StoreLastAccessTS`, `StoreFileAttributes`: toggle what metadata is written.
	- `SetMemoryBlockSize(mode, size)`: control solid block; modes include `Disabled`, `Enabled`, `Bytes`, `KiB`, `MiB`, `GiB`, `TiB` (size required for the unit-based modes).

Examples
--------

Late binding (VBScript)
-----------------------

```VBScript
' Example of late-binding usage

Dim o: Set o = CreateObject("Archiver.7z")
Dim a 'As Archiver.Add
Set a = o.Add("V:\Sandbox\MyArchive.7z")

' Can still use .AddFile over .AddFiles for wildcard adds
a.AddFile "V:\Sandbox\To Archive\*.png"

' Pass Array() for files to add, with optional base directory
a.AddFiles Array( _
				"File1.txt", _
				"File2.md" _
		), "V:\Sandbox\To Archive"

a.CompressionLevel = 9
a.AnalysisLevel = 9
a.Execute()
```

Early binding (VBA)
-------------------

```VBA
Sub Archive()

		' Example of early-binding usage (i.e. VBA)
		' Note that class name is SevenZip, not 7z

		Dim o As New Archiver.SevenZip
		Dim a As Archiver.add
		Set a = o.add("V:\Sandbox\MyArchive.7z")
    
		' Can still use .AddFile over .AddFiles for wildcard adds
		a.AddFile "V:\Sandbox\To Archive\*.png"
    
		' Pass Array() for files to add, with optional base directory
		a.AddFiles Array( _
						"File1.txt", _
						"File2.md" _
				), "V:\Sandbox\To Archive"
    
		a.CompressionLevel = 9
		a.AnalysisLevel = 9
		a.Execute

End Sub
```

Build from source
-----------------
- Open `Archiver.slnx` in Visual Studio 2022 (or newer) with .NET Framework 4.8.1 targeting.
- Build the `Archiver` project for x86 and/or x64 to produce `Archiver.dll` and `Archiver.xml` docs.

Notes
-----
- The library does not bundle 7-Zip; it must be installed separately.
- Wildcards are supported via `AddFile`. Use `AddFiles` for explicit lists with an optional shared base directory.
- `OutputLogLevel` maps directly to 7-Zip `-bb` levels; `Disabled` writes no log output.
- `SetMemoryBlockSize` maps to 7-Zip `-ms` solid mode switches.
