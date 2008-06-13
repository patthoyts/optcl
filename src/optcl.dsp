# Microsoft Developer Studio Project File - Name="optcl" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=optcl - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "optcl.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "optcl.mak" CFG="optcl - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "optcl - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "optcl - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "optcl - Win32 Release Static" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "optcl - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release\Objects"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /YX /FD /c
# ADD CPP /nologo /MD /W3 /GX /Zi /O1 /Ob2 /I "c:\opt\tcl\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /D "USE_TCL_STUBS" /D "USE_TK_STUBS" /D TCL_THREADS=1 /D USE_THREAD_ALLOC=1 /D _REENTRANT=1 /D _THREAD_SAFE=1 /FR /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x809 /d "NDEBUG"
# ADD RSC /l 0x809 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386
# ADD LINK32 tclstub84.lib tkstub84.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib /nologo /dll /debug /machine:I386 /out:"Release/optcl30.dll" /libpath:"c:\opt\tcl\lib"
# SUBTRACT LINK32 /incremental:yes

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug\Objects"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /YX /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /I "c:\opt\tcl\include" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /D "USE_TCL_STUBS" /D "USE_TK_STUBS" /D "ATL_DEBUG_INTERFACES" /FR /YX"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x809 /d "_DEBUG"
# ADD RSC /l 0x809 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 tclstub84.lib tkstub84.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib /nologo /dll /debug /machine:I386 /out:"Debug/optcl30g.dll" /pdbtype:sept /libpath:"c:\opt\tcl\lib"

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release\Static"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release\Static"
# PROP Target_Dir ""
# ADD BASE CPP /nologo /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_MBCS" /D "_LIB" /YX /FD /c
# ADD CPP /nologo /MD /W3 /GX /Zi /O1 /Ob2 /I "c:\opt\tcl\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_LIB" /D "OPTCL_EXPORTS" /D TCL_THREADS=1 /D USE_THREAD_ALLOC=1 /D _REENTRANT=1 /D _THREAD_SAFE=1 /FR /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x809 /d "NDEBUG"
# ADD RSC /l 0x809 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /machine:IX86
# ADD LINK32 /nologo /machine:IX86 /out:"Release\optcl30s.lib" /lib

!ENDIF 

# Begin Target

# Name "optcl - Win32 Release"
# Name "optcl - Win32 Debug"
# Name "optcl - Win32 Release Static"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\ComRecordInfoImpl.cpp
# End Source File
# Begin Source File

SOURCE=.\Container.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\DispParams.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\EventBinding.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\initonce.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# SUBTRACT CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\ObjMap.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\optcl.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\OptclBindPtr.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\OptclObj.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\OptclTypeAttr.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\resource.rc
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# Begin Source File

SOURCE=.\typelib.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\utility.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /YX"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /YX

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# ADD BASE CPP /YX"stdafx.h"
# ADD CPP /YX"stdafx.h"

!ENDIF 

# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\ComRecordInfoImpl.h
# End Source File
# Begin Source File

SOURCE=.\Container.h
# End Source File
# Begin Source File

SOURCE=.\DispParams.h
# End Source File
# Begin Source File

SOURCE=.\EventBinding.h
# End Source File
# Begin Source File

SOURCE=.\ObjMap.h
# End Source File
# Begin Source File

SOURCE=.\optcl.h
# End Source File
# Begin Source File

SOURCE=.\OptclBindPtr.h
# End Source File
# Begin Source File

SOURCE=.\OptclObj.h
# End Source File
# Begin Source File

SOURCE=.\OptclTypeAttr.h
# End Source File
# Begin Source File

SOURCE=.\resource.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# Begin Source File

SOURCE=.\tbase.h
# End Source File
# Begin Source File

SOURCE=.\typelib.h
# End Source File
# Begin Source File

SOURCE=.\utility.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=.\ImageListBox.tcl

!IF  "$(CFG)" == "optcl - Win32 Release"

# Begin Custom Build
InputPath=.\ImageListBox.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# Begin Custom Build
InputPath=.\ImageListBox.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# Begin Custom Build
InputPath=.\ImageListBox.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\optcl.tcl

!IF  "$(CFG)" == "optcl - Win32 Release"

# Begin Custom Build
InputPath=.\optcl.tcl

"c:\progra~1\tcl\lib\optcl\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# Begin Custom Build
InputPath=.\optcl.tcl

"c:\progra~1\tcl\lib\optcl\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# Begin Custom Build
InputPath=.\optcl.tcl

"c:\progra~1\tcl\lib\optcl\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\$(InputPath)

# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Splitter.tcl

!IF  "$(CFG)" == "optcl - Win32 Release"

# Begin Custom Build
InputPath=.\Splitter.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# Begin Custom Build
InputPath=.\Splitter.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# Begin Custom Build
InputPath=.\Splitter.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\TLView.tcl

!IF  "$(CFG)" == "optcl - Win32 Release"

# Begin Custom Build
InputPath=.\TLView.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# Begin Custom Build
InputPath=.\TLView.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# Begin Custom Build
InputPath=.\TLView.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Tooltip.tcl

!IF  "$(CFG)" == "optcl - Win32 Release"

# Begin Custom Build
InputPath=.\Tooltip.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# Begin Custom Build
InputPath=.\Tooltip.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# Begin Custom Build
InputPath=.\Tooltip.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\typelib.tcl

!IF  "$(CFG)" == "optcl - Win32 Release"

# Begin Custom Build
InputPath=.\typelib.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# Begin Custom Build
InputPath=.\typelib.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# Begin Custom Build
InputPath=.\typelib.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\Utilities.tcl

!IF  "$(CFG)" == "optcl - Win32 Release"

# Begin Custom Build
InputPath=.\Utilities.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# Begin Custom Build
InputPath=.\Utilities.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ELSEIF  "$(CFG)" == "optcl - Win32 Release Static"

# Begin Custom Build
InputPath=.\Utilities.tcl

"c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	copy $(InputPath) c:\progra~1\tcl\lib\optcl\scripts\$(InputPath)

# End Custom Build

!ENDIF 

# End Source File
# End Group
# Begin Source File

SOURCE=.\test.tcl
# End Source File
# End Target
# End Project
