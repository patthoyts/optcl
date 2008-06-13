# Microsoft Developer Studio Project File - Name="optcl" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=optcl - Win32 Debug_NoStubs
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "optcl.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "optcl.mak" CFG="optcl - Win32 Debug_NoStubs"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "optcl - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "optcl - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "optcl - Win32 Release_NoStubs" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "optcl - Win32 Debug_NoStubs" (based on "Win32 (x86) Dynamic-Link Library")
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
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /YX /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /Ob2 /I "c:\progra~1\tcl\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /D "USE_TCL_STUBS" /D "USE_TK_STUBS" /FR /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x809 /d "NDEBUG"
# ADD RSC /l 0x809 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib tclstub82.lib tkstub82.lib /nologo /dll /machine:I386 /out:"../install/optclstubs.dll" /libpath:"c:\progra~1\tcl\lib"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "optcl___Win32_Debug"
# PROP BASE Intermediate_Dir "optcl___Win32_Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /YX /FD /GZ /c
# ADD CPP /nologo /MTd /W3 /Gm /Gi /GX /ZI /Od /I "c:\progra~1\tcl\include" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /D "USE_TCL_STUBS" /D "USE_TK_STUBS" /FR /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x809 /d "_DEBUG"
# ADD RSC /l 0x809 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib tclstub84.lib tkstub84.lib /nologo /dll /debug /machine:I386 /out:"../install/optclstubs.dll" /pdbtype:sept /libpath:"c:\progra~1\tcl\lib"

!ELSEIF  "$(CFG)" == "optcl - Win32 Release_NoStubs"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "optcl___Win32_Release_NoStubs"
# PROP BASE Intermediate_Dir "optcl___Win32_Release_NoStubs"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release_NoStubs"
# PROP Intermediate_Dir "Release_NoStubs"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /Ob2 /I "c:\progra~1\tcl\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /D "USE_TCL_STUBS" /D "USE_TK_STUBS" /FR /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MT /W3 /GX /O2 /Ob2 /I "c:\progra~1\tcl\include" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /FR /Yu"stdafx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x809 /d "NDEBUG"
# ADD RSC /l 0x809 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib tclstub82.lib tkstub82.lib /nologo /dll /machine:I386 /libpath:"c:\progra~1\tcl\lib"
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib tcl80.lib tk80.lib /nologo /dll /machine:I386 /out:"../install/optcl80.dll" /libpath:"c:\progra~1\tcl\lib"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug_NoStubs"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "optcl___Win32_Debug_NoStubs"
# PROP BASE Intermediate_Dir "optcl___Win32_Debug_NoStubs"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug_NoStubs"
# PROP Intermediate_Dir "Debug_NoStubs"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /Gi /GX /ZI /Od /I "c:\progra~1\tcl\include" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /D "USE_TCL_STUBS" /D "USE_TK_STUBS" /FR /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MTd /W3 /Gm /Gi /GX /ZI /Od /I "c:\progra~1\tcl\include" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "OPTCL_EXPORTS" /FR /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x809 /d "_DEBUG"
# ADD RSC /l 0x809 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib tclstub82.lib tkstub82.lib /nologo /dll /debug /machine:I386 /pdbtype:sept /libpath:"c:\progra~1\tcl\lib"
# ADD LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib tcl80.lib tk80.lib /nologo /dll /debug /machine:I386 /out:"../install/optcl80.dll" /pdbtype:sept /libpath:"c:\progra~1\tcl\lib"

!ENDIF 

# Begin Target

# Name "optcl - Win32 Release"
# Name "optcl - Win32 Debug"
# Name "optcl - Win32 Release_NoStubs"
# Name "optcl - Win32 Debug_NoStubs"
# Begin Group "Source"

# PROP Default_Filter "cpp"
# Begin Source File

SOURCE=.\Container.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /Yu"StdAfx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

!ELSEIF  "$(CFG)" == "optcl - Win32 Release_NoStubs"

# ADD BASE CPP /Yu"StdAfx.h"
# ADD CPP /Yu"StdAfx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug_NoStubs"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\DispParams.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /Yu"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /Yu"StdAfx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Release_NoStubs"

# ADD BASE CPP /Yu"stdafx.h"
# ADD CPP /Yu"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug_NoStubs"

# ADD BASE CPP /Yu"StdAfx.h"
# ADD CPP /Yu"StdAfx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\EventBinding.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /Yu"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

!ELSEIF  "$(CFG)" == "optcl - Win32 Release_NoStubs"

# ADD BASE CPP /Yu"stdafx.h"
# ADD CPP /Yu"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug_NoStubs"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\initonce.cpp
# SUBTRACT CPP /YX /Yc /Yu
# End Source File
# Begin Source File

SOURCE=.\ObjMap.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /Yu"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /Yu"StdAfx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Release_NoStubs"

# ADD BASE CPP /Yu"stdafx.h"
# ADD CPP /Yu"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug_NoStubs"

# ADD BASE CPP /Yu"StdAfx.h"
# ADD CPP /Yu"StdAfx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\optcl.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /Yu"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /Yu"StdAfx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Release_NoStubs"

# ADD BASE CPP /Yu"stdafx.h"
# ADD CPP /Yu"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug_NoStubs"

# ADD BASE CPP /Yu"StdAfx.h"
# ADD CPP /Yu"StdAfx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\OptclBindPtr.cpp
# End Source File
# Begin Source File

SOURCE=.\OptclObj.cpp

!IF  "$(CFG)" == "optcl - Win32 Release"

# ADD CPP /Yu"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# ADD CPP /Yu"StdAfx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Release_NoStubs"

# ADD BASE CPP /Yu"stdafx.h"
# ADD CPP /Yu"stdafx.h"

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug_NoStubs"

# ADD BASE CPP /Yu"StdAfx.h"
# ADD CPP /Yu"StdAfx.h"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=.\OptclTypeAttr.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"StdAfx.h"
# End Source File
# Begin Source File

SOURCE=.\typelib.cpp
# ADD CPP /Yu"StdAfx.h"
# End Source File
# Begin Source File

SOURCE=.\utility.cpp
# ADD CPP /Yu"StdAfx.h"
# End Source File
# End Group
# Begin Group "Header"

# PROP Default_Filter "h"
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
# Begin Group "Resource"

# PROP Default_Filter ""
# Begin Source File

SOURCE=.\resource.rc
# End Source File
# Begin Source File

SOURCE=.\typelib.tcl
# End Source File
# End Group
# Begin Source File

SOURCE=.\test.tcl

!IF  "$(CFG)" == "optcl - Win32 Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "optcl - Win32 Release_NoStubs"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "optcl - Win32 Debug_NoStubs"

# PROP BASE Exclude_From_Build 1
# PROP Exclude_From_Build 1

!ENDIF 

# End Source File
# End Target
# End Project
