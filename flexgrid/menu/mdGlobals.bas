Attribute VB_Name = "mdGlobals"
'==============================================================================
' mdGlobals.bas
'
'   Subclassing Thunk (SuperClass V2) Project Samples
'   Copyright (c) 2002 by Vlad Vissoultchev <wqweto@myrealbox.com>
'
'   Menu hook impl encapsulation
'
' Modifications:
'
' 2002-10-28    WQW     Initial implementation
'
'==============================================================================
Option Explicit

Public g_oMenuHook      As cHookingThunk
Public g_oMenuHookImpl  As cMenuHook
Public g_oCurrentMenu   As ctxHookMenu

#If DebugMode Then
Public g_lObjCount          As Long
#End If

