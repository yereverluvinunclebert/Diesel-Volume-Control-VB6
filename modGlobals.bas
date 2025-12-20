Attribute VB_Name = "modGlobals"
Option Explicit

Private m_bCurrentMute As Boolean
Private m_dCurrentVolume As Double
Private m_bChangeRemote As Boolean
Private m_bStartupFlg As Boolean

' needs a project reference to Fafalone's oleexp.tlb Modern Shell Interfaces for VB6

'---------------------------------------------------------------------------------------
' Procedure : changeRemote
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get changeRemote() As Boolean

    On Error GoTo changeRemote_Error

    changeRemote = m_bChangeRemote

    On Error GoTo 0
    Exit Property

changeRemote_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeRemote of Module modGlobals"

End Property

'---------------------------------------------------------------------------------------
' Procedure : changeRemote
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let changeRemote(ByVal bchangeRemote As Boolean)

    On Error GoTo changeRemote_Error

    m_bChangeRemote = bchangeRemote

    On Error GoTo 0
    Exit Property

changeRemote_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure changeRemote of Module modGlobals"

End Property

'---------------------------------------------------------------------------------------
' Procedure : currentMute
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get currentMute() As Boolean

    On Error GoTo currentMute_Error

    currentMute = m_bCurrentMute

    On Error GoTo 0
    Exit Property

currentMute_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure currentMute of Module modGlobals"

End Property

'---------------------------------------------------------------------------------------
' Procedure : currentMute
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let currentMute(ByVal bcurrentMute As Boolean)

    On Error GoTo currentMute_Error

    m_bCurrentMute = bcurrentMute

    On Error GoTo 0
    Exit Property

currentMute_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure currentMute of Module modGlobals"

End Property

'---------------------------------------------------------------------------------------
' Procedure : currentVolume
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get currentVolume() As Double

    On Error GoTo currentVolume_Error

    currentVolume = m_dCurrentVolume

    On Error GoTo 0
    Exit Property

currentVolume_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure currentVolume of Module modGlobals"

End Property

'---------------------------------------------------------------------------------------
' Procedure : currentVolume
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let currentVolume(ByVal dcurrentVolume As Double)

    On Error GoTo currentVolume_Error

    m_dCurrentVolume = dcurrentVolume

    On Error GoTo 0
    Exit Property

currentVolume_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure currentVolume of Module modGlobals"

End Property

'---------------------------------------------------------------------------------------
' Procedure : bStartupFlg
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get bStartupFlg() As Boolean

    On Error GoTo bStartupFlg_Error

    bStartupFlg = m_bStartupFlg

    On Error GoTo 0
    Exit Property

bStartupFlg_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bStartupFlg of Module modGlobals"

End Property

'---------------------------------------------------------------------------------------
' Procedure : bStartupFlg
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let bStartupFlg(ByVal bStartupFlg As Boolean)

    On Error GoTo bStartupFlg_Error

    m_bStartupFlg = bStartupFlg

    On Error GoTo 0
    Exit Property

bStartupFlg_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bStartupFlg of Module modGlobals"

End Property
