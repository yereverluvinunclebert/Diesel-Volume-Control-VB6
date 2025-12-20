VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAudioEndpointVolumeCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cAudioEndpointVolumeCallback
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------

'Option Explicit

'Implements IAudioEndpointVolumeCallback

'---------------------------------------------------------------------------------------
' Procedure : IAudioEndpointVolumeCallback_OnNotify
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub IAudioEndpointVolumeCallback_OnNotify(tNotify As AUDIO_VOLUME_NOTIFICATION_DATA)
'
'    On Error GoTo IAudioEndpointVolumeCallback_OnNotify_Error
'
'    'Form1.AppendNotice "EndpointVolCallback: mute=" & tNotify.bMuted
'    'Form1.AppendNotice "Master vol=" & tNotify.fMasterVolume & ",channels=" & tNotify.nChannels
'
'    'changeRemote = True
'
''    If currentVolume <> tNotify.fMasterVolume Then Call changeVolume(tNotify.fMasterVolume)
''    If currentMute <> tNotify.bMuted Then Call changeMute(tNotify.bMuted)
'
'    On Error GoTo 0
'    Exit Sub
'
'IAudioEndpointVolumeCallback_OnNotify_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IAudioEndpointVolumeCallback_OnNotify of Class Module cAudioEndpointVolumeCallback"
'
'End Sub
