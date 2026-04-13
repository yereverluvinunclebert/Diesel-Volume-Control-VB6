Attribute VB_Name = "modAudioCode"
'---------------------------------------------------------------------------------------
' Module    : modAudioCode
' Author    : fafalone
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private mDeviceEnum As MMDeviceEnumerator
Private pEPVolMM As IAudioEndpointVolume
'Private cVolCallback As cAudioEndpointVolumeCallback
Private mDefRenderMM As IMMDevice


'---------------------------------------------------------------------------------------
' Procedure : GetDeviceName
' Author    : Fafalone
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetDeviceName(pCol As IMMDeviceCollection, nIdx As Long) As String
    Dim pDevice As IMMDevice
    Dim sID As String
    Dim pStore As IPropertyStore
    Dim pDesc As IPropertyDescription
    Dim lp As Long
    Dim lpID As Long
    Dim vrProp As Variant
    Dim vProp As Variant
    Dim j As Long
    On Error GoTo GetDeviceName_Error

    If (pCol.Item(nIdx, pDevice)) = S_OK Then
        pDevice.GetId lpID
        sID = LPWSTRtoStr(lpID)
        Debug.Print "Got device(" & nIdx & ") id=" & sID
        If (sID = "") Or (sID = vbNullChar) Then
            Debug.Print "Invalid DeviceID"
            Exit Function
        End If
        pDevice.OpenPropertyStore STGM_READ, pStore
        If (pStore Is Nothing) = False Then
            'these property stores aren't as full-featured as other ones
            'such as those associated with IShellItem's of files.. for
            'example we can't use the superior PSFormatPropertyValue b/c
            'we can't get an IPropertyDescription, or a property name
            'So we're stuck using PropVariants, a nightmare in VB
            Dim pcnt As Long
            pStore.GetCount pcnt
            Debug.Print "prop cnt=" & pcnt
            Debug.Print "outputting propvariant..."
            Dim pk As PROPERTYKEY
            pStore.GetValue PKEY_Device_FriendlyName, vProp
            PropVariantToVariant vProp, vrProp
            Dim vte As VbVarType
            vte = VarType(vrProp)
            If (vte And vbArray) = vbArray Then 'this always seems to be vbString and vbArray, haven't encountered other types
                Debug.Print "PV type is array"
                For j = LBound(vrProp) To UBound(vrProp)
                    Debug.Print "Value(" & j & ")=" & CStr(vrProp(j))
                Next j
    
            Else
                Select Case vte
                    Case vbDataObject, vbObject, vbUserDefinedType
                        GetDeviceName = "<cannot display this type>"
                    Case vbEmpty, vbNull
                        GetDeviceName = "<empty or null>"
                    Case vbError
                        GetDeviceName = "<vbError>"
                    Case Else
                        GetDeviceName = CStr(vrProp)
                End Select
            End If
    '        Debug.Print "enum pstore..."
    '        dbg_enumstore pStore
        Else
            Debug.Print "Failed to get IPropertyStore"
        End If
    Else
        Debug.Print "Failed to get device with pCol.Item, nIdx=" & nIdx
    End If

    On Error GoTo 0
    Exit Function

GetDeviceName_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDeviceName of Form Form1"

End Function




'---------------------------------------------------------------------------------------
' Procedure : GetDeviceNameDirect
' Author    : Fafalone
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function GetDeviceNameDirect(pDevice As IMMDevice) As String
    Dim sID As String
    Dim pStore As IPropertyStore
    Dim pDesc As IPropertyDescription
    Dim lp As Long
    Dim lpID As Long
    Dim vrProp As Variant
    Dim vProp As Variant
    Dim j As Long
    On Error GoTo GetDeviceNameDirect_Error

    pDevice.GetId lpID
    sID = LPWSTRtoStr(lpID)
    Debug.Print "Got device id=" & sID
    If (sID = "") Or (sID = vbNullChar) Then
        Debug.Print "Invalid DeviceID"
        Exit Function
    End If
    pDevice.OpenPropertyStore STGM_READ, pStore
    If (pStore Is Nothing) = False Then
        'these property stores aren't as full-featured as other ones
        'such as those associated with IShellItem's of files.. for
        'example we can't use the superior PSFormatPropertyValue b/c
        'we can't get an IPropertyDescription, or a property name
        'So we're stuck using PropVariants, a nightmare in VB
        Dim pcnt As Long
        pStore.GetCount pcnt
        Debug.Print "prop cnt=" & pcnt
        Debug.Print "outputting propvariant..."
        Dim pk As PROPERTYKEY
        pStore.GetValue PKEY_Device_FriendlyName, vProp
        PropVariantToVariant vProp, vrProp
        Dim vte As VbVarType
        vte = VarType(vrProp)
        If (vte And vbArray) = vbArray Then 'this always seems to be vbString and vbArray, haven't encountered other types
            Debug.Print "PV type is array"
            For j = LBound(vrProp) To UBound(vrProp)
                Debug.Print "Value(" & j & ")=" & CStr(vrProp(j))
            Next j

        Else
            Select Case vte
                Case vbDataObject, vbObject, vbUserDefinedType
                    GetDeviceNameDirect = "<cannot display this type>"
                Case vbEmpty, vbNull
                    GetDeviceNameDirect = "<empty or null>"
                Case vbError
                    GetDeviceNameDirect = "<vbError>"
                Case Else
                    GetDeviceNameDirect = CStr(vrProp)
            End Select
        End If
'        Debug.Print "enum pstore..."
'        dbg_enumstore pStore
    Else
        Debug.Print "Failed to get IPropertyStore"
    End If


    On Error GoTo 0
    Exit Function

GetDeviceNameDirect_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetDeviceNameDirect of Form Form1"

End Function



'---------------------------------------------------------------------------------------
' Procedure : sliVolControl_ScrollCode
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub sliVolControl_ScrollCode(ByVal sliVolControlValue As Integer)
    
    On Error GoTo sliVolControl_ScrollCode_Error

    If changeRemote = True Then
        changeRemote = False
        Exit Sub
    End If

    If bStartupFlg = True Then Exit Sub

    If (pEPVolMM Is Nothing) Then
        Set mDeviceEnum = New MMDeviceEnumerator
        mDeviceEnum.GetDefaultAudioEndpoint eRender, eMultimedia, mDefRenderMM
    
        If (mDefRenderMM Is Nothing) = False Then
            mDefRenderMM.Activate IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, CVar(0), pEPVolMM
        End If
    End If
    
    If (pEPVolMM Is Nothing) = False Then
        pEPVolMM.SetMasterVolumeLevelScalar CSng(sliVolControlValue / 100), UUID_NULL
    End If
    
    'Form1.AppendNotice "sliVolControl_Scroll Value = " & sliVolControlValue

    On Error GoTo 0
    Exit Sub

sliVolControl_ScrollCode_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sliVolControl_ScrollCode of Form Form1"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : getVol
' Author    : Fafalone
' Date      : 19/12/2025
' Purpose   : get the volume and set the volume level of the default audio output
'             called only at form_load and by timers
'---------------------------------------------------------------------------------------
'
Public Function getVol() As Long
    Dim sOut As String
    
    On Error GoTo getVol_Error

    Set mDeviceEnum = New MMDeviceEnumerator
    mDeviceEnum.GetDefaultAudioEndpoint eRender, eMultimedia, mDefRenderMM
    
    If (mDefRenderMM Is Nothing) = False Then
        
        mDefRenderMM.Activate IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, CVar(0), pEPVolMM
        If (pEPVolMM Is Nothing) = False Then
            Dim sngVol As Single
            pEPVolMM.GetMasterVolumeLevelScalar sngVol
            ' causes crash
'            If (cVolCallback Is Nothing) Then
'                Set cVolCallback = New cAudioEndpointVolumeCallback
'                pEPVolMM.RegisterControlChangeNotify cVolCallback
'                AppendNotice "Callback re-registered in GetVol."
'            End If
        Else
            'Form1.AppendNotice "Couldn't get endpoint volume control."
        End If
    End If

    getVol = CLng(sngVol * 100)
    
    On Error GoTo 0
    Exit Function

getVol_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getVol of Form Form1"

End Function




'---------------------------------------------------------------------------------------
' Procedure : getmute
' Author    : Fafalone
' Date      : 19/12/2025
' Purpose   : obtains the initial mute status, called only at form_load and by timers
'---------------------------------------------------------------------------------------
'
Public Function getmute() As Boolean
    Dim sOut As String
    
    On Error GoTo getmute_Error

    If (pEPVolMM Is Nothing) Then
        Set mDeviceEnum = New MMDeviceEnumerator
        mDeviceEnum.GetDefaultAudioEndpoint eRender, eMultimedia, mDefRenderMM
    
        If (mDefRenderMM Is Nothing) = False Then
            mDefRenderMM.Activate IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, CVar(0), pEPVolMM
            If (pEPVolMM Is Nothing) = False Then
'                If (cVolCallback Is Nothing) Then
'                    Set cVolCallback = New cAudioEndpointVolumeCallback
'                    pEPVolMM.RegisterControlChangeNotify cVolCallback
'                    AppendNotice "Callback re-registered in Getmute."
'                End If
                'Form1.AppendNotice "Getting Mute"
            Else
               'Form1.AppendNotice "Couldn't get endpoint volume control."
            End If
        End If
    End If
    Dim fMuted As BOOL
    If (pEPVolMM Is Nothing) = False Then
        pEPVolMM.getmute fMuted
        If fMuted Then getmute = True
    End If
    
    On Error GoTo 0
    Exit Function

getmute_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getmute of Form Form1"

End Function





'---------------------------------------------------------------------------------------
' Procedure : btnMuteCode
' Author    : beededea
' Date      : 19/12/2025
' Purpose   : mute the sound
'---------------------------------------------------------------------------------------
'
Public Sub btnMuteCode()

    On Error GoTo btnMuteCode_Error

        If getmute = True Then
            unMute
        Else
            muteIt
        End If

    On Error GoTo 0
    Exit Sub

btnMuteCode_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure btnMuteCode of Form Form1"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : unMute
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub unMute()

    On Error GoTo unMute_Error

    Set mDeviceEnum = New MMDeviceEnumerator
    
    mDeviceEnum.GetDefaultAudioEndpoint eRender, eMultimedia, mDefRenderMM
    
    If (mDefRenderMM Is Nothing) = False Then

        mDefRenderMM.Activate IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, CVar(0), pEPVolMM
        If (pEPVolMM Is Nothing) = False Then
            If pEPVolMM.SetMute(0, UUID_NULL) = S_OK Then
                'Form1.AppendNotice "unMuted " & GetDeviceNameDirect(mDefRenderMM)
                'Form1.btnMute.Caption = "Mute"
            Else
                'Form1.AppendNotice "Mute op failed on " & GetDeviceNameDirect(mDefRenderMM)
            End If
        Else
            'Form1.AppendNotice "Couldn't get default endpoint device."
        End If

    Else
        'Form1.AppendNotice "Couldn't get default endpoint device."
    End If
    
    On Error GoTo 0
    Exit Sub

unMute_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure unMute of Module modAudioCode"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : muteIt
' Author    : beededea
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub muteIt()

    On Error GoTo muteIt_Error

    Set mDeviceEnum = New MMDeviceEnumerator
    
    mDeviceEnum.GetDefaultAudioEndpoint eRender, eMultimedia, mDefRenderMM
    
    If (mDefRenderMM Is Nothing) = False Then
        mDefRenderMM.Activate IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, CVar(0), pEPVolMM
        If (pEPVolMM Is Nothing) = False Then
            If pEPVolMM.SetMute(1, UUID_NULL) = S_OK Then
                'Form1.AppendNotice "Muted " & GetDeviceNameDirect(mDefRenderMM)
                'Form1.btnMute.Caption = "UnMute"
            Else
                'Form1.AppendNotice "Mute op failed on " & GetDeviceNameDirect(mDefRenderMM)
            End If
        Else
            'Form1.AppendNotice "Couldn't get endpoint volume control."
        End If
    Else
       ' Form1.AppendNotice "Couldn't get default endpoint device."
    End If
    
    On Error GoTo 0
    Exit Sub

muteIt_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure muteIt of Module modAudioCode"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : StartCallback
' Author    : Fafalone
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
'public Sub StartCallback()
'
'    On Error GoTo StartCallback_Error
'
'    If (mDeviceEnum Is Nothing) Then
'        Set mDeviceEnum = New MMDeviceEnumerator
'
'        mDeviceEnum.GetDefaultAudioEndpoint eRender, eMultimedia, mDefRenderMM
'
'        If (mDefRenderMM Is Nothing) = False Then
'
'            mDefRenderMM.Activate IID_IAudioEndpointVolume, CLSCTX_INPROC_SERVER, CVar(0), pEPVolMM
'            If (pEPVolMM Is Nothing) = False Then
'                If (cVolCallback Is Nothing) Then
'                    Set cVolCallback = New cAudioEndpointVolumeCallback
'                    pEPVolMM.RegisterControlChangeNotify cVolCallback
'                    AppendNotice "Callback registered. Adjusting in Explorer or elsewhere will notify this app too."
'                End If
'            End If
'
'        End If
'    End If
'
'    Command1.Enabled = False
'    Command2.Enabled = True
'
'    On Error GoTo 0
'    Exit Sub
'
'StartCallback_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StartCallback of Form Form1"
'
'End Sub


'---------------------------------------------------------------------------------------
' Procedure : StopCallback
' Author    : Fafalone
' Date      : 19/12/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
'public Sub StopCallback()
'    On Error GoTo StopCallback_Error
'
'    If (pEPVolMM Is Nothing) = False Then
'        If (cVolCallback Is Nothing) = False Then
'            pEPVolMM.UnregisterControlChangeNotify cVolCallback
'            AppendNotice "Callback unregistered."
'            Command1.Enabled = True
'            Command2.Enabled = False
'            Set cVolCallback = Nothing
'            Set pEPVolMM = Nothing
'            Set mDefRenderMM = Nothing
'            Set mDeviceEnum = Nothing
'        End If
'    End If
'
'    Command1.Enabled = True
'    Command2.Enabled = False
'
'    On Error GoTo 0
'    Exit Sub
'
'StopCallback_Error:
'
'     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure StopCallback of Form Form1"
'End Sub
