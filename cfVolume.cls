VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
Option Explicit 'simple Alpha-Form (...to be filled with (mostly Img-) Content via .Widgets.Add)

Public WithEvents volumeForm As cWidgetForm
Attribute volumeForm.VB_VarHelpID = -1

Public collPSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

Private PSDWidth As Long
Private PSDHeight As Long

Private mMute As Boolean

Private RenderDev As cMMDevice
Private WithEvents SystemVolume As cSimpleAudioVolume 'use cSimpleAudioVoume with events
Attribute SystemVolume.VB_VarHelpID = -1
Private slidersetMouseDown As Boolean
Private lblPercent As cwLabel
Private lblPercentShadow As cwLabel



'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
   On Error GoTo Class_Initialize_Error

    Set collPSDNonUIElements = New_c.Collection(False)
    
    'retrieve (and Set) the Default-Renderdevice
    Set RenderDev = New_c.MMDeviceEnumerator.GetDefaultAudioEndpoint(eRender, eMultimedia)
    'retrieve (and Set) a cSimpleAudioVolume instance, to get access to Volume-Handling
    Set SystemVolume = RenderDev.Activate(AudioEndpointVolume)
    

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfVolume"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitFromPSD
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :    loads up the Cairo imageList with keyed entries and a surface from each and every PSD layer
'                adds path X,Y and alpha to the excluded collection
'                creates widgets with corresponding keynames, locations and sizes as per PSD for each layer
'                cwOverlay class puts it on the screen
'---------------------------------------------------------------------------------------
'
Public Sub InitFromPSD(ByRef PSD_FileNameOrByteArray As Variant)
    
    Const AlphaWithTaskbarEntry As Integer = 6
    Const AlphaNoTaskbarEntry As Integer = 7
  
    Dim i As Long: i = 0
    Dim W As cWidgetBase
   
    On Error GoTo InitFromPSD_Error

    'create the Alpha-capable Top-Level-Form
    Set volumeForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), widgetName, True, 1, 1)
        volumeForm.WidgetRoot.BackColor = -1
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
        PSDWidth = .Width
        PSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    
        For i = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
            If .LayerByteSize(i) Then  ' this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
                If .LayerAlphaPercent(i) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
                    'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                    Cairo.ImageList.AddSurface .LayerPath(i), .LayerSurface(i)
                    
                    If collPSDNonUIElements.Exists(.LayerPath(i)) Then ' we do not create Widgets for those in this exclude list.
                        
                        'we add layer info. (used in cwOverlay) to the "excluded paths" previously defined in Sub Main,
                        collPSDNonUIElements(.LayerPath(i)) = Array(.LayerX(i), .LayerY(i), Val(gblOpacity) / 100)  'here we update the so far empty slots with the PSD-offsets
                
                    Else 'everything not in the exclude-list will get its own Widget-instance
                        Set W = volumeForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(i)), .LayerX(i), .LayerY(i), .LayerWidth(i), .LayerHeight(i)).Widget
                        'Set W = volumeForm.Widgets.Add(New_W("cwAlphaImg"), LCase$(.LayerPath(I)), .LayerX(I), .LayerY(I), .LayerWidth(I), .LayerHeight(I)).Widget
                            W.ImageKey = W.Key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget
                        If gblOpacity < 1 Then
                            W.Alpha = Val(gblOpacity) / 100
                        Else
                            W.Alpha = 1
                        End If
                        
                        ' set the default characteristics of the non-interactive areas
                        ' note: the clickable layers characteristics are set in adjustMainControls
                        
                        ' all non-clickable Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                        W.HoverColor = -1 ' task: might change this later when new ver or RC6 arrives
                        If gblEnableTooltips = "1" Then W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                        W.MousePointer = IDC_SIZEALL

                    End If
                End If
            End If
        Next i
    End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
    
    Set lblPercentShadow = volumeForm.Widgets.Add(New_W("cwLabel"), "lblPercentShadow", 252, 92, 100, 50)
    Set lblPercent = volumeForm.Widgets.Add(New_W("cwLabel"), "lblPercent", 250, 90, 100, 50)
    
    lblPercent.Widget.FontSize = 38
    lblPercent.Caption = "33%"
    lblPercent.Widget.FontName = "times new roman"
    lblPercent.borderWidth = 0
    lblPercent.Widget.FontBold = True
    lblPercent.Widget.ForeColor = vbWhite
    lblPercent.Widget.BackColor = -1
    lblPercent.Widget.ToolTip = "This is the current volume in %."
    'lblPercent.TextShadowColor = vbBlack '&H80000010
    
    lblPercentShadow.Widget.FontSize = 38
    lblPercentShadow.Caption = "33%"
    lblPercentShadow.Widget.FontName = "times new roman"
    lblPercentShadow.borderWidth = 0
    lblPercentShadow.Widget.FontBold = True
    lblPercentShadow.Widget.ForeColor = vbBlack ' &H80000010
    lblPercentShadow.Widget.BackColor = -1
    lblPercentShadow.Widget.ToolTip = "This is the current volume in %."
    
    Set overlayWidget = volumeForm.Widgets.Add(New cwOverlay, "Overlay", 0, 0, PSDWidth, PSDHeight)
    Set overlayWidget.cCollOverlayImages = collPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
        overlayWidget.cxPerc = 0.4940476 'define the relative center (within cwOverlay, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
        overlayWidget.cyPerc = 0.4727272

   On Error GoTo 0
   Exit Sub

InitFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitFromPSD of Class Module cfVolume"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : volumeForm_BubblingEvent
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_BubblingEvent(ByRef Sender As Object, ByRef EventName As String, ByRef P1 As Variant, ByRef P2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
   
    Dim stdMessage As String: stdMessage = vbNullString
    Dim stdSubject As String: stdSubject = vbNullString
    
    On Error GoTo volumeForm_BubblingEvent_Error

    If gblIgnoreMouse = "1" Then Exit Sub
    
    stdMessage = "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize. "
    stdSubject = "You can turn off the balloon tooltips in the preferences."
    
'    If EventName = "W_MouseDrag" Then
'        Select Case LCase$(Sender.Widget.Key)
'            Case "sliderset"
'        End Select
'    End If
    
    If EventName = "W_MouseMove" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "lockingpin"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip volumeForm.hwnd, "Click me to lock the widget in place. This lock can be released by clicking again on the locking pin, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
            Case "bell"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip volumeForm.hwnd, "This is the mute button. Click me to toggle sound on/off.", _
                     TTIconInfo, widgetName & " Sound Toggle Help", , , , True
            Case "lockingpinunlocked"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip volumeForm.hwnd, "Click me to lock the widget in place. This lock can be released by clicking again on the locking pin, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
            Case "indicatorred"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip volumeForm.hwnd, "This is the mute indicator. A green lamp means that the sound is unmuted. Red lamp means that mute is ON.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
            Case "indicatorgreen"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip volumeForm.hwnd, "This is the mute indicator. Sound is unmuted.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
            Case "sliderset"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip volumeForm.hwnd, "This is the Volume Slider. Click and drag me left/right to change the volume.", _
                     TTIconInfo, widgetName & " Locking pin Help", , , , True
                If slidersetMouseDown = True Then


                End If
            Case "speaker", "cable", "bar", "pipes", "cablewheelset"
                If gblEnableBalloonTooltips = "1" Then CreateToolTip volumeForm.hwnd, stdMessage _
                    & vbCrLf & stdSubject, _
                    TTIconInfo, widgetName & " Configuration Help", , , , True
        End Select
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "lockingpin"
                Call lockingpin_mouseDown(Sender)
            Case "helppin"
                Call helppin_mouseDown(Sender)
            Case "lockingpinunlocked"
                Call lockingpin_mouseDown(Sender)
            Case "bell"
                Call bell_mouseDown(Sender)
            Case "sliderset"
                overlayWidget.Locked = True
                slidersetMouseDown = True
            Case Else
                Debug.Print Sender.Widget.Key
        End Select
    End If
    
    If EventName = "W_MouseUp" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "bell"
                Call bell_mouseUp(Sender)
            Case "sliderset"
                overlayWidget.Locked = False
                slidersetMouseDown = False
            Case Else
                Debug.Print Sender.Widget.Key
        End Select
    End If
    
'    If EventName = "W_Click" Then
'        Select Case LCase$(Sender.Widget.Key)
''            Case "helpbutton"
''                 Call helpbutton_click
''            Case "startbutton"
''                 Call startbutton_click
'            Case Else
'                Debug.Print Sender.Widget.Key
'        End Select
'    End If
   
   On Error GoTo 0
   Exit Sub

volumeForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_BubblingEvent of Class Module cfVolume"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : helpbutton_click
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
'Private Sub helpbutton_click()
'   On Error GoTo helpbutton_click_Error
'
'   Call helpSplash
'
'   On Error GoTo 0
'   Exit Sub
'
'helpbutton_click_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helpbutton_click of Class Module cfVolume"
'
'End Sub



'---------------------------------------------------------------------------------------
' Procedure : lockingpin_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lockingpin_mouseDown(ByRef Sender As Object)
   On Error GoTo lockingpin_mouseDown_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString

    fileToPlay = "lock.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If overlayWidget.Locked = False Then
        Sender.Widget.Alpha = 0
        fVolume.volumeForm.Widgets("lockingpinunlocked").Widget.Alpha = Val(gblOpacity) / 100
    Else
        Sender.Widget.Alpha = Val(gblOpacity) / 100
        fVolume.volumeForm.Widgets("lockingpinunlocked").Widget.Alpha = 0
    End If
    Sender.Widget.Refresh ' this removes the delay in the lockingpin depressing
    Call lockWidget
    
   On Error GoTo 0
   Exit Sub

lockingpin_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockingpin_mouseDown of Class Module cfVolume"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : helppin_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helppin_mouseDown(ByRef Sender As Object)
   On Error GoTo helppin_mouseDown_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString

    fileToPlay = "lock.wav"
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
'    If overlayWidget.Locked = False Then
'        Sender.Widget.Alpha = 0
'        fVolume.volumeForm.Widgets("helppinunlocked").Widget.Alpha = Val(gblOpacity) / 100
'    Else
'        Sender.Widget.Alpha = Val(gblOpacity) / 100
'        fVolume.volumeForm.Widgets("helppinunlocked").Widget.Alpha = 0
'    End If
    Sender.Widget.Refresh ' this removes the delay in the helppin depressing
    
    Call helpSplash
    
   On Error GoTo 0
   Exit Sub

helppin_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helppin_mouseDown of Class Module cfVolume"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : bell_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub bell_mouseDown(ByRef Sender As Object)
   On Error GoTo bell_mouseDown_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString

    fileToPlay = "buzzer.wav"

    If Mute = True Then
        Mute = False
    Else
        Mute = True
    End If
    
    If gblEnableSounds = "1" And fFExists(App.path & "\resources\sounds\" & fileToPlay) Then
        PlaySound App.path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    Sender.Widget.Alpha = 0
    Sender.Widget.Refresh ' this removes the delay in the bell depressing
    
   On Error GoTo 0
   Exit Sub

bell_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bell_mouseDown of Class Module cfVolume"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : bell_mouseUp
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub bell_mouseUp(ByRef Sender As Object)
   On Error GoTo bell_mouseUp_Error
    
    Sender.Widget.Alpha = Val(gblOpacity) / 100
    Sender.Widget.Refresh ' this removes the delay in the bell depressing
    
   On Error GoTo 0
   Exit Sub

bell_mouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bell_mouseUp of Class Module cfVolume"

End Sub


'---------------------------------------------------------------------------------------
' Property  : Mute
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Mute() As Boolean
   On Error GoTo MuteGet_Error

   Mute = mMute

   On Error GoTo 0
   Exit Property

MuteGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Mute of Class Module cwGlobe"
End Property
'---------------------------------------------------------------------------------------
' Property  : Mute
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Mute(ByVal newValue As Boolean)
   On Error GoTo MuteLet_Error

    If mMute <> newValue Then mMute = newValue Else Exit Property
    
    SystemVolume.SetMute Abs(CInt(mMute))
    
    If mMute = True Then
        fVolume.volumeForm.Widgets("indicatorgreen").Widget.Alpha = 0
        fVolume.volumeForm.Widgets("indicatorred").Widget.Alpha = Val(gblOpacity) / 100
    Else
        fVolume.volumeForm.Widgets("indicatorgreen").Widget.Alpha = Val(gblOpacity) / 100
        fVolume.volumeForm.Widgets("indicatorred").Widget.Alpha = 0
    End If
    fVolume.volumeForm.Refresh

   On Error GoTo 0
   Exit Property

MuteLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Mute of Class Module cwGlobe"
End Property


'---------------------------------------------------------------------------------------
' Procedure : volumeForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo volumeForm_DblClick_Error

    '    If gblIgnoreMouse = "1" Then Exit Sub

    If LTrim$(gblDblClickCommand) = vbNullString Then Exit Sub
    
    thisCommand = gblDblClickCommand
        
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If
    
    ' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If
    
    If SHIFT_1 = True Then
        SHIFT_1 = False
        Call ShellExecute(fVolume.volumeForm.hwnd, "Open", gblOpenFile, vbNullString, App.path, 1)
    Else
        Call ShellExecute(fVolume.volumeForm.hwnd, "runas", thisCommand, vbNullString, App.path, 1)
    End If

    On Error GoTo 0
    Exit Sub

volumeForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_DblClick of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : volumeForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo volumeForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift) ' will not catch Shift or CTRL keypresses on their own, must be combination with another key.

    On Error GoTo 0
    Exit Sub

volumeForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_KeyDown of Class Module cfMain"
End Sub
''
'---------------------------------------------------------------------------------------
' Procedure : volumeForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : save the form x,y position when ever the globe/form is dragged
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
    On Error GoTo volumeForm_MouseUp_Error
    
    Call savePosition

    On Error GoTo 0
    Exit Sub

volumeForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_MouseUp of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : volumeForm_MouseMove
' Author    : beededea
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = x: y0 = y 'just store the offsets when no button is down
  
   On Error GoTo volumeForm_MouseMove_Error
    
    If overlayWidget.Locked = True Then Exit Sub
    If gblIgnoreMouse = "1" Then Exit Sub

    If Button = vbLeftButton And Not volumeForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (x - x0) * volumeForm.WidgetRoot.Zoom + volumeForm.Left
      FY = (y - y0) * volumeForm.WidgetRoot.Zoom + volumeForm.Top
      volumeForm.Move FX, FY
    End If


   On Error GoTo 0
   Exit Sub

volumeForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_MouseMove of Class Module cfVolume"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : volumeForm_MouseWheel
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : rotate forward positive = smaller
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   Dim RotationDirection As Long: RotationDirection = 0
   
   On Error GoTo volumeForm_MouseWheel_Error

    If gblScrollWheelDirection = "1" Then
        If Rotation > 0 Then
            RotationDirection = 0 - Abs(Rotation)
        Else
            RotationDirection = Abs(Rotation)
        End If
    Else
        RotationDirection = Rotation
    End If
    
   If MouseKeys = 8 Then Call AdjustZoom(FZ + RotationDirection / 2400)   'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

volumeForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_MouseWheel of Class Module cfVolume"
End Sub
''
''---------------------------------------------------------------------------------------
'' Procedure : volumeForm_ResizeWithDimensionsDIP
'' Author    : beededea
'' Date      : 31/07/2023
'' Purpose   : put any other resizing action in here - unused
''---------------------------------------------------------------------------------------
''
'Private Sub volumeForm_ResizeWithDimensionsDIP(ByVal dx As Single, ByVal dy As Single)
'   On Error GoTo volumeForm_ResizeWithDimensionsDIP_Error
'
'  'If Not CloseWidget Is Nothing Then CloseWidget.Move dx - 32, 0, 32, 32 'place the close-widget at the top-right of the form
'
'   On Error GoTo 0
'   Exit Sub
'
'volumeForm_ResizeWithDimensionsDIP_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_ResizeWithDimensionsDIP of Class Module cfVolume"
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : AdjustZoom
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub AdjustZoom(ByVal NewZoom As Single)
   On Error GoTo AdjustZoom_Error

  FZ = NewZoom
  If FZ < 0.05 Then FZ = 0.05 Else If FZ > 3 Then FZ = 3 'keep FZ within a sane ZoomInterval
  volumeForm.WidgetRoot.Zoom = FZ '* volumeForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
  volumeForm.Move volumeForm.Left, volumeForm.Top, PSDWidth * volumeForm.WidgetRoot.Zoom, PSDHeight * volumeForm.WidgetRoot.Zoom

  ' when resized from the gauge scroll up/down it needs to write it back so the size is correct when prefs opened.
  gblGaugeSize = LTrim$(Str$(NewZoom * 100))
  If widgetPrefs.IsVisible = True Then widgetPrefs.sliGaugeSize = Val(gblGaugeSize)

    
   On Error GoTo 0
   Exit Sub

AdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AdjustZoom of Class Module cfVolume"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : W_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo W_MouseDown_Error

    If Button = vbRightButton Then
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        If gblIgnoreMouse = "1" Then Exit Sub
        ' do whatever you want with a mouseDown here, not doing anything at the moment but we will...
    End If

   On Error GoTo 0
   Exit Sub

W_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_MouseDown of Class Module cwOverlay"
End Sub


