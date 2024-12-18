Attribute VB_Name = "modScaleForm"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If you can read "Attribute VB_Name = "modScaleForm"" on the line above when already in the
' Access VBA editor, please remove this line or the module will not compile.
' When importing the .bas-file the the "import from file" function, this line never appears
' and the module is already correctly named "modScaleForm"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GENERAL INFORMATION
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Module     : modScaleForm
' Version    : 2024-12-18
' Usable     : From Access 2002 (XP) and up, maybe from Access 2000 and up
' Author     : Markus Gruber (markus.gruber@gruber.cc)
' Purpose    : Resize Access forms and scale the controls within according to screen resolution
'              and DPI.
' Usage      : 1) Add "ScaleFormWindow Me" to the OnOpen-Event of all your forms. When setting
'                 the optional parameter "DoNotScaleFormProportionally" to true, the window
'                 does not scale to the same ratio horizontally and vertically, but strict to
'                 the ratios between design resolution and current resolution. This may lead to
'                 unproportionally scaled forms and controls.
'              2) Add "ScaleFormControls Me" to the OnResize-Event of all your forms.
'              1+2) (You can also use the function "InitialInsertCodeIntoForms" to do this initially.
'                   Watch the "immediate window".)
'              3) Set the following constants in the module to match your design environment:
'                 DefaultDesignWidth (default is 1280)
'                 DefaultDesignHeight (default is 1024)
'                 DefaultDesignDPI (default is 120)
'              4) If one form was designed at another resolution, add the following to the
'                 form's tag field (do not forget all the ":"):
'                 DesignRes:<width>x<height>x<DPI>: - for example:
'                 DesignRes:1024x768x96:
' Credits    : .) Thanks to Jamie Czernik for his module modResizeForm, it was much inspiration!
'              .) Thanks to Myke Mayers for his function AdjustColumnWidths, it helped me a lot!
' Requires   : :) This module needs the class module clFormWindow:
'                 Visit http://www.mvps.org/access/forms/frm0042.htm
'                 and download http://www.mvps.org/access/downloads/clFormWindow.bas
' Remarks    : .) I am pretty sure that there are many errors in this module. Feel free
'                 to correct them yourself or simply inform me about them.
'              .) This module is provided "as is", I do not take responsability for usage.
'              .) This module is licensed under the GPL.
'              .) Thanks to all other Access developers posting solutions and code helping me
'                 over the last years!
' Remarks 2  : .) Resizing (sub)forms with many controls takes long (I know of a case, where it
'                 takes 15 seconds).
'              .) Fonts are sized using the smaller value of either horizontal or vertical scale
'                 factor - if this would not be the case, fonts would not fit any more into
'                 their controls.
'              .) Windows are centered after rescaling, so that they can not "disappear" to
'                 non-visible areas. Therefore, I use the class module "clFormWindow". If you only
'                 want the module to set the new size but not center the form, set the optional
'                 parameter "DoNotCenterForm" to "true".
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CHANGELOG
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2024-12-18: .) modScaleForm and clFormWindow are now compatible with 64 bit systems.
'
' 2008-03-10: .) New function PreScale: Can scale forms one time only in design mode to a desired
'                target resolution, could also be called "static scaling"
'             .) New function mSFOpenForm: Similar to docmd.openform with following addition: Opens
'                a given form not directly, but searches for a statically scaled form and opens this
'                one. Returns the name of the finally opened form as string. Same parameters usage as
'                docmd.openform, but no error handling implemented yet.
'             .) Cosmetic changes to the code (no added/changed functionality)
'
' 2008-02-08: .) Corrected version info in the "General Information" header. Thanks to anonymous!
'
' 2008-02-02: .) Added a comment about a possible caveat with the line
'                "Attribute VB_Name = "modScaleForm"". Thanks to Tony D'Ambra from
'                www.accessextra.net for the information!
'             .) Added description about optional parameter "DoNotScaleFormProportionally".
'             .) Added description about font sizing behavior.
'             .) Added description why I use clFormWindow for setting a form's position.
'             .) New optional parameter "DoNotCenterForm".
'
' 2008-01-21: .) Fixed a bug at scaling column widths
'
' 2008-01-18: .) First control element returned from Access has been been included in array (missing
'                "ReDim Preserve" added)
'
' 2008-01-07: .) DPI changes are reflected in scale factors
'             .) Added function "InitialInsertCodeIntoForms"
'
' 2007-12-01: .) Split ModScaleForm in several functions
'
' 2007-11-01: .) First thoughts and research
'             .) Decision to create modScaleForm
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Option Compare Database
Option Explicit

Private Type tRect 'for window sizes
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type tControl 'for control properties
    FormName As String
    NAME As String
    Height As Long
    Width As Long
    Top As Long
    Left As Long
    FontSize As Long
    ColumnWidths As String
    ListWidth As Long
    TabFixedWidth As Long
    TabFixedHeight As Long
End Type

Private Type tDisplay
    Height As Long
    Width As Long
    DPI As Long
End Type


Private Const DefaultDesignWidth As Long = 1920
Private Const DefaultDesignHeight As Long = 1080
Private Const DefaultDesignDPI As Long = 120
Private Const WM_HORZRES As Long = 8
Private Const WM_VERTRES As Long = 10
Private Const WM_LOGPIXELSX As Long = 88

#If VBA7 Then
    Private Declare PtrSafe Function WM_apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function WM_apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function WM_apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function WM_apiGetWindowRect Lib "user32.dll" Alias "GetWindowRect" (ByVal hWnd As LongPtr, lpRect As tRect) As Long
#Else
    Private Declare Function WM_apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
    Private Declare Function WM_apiReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hWnd As Long, ByVal hdc As Long) As Long
    Private Declare Function WM_apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function WM_apiGetWindowRect Lib "user32.dll" Alias "GetWindowRect" (ByVal hWnd As Long, lpRect As tRect) As Long
#End If



Public arrCtlsScaleForm() As tControl 'public array for control properties, available for all forms

Public Sub ScaleFormWindow( _
    ByVal frm As Access.Form, _
    Optional DoNotScaleFormProportionally As Boolean, _
    Optional DoNotCenterForm As Boolean)

Dim CurrentFormHeight As Long
Dim CurrentFormName As String
Dim CurrentFormWidth As Long
Dim CurrentScreenDPI As Long
Dim CurrentScreenHeight As Long
Dim CurrentScreenWidth As Long
Dim DesignResDPI As Long
Dim DesignResHeight As Long
Dim DesignResWidth As Long
Dim GetScreenResolution As tDisplay
#If VBA7 Then
    Dim hDCcaps As LongPtr
#Else
    Dim hDCcaps As Long
#End If
Dim lngRtn As Long
Dim MonitorResHeightRatio As Single
Dim MonitorResWidthRatio As Single
Dim NewForm As New clFormWindow
Dim NewFormHeight As Long
Dim NewFormWidth As Long
Dim OnOpenWindowHeight As Long
Dim OnOpenWindowWidth As Long
Dim rectWindow As tRect
Dim ScaledFormString As String
Dim TagString As String
Dim TagStringArray() As String
Dim TagStringPosition As Long


On Error Resume Next

CurrentFormName = frm.NAME

'Call API to get current resolution
hDCcaps = WM_apiGetDC(0) 'Get display context for desktop (hwnd = 0).
With GetScreenResolution
    .Height = WM_apiGetDeviceCaps(hDCcaps, WM_VERTRES)
    .Width = WM_apiGetDeviceCaps(hDCcaps, WM_HORZRES)
    .DPI = WM_apiGetDeviceCaps(hDCcaps, WM_LOGPIXELSX)
End With
lngRtn = WM_apiReleaseDC(0, hDCcaps) 'Release display context.

CurrentScreenWidth = GetScreenResolution.Width
CurrentScreenHeight = GetScreenResolution.Height
CurrentScreenDPI = GetScreenResolution.DPI

ScaledFormString = ("_mSF" & CurrentScreenWidth & "x" & CurrentScreenHeight & "x" & CurrentScreenDPI)

If isSubform(frm) Then
    Exit Sub
Else
    If Right(CurrentFormName, Len(ScaledFormString)) = ScaledFormString Then
        Exit Sub 'this form is statically scaled for the current resolution
    Else
    'go go go
    End If
End If


'Set size to fit form
DoCmd.RunCommand acCmdSizeToFitForm


'Extract "DesignRes:*:" from Tag-Option of Form
TagString = LCase(frm.Tag) 'lower case tag-field
TagStringPosition = InStr(TagString, "designres:") 'get left position of resolution
If TagStringPosition > 0 Then
    TagString = Mid(TagString, TagStringPosition + 10) 'get string starting with resolution (the "+10" is the char count of "DesignRes:"
    TagStringPosition = InStr(TagString, ":") 'Get Position of ":" behind resolution
    TagString = Left(TagString, TagStringPosition - 1) 'remove everything behind ":" - the variable now contains the design resolution
    ReDim TagStringArray(1)
    TagStringArray = Split(TagString, "x") '"x" is the delimiter
    If UBound(TagStringArray) >= 0 Then DesignResWidth = TagStringArray(0)
    If UBound(TagStringArray) >= 1 Then DesignResHeight = TagStringArray(1)
    If UBound(TagStringArray) >= 2 Then DesignResDPI = TagStringArray(2)
    If (DesignResWidth <= 0) Or (DesignResHeight <= 0) Or (DesignResDPI <= 0) Then
        DesignResWidth = DefaultDesignWidth
        DesignResHeight = DefaultDesignHeight
        DesignResDPI = DefaultDesignDPI
    End If
    Erase TagStringArray
Else
    'DesignRes is not specified in the form so we use default values for scaling
    DesignResWidth = DefaultDesignWidth
    DesignResHeight = DefaultDesignHeight
    DesignResDPI = DefaultDesignDPI
End If


'Extract "OnOpenRes:*:" from Tag-Option of Form
TagString = LCase(frm.Tag) 'lower case tag-field
TagStringPosition = InStr(TagString, "onopenres:") 'get left position of resolution
If TagStringPosition > 0 Then
    TagString = Mid(TagString, TagStringPosition + 10) 'get string starting with resolution (the "+10" is the char count of "OnOpenRes:"
    TagStringPosition = InStr(TagString, ":") 'Get Position of ":" behind resolution
    TagString = Left(TagString, TagStringPosition - 1) 'remove everything behind ":" - the variable now contains the OnOpen resolution
    ReDim TagStringArray(1)
    TagStringArray = Split(TagString, "x") '"x" is the delimiter
    OnOpenWindowWidth = TagStringArray(0)
    OnOpenWindowHeight = TagStringArray(1)
    Erase TagStringArray
Else
    Call WM_apiGetWindowRect(frm.hWnd, rectWindow) 'prepare for getting size of current window on screen
    OnOpenWindowWidth = rectWindow.Right - rectWindow.Left
    OnOpenWindowHeight = rectWindow.Bottom - rectWindow.Top
    frm.Tag = frm.Tag & ", OnOpenRes:" & OnOpenWindowWidth & "x" & OnOpenWindowHeight & ":"
End If



If (DesignResDPI = CurrentScreenDPI) And (DesignResWidth = CurrentScreenWidth) And (DesignResHeight = CurrentScreenHeight) Then
    'design values match current values, no need for scaling
    Exit Sub
Else
    'design values and current values are not equal - go on
End If


'Calculate Ratios
MonitorResHeightRatio = (CurrentScreenHeight / DesignResHeight) * (DesignResDPI / CurrentScreenDPI)
MonitorResWidthRatio = (CurrentScreenWidth / DesignResWidth) * (DesignResDPI / CurrentScreenDPI)


'If DoNotScaleFormProportionally is True, then the ratios remain the same. Else, the smaller value is chosen for both
If DoNotScaleFormProportionally = False Then
    If MonitorResHeightRatio < MonitorResWidthRatio Then MonitorResWidthRatio = MonitorResHeightRatio
    If MonitorResWidthRatio < MonitorResHeightRatio Then MonitorResHeightRatio = MonitorResWidthRatio
Else
    'leave ratios as they are and scale unproportionally, strictly as display resolution ratios are.
End If


'Get width and heights of form
Call WM_apiGetWindowRect(frm.hWnd, rectWindow) 'prepare for getting size of current window on screen
CurrentFormWidth = rectWindow.Right - rectWindow.Left
CurrentFormHeight = rectWindow.Bottom - rectWindow.Top


'Calculate new width and height (add 1 pixel for safety reasons), including percentage relative to whole desktop
NewFormHeight = Round(CurrentFormHeight * MonitorResHeightRatio, 0) '+ 1
NewFormWidth = Round(CurrentFormWidth * MonitorResWidthRatio, 0) '+ 1


'Set form's new size and position
NewForm.hWnd = frm.hWnd
With NewForm
    .Height = NewFormHeight
    .Width = NewFormWidth
End With
'If DoNotCenterForm is True, then the form is not centered
If DoNotCenterForm = False Then
    With NewForm
        .Top = (.Parent.Height - NewFormHeight) / 2
        .Left = (.Parent.Width - NewFormWidth) / 2
    End With
Else
    'form should not be centered
End If
Set NewForm = Nothing


End Sub

Public Sub ScaleFormControls( _
    ByVal frm As Access.Form, _
    Optional NoScalingWhenSubform As Boolean)

Dim ArrayCount As Long
Dim ctl As Control
Dim CurrentWindowHeight As Long
Dim CurrentWindowWidth As Long
Dim FormExists As Boolean
Dim LastScaleHorzFactor As Single
Dim LastScaleVertFactor As Single
Dim OnOpenWindowHeight As Long
Dim OnOpenWindowWidth As Long
Dim rectWindow As tRect
Dim ScaleFontFactor As Single
Dim ScaleHorzFactor As Single
Dim ScaleVertFactor As Single
Dim TagString As String
Dim TagStringArray() As String
Dim TagStringPosition As Long


On Error Resume Next 'no error messages, no debugger will start

'prepare array holding initial control sizes
If (Not arrCtlsScaleForm) = -1 Then
    'array empty, so create one
    ReDim arrCtlsScaleForm(1) 'create array
Else
    'array exists, do nothing
End If


If isSubform(frm) And (NoScalingWhenSubform = True) Then
    Exit Sub
Else
    'go go go
End If


'Extract "OnOpenRes:*:" from Tag-Option of Form
TagString = LCase(frm.Tag) 'lower case tag-field
TagStringPosition = InStr(TagString, "onopenres:") 'get left position of resolution
If TagStringPosition > 0 Then
    TagString = Mid(TagString, TagStringPosition + 10) 'get string starting with resolution (the "+10" is the char count of "OnOpenRes:"
    TagStringPosition = InStr(TagString, ":") 'Get Position of ":" behind resolution
    TagString = Left(TagString, TagStringPosition - 1) 'remove everything behind ":" - the variable now contains the OnOpen resolution
    ReDim TagStringArray(1)
    TagStringArray = Split(TagString, "x") '"x" is the delimiter
    OnOpenWindowWidth = TagStringArray(0)
    OnOpenWindowHeight = TagStringArray(1)
    Erase TagStringArray
Else
    Call WM_apiGetWindowRect(frm.hWnd, rectWindow) 'prepare for getting size of current window on screen
    OnOpenWindowWidth = rectWindow.Right - rectWindow.Left
    OnOpenWindowHeight = rectWindow.Bottom - rectWindow.Top
    frm.Tag = frm.Tag & ", OnOpenRes:" & OnOpenWindowWidth & "x" & OnOpenWindowHeight & ":"
End If


'check if form has already been opened and is part of the array
'if not, write the properties to the array
FormExists = 0 'means name of form is not in the array, initially
For ArrayCount = LBound(arrCtlsScaleForm) To UBound(arrCtlsScaleForm)
    If arrCtlsScaleForm(ArrayCount).FormName = frm.NAME Then FormExists = 1 'form name is already in the array, no need to write array
Next ArrayCount


If FormExists = 0 Then
    ArrayCount = UBound(arrCtlsScaleForm) + 1
    ReDim Preserve arrCtlsScaleForm(ArrayCount) 'Increase the size of the array.
    For Each ctl In frm.Controls
        With arrCtlsScaleForm(ArrayCount)
            .FormName = frm.NAME
            .NAME = ctl.NAME
            .Height = ctl.Height
            .Width = ctl.Width
            .Top = ctl.Top
            .Left = ctl.Left
            .FontSize = ctl.FontSize
            .ColumnWidths = ctl.ColumnWidths
            .ListWidth = ctl.ListWidth
            .TabFixedWidth = ctl.TabFixedWidth
            .TabFixedHeight = ctl.TabFixedHeight
        End With
        ArrayCount = ArrayCount + 1
        ReDim Preserve arrCtlsScaleForm(ArrayCount) 'Increase the size of the array.
    Next ctl
    'save initial heights for header, detail, footer
    ArrayCount = UBound(arrCtlsScaleForm) + 1
    ReDim Preserve arrCtlsScaleForm(ArrayCount)
    With arrCtlsScaleForm(ArrayCount)
        .FormName = frm.NAME
        .NAME = "xxxHeaderxxx"
        .Height = frm.Section(Access.acHeader).Height
    End With
    ArrayCount = ArrayCount + 1
    ReDim Preserve arrCtlsScaleForm(ArrayCount)
    With arrCtlsScaleForm(ArrayCount)
        .FormName = frm.NAME
        .NAME = "xxxDetailxxx"
        .Height = frm.Section(Access.acDetail).Height
    End With
    ArrayCount = ArrayCount + 1
    ReDim Preserve arrCtlsScaleForm(ArrayCount)
    With arrCtlsScaleForm(ArrayCount)
        .FormName = frm.NAME
        .NAME = "xxxFooterxxx"
        .Height = frm.Section(Access.acFooter).Height
    End With
    ArrayCount = ArrayCount + 1
    ReDim Preserve arrCtlsScaleForm(ArrayCount)
    With arrCtlsScaleForm(ArrayCount)
        .FormName = frm.NAME
        .NAME = "xxxLastScaleVertFactorxxx"
        .Height = 1
    End With
    ArrayCount = ArrayCount + 1
    ReDim Preserve arrCtlsScaleForm(ArrayCount)
    With arrCtlsScaleForm(ArrayCount)
        .FormName = frm.NAME
        .NAME = "xxxLastScaleHorzFactorxxx"
        .Height = 1
    End With
End If


Call WM_apiGetWindowRect(frm.hWnd, rectWindow) 'prepare for getting size of current window on screen
CurrentWindowWidth = rectWindow.Right - rectWindow.Left
CurrentWindowHeight = rectWindow.Bottom - rectWindow.Top


For ArrayCount = LBound(arrCtlsScaleForm) To UBound(arrCtlsScaleForm)
    If arrCtlsScaleForm(ArrayCount).FormName = frm.NAME Then
        Select Case arrCtlsScaleForm(ArrayCount).NAME
            Case "xxxLastScaleHorzFactorxxx"
               LastScaleHorzFactor = arrCtlsScaleForm(ArrayCount).Height
            Case "xxxLastScaleVertFactorxxx"
                LastScaleVertFactor = arrCtlsScaleForm(ArrayCount).Height
        End Select
    End If
Next


'Calculate the scaling factors and round them to 2 decimal places.
'This ensures that changes happen only in the area of 1% relative to the original size
ScaleVertFactor = Round((CurrentWindowHeight / OnOpenWindowHeight), 2)
ScaleHorzFactor = Round((CurrentWindowWidth / OnOpenWindowWidth), 2)
ScaleFontFactor = VBA.IIf(ScaleHorzFactor < ScaleVertFactor, ScaleHorzFactor, ScaleVertFactor)

If (ScaleVertFactor = LastScaleVertFactor) And (ScaleHorzFactor = LastScaleHorzFactor) Then
    'nothing to rescale
    Exit Sub
Else
    'there is something to rescale, so we update the array
    For ArrayCount = LBound(arrCtlsScaleForm) To UBound(arrCtlsScaleForm)
        If arrCtlsScaleForm(ArrayCount).FormName = frm.NAME Then
            Select Case arrCtlsScaleForm(ArrayCount).NAME
                Case "xxxLastScaleHorzFactorxxx"
                    arrCtlsScaleForm(ArrayCount).Height = ScaleHorzFactor
                Case "xxxLastScaleVertFactorxxx"
                    arrCtlsScaleForm(ArrayCount).Height = ScaleVertFactor
            End Select
        End If
    Next
End If



frm.Painting = False 'Turn off painting


For ArrayCount = LBound(arrCtlsScaleForm) To UBound(arrCtlsScaleForm)
    If arrCtlsScaleForm(ArrayCount).FormName = frm.NAME Then
        Select Case arrCtlsScaleForm(ArrayCount).NAME
            Case "xxxHeaderxxx"
                frm.Section(Access.acHeader).Height = arrCtlsScaleForm(ArrayCount).Height * ScaleVertFactor
            Case "xxxDetailxxx"
                frm.Section(Access.acDetail).Height = arrCtlsScaleForm(ArrayCount).Height * ScaleVertFactor
            Case "xxxFooterxxx"
                frm.Section(Access.acFooter).Height = arrCtlsScaleForm(ArrayCount).Height * ScaleVertFactor
        End Select
    End If
Next ArrayCount


For ArrayCount = LBound(arrCtlsScaleForm) To UBound(arrCtlsScaleForm)
    If (arrCtlsScaleForm(ArrayCount).FormName = frm.NAME) And (arrCtlsScaleForm(ArrayCount).NAME <> "") Then
        'On Error GoTo err
        If frm.Controls.Item(arrCtlsScaleForm(ArrayCount).NAME).ControlType <> Access.acPage Then  'Ignore pages in Tab controls.
            With frm.Controls.Item(arrCtlsScaleForm(ArrayCount).NAME)
                If ScaleVertFactor <> LastScaleVertFactor Then
                    .Height = arrCtlsScaleForm(ArrayCount).Height * ScaleVertFactor
                    .Top = arrCtlsScaleForm(ArrayCount).Top * ScaleVertFactor
                End If
                If ScaleHorzFactor <> LastScaleHorzFactor Then
                    .Left = arrCtlsScaleForm(ArrayCount).Left * ScaleHorzFactor
                    .Width = arrCtlsScaleForm(ArrayCount).Width * ScaleHorzFactor
                End If
                If .FontSize > 0 Then
                    .FontSize = arrCtlsScaleForm(ArrayCount).FontSize * ScaleFontFactor
                End If
                Select Case frm.Controls.Item(arrCtlsScaleForm(ArrayCount).NAME).ControlType
                    Case Access.acListBox
                        .ColumnWidths = ScaleColumnWidths(arrCtlsScaleForm(ArrayCount).ColumnWidths, ScaleHorzFactor)
                    Case Access.acComboBox
                        .ColumnWidths = ScaleColumnWidths(arrCtlsScaleForm(ArrayCount).ColumnWidths, ScaleHorzFactor)
                        .ListWidth = arrCtlsScaleForm(ArrayCount).ListWidth * ScaleHorzFactor
                    Case Access.acTabCtl
                        .TabFixedWidth = arrCtlsScaleForm(ArrayCount).TabFixedWidth * ScaleHorzFactor
                        .TabFixedHeight = arrCtlsScaleForm(ArrayCount).TabFixedHeight * ScaleVertFactor
                End Select
            End With
        End If
    End If
Next


frm.Painting = True


'Keep this comments for debugging reasons
frm.txtOnOpenwindowWidth = OnOpenWindowWidth
frm.txtOnOpenwindowHeight = OnOpenWindowHeight
frm.txtCurrentwindowWidth = CurrentWindowWidth
frm.txtCurrentwindowHeight = CurrentWindowHeight
frm.txtscalevertfactor = ScaleVertFactor
frm.txtscalehorzfactor = ScaleHorzFactor


'Free up resources
Set ctl = Nothing 'Free up resources.
Set frm = Nothing 'Free up resources.


End Sub
Function isSubform(frmIn As Form) As Boolean

Dim strX As String

On Error Resume Next

strX = frmIn.Parent.NAME
isSubform = Err.Number = 0

End Function
Private Function ScaleColumnWidths(DesignColumnWidths As String, ScaleColumnWidthFactor As Single) As String

On Error Resume Next

Dim DesignColumnWidthsArray() As String
Dim NewColumnWidths As String
Dim TempVar As Long

ReDim DesignColumnWidthsArray(0)

DesignColumnWidthsArray = Split(DesignColumnWidths, ";")
NewColumnWidths = vbNullString

For TempVar = LBound(DesignColumnWidthsArray) To UBound(DesignColumnWidthsArray)
    If Not IsNull(DesignColumnWidthsArray(TempVar)) And DesignColumnWidthsArray(TempVar) <> "" Then
        NewColumnWidths = NewColumnWidths & CSng(DesignColumnWidthsArray(TempVar)) * ScaleColumnWidthFactor & ";"
    End If
Next

ScaleColumnWidths = NewColumnWidths
Erase DesignColumnWidthsArray

End Function

Function InitialInsertCodeIntoForms()

Dim lngProcCountLines As Long
Dim lngProcStartLine As Long
Dim loCont As Container
Dim loDb As Database
Dim loDoc As Document
Dim loForm As Form
Dim loMod As Module
Dim pkType As Long
Dim strCode As String
Dim strName As String


Set loDb = CurrentDb
Set loCont = loDb.Containers("Forms")


For Each loDoc In loCont.Documents
    strName = loDoc.NAME
    DoCmd.OpenForm strName, acDesign
    Set loForm = Forms(strName)
    Debug.Print loForm.NAME; ".HasModule = "; loForm.HasModule
    If loForm.HasModule = True Then
        Set loMod = loForm.Module
        On Error Resume Next
        lngProcStartLine = loMod.ProcStartLine("Form_Open", pkType)
        If Err > 0 Then
            Debug.Print vbTab; "FormOpen Added"
            loForm.OnOpen = "[Event Procedure]"
            lngProcStartLine = loMod.CountOfLines
            strCode = "Private Sub Form_Open(Cancel As Integer)" _
            & vbCrLf _
            & "ScaleFormWindow Me" _
            & vbCrLf _
            & "End Sub"
            loMod.InsertLines lngProcStartLine + 1, strCode
        Else
            Debug.Print vbTab; "FormOpen Edited"
            lngProcCountLines = loMod.ProcCountLines("Form_Open", pkType)
            loMod.InsertLines lngProcStartLine + lngProcCountLines - 1, "ScaleFormWindow Me"
        End If
        If loForm.OnOpen <> "[Event Procedure]" Then Debug.Print vbTab; "Check OnOpen Event, it is not set to [Event Procedure]."
        Err.Clear
    End If
    
    If loForm.HasModule = True Then
        Set loMod = loForm.Module
        On Error Resume Next
        lngProcStartLine = loMod.ProcStartLine("Form_Resize", pkType)
        If Err > 0 Then
            Debug.Print vbTab; "FormResize Added"
            loForm.OnOpen = "[Event Procedure]"
            lngProcStartLine = loMod.CountOfLines
            strCode = "Private Sub Form_Resize()" _
            & vbCrLf _
            & "ScaleFormControls Me" _
            & vbCrLf _
            & "End Sub"
            loMod.InsertLines lngProcStartLine + 1, strCode
        Else
            Debug.Print vbTab; "FormResize Edited"
            lngProcCountLines = loMod.ProcCountLines("Form_Resize", pkType)
            loMod.InsertLines lngProcStartLine + lngProcCountLines - 1, "ScaleFormControls Me"
        End If
        If loForm.OnResize <> "[Event Procedure]" Then Debug.Print vbTab; "Check OnResize Event, it is not set to [Event Procedure]."
        Err.Clear
    End If
    DoCmd.Close acForm, strName, acSaveYes
Next


Set loMod = Nothing
Set loForm = Nothing
Set loDoc = Nothing
Set loCont = Nothing
Set loDb = Nothing


End Function

Function PreScale()

Dim ArrayCount As Long
Dim ctl As Control
Dim DesignResDPI As Long
Dim DesignResHeight As Long
Dim DesignResWidth As Long
Dim FormHasToBeScaled As Boolean
Dim FormNameArray() As String
Dim frm As Object
Dim frm2 As Form
Dim KeepOriginalFormNames As Boolean
Dim NewFormName As String
Dim ScaleFontFactor As Single
Dim ScaleHorzFactor As Single
Dim ScaleVertFactor As Single
Dim TagString As String
Dim TagStringArray() As String
Dim TagStringPosition As Long
Dim TagStringPosition2 As Long
Dim TagStringPosition3 As Long
Dim TagStringTemp As String
Dim TargetResDPI As Long
Dim TargetResHeight As Long
Dim TargetResString As String
Dim TargetResWidth As Long
Dim TempString As String


'Define options here
TargetResWidth = 1400
TargetResHeight = 1050
TargetResDPI = 96
KeepOriginalFormNames = False


Debug.Print "==========================================="
Debug.Print "Starting"
Debug.Print "==========================================="


'Write the names of all forms to an array
ReDim FormNameArray(0)
ArrayCount = UBound(FormNameArray)

Debug.Print
Debug.Print "Writing form names to array"
Debug.Print "==========================================="
Debug.Print

For Each frm In CurrentProject.AllForms
    ReDim Preserve FormNameArray(ArrayCount)
    FormNameArray(ArrayCount) = frm.NAME
    Debug.Print frm.NAME
    ArrayCount = ArrayCount + 1
Next

TargetResString = "_mSF" & TargetResWidth & "x" & TargetResHeight & "x" & TargetResDPI

'delete all existing prescaled forms for target resolution
Debug.Print
Debug.Print "Deleting existing forms for " & TargetResWidth & "x" & TargetResHeight & "x" & TargetResDPI
Debug.Print "==========================================="
Debug.Print

For ArrayCount = LBound(FormNameArray) To UBound(FormNameArray)
    If Right(FormNameArray(ArrayCount), Len(TargetResString)) = TargetResString Then 'this form can be deleted
        On Error Resume Next
        DoCmd.DeleteObject acForm, FormNameArray(ArrayCount)
        Debug.Print FormNameArray(ArrayCount)
    End If
Next

Debug.Print
Debug.Print "Starting form manipulation"
Debug.Print "==========================================="
Debug.Print


For ArrayCount = LBound(FormNameArray) To UBound(FormNameArray)
    Debug.Print FormNameArray(ArrayCount)
    If InStr(1, FormNameArray(ArrayCount), "_msf", vbTextCompare) > 0 Then
        'form has "_msf" in the name, so it has already been scaled for another resolution
        Debug.Print "   skipped, it has already been scaled for another resolution"
    Else
        If Len(FormNameArray(ArrayCount) & TargetResString) <= 64 Then
            FormHasToBeScaled = True
            DoCmd.OpenForm FormNameArray(ArrayCount), acDesign
            Set frm2 = Forms(FormNameArray(ArrayCount))
            'Extract "DesignRes:*:" from Tag-Option of Form
            TagString = LCase(frm2.Tag) 'lower case tag-field
            TagStringPosition = InStr(TagString, "designres:") 'get left position of resolution
            If TagStringPosition > 0 Then
                TagString = Mid(TagString, TagStringPosition + 10) 'get string starting with resolution (the "+10" is the char count of "DesignRes:"
                TagStringPosition = InStr(TagString, ":") 'Get Position of ":" behind resolution
                TagString = Left(TagString, TagStringPosition - 1) 'remove everything behind ":" - the variable now contains the design resolution
                ReDim TagStringArray(1)
                TagStringArray = Split(TagString, "x") '"x" is the delimiter
                    If UBound(TagStringArray) >= 0 Then DesignResWidth = TagStringArray(0)
                    If UBound(TagStringArray) >= 1 Then DesignResHeight = TagStringArray(1)
                    If UBound(TagStringArray) >= 2 Then DesignResDPI = TagStringArray(2)
                    If (DesignResWidth <= 0) Or (DesignResHeight <= 0) Or (DesignResDPI <= 0) Then
                        DesignResWidth = 0
                        DesignResHeight = 0
                        DesignResDPI = 0
                    End If
                    Erase TagStringArray
            Else
                'DesignRes is not specified in the form so we use default values for scaling
                DesignResWidth = 0
                DesignResHeight = 0
                DesignResDPI = 0
            End If
            If (DesignResWidth > 0) And (DesignResHeight > 0) And (DesignResDPI > 0) Then
                If (DesignResWidth = TargetResWidth) And (DesignResHeight = TargetResHeight) And (DesignResDPI = TargetResDPI) Then
                Debug.Print "   Forms DesignRes is equal to the TargetRes, so it does not have to be scaled"
                FormHasToBeScaled = False
                End If
            Else
                If (DefaultDesignWidth = TargetResWidth) And (DefaultDesignHeight = TargetResHeight) And (DefaultDesignDPI = TargetResDPI) Then
                    Debug.Print "   Forms DesignRes is not set, TargetRes is equal to DefaultDesignRes, so it does not have to be scaled"
                    FormHasToBeScaled = False
                End If
            End If
            DoCmd.Close acForm, FormNameArray(ArrayCount), acSaveNo
            If FormHasToBeScaled = True Then
                On Error Resume Next
                If KeepOriginalFormNames = True Then
                    NewFormName = FormNameArray(ArrayCount)
                Else
                    NewFormName = FormNameArray(ArrayCount) & TargetResString
                    DoCmd.CopyObject , NewFormName, acForm, FormNameArray(ArrayCount)
                    Debug.Print "   copying to " & NewFormName
                    Debug.Print "   changing subform names"
                End If
                DoCmd.OpenForm NewFormName, acDesign
                If KeepOriginalFormNames = False Then
                    For Each ctl In Screen.ActiveForm.Controls
                        TempString = ctl.SourceObject & TargetResString
                        ctl.SourceObject = TempString
                    Next
                End If
                Debug.Print "   scaling controls"
                If (DesignResWidth > 0) And (DesignResHeight > 0) And (DesignResDPI > 0) Then
                    ScaleVertFactor = (TargetResHeight / DesignResHeight) * (TargetResDPI / DesignResDPI)
                    ScaleHorzFactor = (TargetResWidth / DesignResWidth) * (TargetResDPI / DesignResDPI)
                    ScaleFontFactor = VBA.IIf(ScaleHorzFactor < ScaleVertFactor, ScaleHorzFactor, ScaleVertFactor)
                Else
                    ScaleVertFactor = (TargetResHeight / DefaultDesignHeight) * (TargetResDPI / DefaultDesignDPI)
                    ScaleHorzFactor = (TargetResWidth / DefaultDesignWidth) * (TargetResDPI / DefaultDesignDPI)
                    ScaleFontFactor = VBA.IIf(ScaleHorzFactor < ScaleVertFactor, ScaleHorzFactor, ScaleVertFactor)
                End If
                Screen.ActiveForm.Width = Screen.ActiveForm.Width * ScaleHorzFactor
                Screen.ActiveForm.Section(Access.acHeader).Height = Screen.ActiveForm.Section(Access.acHeader).Height * ScaleVertFactor
                Screen.ActiveForm.Section(Access.acDetail).Height = Screen.ActiveForm.Section(Access.acDetail).Height * ScaleVertFactor
                Screen.ActiveForm.Section(Access.acFooter).Height = Screen.ActiveForm.Section(Access.acFooter).Height * ScaleVertFactor
                For Each ctl In Screen.ActiveForm.Controls
                    If ctl.ControlType <> Access.acPage Then  'Ignore pages in Tab controls.
                        With ctl
                            .Height = ctl.Height * ScaleVertFactor
                            .Top = ctl.Top * ScaleVertFactor
                            .Left = ctl.Left * ScaleHorzFactor
                            .Width = ctl.Width * ScaleHorzFactor
                            If .FontSize > 0 Then
                                .FontSize = ctl.FontSize * ScaleFontFactor
                            End If
                        Select Case ctl.ControlType
                            Case Access.acListBox
                                .ColumnWidths = ScaleColumnWidths(ctl.ColumnWidths, ScaleHorzFactor)
                            Case Access.acComboBox
                                .ColumnWidths = ScaleColumnWidths(ctl.ColumnWidths, ScaleHorzFactor)
                                .ListWidth = ctl.ListWidth * ScaleHorzFactor
                            Case Access.acTabCtl
                                .TabFixedWidth = ctl.TabFixedWidth * ScaleHorzFactor
                                .TabFixedHeight = ctl.TabFixedHeight * ScaleVertFactor
                        End Select
                    End With
                    End If
                Next
                frm.Painting = True
                'remove temporary OnOpenRes from Tag
                Set frm2 = Forms(NewFormName)
                frm2.Tag = Left(frm2.Tag, (InStr(LCase(frm2.Tag), LCase("onopenres:")) - 3))
                'Update DesignRes
                If InStr(LCase(frm2.Tag), LCase("designres:")) > 0 Then 'designres is set in the form, so remove it
                    TagStringPosition2 = InStr(LCase(frm2.Tag), LCase("designres:"))
                    TagStringPosition3 = InStr(TagStringPosition2 + Len("designres:"), frm2.Tag, ":")
                    TagStringPosition3 = Len(frm2.Tag) - TagStringPosition3
                    TagStringTemp = Left(frm2.Tag, TagStringPosition2 - 1) & Right(frm2.Tag, TagStringPosition3)
                    frm2.Tag = TagStringTemp
                End If
                frm2.Tag = frm2.Tag & ", DesignRes:" & TargetResWidth & "x" & TargetResHeight & "x" & TargetResDPI & ":"
                'Save scaled form
                DoCmd.Close acForm, NewFormName, acSaveYes
            End If
        Else
            Debug.Print NewFormName & " would have more than 64 characters - form was not created."
        End If
    End If
Next

Debug.Print
Debug.Print "==========================================="
Debug.Print "Finished!"
Debug.Print "==========================================="

End Function

Public Function mSFOpenForm( _
                FormName As String, _
                Optional View As AcFormView = acNormal, _
                Optional FilterName As String, _
                Optional WhereCondition As String, _
                Optional DataMode As AcFormOpenDataMode = acFormPropertySettings, _
                Optional WindowMode As AcWindowMode = acWindowNormal, _
                Optional OpenArgs As String) As String
                'this function returns the name of the opened form as String

Dim frm As Object
Dim GetScreenResolution As tDisplay
Dim hDCcaps As Long
Dim lngRtn As Long
Dim ScaledFormName As String


hDCcaps = WM_apiGetDC(0) 'Get display context for desktop (hwnd = 0).
With GetScreenResolution
    .Height = WM_apiGetDeviceCaps(hDCcaps, WM_VERTRES)
    .Width = WM_apiGetDeviceCaps(hDCcaps, WM_HORZRES)
    .DPI = WM_apiGetDeviceCaps(hDCcaps, WM_LOGPIXELSX)
End With
lngRtn = WM_apiReleaseDC(0, hDCcaps) 'Release display context.
    
ScaledFormName = FormName & "_mSF" & GetScreenResolution.Width & "x" & GetScreenResolution.Height & "x" & GetScreenResolution.DPI
For Each frm In CurrentProject.AllForms
    If LCase(frm.NAME) = LCase(ScaledFormName) Then
        FormName = frm.NAME
    End If
Next

DoCmd.OpenForm FormName, View, FilterName, WhereCondition, DataMode, WindowMode, OpenArgs
mSFOpenForm = FormName

End Function





