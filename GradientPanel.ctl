VERSION 5.00
Begin VB.UserControl GradientPanel 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "GradientPanel.ctx":0000
   Begin VB.PictureBox picDrawArea 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      MouseIcon       =   "GradientPanel.ctx":0312
      ScaleHeight     =   795
      ScaleWidth      =   885
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   885
   End
End
Attribute VB_Name = "GradientPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************************************
'Declarations
'**************************************************************************

'Private Control Constants (Used with AboutBox)
Private Const AppName       As String = "Gradient Panel"
Private Const Version       As String = "1.1.0"
Private Const Author        As String = "Stephen Kent"
Private Const SpecialThanks As String = "Special thanks to the people/groups from which I used code from: (In no particular order)" & vbCrLf & _
                                        "Edwin Vermeer, Kath-Rock Software, Microsoft, Stuart Pennington, and Ulli"
Private Const SupInfo       As String = "I have tried to make this as bug free as possible, but I can't guarantee that it is bug free.  If you do find a bug please send me an e-mail with any information you have on it.  SFalcon@Softhome.net" & vbCrLf & _
                                        "NOTE: This has only been tested in the Visual Basic environment and no guarantees are made concerning other environments."

'API Declarations
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RectAPI, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDc As Long, lpRect As RectAPI, ByVal hBrush As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RectAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'API Type Definitions
Private Type RectAPI    'API Rect structure
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'Public Enumerations (For Properties)
Public Enum gpAppearance
    gpaNone = 0
    gpaFlatRaised = 1
    gpaFlatInset = 2
    gpa3DRaised = 3
    gpa3DInset = 4
    gpaEtched = 5
    gpaBevelRaised = 6
    gpaBevelInset = 7
End Enum

Public Enum gpStyle
    gpsStandard = 0
    gpsGradient = 1
    gpsPicture = 2
    gpsTransparent = 3
End Enum

Public Enum gpAlignment
    gpaLeftTop = 0
    gpaLeftMiddle = 1
    gpaLeftBottom = 2
    gpaRightTop = 3
    gpaRightMiddle = 4
    gpaRightBottom = 5
    gpaCenterTop = 6
    gpaCenterMiddle = 7
    gpaCenterBottom = 8
End Enum

Public Enum gpAutoSize
    gpasNone = 0
    gpasPictureToControl = 1
    gpasControlToPicture = 2
End Enum

Public Enum gpCaptionStyle
    gpcStandard = 0
    gpcInsetLight = 1
    gpcInsetHeavy = 2
    gpcRaisedLight = 3
    gpcRaisedHeavy = 4
    gpcDropShadow = 5
End Enum

'DrawText API Constants
Private Const DT_CALCRECT   As Long = &H400     'Used to adjust the bottom of the rectangle to account for all text
Private Const DT_CENTER     As Long = &H1       'Used to center the text
Private Const DT_LEFT       As Long = &H0       'Used to left justify the text
Private Const DT_RIGHT      As Long = &H2       'Used to right justify the text
Private Const DT_SINGLELINE As Long = &H20      'Used to make sure all text is on a single line

'Raster Operation Codes
Private Const DSna = &H220326

'VB Errors
Private Const giINVALID_PICTURE As Integer = 481        'Error code used by Transparent Picture copy routines

'Color Adjustment Constants
Private Const AmountLighten     As Long = 48    'Commonly used AdjustColor amounts
Private Const AmountDarken      As Long = -80

'Default Control Constants
Private Const DefHeight         As Long = 510   'Default initialization height for the control
Private Const DefWidth          As Long = 1230  'Default initialization width for the control

'Local Variables:
Private bAutoSizing             As Boolean      'Used to prevent unnecessary iterations of the Resize event while autosizing.
Private bInitializing           As Boolean      'Used to prevent excess redraws on initialize
Private BorderDark              As Long         'Variable that holds the Dark Border Color
Private BorderDarkest           As Long         'Variable that holds the Darkest Border Color
Private BorderLight             As Long         'Variable that holds the Light Border Color
Private BorderLightest          As Long         'Variable that holds the Lightest Border Color
Private bParentAvailable        As Boolean      'Variable that holds whether the parent object is available with all necessary properties
Private Grad                    As Gradient     'Variable draws all of the gradients (holds instance of class Gradient)
Private GradPic                 As Picture      'Variable to hold the Gradient after it has been drawn

'Default Property Values:
Private Const m_def_Alignment           As Long = gpaCenterMiddle   'Default Alignment: Center Middle (Full Centered)
Private Const m_def_AlignmentCushion    As Long = 3                 'Default Alignment Cushion: 3 Pixels
Private Const m_def_Appearance          As Long = gpaEtched         'Default Appearance: Etched
Private Const m_def_AutoSize            As Long = gpasNone          'Default AutoSizing: None
Private Const m_def_BevelIntensity      As Long = 20                'Default Bevel Intensity: 20
Private Const m_def_BevelWidth          As Long = 3                 'Default Bevel Width: 3 Pixels
Private Const m_def_BorderColor         As Long = vbButtonFace      'Default Border Color: System Button Face
Private Const m_def_Caption             As String = vbNullString    'Default Caption: Empty
Private Const m_def_CaptionStyle        As Long = gpcStandard       'Default Caption Style: Standard
Private Const m_def_GradientAngle       As Double = 0               'Deafult Gradient Angle: 0
Private Const m_def_GradientBlendMode   As Long = gbmRGB            'Default Gradient Blend Mode: RGB Colors
Private Const m_def_GradientColor1      As Long = vbButtonFace      'Default First Gradient Color: System Button Face
Private Const m_def_GradientColor2      As Long = vbButtonFace      'Default Second Gradient Color: System Button Face
Private Const m_def_GradientRepetitions As Double = 1               'Default Gradient Repetitions: 1
Private Const m_def_GradientType        As Long = gtNormal          'Default Gradient Type: Normal (Lines)
Private Const m_def_Style               As Long = gpsStandard       'Default Style: Standard
Private Const m_def_UseClassicBorders   As Boolean = False          'Default Use Classic Borders: False (Using New Borders)

'Property Variables:
Private m_Alignment             As Long             'Local Variable to hold the Alignment of the caption
Private m_AlignmentCushion      As Long             'Local Variable holds the cushion between the edge of the control and the Caption for non-centered alignments
Private m_Appearance            As Long             'Local Variable to hold the Appearance of the control
Private m_AutoSize              As gpAutoSize       'Local Variable to hold the autosizing mode used by the control/Background Picture [Used only for Picture and Graphical Picture Modes]
Private m_BevelIntensity        As Long             'Local Variable to hold the Intesity of the bevel border
Private m_BevelWidth            As Long             'Local Variable to hold the width of the bevel border
Private m_BorderColor           As Long             'Local Variable to hold the border color
Private m_Caption               As String           'Local Variable to hold the Caption
Private m_CaptionStyle          As gpCaptionStyle   'Local Variable to hold the Caption display style
Private m_GradientAngle         As Double           'Local Variable to hold the angle to draw the gradient at
Private m_GradientBlendMode     As GradBlendMode    'Local Variable hold the blending mode to use for gradients
Private m_GradientColor1        As Long             'Local Variable to hold the first gradient color
Private m_GradientColor2        As Long             'Local Variable to hold the second gradient color
Private m_GradientRepetitions   As Double           'Local Variable holds the number of times to repeat the gradient across the button
Private m_GradientType          As GradType         'Local Variable holds the type of gradient to draw. (Only used for Gradient Modes)
Private m_Picture               As Picture          'Local Variable to hold the picture for the picture style
Private m_Style                 As Long             'Local Variable to hold the style of the control
Private m_UseClassicBorders     As Boolean          'Local Variable to hold whether or not to use classic borders

'Event Declarations:
Public Event Click()    'Event fired when the control is clicked
Attribute Click.VB_Description = "Event fired when the user clicks the control."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Public Event DblClick() 'Event fired when the control is double clicked
Attribute DblClick.VB_Description = "Event fired when the user double clicks the control."
Attribute DblClick.VB_UserMemId = -601
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)   'Event fired when a mouse button is pressed over the control
Attribute MouseDown.VB_Description = "Event fired when the user presses a mouse button over the control."
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)   'Event fired when the mouse is moved in the area over the control
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)     'Event fired when a mouse button is released while over the control
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)                   'Event fired when an OLE Drag Drop operation starts
Attribute OLEStartDrag.VB_Description = "Event fired when an OLE Drag and Drop operation starts."
Public Event OLESetData(Data As DataObject, DataFormat As Integer)                      'Event fired when data is requested from an OLE operation that was not previously set
Attribute OLESetData.VB_Description = "Event fired when the Target requests data that has not yet been set in the OLE object."
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)                 'Event fired after the OLE operation completes and the Target wants to return some info to the Source
Attribute OLEGiveFeedback.VB_Description = "Event fired after an OLE drag operation completes and the Target needs to return data to the Source."
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)   'Event fired when an OLE object is dragged over the control
Attribute OLEDragOver.VB_Description = "Event fired when an OLE object is dragged over the control."
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)     'Event fired when an OLE object is dropped on the control
Attribute OLEDragDrop.VB_Description = "Event fired when an OLE object is dropped on the control."
Public Event OLECompleteDrag(Effect As Long)        'Event fired when the OLE operation has been completed
Attribute OLECompleteDrag.VB_Description = "Event fired when an OLE Drag and Drop operation is completed."
Public Event Resize()   'Event fired when the control is resized
Attribute Resize.VB_Description = "Event fired when the control is resized."

'**************************************************************************
'Properties & Methods
'**************************************************************************

'Sub-Procedure to display information about the control.
Public Sub About()
Attribute About.VB_Description = "Displays an about box giving information about the control."
Attribute About.VB_UserMemId = -552
    Load frmAbout   'Load aboutbox in background
    frmAbout.strApplication = AppName   'Assign the name of the application to the about box
    frmAbout.strVersion = Version       'Assign the version number of the program
    frmAbout.strAuthor = Author         'Assign the about box the author of the control (Me)
    frmAbout.strThanks = SpecialThanks  'List any and all special thanks recipients
    frmAbout.strAddInfo = SupInfo       'Assign the supplemental information/disclaimer
    Set frmAbout.picLogo = picDrawArea.MouseIcon    'Set the Icon to display on the about box (MouseIcon was used because TrueColor icons don't store well in res files.
    frmAbout.Show 1         'Open the aboutbox as modal
End Sub

'Property set for Caption Alignment
Public Property Get Alignment() As gpAlignment
Attribute Alignment.VB_Description = "Returns/sets the alignment of the caption."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As gpAlignment)
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint to update the control with the caption's new positioning
End Property
'End Alignment Property set

'Property set for Alignment cushion for the Caption
Public Property Get AlignmentCushion() As Long
Attribute AlignmentCushion.VB_Description = "Returns/sets the number of pixels between the Caption and the edge of the button as a cushion zone."
Attribute AlignmentCushion.VB_ProcData.VB_Invoke_Property = ";Appearance"
    AlignmentCushion = m_AlignmentCushion
End Property

Public Property Let AlignmentCushion(ByVal New_AlignmentCushion As Long)
    If New_AlignmentCushion < 0 Then Exit Property  'Make sure we have a valid cushion
    m_AlignmentCushion = New_AlignmentCushion
    PropertyChanged "AlignmentCushion"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint for the caption to be redrawn with the new alignment
End Property
'End AlignmentCushion Property set

'Property set for Border Appearance
Public Property Get Appearance() As gpAppearance
Attribute Appearance.VB_Description = "Returns/sets the border appearance of the pane.  (None / Flat Raised / Flat Inset / 3D Raised / 3D Inset / Etched / Bevel Raised / Bevel Inset)"
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As gpAppearance)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint to update the control with the new border appearance
End Property
'End Appearance Property set

'Property set for the autosizing property of the control
Public Property Get AutoSize() As gpAutoSize
Attribute AutoSize.VB_Description = "Returns/sets what AutoSizing method to use for the background picture / control.  (Only used for Picture and Graphical Picture modes) [None / Picture to Control / Control to Picture]."
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As gpAutoSize)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    UserControl_Resize  'Call the resize event so that the control is resized if necessary
End Property
'End BackColor Property set

'Property set for Back Color
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used for the control.  Also sets the transparency color for Style: Transparent."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = picDrawArea.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picDrawArea.BackColor() = New_BackColor
    PropertyChanged "BackColor"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If (m_Style = gpsStandard) Or (m_Style = gpsPicture) Or (m_Style = gpsTransparent) Then PaintFast   'If the style is one of the style's affected by a background color change then do a fast paint to update
End Property
'End BackColor Property set

'Property set for Bevel Intensity (3D effect)
Public Property Get BevelIntensity() As Long
Attribute BevelIntensity.VB_Description = "Returns/set the intensity of the bevel border.  (Higher intensity leads to a greater 3D effect with a smaller width)"
Attribute BevelIntensity.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BevelIntensity = m_BevelIntensity
End Property

Public Property Let BevelIntensity(ByVal New_BevelIntensity As Long)
    m_BevelIntensity = New_BevelIntensity
    PropertyChanged "BevelIntensity"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If (m_Appearance = gpaBevelRaised) Or (m_Appearance = gpaBevelInset) Then PaintFast     'If the border is a bevel then do a fast paint to update with the new settings
End Property
'End BevelIntensity Property set

'Property set for Bevel Width (In Pixels)
Public Property Get BevelWidth() As Long
Attribute BevelWidth.VB_Description = "Returns/set the width of the Beveled border."
Attribute BevelWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BevelWidth = m_BevelWidth
End Property

Public Property Let BevelWidth(ByVal New_BevelWidth As Long)
    m_BevelWidth = New_BevelWidth
    PropertyChanged "BevelWidth"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If (m_Appearance = gpaBevelRaised) Or (m_Appearance = gpaBevelInset) Then PaintFast     'If the border is a bevel then do a fast paint to update with the new settings
End Property
'End BevelWidth Property set

'Property set for Border Color
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color to use when drawing the borders except for bevel.  (Only used for Classic Borders)"
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderColor.VB_UserMemId = -503
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    SetColors   'call setcolors to update the internal variables with the new color scheme
    If (m_UseClassicBorders) Then PaintFast     'If using classic borders then do a fast paint to update the control
End Property
'End BorderColor Property set

'Property set for Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text for the caption."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint to update the control with the new caption
End Property
'End Caption Property set

'Property set for Caption Style
Public Property Get CaptionStyle() As gpCaptionStyle
Attribute CaptionStyle.VB_Description = "Returns/sets the Caption Style for the button.  (Standard / Light Inset / Heavy Inset / Light Raised / Heavy Raised / Drop Shadow)"
Attribute CaptionStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CaptionStyle = m_CaptionStyle
End Property

Public Property Let CaptionStyle(ByVal New_CaptionStyle As gpCaptionStyle)
    If (New_CaptionStyle < gpcStandard) Or (New_CaptionStyle > gpcDropShadow) Then Exit Property    'If new value isn't valid then exit the property doing nothing
    m_CaptionStyle = New_CaptionStyle
    PropertyChanged "CaptionStyle"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint to refresh the caption with the new style
End Property
'End CaptionStyle Property set

'Property set for whether the control is enabled or not
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
End Property
'End Enabled Property set

'Property set for the font used for the caption
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = picDrawArea.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set picDrawArea.Font = New_Font
    PropertyChanged "Font"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint to update the caption with the new font
End Property
'End Font Property set

'Property set for the bold attribute of the font
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
Attribute FontBold.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontBold.VB_MemberFlags = "400"
    FontBold = picDrawArea.Font.Bold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    picDrawArea.Font.Bold = New_FontBold
    PropertyChanged "FontBold"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint to update the caption with the new font
End Property
'End FontBold Property set

'Property set for the italic attribute of the font
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
Attribute FontItalic.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontItalic.VB_MemberFlags = "400"
    FontItalic = picDrawArea.Font.Italic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    picDrawArea.Font.Italic = New_FontItalic
    PropertyChanged "FontItalic"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint to update the caption with the new font
End Property
'End FontItalic Property set

'Property set for the name of the font
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
Attribute FontName.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontName.VB_MemberFlags = "400"
    FontName = picDrawArea.Font.Name
End Property

Public Property Let FontName(ByVal New_FontName As String)
    picDrawArea.Font.Name = New_FontName
    PropertyChanged "FontName"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint to update the caption with the new font
End Property
'End FontName Property set

'Property set for the size attribute of the font
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
Attribute FontSize.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontSize.VB_MemberFlags = "400"
    FontSize = picDrawArea.Font.Size
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    picDrawArea.Font.Size = New_FontSize
    PropertyChanged "FontSize"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint to update the caption with the new font
End Property
'End FontSize Property set

'Property set for the strike through attribute of the font
Public Property Get FontStrikethrough() As Boolean
Attribute FontStrikethrough.VB_Description = "Returns/sets strikethrough font styles."
Attribute FontStrikethrough.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontStrikethrough.VB_MemberFlags = "400"
    FontStrikethru = picDrawArea.Font.Strikethrough
End Property

Public Property Let FontStrikethrough(ByVal New_FontStrikethrough As Boolean)
    picDrawArea.Font.Strikethrough = New_FontStrikethrough
    PropertyChanged "FontStrikethrough"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint to update the caption with the new font
End Property
'End FontStrikeThrough Property set

'Property set for the Underline attribute of the font
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
Attribute FontUnderline.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute FontUnderline.VB_MemberFlags = "400"
    FontUnderline = picDrawArea.Font.Underline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    picDrawArea.Font.Underline = New_FontUnderline
    PropertyChanged "FontUnderline"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint to update the caption with the new font
End Property
'End FontUnderline Property set

'Property set for the Fore Color of the caption
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = picDrawArea.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    picDrawArea.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast   'Do a fast paint to update the caption with the new font
End Property
'End ForeColor Property set

'Property set for the angle that the gradient is to be drawn at
Public Property Get GradientAngle() As Double
Attribute GradientAngle.VB_Description = "Returns/set the angle at which to draw the gradient.  (Only used for Gradient mode)"
Attribute GradientAngle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GradientAngle = m_GradientAngle
End Property

Public Property Let GradientAngle(ByVal New_GradientAngle As Double)
    m_GradientAngle = New_GradientAngle
    PropertyChanged "GradientAngle"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If (m_Style = gpsGradient) Then PaintAll    'If style is in gradient mode then do a paint all to update the gradient with the new settings
End Property
'End GradientAngle Property set

'Property set for the Gradient Blending Mode
Public Property Get GradientBlendMode() As GradBlendMode
Attribute GradientBlendMode.VB_Description = "Returns/sets the blending mode to use with the gradient modes."
Attribute GradientBlendMode.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GradientBlendMode = m_GradientBlendMode
End Property

Public Property Let GradientBlendMode(ByVal New_GradientBlendMode As GradBlendMode)
    If (New_GradientBlendMode < gbmRGB) Or (New_GradientBlendMode > gbmHSL) Then Exit Property  'Make sure we have a valid value
    m_GradientBlendMode = New_GradientBlendMode
    PropertyChanged "GradientBlendMode"
    PaintAll        'Do a paint all so that the Gradient is redrawn
End Property
'End GradientBlendMode Property set

'Property set for the First Gradient Color
Public Property Get GradientColor1() As OLE_COLOR
Attribute GradientColor1.VB_Description = "Returns/sets the first color to use in the gradient.  (Only used for Gradient mode)"
Attribute GradientColor1.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GradientColor1 = m_GradientColor1
End Property

Public Property Let GradientColor1(ByVal New_GradientColor1 As OLE_COLOR)
    m_GradientColor1 = New_GradientColor1
    PropertyChanged "GradientColor1"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If (m_Style = gpsGradient) Then PaintAll    'If style is in gradient mode then do a paint all to update the gradient with the new settings
End Property
'End GradientColor1 Property set

'Property set for the Second Gradient Color
Public Property Get GradientColor2() As OLE_COLOR
Attribute GradientColor2.VB_Description = "Returns/sets the second color to use in the gradient.  (Only used for Gradient mode)"
Attribute GradientColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GradientColor2 = m_GradientColor2
End Property

Public Property Let GradientColor2(ByVal New_GradientColor2 As OLE_COLOR)
    m_GradientColor2 = New_GradientColor2
    PropertyChanged "GradientColor2"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If (m_Style = gpsGradient) Then PaintAll    'If style is in gradient mode then do a paint all to update the gradient with the new settings
End Property
'End GradientColor2 Property set

'Property set for number of times to repeat the gradient
Public Property Get GradientRepetitions() As Double
Attribute GradientRepetitions.VB_Description = "Returns/sets the number times to show the gradient.  Valid values are 1 to 255.  (Used only for Gradient and Graphical Gradient Modes)"
Attribute GradientRepetitions.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GradientRepetitions = m_GradientRepetitions
End Property

Public Property Let GradientRepetitions(ByVal New_GradientRepetitions As Double)
    If (New_GradientRepetitions < 1) Or (New_GradientRepetitions > 45) Then Exit Property   'Make sure within valid range
    m_GradientRepetitions = New_GradientRepetitions
    PropertyChanged "GradientRepetitions"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintAll        'Do a paint all so that the Gradient is redrawn
End Property
'End GradientRepetitions Property set

'Property set for the type of gradient
Public Property Get GradientType() As GradType
Attribute GradientType.VB_Description = "Returns/sets the type of Gradient to Draw (Normal / Elliptical / Rectangular)"
Attribute GradientType.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GradientType = m_GradientType
End Property

Public Property Let GradientType(ByVal New_GradientType As GradType)
    If (New_GradientType < gtNormal) Or (New_GradientType > gtRectangular) Then Exit Property   'Make sure we have a valid value
    m_GradientType = New_GradientType
    PropertyChanged "GradientType"        'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintAll        'Do a paint all so that the Gradient is redrawn
End Property
'End GradientType Property set

'Property to get the control's device context handle
Public Property Get hDc() As Long
Attribute hDc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
Attribute hDc.VB_ProcData.VB_Invoke_Property = ";Data"
    hDc = UserControl.hDc       'Pass along the control's device context for the developer to use with the windows API if necessary
End Property

'Property to get the control's window handle
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hWnd.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd     'Pass along the control's window handle for the developer to use with the windows API if necessary
End Property

'Property set for Mouse Icon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
End Property
'End MouseIcon Property set

'Property set for the Mouse Pointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"      'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
End Property
'End MousePointer Property set

'Property set for OLE Drop Mode
Public Property Get OLEDropMode() As OLEDropConstants
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
Attribute OLEDropMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
End Property
'End OLEDropMode Property set

'Method to allow user to start an OLE drag operation
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag     'Start an OLE drag operation
End Sub

'Property set for Picture (to use as background)
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"       'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    If (m_Style = gpsPicture) Then PaintFast        'If the style is Picture then do a fast paint to update the control
End Property
'End Picture Property set

'Method to force control to redraw/repaint itself
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
     PaintAll       'Do a paint all because the user has called for a refresh of the control
End Sub

'Property set for Control Style
Public Property Get Style() As gpStyle
Attribute Style.VB_Description = "Returns/set the style of the control.  (Standard / Gradient / Picture / Transparent)"
Attribute Style.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As gpStyle)
    m_Style = New_Style
    PropertyChanged "Style"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintAll        'Do a paint all to update control with the new style
End Property
'End Style Property set

'Property set for Use of Classic Borders
Public Property Get UseClassicBorders() As Boolean
Attribute UseClassicBorders.VB_Description = "Returns/sets whether to use the original style of borders or the blending style.  (Not valid for Beveled borders)"
Attribute UseClassicBorders.VB_ProcData.VB_Invoke_Property = ";Appearance"
    UseClassicBorders = m_UseClassicBorders
End Property

Public Property Let UseClassicBorders(ByVal New_UseClassicBorders As Boolean)
    m_UseClassicBorders = New_UseClassicBorders
    PropertyChanged "UseClassicBorders"     'Signal that Property has been assigned (not necessarily changed because I'm not checking to see if changed)
    PaintFast       'Do a fast paint so that the borders may be updated if necessary
End Property
'End UseClassicBorders Property set

'**************************************************************************
'Internal Event Coding & Implementation
'**************************************************************************

Private Sub UserControl_Click()
    RaiseEvent Click        'Pass the event along to the developer
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick     'Pass the event along to the developer
End Sub

Private Sub UserControl_Initialize()
    Set Grad = New Gradient     'Initialize the gradient object whenever the control is created
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    On Error Resume Next
    bInitializing = True    'Set initializing flag so we don't have unnecessary redraws
    'Initialize variables with defaults
    m_Alignment = m_def_Alignment
    m_AlignmentCushion = m_def_AlignmentCushion
    m_Appearance = m_def_Appearance
    m_AutoSize = m_def_AutoSize
    m_BevelIntensity = m_def_BevelIntensity
    m_BevelWidth = m_def_BevelWidth
    m_BorderColor = m_def_BorderColor
    m_Caption = Ambient.DisplayName     'Get display name from VB
    m_CaptionStyle = m_def_CaptionStyle
    Set picDrawArea.Font = Ambient.Font 'Get the current default font from VB
    m_GradientAngle = m_def_GradientAngle
    m_GradientBlendMode = m_def_GradientBlendMode
    m_GradientColor1 = m_def_GradientColor1
    m_GradientColor2 = m_def_GradientColor2
    m_GradientRepetitions = m_def_GradientRepetitions
    m_GradientType = m_def_GradientType
    Set m_Picture = Nothing
    m_Style = m_def_Style
    m_UseClassicBorders = m_def_UseClassicBorders
    Height = DefHeight      'This should give an actual default height of 525 twips
    Width = DefWidth        'This should give an actual default width of 1245 twips
    CheckParent             'Check to make sure we have parent information available
    bInitializing = False   'We're done intializing so we can remove the flag so the control will update properly
    PaintAll        'Do an initial paint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    If bParentAvailable Then    'If parent information is available then
        RaiseEvent MouseDown(Button, Shift, ScaleX(X, ScaleMode, UserControl.Parent.ScaleMode), ScaleY(Y, ScaleMode, UserControl.Parent.ScaleMode))
    Else
        RaiseEvent MouseDown(Button, Shift, ScaleX(X, ScaleMode, vbTwips), ScaleY(Y, ScaleMode, vbTwips))
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    Dim PassX As Single     'Temporary variable to hold the translated X co-ordinate
    Dim PassY As Single     'Temporary variable to hold the translated Y co-ordinate

    If bParentAvailable Then    'If parent information is available then
        PassX = ScaleX(X, ScaleMode, UserControl.Parent.ScaleMode)      'Translate the X co-ordinate to match the parent's scale mode
        PassY = ScaleY(Y, ScaleMode, UserControl.Parent.ScaleMode)      'Translate the Y co-ordinate to match the parent's scale mode
    Else
        PassX = ScaleX(X, ScaleMode, vbTwips)   'Translate the X co-ordinate to twips
        PassY = ScaleY(Y, ScaleMode, vbTwips)   'Translate the Y co-ordinate to twips
    End If
    RaiseEvent MouseMove(Button, Shift, PassX, PassY)   'Now raise the MouseMove event
End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    If bParentAvailable Then    'If parent information is available then
        RaiseEvent MouseUp(Button, Shift, ScaleX(X, ScaleMode, UserControl.Parent.ScaleMode), ScaleY(Y, ScaleMode, UserControl.Parent.ScaleMode))       'Pass the event along to the developer (Although rescale the X and Y co-ordinates)
    Else
        RaiseEvent MouseUp(Button, Shift, ScaleX(X, ScaleMode, vbTwips), ScaleY(Y, ScaleMode, vbTwips))       'Pass the event along to the developer (Although rescale the X and Y co-ordinates)
    End If
End Sub


Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)      'Pass the event along to the developer
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, _
                                Effect As Long, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    If bParentAvailable Then    'If parent information is available then
        RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, ScaleX(X, ScaleMode, UserControl.Parent.ScaleMode), ScaleY(Y, ScaleMode, UserControl.Parent.ScaleMode))     'Pass the event along to the developer (Although rescale the X and Y co-ordinates)
    Else
        RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, ScaleX(X, ScaleMode, vbTwips), ScaleY(Y, ScaleMode, vbTwips))   'Pass the event along to the developer (After rescale to twips)
    End If
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, _
                                Effect As Long, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single, _
                                State As Integer)
    If bParentAvailable Then    'If parent information is available then
        RaiseEvent OLEDragOver(Data, Effect, Button, Shift, ScaleX(X, ScaleMode, UserControl.Parent.ScaleMode), ScaleY(Y, ScaleMode, UserControl.Parent.ScaleMode), State)      'Pass the event along to the developer (Although rescale the X and Y co-ordinates)
    Else
        RaiseEvent OLEDragOver(Data, Effect, Button, Shift, ScaleX(X, ScaleMode, vbTwips), ScaleY(Y, ScaleMode, vbTwips), State)    'Pass the event along to the developer (After rescale to twips)
    End If
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, _
                                DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)      'Pass the event along to the developer
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, _
                                DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)     'Pass the event along to the developer
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, _
                                AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)       'Pass the event along to the developer
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    'Read properties from persisted data
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    m_AlignmentCushion = PropBag.ReadProperty("AlignmentCushion", m_def_AlignmentCushion)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    picDrawArea.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    m_BevelIntensity = PropBag.ReadProperty("BevelIntensity", m_def_BevelIntensity)
    m_BevelWidth = PropBag.ReadProperty("BevelWidth", m_def_BevelWidth)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_CaptionStyle = PropBag.ReadProperty("CaptionStyle", m_def_CaptionStyle)
    Me.Enabled = PropBag.ReadProperty("Enabled", True)
    Set picDrawArea.Font = PropBag.ReadProperty("Font", Ambient.Font)
    picDrawArea.ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
    m_GradientAngle = PropBag.ReadProperty("GradientAngle", m_def_GradientAngle)
    m_GradientBlendMode = PropBag.ReadProperty("GradientBlendMode", m_def_GradientBlendMode)
    m_GradientColor1 = PropBag.ReadProperty("GradientColor1", m_def_GradientColor1)
    m_GradientColor2 = PropBag.ReadProperty("GradientColor2", m_def_GradientColor2)
    m_GradientRepetitions = PropBag.ReadProperty("GradientRepetitions", m_def_GradientRepetitions)
    m_GradientType = PropBag.ReadProperty("GradientType", m_def_GradientType)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    m_UseClassicBorders = PropBag.ReadProperty("UseClassicBorders", m_def_UseClassicBorders)
    CheckParent             'Check to make sure we have parent information available
    PaintAll
End Sub

Private Sub UserControl_Resize()
    If Not (bAutoSizing) Then
        If (m_Style = gpsPicture) And (m_AutoSize = gpasControlToPicture) Then  'If it is a Picture mode and we are autosizing control to picture then
            If Not (m_Picture Is Nothing) Then
                bAutoSizing = True
                Width = ScaleX(m_Picture.Width, vbHimetric, vbTwips)
                Height = ScaleY(m_Picture.Height, vbHimetric, vbTwips)
                bAutoSizing = False
            End If
        End If
        'We need to execute a resize to make sure the control looks correct
        picDrawArea.Width = ScaleWidth      'Resize the Drawing area to be the same size as the control
        picDrawArea.Height = ScaleHeight
        If Not bInitializing Then   'If we're not initializing the control then
            PaintAll                'Do a paint all to refresh the complete control and acount for size differences
        End If
        RaiseEvent Resize       'Pass the event along to the developer
    End If
End Sub

Private Sub UserControl_Show()
    PaintFast   'Do a fast paint after the control is shown making sure that a correct transparency is made
End Sub

Private Sub UserControl_Terminate()
    Set Grad = Nothing      'Kill the Gradient Object as the button is being destroyed
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    'Write properties to be persisted (saved for later load)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
    Call PropBag.WriteProperty("AlignmentCushion", m_AlignmentCushion, m_def_AlignmentCushion)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("BackColor", picDrawArea.BackColor, vbButtonFace)
    Call PropBag.WriteProperty("BevelIntensity", m_BevelIntensity, m_def_BevelIntensity)
    Call PropBag.WriteProperty("BevelWidth", m_BevelWidth, m_def_BevelWidth)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("CaptionStyle", m_CaptionStyle, m_def_CaptionStyle)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", picDrawArea.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", picDrawArea.ForeColor, vbButtonText)
    Call PropBag.WriteProperty("GradientAngle", m_GradientAngle, m_def_GradientAngle)
    Call PropBag.WriteProperty("GradientBlendMode", m_GradientBlendMode, m_def_GradientBlendMode)
    Call PropBag.WriteProperty("GradientColor1", m_GradientColor1, m_def_GradientColor1)
    Call PropBag.WriteProperty("GradientColor2", m_GradientColor2, m_def_GradientColor2)
    Call PropBag.WriteProperty("GradientRepetitions", m_GradientRepetitions, m_def_GradientRepetitions)
    Call PropBag.WriteProperty("GradientType", m_GradientType, m_def_GradientType)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
    Call PropBag.WriteProperty("UseClassicBorders", m_UseClassicBorders, m_def_UseClassicBorders)
End Sub

'**************************************************************************
'Internal Routines
'**************************************************************************

Private Function AdjustColor(ByVal RGBColor As Long, _
                            ByVal Amount As Long) As Long
    Dim Blue As Long    'Variable to hold the Blue Value while in this procedure
    Dim Green As Long   'Variable to hold the Green Value while in this procedure
    Dim Red As Long     'Variable to hold the Red Value while in this procedure

    If (RGBColor = (RGBColor Or &H80000000)) Then RGBColor = GetSysColor(RGBColor Xor &H80000000)   'If working with a system color get the RGB equivilent
    RGBColor = Abs(RGBColor)        'Make sure working with a positive number
    Blue = RGBColor \ 65536         'Seperate out the Blue
    RGBColor = RGBColor Mod 65536   'Remove Blue from Value
    Green = RGBColor \ 256          'Seperate out the Green
    Red = RGBColor Mod 256          'Remove Green from Value and place in Red
    Red = Red + Amount              'Change Red by Amount
    If Red < 0 Then Red = 0         'If less than 0 then make it 0
    If Red > 255 Then Red = 255     'If greater than 255 then make it 255
    Green = Green + Amount          'Change Green by Amount
    If Green < 0 Then Green = 0     'If less than 0 then make it 0
    If Green > 255 Then Green = 255 'If greater than 255 then make it 255
    Blue = Blue + Amount            'Change Blue by Amount
    If Blue < 0 Then Blue = 0       'If less than 0 then make it 0
    If Blue > 255 Then Blue = 255   'If greater than 255 then make it 255
    AdjustColor = RGB(Red, Green, Blue)  'Combine the colors and pass back the value
End Function

Private Sub Border3D(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
    If m_UseClassicBorders Then     'If we're using the old style borders then
        Border3DOld DefaultOffset, Depressed    'Call Old Border Draw Routine
    Else                            'Otherwise
        Border3DNew DefaultOffset, Depressed    'Call New Border Draw Routine
    End If
End Sub

Private Sub Border3DNew(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
    Dim hPic As Long            'Variable to hold the Device Context handle of the drawing area
    Dim CurColor As Long        'Variable to hold the color of the pixel currently being worked with
    Dim i As Long               'Variable used as a counter in the for loops
    Dim picWidth As Long        'Variable to hold the width of the drawing area in pixels
    Dim picHeight As Long       'Variable to hold the height of the drawing area in pixels

    hPic = picDrawArea.hDc          'Get the Device Context handle and store it in a local variable for quick access
    With picDrawArea
        picWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)   'Get the width of the drawing area in pixels
        picHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels) 'Get the height of the drawing area in pixels
    End With
    If Depressed Then   'If button is depressed then
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Horizontal points from Left to Right on the inside of the border
        For i = (DefaultOffset + 1) To (picWidth - DefaultOffset - 2)
            CurColor = GetPixel(hPic, i, DefaultOffset + 1) 'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset + 1, AdjustColor(CurColor, AmountDarken)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 2)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 2, AdjustColor(CurColor, AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the inside of the border
        For i = (DefaultOffset + 2) To (picHeight - DefaultOffset - 3)
            CurColor = GetPixel(hPic, DefaultOffset + 1, i) 'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset + 1, i, AdjustColor(CurColor, AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 2, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 2, i, AdjustColor(CurColor, AmountLighten)
        Next
    Else        'Button should be drawn up
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountLighten)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountDarken)
        Next
        'Start a loop through the Horizontal points from Left to Right on the inside of the border
        For i = (DefaultOffset + 1) To (picWidth - DefaultOffset - 2)
            CurColor = GetPixel(hPic, i, DefaultOffset + 1) 'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset + 1, AdjustColor(CurColor, AmountLighten)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 2)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 2, AdjustColor(CurColor, AmountDarken)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountLighten)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountDarken)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the inside of the border
        For i = (DefaultOffset + 2) To (picHeight - DefaultOffset - 3)
            CurColor = GetPixel(hPic, DefaultOffset + 1, i) 'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset + 1, i, AdjustColor(CurColor, AmountLighten)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 2, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 2, i, AdjustColor(CurColor, AmountDarken)
        Next
    End If
End Sub

Private Sub Border3DOld(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
    If Depressed Then   'If button is depressed then
        picDrawArea.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest, B    'Draw a box that could be considered the inside of the border and will become the Bottom and Right Inside lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLight, B               'Draw a box that could be considered the outside of the border and will become the Bottom and Right Outside lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode)), BorderDarkest      'Draw two lines that will be the Top and Left Outside border lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest
        picDrawArea.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDark     'Draw two lines that will be the Top and Left Inside border lines
        picDrawArea.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderDark
    Else        'Button should be drawn up
        picDrawArea.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderDark, B    'Draw a box that could be considered the inside of the border and will become the Bottom and Right Inside lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest, B         'Draw a box that could be considered the outside of the border and will become the Bottom and Right Outside lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode)), BorderLightest     'Draw two lines that will be the Top and Left Outside border lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest
        picDrawArea.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLight    'Draw two lines that will be the Top and Left Inside border lines
        picDrawArea.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderLight
    End If
End Sub

Private Sub BorderBevel(ByVal DefaultOffset As Long, _
                                ByVal BeveledWidth As Long, _
                                Optional ByVal Depressed As Boolean = False)
    Dim hPic As Long            'Variable to hold the Device Context handle of the drawing area
    Dim CurColor As Long        'Variable to hold the color of the pixel currently being worked with
    Dim i As Long               'Variable used as a counter in the for loops
    Dim picWidth As Long        'Variable to hold the width of the drawing area in pixels
    Dim picHeight As Long       'Variable to hold the height of the drawing area in pixels
    Dim ColorAdjust As Long     'Variable containing the amount to change the color of each pixel
    Dim modDepress As Integer   'Variable to hold a modifier determining if the border should be drawn up or depressed (color adjusting)

    ColorAdjust = m_BevelIntensity  'Initialize the color adjustment (Intensity)
    hPic = picDrawArea.hDc          'Get the Device Context handle and store it in a local variable for quick access
    With picDrawArea
        picWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)   'Get the width of the drawing area in pixels
        picHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels) 'Get the height of the drawing area in pixels
    End With
    If Depressed Then       'If border should be draw as depressed then
        modDepress = -1     'Set a negative modifier to reverse the color adjustment
    Else                    'Otherwise
        modDepress = 1      'Set a positive modifier to leave color adjustment alone
    End If
    Do      'Loop through the following block until we've finished the bevel (BeveledWidth = -1) moving from the inside > outside
        'Start a loop through the Horizontal points from Left to Right on the y positions of the Bevel
        For i = (BeveledWidth + DefaultOffset) To (picWidth - BeveledWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, BeveledWidth + DefaultOffset)  'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the left side
            SetPixelV hPic, i, BeveledWidth + DefaultOffset, AdjustColor(CurColor, ColorAdjust * modDepress)
            CurColor = GetPixel(hPic, i, picHeight - BeveledWidth - DefaultOffset - 1)  'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, i, picHeight - BeveledWidth - DefaultOffset - 1, AdjustColor(CurColor, -ColorAdjust * modDepress)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the x positions of the Bevel
        For i = (BeveledWidth + DefaultOffset + 1) To (picHeight - BeveledWidth - DefaultOffset - 2)
            CurColor = GetPixel(hPic, BeveledWidth + DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the top
            SetPixelV hPic, BeveledWidth + DefaultOffset, i, AdjustColor(CurColor, ColorAdjust * modDepress)
            CurColor = GetPixel(hPic, picWidth - BeveledWidth - DefaultOffset - 1, i)   'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the bottom
            SetPixelV hPic, picWidth - BeveledWidth - DefaultOffset - 1, i, AdjustColor(CurColor, -ColorAdjust * modDepress)
        Next
        BeveledWidth = BeveledWidth - 1     'Reduce the bevel width left to draw
        ColorAdjust = ColorAdjust + m_BevelIntensity    'Increase the Intensity of the color change
    Loop Until BeveledWidth = -1
End Sub

Private Sub BorderEtched(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
    If m_UseClassicBorders Then     'If we're using the old style borders then
        BorderEtchedOld DefaultOffset, Depressed    'Call Old Border Draw Routine
    Else                            'Otherwise
        BorderEtchedNew DefaultOffset, Depressed    'Call New Border Draw Routine
    End If
End Sub

Private Sub BorderEtchedNew(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
    Dim hPic As Long            'Variable to hold the Device Context handle of the drawing area
    Dim CurColor As Long        'Variable to hold the color of the pixel currently being worked with
    Dim i As Long               'Variable used as a counter in the for loops
    Dim picWidth As Long        'Variable to hold the width of the drawing area in pixels
    Dim picHeight As Long       'Variable to hold the height of the drawing area in pixels

    hPic = picDrawArea.hDc          'Get the Device Context handle and store it in a local variable for quick access
    With picDrawArea
        picWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)   'Get the width of the drawing area in pixels
        picHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels) 'Get the height of the drawing area in pixels
    End With
    If Depressed Then   'If button is depressed then
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Horizontal points from Left to Right on the inside of the border
        For i = (DefaultOffset + 1) To (picWidth - DefaultOffset - 2)
            CurColor = GetPixel(hPic, i, DefaultOffset + 1) 'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset + 1, AdjustColor(CurColor, AmountDarken)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 2)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 2, AdjustColor(CurColor, AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the inside of the border
        For i = (DefaultOffset + 2) To (picHeight - DefaultOffset - 3)
            CurColor = GetPixel(hPic, DefaultOffset + 1, i) 'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset + 1, i, AdjustColor(CurColor, AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 2, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 2, i, AdjustColor(CurColor, AmountLighten)
        Next
    Else        'Button should be drawn up
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            If Not (i = (picWidth - DefaultOffset - 1)) Then
                SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountDarken)
            Else
                SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountLighten)
            End If
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Horizontal points from Left to Right on the inside of the border
        For i = (DefaultOffset + 1) To (picWidth - DefaultOffset - 2)
            CurColor = GetPixel(hPic, i, DefaultOffset + 1) 'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            If Not (i = (picWidth - DefaultOffset - 2)) Then
                SetPixelV hPic, i, DefaultOffset + 1, AdjustColor(CurColor, 2 * AmountLighten)
            Else
                SetPixelV hPic, i, DefaultOffset + 1, AdjustColor(CurColor, 2 * AmountDarken)
            End If
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 2)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 2, AdjustColor(CurColor, 2 * AmountDarken)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the inside of the border
        For i = (DefaultOffset + 2) To (picHeight - DefaultOffset - 3)
            CurColor = GetPixel(hPic, DefaultOffset + 1, i) 'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset + 1, i, AdjustColor(CurColor, 2 * AmountLighten)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 2, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 2, i, AdjustColor(CurColor, 2 * AmountDarken)
        Next
    End If
End Sub

Private Sub BorderEtchedOld(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
    If Depressed Then   'If button is depressed then
        picDrawArea.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest, B    'Draw a box that could be considered the inside of the border and will become the Bottom and Right Inside lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLight, B               'Draw a box that could be considered the outside of the border and will become the Bottom and Right Outside lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode)), BorderDarkest          'Draw two lines that will be the Top and Left Outside border lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest
        picDrawArea.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDark     'Draw two lines that will be the Top and Left Inside border lines
        picDrawArea.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderDark
    Else        'Button should be drawn up
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest, B
        picDrawArea.Line (ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(1 + DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest, B
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(2 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(2 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest, B
    End If
End Sub

Private Sub BorderFlat(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
    If m_UseClassicBorders Then     'If we're using the old style borders then
        BorderFlatOld DefaultOffset, Depressed  'Call Old Border Draw Routine
    Else                            'Otherwise
        BorderFlatNew DefaultOffset, Depressed  'Call New Border Draw Routine
    End If
End Sub

Private Sub BorderFlatNew(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
    Dim hPic As Long            'Variable to hold the Device Context handle of the drawing area
    Dim CurColor As Long        'Variable to hold the color of the pixel currently being worked with
    Dim i As Long               'Variable used as a counter in the for loops
    Dim picWidth As Long        'Variable to hold the width of the drawing area in pixels
    Dim picHeight As Long       'Variable to hold the height of the drawing area in pixels

    hPic = picDrawArea.hDc          'Get the Device Context handle and store it in a local variable for quick access
    With picDrawArea
        picWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbPixels)   'Get the width of the drawing area in pixels
        picHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbPixels) 'Get the height of the drawing area in pixels
    End With
    If Depressed Then   'If button is depressed then
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountLighten)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountDarken)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountLighten)
        Next
    Else        'Button should be drawn up
        'Start a loop through the Horizontal points from Left to Right on the outside of the border
        For i = (DefaultOffset) To (picWidth - DefaultOffset - 1)
            CurColor = GetPixel(hPic, i, DefaultOffset)   'Get Color of the current pixel on the Left Side
            'Adjust the color and write the pixel back to the Top
            SetPixelV hPic, i, DefaultOffset, AdjustColor(CurColor, 2 * AmountLighten)
            CurColor = GetPixel(hPic, i, picHeight - DefaultOffset - 1)   'Get Color of the current pixel on the Right Side
            'Adjust the color and write the pixel back to the Bottom
            SetPixelV hPic, i, picHeight - DefaultOffset - 1, AdjustColor(CurColor, 2 * AmountDarken)
        Next
        'Start a loop through the Vertical poitns from Top to Bottom on the outside of the border
        For i = (DefaultOffset + 1) To (picHeight - DefaultOffset - 2)
            CurColor = GetPixel(hPic, DefaultOffset, i)  'Get Color of the current pixel at the top
            'Adjust the color and write the pixel back to the Left Side
            SetPixelV hPic, DefaultOffset, i, AdjustColor(CurColor, 2 * AmountLighten)
            CurColor = GetPixel(hPic, picWidth - DefaultOffset - 1, i)    'Get Color of the current pixel on the Bottom
            'Adjust the color and write the pixel back to the Right side
            SetPixelV hPic, picWidth - DefaultOffset - 1, i, AdjustColor(CurColor, 2 * AmountDarken)
        Next
    End If
End Sub

Private Sub BorderFlatOld(ByVal DefaultOffset As Long, _
                                Optional ByVal Depressed As Boolean = False)
    If Depressed Then   'If button is depressed then
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest, B    'Draw a box on the border of the control that will become the Bottom and Right border lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest     'Draw the Top and Left Border Lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode)), BorderDarkest
    Else        'Button should be drawn up
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderDarkest, B     'Draw a box on the border of the control that will become the Bottom and Right border lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleHeight - ScaleY(1 + DefaultOffset, vbPixels, ScaleMode)), BorderLightest    'Draw the Top and Left Border Lines
        picDrawArea.Line (ScaleX(DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode))-(ScaleWidth - ScaleX(1 + DefaultOffset, vbPixels, ScaleMode), ScaleY(DefaultOffset, vbPixels, ScaleMode)), BorderLightest
    End If
End Sub

Private Sub CheckParent()
    On Error GoTo NoParent      'Start error handling in case the parent object doesn't support what we need
    If Not (UserControl.Parent.ScaleMode = vbUser) Then 'If the scalemode is not User defined then
        bParentAvailable = True 'Indicate we have parent available and it's not in user mode
    End If
    Exit Sub    'Finish this sub procedure

NoParent:
    bParentAvailable = False    'We don't have a parent with the scalemode property so indicate such to the control.
End Sub

Private Sub CopyGradient()
    If (m_Style = gpsGradient) Then     'If it is Gradient mode then
        PaintTransparentStdPic picDrawArea.hDc, 0, 0, ScaleX(ScaleWidth, ScaleMode, vbPixels), ScaleY(ScaleHeight, ScaleMode, vbPixels), GradPic, 0, 0, picDrawArea.BackColor   'Copy the gradient from the picture to the drawing area
        picDrawArea.Refresh             'Refresh so that gradient will show up
    End If
End Sub

Private Sub CreateTransparencyMask()
    Dim ctl As Control          'Variable to hold a reference to a contained control
    Dim NonTransColor As Long   'Variable to hold a color that will not show up as transparent for the mask

    If (m_Style = gpsTransparent) And (Ambient.UserMode) Then       'If not in design mode and style is transparent then
        picDrawArea.Picture = picDrawArea.Image
        If AdjustColor(picDrawArea.BackColor, 0) = vbBlack Then     'If the background color of the drawing area is black then
            NonTransColor = vbWhite     'Use white as the non transparent color
        Else    'Background is not black
            NonTransColor = vbBlack     'Use black as the non transparent color
        End If
        UserControl.MaskColor = picDrawArea.BackColor   'Set the control's mask color to the background color of the drawing area
        On Error Resume Next        'If we encounter an error while going through the controls then ignore it and continue on
        For Each ctl In ContainedControls   'Loop through all of the contained controls
            'Create a filled box the size and position of the current contained control out of the nontransparent color (to make sure the user added controls will show up)
            picDrawArea.Line (ScaleX(ctl.Left, ScaleMode, vbTwips), ScaleY(ctl.Top, ScaleMode, vbTwips))-Step(ScaleX(ctl.Width, ScaleMode, vbTwips) - ScaleX(1, vbPixels, vbTwips), ScaleY(ctl.Height, ScaleMode, vbTwips) - ScaleY(1, vbPixels, vbTwips)), NonTransColor, BF
        Next
        UserControl.MaskPicture = picDrawArea.Image     'Set the mask image from what we created on the drawing area
        UserControl.BackStyle = 0   'Set control's backstyle to Transparent
        picDrawArea.Cls             'Clear any extra from the drawing area
        picDrawArea.Refresh         'Refresh the drawing area graphics so the correct image is displayed
    Else        'Control is either in design mode or in a style that is not transparent.
        UserControl.BackStyle = 1   'Set control's backstyle to Opague
    End If
End Sub

Private Sub DrawBorders()
    Select Case m_Appearance        'Select the correct Appearance Scheme
        Case gpaNone        'No Border Appearance
        Case gpaFlatRaised  'Flat Raised
            BorderFlat 0    'Use the Flat Border non-depressed (DefaultOffset is always 0 for a frame control because it never has focus)
        Case gpaFlatInset   'Flat Inset
            BorderFlat 0, True      'Use the Flat Border depressed (DefaultOffset is always 0 for a frame control because it never has focus)
        Case gpa3DRaised    '3D Raised
            Border3D 0      'Use the 3D Border non-depressed (DefaultOffset is always 0 for a frame control because it never has focus)
        Case gpa3DInset     '3D Inset
            Border3D 0, True        'Use the 3D Border non-depressed (DefaultOffset is always 0 for a frame control because it never has focus)
        Case gpaEtched      'Etched
            BorderEtched 0  'Use the Etched Border non-depressed (DefaultOffset is always 0 for a frame control because it never has focus)
        Case gpaBevelRaised 'Bevel Raised
            BorderBevel 0, m_BevelWidth         'Use the Beveled Border non-depressed (DefaultOffset is always 0 for a frame control because it never has focus)
        Case gpaBevelInset  'Bevel Inset
            BorderBevel 0, m_BevelWidth, True   'Use the Beveled Border non-depressed (DefaultOffset is always 0 for a frame control because it never has focus)
        Case Else           'Invalid appearance so use the default Etched
            BorderEtched 0  'Use the Etched Border non-depressed (DefaultOffset is always 0 for a frame control because it never has focus)
    End Select
End Sub

Private Sub DrawCaption(ByVal Caption As String, _
                        ByRef Target As Control, _
                        Optional ByVal Alignment As gpAlignment, _
                        Optional ByVal CaptionStyle As gpCaptionStyle = gpcStandard)
    'Primitives
    Dim Flags As Long               'To hold alignment flags for drawing the text
    Dim Line As String              'The current working line
    Dim OrigColor As Long           'Variable to hold the original text color
    Dim Height As Long              'Variable to hold the Height of the caption
    Dim Width As Long               'Variable to hold the Width of the caption
    Dim ret As Long                 'To hold the DrawText return value (Text Height)
    'Structures
    Dim rct As RectAPI              'Structure to hold the rectangle information about where to draw the Caption
    Dim rctAdjust As RectAPI        'Structure to hold the rectangle information after adjustments for effects.

    Line = Caption                  'Assign the caption to the internal Line variable for processing (if necessary)
    GetClientRect UserControl.hWnd, rct     'Get the rectangular area of the control
    ret = DrawText(Target.hDc, Line, Len(Line), rct, DT_CALCRECT)       'Have the API calculate the correct rect area
    Width = rct.Right - rct.Left    'Store the width of the caption rect
    Height = rct.Bottom - rct.Top   'Store the height of the caption rect
    Select Case Alignment           'Select the alignment to be used
        Case gpaLeftTop             'Top Left Corner
            rct.Top = AlignmentCushion  'At Top vertically (Everything needs to be off one pixel in case caption styles are used)
            rct.Left = AlignmentCushion 'Left side
            'Fill out rest of rect from width and height
            rct.Bottom = rct.Top + Height
            rct.Right = rct.Left + Width
        Case gpaLeftMiddle          'Left Centered
            rct.Top = CLng((ScaleY(ScaleHeight, ScaleMode, vbPixels) - Height) / 2)     'Center vertically
            rct.Left = AlignmentCushion 'Left side
            'Fill out rest of rect from width and height
            rct.Bottom = rct.Top + Height
            rct.Right = rct.Left + Width
        Case gpaLeftBottom          'Bottom Left Corner
            rct.Top = CLng((ScaleY(ScaleHeight, ScaleMode, vbPixels) - Height) - (AlignmentCushion + 1))     'At Bottom vertically
            rct.Left = AlignmentCushion 'Left side
            'Fill out rest of rect from width and height
            rct.Bottom = rct.Top + Height
            rct.Right = rct.Left + Width
        Case gpaRightTop            'Top Right Corner
            rct.Top = AlignmentCushion  'At Top vertically
            rct.Left = CLng((ScaleX(ScaleWidth, ScaleMode, vbPixels) - Width) - (AlignmentCushion + 1))      'Right side
            'Fill out rest of rect from width and height
            rct.Bottom = rct.Top + Height
            rct.Right = rct.Left + Width
        Case gpaRightMiddle         'Right Centered
            rct.Top = CLng((ScaleY(ScaleHeight, ScaleMode, vbPixels) - Height) / 2)     'Center vertically
            rct.Left = CLng((ScaleX(ScaleWidth, ScaleMode, vbPixels) - Width) - (AlignmentCushion + 1))      'Right side
            'Fill out rest of rect from width and height
            rct.Bottom = rct.Top + Height
            rct.Right = rct.Left + Width
        Case gpaRightBottom         'Bottom Right Corner
            rct.Top = CLng((ScaleY(ScaleHeight, ScaleMode, vbPixels) - Height) - (AlignmentCushion + 1))     'At Bottom vertically
            rct.Left = CLng((ScaleX(ScaleWidth, ScaleMode, vbPixels) - Width) - (AlignmentCushion + 1))      'Right side
            'Fill out rest of rect from width and height
            rct.Bottom = rct.Top + Height
            rct.Right = rct.Left + Width
        Case gpaCenterTop           'Top Centered
            rct.Top = AlignmentCushion  'At Top vertically
            rct.Left = CLng((ScaleX(ScaleWidth, ScaleMode, vbPixels) - Width) / 2)      'Center Horizontally
            'Fill out rest of rect from width and height
            rct.Bottom = rct.Top + Height
            rct.Right = rct.Left + Width
        Case gpaCenterMiddle        'Full Centered
            rct.Top = CLng((ScaleY(ScaleHeight, ScaleMode, vbPixels) - Height) / 2)     'Center vertically
            rct.Left = CLng((ScaleX(ScaleWidth, ScaleMode, vbPixels) - Width) / 2)      'Center Horizontally
            'Fill out rest of rect from width and height
            rct.Bottom = rct.Top + Height
            rct.Right = rct.Left + Width
        Case gpaCenterBottom        'Bottom Centered
            rct.Top = CLng((ScaleY(ScaleHeight, ScaleMode, vbPixels) - Height) - (AlignmentCushion + 1))     'At Bottom vertically
            rct.Left = CLng((ScaleX(ScaleWidth, ScaleMode, vbPixels) - Width) / 2)      'Center Horizontally
            'Fill out rest of rect from width and height
            rct.Bottom = rct.Top + Height
            rct.Right = rct.Left + Width
        Case Else                   'Invalid so do full centering
            rct.Top = CLng((ScaleY(ScaleHeight, ScaleMode, vbPixels) - Height) / 2)     'Center vertically
            rct.Left = CLng((ScaleX(ScaleWidth, ScaleMode, vbPixels) - Width) / 2)      'Center Horizontally
            'Fill out rest of rect from width and height
            rct.Bottom = rct.Top + Height
            rct.Right = rct.Left + Width
    End Select
    OrigColor = Target.ForeColor    'Store the original color so that it may be restored after processing is done
    Flags = DT_SINGLELINE           'Ensure that the Caption will remain on one line
    Select Case CaptionStyle        'Choose the caption style
        Case gpcRaisedLight         'Light Inset Effect
            rctAdjust = rct             'Assign all the rectangle information to rctAdjust
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top - 1         'Move rectangle up one pixel
                .Bottom = .Bottom - 1
                .Left = .Left - 1       'Move rectangle left one pixel
                .Right = .Right - 1
            End With
            Target.ForeColor = vbWhite  'Change text color to white for highlighting
            ret = DrawText(Target.hDc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            Target.ForeColor = OrigColor            'Restore original color for caption
        Case gpcRaisedHeavy          'Heavy Inset Effect
            rctAdjust = rct             'Assign all the rectangle information to rctAdjust
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top - 1         'Move rectangle up one pixel
                .Bottom = .Bottom - 1
                .Left = .Left - 1       'Move rectangle left one pixel
                .Right = .Right - 1
            End With
            Target.ForeColor = vbWhite  'Change text color to white for highlighting
            ret = DrawText(Target.hDc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top + 2         'Move rectangle down two pixels
                .Bottom = .Bottom + 2
                .Left = .Left + 2       'Move rectangle right two pixels
                .Right = .Right + 2
            End With
            Target.ForeColor = AdjustColor(OrigColor, 60)   'Adjust text color to a lighter shade
            ret = DrawText(Target.hDc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            Target.ForeColor = OrigColor            'Restore original color for caption
        Case gpcInsetLight         'Light Raised Effect
            rctAdjust = rct             'Assign all the rectangle information to rctAdjust
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top + 1         'Move rectangle down one pixel
                .Bottom = .Bottom + 1
                .Left = .Left + 1       'Move rectangle right one pixel
                .Right = .Right + 1
            End With
            Target.ForeColor = vbWhite  'Change text color to white for highlighting
            ret = DrawText(Target.hDc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            Target.ForeColor = OrigColor            'Restore original color for caption
        Case gpcInsetHeavy         'Heavy Raised Effect
            rctAdjust = rct             'Assign all the rectangle information to rctAdjust
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top - 1         'Move rectangle up one pixel
                .Bottom = .Bottom - 1
                .Left = .Left - 1       'Move rectangle left one pixel
                .Right = .Right - 1
            End With
            Target.ForeColor = AdjustColor(OrigColor, 60)   'Adjust text color to a lighter shade
            ret = DrawText(Target.hDc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top + 2         'Move rectangle down two pixels
                .Bottom = .Bottom + 2
                .Left = .Left + 2       'Move rectangle right two pixels
                .Right = .Right + 2
            End With
            Target.ForeColor = vbWhite  'Change text color to white for highlighting
            ret = DrawText(Target.hDc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            Target.ForeColor = OrigColor            'Restore original color for caption
        Case gpcDropShadow          'Drop Shadow Effect
            rctAdjust = rct             'Assign all the rectangle information to rctAdjust
            With rctAdjust              'Use rctAdjust as the default for this block
                .Top = .Top + 1         'Move rectangle down one pixel
                .Bottom = .Bottom + 1
                .Left = .Left + 1       'Move rectangle right one pixel
                .Right = .Right + 1
            End With
            Target.ForeColor = AdjustColor(OrigColor, 60)   'Adjust text color to a lighter shade
            ret = DrawText(Target.hDc, Line, Len(Line), rctAdjust, Flags)     'Draw a lighter shadow 1 pixel above and to the left of where the real caption will go
            Target.ForeColor = OrigColor            'Restore original color for caption
        Case Else                   'No Effect/Invalid Effect
            'Do nothing
    End Select
    ret = DrawText(Target.hDc, Line, Len(Line), rct, Flags)     'Actually Draw on the Caption
End Sub

Private Sub PaintAll()
    Set picDrawArea.Picture = Nothing   'Clear the picture property because it may have been used for the transparency
    picDrawArea.Cls 'Clear the Temporary graphic holder
    SetColors       'Set the colors just in case any color corruption occured on screen
    PaintGradient   'Paint the gradient into a seperate picture box
    CopyGradient    'Copy the gradient onto the control
    SetBackPicture  'Paint on the Picture if appropriate
    DrawBorders     'Draw the borders
    DrawCaption m_Caption, picDrawArea, m_Alignment, m_CaptionStyle
    CreateTransparencyMask  'Create a transparency mask for the panel and turn on transparency
    PaintToControl  'Paint the drawn control face onto the control
End Sub

Private Sub PaintFast()
    Set picDrawArea.Picture = Nothing   'Clear the picture property because it may have been used for the transparency
    picDrawArea.Cls 'Clear the Temporary graphic holder
    CopyGradient    'Copy the gradient onto the control
    SetBackPicture  'Paint on the Picture if appropriate
    DrawBorders     'Draw the borders
    DrawCaption m_Caption, picDrawArea, m_Alignment, m_CaptionStyle
    CreateTransparencyMask  'Create a transparency mask for the panel and turn on transparency
    PaintToControl  'Paint the drawn control face onto the control
End Sub

Private Sub PaintGradient()
    If (m_Style = gpsGradient) Then         'If it is in Gradient mode then
        'Draws the Gradient onto a PictureBox using the gradient class
        Grad.Color1 = m_GradientColor1      'Configure 1st color
        Grad.Color2 = m_GradientColor2      'Configure 2nd color
        Grad.Angle = m_GradientAngle        'Set the angle to draw at
        Grad.BlendMode = m_GradientBlendMode    'Set the blending mode
        Grad.Repetitions = m_GradientRepetitions    'Set the number of gradient repetitions
        Grad.GradientType = m_GradientType  'Set the gradient type to draw
        Grad.Draw picDrawArea               'Actually draws the gradient on the picture box
        picDrawArea.Picture = picDrawArea.Image     'Move the picture from a temporary position into the picture so that it can be copied and used as a standard picture
        Set GradPic = picDrawArea.Picture   'Copy the gradient into the picture variable
    Else
        picDrawArea.Picture = LoadPicture() 'Remove any trace gradient elements in the drawing area
    End If
End Sub

Private Sub PaintToControl()
    BitBlt UserControl.hDc, 0, 0, ScaleX(ScaleWidth, ScaleMode, vbPixels), ScaleY(ScaleHeight, ScaleMode, vbPixels), picDrawArea.hDc, 0, 0, vbSrcCopy       'Copy pre-painted control picture onto control
    UserControl.Refresh     'Do a refresh on the control so that what we just copied will show up
End Sub

'Provided with comments by Microsoft
Private Sub PaintStretchedStdPic(ByVal hdcDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal tWidth As Long, _
                                    ByVal tHeight As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
                                    ByVal sWidth As Long, _
                                    ByVal sHeight As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
    Dim hdcSrc As Long         'HDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RectAPI
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long

    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then GoTo PaintStretchedStdPic_InvalidParam

    Select Case picSource.Type
        Case vbPicTypeBitmap
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            'Select passed picture into an HDC
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.Handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            StretchBlt hdcDest, xDest, yDest, tWidth, tHeight, hdcSrc, xSrc, ySrc, sWidth, sHeight, vbSrcCopy
            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.Handle
            'Draw Transparent image
            StretchBlt hdcDest, xDest, yDest, tWidth, tHeight, hdcSrc, xSrc, ySrc, sWidth, sHeight, vbSrcCopy
            'Clean up
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case Else
            GoTo PaintStretchedStdPic_InvalidParam
    End Select
    Exit Sub

PaintStretchedStdPic_InvalidParam:
    Err.Raise giINVALID_PICTURE
    Exit Sub
End Sub

Private Sub PaintTransparentDC(ByVal hdcDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal hdcSrc As Long, _
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
    Dim hdcMask As Long        'HDC of the created mask image
    Dim hdcColor As Long       'HDC of the created color image
    Dim hbmMask As Long        'Bitmap handle to the mask image
    Dim hbmColor As Long       'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long         'Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long

    hdcScreen = GetDC(0&)
    'Validate palette
    If hPal = 0 Then
        'Create halftone palette
        hPal = CreateHalftonePalette(hdcScreen)
    End If
    OleTranslateColor clrMask, hPal, lMaskColor

    'Create a color bitmap to server as a copy of the destination
    'Do all work on this bitmap and then copy it back over the destination
    'when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    'Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcDest, xDest, yDest, vbSrcCopy

    'Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    'hdcSrc, because this will create a DIB section if the original bitmap
    'is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, Width, Height)
    'Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this first
    'and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome bitmap
    'does a nearest-color selection rather than painting based on the
    'backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hdcSrc contains a monochrome bitmap, we must set the destination
    'foreground/background colors according to those currently set in hdcSrc
    '(because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent color
    'from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.  All
    'other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color, and
    'the original colors everywhere else.  To do this, we first
    'paint the original onto the cover (which we already did), then we
    'AND the inverse of the mask onto that using the DSna ternary raster
    'operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    'Operation Codes", "Ternary Raster Operations", or search in MSDN
    'for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows transforms all white
    'bits (1) to the background color of the destination hdc.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, 0, 0, DSna
    'Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, 0, 0, vbSrcAnd
    'Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcPaint
    'Copy the screen buffer to the screen
    BitBlt hdcDest, xDest, yDest, Width, Height, hdcScnBuffer, 0, 0, vbSrcCopy
    'All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer

    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
    DeleteObject hPal
End Sub

Private Sub PaintTransparentStdPic(ByVal hdcDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xSrc As Long, _
                                    ByVal ySrc As Long, _
                                    ByVal clrMask As OLE_COLOR, _
                                    Optional ByVal hPal As Long = 0)
    Dim hdcSrc As Long         'HDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RectAPI
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long

    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then GoTo PaintTransparentStdPic_InvalidParam

    Select Case picSource.Type
        Case vbPicTypeBitmap
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            'Select passed picture into an HDC
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.Handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            PaintTransparentDC hdcDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, clrMask, hPal
            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                'Create halftone palette
                hPal = CreateHalftonePalette(hdcScreen)
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.Handle
            'Draw Transparent image
            PaintTransparentDC hdcDest, xDest, yDest, Width, Height, hdcSrc, 0, 0, lMaskColor, hPal
            'Clean up
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
            DeleteObject hPal
        Case Else
            GoTo PaintTransparentStdPic_InvalidParam
    End Select
    Exit Sub

PaintTransparentStdPic_InvalidParam:
    Err.Raise giINVALID_PICTURE
    Exit Sub
End Sub

Private Sub SetBackPicture()
    If (m_Style = gpsPicture) And Not (m_Picture Is Nothing) Then   'If it is in picture mode and the picture isn't nothing then
        'Paint the picture into the picturebox to prepare for overlay onto the control
        If m_AutoSize = gpasPictureToControl Then
            PaintStretchedStdPic picDrawArea.hDc, 0, 0, ScaleX(Width, vbTwips, vbPixels), ScaleY(Height, vbTwips, vbPixels), m_Picture, 0, 0, ScaleX(m_Picture.Width, vbHimetric, vbPixels), ScaleY(m_Picture.Height, vbHimetric, vbPixels), picDrawArea.BackColor
        Else
            PaintTransparentStdPic picDrawArea.hDc, 0, 0, ScaleX(m_Picture.Width, vbHimetric, vbPixels), ScaleY(m_Picture.Height, vbHimetric, vbPixels), m_Picture, 0, 0, picDrawArea.BackColor
        End If
        picDrawArea.Refresh     'Refresh the control so that the picture will show up
    End If
End Sub

Private Sub SetColors()
    Select Case m_Style         'Select the style so we know what colors to load
        Case gpsStandard        'Standard modes use system colors
            BorderDark = GetSysColor(vb3DShadow Xor &H80000000)         'Dark
            BorderDarkest = GetSysColor(vb3DDKShadow Xor &H80000000)    'Darkest
            BorderLight = GetSysColor(vb3DLight Xor &H80000000)         'Light
            BorderLightest = GetSysColor(vb3DHighlight Xor &H80000000)  'Lightest
        Case gpsGradient, gpsPicture, gpsTransparent    'New Modes use the color that was selected by the Developer
            BorderDark = AdjustColor(m_BorderColor, AmountDarken)           'Take the user selected color and Darken it
            BorderDarkest = AdjustColor(m_BorderColor, 2 * AmountDarken)    'Take the user selected color and double Darken it
            BorderLight = AdjustColor(m_BorderColor, AmountLighten)         'Take the user selected color and Lighten it
            BorderLightest = AdjustColor(m_BorderColor, 2 * AmountLighten)  'Take the user selected color and double Lighten it
        Case Else       'Invalid style use system colors
            BorderDark = GetSysColor(vb3DShadow Xor &H80000000)         'Dark
            BorderDarkest = GetSysColor(vb3DDKShadow Xor &H80000000)    'Darkest
            BorderLight = GetSysColor(vb3DLight Xor &H80000000)         'Light
            BorderLightest = GetSysColor(vb3DHighlight Xor &H80000000)  'Lightest
    End Select
End Sub
