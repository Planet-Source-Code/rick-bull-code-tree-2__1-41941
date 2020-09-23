VERSION 5.00
Begin VB.UserControl Seperator 
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   PropertyPages   =   "Seperator.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   960
   ToolboxBitmap   =   "Seperator.ctx":0014
End
Attribute VB_Name = "Seperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Public Enum OrientationConstants
    [Horizontal]
    [Vertical]
End Enum
'Default Property Values:
Const m_def_Color1 = vb3DShadow
Const m_def_Color2 = vb3DHighlight
Const m_def_AutoSize = True
Const m_def_Orientation = Horizontal
'Property Variables:
Dim m_Color1 As OLE_COLOR
Dim m_Color2 As OLE_COLOR
Dim m_AutoSize As Boolean
Dim m_Orientation As OrientationConstants

Private Sub DrawControl()
    On Local Error Resume Next
    With UserControl
        .Cls
        If m_Orientation = [Vertical] Then
            UserControl.Line (0, 0)-(.ScaleWidth, .ScaleHeight), _
                m_Color2, BF
            UserControl.Line (0, 0)-((.ScaleWidth \ 2) - TwipsX, .ScaleHeight - TwipsY(2)), _
                m_Color1, BF
        Else
            UserControl.Line (0, 0)-(.ScaleWidth, .ScaleHeight), _
                m_Color2, BF
            UserControl.Line (0, 0)-(.ScaleWidth - TwipsX(2), (.ScaleHeight \ 2) - TwipsY), _
                m_Color1, BF
        End If
        If .AutoRedraw Then .Refresh
    End With
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    If m_AutoSize Then
        With UserControl
            If m_Orientation = Horizontal Then
                .Height = TwipsY(2)
                If .Width < TwipsX(5) Then .Width = TwipsX(20)
            ElseIf m_Orientation = Vertical Then
                .Width = TwipsX(2)
                If .Height < TwipsY(5) Then .Height = TwipsY(20)
            End If
        End With
    End If
    Call DrawControl
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    m_Color1 = PropBag.ReadProperty("Color1", m_def_Color1)
    m_Color2 = PropBag.ReadProperty("Color2", m_def_Color2)
    Call DrawControl
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("Color1", m_Color1, m_def_Color1)
    Call PropBag.WriteProperty("Color2", m_Color2, m_def_Color2)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,Horizontal
Public Property Get Orientation() As OrientationConstants
Attribute Orientation.VB_Description = "Returns/sets the orientation for the control (horizontal = lines from left to right, vertical = lines from top to bottom)."
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As OrientationConstants)
    m_Orientation = New_Orientation
    PropertyChanged "Orientation"
    Call DrawControl
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Orientation = m_def_Orientation
    m_AutoSize = m_def_AutoSize
    m_Color1 = m_def_Color1
    m_Color2 = m_def_Color2
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    Call UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vb3DShadow
Public Property Get Color1() As OLE_COLOR
Attribute Color1.VB_Description = "Returns/sets the first colour to be used for the control."
    Color1 = m_Color1
End Property

Public Property Let Color1(ByVal New_Color1 As OLE_COLOR)
    m_Color1 = New_Color1
    PropertyChanged "Color1"
    Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vb3DHighlight
Public Property Get Color2() As OLE_COLOR
Attribute Color2.VB_Description = "Returns/sets the second colour to be used for the control."
    Color2 = m_Color2
End Property

Public Property Let Color2(ByVal New_Color2 As OLE_COLOR)
    m_Color2 = New_Color2
    PropertyChanged "Color2"
    Call DrawControl
End Property

Public Sub About()
Attribute About.VB_Description = "Shows the about dialog."
Attribute About.VB_UserMemId = -552
    On Local Error Resume Next
    Call MsgBox("Written by Rick Bull Sat, 29 June 2002 16:07", _
        vbOKOnly Or vbInformation, "About")
End Sub
