VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PE File Format Viewer"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9675
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frContainer 
      BorderStyle     =   0  'None
      Height          =   5505
      Left            =   180
      TabIndex        =   2
      Top             =   495
      Width           =   9375
      Begin VB.TextBox txtFull 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5370
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   9330
      End
      Begin VB.TextBox txtOH 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5370
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   9330
      End
      Begin VB.TextBox txtPE 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5370
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   9330
      End
      Begin VB.TextBox txtMZ 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5370
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   9330
      End
      Begin VB.Shape shShadow 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   5370
         Left            =   0
         Top             =   45
         Width           =   9375
      End
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Im Don&e"
      Height          =   330
      Left            =   7560
      TabIndex        =   1
      Top             =   6120
      Width           =   2040
   End
   Begin ComctlLib.TabStrip tabx 
      Height          =   5955
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   10504
      TabWidthStyle   =   2
      TabFixedWidth   =   4145
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "MZ Headers"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "PE Headers"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Optional Headers"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Full View"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      MousePointer    =   1
      Enabled         =   0   'False
   End
   Begin MSComDlg.CommonDialog dlgLoadEXE 
      Left            =   90
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "*.exe"
      DialogTitle     =   "Select PE-Executable"
      Filter          =   "Executable Files (*.exe)|*.exe|Dynamic Link Library (*.dll)|*.dll|All Files (*.*)|*.*"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open PE File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnudash000 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' This is a DEMO Application on Using my FileMapping VB Class Module
'   to Simulate the use of File Mapping APIs
'
'       Used in Visual C++ and Win32 Assembly, Why in Visual Basic, Why Not?
'
'
'   This code checks for a valid PE Executable Format
'   And display the Headers Visually.
'
'   Added some Checkings on Machine Types, Characteristics and Browser to
'   Section Headers of the PE File Format.
'
'   To understand this Application, you need to consult your nearest
'   PE Documentation.
'
'   Win32 Assembly Codes are included in Comments are 100% working on
'   TASM32 Compiler
'
'
'
'   Created by: Chris Vega [gwapo@models.com]
'               http://trider.8m.com
'

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    InitDescriptions
    mnuHelp_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Copyright 2001 by Chris Vega [gwapo@models.com]", _
           vbInformation, "Chris Vega [gwapo@models.com]"
End Sub

Private Sub mnuexit_Click()
    Unload Me
End Sub

Private Sub mnuHelp_Click()
    MsgBox "Click ""File -> Open PE File"" or Press Ctrl+O " _
           & vbCrLf & "to Start the Viewing a PE File Headers", _
           vbInformation, _
           "Chris Vega [gwapo@models.com]"
End Sub

Private Sub mnuopen_Click()
    dlgLoadEXE.ShowOpen
    If Not (Err) Then
        If OpenPE(dlgLoadEXE.FileName) Then
            Caption = "PE File Format Viewer [" & dlgLoadEXE.FileName & "]"
            
            Build_MZ_Headers
            Build_PE_Headers
            Build_OH_Headers
            Build_Full_View
            
            tabx.Enabled = True
            tabx_Click
        Else
            Caption = "PE File Format Viewer"
            tabx.Enabled = False
            tabx_Click
            MsgBox LastError, _
                   vbExclamation, _
                   "Chris Vega [gwapo@models.com]"
        End If
    Else
        tabx.Enabled = False
        tabx.Tabs(1).Selected = False
    End If
End Sub

Private Sub tabx_Click()
    If tabx.Enabled Then
        Select Case tabx.SelectedItem.Index
            Case 1
                txtMZ.Visible = True
                txtPE.Visible = False
                txtOH.Visible = False
                txtFull.Visible = False
            Case 2
                txtMZ.Visible = False
                txtPE.Visible = True
                txtOH.Visible = False
                txtFull.Visible = False
            Case 3
                txtMZ.Visible = False
                txtPE.Visible = False
                txtOH.Visible = True
                txtFull.Visible = False
            Case 4
                txtMZ.Visible = False
                txtPE.Visible = False
                txtOH.Visible = False
                txtFull.Visible = True
        End Select
    Else
        txtMZ.Visible = False
    End If
End Sub

Private Sub Build_MZ_Headers()
    Dim xOffset As Long
    Dim i As Long
    
    xOffset = 0
    txtMZ = etxt(LoadResString(100))
            
    For i = 0 To UBound(MZ_Header)
        txtMZ = txtMZ & _
                LCase(TwoH(Hex(xOffset))) & "h" & _
                vbTab & StringIt(MZ_Header(i), 4) & _
                vbTab & vbTab & MZ_Header_d(i) & _
                vbCrLf
        xOffset = xOffset + 2
    Next

    txtMZ = txtMZ & vbCrLf & etxt(LoadResString(101)) & vbCrLf

    For i = 0 To UBound(MZ_Res01)
        txtMZ = txtMZ & _
                LCase(TwoH(Hex(xOffset))) & "h" & _
                vbTab & StringIt(MZ_Res01(i), 4) & _
                vbTab & vbTab & MZ_Res01_d(i) & _
                vbCrLf
        xOffset = xOffset + 2
    Next

    txtMZ = txtMZ & vbCrLf & etxt(LoadResString(102)) & vbCrLf

    For i = 0 To UBound(MZ_OEM)
        txtMZ = txtMZ & _
                LCase(TwoH(Hex(xOffset))) & "h" & _
                vbTab & StringIt(MZ_OEM(i), 4) & _
                vbTab & vbTab & MZ_OEM_d(i) & _
                vbCrLf
        xOffset = xOffset + 2
    Next

    txtMZ = txtMZ & vbCrLf & etxt(LoadResString(103)) & vbCrLf
    txtMZ = txtMZ & _
                LCase(TwoH(Hex(xOffset))) & "h" & _
                vbTab & StringIt(MZ_lfanew, 8) & _
                vbTab & MZ_lfanew_d & _
                vbCrLf
End Sub

Private Sub Build_PE_Headers()
    Dim xOffset As Long
    Dim i As Long, sizeX, xTab
    
    xOffset = 0
    txtPE = Replace(etxt(LoadResString(104)), _
            "%%PE_Head%%", StringIt(MZ_lfanew, 8) & "h")

    For i = 0 To UBound(PE_Header)
        If i = 1 Or i = 2 Or i = 6 Or i = 7 Then
            sizeX = 4
            xTab = vbTab
        Else
            sizeX = 8
            xTab = ""
        End If
        txtPE = txtPE & _
                LCase(TwoH(Hex(xOffset))) & "h" & _
                vbTab & StringIt(PE_Header(i), sizeX) & _
                vbTab & xTab & PE_Header_d(i) & _
                vbCrLf
        If i = 1 Then printMachineType
        xOffset = xOffset + (sizeX \ 2)
    Next
    
    printCharacteristics
End Sub

Private Sub Build_OH_Headers()
    Dim xOffset As Long
    Dim i As Long, sizeX, xTab
    Dim size_of_NT As Long
    
    xOffset = 0
    txtOH = Replace(etxt(LoadResString(105)), _
            "%%OH Header%%", StringIt(OH_VA - ImageBase, 8) & "h")

    For i = 0 To UBound(OH_Header)
        Select Case i
            Case 0
                sizeX = 4
                xTab = vbTab
            Case 1, 2
                sizeX = 2
                xTab = vbTab
            Case Else
                sizeX = 8
                xTab = ""
        End Select
        
        txtOH = txtOH & _
                LCase(TwoH(Hex(xOffset))) & "h" & vbTab & _
                LCase(TwoH(Hex(xOffset + 24))) & "h" & vbTab & _
                vbTab & StringIt(OH_Header(i), sizeX) & _
                vbTab & xTab & OH_Header_d(i) & _
                vbCrLf

        xOffset = xOffset + (sizeX \ 2)
    Next

    If PE32 Then size_of_NT = 68 Else size_of_NT = 88

    txtOH = txtOH & vbCrLf & Replace(Replace(etxt(LoadResString(106)), _
            "%%PEType%%", PEType), "%%OHSize%%", size_of_NT) & vbCrLf

    For i = 0 To UBound(OH_NT)
        Select Case i
            Case 0, 15, 16, 17, 18
                If PE32 Then
                    sizeX = 8
                Else
                    sizeX = 16
                End If
                xTab = ""
            Case 3, 4, 5, 6, 7, 8, 13, 14
                sizeX = 4
                xTab = vbTab
            Case Else
                sizeX = 8
                xTab = ""
        End Select
        
        txtOH = txtOH & _
                LCase(TwoH(Hex(xOffset))) & "h" & vbTab & _
                LCase(TwoH(Hex(xOffset + 24))) & "h" & vbTab & _
                vbTab & StringIt(OH_NT(i), sizeX) & _
                vbTab & xTab & OH_NT_d(i) & _
                vbCrLf
        
        If i = 12 Then printCheckSum Else _
        If i = 13 Then printSubSytem Else _
        If i = 14 Then printDLLCharacteristics

        xOffset = xOffset + (sizeX \ 2)
    Next

End Sub

Private Sub Build_Full_View()
    Dim ChrisVega As String
    
    ChrisVega = etxt(LoadResString(666))
    
    txtFull = Replace(txtMZ, ChrisVega, "") & vbCrLf & vbCrLf & _
              Replace(txtPE, ChrisVega, "") & vbCrLf & vbCrLf & _
              Replace(txtOH, ChrisVega, "")
    
    ChrisVega = Replace( _
                Replace( _
                Replace( _
                Replace( _
                        etxt(LoadResString(999)), _
                        "%%Filename%%", _
                        dlgLoadEXE.FileName), _
                        "%%Filesize%%", _
                        ImageFileSize), _
                        "%%PEType%%", _
                        PEType), _
                        "%%Checksum%%", _
                        StringIt(ImageCheckSum, 8))
    
    txtFull = ChrisVega & txtFull
End Sub

Private Sub printMachineType()
    Dim macX, i As Long
    
    macX = Hex(PE_Header(1))
    
    For i = 0 To UBound(PE_Machines_d)
        If macX = Hex(PE_Machines_v(i)) Then
            txtPE = txtPE & _
                    vbTab & _
                    vbTab & vbTab & "(" & _
                    PE_Machines_d(i) & ")" & _
                    vbCrLf
            i = UBound(PE_Machines_d) + 1   ' Exit For
        End If
    Next
End Sub

Private Sub printCharacteristics()
    Dim macX, i As Long
    
    macX = PE_Header(7)
    
    For i = 0 To UBound(PE_Characteristics_d)
        If macX And PE_Characteristics_v(i) Then
            txtPE = txtPE & _
                    vbTab & _
                    vbTab & vbTab & "(" & _
                    PE_Characteristics_d(i) & ")" & _
                    vbCrLf
        End If
    Next
End Sub

Private Sub printCheckSum()
    txtOH = txtOH & _
            vbTab & _
            vbTab & vbTab & vbTab & vbTab & "(" & _
            StringIt(ImageCheckSum, 8) & ")" & _
            vbCrLf
End Sub

Private Sub printSubSytem()
    Dim macX, i As Long
    
    macX = Hex(OH_NT(13))
    
    For i = 0 To UBound(OH_Subsystems_d)
        If macX = Hex(OH_Subsystems_v(i)) Then
            txtOH = txtOH & _
                    vbTab & vbTab & vbTab & _
                    vbTab & vbTab & "(" & _
                    OH_Subsystems_d(i) & ")" & _
                    vbCrLf
        End If
    Next
End Sub

Private Sub printDLLCharacteristics()
    Dim macX, i As Long
    
    macX = OH_NT(14)
    
    For i = 0 To UBound(OH_DLL_d)
        If macX And OH_DLL_v(i) Then
            txtOH = txtOH & _
                    vbTab & vbTab & vbTab & _
                    vbTab & vbTab & "(" & _
                    OH_DLL_d(i) & ")" & _
                    vbCrLf
        End If
    Next
End Sub
