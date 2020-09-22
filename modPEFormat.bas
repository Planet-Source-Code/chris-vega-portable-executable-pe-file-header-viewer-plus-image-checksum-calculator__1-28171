Attribute VB_Name = "modPEFormat"
Option Explicit

Private Declare Function CheckSumMappedFile _
                         Lib "Imagehlp" _
                         (ByVal BaseAddress As Long, _
                          ByVal FileLength As Long, _
                          ByVal HeaderSum As Long, _
                          ByVal CheckSum As Long) As Long
                          
Private Const e_key = 128

Public LastError As String
Public xPE_File As New clsFileMapping

Public ImageBase As Long

Public MZ_Header()
Public MZ_Res01()
Public MZ_OEM()
Public MZ_lfanew As Long
Public MZ_Header_d(13) As String
Public MZ_Res01_d(3) As String
Public MZ_OEM_d(11) As String
Public MZ_lfanew_d As String

Public PE_VA As Long

Public PE_Header(7)
Public PE_Header_d(7) As String
Public PE_Machines_v
Public PE_Machines_d(16) As String
Public PE_Characteristics_v
Public PE_Characteristics_d(14) As String

Public OH_VA As Long
Public Size_Of_OH As Integer

Public OH_Header()
Public OH_NT(20)

Public OH_Header_d(8)
Public OH_NT_d(20)

Public OH_Subsystems_d(8)
Public OH_Subsystems_v
Public OH_DLL_d(6)
Public OH_DLL_v

Public PE32 As Boolean
Public PEType As String

Public ImageCheckSum As Long
Public ImageFileSize As Long

Public Sub InitDescriptions()
    Dim i As Long
    
    ' MZ Headers Description
    For i = 0 To 13
        MZ_Header_d(i) = Trim(LoadResString(1000 + i))
    Next
    ' MZ Reserved Descriptions
    For i = 0 To 3
        MZ_Res01_d(i) = Trim(LoadResString(1200 + i))
    Next
    ' MZ OEM Descriptions
    For i = 0 To 11
        MZ_OEM_d(i) = Trim(LoadResString(1400 + i))
    Next
    ' Pointer to RVA Description
    MZ_lfanew_d = Trim(LoadResString(1666))
    
    ' PE Headers
    For i = 0 To 7
        PE_Header_d(i) = Trim(LoadResString(2000 + i))
    Next

    ' PE Headers (Machine Types)
    For i = 0 To 16
        PE_Machines_d(i) = Trim(LoadResString(2060 + i))
    Next
    
    ' PE Headers (Characteristics)
    For i = 0 To 14
        PE_Characteristics_d(i) = Trim(LoadResString(2100 + i))
    Next
    
    PE_Machines_v = Split(Trim(LoadResString(2198)), ",")
    PE_Characteristics_v = Split(Trim(LoadResString(2199)), ",")

    ' OH Headers
    For i = 0 To 8
        OH_Header_d(i) = Trim(LoadResString(3000 + i))
    Next

    ' OH Headers (NT Specific)
    For i = 0 To 20
        OH_NT_d(i) = Trim(LoadResString(3600 + i))
    Next

    ' PE Subsystems
    For i = 0 To 8
        OH_Subsystems_d(i) = Trim(LoadResString(3800 + i))
    Next
    
    ' PE DLL Characteristics
    For i = 0 To 6
        OH_DLL_d(i) = Trim(LoadResString(3900 + i))
    Next
    
    OH_Subsystems_v = Split(Trim(LoadResString(3998)), ",")
    OH_DLL_v = Split(Trim(LoadResString(3998)), ",")
End Sub

Public Function OpenPE(PE_Filename) As Boolean
    Dim i As Long, j As Long
    OpenPE = False
    LastError = ""

    With xPE_File
        If .OpenFile(Trim(PE_Filename)) Then
            If .MapFile Then
                If .OpenView Then
                    If Hex(.lodsw) <> "5A4D" Then
                        LastError = "Not a Valid Executable File"
                        OpenPE = False
                    Else
                        ImageBase = .GetFileEntryPoint
                        ImageFileSize = .GetFileSizeX
                        
                        LastError = ""
                        OpenPE = True
                        ' Here we got a valid Executable File

                        ' =========================================================
                        ' Extract MZ Headers
                        ' =========================================================
                        MZ_Header = .ReadStream(14, , DefineWord)   ' 14 Word Values
                        MZ_Res01 = .ReadStream(4, , DefineWord)     '  4 Word Values
                        MZ_OEM = .ReadStream(12, , DefineWord)      ' 12 Word Values
                        MZ_lfanew = .lodsd                ' RVA Pointer to PE Header
                        
                        PE_VA = ImageBase + MZ_lfanew    ' Align PE Address
                        
                        .SetFilePointer PE_VA, SetReplaceCurrent      ' Point to the
                                                                      '  Location
                        
                        If Hex(.lodsw) = "4550" Then
                        
                            ' Fill-up PE_Header Structure
                            For i = 0 To 7
                                Select Case i
                                    Case 0, 3, 4, 5
                                        PE_Header(i) = _
                                                .ReadData(DefineDoubleWord)
                                    Case Else
                                        PE_Header(i) = _
                                                .ReadData(DefineWord)
                                End Select
                            Next
                            
                            ' Now we got the PE Header and we are pointer directly to
                            ' Optional Header
                            
                            OH_VA = .GetFilePointer         ' VA to Optional Headers
                            Size_Of_OH = PE_Header(6)       ' Size of Optional Headers
                            
                            ' Determine PE Type
                            If Right(Hex(.lodsd), 4) = "010B" Then PE32 = True _
                                                              Else PE32 = False
                            PEType = "PE32"
                            If Not PE32 Then PEType = PEType & "+/PE64"

                            If PE32 Then j = 8 Else j = 7
                            
                            ' Fill-up OH_Header Structure
                            For i = 0 To j
                                ReDim Preserve OH_Header(i)
                                Select Case i
                                    Case 0
                                        OH_Header(i) = _
                                                .ReadData(DefineWord)
                                    Case 1, 2
                                        OH_Header(i) = _
                                                .ReadData(DefineByte)
                                    Case Else
                                        OH_Header(i) = _
                                                .ReadData(DefineDoubleWord)
                                End Select
                            Next

                            ' Fill-up OH_Header (NT-Specific) Structure
                            For i = 0 To 20
                                Select Case i
                                    Case 0, 15, 16, 17, 18
                                        If PE32 Then
                                            ' DWORD Image Base
                                            OH_NT(i) = _
                                                    .ReadData(DefineDoubleWord)
                                        Else
                                            ' QWord Image Base
                                            OH_NT(i) = _
                                                    .ReadData(DefineDoubleWord) & _
                                                    "," & _
                                                    .ReadData(DefineDoubleWord)
                                        End If
                                    Case 3, 4, 5, 6, 7, 8, 13, 14
                                        OH_NT(i) = _
                                                .ReadData(DefineWord)
                                    Case Else
                                        OH_NT(i) = _
                                                .ReadData(DefineDoubleWord)
                                End Select
                            Next
                            
                            ' Re-Calculate the Image Checksum
                            CheckSumMappedFile ImageBase, _
                                               ImageFileSize, _
                                               VarPtr(i), _
                                               VarPtr(ImageCheckSum)
                        Else
                            LastError = "Not a Valid PE Executable File Format"
                            OpenPE = False
                        End If
                    End If

                    .CloseView
                    .CloseMap
                    .CloseFile
                Else
                    LastError = "Failed to Open Views of File"
                    .CloseMap
                    .CloseFile
                End If
            Else
                LastError = "Failed to Map the File"
                .CloseFile
            End If
        Else
            LastError = "Failed to Open the File"
        End If
    End With
End Function

Public Function StringIt(xValue, SizeIn) As String
    Dim k As Long
    If SizeIn <= 8 Then
        StringIt = Hex(xValue)
        For k = Len(Hex(xValue)) To (SizeIn - 1)
            StringIt = "0" & StringIt
        Next
    Else
        Dim xVal
        xVal = Split(xValue, ",")
        StringIt = Hex(xVal(0)) & Hex(xVal(1))
        For k = Len(Hex(xValue)) To (SizeIn - 1)
            StringIt = "0" & StringIt
        Next
    End If
End Function

Public Function getSize(StrToGetSize)
    getSize = Len(StrToGetSize) \ 2
End Function

Public Function TwoH(StrHex)
    If Len(StrHex) <= 1 Then TwoH = "0" & _
        StrHex Else TwoH = StrHex
End Function

Public Function etxt(strx As String)
    Dim i As Long
    etxt = ""
    For i = 1 To Len(strx)
        cat etxt, Chr(Asc(Mid(strx, i, 1)) Xor e_key)
    Next
End Function

Public Sub cat(str1, str2)
    str1 = str1 & str2
End Sub
