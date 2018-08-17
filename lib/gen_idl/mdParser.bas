Attribute VB_Name = "mdParser"
' Auto-generated on 17.8.2018 13:58:39
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

Private Const LOCALE_USER_DEFAULT           As Long = &H400
Private Const NORM_IGNORECASE               As Long = 1
Private Const CSTR_EQUAL                    As Long = 2

Private Declare Function CompareStringW Lib "kernel32" (ByVal Locale As Long, ByVal dwCmpFlags As Long, lpString1 As Any, ByVal cchCount1 As Long, lpString2 As Any, ByVal cchCount2 As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const LNG_MAXINT            As Long = 2 ^ 31 - 1

'= generated enum ========================================================

Private Enum UcsParserActionsEnum
    ucsAct_3_StmtList
    ucsAct_2_StmtList
    ucsAct_1_StmtList
    ucsAct_1_TypedefDecl
    ucsAct_1_TypedefCallback
    ucsAct_3_EnumDecl
    ucsAct_2_EnumDecl
    ucsAct_1_EnumDecl
    ucsAct_5_StructDecl
    ucsAct_4_StructDecl
    ucsAct_3_StructDecl
    ucsAct_2_StructDecl
    ucsAct_1_StructDecl
    ucsAct_1_FunDecl
    ucsAct_1_SkipStmt
    ucsAct_1_Type
    ucsAct_1_ID
    ucsAct_1_TypeUnlimited
    ucsAct_3_ParamList
    ucsAct_2_ParamList
    ucsAct_1_ParamList
    ucsAct_4_EnumValue
    ucsAct_3_EnumValue
    ucsAct_2_EnumValue
    ucsAct_1_EnumValue
    ucsAct_1_EMPTY
    ucsAct_3_IDList
    ucsAct_2_IDList
    ucsAct_1_IDList
    ucsAct_3_StuctMemList
    ucsAct_2_StuctMemList
    ucsAct_1_StuctMemList
    ucsAct_3_ArraySuffixList
    ucsAct_2_ArraySuffixList
    ucsAct_1_ArraySuffixList
    ucsAct_1_Param
    ucsAct_1_ArraySuffix
    ucsAct_1_EnumValueToken
    ucsActVarAlloc = -1
    ucsActVarSet = -2
    ucsActResultClear = -3
    ucsActResultSet = -4
End Enum

Private Type UcsParserThunkType
    Action              As Long
    CaptureBegin        As Long
    CaptureEnd          As Long
End Type

Private Type UcsParserType
    Contents            As String
    BufData()           As Integer
    BufPos              As Long
    BufSize             As Long
    ThunkData()         As UcsParserThunkType
    ThunkPos            As Long
    CaptureBegin        As Long
    CaptureEnd          As Long
    LastExpected        As String
    LastError           As String
    LastBufPos          As Long
    UserData            As Variant
    VarResult           As Variant
    VarStack()          As Variant
    VarPos              As Long
End Type

Private ctx                     As UcsParserType

'=========================================================================
' Properties
'=========================================================================

Property Get VbPegLastError() As String
    VbPegLastError = ctx.LastError
End Property

Property Get VbPegLastOffset() As Long
    VbPegLastOffset = ctx.LastBufPos + 1
End Property

Property Get VbPegParserVersion() As String
    VbPegParserVersion = "17.8.2018 13:58:39"
End Property

Property Get VbPegContents(Optional ByVal lOffset As Long = 1, Optional ByVal lSize As Long = LNG_MAXINT) As String
    VbPegContents = Mid$(ctx.Contents, lOffset, lSize)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function VbPegMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant, Optional Result As Variant) As Long
    If VbPegBeginMatch(sSubject, StartPos, UserData) Then
        If VbPegParseStmtList() Then
            VbPegMatch = VbPegEndMatch(Result)
        Else
            With ctx
                If LenB(.LastError) = 0 Then
                    If LenB(.LastExpected) = 0 Then
                        .LastError = "Fail"
                    Else
                        .LastError = "Expected " & Join(Split(Mid$(.LastExpected, 2, Len(.LastExpected) - 2), vbNullChar), " or ")
                    End If
                End If
            End With
        End If
    End If
End Function

Public Function VbPegBeginMatch(sSubject As String, Optional ByVal StartPos As Long, Optional UserData As Variant) As Boolean
    With ctx
        .LastBufPos = 0
        If LenB(sSubject) = 0 Then
            .LastError = "Cannot match empty input"
            Exit Function
        End If
        .Contents = sSubject
        ReDim .BufData(0 To Len(sSubject) + 3) As Integer
        Call CopyMemory(.BufData(0), ByVal StrPtr(sSubject), LenB(sSubject))
        .BufPos = StartPos
        .BufSize = Len(sSubject)
        .BufData(.BufSize) = -1 '-- EOF anchor
        ReDim .ThunkData(0 To 4) As UcsParserThunkType
        .ThunkPos = 0
        .CaptureBegin = 0
        .CaptureEnd = 0
        If IsObject(UserData) Then
            Set .UserData = UserData
        Else
            .UserData = UserData
        End If
    End With
    VbPegBeginMatch = True
End Function

Public Function VbPegEndMatch(Optional Result As Variant) As Long
    Dim lIdx            As Long

    With ctx
        ReDim .VarStack(0 To 1024) As Variant
        For lIdx = 0 To .ThunkPos - 1
            Select Case .ThunkData(lIdx).Action
            Case ucsActVarAlloc
                .VarPos = .VarPos + .ThunkData(lIdx).CaptureBegin
            Case ucsActVarSet
                If IsObject(.VarResult) Then
                    Set .VarStack(.VarPos - .ThunkData(lIdx).CaptureBegin) = .VarResult
                Else
                    .VarStack(.VarPos - .ThunkData(lIdx).CaptureBegin) = .VarResult
                End If
            Case ucsActResultClear
                .VarResult = Empty
            Case ucsActResultSet
                With .ThunkData(lIdx)
                    ctx.VarResult = Mid$(ctx.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)
                End With
            Case Else
                With .ThunkData(lIdx)
                    pvImplAction .Action, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin
                End With
            End Select
        Next
        If IsObject(.VarResult) Then
            Set Result = .VarResult
        Else
            Result = .VarResult
        End If
        VbPegEndMatch = .BufPos + 1
        .Contents = vbNullString
        Erase .BufData
        .BufPos = 0
        .BufSize = 0
        Erase .ThunkData
        .ThunkPos = 0
        .CaptureBegin = 0
        .CaptureEnd = 0
    End With
End Function

Private Sub pvPushThunk(ByVal eAction As UcsParserActionsEnum, Optional ByVal lBegin As Long, Optional ByVal lEnd As Long)
    With ctx
        If UBound(.ThunkData) < .ThunkPos Then
            ReDim Preserve .ThunkData(0 To 2 * UBound(.ThunkData)) As UcsParserThunkType
        End If
        With .ThunkData(.ThunkPos)
            .Action = eAction
            .CaptureBegin = lBegin
            .CaptureEnd = lEnd
        End With
        .ThunkPos = .ThunkPos + 1
    End With
End Sub

Private Function pvMatchString(sText As String, Optional ByVal CmpFlags As Long) As Boolean
    With ctx
        If .BufPos + Len(sText) <= .BufSize Then
            pvMatchString = CompareStringW(LOCALE_USER_DEFAULT, CmpFlags, ByVal StrPtr(sText), Len(sText), .BufData(.BufPos), Len(sText)) = CSTR_EQUAL
        End If
    End With
End Function

Private Sub pvSetAdvance()
    With ctx
        If .BufPos > .LastBufPos Then
            .LastExpected = vbNullString
            .LastError = vbNullString
            .LastBufPos = .BufPos
        End If
    End With
End Sub

'= generated functions ===================================================

Public Function VbPegParseStmtList() As Boolean
    Dim p22 As Long
    Dim q22 As Long
    Dim p12 As Long
    Dim q12 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        pvPushThunk ucsActVarSet, 1
        pvPushThunk ucsAct_1_StmtList, .CaptureBegin, .CaptureEnd
        Do
            p22 = .BufPos
            q22 = .ThunkPos
            pvPushThunk ucsActResultClear
            p12 = .BufPos
            q12 = .ThunkPos
            If VbPegParseTypedefDecl() Then
                pvPushThunk ucsActVarSet, 2
                GoTo L1
            Else
                .BufPos = p12
                .ThunkPos = q12
            End If
            If VbPegParseTypedefCallback() Then
                pvPushThunk ucsActVarSet, 2
                GoTo L1
            Else
                .BufPos = p12
                .ThunkPos = q12
            End If
            If VbPegParseEnumDecl() Then
                pvPushThunk ucsActVarSet, 2
                GoTo L1
            Else
                .BufPos = p12
                .ThunkPos = q12
            End If
            If VbPegParseStructDecl() Then
                pvPushThunk ucsActVarSet, 2
                GoTo L1
            Else
                .BufPos = p12
                .ThunkPos = q12
            End If
            If VbPegParseFunDecl() Then
                pvPushThunk ucsActVarSet, 2
                GoTo L1
            Else
                .BufPos = p12
                .ThunkPos = q12
            End If
            If VbPegParseSkipStmt() Then
                pvPushThunk ucsActVarSet, 2
                GoTo L1
            Else
                .BufPos = p12
                .ThunkPos = q12
            End If
            .BufPos = p22
            .ThunkPos = q22
            Exit Do
L1:
            pvPushThunk ucsAct_2_StmtList, .CaptureBegin, .CaptureEnd
        Loop
        If ParseEOL() Then
            pvPushThunk ucsAct_3_StmtList, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            VbPegParseStmtList = True
        End If
    End With
End Function

Public Function VbPegParseTypedefDecl() As Boolean
    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If ParseTYPEDEF() Then
            pvPushThunk ucsActResultClear
            If VbPegParseType() Then
                pvPushThunk ucsActVarSet, 1
                pvPushThunk ucsActResultClear
                If ParseID() Then
                    pvPushThunk ucsActVarSet, 2
                    If ParseSEMI() Then
                        pvPushThunk ucsAct_1_TypedefDecl, .CaptureBegin, .CaptureEnd
                        pvPushThunk ucsActVarAlloc, -2
                        VbPegParseTypedefDecl = True
                    End If
                End If
            End If
        End If
    End With
End Function

Public Function VbPegParseTypedefCallback() As Boolean
    Dim p46 As Long
    Dim p69 As Long
    Dim q69 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        If ParseTYPEDEF() Then
            p46 = .BufPos
            If Not (VbPegParseLinkage()) Then
                .BufPos = p46
            End If
            pvPushThunk ucsActResultClear
            If VbPegParseTypeUnlimited() Then
                pvPushThunk ucsActVarSet, 1
                If ParseLPAREN() Then
                    If ParseCC_STDCALL() Then
                        If ParseSTAR() Then
                            pvPushThunk ucsActResultClear
                            If ParseID() Then
                                pvPushThunk ucsActVarSet, 2
                                If ParseRPAREN() Then
                                    If ParseLPAREN() Then
                                        pvPushThunk ucsActResultClear
                                        p69 = .BufPos
                                        q69 = .ThunkPos
                                        If Not (VbPegParseParamList()) Then
                                            .BufPos = p69
                                            .ThunkPos = q69
                                        End If
                                        pvPushThunk ucsActVarSet, 3
                                        If ParseRPAREN() Then
                                            If ParseSEMI() Then
                                                pvPushThunk ucsAct_1_TypedefCallback, .CaptureBegin, .CaptureEnd
                                                pvPushThunk ucsActVarAlloc, -3
                                                VbPegParseTypedefCallback = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Public Function VbPegParseEnumDecl() As Boolean
    Dim p76 As Long
    Dim p82 As Long
    Dim q82 As Long
    Dim i111 As Long
    Dim p96 As Long
    Dim q96 As Long
    Dim p106 As Long
    Dim q106 As Long
    Dim p108 As Long
    Dim q108 As Long
    Dim p116 As Long
    Dim q116 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 5
        p76 = .BufPos
        If Not (ParseTYPEDEF()) Then
            .BufPos = p76
        End If
        If ParseENUM() Then
            pvPushThunk ucsActResultClear
            p82 = .BufPos
            q82 = .ThunkPos
            If Not (ParseID()) Then
                .BufPos = p82
                .ThunkPos = q82
            End If
            pvPushThunk ucsActVarSet, 1
            If ParseLBRACE() Then
                pvPushThunk ucsActResultClear
                pvPushThunk ucsActVarSet, 2
                pvPushThunk ucsAct_1_EnumDecl, .CaptureBegin, .CaptureEnd
                For i111 = 0 To LNG_MAXINT
                    p96 = .BufPos
                    q96 = .ThunkPos
                    pvPushThunk ucsActResultClear
                    If ParseID() Then
                        pvPushThunk ucsActVarSet, 3
                    Else
                        .BufPos = p96
                        .ThunkPos = q96
                        Exit For
                    End If
                    p106 = .BufPos
                    q106 = .ThunkPos
                    pvPushThunk ucsActResultClear
                    If VbPegParseEnumValue() Then
                        pvPushThunk ucsActVarSet, 4
                    Else
                        .BufPos = p106
                        .ThunkPos = q106
                        pvPushThunk ucsActResultClear
                        Call ParseEMPTY
                        pvPushThunk ucsActVarSet, 4
                    End If
                    p108 = .BufPos
                    q108 = .ThunkPos
                    If Not (ParseCOMMA()) Then
                        .BufPos = p108
                        .ThunkPos = q108
                    End If
                    pvPushThunk ucsAct_2_EnumDecl, .CaptureBegin, .CaptureEnd
                Next
                If i111 <> 0 Then
                    If ParseRBRACE() Then
                        pvPushThunk ucsActResultClear
                        p116 = .BufPos
                        q116 = .ThunkPos
                        If Not (VbPegParseIDList()) Then
                            .BufPos = p116
                            .ThunkPos = q116
                        End If
                        pvPushThunk ucsActVarSet, 5
                        If ParseSEMI() Then
                            pvPushThunk ucsAct_3_EnumDecl, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -5
                            VbPegParseEnumDecl = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Public Function VbPegParseStructDecl() As Boolean
    Dim p122 As Long
    Dim p128 As Long
    Dim q128 As Long
    Dim i179 As Long
    Dim p174 As Long
    Dim q174 As Long
    Dim p149 As Long
    Dim q149 As Long
    Dim p167 As Long
    Dim q167 As Long
    Dim p182 As Long
    Dim q182 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 7
        p122 = .BufPos
        If Not (ParseTYPEDEF()) Then
            .BufPos = p122
        End If
        If ParseSTRUCT() Then
            pvPushThunk ucsActResultClear
            p128 = .BufPos
            q128 = .ThunkPos
            If Not (ParseID()) Then
                .BufPos = p128
                .ThunkPos = q128
            End If
            pvPushThunk ucsActVarSet, 1
            If ParseLBRACE() Then
                pvPushThunk ucsActResultClear
                pvPushThunk ucsActVarSet, 2
                pvPushThunk ucsAct_1_StructDecl, .CaptureBegin, .CaptureEnd
                For i179 = 0 To LNG_MAXINT
                    p174 = .BufPos
                    q174 = .ThunkPos
                    pvPushThunk ucsActResultClear
                    If VbPegParseType() Then
                        pvPushThunk ucsActVarSet, 3
                        pvPushThunk ucsActResultClear
                        If VbPegParseStuctMemList() Then
                            pvPushThunk ucsActVarSet, 4
                            pvPushThunk ucsActResultClear
                            p149 = .BufPos
                            q149 = .ThunkPos
                            If Not (VbPegParseArraySuffixList()) Then
                                .BufPos = p149
                                .ThunkPos = q149
                            End If
                            pvPushThunk ucsActVarSet, 5
                            If ParseSEMI() Then
                                pvPushThunk ucsAct_2_StructDecl, .CaptureBegin, .CaptureEnd
                            Else
                                .BufPos = p174
                                .ThunkPos = q174
                                pvPushThunk ucsActResultClear
                                If VbPegParseTypeUnlimited() Then
                                    pvPushThunk ucsActVarSet, 3
                                    If ParseLPAREN() Then
                                        If ParseCC_STDCALL() Then
                                            If ParseSTAR() Then
                                                pvPushThunk ucsActResultClear
                                                If VbPegParseIDList() Then
                                                    pvPushThunk ucsActVarSet, 4
                                                    If ParseRPAREN() Then
                                                        If ParseLPAREN() Then
                                                            pvPushThunk ucsActResultClear
                                                            p167 = .BufPos
                                                            q167 = .ThunkPos
                                                            If Not (VbPegParseParamList()) Then
                                                                .BufPos = p167
                                                                .ThunkPos = q167
                                                            End If
                                                            pvPushThunk ucsActVarSet, 6
                                                            If ParseRPAREN() Then
                                                                If ParseSEMI() Then
                                                                    pvPushThunk ucsAct_3_StructDecl, .CaptureBegin, .CaptureEnd
                                                                Else
                                                                    .BufPos = p174
                                                                    .ThunkPos = q174
                                                                    pvPushThunk ucsActResultClear
                                                                    If VbPegParseStructDecl() Then
                                                                        pvPushThunk ucsActVarSet, 3
                                                                        pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                                    Else
                                                                        .BufPos = p174
                                                                        .ThunkPos = q174
                                                                        .BufPos = p174
                                                                        .ThunkPos = q174
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            Else
                                                                .BufPos = p174
                                                                .ThunkPos = q174
                                                                pvPushThunk ucsActResultClear
                                                                If VbPegParseStructDecl() Then
                                                                    pvPushThunk ucsActVarSet, 3
                                                                    pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                                Else
                                                                    .BufPos = p174
                                                                    .ThunkPos = q174
                                                                    .BufPos = p174
                                                                    .ThunkPos = q174
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Else
                                                            .BufPos = p174
                                                            .ThunkPos = q174
                                                            pvPushThunk ucsActResultClear
                                                            If VbPegParseStructDecl() Then
                                                                pvPushThunk ucsActVarSet, 3
                                                                pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                            Else
                                                                .BufPos = p174
                                                                .ThunkPos = q174
                                                                .BufPos = p174
                                                                .ThunkPos = q174
                                                                Exit For
                                                            End If
                                                        End If
                                                    Else
                                                        .BufPos = p174
                                                        .ThunkPos = q174
                                                        pvPushThunk ucsActResultClear
                                                        If VbPegParseStructDecl() Then
                                                            pvPushThunk ucsActVarSet, 3
                                                            pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                        Else
                                                            .BufPos = p174
                                                            .ThunkPos = q174
                                                            .BufPos = p174
                                                            .ThunkPos = q174
                                                            Exit For
                                                        End If
                                                    End If
                                                Else
                                                    .BufPos = p174
                                                    .ThunkPos = q174
                                                    pvPushThunk ucsActResultClear
                                                    If VbPegParseStructDecl() Then
                                                        pvPushThunk ucsActVarSet, 3
                                                        pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                    Else
                                                        .BufPos = p174
                                                        .ThunkPos = q174
                                                        .BufPos = p174
                                                        .ThunkPos = q174
                                                        Exit For
                                                    End If
                                                End If
                                            Else
                                                .BufPos = p174
                                                .ThunkPos = q174
                                                pvPushThunk ucsActResultClear
                                                If VbPegParseStructDecl() Then
                                                    pvPushThunk ucsActVarSet, 3
                                                    pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                Else
                                                    .BufPos = p174
                                                    .ThunkPos = q174
                                                    .BufPos = p174
                                                    .ThunkPos = q174
                                                    Exit For
                                                End If
                                            End If
                                        Else
                                            .BufPos = p174
                                            .ThunkPos = q174
                                            pvPushThunk ucsActResultClear
                                            If VbPegParseStructDecl() Then
                                                pvPushThunk ucsActVarSet, 3
                                                pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                            Else
                                                .BufPos = p174
                                                .ThunkPos = q174
                                                .BufPos = p174
                                                .ThunkPos = q174
                                                Exit For
                                            End If
                                        End If
                                    Else
                                        .BufPos = p174
                                        .ThunkPos = q174
                                        pvPushThunk ucsActResultClear
                                        If VbPegParseStructDecl() Then
                                            pvPushThunk ucsActVarSet, 3
                                            pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                        Else
                                            .BufPos = p174
                                            .ThunkPos = q174
                                            .BufPos = p174
                                            .ThunkPos = q174
                                            Exit For
                                        End If
                                    End If
                                Else
                                    .BufPos = p174
                                    .ThunkPos = q174
                                    pvPushThunk ucsActResultClear
                                    If VbPegParseStructDecl() Then
                                        pvPushThunk ucsActVarSet, 3
                                        pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                    Else
                                        .BufPos = p174
                                        .ThunkPos = q174
                                        .BufPos = p174
                                        .ThunkPos = q174
                                        Exit For
                                    End If
                                End If
                            End If
                        Else
                            .BufPos = p174
                            .ThunkPos = q174
                            pvPushThunk ucsActResultClear
                            If VbPegParseTypeUnlimited() Then
                                pvPushThunk ucsActVarSet, 3
                                If ParseLPAREN() Then
                                    If ParseCC_STDCALL() Then
                                        If ParseSTAR() Then
                                            pvPushThunk ucsActResultClear
                                            If VbPegParseIDList() Then
                                                pvPushThunk ucsActVarSet, 4
                                                If ParseRPAREN() Then
                                                    If ParseLPAREN() Then
                                                        pvPushThunk ucsActResultClear
                                                        p167 = .BufPos
                                                        q167 = .ThunkPos
                                                        If Not (VbPegParseParamList()) Then
                                                            .BufPos = p167
                                                            .ThunkPos = q167
                                                        End If
                                                        pvPushThunk ucsActVarSet, 6
                                                        If ParseRPAREN() Then
                                                            If ParseSEMI() Then
                                                                pvPushThunk ucsAct_3_StructDecl, .CaptureBegin, .CaptureEnd
                                                            Else
                                                                .BufPos = p174
                                                                .ThunkPos = q174
                                                                pvPushThunk ucsActResultClear
                                                                If VbPegParseStructDecl() Then
                                                                    pvPushThunk ucsActVarSet, 3
                                                                    pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                                Else
                                                                    .BufPos = p174
                                                                    .ThunkPos = q174
                                                                    .BufPos = p174
                                                                    .ThunkPos = q174
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Else
                                                            .BufPos = p174
                                                            .ThunkPos = q174
                                                            pvPushThunk ucsActResultClear
                                                            If VbPegParseStructDecl() Then
                                                                pvPushThunk ucsActVarSet, 3
                                                                pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                            Else
                                                                .BufPos = p174
                                                                .ThunkPos = q174
                                                                .BufPos = p174
                                                                .ThunkPos = q174
                                                                Exit For
                                                            End If
                                                        End If
                                                    Else
                                                        .BufPos = p174
                                                        .ThunkPos = q174
                                                        pvPushThunk ucsActResultClear
                                                        If VbPegParseStructDecl() Then
                                                            pvPushThunk ucsActVarSet, 3
                                                            pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                        Else
                                                            .BufPos = p174
                                                            .ThunkPos = q174
                                                            .BufPos = p174
                                                            .ThunkPos = q174
                                                            Exit For
                                                        End If
                                                    End If
                                                Else
                                                    .BufPos = p174
                                                    .ThunkPos = q174
                                                    pvPushThunk ucsActResultClear
                                                    If VbPegParseStructDecl() Then
                                                        pvPushThunk ucsActVarSet, 3
                                                        pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                    Else
                                                        .BufPos = p174
                                                        .ThunkPos = q174
                                                        .BufPos = p174
                                                        .ThunkPos = q174
                                                        Exit For
                                                    End If
                                                End If
                                            Else
                                                .BufPos = p174
                                                .ThunkPos = q174
                                                pvPushThunk ucsActResultClear
                                                If VbPegParseStructDecl() Then
                                                    pvPushThunk ucsActVarSet, 3
                                                    pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                Else
                                                    .BufPos = p174
                                                    .ThunkPos = q174
                                                    .BufPos = p174
                                                    .ThunkPos = q174
                                                    Exit For
                                                End If
                                            End If
                                        Else
                                            .BufPos = p174
                                            .ThunkPos = q174
                                            pvPushThunk ucsActResultClear
                                            If VbPegParseStructDecl() Then
                                                pvPushThunk ucsActVarSet, 3
                                                pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                            Else
                                                .BufPos = p174
                                                .ThunkPos = q174
                                                .BufPos = p174
                                                .ThunkPos = q174
                                                Exit For
                                            End If
                                        End If
                                    Else
                                        .BufPos = p174
                                        .ThunkPos = q174
                                        pvPushThunk ucsActResultClear
                                        If VbPegParseStructDecl() Then
                                            pvPushThunk ucsActVarSet, 3
                                            pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                        Else
                                            .BufPos = p174
                                            .ThunkPos = q174
                                            .BufPos = p174
                                            .ThunkPos = q174
                                            Exit For
                                        End If
                                    End If
                                Else
                                    .BufPos = p174
                                    .ThunkPos = q174
                                    pvPushThunk ucsActResultClear
                                    If VbPegParseStructDecl() Then
                                        pvPushThunk ucsActVarSet, 3
                                        pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                    Else
                                        .BufPos = p174
                                        .ThunkPos = q174
                                        .BufPos = p174
                                        .ThunkPos = q174
                                        Exit For
                                    End If
                                End If
                            Else
                                .BufPos = p174
                                .ThunkPos = q174
                                pvPushThunk ucsActResultClear
                                If VbPegParseStructDecl() Then
                                    pvPushThunk ucsActVarSet, 3
                                    pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                Else
                                    .BufPos = p174
                                    .ThunkPos = q174
                                    .BufPos = p174
                                    .ThunkPos = q174
                                    Exit For
                                End If
                            End If
                        End If
                    Else
                        .BufPos = p174
                        .ThunkPos = q174
                        pvPushThunk ucsActResultClear
                        If VbPegParseTypeUnlimited() Then
                            pvPushThunk ucsActVarSet, 3
                            If ParseLPAREN() Then
                                If ParseCC_STDCALL() Then
                                    If ParseSTAR() Then
                                        pvPushThunk ucsActResultClear
                                        If VbPegParseIDList() Then
                                            pvPushThunk ucsActVarSet, 4
                                            If ParseRPAREN() Then
                                                If ParseLPAREN() Then
                                                    pvPushThunk ucsActResultClear
                                                    p167 = .BufPos
                                                    q167 = .ThunkPos
                                                    If Not (VbPegParseParamList()) Then
                                                        .BufPos = p167
                                                        .ThunkPos = q167
                                                    End If
                                                    pvPushThunk ucsActVarSet, 6
                                                    If ParseRPAREN() Then
                                                        If ParseSEMI() Then
                                                            pvPushThunk ucsAct_3_StructDecl, .CaptureBegin, .CaptureEnd
                                                        Else
                                                            .BufPos = p174
                                                            .ThunkPos = q174
                                                            pvPushThunk ucsActResultClear
                                                            If VbPegParseStructDecl() Then
                                                                pvPushThunk ucsActVarSet, 3
                                                                pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                            Else
                                                                .BufPos = p174
                                                                .ThunkPos = q174
                                                                .BufPos = p174
                                                                .ThunkPos = q174
                                                                Exit For
                                                            End If
                                                        End If
                                                    Else
                                                        .BufPos = p174
                                                        .ThunkPos = q174
                                                        pvPushThunk ucsActResultClear
                                                        If VbPegParseStructDecl() Then
                                                            pvPushThunk ucsActVarSet, 3
                                                            pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                        Else
                                                            .BufPos = p174
                                                            .ThunkPos = q174
                                                            .BufPos = p174
                                                            .ThunkPos = q174
                                                            Exit For
                                                        End If
                                                    End If
                                                Else
                                                    .BufPos = p174
                                                    .ThunkPos = q174
                                                    pvPushThunk ucsActResultClear
                                                    If VbPegParseStructDecl() Then
                                                        pvPushThunk ucsActVarSet, 3
                                                        pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                    Else
                                                        .BufPos = p174
                                                        .ThunkPos = q174
                                                        .BufPos = p174
                                                        .ThunkPos = q174
                                                        Exit For
                                                    End If
                                                End If
                                            Else
                                                .BufPos = p174
                                                .ThunkPos = q174
                                                pvPushThunk ucsActResultClear
                                                If VbPegParseStructDecl() Then
                                                    pvPushThunk ucsActVarSet, 3
                                                    pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                                Else
                                                    .BufPos = p174
                                                    .ThunkPos = q174
                                                    .BufPos = p174
                                                    .ThunkPos = q174
                                                    Exit For
                                                End If
                                            End If
                                        Else
                                            .BufPos = p174
                                            .ThunkPos = q174
                                            pvPushThunk ucsActResultClear
                                            If VbPegParseStructDecl() Then
                                                pvPushThunk ucsActVarSet, 3
                                                pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                            Else
                                                .BufPos = p174
                                                .ThunkPos = q174
                                                .BufPos = p174
                                                .ThunkPos = q174
                                                Exit For
                                            End If
                                        End If
                                    Else
                                        .BufPos = p174
                                        .ThunkPos = q174
                                        pvPushThunk ucsActResultClear
                                        If VbPegParseStructDecl() Then
                                            pvPushThunk ucsActVarSet, 3
                                            pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                        Else
                                            .BufPos = p174
                                            .ThunkPos = q174
                                            .BufPos = p174
                                            .ThunkPos = q174
                                            Exit For
                                        End If
                                    End If
                                Else
                                    .BufPos = p174
                                    .ThunkPos = q174
                                    pvPushThunk ucsActResultClear
                                    If VbPegParseStructDecl() Then
                                        pvPushThunk ucsActVarSet, 3
                                        pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                    Else
                                        .BufPos = p174
                                        .ThunkPos = q174
                                        .BufPos = p174
                                        .ThunkPos = q174
                                        Exit For
                                    End If
                                End If
                            Else
                                .BufPos = p174
                                .ThunkPos = q174
                                pvPushThunk ucsActResultClear
                                If VbPegParseStructDecl() Then
                                    pvPushThunk ucsActVarSet, 3
                                    pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                                Else
                                    .BufPos = p174
                                    .ThunkPos = q174
                                    .BufPos = p174
                                    .ThunkPos = q174
                                    Exit For
                                End If
                            End If
                        Else
                            .BufPos = p174
                            .ThunkPos = q174
                            pvPushThunk ucsActResultClear
                            If VbPegParseStructDecl() Then
                                pvPushThunk ucsActVarSet, 3
                                pvPushThunk ucsAct_4_StructDecl, .CaptureBegin, .CaptureEnd
                            Else
                                .BufPos = p174
                                .ThunkPos = q174
                                .BufPos = p174
                                .ThunkPos = q174
                                Exit For
                            End If
                        End If
                    End If
                Next
                If i179 <> 0 Then
                    If ParseRBRACE() Then
                        pvPushThunk ucsActResultClear
                        p182 = .BufPos
                        q182 = .ThunkPos
                        If Not (VbPegParseIDList()) Then
                            .BufPos = p182
                            .ThunkPos = q182
                        End If
                        pvPushThunk ucsActVarSet, 7
                        If ParseSEMI() Then
                            pvPushThunk ucsAct_5_StructDecl, .CaptureBegin, .CaptureEnd
                            pvPushThunk ucsActVarAlloc, -7
                            VbPegParseStructDecl = True
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Public Function VbPegParseFunDecl() As Boolean
    Dim p188 As Long
    Dim p204 As Long
    Dim q204 As Long
    Dim p209 As Long
    Dim q209 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        p188 = .BufPos
        If Not (VbPegParseLinkage()) Then
            .BufPos = p188
        End If
        pvPushThunk ucsActResultClear
        If VbPegParseType() Then
            pvPushThunk ucsActVarSet, 1
            If ParseCC_STDCALL() Then
                pvPushThunk ucsActResultClear
                If ParseID() Then
                    pvPushThunk ucsActVarSet, 2
                    If ParseLPAREN() Then
                        p204 = .BufPos
                        q204 = .ThunkPos
                        pvPushThunk ucsActResultClear
                        If VbPegParseParamList() Then
                            pvPushThunk ucsActVarSet, 3
                        Else
                            .BufPos = p204
                            .ThunkPos = q204
                        End If
                        If ParseRPAREN() Then
                            p209 = .BufPos
                            q209 = .ThunkPos
                            If ParseSEMI() Then
                                pvPushThunk ucsAct_1_FunDecl, .CaptureBegin, .CaptureEnd
                                pvPushThunk ucsActVarAlloc, -3
                                VbPegParseFunDecl = True
                                Exit Function
                            Else
                                .BufPos = p209
                                .ThunkPos = q209
                            End If
                            If ParseLBRACE() Then
                                pvPushThunk ucsAct_1_FunDecl, .CaptureBegin, .CaptureEnd
                                pvPushThunk ucsActVarAlloc, -3
                                VbPegParseFunDecl = True
                                Exit Function
                            Else
                                .BufPos = p209
                                .ThunkPos = q209
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Function

Public Function VbPegParseSkipStmt() As Boolean
    Dim lCaptureBegin As Long
    Dim p225 As Long
    Dim p216 As Long
    Dim p215 As Long
    Dim p221 As Long
    Dim p230 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Do
            p225 = .BufPos
            p216 = .BufPos
            p215 = .BufPos
            If ParseNL() Then
                .BufPos = p225
                Exit Do
            Else
                .BufPos = p215
            End If
            If ParseSEMI() Then
                .BufPos = p225
                Exit Do
            Else
                .BufPos = p215
            End If
            .BufPos = p216
            p221 = .BufPos
            If Not (ParseBLOCKCOMMENT()) Then
                .BufPos = p221
                If Not (ParseLINECOMMENT()) Then
                    .BufPos = p221
                    If Not (ParseWS()) Then
                        .BufPos = p221
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p225
                            Exit Do
                        End If
                    End If
                End If
            End If
        Loop
        p230 = .BufPos
        If ParseNL() Then
            Call Parse_
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_SkipStmt, lCaptureBegin, lCaptureEnd
            VbPegParseSkipStmt = True
            Exit Function
        Else
            .BufPos = p230
        End If
        If ParseSEMI() Then
            Call Parse_
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_SkipStmt, lCaptureBegin, lCaptureEnd
            VbPegParseSkipStmt = True
            Exit Function
        Else
            .BufPos = p230
        End If
    End With
End Function

Private Function ParseEOL() As Boolean
    Dim p688 As Long

    With ctx
        p688 = .BufPos
        If Not (.BufPos < .BufSize) Then
            .BufPos = p688
            ParseEOL = True
        End If
    End With
End Function

Private Function ParseTYPEDEF() As Boolean
    Dim p477 As Long

    With ctx
        If pvMatchString("typedef", NORM_IGNORECASE) Then ' "typedef"i
            .BufPos = .BufPos + 7
            p477 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p477
                Call Parse_
                Call pvSetAdvance
                ParseTYPEDEF = True
            End Select
        End If
    End With
End Function

Public Function VbPegParseType() As Boolean
    Dim lCaptureBegin As Long
    Dim p247 As Long
    Dim q247 As Long
    Dim p237 As Long
    Dim p250 As Long
    Dim q250 As Long
    Dim e250 As String
    Dim p252 As Long
    Dim q252 As Long
    Dim e252 As String
    Dim p254 As Long
    Dim q254 As Long
    Dim lCaptureEnd As Long
    Dim i243 As Long
    Dim p242 As Long
    Dim q242 As Long
    Dim p244 As Long
    Dim q244 As Long

    With ctx
        lCaptureBegin = .BufPos
        p247 = .BufPos
        q247 = .ThunkPos
        Do
            p237 = .BufPos
            If Not (VbPegParseTypePrefix()) Then
                .BufPos = p237
                Exit Do
            End If
        Loop
        If VbPegParseTypeBody() Then
            p250 = .BufPos
            q250 = .ThunkPos
            e250 = .LastExpected
            If Not (ParseLPAREN()) Then
                .BufPos = p250
                .ThunkPos = q250
                .LastExpected = e250
                p252 = .BufPos
                q252 = .ThunkPos
                e252 = .LastExpected
                If Not (ParseSEMI()) Then
                    .BufPos = p252
                    .ThunkPos = q252
                    .LastExpected = e252
                    p254 = .BufPos
                    q254 = .ThunkPos
                    Call VbPegParseTypeSuffix
                    lCaptureEnd = .BufPos
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_1_Type, lCaptureBegin, lCaptureEnd
                    VbPegParseType = True
                    Exit Function
                End If
            End If
        Else
            .BufPos = p247
            .ThunkPos = q247
        End If
        For i243 = 0 To LNG_MAXINT
            p242 = .BufPos
            q242 = .ThunkPos
            If Not (VbPegParseTypePrefix()) Then
                .BufPos = p242
                .ThunkPos = q242
                Exit For
            End If
        Next
        If i243 <> 0 Then
            p244 = .BufPos
            q244 = .ThunkPos
            If Not (VbPegParseTypeBody()) Then
                .BufPos = p244
                .ThunkPos = q244
            End If
            p250 = .BufPos
            q250 = .ThunkPos
            e250 = .LastExpected
            If Not (ParseLPAREN()) Then
                .BufPos = p250
                .ThunkPos = q250
                .LastExpected = e250
                p252 = .BufPos
                q252 = .ThunkPos
                e252 = .LastExpected
                If Not (ParseSEMI()) Then
                    .BufPos = p252
                    .ThunkPos = q252
                    .LastExpected = e252
                    p254 = .BufPos
                    q254 = .ThunkPos
                    Call VbPegParseTypeSuffix
                    lCaptureEnd = .BufPos
                    .CaptureBegin = lCaptureBegin
                    .CaptureEnd = lCaptureEnd
                    pvPushThunk ucsAct_1_Type, lCaptureBegin, lCaptureEnd
                    VbPegParseType = True
                    Exit Function
                End If
            End If
        Else
            .BufPos = p247
            .ThunkPos = q247
        End If
    End With
End Function

Private Function ParseID() As Boolean
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95                ' [a-zA-Z_]
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                    .BufPos = .BufPos + 1
                Case Else
                    Exit Do
                End Select
            Loop
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_ID, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            ParseID = True
        End Select
    End With
End Function

Private Function ParseSEMI() As Boolean
    With ctx
        If .BufData(.BufPos) = 59 Then              ' ";"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseSEMI = True
        End If
    End With
End Function

Public Function VbPegParseLinkage() As Boolean
    Dim p410 As Long
    Dim p416 As Long

    With ctx
        p410 = .BufPos
        If ParseEXTERN() Then
            p416 = .BufPos
            If Not (ParseINLINE()) Then
                .BufPos = p416
            End If
            VbPegParseLinkage = True
            Exit Function
        Else
            .BufPos = p410
        End If
        If ParseSTATIC() Then
            p416 = .BufPos
            If Not (ParseINLINE()) Then
                .BufPos = p416
            End If
            VbPegParseLinkage = True
            Exit Function
        Else
            .BufPos = p410
        End If
        If ParseCAIRO_PUBLIC() Then
            p416 = .BufPos
            If Not (ParseINLINE()) Then
                .BufPos = p416
            End If
            VbPegParseLinkage = True
            Exit Function
        Else
            .BufPos = p410
        End If
        If ParseCAIRO_WARN() Then
            p416 = .BufPos
            If Not (ParseINLINE()) Then
                .BufPos = p416
            End If
            VbPegParseLinkage = True
            Exit Function
        Else
            .BufPos = p410
        End If
    End With
End Function

Public Function VbPegParseTypeUnlimited() As Boolean
    Dim lCaptureBegin As Long
    Dim p268 As Long
    Dim q268 As Long
    Dim p259 As Long
    Dim p270 As Long
    Dim q270 As Long
    Dim lCaptureEnd As Long
    Dim i264 As Long
    Dim p263 As Long
    Dim q263 As Long
    Dim p265 As Long
    Dim q265 As Long

    With ctx
        lCaptureBegin = .BufPos
        p268 = .BufPos
        q268 = .ThunkPos
        Do
            p259 = .BufPos
            If Not (VbPegParseTypePrefix()) Then
                .BufPos = p259
                Exit Do
            End If
        Loop
        If VbPegParseTypeBody() Then
            p270 = .BufPos
            q270 = .ThunkPos
            Call VbPegParseTypeSuffix
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_TypeUnlimited, lCaptureBegin, lCaptureEnd
            VbPegParseTypeUnlimited = True
            Exit Function
        Else
            .BufPos = p268
            .ThunkPos = q268
        End If
        For i264 = 0 To LNG_MAXINT
            p263 = .BufPos
            q263 = .ThunkPos
            If Not (VbPegParseTypePrefix()) Then
                .BufPos = p263
                .ThunkPos = q263
                Exit For
            End If
        Next
        If i264 <> 0 Then
            p265 = .BufPos
            q265 = .ThunkPos
            If Not (VbPegParseTypeBody()) Then
                .BufPos = p265
                .ThunkPos = q265
            End If
            p270 = .BufPos
            q270 = .ThunkPos
            Call VbPegParseTypeSuffix
            lCaptureEnd = .BufPos
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_TypeUnlimited, lCaptureBegin, lCaptureEnd
            VbPegParseTypeUnlimited = True
            Exit Function
        Else
            .BufPos = p268
            .ThunkPos = q268
        End If
    End With
End Function

Private Function ParseLPAREN() As Boolean
    With ctx
        If .BufData(.BufPos) = 40 Then              ' "("
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseLPAREN = True
        End If
    End With
End Function

Private Function ParseCC_STDCALL() As Boolean
    Dim p575 As Long

    With ctx
        If pvMatchString("CAIRO_CALLCONV") Then     ' "CAIRO_CALLCONV"
            .BufPos = .BufPos + 14
            p575 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p575
                Call Parse_
                Call pvSetAdvance
                ParseCC_STDCALL = True
            End Select
        End If
        If pvMatchString("WINAPI") Then             ' "WINAPI"
            .BufPos = .BufPos + 6
            p575 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p575
                Call Parse_
                Call pvSetAdvance
                ParseCC_STDCALL = True
            End Select
        End If
        p575 = .BufPos
        Select Case .BufData(.BufPos)
        Case 97 To 122, 65 To 90, 95, 48 To 57, 35  ' [a-zA-Z_0-9#]
            '--- do nothing
        Case Else
            .BufPos = p575
            Call Parse_
            Call pvSetAdvance
            ParseCC_STDCALL = True
        End Select
    End With
End Function

Private Function ParseSTAR() As Boolean
    With ctx
        If .BufData(.BufPos) = 42 Then              ' "*"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseSTAR = True
        End If
    End With
End Function

Private Function ParseRPAREN() As Boolean
    With ctx
        If .BufData(.BufPos) = 41 Then              ' ")"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseRPAREN = True
        End If
    End With
End Function

Public Function VbPegParseParamList() As Boolean
    Dim p338 As Long
    Dim q338 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If VbPegParseParam() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_ParamList, .CaptureBegin, .CaptureEnd
            Do
                p338 = .BufPos
                q338 = .ThunkPos
                If Not (ParseCOMMA()) Then
                    .BufPos = p338
                    .ThunkPos = q338
                    Exit Do
                End If
                pvPushThunk ucsActResultClear
                If VbPegParseParam() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p338
                    .ThunkPos = q338
                    Exit Do
                End If
                pvPushThunk ucsAct_2_ParamList, .CaptureBegin, .CaptureEnd
            Loop
            pvPushThunk ucsAct_3_ParamList, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            VbPegParseParamList = True
        End If
    End With
End Function

Private Function ParseENUM() As Boolean
    Dim p580 As Long

    With ctx
        If pvMatchString("enum") Then               ' "enum"
            .BufPos = .BufPos + 4
            p580 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p580
                Call Parse_
                Call pvSetAdvance
                ParseENUM = True
            End Select
        End If
    End With
End Function

Private Function ParseLBRACE() As Boolean
    With ctx
        If .BufData(.BufPos) = 123 Then             ' "{"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseLBRACE = True
        End If
    End With
End Function

Public Function VbPegParseEnumValue() As Boolean
    Dim p374 As Long
    Dim q374 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        If ParseEQ() Then
            pvPushThunk ucsActResultClear
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_EnumValue, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActResultClear
            If VbPegParseEnumValueToken() Then
                pvPushThunk ucsActVarSet, 2
                pvPushThunk ucsAct_2_EnumValue, .CaptureBegin, .CaptureEnd
                Do
                    p374 = .BufPos
                    q374 = .ThunkPos
                    pvPushThunk ucsActResultClear
                    If VbPegParseEnumValueToken() Then
                        pvPushThunk ucsActVarSet, 2
                    Else
                        .BufPos = p374
                        .ThunkPos = q374
                        Exit Do
                    End If
                    pvPushThunk ucsAct_3_EnumValue, .CaptureBegin, .CaptureEnd
                Loop
                pvPushThunk ucsAct_4_EnumValue, .CaptureBegin, .CaptureEnd
                pvPushThunk ucsActVarAlloc, -2
                VbPegParseEnumValue = True
            End If
        End If
    End With
End Function

Private Sub ParseEMPTY()
    Dim lCaptureBegin As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        lCaptureEnd = .BufPos
        .CaptureBegin = lCaptureBegin
        .CaptureEnd = lCaptureEnd
        pvPushThunk ucsAct_1_EMPTY, lCaptureBegin, lCaptureEnd
L39:
    End With
End Sub

Private Function ParseCOMMA() As Boolean
    With ctx
        If .BufData(.BufPos) = 44 Then              ' ","
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseCOMMA = True
        End If
    End With
End Function

Private Function ParseRBRACE() As Boolean
    With ctx
        If .BufData(.BufPos) = 125 Then             ' "}"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseRBRACE = True
        End If
    End With
End Function

Public Function VbPegParseIDList() As Boolean
    Dim p429 As Long
    Dim q429 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseID() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_IDList, .CaptureBegin, .CaptureEnd
            Do
                p429 = .BufPos
                q429 = .ThunkPos
                If Not (ParseCOMMA()) Then
                    .BufPos = p429
                    .ThunkPos = q429
                    Exit Do
                End If
                pvPushThunk ucsActResultClear
                If ParseID() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p429
                    .ThunkPos = q429
                    Exit Do
                End If
                pvPushThunk ucsAct_2_IDList, .CaptureBegin, .CaptureEnd
            Loop
            pvPushThunk ucsAct_3_IDList, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            VbPegParseIDList = True
        End If
    End With
End Function

Private Function ParseSTRUCT() As Boolean
    Dim p585 As Long

    With ctx
        If pvMatchString("struct") Then             ' "struct"
            .BufPos = .BufPos + 6
            p585 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p585
                Call Parse_
                Call pvSetAdvance
                ParseSTRUCT = True
            End Select
        End If
    End With
End Function

Public Function VbPegParseStuctMemList() As Boolean
    Dim p441 As Long
    Dim q441 As Long
    Dim p438 As Long
    Dim q438 As Long
    Dim p442 As Long
    Dim q442 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If ParseID() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_StuctMemList, .CaptureBegin, .CaptureEnd
            Do
                p441 = .BufPos
                q441 = .ThunkPos
                p438 = .BufPos
                q438 = .ThunkPos
                If Not (VbPegParseArraySuffix()) Then
                    .BufPos = p438
                    .ThunkPos = q438
                End If
                If Not (ParseCOMMA()) Then
                    .BufPos = p441
                    .ThunkPos = q441
                    Exit Do
                End If
                p442 = .BufPos
                q442 = .ThunkPos
                If Not (ParseSTAR()) Then
                    .BufPos = p442
                    .ThunkPos = q442
                End If
                pvPushThunk ucsActResultClear
                If ParseID() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p441
                    .ThunkPos = q441
                    Exit Do
                End If
                pvPushThunk ucsAct_2_StuctMemList, .CaptureBegin, .CaptureEnd
            Loop
            pvPushThunk ucsAct_3_StuctMemList, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            VbPegParseStuctMemList = True
        End If
    End With
End Function

Public Function VbPegParseArraySuffixList() As Boolean
    Dim p459 As Long
    Dim q459 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 2
        pvPushThunk ucsActResultClear
        If VbPegParseArraySuffix() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsAct_1_ArraySuffixList, .CaptureBegin, .CaptureEnd
            Do
                p459 = .BufPos
                q459 = .ThunkPos
                pvPushThunk ucsActResultClear
                If VbPegParseArraySuffix() Then
                    pvPushThunk ucsActVarSet, 2
                Else
                    .BufPos = p459
                    .ThunkPos = q459
                    Exit Do
                End If
                pvPushThunk ucsAct_2_ArraySuffixList, .CaptureBegin, .CaptureEnd
            Loop
            pvPushThunk ucsAct_3_ArraySuffixList, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -2
            VbPegParseArraySuffixList = True
        End If
    End With
End Function

Private Function ParseNL() As Boolean
    Dim p657 As Long

    With ctx
        If .BufData(.BufPos) = 13 Then              ' "\r"
            .BufPos = .BufPos + 1
        End If
        If .BufData(.BufPos) = 10 Then              ' "\n"
            .BufPos = .BufPos + 1
            p657 = .BufPos
            If Not (ParsePREPRO()) Then
                .BufPos = p657
            End If
            Call pvSetAdvance
            ParseNL = True
        End If
    End With
End Function

Private Function ParseBLOCKCOMMENT() As Boolean
    Dim p673 As Long
    Dim p666 As Long
    Dim p672 As Long
    Dim p668 As Long

    With ctx
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 42 Then ' "/*"
            .BufPos = .BufPos + 2
            Do
                p673 = .BufPos
                p666 = .BufPos
                If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                    .BufPos = p673
                    Exit Do
                Else
                    .BufPos = p666
                End If
                p672 = .BufPos
                p668 = .BufPos
                If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 42 Then ' "/*"
                    .BufPos = p668
                    If Not (ParseBLOCKCOMMENT()) Then
                        .BufPos = p672
                        If .BufPos < .BufSize Then
                            .BufPos = .BufPos + 1
                        Else
                            .BufPos = p673
                            Exit Do
                        End If
                    End If
                Else
                    .BufPos = p672
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        .BufPos = p673
                        Exit Do
                    End If
                End If
            Loop
            If .BufData(.BufPos) = 42 And .BufData(.BufPos + 1) = 47 Then ' "*/"
                .BufPos = .BufPos + 2
                Call pvSetAdvance
                ParseBLOCKCOMMENT = True
            End If
        End If
    End With
End Function

Private Function ParseLINECOMMENT() As Boolean
    With ctx
        If .BufData(.BufPos) = 47 And .BufData(.BufPos + 1) = 47 Then ' "//"
            .BufPos = .BufPos + 2
            Do
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    Exit Do
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        Exit Do
                    End If
                End Select
            Loop
            If ParseNL() Then
                Call pvSetAdvance
                ParseLINECOMMENT = True
            End If
        End If
    End With
End Function

Private Function ParseWS() As Boolean
    Dim i651 As Long
    Dim p650 As Long

    With ctx
        For i651 = 0 To LNG_MAXINT
            p650 = .BufPos
            Select Case .BufData(.BufPos)
            Case 32, 9                              ' [ \t]
                .BufPos = .BufPos + 1
            Case Else
                If Not (ParseNL()) Then
                    .BufPos = p650
                    .BufPos = p650
                    Exit For
                End If
            End Select
        Next
        If i651 <> 0 Then
            Call pvSetAdvance
            ParseWS = True
        End If
    End With
End Function

Private Sub Parse_()
    Dim p645 As Long

    With ctx
        Do
            p645 = .BufPos
            If Not (ParseBLOCKCOMMENT()) Then
                .BufPos = p645
                If Not (ParseLINECOMMENT()) Then
                    .BufPos = p645
                    If Not (ParseWS()) Then
                        .BufPos = p645
                        .BufPos = p645
                        Exit Do
                    End If
                End If
            End If
        Loop
    End With
End Sub

Public Function VbPegParseTypePrefix() As Boolean
    Dim p278 As Long

    With ctx
        p278 = .BufPos
        If ParseCONST() Then
            VbPegParseTypePrefix = True
            Exit Function
        Else
            .BufPos = p278
        End If
        If ParseUNSIGNED() Then
            VbPegParseTypePrefix = True
            Exit Function
        Else
            .BufPos = p278
        End If
        If ParseSTRUCT() Then
            VbPegParseTypePrefix = True
            Exit Function
        Else
            .BufPos = p278
        End If
        If ParseENUM() Then
            VbPegParseTypePrefix = True
            Exit Function
        Else
            .BufPos = p278
        End If
    End With
End Function

Public Function VbPegParseTypeBody() As Boolean
    Dim p285 As Long
    Dim q285 As Long

    With ctx
        p285 = .BufPos
        q285 = .ThunkPos
        If ParseINT() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseCHAR() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseUNSIGNED() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseVOID() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseINT_A_T() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseINT_B_T() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseINT_C_T() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseUINT_A_T() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseUINT_B_T() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseUINT_C_T() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseUINT_D_T() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseUINTPTR_T() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseSIZE_T() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseDOUBLE() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseLONG() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseLONG_LONG() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If ParseBOOL() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
        If VbPegParseRefType() Then
            VbPegParseTypeBody = True
            Exit Function
        Else
            .BufPos = p285
            .ThunkPos = q285
        End If
    End With
End Function

Public Sub VbPegParseTypeSuffix()
    Dim p320 As Long
    Dim p317 As Long

    With ctx
        Do
            p320 = .BufPos
            p317 = .BufPos
            If Not (ParseCONST()) Then
                .BufPos = p317
            End If
            If Not (ParseSTAR()) Then
                .BufPos = p320
                Exit Do
            End If
        Loop
    End With
End Sub

Private Function ParseCONST() As Boolean
    Dim p497 As Long

    With ctx
        If pvMatchString("const") Then              ' "const"
            .BufPos = .BufPos + 5
            p497 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p497
                Call Parse_
                Call pvSetAdvance
                ParseCONST = True
            End Select
        End If
    End With
End Function

Private Function ParseUNSIGNED() As Boolean
    Dim p492 As Long

    With ctx
        If pvMatchString("unsigned") Then           ' "unsigned"
            .BufPos = .BufPos + 8
            p492 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p492
                Call Parse_
                Call pvSetAdvance
                ParseUNSIGNED = True
            End Select
        End If
    End With
End Function

Private Function ParseINT() As Boolean
    Dim p482 As Long

    With ctx
        If pvMatchString("int") Then                ' "int"
            .BufPos = .BufPos + 3
            p482 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p482
                Call Parse_
                Call pvSetAdvance
                ParseINT = True
            End Select
        End If
    End With
End Function

Private Function ParseCHAR() As Boolean
    Dim p487 As Long

    With ctx
        If pvMatchString("char") Then               ' "char"
            .BufPos = .BufPos + 4
            p487 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p487
                Call Parse_
                Call pvSetAdvance
                ParseCHAR = True
            End Select
        End If
    End With
End Function

Private Function ParseVOID() As Boolean
    Dim p502 As Long

    With ctx
        If pvMatchString("void") Then               ' "void"
            .BufPos = .BufPos + 4
            p502 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p502
                Call Parse_
                Call pvSetAdvance
                ParseVOID = True
            End Select
        End If
    End With
End Function

Private Function ParseINT_A_T() As Boolean
    Dim p507 As Long

    With ctx
        If pvMatchString("int8_t") Then             ' "int8_t"
            .BufPos = .BufPos + 6
            p507 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p507
                Call Parse_
                Call pvSetAdvance
                ParseINT_A_T = True
            End Select
        End If
    End With
End Function

Private Function ParseINT_B_T() As Boolean
    Dim p512 As Long

    With ctx
        If pvMatchString("int32_t") Then            ' "int32_t"
            .BufPos = .BufPos + 7
            p512 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p512
                Call Parse_
                Call pvSetAdvance
                ParseINT_B_T = True
            End Select
        End If
    End With
End Function

Private Function ParseINT_C_T() As Boolean
    Dim p517 As Long

    With ctx
        If pvMatchString("int64_t") Then            ' "int64_t"
            .BufPos = .BufPos + 7
            p517 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p517
                Call Parse_
                Call pvSetAdvance
                ParseINT_C_T = True
            End Select
        End If
    End With
End Function

Private Function ParseUINT_A_T() As Boolean
    Dim p522 As Long

    With ctx
        If pvMatchString("uint8_t") Then            ' "uint8_t"
            .BufPos = .BufPos + 7
            p522 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p522
                Call Parse_
                Call pvSetAdvance
                ParseUINT_A_T = True
            End Select
        End If
    End With
End Function

Private Function ParseUINT_B_T() As Boolean
    Dim p527 As Long

    With ctx
        If pvMatchString("uint32_t") Then           ' "uint32_t"
            .BufPos = .BufPos + 8
            p527 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p527
                Call Parse_
                Call pvSetAdvance
                ParseUINT_B_T = True
            End Select
        End If
    End With
End Function

Private Function ParseUINT_C_T() As Boolean
    Dim p532 As Long

    With ctx
        If pvMatchString("uint64_t") Then           ' "uint64_t"
            .BufPos = .BufPos + 8
            p532 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p532
                Call Parse_
                Call pvSetAdvance
                ParseUINT_C_T = True
            End Select
        End If
    End With
End Function

Private Function ParseUINT_D_T() As Boolean
    Dim p537 As Long

    With ctx
        If pvMatchString("uint16_t") Then           ' "uint16_t"
            .BufPos = .BufPos + 8
            p537 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p537
                Call Parse_
                Call pvSetAdvance
                ParseUINT_D_T = True
            End Select
        End If
    End With
End Function

Private Function ParseUINTPTR_T() As Boolean
    Dim p542 As Long

    With ctx
        If pvMatchString("uintptr_t") Then          ' "uintptr_t"
            .BufPos = .BufPos + 9
            p542 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p542
                Call Parse_
                Call pvSetAdvance
                ParseUINTPTR_T = True
            End Select
        End If
    End With
End Function

Private Function ParseSIZE_T() As Boolean
    Dim p547 As Long

    With ctx
        If pvMatchString("size_t") Then             ' "size_t"
            .BufPos = .BufPos + 6
            p547 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p547
                Call Parse_
                Call pvSetAdvance
                ParseSIZE_T = True
            End Select
        End If
    End With
End Function

Private Function ParseDOUBLE() As Boolean
    Dim p552 As Long

    With ctx
        If pvMatchString("double") Then             ' "double"
            .BufPos = .BufPos + 6
            p552 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p552
                Call Parse_
                Call pvSetAdvance
                ParseDOUBLE = True
            End Select
        End If
    End With
End Function

Private Function ParseLONG() As Boolean
    Dim p557 As Long

    With ctx
        If pvMatchString("long") Then               ' "long"
            .BufPos = .BufPos + 4
            p557 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p557
                Call Parse_
                Call pvSetAdvance
                ParseLONG = True
            End Select
        End If
    End With
End Function

Private Function ParseLONG_LONG() As Boolean
    Dim p562 As Long

    With ctx
        If pvMatchString("long long") Then          ' "long long"
            .BufPos = .BufPos + 9
            p562 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p562
                Call Parse_
                Call pvSetAdvance
                ParseLONG_LONG = True
            End Select
        End If
    End With
End Function

Private Function ParseBOOL() As Boolean
    Dim p567 As Long

    With ctx
        If pvMatchString("bool") Then               ' "bool"
            .BufPos = .BufPos + 4
            p567 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p567
                Call Parse_
                Call pvSetAdvance
                ParseBOOL = True
            End Select
        End If
    End With
End Function

Public Function VbPegParseRefType() As Boolean
    Dim p323 As Long
    Dim e323 As String

    With ctx
        p323 = .BufPos
        e323 = .LastExpected
        If Not (ParseCC_STDCALL()) Then
            .BufPos = p323
            .LastExpected = e323
            If ParseID() Then
                If IsRefType(Mid$(.Contents, .CaptureBegin + 1, .CaptureEnd - .CaptureBegin)) Then
                    VbPegParseRefType = True
                End If
            End If
        End If
    End With
End Function

Public Function VbPegParseParam() As Boolean
    Dim p347 As Long
    Dim q347 As Long
    Dim p353 As Long
    Dim q353 As Long

    With ctx
        pvPushThunk ucsActVarAlloc, 3
        pvPushThunk ucsActResultClear
        If VbPegParseType() Then
            pvPushThunk ucsActVarSet, 1
            pvPushThunk ucsActResultClear
            p347 = .BufPos
            q347 = .ThunkPos
            If Not (ParseID()) Then
                .BufPos = p347
                .ThunkPos = q347
            End If
            pvPushThunk ucsActVarSet, 2
            pvPushThunk ucsActResultClear
            p353 = .BufPos
            q353 = .ThunkPos
            If Not (VbPegParseArraySuffix()) Then
                .BufPos = p353
                .ThunkPos = q353
            End If
            pvPushThunk ucsActVarSet, 3
            pvPushThunk ucsAct_1_Param, .CaptureBegin, .CaptureEnd
            pvPushThunk ucsActVarAlloc, -3
            VbPegParseParam = True
        End If
    End With
End Function

Public Function VbPegParseArraySuffix() As Boolean
    Dim lCaptureBegin As Long
    Dim p400 As Long
    Dim p398 As Long
    Dim e398 As String
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        If ParseLBRACKET() Then
            Do
                p400 = .BufPos
                p398 = .BufPos
                e398 = .LastExpected
                If ParseRBRACKET() Then
                    .BufPos = p400
                    Exit Do
                Else
                    .BufPos = p398
                    .LastExpected = e398
                End If
                If .BufPos < .BufSize Then
                    .BufPos = .BufPos + 1
                Else
                    .BufPos = p400
                    Exit Do
                End If
            Loop
            If ParseRBRACKET() Then
                lCaptureEnd = .BufPos
                Call Parse_
                .CaptureBegin = lCaptureBegin
                .CaptureEnd = lCaptureEnd
                pvPushThunk ucsAct_1_ArraySuffix, lCaptureBegin, lCaptureEnd
                VbPegParseArraySuffix = True
            End If
        End If
    End With
End Function

Private Function ParseEQ() As Boolean
    With ctx
        If .BufData(.BufPos) = 61 Then              ' "="
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseEQ = True
        End If
    End With
End Function

Public Function VbPegParseEnumValueToken() As Boolean
    Dim lCaptureBegin As Long
    Dim i387 As Long
    Dim p386 As Long
    Dim p384 As Long
    Dim p381 As Long
    Dim lCaptureEnd As Long

    With ctx
        lCaptureBegin = .BufPos
        For i387 = 0 To LNG_MAXINT
            p386 = .BufPos
            p384 = .BufPos
            p381 = .BufPos
            If ParseBLOCKCOMMENT() Then
                .BufPos = p386
                Exit For
            Else
                .BufPos = p381
            End If
            If ParseLINECOMMENT() Then
                .BufPos = p386
                Exit For
            Else
                .BufPos = p381
            End If
            If ParseWS() Then
                .BufPos = p386
                Exit For
            Else
                .BufPos = p381
            End If
            Select Case .BufData(.BufPos)
            Case 44, 125                            ' [,}]
                .BufPos = .BufPos + 1
                .BufPos = p386
                Exit For
            End Select
            .BufPos = p384
            If .BufPos < .BufSize Then
                .BufPos = .BufPos + 1
            Else
                .BufPos = p386
                Exit For
            End If
        Next
        If i387 <> 0 Then
            lCaptureEnd = .BufPos
            Call Parse_
            .CaptureBegin = lCaptureBegin
            .CaptureEnd = lCaptureEnd
            pvPushThunk ucsAct_1_EnumValueToken, lCaptureBegin, lCaptureEnd
            Call pvSetAdvance
            VbPegParseEnumValueToken = True
        End If
    End With
End Function

Private Function ParseLBRACKET() As Boolean
    With ctx
        If .BufData(.BufPos) = 91 Then              ' "["
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseLBRACKET = True
        End If
    End With
End Function

Private Function ParseRBRACKET() As Boolean
    With ctx
        If .BufData(.BufPos) = 93 Then              ' "]"
            .BufPos = .BufPos + 1
            Call Parse_
            Call pvSetAdvance
            ParseRBRACKET = True
        End If
    End With
End Function

Private Function ParseEXTERN() As Boolean
    Dim p590 As Long

    With ctx
        If pvMatchString("extern") Then             ' "extern"
            .BufPos = .BufPos + 6
            p590 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p590
                Call Parse_
                Call pvSetAdvance
                ParseEXTERN = True
            End Select
        End If
    End With
End Function

Private Function ParseSTATIC() As Boolean
    Dim p595 As Long

    With ctx
        If pvMatchString("static") Then             ' "static"
            .BufPos = .BufPos + 6
            p595 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p595
                Call Parse_
                Call pvSetAdvance
                ParseSTATIC = True
            End Select
        End If
    End With
End Function

Private Function ParseCAIRO_PUBLIC() As Boolean
    Dim p605 As Long

    With ctx
        If pvMatchString("cairo_public") Then       ' "cairo_public"
            .BufPos = .BufPos + 12
            p605 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p605
                Call Parse_
                Call pvSetAdvance
                ParseCAIRO_PUBLIC = True
            End Select
        End If
    End With
End Function

Private Function ParseCAIRO_WARN() As Boolean
    Dim p610 As Long

    With ctx
        If pvMatchString("cairo_warn") Then         ' "cairo_warn"
            .BufPos = .BufPos + 10
            p610 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p610
                Call Parse_
                Call pvSetAdvance
                ParseCAIRO_WARN = True
            End Select
        End If
    End With
End Function

Private Function ParseINLINE() As Boolean
    Dim p600 As Long

    With ctx
        If pvMatchString("inline") Then             ' "inline"
            .BufPos = .BufPos + 6
            p600 = .BufPos
            Select Case .BufData(.BufPos)
            Case 97 To 122, 65 To 90, 95, 48 To 57, 35 ' [a-zA-Z_0-9#]
                '--- do nothing
            Case Else
                .BufPos = p600
                Call Parse_
                Call pvSetAdvance
                ParseINLINE = True
            End Select
        End If
    End With
End Function

Private Function ParsePREPRO() As Boolean
    With ctx
        If .BufData(.BufPos) = 35 Then              ' "#"
            .BufPos = .BufPos + 1
            Do
                Select Case .BufData(.BufPos)
                Case 13, 10                         ' [\r\n]
                    Exit Do
                Case Else
                    If .BufPos < .BufSize Then
                        .BufPos = .BufPos + 1
                    Else
                        Exit Do
                    End If
                End Select
            Loop
            If ParseNL() Then
                Call pvSetAdvance
                ParsePREPRO = True
            End If
        End If
    End With
End Function

Private Sub pvImplAction(ByVal eAction As UcsParserActionsEnum, ByVal lOffset As Long, ByVal lSize As Long)
    Dim oJson As Object
    Dim oEl As Object
    Dim vElem As Variant
    
    With ctx
    Select Case eAction
    Case ucsAct_3_StmtList
           Set ctx.VarResult = ctx.VarStack(ctx.VarPos - 1)
    Case ucsAct_2_StmtList
           Set oJson = ctx.VarStack(ctx.VarPos - 1) : JsonItem(oJson, -1) = ctx.VarStack(ctx.VarPos - 2)
    Case ucsAct_1_StmtList
           JsonItem(oJson, -1) = Empty
                                                                            Set ctx.VarStack(ctx.VarPos - 1) = oJson
    Case ucsAct_1_TypedefDecl
           JsonItem(oJson, "Tag") = "TypedefDecl"
                                                                            JsonItem(oJson, "Name") = ctx.VarStack(ctx.VarPos - 2)
                                                                            JsonItem(oJson, "Type") = ctx.VarStack(ctx.VarPos - 1)
                                                                            Set ctx.VarResult = oJson
    Case ucsAct_1_TypedefCallback
           JsonItem(oJson, "Tag") = "TypedefCallback"
                                                                            JsonItem(oJson, "Name") = ctx.VarStack(ctx.VarPos - 2)
                                                                            JsonItem(oJson, "Type") = ctx.VarStack(ctx.VarPos - 1)
                                                                            JsonItem(oJson, "Params") = ctx.VarStack(ctx.VarPos - 3)
                                                                            Set ctx.VarResult = oJson
    Case ucsAct_3_EnumDecl
           Set oJson = ctx.VarStack(ctx.VarPos - 2)
                                                                            JsonItem(oJson, "Name") = ctx.VarStack(ctx.VarPos - 1)
                                                                            JsonItem(oJson, "Names") = ctx.VarStack(ctx.VarPos - 5)
                                                                            Set ctx.VarResult = oJson
    Case ucsAct_2_EnumDecl
           JsonItem(oEl, "Name") = ctx.VarStack(ctx.VarPos - 3)
                                                                            JsonItem(oEl, "Value") = zn(CStr(ctx.VarStack(ctx.VarPos - 4)), Empty)
                                                                            Set oJson = ctx.VarStack(ctx.VarPos - 2)
                                                                            JsonItem(oJson, "Items/-1") = oEl
    Case ucsAct_1_EnumDecl
           JsonItem(oJson, "Tag") = "EnumDecl" 
                                                                            JsonItem(oJson, "Items/-1") = Empty
                                                                            Set ctx.VarStack(ctx.VarPos - 2) = oJson
    Case ucsAct_5_StructDecl
           Set oJson = ctx.VarStack(ctx.VarPos - 2)
                                                                            JsonItem(oJson, "Name") = ctx.VarStack(ctx.VarPos - 1)
                                                                            JsonItem(oJson, "Names") = ctx.VarStack(ctx.VarPos - 7)
                                                                            Set ctx.VarResult = oJson
    Case ucsAct_4_StructDecl
           Set oEl = ctx.VarStack(ctx.VarPos - 3)
                                                                            For Each vElem In JsonKeys(oEl, "Items")
                                                                                JsonItem(oJson, "Items/-1") = JsonItem(oEl, "Items/" & vElem)
                                                                            Next
    Case ucsAct_3_StructDecl
           JsonItem(oEl, "Names") = ctx.VarStack(ctx.VarPos - 4)
                                                                            JsonItem(oEl, "Type") = "uintptr_t" 
                                                                            JsonItem(oEl, "PfnType") = ctx.VarStack(ctx.VarPos - 3)
                                                                            JsonItem(oEl, "PfnParams") = ctx.VarStack(ctx.VarPos - 6)
                                                                            Set oJson = ctx.VarStack(ctx.VarPos - 2)
                                                                            JsonItem(oJson, "Items/-1") = oEl
    Case ucsAct_2_StructDecl
           JsonItem(oEl, "Names") = ctx.VarStack(ctx.VarPos - 4)
                                                                            JsonItem(oEl, "Type") = ctx.VarStack(ctx.VarPos - 3)
                                                                            JsonItem(oEl, "ArraySuffixes") = ctx.VarStack(ctx.VarPos - 5)
                                                                            Set oJson = ctx.VarStack(ctx.VarPos - 2)
                                                                            JsonItem(oJson, "Items/-1") = oEl
    Case ucsAct_1_StructDecl
           JsonItem(oJson, "Tag") = "StructDecl" 
                                                                            JsonItem(oJson, "Items/-1") = Empty
                                                                            Set ctx.VarStack(ctx.VarPos - 2) = oJson
    Case ucsAct_1_FunDecl
           JsonItem(oJson, "Tag") = "FunDecl"
                                                                            JsonItem(oJson, "Name") = ctx.VarStack(ctx.VarPos - 2)
                                                                            JsonItem(oJson, "Type") = ctx.VarStack(ctx.VarPos - 1)
                                                                            JsonItem(oJson, "Params") = ctx.VarStack(ctx.VarPos - 3)
                                                                            Set ctx.VarResult = oJson
    Case ucsAct_1_SkipStmt
           JsonItem(oJson, "Tag") = "SkipStmt"
                                                                            JsonItem(oJson, "Text") = Mid$(ctx.Contents, lOffset, lSize)
                                                                            Set ctx.VarResult = oJson
    Case ucsAct_1_Type
         ctx.VarResult = Trim$(Replace(Mid$(ctx.Contents, lOffset, lSize), vbTab, " "))
    Case ucsAct_1_ID
           ctx.VarResult = Mid$(ctx.Contents, lOffset, lSize)
    Case ucsAct_1_TypeUnlimited
         ctx.VarResult = Trim$(Replace(Mid$(ctx.Contents, lOffset, lSize), vbTab, " "))
    Case ucsAct_3_ParamList
           Set ctx.VarResult = ctx.VarStack(ctx.VarPos - 1)
    Case ucsAct_2_ParamList
           Set oJson = ctx.VarStack(ctx.VarPos - 1) : JsonItem(oJson, -1) = ctx.VarStack(ctx.VarPos - 2)
    Case ucsAct_1_ParamList
           JsonItem(oJson, -1) = ctx.VarStack(ctx.VarPos - 1) : Set ctx.VarStack(ctx.VarPos - 1) = oJson
    Case ucsAct_4_EnumValue
           Set oJson = ctx.VarStack(ctx.VarPos - 1) : ctx.VarResult = ConcatCollection(oJson, " ")
    Case ucsAct_3_EnumValue
           ctx.VarStack(ctx.VarPos - 1).Add ctx.VarStack(ctx.VarPos - 2)
    Case ucsAct_2_EnumValue
           ctx.VarStack(ctx.VarPos - 1).Add ctx.VarStack(ctx.VarPos - 2)
    Case ucsAct_1_EnumValue
           Set ctx.VarStack(ctx.VarPos - 1) = New Collection
    Case ucsAct_1_EMPTY
           ctx.VarResult = Mid$(ctx.Contents, lOffset, lSize)
    Case ucsAct_3_IDList
           Set ctx.VarResult = ctx.VarStack(ctx.VarPos - 1)
    Case ucsAct_2_IDList
           Set oJson = ctx.VarStack(ctx.VarPos - 1) : JsonItem(oJson, -1) = ctx.VarStack(ctx.VarPos - 2)
    Case ucsAct_1_IDList
           JsonItem(oJson, -1) = ctx.VarStack(ctx.VarPos - 1) : Set ctx.VarStack(ctx.VarPos - 1) = oJson
    Case ucsAct_3_StuctMemList
           Set ctx.VarResult = ctx.VarStack(ctx.VarPos - 1)
    Case ucsAct_2_StuctMemList
           Set oJson = ctx.VarStack(ctx.VarPos - 1) : JsonItem(oJson, -1) = ctx.VarStack(ctx.VarPos - 2)
    Case ucsAct_1_StuctMemList
           JsonItem(oJson, -1) = ctx.VarStack(ctx.VarPos - 1) : Set ctx.VarStack(ctx.VarPos - 1) = oJson
    Case ucsAct_3_ArraySuffixList
           Set ctx.VarResult = ctx.VarStack(ctx.VarPos - 1)
    Case ucsAct_2_ArraySuffixList
           Set oJson = ctx.VarStack(ctx.VarPos - 1) : JsonItem(oJson, -1) = ctx.VarStack(ctx.VarPos - 2)
    Case ucsAct_1_ArraySuffixList
           JsonItem(oJson, -1) = ctx.VarStack(ctx.VarPos - 1) : Set ctx.VarStack(ctx.VarPos - 1) = oJson
    Case ucsAct_1_Param
           JsonItem(oJson, "Type") = ctx.VarStack(ctx.VarPos - 1)
                                                                            JsonItem(oJson, "Name") = ctx.VarStack(ctx.VarPos - 2)
                                                                            JsonItem(oJson, "ArraySuffix") = zn(CStr(ctx.VarStack(ctx.VarPos - 3)), Empty)
                                                                            Set ctx.VarResult = oJson
    Case ucsAct_1_ArraySuffix
           ctx.VarResult = Mid$(ctx.Contents, lOffset, lSize)
    Case ucsAct_1_EnumValueToken
           ctx.VarResult = Mid$(ctx.Contents, lOffset, lSize)
    End Select
    End With
End Sub

