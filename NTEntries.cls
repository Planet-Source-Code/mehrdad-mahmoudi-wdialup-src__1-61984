VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RASEntriesNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'     This class is written by w0lf but original functions is written by SNE
'   of Russia in a Module!
'     I've converted module of SNE to a class to easily use!
'
'     Don't forget to add VBHWRasDialParams Type to a public module or ...
'
'     sOptions details:
'       its a 6 character string that shows us which details of entry will be
'      shown:
'           1st: Callback number
'           2nd: Domain
'           3th: Phone number
'           4th: Username
'           5th: Password
'           6th: Connection ID
'
'     Special thanks to SNE (website: http://vbnet.ru) !!!
'
'     My E-mail: sir.w0lf@gmail.com
'     ICQ: 231-707-516
'
'   Regards,
'   w0lf
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§




'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ API Calls §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal lpcstr As String, ByVal lpcstr As String, ByRef lprasentrynamea As Any, ByRef lpdword As Long, ByRef lpdword As Long) As Long
Private Declare Function RasGetEntryDialParams Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" (ByVal lpcstr As String, ByRef lprasdialparamsa As Byte, ByRef lpbool As Long) As Long
Private Declare Function RasGetEntryProperties Lib "rasapi32.dll" Alias "RasGetEntryPropertiesA" (ByVal lpszPhonebook As String, ByVal lpszEntry As String, ByRef lpRasEntry As Any, ByRef lpdwEntryInfoSize As Long, ByRef lpbDeviceInfo As Any, ByRef lpdwDeviceInfoSize As Long) As Long

Private Declare Function LsaOpenPolicy Lib "ADVAPI32.dll" (ByRef SystemName As Long, ByRef ObjectAttributes As LSA_OBJECT_ATTRIBUTES, ByVal DesiredAccess As Long, ByRef PolicyHandle As Long) As Long
Private Declare Function LsaRetrievePrivateData Lib "ADVAPI32.dll" (ByVal PolicyHandle As Long, ByRef KeyName As LSA_UNICODE_STRING, ByVal PrivateData As Long) As Long
Private Declare Function LsaClose Lib "ADVAPI32.dll" (ByVal ObjectHandle As Long) As Long
Private Declare Function LsaFreeMemory Lib "ADVAPI32.dll" (ByVal Buffer As Long) As Long

Private Declare Function MultiByteToWideChar Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As String, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Declare Function ConvertSidToStringSid Lib "ADVAPI32.dll" Alias "ConvertSidToStringSidA" (ByVal Sid As String, ByRef lpStringSid As Long) As Long
Private Declare Function IsValidSid Lib "ADVAPI32.dll" (ByRef pSid As Any) As Long
Private Declare Function LookupAccountName Lib "ADVAPI32.dll" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, ByVal Sid As String, ByRef cbSid As Long, ByVal ReferencedDomainName As String, ByRef cbReferencedDomainName As Long, ByRef peUse As Long) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal pszPath As String, ByVal csidl As Long, ByVal fCreate As Long) As Long
Private Declare Function GetUserName Lib "ADVAPI32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GlobalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function LocalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Consts §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Const LMEM_ZEROINIT                     As Long = &H40
Private Const POLICY_GET_PRIVATE_INFORMATION    As Long = &H4&
Private Const MAX_PATH                          As Long = 260
Private Const ERROR_BUFFER_TOO_SMALL            As Long = 603
Private Const VER_PLATFORM_WIN32_NT             As Long = &H2
Private Const RAS_MaxEntryName                  As Long = &H100
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Types §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Type OSVERSIONINFO
    dwOSVersionInfoSize          As Long
    dwMajorVersion               As Long
    dwMinorVersion               As Long
    dwBuildNumber                As Long
    dwPlatformId                 As Long
    szCSDVersion                 As String * 128
End Type

Private Type LSA_UNICODE_STRING
  Length                        As Integer
  MaximumLength                 As Integer
  Buffer                        As Long
End Type
Private Type LSA_OBJECT_ATTRIBUTES
    Length                      As Long
    RootDirectory               As Long
    ObjectName                  As Long
    Attributes                  As Long
    SecurityDescriptor          As Long
    SecurityQualityOfService    As Long
End Type

Private Type RASENTRYNAME
   dwSize                        As Long
   szEntryName(RAS_MaxEntryName) As Byte
End Type

Private strOutput() As VBHWRasDialParams
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Coded by SNE §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Function GetRasEntrLst(ByRef soutArray() As String) As Long
    Dim rsname()  As RASENTRYNAME
    Dim lSize     As Long

    ReDim rsname(0) As RASENTRYNAME

    rsname(0).dwSize = &H108
    lSize = rsname(0).dwSize
                                                                    
    Select Case RasEnumEntries(vbNullString, vbNullString, rsname(0), lSize, GetRasEntrLst)
        Case Is = ERROR_BUFFER_TOO_SMALL
            ReDim rsname(GetRasEntrLst - vbNull)

            rsname(0).dwSize = &H108
            lSize = GetRasEntrLst * rsname(0).dwSize

            If RasEnumEntries(vbNullString, vbNullString, rsname(0), lSize, GetRasEntrLst) Then GoTo err
        Case Is = 0&
        Case Else
            GoTo err
    End Select

    ReDim soutArray(GetRasEntrLst - vbNull)
    For lSize = 0 To GetRasEntrLst - vbNull
        soutArray(lSize) = StrConv(rsname(lSize).szEntryName, vbUnicode)
        soutArray(lSize) = Left$(soutArray(lSize), InStr(vbNull, soutArray(lSize), vbNullChar, vbBinaryCompare) - vbNull)
    Next

Exit Function
err:
    Erase soutArray
    GetRasEntrLst = 0&

End Function
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Coded by SNE §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Function ProcessLSABuffer(ByRef sBuffer As String, ByVal BufferLen As Integer) As String
    Dim i       As Long, sPos       As Long, wchar      As Integer
    Dim BookID  As String, sTmpBuff As String

    For i = 0 To BufferLen Step 2&
        Call CopyMemory(wchar, ByVal Mid$(sBuffer, i + 1&, 2&), 2&)

        If wchar = 0& Then
            sPos = sPos + vbNull

            Select Case sPos
                Case Is = 1&
                    BookID = sTmpBuff

                Case Is = 7&
                    If Len(sTmpBuff) Then ProcessLSABuffer = ProcessLSABuffer & BookID & vbNullChar & _
                                                                                sTmpBuff & vbNullChar & vbNullChar
            End Select
            sTmpBuff = vbNullString

        Else
            sTmpBuff = sTmpBuff & ChrW$(wchar)
            If sPos = 9& Then sPos = 0&: BookID = vbNullString
        End If
    Next
                                                                   
    ProcessLSABuffer = Left$(ProcessLSABuffer, Len(ProcessLSABuffer) - 2&)
End Function
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Coded by SNE §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Function GetLocalSid() As String
    Dim UserName As String, UserNameSize As Long, _
        Sid      As String, SidSize       As Long, _
        Domain   As String, DomainSize    As Long, _
                            snu           As Long

    UserName = String$(255, 0&)
    Sid = String$(255, 0&)
    Domain = String$(255, 0&)
    GetLocalSid = String$(255, 0&)
    
    UserNameSize = 255
    SidSize = 255
    DomainSize = 255

    If GetUserName(UserName, UserNameSize) = 0& Then Exit Function
    UserName = Left$(UserName, InStr(vbNull, UserName, vbNullChar, vbBinaryCompare) - vbNull)
                                                                    
    Call LookupAccountName(vbNullString, UserName, Sid, SidSize, Domain, DomainSize, snu)
    If IsValidSid(ByVal Sid) = 0& Then Exit Function

    Call ConvertSidToStringSid(Sid, snu)
    Call CopyMemory(ByVal GetLocalSid, ByVal snu, 255)

    GetLocalSid = Left$(GetLocalSid, InStr(vbNull, GetLocalSid, vbNullChar, vbBinaryCompare) - vbNull)
    Call GlobalFree(snu)
End Function
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Coded by SNE §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Function GetLsaData(ByVal Policy As Long, ByVal KeyName As String, ByRef lpOutData As LSA_UNICODE_STRING) As Boolean
    Dim LsaObjectAttribs    As LSA_OBJECT_ATTRIBUTES
    Dim LsaHandle           As Long
    Dim LsaKeyName          As LSA_UNICODE_STRING
    Dim hMem                As Long

    If Not LsaOpenPolicy(ByVal 0&, LsaObjectAttribs, Policy, LsaHandle) = 0 Then Exit Function

    LsaKeyName.Length = LenB(KeyName)
    LsaKeyName.MaximumLength = LsaKeyName.Length + &H2
                                                                    
    LsaKeyName.Buffer = LocalAlloc(LMEM_ZEROINIT, &HFF)
    Call MultiByteToWideChar(0&, 0&, KeyName, Len(KeyName), LsaKeyName.Buffer, LsaKeyName.MaximumLength)

    If Not LsaRetrievePrivateData(LsaHandle, LsaKeyName, VarPtr(hMem)) = 0& Then
        Call LsaClose(LsaHandle)
        Exit Function
    Else
        Call CopyMemory(lpOutData, ByVal hMem, Len(lpOutData))
        GetLsaData = True
    End If

    Call LsaFreeMemory(LsaKeyName.Buffer)
    Call LsaClose(LsaHandle)
End Function
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Coded by SNE §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Function MakePhoneBookPath(ByVal sValue As String) As String
    MakePhoneBookPath = Left$(sValue, InStr(vbNull, sValue, vbNullChar, vbBinaryCompare) - vbNull)
    
    If Not Right$(MakePhoneBookPath, vbNull) = "\" Then MakePhoneBookPath = MakePhoneBookPath & "\"
    MakePhoneBookPath = MakePhoneBookPath & "Microsoft\Network\Connections\pbk\rasphone.pbk"
End Function
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Coded by SNE §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Function Byte2Str(bPos As Byte, lngMaxLen As Long) As String
    Byte2Str = String$(lngMaxLen + vbNull, &H0)
    Call CopyMemory(ByVal Byte2Str, bPos, lngMaxLen)

    Byte2Str = Left$(Byte2Str, InStr(Byte2Str, vbNullChar) - vbNull)
End Function
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Coded by SNE §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Function StrToUTF8(inString As String) As String
        Dim hMemLock1   As Long, hMemLock2  As Long
        Dim iStrSize    As Integer
        
        hMemLock1 = LocalAlloc(LMEM_ZEROINIT, &H100)
        hMemLock2 = LocalAlloc(LMEM_ZEROINIT, &H100)

        iStrSize = MultiByteToWideChar(0&, 0&, inString, &HFFFF, hMemLock1, &H100)
        iStrSize = WideCharToMultiByte(65001, 0&, hMemLock1, iStrSize, hMemLock2, &H100, 0&, 0&)  ' CP_UTF8

        StrToUTF8 = String$(iStrSize, 0&)
        Call CopyMemory(ByVal StrToUTF8, ByVal hMemLock2, iStrSize)

        Call LocalFree(hMemLock1)
        Call LocalFree(hMemLock2)
End Function
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Coded by w0lf §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Function GetLSAPasswords(sIndex As Integer) As String                                         ' Ïîëó÷åíèå ïàðîëåé, çàòåì îáðàáîòêà äàííûõ
    Dim pData As LSA_UNICODE_STRING
    Dim sBuffer As String
    
    If sIndex = 0 Then
        If GetLsaData(POLICY_GET_PRIVATE_INFORMATION, "RasDialParams!" & GetLocalSid & "#0", pData) Then GoTo sLPart
        If GetLsaData(POLICY_GET_PRIVATE_INFORMATION, "L$_RasDefaultCredentials#0", pData) Then GoTo sLPart
    Else
        If GetLsaData(POLICY_GET_PRIVATE_INFORMATION, "L$_RasDefaultCredentials#0", pData) Then GoTo sLPart
        If GetLsaData(POLICY_GET_PRIVATE_INFORMATION, "RasDialParams!" & GetLocalSid & "#0", pData) Then GoTo sLPart
    End If

Exit Function
sLPart:
    sBuffer = String$(pData.MaximumLength, 0&)
    Call CopyMemory(ByVal sBuffer, ByVal pData.Buffer, pData.Length)
                                                                          
    GetLSAPasswords = ProcessLSABuffer(sBuffer, pData.Length)
    Call LsaFreeMemory(pData.Buffer)
End Function
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§ Modified by w0lf to work on all systems §§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Function GetRasEntries2k(outArray() As VBHWRasDialParams, sIndex As Integer) As Long
    Dim RasArray()  As String, i            As Long, j As Long
    Dim btArray()   As Byte, DialParamsUID  As Long
    Dim Book1       As String, Book2        As String
    Dim LSA_Pass()  As String, sTempBuffer(1)  As String
    Dim osi         As OSVERSIONINFO

    GetRasEntries2k = GetRasEntrLst(RasArray) - vbNull
    ReDim outArray(GetRasEntries2k)

    osi.dwOSVersionInfoSize = Len(osi)
    Call GetVersionEx(osi)

    Book1 = String$(MAX_PATH + vbNull, 0&)
    Book2 = String$(MAX_PATH + vbNull, 0&)

    If (osi.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osi.dwMajorVersion >= 5) Then
        If (SHGetSpecialFolderPath(0&, Book1, &H1A, False)) Then _
            Book1 = MakePhoneBookPath(Book1)
        If (SHGetSpecialFolderPath(0, Book2, &H23, False)) Then _
            Book2 = MakePhoneBookPath(Book2)

        sTempBuffer(sIndex) = GetLSAPasswords(sIndex)
        If Len(sTempBuffer(sIndex)) = 0& Then GoTo err

        LSA_Pass = Split(sTempBuffer(sIndex), vbNullChar & vbNullChar)
    End If

    For i = 0 To GetRasEntries2k
        ReDim btArray(1060&)

        Call CopyMemory(btArray(0), 1060&, 4&)
        Call CopyMemory(btArray(4), ByVal RasArray(i), Len(RasArray(i)))
        Call RasGetEntryDialParams(vbNullString, btArray(0), vbNull)

        If (osi.dwPlatformId = VER_PLATFORM_WIN32_NT) And (osi.dwMajorVersion >= 5) Then
            sTempBuffer(sIndex) = StrToUTF8(RasArray(i))

            DialParamsUID = GetPrivateProfileInt(RasArray(i), "DialParamsUID", 0&, Book1)
            If DialParamsUID = 0 Then _
                DialParamsUID = GetPrivateProfileInt(RasArray(i), "DialParamsUID", 0&, Book2)
            If DialParamsUID = 0 Then _
                DialParamsUID = GetPrivateProfileInt(sTempBuffer(sIndex), "DialParamsUID", 0&, Book1)
            If DialParamsUID = 0 Then _
                DialParamsUID = GetPrivateProfileInt(sTempBuffer(sIndex), "DialParamsUID", 0&, Book2)

            If Not DialParamsUID = 0 Then
                For j = 0 To UBound(LSA_Pass)
                    If Len(LSA_Pass(j)) Then
                        If DialParamsUID = Val(Split(LSA_Pass(j), vbNullChar)(0)) Then
                            outArray(i).ConnectionID = DialParamsUID

                            If InStr(vbNull, LSA_Pass(j), vbNullChar) Then
                                outArray(i).Password = Split(LSA_Pass(j), vbNullChar)(1)
                            End If
                        End If
                    End If
                Next
            End If

            With outArray(i)
                .EntryName = RasArray(i)
                .CallbackNumber = Byte2Str(btArray(390), 129)
                .UserName = Byte2Str(btArray(519), 257)
                .Domain = Byte2Str(btArray(1033), 16)

                Call RasGetEntryProperties(vbNullString, vbNullString, ByVal 0&, j, ByVal 0&, ByVal 0&)
                If Not RasGetEntryProperties(vbNullString, .EntryName, ByVal 0&, DialParamsUID, ByVal 0&, ByVal 0&) = 603& Then GoTo err

                ReDim btArray(DialParamsUID - vbNull) As Byte
                Call CopyMemory(btArray(0), j, 4&)
                If Not RasGetEntryProperties(vbNullString, .EntryName, btArray(0), DialParamsUID, ByVal 0&, ByVal 0&) = 0& Then GoTo err

                .PhoneNumber = Byte2Str(btArray(27), 129)

            End With
        End If
    Next

Exit Function
err:
    Erase outArray
    GetRasEntries2k = &HFFFF
End Function
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§ Modified by w0lf §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Private Function GetRasEntriesNT(ByRef sArray() As VBHWRasDialParams) As Boolean
Dim i As Long, RasDataCU() As VBHWRasDialParams, RasDataAU() As VBHWRasDialParams
Dim RasData() As VBHWRasDialParams
ReDim RasDataCU(0 To 255)
ReDim RasDataAU(0 To 255)
Dim Num_Items As Byte

For i = 0 To GetRasEntries2k(RasData, 0)
    RasDataCU(i).CallbackNumber = RasData(i).CallbackNumber
    RasDataCU(i).ConnectionID = RasData(i).ConnectionID
    RasDataCU(i).Domain = RasData(i).Domain
    RasDataCU(i).EntryName = RasData(i).EntryName
    RasDataCU(i).Password = RasData(i).Password
    RasDataCU(i).PhoneNumber = RasData(i).PhoneNumber
    RasDataCU(i).UserName = RasData(i).UserName
Next

For i = 0 To GetRasEntries2k(RasData, 1)
    RasDataAU(i).CallbackNumber = RasData(i).CallbackNumber
    RasDataAU(i).ConnectionID = RasData(i).ConnectionID
    RasDataAU(i).Domain = RasData(i).Domain
    RasDataAU(i).EntryName = RasData(i).EntryName
    RasDataAU(i).Password = RasData(i).Password
    RasDataAU(i).PhoneNumber = RasData(i).PhoneNumber
    RasDataAU(i).UserName = RasData(i).PhoneNumber
Next

Num_Items = i - 1

ReDim sArray(Num_Items)

For i = 0 To Num_Items
    
    sArray(i).EntryName = RasDataCU(i).EntryName
    sArray(i).CallbackNumber = RasDataCU(i).CallbackNumber
    sArray(i).Domain = RasDataCU(i).Domain
    sArray(i).PhoneNumber = RasDataCU(i).PhoneNumber
    sArray(i).UserName = RasDataCU(i).UserName
    
    If RasDataCU(i).ConnectionID = "0" And RasDataCU(i).Password = "" And Not RasDataAU(i).ConnectionID = "0" And Not RasDataAU(i).Password = "" Then
        sArray(i).Password = RasDataAU(i).Password
        sArray(i).ConnectionID = RasDataAU(i).ConnectionID
    Else
        sArray(i).Password = RasDataCU(i).Password
        sArray(i).ConnectionID = RasDataCU(i).ConnectionID
    End If
Next i
GetRasEntriesNT = True
End Function
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§ Coded by w0lf §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
'§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
Public Function RASEntries(sOptions As String) As String
    Dim sOut As String
    sOut = "§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§" & vbCrLf
    sOut = sOut & vbTab & "    Originally coded by SNE and modified by w0lf" & vbCrLf
    sOut = sOut & "§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§" + vbCrLf + vbCrLf
    For i = LBound(strOutput) To UBound(strOutput)
        sOut = sOut & "Entry Name: " & strOutput(i).EntryName & vbCrLf
        
        If CBool(Mid(sOptions, 1, 1)) Then _
            sOut = sOut & vbTab & "CallbackNumber:" & vbTab & strOutput(i).CallbackNumber & vbCrLf
        If CBool(Mid(sOptions, 2, 1)) Then _
            sOut = sOut & vbTab & "Domain:" & vbTab & vbTab & strOutput(i).Domain & vbCrLf
        If CBool(Mid(sOptions, 3, 1)) Then _
            sOut = sOut & vbTab & "PhoneNumber:" & vbTab & strOutput(i).PhoneNumber & vbCrLf
        If CBool(Mid(sOptions, 4, 1)) Then _
            sOut = sOut & vbTab & "UserName:" & vbTab & strOutput(i).UserName & vbCrLf
        If CBool(Mid(sOptions, 5, 1)) Then _
            sOut = sOut & vbTab & "Password:" & vbTab & strOutput(i).Password & vbCrLf
        If CBool(Mid(sOptions, 6, 1)) Then _
            sOut = sOut & vbTab & "ConnectionID:" & vbTab & strOutput(i).ConnectionID & vbCrLf
        
        sOut = sOut + vbCrLf
    Next
    sOut = sOut & "§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§"
    RASEntries = sOut
End Function

Private Function cLoadEntries()
    GetRasEntriesNT strOutput
End Function

Private Sub Class_Initialize()
    cLoadEntries
End Sub

