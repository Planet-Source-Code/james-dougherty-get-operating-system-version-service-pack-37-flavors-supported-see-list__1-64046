VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AP_OperatingSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//::|ñññññññññññññññññññññññññññññññññññññññññññññññññññññññññññ|::
'//::|ññ ----------------------------------------------------- ññ|::
'//::|ññ|       ~ A R I E L  P R O D U C T I O N S ~          |ññ|::
'//::|ññ|             ~ ALL RIGHTS RESERVED ~                 |ññ|::
'//::|ññ ----------------------------------------------------- ññ|::
'//::|ññ| ~ PROGRAM DATA ~                           /------/ |ññ|::
'//::|ññ|                                          /------/ | |ññ|::
'//::|ññ| PROGRAMMER: James Dougherty              | |    | | |ññ|::
'//::|ññ| COPYRIGHT : ©2005 Ariel Productions      | |    | | |ññ|::
'//::|ññ| TYPE      : N/A                          | /----|-/ |ññ|::
'//::|ññ|                                          /------/   |ññ|::
'//::|ññ ----------------------------------------------------- ññ|::
'//::|ññ| ~ SOURCE DATA ~                                     |ññ|::
'//::|ññ|                                                     |ññ|::
'//::|ññ| FILENAME :: AP_OperatingSystemV2.cls                |ññ|::
'//::|ññ| DATE     :: 01.15.06                                |ññ|::
'//::|ññ| VERSION  :: 1.0                                     |ññ|::
'//::|ññ|                                                     |ññ|::
'//::|ññ| HISTORY DATA                                        |ññ|::
'//::|ññ| |                                                   |ññ|::
'//::|ññ| |-N/A                                               |ññ|::
'//::|ññ| |                                                   |ññ|::
'//::|ññ|                                                     |ññ|::
'//::|ññ| BUGS                                                |ññ|::
'//::|ññ| |                                                   |ññ|::
'//::|ññ| |-N/A                                               |ññ|::
'//::|ññ| |                                                   |ññ|::
'//::|ññ|                                                     |ññ|::
'//::|ññ ----------------------------------------------------- ññ|::
'//::|ñññññññññññññññññññññññññññññññññññññññññññññññññññññññññññ|::
Option Explicit

Public Enum OSPLATFORM_PRODUCT
    OSPP_UNKNOWN = 0
    OSPP_WINDOWS_32S = 1
    OSPP_WINDOWS_95 = 2
    OSPP_WINDOWS_95_SP1 = 3
    OSPP_WINDOWS_95_OSR2 = 4
    OSPP_WINDOWS_95_OSR25 = 5
    OSPP_WINDOWS_98 = 6
    OSPP_WINDOWS_98_SP1 = 7
    OSPP_WINDOWS_98_SE = 8
    OSPP_WINDOWS_ME = 9
    OSPP_WINDOWS_2000 = 10
    OSPP_WINDOWS_2000_PROFESSIONAL = 11
    OSPP_WINDOWS_2000_SERVER = 12
    OSPP_WINDOWS_2000_SERVER_DATACENTER = 13
    OSPP_WINDOWS_2000_SERVER_ADVANCED = 14
    OSPP_WINDOWS_2003 = 15
    OSPP_WINDOWS_2003_R2 = 16
    OSPP_WINDOWS_2003_ENTERPRISE = 17
    OSPP_WINDOWS_2003_DATACENTER = 18
    OSPP_WINDOWS_2003_WEB = 19
    OSPP_WINDOWS_2003_X64 = 20
    OSPP_WINDOWS_2003_ENTERPRISE_X64 = 21
    OSPP_WINDOWS_2003_DATACENTER_X64 = 22
    OSPP_WINDOWS_2003_ENTERPRISE_ITANIUM = 23
    OSPP_WINDOWS_2003_DATACENTER_ITANIUM = 24
    OSPP_WINDOWS_XP = 25
    OSPP_WINDOWS_XP_X64 = 26
    OSPP_WINDOWS_XP_HOME = 27
    OSPP_WINDOWS_XP_PROFESSIONAL = 28
    OSPP_WINDOWS_NT = 29
    OSPP_WINDOWS_NT_351 = 30
    OSPP_WINDOWS_NT_400 = 31
    OSPP_WINDOWS_NT_400_WORKSTATION = 32
    OSPP_WINDOWS_NT_400_SERVER_ENTERPRISE = 33
    OSPP_WINDOWS_NT_400_SERVER = 34
    OSPP_WINDOWS_VISTA = 35
    OSPP_WINDOWS_SERVER_LONGHORN = 36
    OSPP_WINDOWS_BLACKCOMB = 37
End Enum

Public Enum OSPLATFORM_PRODUCT_TYPE
    OSPPT_WORKSTATION = 1
    OSPPT_DOMAIN_CONTROLLER = 2
    OSPPT_SERVER = 3
End Enum

Public Enum OSPLATFORM_SUITE
    OSPS_SMALLBUSINESS = 1
    OSPS_ENTERPRISE = 2
    OSPS_TERMINAL = 16
    OSPS_DATACENTER = 128
    OSPS_SINGLEUSERTS = 256
    OSPS_PERSONAL = 512
    OSPS_BLADE = 1024
End Enum

Public Enum PROCESSOR_ARCHITECTURE
    PA_UNKNOWN = -1
    PA_INTEL = 0
    PA_IA64 = 6
    PA_AMD64 = 9
    PA_IA32_ON_WIN64 = 9
End Enum

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String
End Type

Private Type SYSTEM_INFO
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private m_Aquired As Boolean
Private m_Extended As Boolean
Private m_OSVersion As OSVERSIONINFOEX
Private m_SystemInformation As SYSTEM_INFO

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFOEX) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Sub GetNativeSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Private Sub Class_Initialize()
    
    Call InitializeSystem

End Sub

Private Function InitializeSystem() As Boolean
    Dim OSVersionLength As OSVERSIONINFO

    If m_Aquired Then
        InitializeSystem = True
        Exit Function
    End If

    m_Aquired = True
    m_Extended = True
    Call GetNativeSystemInfo(m_SystemInformation)
    DoEvents

    m_OSVersion.dwOSVersionInfoSize = Len(m_OSVersion)
    If GetVersionEx(m_OSVersion) = 0 Then
        m_Extended = False
        m_OSVersion.dwOSVersionInfoSize = Len(OSVersionLength)
        If GetVersionEx(m_OSVersion) = 0 Then
            m_Aquired = False
            Call MsgBox("Unable to obtain the system information" & vbCrLf & vbCrLf & "This class will be disabled", vbCritical Or vbOKOnly, "Operating System Version")
            Exit Function
        End If
    End If
    InitializeSystem = True

End Function

Public Function Platform_IsWin32s() As Boolean

    Platform_IsWin32s = IIf(m_OSVersion.dwPlatformId = 0 And m_Aquired, True, False)

End Function

Public Function Platform_IsWindows() As Boolean

    Platform_IsWindows = IIf(m_OSVersion.dwPlatformId = 1, True, False)

End Function

Public Function Platform_IsNT() As Boolean

    Platform_IsNT = IIf(m_OSVersion.dwPlatformId = 2, True, False)

End Function

Public Function Platform_MajorVersion() As Long

    Platform_MajorVersion = m_OSVersion.dwMajorVersion

End Function

Public Function Platform_MinorVersion() As Long

    Platform_MinorVersion = m_OSVersion.dwMinorVersion

End Function

Public Function Platform_BuildNumber() As Long

    Platform_BuildNumber = (m_OSVersion.dwBuildNumber And -1)

End Function

Public Function Platform_ProductType() As OSPLATFORM_PRODUCT_TYPE

    Platform_ProductType = m_OSVersion.wProductType

End Function

Public Function Platform_ProductTypeDescription() As String

    Select Case Platform_ProductType()
        Case OSPPT_WORKSTATION:       Platform_ProductTypeDescription = "Workstation"
        Case OSPPT_DOMAIN_CONTROLLER: Platform_ProductTypeDescription = "Domain Controller"
        Case OSPPT_SERVER:            Platform_ProductTypeDescription = "Server"
        Case Else:                    Platform_ProductTypeDescription = "Unknown"
    End Select

End Function

Public Function Platform_ProductSuite() As OSPLATFORM_SUITE

    If m_OSVersion.wSuiteMask And &H1& Then
        Platform_ProductSuite = OSPS_SMALLBUSINESS
    ElseIf m_OSVersion.wSuiteMask And &H2& Then
        Platform_ProductSuite = OSPS_ENTERPRISE
    ElseIf m_OSVersion.wSuiteMask And &H10& Then
        Platform_ProductSuite = OSPS_TERMINAL
    ElseIf m_OSVersion.wSuiteMask And &H80& Then
        Platform_ProductSuite = OSPS_DATACENTER
    ElseIf m_OSVersion.wSuiteMask And &H100& Then
        Platform_ProductSuite = OSPS_SINGLEUSERTS
    ElseIf m_OSVersion.wSuiteMask And &H200& Then
        Platform_ProductSuite = OSPS_PERSONAL
    ElseIf m_OSVersion.wSuiteMask And &H400& Then
        Platform_ProductSuite = OSPS_BLADE
    End If

End Function

Public Function Platform_ProductSuiteDescription() As String

    Select Case Platform_ProductType()
        Case OSPS_SMALLBUSINESS: Platform_ProductSuiteDescription = "Small Business"
        Case OSPS_ENTERPRISE:    Platform_ProductSuiteDescription = "Enterprise"
        Case OSPS_TERMINAL:      Platform_ProductSuiteDescription = "Terminal Services"
        Case OSPS_DATACENTER:    Platform_ProductSuiteDescription = "Datacenter"
        Case OSPS_SINGLEUSERTS:  Platform_ProductSuiteDescription = "Single User Terminal Server"
        Case OSPS_PERSONAL:      Platform_ProductSuiteDescription = "Personal Terminal Server"
        Case OSPS_BLADE:         Platform_ProductSuiteDescription = "Web Server"
        Case Else:               Platform_ProductSuiteDescription = "Unknown"
    End Select

End Function

Public Function Platform_ServicePack() As String
    Dim TerminatorPosition As Long
    Dim KeyHandle As Long
    
    If Platform_MajorVersion() = 4 And m_OSVersion.szCSDVersion = "Service Pack 6" Then
        If RegOpenKeyEx(&H80000002, "SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Hotfix\\Q246009", 0, &H1, KeyHandle) = 0& Then
            Platform_ServicePack = "Service Pack 6a"
            Call RegCloseKey(KeyHandle)
        Else
            Platform_ServicePack = m_OSVersion.szCSDVersion
        End If
    Else
        Platform_ServicePack = m_OSVersion.szCSDVersion
    End If

    TerminatorPosition = InStr(1, Platform_ServicePack, Chr$(0))
    If TerminatorPosition > 0 Then
        Platform_ServicePack = Left$(Platform_ServicePack, TerminatorPosition - 1)
    End If

End Function

Public Function Platform_Product() As OSPLATFORM_PRODUCT

    If Not (m_Aquired) Then
        Platform_Product = OSPP_UNKNOWN
        Exit Function
    End If

    Select Case m_OSVersion.dwPlatformId
        Case 0: Platform_Product = Platform_GetWin32sProduct()
        Case 1: Platform_Product = Platform_GetWindowsProduct()
        Case 2: Platform_Product = Platform_GetNTProduct()
    End Select

End Function

Public Function Platform_ProductDescription() As String

    Select Case Platform_Product()
        Case OSPP_WINDOWS_32S:                      Platform_ProductDescription = "Microsoft Windows 32S"
        Case OSPP_WINDOWS_95:                       Platform_ProductDescription = "Microsoft Windows 95"
        Case OSPP_WINDOWS_95_SP1:                   Platform_ProductDescription = "Microsoft Windows 95 SP1"
        Case OSPP_WINDOWS_95_OSR2:                  Platform_ProductDescription = "Microsoft Windows 95 OSR2"
        Case OSPP_WINDOWS_95_OSR25:                 Platform_ProductDescription = "Microsoft Windows 95 OSR2.5"
        Case OSPP_WINDOWS_98:                       Platform_ProductDescription = "Microsoft Windows 98"
        Case OSPP_WINDOWS_98_SP1:                   Platform_ProductDescription = "Microsoft Windows 98 SP1"
        Case OSPP_WINDOWS_98_SE:                    Platform_ProductDescription = "Microsoft Windows 98 Second Edition"
        Case OSPP_WINDOWS_ME:                       Platform_ProductDescription = "Microsoft Windows Millennium Edition"
        Case OSPP_WINDOWS_2000:                     Platform_ProductDescription = "Microsoft Windows 2000"
        Case OSPP_WINDOWS_2000_PROFESSIONAL:        Platform_ProductDescription = "Microsoft Windows 2000 Professional"
        Case OSPP_WINDOWS_2000_SERVER:              Platform_ProductDescription = "Microsoft Windows 2000 Server"
        Case OSPP_WINDOWS_2000_SERVER_DATACENTER:   Platform_ProductDescription = "Microsoft Windows 2000 Server Datacenter"
        Case OSPP_WINDOWS_2000_SERVER_ADVANCED:     Platform_ProductDescription = "Microsoft Windows 2000 Server Advanced"
        Case OSPP_WINDOWS_2003:                     Platform_ProductDescription = "Microsoft Windows 2003"
        Case OSPP_WINDOWS_2003_R2:                  Platform_ProductDescription = "Microsoft Windows 2003 R2"
        Case OSPP_WINDOWS_2003_ENTERPRISE:          Platform_ProductDescription = "Microsoft Windows 2003 Enterprise"
        Case OSPP_WINDOWS_2003_DATACENTER:          Platform_ProductDescription = "Microsoft Windows 2003 Datacenter"
        Case OSPP_WINDOWS_2003_WEB:                 Platform_ProductDescription = "Microsoft Windows 2003 Web Edition"
        Case OSPP_WINDOWS_2003_X64:                 Platform_ProductDescription = "Microsoft Windows 2003 X64"
        Case OSPP_WINDOWS_2003_ENTERPRISE_X64:      Platform_ProductDescription = "Microsoft Windows 2003 Enterprise X64"
        Case OSPP_WINDOWS_2003_DATACENTER_X64:      Platform_ProductDescription = "Microsoft Windows 2003 Datacenter X64"
        Case OSPP_WINDOWS_2003_ENTERPRISE_ITANIUM:  Platform_ProductDescription = "Microsoft Windows 2003 Enterprise Itanium"
        Case OSPP_WINDOWS_2003_DATACENTER_ITANIUM:  Platform_ProductDescription = "Microsoft Windows 2003 Datacenter Itanium"
        Case OSPP_WINDOWS_XP:                       Platform_ProductDescription = "Microsoft Windows XP"
        Case OSPP_WINDOWS_XP_X64:                   Platform_ProductDescription = "Microsoft Windows XP X64 Professional"
        Case OSPP_WINDOWS_XP_HOME:                  Platform_ProductDescription = "Microsoft Windows XP Home Edition"
        Case OSPP_WINDOWS_XP_PROFESSIONAL:          Platform_ProductDescription = "Microsoft Windows XP Professional"
        Case OSPP_WINDOWS_NT:                       Platform_ProductDescription = "Microsoft Windows NT"
        Case OSPP_WINDOWS_NT_351:                   Platform_ProductDescription = "Microsoft Windows NT 3.51"
        Case OSPP_WINDOWS_NT_400:                   Platform_ProductDescription = "Microsoft Windows NT 4.00"
        Case OSPP_WINDOWS_NT_400_WORKSTATION:       Platform_ProductDescription = "Microsoft Windows NT 4.00 Workstation"
        Case OSPP_WINDOWS_NT_400_SERVER_ENTERPRISE: Platform_ProductDescription = "Microsoft Windows NT 4.00 Server Enterprise"
        Case OSPP_WINDOWS_NT_400_SERVER:            Platform_ProductDescription = "Microsoft Windows NT 4.00 Server"
        Case OSPP_WINDOWS_VISTA:                    Platform_ProductDescription = "Microsoft Windows Vista"
        Case OSPP_WINDOWS_SERVER_LONGHORN:          Platform_ProductDescription = "Microsoft Windows Server Longhorn"
        Case OSPP_WINDOWS_BLACKCOMB:                Platform_ProductDescription = "Microsoft Windows Blackcomb"
        Case Else:                                  Platform_ProductDescription = "Unknown"
    End Select

End Function

Private Function Platform_GetWin32sProduct() As OSPLATFORM_PRODUCT

    Platform_GetWin32sProduct = OSPP_WINDOWS_32S

End Function

Private Function Platform_GetWindowsProduct() As OSPLATFORM_PRODUCT
    Dim Major As Long
    Dim Minor As Long
    Dim Build As Long
    
    Major = Platform_MajorVersion()
    Minor = Platform_MinorVersion()
    Build = Platform_BuildNumber()
    Platform_GetWindowsProduct = OSPP_UNKNOWN
    
    If Major <> 4 Then
        Exit Function
    End If
    
    Select Case Minor
        Case Is < 10:
            Select Case Left$(UCase(m_OSVersion.szCSDVersion), 1)
                Case "B":
                    Platform_GetWindowsProduct = OSPP_WINDOWS_95_OSR2

                Case "C":
                    Platform_GetWindowsProduct = OSPP_WINDOWS_95_OSR25

                Case Else
                    If Build > 950 And Build < 1080 Then
                        Platform_GetWindowsProduct = OSPP_WINDOWS_95_SP1
                    Else
                        Platform_GetWindowsProduct = OSPP_WINDOWS_95
                    End If
            End Select

        Case 10:
            If Left$(UCase(m_OSVersion.szCSDVersion), 1) = "A" Or Left$(UCase(m_OSVersion.szCSDVersion), 1) = "B" Then
                Platform_GetWindowsProduct = OSPP_WINDOWS_98_SE
            Else
                If Build > 1998 And Build < 2183 Then
                    Platform_GetWindowsProduct = OSPP_WINDOWS_98_SP1
                Else
                    Platform_GetWindowsProduct = OSPP_WINDOWS_98
                End If
            End If

        Case 90:
            Platform_GetWindowsProduct = OSPP_WINDOWS_ME
    End Select

End Function

Private Function Platform_GetNTProduct() As OSPLATFORM_PRODUCT
    Dim Architecture As PROCESSOR_ARCHITECTURE
    Dim ProductType As OSPLATFORM_PRODUCT_TYPE
    Dim ProductSuite As OSPLATFORM_SUITE
    Dim TerminatorPosition As Long
    Dim NTProductType As String
    Dim KeyHandle As Long
    Dim Major As Long
    Dim Minor As Long
    Dim Build As Long

    Major = Platform_MajorVersion()
    Minor = Platform_MinorVersion()
    Build = Platform_BuildNumber()
    ProductType = Platform_ProductType()
    ProductSuite = Platform_ProductSuite()
    Architecture = System_ProcessorArchitecture()

    If Major = 7 And Minor = 0 Then
        Platform_GetNTProduct = OSPP_WINDOWS_BLACKCOMB
    ElseIf Major = 6 And Minor >= 0 Then
        If ProductType = OSPPT_WORKSTATION Then
            Platform_GetNTProduct = OSPP_WINDOWS_VISTA
        Else
            Platform_GetNTProduct = OSPP_WINDOWS_SERVER_LONGHORN
        End If
    ElseIf Major = 5 And Minor = 0 Then
        Platform_GetNTProduct = OSPP_WINDOWS_2000
    ElseIf Major = 5 And Minor = 1 Then
        Platform_GetNTProduct = OSPP_WINDOWS_XP
    ElseIf Major = 5 And Minor = 2 Then
        If GetSystemMetrics(89) Then
            Platform_GetNTProduct = OSPP_WINDOWS_2003_R2
        ElseIf ProductType = OSPPT_WORKSTATION And Architecture = PA_IA64 Then
            Platform_GetNTProduct = OSPP_WINDOWS_XP_X64
        Else
            Platform_GetNTProduct = OSPP_WINDOWS_2003
        End If
    ElseIf Major = 4 Then
        Platform_GetNTProduct = OSPP_WINDOWS_NT_400
    ElseIf Major = 3 Then
        Platform_GetNTProduct = OSPP_WINDOWS_NT_351
    ElseIf Major < 3 Then
        Platform_GetNTProduct = OSPP_WINDOWS_NT
    End If
    
    If m_Extended Then
        If ProductType = OSPPT_WORKSTATION Then
            If Major = 4 Then
                Platform_GetNTProduct = OSPP_WINDOWS_NT_400_WORKSTATION
            ElseIf ProductSuite = OSPS_PERSONAL Then
                Platform_GetNTProduct = OSPP_WINDOWS_XP_HOME
            Else
                If Minor = 0 Then
                    Platform_GetNTProduct = OSPP_WINDOWS_2000_PROFESSIONAL
                ElseIf Minor = 1 Then
                    Platform_GetNTProduct = OSPP_WINDOWS_XP_PROFESSIONAL
                End If
            End If
        ElseIf ProductType = OSPPT_SERVER Or ProductType = OSPPT_DOMAIN_CONTROLLER Then
            If Major = 5 And Minor = 0 Then
                If ProductSuite = OSPS_DATACENTER Then
                    Platform_GetNTProduct = OSPP_WINDOWS_2000_SERVER_DATACENTER
                ElseIf ProductSuite = OSPS_ENTERPRISE Then
                    Platform_GetNTProduct = OSPP_WINDOWS_2000_SERVER_ADVANCED
                Else
                    Platform_GetNTProduct = OSPP_WINDOWS_2000_SERVER
                End If
            ElseIf Major = 5 And Minor = 2 Then
                If Architecture = PA_IA64 Then
                    If ProductSuite = OSPS_DATACENTER Then
                        Platform_GetNTProduct = OSPP_WINDOWS_2003_DATACENTER_ITANIUM
                    ElseIf ProductSuite = OSPS_ENTERPRISE Then
                        Platform_GetNTProduct = OSPP_WINDOWS_2003_ENTERPRISE_ITANIUM
                    End If
                ElseIf Architecture = PA_AMD64 Then
                    If ProductSuite = OSPS_DATACENTER Then
                        Platform_GetNTProduct = OSPP_WINDOWS_2003_DATACENTER_X64
                    ElseIf ProductSuite = OSPS_ENTERPRISE Then
                        Platform_GetNTProduct = OSPP_WINDOWS_2003_ENTERPRISE_X64
                    Else
                        Platform_GetNTProduct = OSPP_WINDOWS_2003_X64
                    End If
                Else
                    If ProductSuite = OSPS_DATACENTER Then
                        Platform_GetNTProduct = OSPP_WINDOWS_2003_DATACENTER
                    ElseIf ProductSuite = OSPS_ENTERPRISE Then
                        Platform_GetNTProduct = OSPP_WINDOWS_2003_ENTERPRISE
                    ElseIf ProductSuite = OSPS_BLADE Then
                        Platform_GetNTProduct = OSPP_WINDOWS_2003_WEB
                    Else
                        Platform_GetNTProduct = OSPP_WINDOWS_2003
                    End If
                End If
            Else
                If ProductSuite = OSPS_ENTERPRISE Then
                    Platform_GetNTProduct = OSPP_WINDOWS_NT_400_SERVER_ENTERPRISE
                Else
                    Platform_GetNTProduct = OSPP_WINDOWS_NT_400_SERVER
                End If
            End If
        End If
    Else
        If RegOpenKeyEx(&H80000002, "SYSTEM\\CurrentControlSet\\Control\\ProductOptions", 0, &H1, KeyHandle) = 0& Then
            NTProductType = String(512, Chr$(0))
            If RegQueryValueEx(KeyHandle, "ProductType", 0, 0, NTProductType, 512) <> 0& Then
                Call RegCloseKey(KeyHandle)
                Exit Function
            End If
            Call RegCloseKey(KeyHandle)

            TerminatorPosition = InStr(1, NTProductType, Chr$(0))
            If TerminatorPosition > 0 Then
                ProductType = Left$(NTProductType, TerminatorPosition - 1)
            End If

            If NTProductType = "WINNT" Then
                Platform_GetNTProduct = OSPP_WINDOWS_NT_400_WORKSTATION
            ElseIf NTProductType = "LANMANNT" Then
                If Major = 5 And Minor = 2 Then
                    Platform_GetNTProduct = OSPP_WINDOWS_2003
                ElseIf Major = 5 And Minor = 0 Then
                    Platform_GetNTProduct = OSPP_WINDOWS_2000_SERVER
                ElseIf Major <= 4 Then
                    Platform_GetNTProduct = OSPP_WINDOWS_NT_400_SERVER
                End If
            ElseIf NTProductType = "SERVERNT" Then
                If Major = 5 And Minor = 2 Then
                    Platform_GetNTProduct = OSPP_WINDOWS_2003_ENTERPRISE
                ElseIf Major = 5 And Minor = 0 Then
                    Platform_GetNTProduct = OSPP_WINDOWS_2000_SERVER_ADVANCED
                ElseIf Major <= 4 Then
                    Platform_GetNTProduct = OSPP_WINDOWS_NT_400_SERVER_ENTERPRISE
                End If
            End If
        End If
    End If

End Function

Public Function System_ProcessorArchitecture() As PROCESSOR_ARCHITECTURE

    System_ProcessorArchitecture = m_SystemInformation.wProcessorArchitecture

End Function

Public Function System_ProcessorArchitectureDescription() As String

    Select Case System_ProcessorArchitecture()
        Case PA_INTEL:         System_ProcessorArchitectureDescription = "X86"
        Case PA_IA64:          System_ProcessorArchitectureDescription = "Intel Itanium Processor Family (IPF)"
        Case PA_AMD64:         System_ProcessorArchitectureDescription = "X64 (AMD or Intel)"
        Case PA_IA32_ON_WIN64: System_ProcessorArchitectureDescription = "WOW64"
        Case Else:             System_ProcessorArchitectureDescription = "Unknown"
    End Select

End Function
