Imports Microsoft.UI.Xaml
Imports Microsoft.UI.Xaml.Markup
Imports Microsoft.UI.Dispatching
Imports System.Threading
Imports Microsoft.UI.Xaml.XamlTypeInfo
Imports System.Runtime
Imports Microsoft.Windows.ApplicationModel.DynamicDependency
Imports System.Runtime.CompilerServices
Imports VB_WinUI3_DesktopWindowXamlSource.Microsoft.WindowsAppSDK
Imports Windows.Management.Deployment
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Text.Json
Imports System.Environment

Friend Class Program
    Inherits Application
    Implements IXamlMetadataProvider

    Private Shared xamlMetaDataProvider As XamlControlsXamlMetaDataProvider

    Public Sub New()
        System.Windows.Forms.Application.Run(New Form1())
    End Sub

    Protected Overrides Sub OnLaunched(args As LaunchActivatedEventArgs)
        XamlControlsXamlMetaDataProvider.Initialize()
        xamlMetaDataProvider = New XamlControlsXamlMetaDataProvider()
        Me.Resources.MergedDictionaries.Add(New Controls.XamlControlsResources())
    End Sub

    ''' <summary>
    ''' The main entry point for the application.
    ''' </summary>
    <STAThread>
    Public Shared Sub Main()
        'ApplicationConfiguration.Initialize()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(False)
        System.Windows.Forms.Application.SetHighDpiMode(HighDpiMode.SystemAware)

        ' Try to read values from Nuget in Debug, otherwise they must be set from Windows App SDK version
        ' in Namespace Microsoft.WindowsAppSDK (see MddBootstrapAutoInitializer.cs and WindowsAppSDK-VersionInfo.cs in C#)
#If DEBUG Then
        Dim sNugetPath = NugetInfo.GetGlobalNugetPath()
        ' Dim info = WindowsAppSDKInfo.ReadSDKJsonInfo("1.6.241114003", sNugetPath) ' info = (65542, "", "6000.318.2304.0")
        Dim info = WindowsAppSDKInfo.ReadSDKJsonInfo("1.7.250401001", sNugetPath) 'info = (65543, "", "7000.456.1632.0")
        AutoInitialize.AccessWindowsAppSDK(info.majorMinor, info.versionTag, info.dotQuad)
#Else
         ' Change values in Namespace Microsoft.WindowsAppSDK
        AutoInitialize.AccessWindowsAppSDK(nothing, nothing, nothing)
#End If

        Application.Start(Sub(p)
                              Dim syncContext = New DispatcherQueueSynchronizationContext(DispatcherQueue.GetForCurrentThread())
                              SynchronizationContext.SetSynchronizationContext(syncContext)

                              Dim app = New Program()

                              Dim currentApp = Application.Current
                              If currentApp IsNot Nothing Then
                                  currentApp.Exit()
                              End If
                          End Sub)
    End Sub

    Public Function GetXamlType(type As Type) As IXamlType Implements IXamlMetadataProvider.GetXamlType
        Return xamlMetaDataProvider.GetXamlType(type)
    End Function

    Public Function GetXamlType(fullName As String) As IXamlType Implements IXamlMetadataProvider.GetXamlType
        Return xamlMetaDataProvider.GetXamlType(fullName)
    End Function

    Public Function GetXmlnsDefinitions() As XmlnsDefinition() Implements IXamlMetadataProvider.GetXmlnsDefinitions
        Return xamlMetaDataProvider.GetXmlnsDefinitions()
    End Function
End Class

' Copied from WindowsAppSDK-VersionInfo.cs for version 1.6.241114003
' C:\Users\Christian\.nuget\packages\microsoft.windowsappsdk\1.6.241114003\include\WindowsAppSDK-VersionInfo.cs

Namespace Microsoft.WindowsAppSDK

    ' Release information
    Friend Class Release
        ''' <summary>The major version of the Windows App SDK release.</summary>
        Friend Const Major As UShort = 1

        ''' <summary>The minor version of the Windows App SDK release.</summary>
        Friend Const Minor As UShort = 6

        ''' <summary>The patch version of the Windows App SDK release.</summary>
        Friend Const Patch As UShort = 0

        ''' <summary>The major and minor version encoded as 0xMMMMNNNN</summary>
        Friend Const MajorMinor As UInteger = &H10006

        ''' <summary>Channel, like "preview", or empty string for stable.</summary>
        Friend Const Channel As String = "stable"

        Friend Const VersionTag As String = ""
        Friend Const VersionShortTag As String = ""
        Friend Const FormattedVersionTag As String = ""
        Friend Const FormattedVersionShortTag As String = ""
    End Class

    ' Runtime information
    Namespace Runtime

        Friend Class Identity
            Friend Const Publisher As String = "CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US"
            Friend Const PublisherId As String = "8wekyb3d8bbwe"
        End Class

        Friend Class Version
            Friend Const Major As UShort = 6000
            Friend Const Minor As UShort = 318
            Friend Const Build As UShort = 2304
            Friend Const Revision As UShort = 0
            Friend Const UInt64 As ULong = &H1770013E09000000UL
            Friend Const DotQuadString As String = "6000.318.2304.0"
        End Class

        Namespace Packages

            Friend Class Framework
                Friend Const PackageFamilyName As String = "Microsoft.WindowsAppRuntime.1.6_8wekyb3d8bbwe"
            End Class

            Friend Class Main
                Friend Const PackageFamilyName As String = "MicrosoftCorporationII.WinAppRuntime.Main.1.6_8wekyb3d8bbwe"
            End Class

            Friend Class Singleton
                Friend Const PackageFamilyName As String = "MicrosoftCorporationII.WinAppRuntime.Singleton_8wekyb3d8bbwe"
            End Class

            Namespace DDLM

                Friend Class X86
                    Friend Const PackageFamilyName As String = "Microsoft.WinAppRuntime.DDLM.6000.318.2304.0-x8_8wekyb3d8bbwe"
                End Class

                Friend Class X64
                    Friend Const PackageFamilyName As String = "Microsoft.WinAppRuntime.DDLM.6000.318.2304.0-x6_8wekyb3d8bbwe"
                End Class

                Friend Class Arm64
                    Friend Const PackageFamilyName As String = "Microsoft.WinAppRuntime.DDLM.6000.318.2304.0-a6_8wekyb3d8bbwe"
                End Class

            End Namespace
        End Namespace

    End Namespace
End Namespace

' Copied from MddBootstrapAutoInitializer.cs for version 1.6.241114003

#If Not MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_DEFAULT Then

' Default isn't defined. Check if any other options are defined
#If MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_NONE Then
#ElseIf MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_ONERROR_DEBUGBREAK Then
#ElseIf MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_ONERROR_DEBUGBREAK_IFDEBUGGERATTACHED Then
#ElseIf MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_ONERROR_FAILFAST Then
#ElseIf MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_ONNOMATCH_SHOWUI Then
#ElseIf MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_ONPACKAGEIDENTITY_NOOP Then
#Else
' No options specified! Define the default
#Const MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_DEFAULT = True
#End If

#End If

Friend NotInheritable Class AutoInitialize

    Friend Shared Sub AccessWindowsAppSDK(majorMinor As UInteger, versionTag As String, dotQuad As String)
        Dim nMajorMinorVersion As UInteger = Release.MajorMinor
        Dim sVersionTag As String = Release.VersionTag
        Dim minVersion As New PackageVersion(Runtime.Version.UInt64)
        If Not (majorMinor = Nothing And versionTag = Nothing And dotQuad = Nothing) Then
            nMajorMinorVersion = majorMinor
            sVersionTag = versionTag
            minVersion = DotQuadToPackageVersion(dotQuad)
        End If

        Dim options As Bootstrap.InitializeOptions = _Options
        Dim hr As Integer = 0

        ' for 1.6.241114003
        ' nMajorMinorVersion = 65542
        ' sVersionTag = ""
        ' minVersion = {6000.318.2304.0}

        ' 0x80670016 MddBootstrapInitialize2
        If Not Bootstrap.TryInitialize(nMajorMinorVersion, sVersionTag, minVersion, options, hr) Then
            Environment.Exit(hr)
        End If
    End Sub

    Friend Shared Function DotQuadToPackageVersion(dotQuad As String) As PackageVersion
        ' Split the dotQuad string into its components (Major, Minor, Build, Revision)
        Dim parts = dotQuad.Split("."c)

        If parts.Length <> 4 Then
            Throw New FormatException("Invalid DotQuad version format.")
        End If

        ' Parse each part to integers
        Dim nMajor = Convert.ToUInt32(parts(0))
        Dim nMinor = Convert.ToUInt32(parts(1))
        Dim nBuild = Convert.ToUInt32(parts(2))
        Dim nRevision = Convert.ToUInt32(parts(3))

        Dim nVersion64 As ULong = (CLng(nMajor) << 48) Or (CLng(nMinor) << 32) Or (CLng(nBuild) << 16) Or CLng(nRevision)

        Return New PackageVersion(nVersion64)
    End Function

    Friend Shared ReadOnly Property _Options As Bootstrap.InitializeOptions
        Get
#If MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_DEFAULT Then
            ' Use the default options
            Return Bootstrap.InitializeOptions.OnNoMatch_ShowUI
#ElseIf MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_NONE Then
            ' No options
            Return Bootstrap.InitializeOptions.None
#Else
            ' Use the specified options
            Dim opts As Bootstrap.InitializeOptions = Bootstrap.InitializeOptions.None
#If MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_ONERROR_DEBUGBREAK Then
            opts = opts Or Bootstrap.InitializeOptions.OnError_DebugBreak
#End If
#If MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_ONERROR_DEBUGBREAK_IFDEBUGGERATTACHED Then
            opts = opts Or Bootstrap.InitializeOptions.OnError_DebugBreak_IfDebuggerAttached
#End If
#If MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_ONERROR_FAILFAST Then
            opts = opts Or Bootstrap.InitializeOptions.OnError_FailFast
#End If
#If MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_ONNOMATCH_SHOWUI Then
            opts = opts Or Bootstrap.InitializeOptions.OnNoMatch_ShowUI
#End If
#If MICROSOFT_WINDOWSAPPSDK_BOOTSTRAP_AUTO_INITIALIZE_OPTIONS_ONPACKAGEIDENTITY_NOOP Then
            opts = opts Or Bootstrap.InitializeOptions.OnPackageIdentity_NOOP
#End If
            Return opts
#End If
        End Get
    End Property

End Class

Friend Class WindowsAppSDKInfo
    Friend Shared Function ReadSDKJsonInfo(version As String, Optional sNugetPath As String = Nothing) As (majorMinor As UInteger, versionTag As String, dotQuad As String)
        Dim sUserProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        Dim sJSONPath As String
        If String.IsNullOrEmpty(sNugetPath) Then
            sJSONPath = Path.Combine(sUserProfile, ".nuget", "packages", "microsoft.windowsappsdk", version, "WindowsAppSDK-VersionInfo.json")
        Else
            sJSONPath = Path.Combine(sNugetPath, "microsoft.windowsappsdk", version, "WindowsAppSDK-VersionInfo.json")
        End If
        If Not File.Exists(sJSONPath) Then
            Throw New FileNotFoundException("WindowsAppSDK-VersionInfo.json not found at: " & sJSONPath)
        End If

        Dim sJSONText = File.ReadAllText(sJSONPath)
        Using doc = JsonDocument.Parse(sJSONText)
            Dim root = doc.RootElement

            Dim majorMinor = root.GetProperty("Release").GetProperty("MajorMinor").GetProperty("UInt32").GetUInt32()
            Dim sVersionTag = root.GetProperty("Release").GetProperty("VersionTag").GetString()
            Dim sDotQuad = root.GetProperty("Runtime").GetProperty("Version").GetProperty("DotQuadNumber").GetString()

            Return (majorMinor, sVersionTag, sDotQuad)
        End Using
    End Function
End Class

Friend Class NugetInfo
    Friend Shared Function GetGlobalNugetPath() As String
        Dim sEnvPath = GetEnvironmentVariable("NUGET_PACKAGES")
        If Not String.IsNullOrEmpty(sEnvPath) Then
            Return sEnvPath
        End If

        Return Path.Combine(GetFolderPath(SpecialFolder.UserProfile), ".nuget", "packages")
    End Function
End Class


