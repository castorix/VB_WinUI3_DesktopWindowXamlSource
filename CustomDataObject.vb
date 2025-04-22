Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

' Adapted from WinUI 3 Gallery

Public Class CustomDataObject
    Public Property Title As String
    Public Property ImageLocation As String
    Public Property Description As String

    Public Sub New()
    End Sub

    Public Shared Function GetDataObjects(sRelativePath As String) As List(Of CustomDataObject)
        Dim objects As New List(Of CustomDataObject)()
        Dim sExePath As String = AppContext.BaseDirectory
        Dim sImageDirectory As String = Path.Combine(sExePath, sRelativePath)
        Dim images = Directory.EnumerateFiles(sImageDirectory, "*.*", SearchOption.AllDirectories)

        For Each sCurrentFile In images
            Dim pPropertyStore As IPropertyStore = Nothing
            Dim PropertyStoreGuid As Guid = GetType(IPropertyStore).GUID
            Dim hr As HRESULT = SHGetPropertyStoreFromParsingName(sCurrentFile, IntPtr.Zero, GETPROPERTYSTOREFLAGS.GPS_READWRITE, PropertyStoreGuid, pPropertyStore)

            Dim sTitle As String = ""
            If hr = HRESULT.S_OK Then
                Dim pv As PROPVARIANT = New PROPVARIANT()
                hr = pPropertyStore.GetValue(PKEY_Title, pv)
                If hr = HRESULT.S_OK Then
                    sTitle = Marshal.PtrToStringUni(pv.pwszVal)
                End If
                Marshal.ReleaseComObject(pPropertyStore)
            End If

            objects.Add(New CustomDataObject() With {
                        .Title = sTitle,
                        .ImageLocation = sCurrentFile
                        })
        Next

        Return objects
    End Function

    Public Enum HRESULT As Integer
        S_OK = 0
        S_FALSE = 1
        E_NOINTERFACE = &H80004002
        E_NOTIMPL = &H80004001
        E_FAIL = &H80004005
        E_UNEXPECTED = &H8000FFFF
        E_OUTOFMEMORY = &H8007000E
    End Enum

    <DllImport("Shell32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function SHGetPropertyStoreFromParsingName(<MarshalAs(UnmanagedType.LPWStr)> pszPath As String, pbc As IntPtr,
                                                             flags As GETPROPERTYSTOREFLAGS, ByRef iid As Guid,
                                                             <MarshalAs(UnmanagedType.Interface)> ByRef propertyStore As IPropertyStore) As HRESULT
    End Function

    Public Enum GETPROPERTYSTOREFLAGS
        GPS_DEFAULT = 0
        GPS_HANDLERPROPERTIESONLY = &H1
        GPS_READWRITE = &H2
        GPS_TEMPORARY = &H4
        GPS_FASTPROPERTIESONLY = &H8
        GPS_OPENSLOWITEM = &H10
        GPS_DELAYCREATION = &H20
        GPS_BESTEFFORT = &H40
        GPS_NO_OPLOCK = &H80
        GPS_PREFERQUERYPROPERTIES = &H100
        GPS_EXTRINSICPROPERTIES = &H200
        GPS_EXTRINSICPROPERTIESONLY = &H400
        GPS_VOLATILEPROPERTIES = &H800
        GPS_VOLATILEPROPERTIESONLY = &H1000
        GPS_MASK_VALID = &H1FFF
    End Enum

    <ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IPropertyStore
        Function GetCount(<Out> ByRef propertyCount As UInteger) As HRESULT
        Function GetAt(propertyIndex As UInteger, <Out> ByRef key As PROPERTYKEY) As HRESULT
        Function GetValue(<[In]> ByRef key As PROPERTYKEY, <Out> ByRef pv As PROPVARIANT) As HRESULT
        Function SetValue(<[In]> ByRef key As PROPERTYKEY, <[In]> ByRef pv As PROPVARIANT) As HRESULT
        Function Commit() As HRESULT
    End Interface

    <StructLayout(LayoutKind.Sequential, Pack:=4)>
    Public Structure PROPERTYKEY
        Private ReadOnly _fmtid As Guid
        Private ReadOnly _pid As UInteger

        Public Sub New(fmtid As Guid, pid As UInteger)
            _fmtid = fmtid
            _pid = pid
        End Sub
    End Structure

    Public Shared ReadOnly PKEY_Title As New PROPERTYKEY(New Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"), 2)
    Public Shared ReadOnly PKEY_Keywords As New PROPERTYKEY(New Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"), 5)
    Public Shared ReadOnly PKEY_Comment As New PROPERTYKEY(New Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9"), 6)

    <StructLayout(LayoutKind.Sequential)>
    Public Structure PROPARRAY
        Public cElems As UInteger
        Public pElems As IntPtr
    End Structure

    <StructLayout(LayoutKind.Explicit, Pack:=1)>
    Public Structure PROPVARIANT
        <FieldOffset(0)> Public varType As UShort
        <FieldOffset(2)> Public wReserved1 As UShort
        <FieldOffset(4)> Public wReserved2 As UShort
        <FieldOffset(6)> Public wReserved3 As UShort

        <FieldOffset(8)> Public bVal As Byte
        <FieldOffset(8)> Public cVal As SByte
        <FieldOffset(8)> Public uiVal As UShort
        <FieldOffset(8)> Public iVal As Short
        <FieldOffset(8)> Public uintVal As UInteger
        <FieldOffset(8)> Public intVal As Integer
        <FieldOffset(8)> Public ulVal As ULong
        <FieldOffset(8)> Public lVal As Long
        <FieldOffset(8)> Public fltVal As Single
        <FieldOffset(8)> Public dblVal As Double
        <FieldOffset(8)> Public boolVal As Short
        <FieldOffset(8)> Public pclsidVal As IntPtr
        <FieldOffset(8)> Public pszVal As IntPtr
        <FieldOffset(8)> Public pwszVal As IntPtr
        <FieldOffset(8)> Public punkVal As IntPtr
        <FieldOffset(8)> Public ca As PROPARRAY
        <FieldOffset(8)> Public filetime As FILETIME
    End Structure

    Public Enum VARENUM
        VT_EMPTY = 0
        VT_NULL = 1
        VT_I2 = 2
        VT_I4 = 3
        VT_R4 = 4
        VT_R8 = 5
        VT_CY = 6
        VT_DATE = 7
        VT_BSTR = 8
        VT_DISPATCH = 9
        VT_ERROR = 10
        VT_BOOL = 11
        VT_VARIANT = 12
        VT_UNKNOWN = 13
        VT_DECIMAL = 14
        VT_I1 = 16
        VT_UI1 = 17
        VT_UI2 = 18
        VT_UI4 = 19
        VT_I8 = 20
        VT_UI8 = 21
        VT_INT = 22
        VT_UINT = 23
        VT_VOID = 24
        VT_HRESULT = 25
        VT_PTR = 26
        VT_SAFEARRAY = 27
        VT_CARRAY = 28
        VT_USERDEFINED = 29
        VT_LPSTR = 30
        VT_LPWSTR = 31
        VT_RECORD = 36
        VT_INT_PTR = 37
        VT_UINT_PTR = 38
        VT_FILETIME = 64
        VT_BLOB = 65
        VT_STREAM = 66
        VT_STORAGE = 67
        VT_STREAMED_OBJECT = 68
        VT_STORED_OBJECT = 69
        VT_BLOB_OBJECT = 70
        VT_CF = 71
        VT_CLSID = 72
        VT_VERSIONED_STREAM = 73
        VT_BSTR_BLOB = &HFFF
        VT_VECTOR = &H1000
        VT_ARRAY = &H2000
        VT_BYREF = &H4000
        VT_RESERVED = &H8000
        VT_ILLEGAL = &HFFFF
        VT_ILLEGALMASKED = &HFFF
        VT_TYPEMASK = &HFFF
    End Enum
End Class

