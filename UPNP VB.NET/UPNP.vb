Imports System.Runtime.InteropServices

Friend Class UPnP
    Implements IDisposable

    Public Shared Sub ForceTCP(ByVal Port As Integer, Optional SupressError As Boolean = True)
        Force(Port, Protocol.TCP, SupressError)
    End Sub

    Public Shared Sub ForceUDP(ByVal Port As Integer, Optional SupressError As Boolean = True)
        Force(Port, Protocol.UDP, SupressError)
    End Sub

    Public Shared Sub Force(ByVal Port As Integer, Prot As Protocol, Optional SupressError As Boolean = True)
        Dim U As New UPnP

        If U.Exists(Port, Prot) Then
            U.Remove(Port, Prot, SupressError)
        End If

        U.Add(Port, Prot, Application.ProductName, SupressError)
    End Sub



    Private upnpnat As NATUPNPLib.UPnPNAT
    Private staticMapping As NATUPNPLib.IStaticPortMappingCollection
    Private dynamicMapping As NATUPNPLib.IDynamicPortMappingCollection

    Private staticEnabled As Boolean = True
    Private dynamicEnabled As Boolean = True

    Public Enum Protocol
        TCP
        UDP
    End Enum

    Public ReadOnly Property UPnPEnabled As Boolean
        Get
            Return staticEnabled = True OrElse dynamicEnabled = True
        End Get
    End Property

    Public Sub New()
        upnpnat = New NATUPNPLib.UPnPNAT
        Me.GetStaticMappings()
        Me.GetDynamicMappings()
    End Sub

    Private Sub GetStaticMappings()
        Try
            staticMapping = upnpnat.StaticPortMappingCollection()
        Catch ex As NotImplementedException
            staticEnabled = False
        End Try
    End Sub

    Private Sub GetDynamicMappings()
        Try
            dynamicMapping = upnpnat.DynamicPortMappingCollection()
        Catch ex As NotImplementedException
            dynamicEnabled = False
        End Try
    End Sub

    Public Sub Add(ByVal Port As Integer, ByVal prot As Protocol, ByVal desc As String, Optional SupressError As Boolean = True)
        Try
            Dim LocalIP As String = GetInternalIPV4()
            If Exists(Port, prot) Then Throw New ArgumentException("This mapping already exists!", "Port;prot")
            If Not IsPrivateIP(LocalIP) Then Throw New ArgumentException("This is not a local IP address!", "localIP")
            If Not staticEnabled Then Throw New ApplicationException("UPnP is not enabled, or there was an error with UPnP Initialization.")
            staticMapping.Add(Port, prot.ToString(), Port, LocalIP, True, desc)
        Catch ex As Exception When SupressError
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub Add(ByVal localIP As String, ByVal Port As Integer, ByVal prot As Protocol, ByVal desc As String, Optional SupressError As Boolean = True)
        Try
            If Exists(Port, prot) Then Throw New ArgumentException("This mapping already exists!", "Port;prot")
            If Not IsPrivateIP(localIP) Then Throw New ArgumentException("This is not a local IP address!", "localIP")
            If Not staticEnabled Then Throw New ApplicationException("UPnP is not enabled, or there was an error with UPnP Initialization.")
            staticMapping.Add(Port, prot.ToString(), Port, localIP, True, desc)
        Catch ex As Exception When SupressError
        End Try
    End Sub

    Public Sub Remove(ByVal Port As Integer, ByVal Prot As Protocol, Optional SupressError As Boolean = True)
        Try
            If Not Exists(Port, Prot) Then Throw New ArgumentException("This mapping doesn't exist!", "Port;prot")
            If Not staticEnabled Then Throw New ApplicationException("UPnP is not enabled, or there was an error with UPnP Initialization.")
            staticMapping.Remove(Port, Prot.ToString)
        Catch ex As Exception When SupressError
        End Try
    End Sub

    Public Function Exists(ByVal Port As Integer, ByVal Prot As Protocol) As Boolean
        Try
            If Not staticEnabled Then Throw New ApplicationException("UPnP is not enabled, or there was an error with UPnP Initialization.")
            For Each mapping As NATUPNPLib.IStaticPortMapping In staticMapping
                If mapping.ExternalPort.Equals(Port) AndAlso mapping.Protocol.ToString.Equals(Prot.ToString) Then Return True
            Next
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function LocalIP() As String
        Dim IPList As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName)
        For Each IPaddress In IPList.AddressList
            If (IPaddress.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork) AndAlso IsPrivateIP(IPaddress.ToString()) Then
                Return IPaddress.ToString
            End If
        Next
        Return String.Empty
    End Function

    Private Shared Function IsPrivateIP(ByVal CheckIP As String) As Boolean
        Dim Quad1, Quad2 As Integer

        Quad1 = CInt(CheckIP.Substring(0, CheckIP.IndexOf(".")))
        Quad2 = CInt(CheckIP.Substring(CheckIP.IndexOf(".") + 1).Substring(0, CheckIP.IndexOf(".")))
        Select Case Quad1
            Case 10
                Return True
            Case 172
                If Quad2 >= 16 And Quad2 <= 31 Then Return True
            Case 192
                If Quad2 = 168 Then Return True
        End Select
        Return False
    End Function

    Protected Overridable Sub Dispose(disposing As Boolean)
        Marshal.ReleaseComObject(staticMapping)
        Marshal.ReleaseComObject(dynamicMapping)
        Marshal.ReleaseComObject(upnpnat)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Public Function Print() As List(Of String)
        Dim L As New List(Of String)
        Try
            If staticEnabled Then
                For Each mapping As NATUPNPLib.IStaticPortMapping In staticMapping
                    L.Add("--------------------------------------")
                    L.Add(String.Format("IP: {0}", mapping.InternalClient))
                    L.Add(String.Format("Port: {0}", mapping.InternalPort))
                    L.Add(String.Format("Description: {0}", mapping.Description))
                Next
            End If
        Catch ex As Exception
        End Try

        L.Add("--------------------------------------")
        Return L
    End Function

    Public Function List() As List(Of PortMap)

        Dim Maps As New List(Of PortMap)

        Try

            If staticEnabled Then
                For Each mapping As NATUPNPLib.IStaticPortMapping In staticMapping

                    Dim Map As New PortMap
                    Map.InternalIP = mapping.InternalClient
                    Map.InternalPort = mapping.InternalPort

                    Map.ExternalIP = mapping.ExternalIPAddress
                    Map.ExternalPort = mapping.ExternalPort

                    Map.Descriptioin = mapping.Description
                    Map.Protocol = mapping.Protocol
                    Map.Enabled = mapping.Enabled

                    Maps.Add(Map)
                Next
            End If

        Catch ex As Exception
        End Try

        Return maps

    End Function

    Public Function GetInternalIPV4(Optional ByVal Index As Integer = 0) As String
        Dim h As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName)
        Dim internalip As String = h.AddressList.GetValue(Index).ToString
        Return internalip
    End Function

    Public Function GetInternalIPV6(Optional ByVal Index As Integer = 0) As String
        Dim h As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName)
        Dim internalip As String = h.AddressList.GetValue(Index).ToString
        Return internalip
    End Function

End Class

Friend Class PortMap
    Property InternalIP As String
    Property InternalPort As Integer

    Property ExternalIP As String
    Property ExternalPort As Integer

    Property Descriptioin As String

    Property Protocol As String

    Property Enabled As Boolean

End Class