Imports System.Net.Sockets
Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Text

Public Class ServerConnection
    Implements IDisposable

    Public connected As Boolean

    Private serverIP As String
    Private serverPort As String
 
    Private client As TcpClient
    Private serverWriter As StreamWriter
    Private serverReader As StreamReader

    Public Sub New(ByVal ipAddress As String, ByVal port As String)
        Me.serverIP = ipAddress
        Me.serverPort = port
    End Sub

    Public Sub connectToServer(ByVal reconnect As Boolean, ByVal gunVariables As Dictionary(Of String, String))

        If (reconnect) Then 'clean up previous connection
            Try
                client.Close()
            Catch ex As Exception
            End Try
        End If

        client = New TcpClient()

        Dim serverEndPoint As IPEndPoint = New IPEndPoint(IPAddress.Parse(Me.serverIP), serverPort)
        client.Connect(serverEndPoint)
        connected = True
        Dim clientStream As NetworkStream = client.GetStream()
        serverWriter = New StreamWriter(clientStream)
        serverReader = New StreamReader(clientStream)

        If (reconnect) Then
            'tell server we are reconnecting and provide the session no to reconnect into
            Dim writeToServerString = String.Format("SESSION={0}|USER={1}|PASS={2}|GUNID={3}|GUNIP={4}", gunVariables("SESSION"), gunVariables("USER"), gunVariables("PASS"), gunVariables("GUNID"), gunVariables("GUNIP"))
            If gunVariables.ContainsKey("SCHEMA") Then
                writeToServerString &= "|SCHEMA" & gunVariables("SCHEMA")
            End If
            writeToServer(writetoserverstring)
        End If
    End Sub

    Private Function writeToServer(ByVal msg As String) As Boolean
        Dim connectionAttempts = -1, writeSuccess = False

        Do
            Try
                connectionAttempts += 1

                serverWriter.Write(msg)
                serverWriter.Flush()
                writeSuccess = True
            Catch ex As Exception
                Thread.Sleep(1000)
            End Try
        Loop While (connectionAttempts < 0 And writeSuccess = False)

        Return writeSuccess
    End Function

    Public Sub sendMessageToServer(ByVal serverMessages As Dictionary(Of String, String))
        Dim serverMessage As New StringBuilder()

        Dim delimiter As String = ""
        For Each key In serverMessages.Keys
            serverMessage.Append(delimiter)
            serverMessage.Append(key)
            serverMessage.Append("=")
            serverMessage.Append(serverMessages(key))
            delimiter = "|"
        Next

        writeToServer(serverMessage.ToString())
    End Sub

    Public Function readFromServer() As String
        Dim buffer(2048) As Char


        Dim received = serverReader.Read(buffer, 0, buffer.Length)
        Dim recString = New String(buffer, 0, received)
        Return recString.Substring(3) 'server messages are preceded with "ACK" (in the event an SP returns null to prevent read from blocking forever)
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                client.Close()
            End If

        End If
        Me.disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class