Imports System.Messaging

Public Class ASCQUEUE

    Public QueueName As String = String.Empty
    Public LastError As String = String.Empty

    Private WaitingForAMessage As Boolean = False

    Public Sub New()

    End Sub

    Public Sub New(ByVal vQueueName As String)
        QueueName = vQueueName
    End Sub

    ''' <summary>
    ''' Creates a Queue for the Class Queue name
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateQueue() As Boolean
        Return CreateQueue(QueueName)
    End Function

    ''' <summary>
    ''' Creates a Queue for the Queue Name provided
    ''' </summary>
    ''' <param name="vQueueName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateQueue(ByVal vQueueName As String) As Boolean
        LastError = String.Empty

        'Try
        'create a queue on the local machine named 
        'Dim mg As MessageQueue = New System.Messaging.MessageQueue(vQueueName)
        MessageQueue.Create(vQueueName)
        QueueName = vQueueName
        Return True
        'Catch ex As Exception
        '    LastError = "Exception was thrown: " & ex.Source & ": " & ex.Message
        '    Return False
        'End Try
    End Function

    ''' <summary>
    ''' Deletes the Class Queue
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function DeleteQueue() As Boolean
        LastError = String.Empty
        'Try
        ' note that queue names are not case-specific
        MessageQueue.Delete(QueueName)
        Return True
        'Catch ex As Exception
        '    LastError = "Exception was thrown: " & ex.Source & ": " & ex.Message
        '    Return False
        'End Try
    End Function

    ''' <summary>
    ''' Purges the Class Queue
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PurgeQueue() As Boolean
        Dim myQueue As MessageQueue
        LastError = String.Empty

        'Try
        ' Create an instance variable before calling the Purge method
        myQueue = New MessageQueue(QueueName)
        myQueue.Purge()
        Return True
        'Catch ex As Exception
        '    LastError = "Exception was thrown: " & ex.Source & ": " & ex.Message
        '    Return False
        'End Try
    End Function

    ''' <summary>
    ''' Boolean indicating if a Queue Exists
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function QueueExists() As Boolean
        Return QueueExists(QueueName)
    End Function

    ''' <summary>
    ''' Boolean indicating if a Queue Exists 
    ''' </summary>
    ''' <param name="vQueueName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function QueueExists(ByVal vQueueName As String) As Boolean
        LastError = String.Empty
        'Try
        Return MessageQueue.Exists(vQueueName)
        'Catch ex As Exception
        '    LastError = "Exception was thrown: " & ex.Source & ": " & ex.Message
        '    Return False
        'End Try

    End Function

    ''' <summary>
    ''' Grant the Queue's Permissions
    ''' </summary>
    ''' <param name="User">The individual, group or computer that gets the rights to the queue</param>
    ''' <param name="AccessRightsType">Set of rights to the Queue for the user passed in</param>
    ''' <param name="AccessControlEntryType">Whether to grant, deny, revoke the permission specified by the rights parameter</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GrantQueuePermissions(ByVal User As String, ByVal AccessRightsType As MessageQueueAccessRights, ByVal AccessControlEntryType As AccessControlEntryType) As Boolean

        Dim myQueue As MessageQueue
        LastError = String.Empty
        'Try
        ' grant Everyone Full Control permissions to my queue
        myQueue = New MessageQueue(QueueName)
        'myQueue.SetPermissions("Everyone", MessageQueueAccessRights.FullControl, AccessControlEntryType.Set)
        myQueue.SetPermissions(User, AccessRightsType, AccessControlEntryType)
        Return True
        'Catch ex As Exception
        '    LastError = "Exception was thrown: " & ex.Source & ": " & ex.Message
        '    Return False
        'End Try
    End Function

    ''' <summary>
    ''' Revokes a user's permisison to a queue
    ''' </summary>
    ''' <param name="User">The individual, group or computer that get its rights revoked</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RevokeUserPermissions(ByVal User As String) As Boolean

        Dim myQueue As MessageQueue
        LastError = String.Empty

        'Try

        ' revoke the Full Control permission from Everyone
        myQueue = New MessageQueue(QueueName)
        'myQueue.SetPermissions("Everyone", MessageQueueAccessRights.FullControl, AccessControlEntryType.Revoke)
        myQueue.SetPermissions(User, MessageQueueAccessRights.FullControl, AccessControlEntryType.Revoke)
        Return True
        'Catch ex As Exception
        '    LastError = "Exception was thrown: " & ex.Source & ": " & ex.Message
        '    Return False
        'End Try
    End Function

    ''' <summary>
    ''' Returns a list of all public queuese on a Machine
    ''' </summary>
    ''' <param name="MachineName">Machine Name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPublicQueues(ByVal MachineName As String) As List(Of String)

        Dim queues As List(Of String) = New List(Of String)

        Dim myQueues() As MessageQueue

        'Try
        ' get a list of all public queues on the specified machine
        myQueues = MessageQueue.GetPublicQueuesByMachine(MachineName)

        ' loop through each queue in the array writing out the name of each queue
        For Each myQueue As MessageQueue In myQueues
            queues.Add(myQueue.QueueName)
        Next
        Return queues
        'Catch ex As Exception
        '    LastError = "Exception was thrown: " & ex.Source & ": " & ex.Message
        '    Return New List(Of String)
        'End Try

    End Function

    ''' <summary>
    ''' Returns a list of all private queuese on a Machine
    ''' </summary>
    ''' <param name="MachineName">Machine Name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPrivateQueues(ByVal MachineName As String) As List(Of String)

        Dim queues As List(Of String) = New List(Of String)

        Dim myQueues() As MessageQueue

        'Try
        ' get a list of all public queues on the specified machine
        myQueues = MessageQueue.GetPrivateQueuesByMachine(MachineName)

        ' loop through each queue in the array writing out the name of each queue
        For Each myQueue As MessageQueue In myQueues
            queues.Add(myQueue.QueueName)
        Next
        Return queues
        'Catch ex As Exception
        '    LastError = "Exception was thrown: " & ex.Source & ": " & ex.Message
        '    Return New List(Of String)
        'End Try

    End Function

    ''' <summary>
    ''' Sends and Object to the Queue
    ''' </summary>
    ''' <param name="QueueObject"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PostMessage(ByRef QueueObject As Object) As Boolean

        Return PostMessage(QueueName, QueueObject)
    End Function

    ''' <summary>
    ''' Sends an Object to A Queue
    ''' </summary>
    ''' <param name="Queuename"></param>
    ''' <param name="QueueObject"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PostMessage(ByVal Queuename As String, ByRef QueueObject As Object) As Boolean

        Dim myQueue As MessageQueue
        LastError = String.Empty

        'Try

        ' connect to the queue and send the message
        myQueue = New MessageQueue(Queuename)
        myQueue.Send(QueueObject)
        Return True

        'Catch ex As Exception

        '    LastError = "Exception was thrown: " & ex.Source & ": " & ex.Message
        '    Return False

        'End Try
    End Function

    ''' <summary>
    ''' Retrieves a message from the queue
    ''' </summary>
    ''' <param name="QueueObject">String used to hold the Message in the Queue</param>
    ''' <param name="SecondsToWait">Number of Seconds to wait fot the request</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetQueueMessage(ByRef QueueObject As Object, ByVal SecondsToWait As Integer) As Boolean

        Dim myQueue As MessageQueue
        Dim myMsg As System.Messaging.Message
        LastError = String.Empty

        If SecondsToWait <= 0 Then SecondsToWait = 5

        Dim minutes As Integer = SecondsToWait \ 60
        Dim seconds As Integer = SecondsToWait Mod 60

        Try

            ' connect to the queue and set the message formatter
            myQueue = New MessageQueue(QueueName)
            myQueue.Formatter = New XmlMessageFormatter(New Type() {QueueObject.GetType})

            'receive the message, waiting no longer than 5 seconds
            myMsg = myQueue.Receive(New TimeSpan(0, minutes, seconds))

            QueueObject = myMsg.Body
            Return True

        Catch exQueue As MessageQueueException

            ' this exception will be thrown when no message was received
            LastError = exQueue.Message
            Return False

        Catch ex As Exception
            LastError = "Generic Exception was thrown: " & ex.Source & ": " & ex.Message
            Return False

        End Try
    End Function

End Class
