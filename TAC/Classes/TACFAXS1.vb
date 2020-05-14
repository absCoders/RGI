Imports System
Imports System.IO
Imports System.Threading
Imports System.Web.Services.Protocols

    Public Class TACFAXS1

    Public fax_Username As String
    Public fax_Password As String
    Public fax_CoverFile As String
    Public fax_FaxAttachment As String

    Public SEND_FROM As String
    Public SEND_FROM_NAME As String
    Public SEND_TO As String
    Public SEND_TO_NAME As String
    Public SEND_SUBJECT As String
    Public SEND_BODY As String
    Public SEND_CODE As String
    Public SEND_NAME As String

    Public fax_transportID As Integer
    Public fax_log As New System.Text.StringBuilder

    Public Sub New()

    End Sub

    ' Method used to read data from a file and store them in a Web Service file object.
    Private Function ReadFile(ByVal filename As String) As ODSubmission.WSFile
        Dim wsFile As New ODSubmission.WSFile
        wsFile.mode = ODSubmission.WSFILE_MODE.MODE_INLINED
        wsFile.name = Me.shortFileName(filename)
        Dim myFile As FileStream = File.OpenRead(filename)
        wsFile.content = New Byte(myFile.Length - 1) {}
        myFile.Read(wsFile.content, 0, CType(myFile.Length, Integer))
        myFile.Close()
        Return wsFile
    End Function

    ' Helper method to allocate and fill in Variable objects.
    Private Function CreateValue(ByVal AttributeName As String, ByVal AttributeValue As String) As ODSubmission.Var
        Dim var As New ODSubmission.Var
        var.attribute = AttributeName
        var.simpleValue = AttributeValue
        Return var
    End Function

    ' Helper method to extract the short file name from a full file path
    Public Function shortFileName(ByVal filename As String) As String
        Dim state As Integer = filename.LastIndexOf("\")
        If (state < 0) Then
            Return filename
        End If
        Return filename.Substring((state + 1))
    End Function

    Public Sub SendFax()
        '//////////////////////////////////////////////////////////////////////
        '// STEP #2 : Initialization + Authentication
        '//////////////////////////////////////////////////////////////////////

        fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Retrieving bindings")

        Dim session As New ODSession.SessionService

        ' Retrieve the bindings on the Application Server (location of the Web Services)
        Dim bindings As ODSession.BindingResult = Nothing
        Try
            'bindings = session.GetBindings("")
            bindings = session.GetBindings(fax_Username)
        Catch ex As SoapException
            Dim detail = ex.Detail.SelectSingleNode("APIErrorMessage").InnerText
            Dim errorCode = ex.Detail.SelectSingleNode("APIErrorCode").InnerText
            fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Call to GetBindings() failed with message: " & ex.Message & " [" & errorCode + "/" & detail & "]")
            Return
        End Try

        fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Binding = " & bindings.sessionServiceLocation)

        ' Now uses the returned URL with our session object, in case the Application Server redirected us.
        session.Url = bindings.sessionServiceLocation

        fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Authenticating session")

        ' Authenticate the user on this session object to retrieve a sessionID
        Dim login As ODSession.LoginResult = Nothing
        Try
            login = session.Login(fax_Username, fax_Password)
        Catch ex As SoapException
            Dim detail = ex.Detail.SelectSingleNode("APIErrorMessage").InnerText
            Dim errorCode = ex.Detail.SelectSingleNode("APIErrorCode").InnerText
            fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Call to Login() failed with message: " & ex.Message & " [" & errorCode + "/" & detail & "]")
            Return
        End Try

        ' This sessionID is an impersonation token representing the logged on user
        ' You can use it with other Web Services objects, until you call Logout (which releases the
        ' current(sessionID And it) 's associated resources), or until the session times out (default is 10
        ' minutes on the Application Server).
        fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "SessionID = " & login.sessionID)

        '//////////////////////////////////////////////////////////////////////
        '// STEP #3 : Simple fax submission
        '//////////////////////////////////////////////////////////////////////

        ' Creating and initializing a SubmissionService object.
        Dim submissionService As New ODSubmission.SubmissionService

        ' Set the service URL with the location retrieved above with GetBindings()
        submissionService.Url = bindings.submissionServiceLocation
        ' Set the sessionID with the one retrieved above with Login()
        ' Every action performed on this object will now use the authenticated context created in step 1
        submissionService.SessionHeaderValue = New ODSubmission.SessionHeader
        submissionService.SessionHeaderValue.sessionID = login.sessionID

        ' Cover file resource is now made available on the server for the current user
        ' The cover resource file is specific to the fax transport, and should be ignored when submitting
        ' other transport types.
        ' Once it is registered on the server, you should not have to upload it each time a transport is submitted.
        ' Unlike other files, resources are permanently stored on the server even after a call to Logout()
        fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Registering cover resource")

        Try
            Dim cover As ODSubmission.WSFile = Me.ReadFile(fax_CoverFile)
            submissionService.RegisterResource(cover, ODSubmission.RESOURCE_TYPE.TYPE_COVER, False, True)
        Catch ex As SoapException
            Dim detail = ex.Detail.SelectSingleNode("APIErrorMessage").InnerText
            Dim errorCode = ex.Detail.SelectSingleNode("APIErrorCode").InnerText
            fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Call to UploadResources() failed with message: " & ex.Message & " [" & errorCode + "/" & detail & "]")
            Return
        End Try

        fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Sending Fax Request")

        ' Now allocate a transport with transportName = "Fax"
        Dim transport As New ODSubmission.Transport
        transport.transportName = "Fax"

        ' Specifies fax variables (see documentation for their definitions)
        transport.vars = New ODSubmission.Var(8) {}
        transport.vars(0) = CreateValue("Subject", SEND_SUBJECT)
        transport.vars(1) = CreateValue("FaxNumber", SEND_TO)
        transport.vars(2) = CreateValue("Message", SEND_BODY)
        transport.vars(3) = CreateValue("FromName", SEND_FROM)
        transport.vars(4) = CreateValue("FromCompany", SEND_FROM_NAME)
        transport.vars(5) = CreateValue("FromFax", "")
        transport.vars(6) = CreateValue("ToName", SEND_TO_NAME)
        transport.vars(7) = CreateValue("ToCompany", SEND_CODE & " - " & SEND_NAME)
        transport.vars(8) = CreateValue("CoverTemplate", Me.shortFileName(fax_CoverFile))

        ' Specify a pdf attachment to append to the fax.
        ' The attachment content is inlined in the transport description

        'transport.attachments = New ODSubmission.Attachment(1) {}
        'transport.attachments(0) = New ODSubmission.Attachment
        'transport.attachments(0).sourceAttachment = Me.ReadFile(fax_FaxAttachment)

        Dim numAttachments As Integer = 0
        For Each attachment As String In fax_FaxAttachment.Split(";")
            If attachment.Length = 0 Then Continue For
            numAttachments += 1
        Next

        If numAttachments = 0 Then numAttachments = 1
        transport.attachments = New ODSubmission.Attachment(numAttachments) {}

        Dim attachNo As Integer = 0
        For Each attachment As String In fax_FaxAttachment.Split(";")
            If attachment.Length = 0 Then Continue For
            transport.attachments(attachNo) = New ODSubmission.Attachment
            transport.attachments(attachNo).sourceAttachment = Me.ReadFile(attachment)
            attachNo += 1
        Next

        ' Submit the complete transport description to the Application Server
        Dim result As ODSubmission.SubmissionResult = Nothing
        Try
            result = submissionService.SubmitTransport(transport)
        Catch ex As SoapException
            Dim detail = ex.Detail.SelectSingleNode("APIErrorMessage").InnerText
            Dim errorCode = ex.Detail.SelectSingleNode("APIErrorCode").InnerText
            fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Call to SubmitTransport() failed with message: " & ex.Message & " [" & errorCode + "/" & detail & "]")
            Return
        End Try

        fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Request submitted with transportID " & result.transportID)

        fax_transportID = result.transportID


        '//////////////////////////////////////////////////////////////////////
        '// STEP #5 : Release the session and its allocated resources
        '//////////////////////////////////////////////////////////////////////


        ' As soon as you call Logout(), the files allocated on the server during this session won't be available
        ' anymore, so keep in mind that former urls are now useless...

        fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Releasing session and server files")

        Try
            session.Logout()
        Catch ex As SoapException
            fax_log.AppendLine(Format(Now, "MM/dd/yy HH:mm:ss") & " " & "Call to Logout() failed with message: " & ex.Message)
            Return
        End Try
    End Sub
End Class