Imports System.IO
Imports System.Reflection
Imports System.Net.Mail
Imports Microsoft.Exchange.WebServices.Data

Public Module TACMAIL1
    Sub New()
    End Sub

    <System.Runtime.CompilerServices.Extension()>
    Public Sub Save(ByVal Message As MailMessage, ByVal FileName As String)
        Dim assembly As Assembly = GetType(SmtpClient).Assembly
        Dim _mailWriterType As Type = assembly.[GetType]("System.Net.Mail.MailWriter")
        Using _fileStream As New FileStream(FileName, FileMode.Create)
            ' Get reflection info for MailWriter contructor
            Dim _mailWriterContructor As ConstructorInfo = _mailWriterType.GetConstructor(BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Type() {GetType(Stream)}, Nothing)

            ' Construct MailWriter object with our FileStream
            Dim _mailWriter As Object = _mailWriterContructor.Invoke(New Object() {_fileStream})

            ' Get reflection info for Send() method on MailMessage
            Dim _sendMethod As MethodInfo = GetType(MailMessage).GetMethod("Send", BindingFlags.Instance Or BindingFlags.NonPublic)

            ' Call method passing in MailWriter
            '_sendMethod.Invoke(Message, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {_mailWriter, True}, Nothing)
            '_sendMethod.Invoke(Message, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {_mailWriter, True, True}, Nothing)

            If ASCMAIN1.Running_in_VS Or (ASCMAIN1.DBS_SERVER = "RGI" Or ASCMAIN1.DBS_COMPANY = "RGI") Or (ASCMAIN1.DBS_SERVER = "NYA" Or ASCMAIN1.DBS_COMPANY = "NYA") Then
                Try
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox("TRYING 1a")
                    _sendMethod.Invoke(Message, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {_mailWriter, True, True}, Nothing)
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox("email.save method 1a")
                Catch ex As Exception
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox("TRYING 1B")
                    _sendMethod.Invoke(Message, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {_mailWriter, True}, Nothing)
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox("email.save method 1b")
                End Try
            Else
                Try
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox("TRYING 2a")
                    _sendMethod.Invoke(Message, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {_mailWriter, True}, Nothing)
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox("email.save method 2a")
                Catch ex As Exception
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox("TRYING 2B")
                    _sendMethod.Invoke(Message, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {_mailWriter, True, True}, Nothing)
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox("email.save method 2b")
                End Try
            End If

            'may need to make the change above if we move to the next version of .Net framework
            ' wierd that I need the extra True parameter on my laptop and not in production
            'http://techsharehub.blogspot.com/2013/11/parameter-count-mismatch.html
            'sendMethod.Invoke(m, BindingFlags.Instance | BindingFlags.NonPublic, null, new[] { mailWriter, true}, null)

            ' Finally get reflection info for Close() method on our MailWriter
            Dim _closeMethod As MethodInfo = _mailWriter.[GetType]().GetMethod("Close", BindingFlags.Instance Or BindingFlags.NonPublic)

            ' Call close method
            _closeMethod.Invoke(_mailWriter, BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, New Object() {}, Nothing)
        End Using
    End Sub

    <System.Runtime.CompilerServices.Extension()>
    Public Sub SaveToFile(ByVal Message As EmailMessage, ByVal FileName As String)
        Message.Load(New PropertySet(ItemSchema.MimeContent))
        Dim mimcon As MimeContent = Message.MimeContent
        Using fStream As New FileStream(FileName, FileMode.Create)
            fStream.Write(mimcon.Content, 0, mimcon.Content.Length)
            fStream.Close()
        End Using
    End Sub

End Module
