Imports System.Environment
Imports System.Net
Imports EASendMail

Public Class clsEmail
    Private _SMTP_Server As String = String.Empty
    Private _EmailFrom As String = String.Empty
    Private _EmailTo As String = String.Empty
    Private _EmailCc As String = String.Empty
    Private _AuthType As String = String.Empty
    Private _Username As String = String.Empty
    Private _Password As String = String.Empty
    Private _Attachment As String = String.Empty
    Private _PortNum As Integer = 0
    Private _LocalIPAddress As String = ""
    Private _EnableSSL As String = ""
    Private _EmailPath As String = ""
    Private _BPName As String = ""
    Private _IsOffice365 As String = "N"
    Private _DocNum As String = ""


    Public Property EnableSSL() As String
        Get
            Return _EnableSSL
        End Get
        Set(ByVal value As String)
            _EnableSSL = value
        End Set
    End Property
    Public Property LocalIPAddress() As String
        Get
            Return _LocalIPAddress
        End Get
        Set(ByVal value As String)
            _LocalIPAddress = value
        End Set
    End Property
    Public Property PortNum() As Integer
        Get
            Return _PortNum
        End Get
        Set(ByVal value As Integer)
            _PortNum = value
        End Set
    End Property
    Public Property SMTP_Server() As String
        Get
            Return _SMTP_Server
        End Get
        Set(ByVal value As String)
            _SMTP_Server = value
        End Set
    End Property
    Public Property EmailFrom() As String
        Get
            Return _EmailFrom
        End Get
        Set(ByVal value As String)
            _EmailFrom = value
        End Set
    End Property
    Public Property EmailTo() As String
        Get
            Return _EmailTo
        End Get
        Set(ByVal value As String)
            _EmailTo = value
        End Set
    End Property
    Public Property EmailCc() As String
        Get
            Return _EmailCc
        End Get
        Set(ByVal value As String)
            _EmailCc = value
        End Set
    End Property
    Public Property AuthorizationType() As String
        Get
            Return _AuthType
        End Get
        Set(ByVal value As String)
            _AuthType = value
        End Set
    End Property
    Public Property Username() As String
        Get
            Return _Username
        End Get
        Set(ByVal value As String)
            _Username = value
        End Set
    End Property
    Public Property Password() As String
        Get
            Return _Password
        End Get
        Set(ByVal value As String)
            _Password = value
        End Set
    End Property
    Public Property Attachment() As String
        Get
            Return _Attachment
        End Get
        Set(ByVal value As String)
            _Attachment = value
        End Set
    End Property
    Public Property CardName() As String
        Get
            Return _BPName
        End Get
        Set(ByVal value As String)
            _BPName = value
        End Set
    End Property
    Public Property DocNum() As String
        Get
            Return _DocNum
        End Get
        Set(ByVal value As String)
            _DocNum = value
        End Set
    End Property

    Public Function GetSetting(ByVal Code As String) As Boolean
        Try
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = "SELECT IFNULL(""U_MailFrom"",''), IFNULL(""U_SMTP"",''), IFNULL(""U_Username"",''), IFNULL(""U_Password"",''), IFNULL(""U_AuthType"",0), IFNULL(""U_PortNum"",0) ""PortNumber"", IFNULL(""U_LocalIP"",'') ""LocalIPAddress"", IFNULL(""U_EnableSSL"",'N') ""EnableSSL"", IFNULL(""U_EmailPath"",'') ""EmailPath"", IFNULL(""U_Office365"",'N') ""Office365""  FROM ""@NCM_EMAIL_CONFIG"" WHERE ""Code"" = '" & Code & "'"
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                _EmailFrom = oRecord.Fields.Item(0).Value
                _SMTP_Server = oRecord.Fields.Item(1).Value
                _Username = oRecord.Fields.Item(2).Value
                _Password = oRecord.Fields.Item(3).Value
                _AuthType = oRecord.Fields.Item(4).Value
                _PortNum = oRecord.Fields.Item("PortNumber").Value
                _LocalIPAddress = oRecord.Fields.Item("LocalIPAddress").Value.ToString.Trim
                _EnableSSL = oRecord.Fields.Item("EnableSSL").Value.ToString.Trim
                _EmailPath = oRecord.Fields.Item("EmailPath").Value.ToString.Trim
                _IsOffice365 = oRecord.Fields.Item("Office365").Value.ToString.Trim

                Return True
            Else
                SBO_Application.MessageBox("[clsEmail].[GetSetting] - Please configure email setting.", 1, "OK", String.Empty, String.Empty)
                Return False
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[clsEmail].[GetSetting] - " & ex.Message, 1, "OK", String.Empty, String.Empty)
            Return False
        End Try
        Return False
    End Function
    Public Function SendEmail(ByRef ErrorMessage As String, ByVal AsAtDate As DateTime) As Boolean
        Try
            GetSetting("SOA")
            _EmailCc = GetEmailCCFromUDT()

            'Select Case _IsOffice365
            '    Case "Y"
            'Dim s() As String = _EmailTo.Split(";")
            'Dim sOutput As String = String.Empty
            'Dim sOutput2 As String = String.Empty
            'Dim bIsHTML As Boolean = False
            'Dim sFilePath As String = ""

            'If _EmailPath.Trim = "" Then
            '    sFilePath = "EmailBody.html"
            'Else
            '    sFilePath = _EmailPath
            'End If

            'Dim oMail As New SmtpMail("TryIt")
            'Dim oSmtp As New SmtpClient()

            '' Your hotmail/outlook email address
            'oMail.From = _EmailFrom

            '' Set recipient email address, please change it to yours
            'For i As Integer = 0 To s.Length - 1
            '    oMail.To.Add((s(i).Trim))
            'Next

            'Dim Cc() As String
            'If _EmailCc.Trim.Length > 0 Then
            '    Cc = _EmailCc.Split(";")
            '    For i As Integer = 0 To Cc.Length - 1
            '        oMail.Cc.Add((Cc(i).Trim))
            '    Next
            'End If

            '' Set email subject
            'oMail.Subject = oCompany.CompanyName & " - Statement Of Account - " & _BPName

            '' Set email body
            'If System.IO.File.Exists(sFilePath) Then
            '    bIsHTML = True
            '    sOutput = System.IO.File.ReadAllText(sFilePath)
            '    If sOutput.Contains("{0}") Then
            '        sOutput2 = sOutput.Replace("{0}", AsAtDate.ToString("dd/MM/yyyy"))
            '    Else
            '        sOutput2 = sOutput
            '    End If
            'Else
            '    sOutput2 = "Please refer to attachment."
            'End If

            'If bIsHTML Then
            '    oMail.HtmlBody = sOutput2
            'Else
            '    oMail.TextBody = sOutput2
            'End If

            'oMail.AddAttachment(_Attachment)

            ' Hotmail/Outlook SMTP server address
            'Dim oServer As New SmtpServer(_SMTP_Server)

            'oServer.User = _Username
            'oServer.Password = _Password

            '' use 587 port
            'oServer.Port = _PortNum

            '' detect SSL/TLS connection automatically
            'oServer.ConnectType = SmtpConnectType.ConnectSSLAuto

            'oSmtp.SendMail(oServer, oMail)

            'Case "N"

            Dim tmpMailFr As New System.Net.Mail.MailAddress(_EmailFrom)
            Dim s() As String = _EmailTo.Split(";")
            Dim a As New System.Net.Mail.MailMessage()
            Dim sOutput As String = String.Empty
            Dim sOutput2 As String = String.Empty
            Dim bIsHTML As Boolean = False
            Dim sFilePath As String = ""

            If _EmailPath.Trim = "" Then
                sFilePath = "EmailBody.html"
            Else
                sFilePath = _EmailPath
            End If

            a.From = tmpMailFr
            For i As Integer = 0 To s.Length - 1
                a.To.Add(s(i))
            Next

            Dim Cc() As String
            If _EmailCc.Trim.Length > 0 Then
                Cc = _EmailCc.Split(";")
                For i As Integer = 0 To Cc.Length - 1
                    a.CC.Add((Cc(i).Trim))
                Next
            End If

            a.Subject = oCompany.CompanyName & " - Statement Of Account - " & _BPName       ' HANA
            'a.Subject = oCompany.CompanyName & " - " & _BPName & " - Statement Of Account " ' SQL

            If System.IO.File.Exists(sFilePath) Then
                a.IsBodyHtml = True
                bIsHTML = True
                sOutput = System.IO.File.ReadAllText(sFilePath)
                If sOutput.Contains("{0}") Then
                    sOutput2 = sOutput.Replace("{0}", AsAtDate.ToString("dd/MM/yyyy"))
                Else
                    sOutput2 = sOutput
                End If
            Else
                sOutput2 = "Please refer to attachment."
            End If

            If bIsHTML Then
                Try
                    If _EmailPath = "" Then
                        If IO.File.Exists("image001.gif") Then
                            Dim bodyAltView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(sOutput2, Nothing, "text/html")
                            Dim imageResourceEs1 As New System.Net.Mail.LinkedResource("image001.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                            imageResourceEs1.ContentId = "image1"
                            imageResourceEs1.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                            bodyAltView.LinkedResources.Add(imageResourceEs1)
                            a.AlternateViews.Add(bodyAltView)
                        End If
                    End If
                Catch ex As Exception

                End Try
            End If

            a.Body = sOutput2

            Dim b As New System.Net.Mail.Attachment(_Attachment)
            a.Attachments.Add(b)

            Dim c As New System.Net.Mail.SmtpClient(_SMTP_Server)

            If _AuthType = "1" Then
                Dim d As New System.Net.NetworkCredential(_Username, _Password)

                ' START - added by ES 22.01.2016 for future enhancement.
                If _LocalIPAddress.Trim <> "" Then
                    c.Host = _LocalIPAddress
                End If

                If _IsOffice365 = "Y" Then
                    c.EnableSsl = True
                Else
                    If _EnableSSL = "Y" Then
                        c.EnableSsl = True
                    Else
                        c.EnableSsl = False
                    End If
                End If

                If _PortNum > 0 Then
                    c.Port = _PortNum
                End If
                ' END - added by ES 22.01.2016 for future enhancement.

                c.DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network
                c.UseDefaultCredentials = False
                c.Credentials = d
            End If

            c.Send(a)
            b.Dispose()

            c = Nothing
            a = Nothing
            b = Nothing

            ' End Select

            Return True
        Catch ex As Exception
            ErrorMessage = ex.Message
            Return False
        End Try
        Return False
    End Function
    Public Function SendPVEmail(ByRef ErrorMessage As String, ByVal AsAtDate As DateTime) As Boolean 'AT Added on 04.05.2019
        Try
            GetSetting("PV")
            _EmailCc = "" 'GetEmailCCFromUDT()

            Dim tmpMailFr As New System.Net.Mail.MailAddress(_EmailFrom)
            Dim s() As String = _EmailTo.Split(";")
            Dim a As New System.Net.Mail.MailMessage()
            Dim sOutput As String = String.Empty
            Dim sOutput2 As String = String.Empty
            Dim bIsHTML As Boolean = False
            Dim sFilePath As String = ""

            If _EmailPath.Trim = "" Then
                sFilePath = "EmailBodyPV.html"
            Else
                sFilePath = _EmailPath
            End If

            a.From = tmpMailFr
            For i As Integer = 0 To s.Length - 1
                a.To.Add(s(i))
            Next

            Dim Cc() As String
            If _EmailCc.Trim.Length > 0 Then
                Cc = _EmailCc.Split(";")
                For i As Integer = 0 To Cc.Length - 1
                    a.CC.Add((Cc(i).Trim))
                Next
            End If

            a.Subject = oCompany.CompanyName & " - PV No. " & DocNum & " - " & _BPName       ' HANA
            'a.Subject = oCompany.CompanyName & " - " & _BPName & " - Statement Of Account " ' SQL

            If System.IO.File.Exists(sFilePath) Then
                a.IsBodyHtml = True
                bIsHTML = True
                sOutput = System.IO.File.ReadAllText(sFilePath)
                If sOutput.Contains("{0}") Then
                    sOutput2 = sOutput.Replace("{0}", DocNum)
                Else
                    sOutput2 = sOutput
                End If
            Else
                sOutput2 = "Please refer to attachment."
            End If

            If bIsHTML Then
                Try
                    If _EmailPath = "" Then
                        If IO.File.Exists("ImagePV001.gif") Then
                            Dim bodyAltView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(sOutput2, Nothing, "text/html")
                            Dim imageResourceEs1 As New System.Net.Mail.LinkedResource("ImagePV001.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                            imageResourceEs1.ContentId = "image1"
                            imageResourceEs1.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                            bodyAltView.LinkedResources.Add(imageResourceEs1)
                            a.AlternateViews.Add(bodyAltView)
                        End If
                    End If
                Catch ex As Exception

                End Try
            End If

            a.Body = sOutput2

            Dim b As New System.Net.Mail.Attachment(_Attachment)
            a.Attachments.Add(b)

            Dim c As New System.Net.Mail.SmtpClient(_SMTP_Server)

            If _AuthType = "1" Then
                Dim d As New System.Net.NetworkCredential(_Username, _Password)

                ' START - added by ES 22.01.2016 for future enhancement.
                If _LocalIPAddress.Trim <> "" Then
                    c.Host = _LocalIPAddress
                End If

                If _IsOffice365 = "Y" Then
                    c.EnableSsl = True
                Else
                    If _EnableSSL = "Y" Then
                        c.EnableSsl = True
                    Else
                        c.EnableSsl = False
                    End If
                End If

                If _PortNum > 0 Then
                    c.Port = _PortNum
                End If
                ' END - added by ES 22.01.2016 for future enhancement.

                c.DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network
                c.UseDefaultCredentials = False
                c.Credentials = d
            End If

            c.Send(a)
            b.Dispose()

            c = Nothing
            a = Nothing
            b = Nothing

            ' End Select

            Return True
        Catch ex As Exception
            ErrorMessage = ex.Message
            Return False
        End Try
        Return False
    End Function

End Class
