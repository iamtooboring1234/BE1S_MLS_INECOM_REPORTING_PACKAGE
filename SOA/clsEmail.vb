'' © Copyright © 2007-2020, Inecom Pte Ltd, All rights reserved.
'' =============================================================

Imports System.IO
Imports System.Environment
Imports System.Net
Imports outlook = Microsoft.Office.Interop.Outlook

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
    Private _EmailSubject As String = ""
    Private _BPName As String = ""
    Private _IsOffice365 As String = "N"
    Private _IsOutlook As String = "N"
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

    Private Function CheckImage1(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image1") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage1] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function CheckImage2(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image2") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage2] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function CheckImage3(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image3") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage3] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function CheckImage4(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image4") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage4] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function CheckImage5(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image5") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage5] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function
    Private Function CheckImage6(ByVal sInput As String) As Boolean
        Try
            If sInput.Contains("image6") Then
                Return True
            End If
            Return False
        Catch ex As Exception
            SBO_Application.MessageBox("[CheckImage6] - " & ex.Message, 1, "Ok", String.Empty, String.Empty)
            Return False
        End Try
    End Function

    Public Function GetSetting(ByVal Code As String) As Boolean
        Try
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = "SELECT IFNULL(""U_MailFrom"",''), IFNULL(""U_SMTP"",''), IFNULL(""U_Username"",''), "
            sQuery &= " IFNULL(""U_Password"",''), IFNULL(""U_AuthType"",0), IFNULL(""U_PortNum"",0) ""PortNumber"", "
            sQuery &= " IFNULL(""U_LocalIP"",'') ""LocalIPAddress"", IFNULL(""U_EnableSSL"",'N') ""EnableSSL"", "
            sQuery &= " IFNULL(""U_EmailPath"",'') ""EmailPath"", IFNULL(""U_Office365"",'N') ""Office365"", "
            sQuery &= " IFNULL(""U_Outlook"",'N') ""Outlook"", IFNULL(""U_EmailSub"",'') ""EmailSubject""  "
            sQuery &= " FROM ""@NCM_EMAIL_CONFIG"" "
            sQuery &= " WHERE ""Code"" = '" & Code & "'"

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
                _IsOutlook = oRecord.Fields.Item("Outlook").Value.ToString.Trim
                _EmailSubject = oRecord.Fields.Item("EmailSubject").Value.ToString.Trim

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
            Dim bImage1 As Boolean = False
            Dim bImage2 As Boolean = False
            Dim bImage3 As Boolean = False
            Dim bImage4 As Boolean = False
            Dim bImage5 As Boolean = False
            Dim bImage6 As Boolean = False

            GetSetting("SOA")
            _EmailCc = GetEmailCCFromUDT()

            Select Case _IsOutlook
                Case "Y"
                    Dim OutlookMessage As outlook.MailItem
                    Dim AppOutlook As New outlook.Application
                    Dim sOutput As String = ""
                    Dim sOutput2 As String = ""
                    Dim propertyAccessor As outlook.PropertyAccessor
                    Dim image1 As outlook.Attachment
                    Dim image2 As outlook.Attachment
                    Dim image3 As outlook.Attachment
                    Dim image4 As outlook.Attachment
                    Dim image5 As outlook.Attachment
                    Dim image6 As outlook.Attachment

                    Dim attachments As outlook.Attachments = Nothing
                    Dim bIsHTML As Boolean = False
                    Dim sFilePath As String = ""

                    If _EmailPath.Trim = "" Then
                        sFilePath = Directory.GetCurrentDirectory & "\EmailBody.html"
                    Else
                        sFilePath = _EmailPath
                    End If

                    OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem)
                    Dim Recipients As outlook.Recipients = OutlookMessage.Recipients
                    Dim s() As String = _EmailTo.Split(";")

                    For i As Integer = 0 To s.Length - 1
                        Recipients.Add(s(i).Trim)
                    Next

                    'Dim Cc() As String
                    'If _EmailCc.Trim.Length > 0 Then
                    '    Cc = _EmailCc.Split(";")
                    '    For i As Integer = 0 To Cc.Length - 1
                    '        Recipients.Add((Cc(i).Trim))
                    '    Next
                    'End If

                    If System.IO.File.Exists(sFilePath) Then
                        OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
                        bIsHTML = True
                        sOutput = System.IO.File.ReadAllText(sFilePath)
                        bImage1 = CheckImage1(sOutput)
                        bImage2 = CheckImage2(sOutput)
                        bImage3 = CheckImage3(sOutput)
                        bImage4 = CheckImage4(sOutput)
                        bImage5 = CheckImage5(sOutput)
                        bImage6 = CheckImage6(sOutput)

                        If sOutput.Contains("{0}") Then
                            sOutput2 = sOutput.Replace("{0}", AsAtDate.ToString("dd/MM/yyyy"))
                        Else
                            sOutput2 = sOutput
                        End If
                    Else
                        OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
                        sOutput2 = "Please refer to attachment."
                    End If

                    If bIsHTML Then
                        Try
                            attachments = OutlookMessage.Attachments

                            If bImage1 And File.Exists(Directory.GetCurrentDirectory & "\image001.gif") Then
                                image1 = attachments.Add(Directory.GetCurrentDirectory & "\image001.gif")
                                propertyAccessor = image1.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image1")
                            End If

                            If bImage2 And File.Exists(Directory.GetCurrentDirectory & "\image002.gif") Then
                                image2 = attachments.Add(Directory.GetCurrentDirectory & "\image002.gif")
                                propertyAccessor = image2.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image2")
                            End If

                            If bImage3 And File.Exists(Directory.GetCurrentDirectory & "\image003.gif") Then
                                image3 = attachments.Add(Directory.GetCurrentDirectory & "\image003.gif")
                                propertyAccessor = image3.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image3")
                            End If

                            If bImage4 And File.Exists(Directory.GetCurrentDirectory & "\image004.gif") Then
                                image4 = attachments.Add(Directory.GetCurrentDirectory & "\image004.gif")
                                propertyAccessor = image4.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image4")
                            End If

                            If bImage5 And File.Exists(Directory.GetCurrentDirectory & "\image005.gif") Then
                                image5 = attachments.Add(Directory.GetCurrentDirectory & "\image005.gif")
                                propertyAccessor = image5.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image5")
                            End If

                            If bImage6 And File.Exists(Directory.GetCurrentDirectory & "\image006.gif") Then
                                image6 = attachments.Add(Directory.GetCurrentDirectory & "\image006.gif")
                                propertyAccessor = image6.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image6")
                            End If

                        Catch ex As Exception

                        End Try
                    End If

                    If _EmailCc.Trim.Length > 0 Then
                        OutlookMessage.CC = _EmailCc
                    End If
                    OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
                    OutlookMessage.HTMLBody = sOutput2
                    OutlookMessage.Attachments.Add(_Attachment)
                    OutlookMessage.Subject = oCompany.CompanyName & " - Statement Of Account - " & _BPName
                    OutlookMessage.Send()

                    attachments = Nothing
                    OutlookMessage = Nothing
                    AppOutlook = Nothing


                Case Else
                    ' VIA SMTP
                    ' ==========================================================
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

                    If System.IO.File.Exists(sFilePath) Then
                        a.IsBodyHtml = True
                        bIsHTML = True
                        sOutput = System.IO.File.ReadAllText(sFilePath)
                        bImage1 = CheckImage1(sOutput)
                        bImage2 = CheckImage2(sOutput)
                        bImage3 = CheckImage3(sOutput)
                        bImage4 = CheckImage4(sOutput)
                        bImage5 = CheckImage5(sOutput)
                        bImage6 = CheckImage6(sOutput)

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
                            Dim bodyAltView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(sOutput2, Nothing, "text/html")
                            Dim imageRes1 As System.Net.Mail.LinkedResource
                            Dim imageRes2 As System.Net.Mail.LinkedResource
                            Dim imageRes3 As System.Net.Mail.LinkedResource
                            Dim imageRes4 As System.Net.Mail.LinkedResource
                            Dim imageRes5 As System.Net.Mail.LinkedResource
                            Dim imageRes6 As System.Net.Mail.LinkedResource

                            Try
                                If bImage1 AndAlso IO.File.Exists("image001.gif") Then
                                    imageRes1 = New System.Net.Mail.LinkedResource("image001.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                                    imageRes1.ContentId = "image1"
                                    imageRes1.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                                    bodyAltView.LinkedResources.Add(imageRes1)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage2 AndAlso IO.File.Exists("image002.gif") Then
                                    imageRes2 = New System.Net.Mail.LinkedResource("image002.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                                    imageRes2.ContentId = "image2"
                                    imageRes2.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                                    bodyAltView.LinkedResources.Add(imageRes2)
                                End If
                            Catch ex As Exception

                            End Try
                            Try
                                If bImage3 AndAlso IO.File.Exists("image003.gif") Then
                                    imageRes3 = New System.Net.Mail.LinkedResource("image003.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                                    imageRes3.ContentId = "image3"
                                    imageRes3.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                                    bodyAltView.LinkedResources.Add(imageRes3)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage4 AndAlso IO.File.Exists("image004.gif") Then
                                    imageRes4 = New System.Net.Mail.LinkedResource("image004.gif", "image/gif")
                                    imageRes4.ContentId = "image4"
                                    imageRes4.TransferEncoding = Net.Mime.TransferEncoding.Base64
                                    bodyAltView.LinkedResources.Add(imageRes4)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage5 AndAlso IO.File.Exists("image005.gif") Then
                                    imageRes5 = New System.Net.Mail.LinkedResource("image005.gif", "image/gif")
                                    imageRes5.ContentId = "image5"
                                    imageRes5.TransferEncoding = Net.Mime.TransferEncoding.Base64
                                    bodyAltView.LinkedResources.Add(imageRes5)

                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage6 AndAlso IO.File.Exists("image006.gif") Then
                                    imageRes6 = New System.Net.Mail.LinkedResource("image006.gif", "image/gif")
                                    imageRes6.ContentId = "image6"
                                    imageRes6.TransferEncoding = Net.Mime.TransferEncoding.Base64
                                    bodyAltView.LinkedResources.Add(imageRes6)

                                End If
                            Catch ex As Exception

                            End Try

                            a.AlternateViews.Add(bodyAltView)
                            a.Body = sOutput2

                        Catch ex As Exception

                        End Try
                    Else
                        a.Body = sOutput2
                    End If

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
            End Select

            Return True
        Catch ex As Exception
            ErrorMessage = ex.Message
            Return False
        End Try
        Return False
    End Function
    Public Function SendPVEmail(ByRef ErrorMessage As String, ByVal AsAtDate As DateTime) As Boolean 'AT Added on 04.05.2019
        Try
            Dim bImage1 As Boolean = False
            Dim bImage2 As Boolean = False
            Dim bImage3 As Boolean = False
            Dim bImage4 As Boolean = False
            Dim bImage5 As Boolean = False
            Dim bImage6 As Boolean = False

            GetSetting("PV")
            _EmailCc = "" 'GetEmailCCFromUDT()

            Select Case _IsOutlook
                Case "Y"
                    Dim OutlookMessage As outlook.MailItem
                    Dim AppOutlook As New outlook.Application
                    Dim sOutput As String = ""
                    Dim sOutput2 As String = ""
                    Dim propertyAccessor As outlook.PropertyAccessor
                    Dim image1 As outlook.Attachment
                    Dim image2 As outlook.Attachment
                    Dim image3 As outlook.Attachment
                    Dim image4 As outlook.Attachment
                    Dim image5 As outlook.Attachment
                    Dim image6 As outlook.Attachment

                    Dim attachments As outlook.Attachments = Nothing
                    Dim bIsHTML As Boolean = False
                    Dim sFilePath As String = ""

                    If _EmailPath.Trim = "" Then
                        sFilePath = Directory.GetCurrentDirectory & "\EmailBodyPV.html"
                    Else
                        sFilePath = _EmailPath
                    End If

                    OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem)
                    Dim Recipients As outlook.Recipients = OutlookMessage.Recipients
                    Dim s() As String = _EmailTo.Split(";")

                    For i As Integer = 0 To s.Length - 1
                        Recipients.Add(s(i).Trim)
                    Next

                    If System.IO.File.Exists(sFilePath) Then
                        OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
                        bIsHTML = True
                        sOutput = System.IO.File.ReadAllText(sFilePath)
                        bImage1 = CheckImage1(sOutput)
                        bImage2 = CheckImage2(sOutput)
                        bImage3 = CheckImage3(sOutput)
                        bImage4 = CheckImage4(sOutput)
                        bImage5 = CheckImage5(sOutput)
                        bImage6 = CheckImage6(sOutput)

                        If sOutput.Contains("{0}") Then
                            sOutput2 = sOutput.Replace("{0}", AsAtDate.ToString("dd/MM/yyyy"))
                        Else
                            sOutput2 = sOutput
                        End If

                        If sOutput2.Contains("{1}") Then
                            sOutput2 = sOutput2.Replace("{1}", _DocNum)
                        Else
                            sOutput2 = sOutput2
                        End If
                    Else
                        OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
                        sOutput2 = "Please refer to attachment."
                    End If

                    If bIsHTML Then
                        Try
                            attachments = OutlookMessage.Attachments

                            If bImage1 And File.Exists(Directory.GetCurrentDirectory & "\ImagePV001.gif") Then
                                image1 = attachments.Add(Directory.GetCurrentDirectory & "\ImagePV001.gif")
                                propertyAccessor = image1.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image1")
                            End If

                            If bImage2 And File.Exists(Directory.GetCurrentDirectory & "\ImagePV002.gif") Then
                                image2 = attachments.Add(Directory.GetCurrentDirectory & "\ImagePV002.gif")
                                propertyAccessor = image2.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image2")
                            End If

                            If bImage3 And File.Exists(Directory.GetCurrentDirectory & "\ImagePV003.gif") Then
                                image3 = attachments.Add(Directory.GetCurrentDirectory & "\ImagePV003.gif")
                                propertyAccessor = image3.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image3")
                            End If

                            If bImage4 And File.Exists(Directory.GetCurrentDirectory & "\ImagePV004.gif") Then
                                image4 = attachments.Add(Directory.GetCurrentDirectory & "\ImagePV004.gif")
                                propertyAccessor = image4.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image4")
                            End If

                            If bImage5 And File.Exists(Directory.GetCurrentDirectory & "\ImagePV005.gif") Then
                                image5 = attachments.Add(Directory.GetCurrentDirectory & "\ImagePV005.gif")
                                propertyAccessor = image5.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image5")
                            End If

                            If bImage6 And File.Exists(Directory.GetCurrentDirectory & "\ImagePV006.gif") Then
                                image6 = attachments.Add(Directory.GetCurrentDirectory & "\ImagePV006.gif")
                                propertyAccessor = image6.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image6")
                            End If

                        Catch ex As Exception

                        End Try
                    End If

                    If _EmailCc.Trim.Length > 0 Then
                        OutlookMessage.CC = _EmailCc
                    End If
                    OutlookMessage.BodyFormat = outlook.OlBodyFormat.olFormatHTML
                    OutlookMessage.HTMLBody = sOutput2
                    OutlookMessage.Attachments.Add(_Attachment)

                    If _EmailSubject.Length > 0 Then
                        OutlookMessage.Subject = _EmailSubject & " - PV No. " & DocNum & " - " & _BPName
                    Else
                        OutlookMessage.Subject = "PV No. " & DocNum & " - " & _BPName
                    End If

                    OutlookMessage.Send()

                    attachments = Nothing
                    OutlookMessage = Nothing
                    AppOutlook = Nothing

                Case Else
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

                    If _EmailSubject.Length > 0 Then
                        a.Subject = _EmailSubject & " - PV No. " & DocNum & " - " & _BPName
                    Else
                        a.Subject = "PV No. " & DocNum & " - " & _BPName       ' HANA
                    End If

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
            End Select

            Return True
        Catch ex As Exception
            ErrorMessage = ex.Message
            Return False
        End Try
        Return False
    End Function

End Class
