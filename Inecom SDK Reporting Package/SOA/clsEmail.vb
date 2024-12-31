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
    Private _Attachment2 As String = String.Empty
    Private _Attachment3 As String = String.Empty
    Private _Attachment4 As String = String.Empty
    Private _EmailType As String = ""
    Private _BPName As String = ""
    Private _BPCode As String = ""

    Private _PlainText As String = ""
    Private _PortNum As Integer = 0
    Private _LocalIPAddress As String = ""
    Private _EnableSSL As String = ""
    Private _EmailPath As String = ""
    Private _EmailSubject As String = ""
    Private _IsOffice365 As String = "N"
    Private _IsOutlook As String = "N"
    Private _IsGeneric As String = ""
    Private _DocNum As String = ""
    Private _ReportType As String = "ARSOA"
    Private _CardOption As String = "C"

    Private _ARSOA As String = ""
    Private _ARINV As String = ""
    Private _ARDPI As String = ""
    Private _ARRIN As String = ""
    Private _PAYPV As String = ""
    Private _PAYRA As String = ""
    Private _DODEL As String = "" 'SY add on 12112020
    Private _DOUNDEL As String = "" 'SY add on 12112020
    Private _ARSOA_EmailSubject As String = ""

    Public Property ReportType() As String
        Get
            Return _ReportType
        End Get
        Set(ByVal value As String)
            _ReportType = value
        End Set
    End Property
    Public Property IsGeneric() As String
        Get
            Return _IsGeneric
        End Get
        Set(ByVal value As String)
            _IsGeneric = value
        End Set
    End Property
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
    Public Property Attachment2() As String
        Get
            Return _Attachment2
        End Get
        Set(ByVal value As String)
            _Attachment2 = value
        End Set
    End Property
    Public Property Attachment3() As String
        Get
            Return _Attachment3
        End Get
        Set(ByVal value As String)
            _Attachment3 = value
        End Set
    End Property
    Public Property Attachment4() As String
        Get
            Return _Attachment4
        End Get
        Set(ByVal value As String)
            _Attachment4 = value
        End Set
    End Property

    Public Property CardCode() As String
        Get
            Return _BPCode
        End Get
        Set(ByVal value As String)
            _BPCode = value
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
            Dim sQueryEmail As String = ""
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            Dim sQuery As String = "SELECT IFNULL(""U_MailFrom"",''), IFNULL(""U_SMTP"",''), IFNULL(""U_Username"",''), "
            sQuery &= " IFNULL(""U_Password"",''), IFNULL(""U_AuthType"",0), IFNULL(""U_PortNum"",0) ""PortNumber"", "
            sQuery &= " IFNULL(""U_LocalIP"",'') ""LocalIPAddress"", IFNULL(""U_EnableSSL"",'N') ""EnableSSL"", "
            sQuery &= " IFNULL(""U_EmailPath"",'') ""EmailPath"", IFNULL(""U_Office365"",'N') ""Office365"", "
            sQuery &= " IFNULL(""U_Outlook"",'N') ""Outlook"", IFNULL(""U_EmailSub"",'') ""EmailSubject"" "
            sQuery &= " FROM ""@NCM_EMAIL_CONFIG"" "
            sQuery &= " WHERE ""Code"" = '" & Code & "'"

            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                _EmailFrom = oRecord.Fields.Item(0).Value.ToString.Trim
                _SMTP_Server = oRecord.Fields.Item(1).Value.ToString.Trim
                _Username = oRecord.Fields.Item(2).Value.ToString.Trim
                _Password = oRecord.Fields.Item(3).Value.ToString.Trim
                _AuthType = oRecord.Fields.Item(4).Value.ToString.Trim
                _PortNum = oRecord.Fields.Item("PortNumber").Value.ToString.Trim
                _LocalIPAddress = oRecord.Fields.Item("LocalIPAddress").Value.ToString.Trim
                _EnableSSL = oRecord.Fields.Item("EnableSSL").Value.ToString.Trim
                _EmailPath = oRecord.Fields.Item("EmailPath").Value.ToString.Trim
                _IsOffice365 = oRecord.Fields.Item("Office365").Value.ToString.Trim
                _IsOutlook = oRecord.Fields.Item("Outlook").Value.ToString.Trim
                _EmailSubject = oRecord.Fields.Item("EmailSubject").Value.ToString.Trim
            Else
                SBO_Application.MessageBox("[GetSetting] - Please configure email setting.", 1, "OK", String.Empty, String.Empty)
                Return False
            End If

            If Code = "PV" Then
                _PlainText = "Please refer to the attached Payment Voucher file for your kind perusal."
                _EmailType = "H"
                _CardOption = ""
            Else
                Select Case _ReportType
                    Case "ARSOA"
                        sQueryEmail = " SELECT IFNULL(""U_PlainText"",''),    IFNULL(""U_EmailType"",'H'),    IFNULL(""U_CardOption"",'C') FROM ""@NCM_NEW_SETTING"" "
                    Case "ARINV"
                        sQueryEmail = " SELECT IFNULL(""U_InvPlainText"",''), IFNULL(""U_InvEmailType"",'H'), IFNULL(""U_CardOption"",'C') FROM ""@NCM_NEW_SETTING"" "
                    Case "ARRIN"
                        sQueryEmail = " SELECT IFNULL(""U_RinPlainText"",''), IFNULL(""U_RinEmailType"",'H'), IFNULL(""U_CardOption"",'C') FROM ""@NCM_NEW_SETTING"" "
                    Case "ARDPI"
                        sQueryEmail = " SELECT IFNULL(""U_DpiPlainText"",''), IFNULL(""U_DpiEmailType"",'H'), IFNULL(""U_CardOption"",'C') FROM ""@NCM_NEW_SETTING"" "
                    Case Else
                        sQueryEmail = " SELECT IFNULL(""U_PlainText"",''),    IFNULL(""U_EmailType"",'H'),    IFNULL(""U_CardOption"",'C') FROM ""@NCM_NEW_SETTING"" "
                End Select
                oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQueryEmail)
                If oRec.RecordCount > 0 Then
                    oRec.MoveFirst()
                    _PlainText = oRec.Fields.Item(0).Value.ToString.Trim
                    _EmailType = oRec.Fields.Item(1).Value.ToString.Trim
                    _CardOption = oRec.Fields.Item(2).Value.ToString.Trim
                Else
                    SBO_Application.MessageBox("[GetSetting] - Please configure email type in IRP Configuration.", 1, "Ok", String.Empty, String.Empty)
                    Return False
                End If
            End If

            oRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'SY added U_DODEL and U_DOUDEL 
            oRec.DoQuery("SELECT IFNULL(""U_ARSOA"",''), IFNULL(""U_ARINV"",''), IFNULL(""U_ARRIN"",''), IFNULL(""U_ARDPI"",''), IFNULL(""U_PAYPV"",''), IFNULL(""U_PAYRA"",'') , IFNULL(""U_DODEL"",''), IFNULL(""U_DOUDEL"",'') FROM ""@NCM_EMAIL_HTML"" ")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                _ARSOA = oRec.Fields.Item(0).Value.ToString.Trim
                _ARINV = oRec.Fields.Item(1).Value.ToString.Trim
                _ARRIN = oRec.Fields.Item(2).Value.ToString.Trim
                _ARDPI = oRec.Fields.Item(3).Value.ToString.Trim
                _PAYPV = oRec.Fields.Item(4).Value.ToString.Trim
                _PAYRA = oRec.Fields.Item(5).Value.ToString.Trim
                _DODEL = oRec.Fields.Item(6).Value.ToString.Trim 'SY added 
                _DOUNDEL = oRec.Fields.Item(7).Value.ToString.Trim 'SY added 
            Else
                SBO_Application.MessageBox("[GetSetting] - Please configure email HTML in Email Configuration.", 1, "Ok", String.Empty, String.Empty)
                Return False
            End If

        Catch ex As Exception
            SBO_Application.MessageBox("[GetSetting] - " & ex.Message, 1, "OK", String.Empty, String.Empty)
            Return False
        End Try
        Return False
    End Function
    Public Function SendEmail_INV(ByRef ErrorMessage As String, Optional ByVal AsAtDate As DateTime = Nothing, Optional ByVal DocType As String = "ARINV", Optional ByVal DocNum As String = "") As Boolean
        Try
            Dim bImage1 As Boolean = False
            Dim bImage2 As Boolean = False
            Dim bImage3 As Boolean = False
            Dim bImage4 As Boolean = False
            Dim bImage5 As Boolean = False
            Dim bImage6 As Boolean = False
            Dim sFilePath As String = ""

            GetSetting("SOA")
            _EmailCc = GetEmailCCFromUDT()

            Select Case DocType
                Case "ARINV"
                    If _ARINV.Trim = "" Then
                        sFilePath = Directory.GetCurrentDirectory & "\EmailBody.html"
                    Else
                        sFilePath = _ARINV
                    End If
                Case "ARRIN"
                    If _ARRIN.Trim = "" Then
                        sFilePath = Directory.GetCurrentDirectory & "\EmailBody.html"
                    Else
                        sFilePath = _ARRIN
                    End If
                Case "ARDPI"
                    If _ARDPI.Trim = "" Then
                        sFilePath = Directory.GetCurrentDirectory & "\EmailBody.html"
                    Else
                        sFilePath = _ARDPI
                    End If
                Case "ARDO"
                    If _DODEL.Trim = "" Then
                        sFilePath = Directory.GetCurrentDirectory & "\EmailBody.html"
                    Else
                        sFilePath = _DODEL
                    End If
            End Select

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


                    OutlookMessage = AppOutlook.CreateItem(outlook.OlItemType.olMailItem)
                    Dim Recipients As outlook.Recipients = OutlookMessage.Recipients
                    Dim s() As String = _EmailTo.Split(";")

                    For i As Integer = 0 To s.Length - 1
                        Recipients.Add(s(i).Trim)
                    Next

                    bIsHTML = False
                    Select Case _EmailType
                        Case "H"
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
                        Case Else
                            sOutput2 = _PlainText
                    End Select

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

                    Try
                        If _Attachment2.Trim.Length > 0 Then
                            OutlookMessage.Attachments.Add(_Attachment2)
                        End If

                        If _Attachment3.Trim.Length > 0 Then
                            OutlookMessage.Attachments.Add(_Attachment3)
                        End If

                        If _Attachment4.Trim.Length > 0 Then
                            OutlookMessage.Attachments.Add(_Attachment4)
                        End If
                    Catch ex As Exception

                    End Try

                    Select Case DocType
                        Case "ARINV"
                            OutlookMessage.Subject = oCompany.CompanyName & " - Electronic Invoice No. - " & DocNum
                        Case "ARRIN"
                            OutlookMessage.Subject = oCompany.CompanyName & " - Electronic Credit Note No. - " & DocNum
                        Case "ARDPI"
                            OutlookMessage.Subject = oCompany.CompanyName & " - Electronic DP Invoice No. - " & DocNum
                        Case "ARDO"
                            OutlookMessage.Subject = oCompany.CompanyName & " - Electronic Delivery Order No. - " & DocNum
                    End Select

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

                    Select Case DocType
                        Case "ARINV"
                            a.Subject = oCompany.CompanyName & " - Electronic Invoice No. - " & DocNum
                        Case "ARRIN"
                            a.Subject = oCompany.CompanyName & " - Electronic Credit Note No. - " & DocNum
                        Case "ARDPI"
                            a.Subject = oCompany.CompanyName & " - Electronic DP Invoice No. - " & DocNum
                        Case "ARDO"
                            a.Subject = oCompany.CompanyName & " - Electronic Delivery Order No. - " & DocNum
                    End Select

                    Select Case _EmailType
                        Case "H"
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
                        Case Else
                            sOutput2 = _PlainText
                    End Select

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

                    Try
                        If _Attachment2.Trim.Length > 0 Then
                            Dim b2 As New System.Net.Mail.Attachment(_Attachment2)
                            a.Attachments.Add(b2)
                        End If

                        If _Attachment3.Trim.Length > 0 Then
                            Dim b3 As New System.Net.Mail.Attachment(_Attachment3)
                            a.Attachments.Add(b3)
                        End If

                        If _Attachment4.Trim.Length > 0 Then
                            Dim b4 As New System.Net.Mail.Attachment(_Attachment4)
                            a.Attachments.Add(b4)
                        End If
                    Catch ex As Exception

                    End Try

                    Dim c As New System.Net.Mail.SmtpClient(_SMTP_Server)

                    If _AuthType = "1" Then
                        Dim d As New System.Net.NetworkCredential(_Username, _Password)

                        ' START - added by ES 22.01.2016 for future enhancement.
                        If _LocalIPAddress.Trim <> "" Then
                            c.Host = _LocalIPAddress
                        End If

                        If _IsOffice365 = "Y" Then
                            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                            c.EnableSsl = True
                        Else
                            If _EnableSSL = "Y" Then
                                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
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
    Public Function SendEmail(ByRef ErrorMessage As String, Optional ByVal AsAtDate As DateTime = Nothing, Optional ByVal DocType As String = "SOA", Optional ByVal DocNum As String = "") As Boolean
        Try
            Dim bImage1 As Boolean = False
            Dim bImage2 As Boolean = False
            Dim bImage3 As Boolean = False
            Dim bImage4 As Boolean = False
            Dim bImage5 As Boolean = False
            Dim bImage6 As Boolean = False
            Dim sDirectoryName As String = ""

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

                        Dim fi As New IO.FileInfo(sFilePath)
                        sDirectoryName = fi.DirectoryName

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

                            If bImage1 And File.Exists(sDirectoryName & "\image001.gif") Then
                                image1 = attachments.Add(sDirectoryName & "\image001.gif")
                                propertyAccessor = image1.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image1")
                            End If

                            If bImage2 And File.Exists(sDirectoryName & "\image002.gif") Then
                                image2 = attachments.Add(sDirectoryName & "\image002.gif")
                                propertyAccessor = image2.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image2")
                            End If

                            If bImage3 And File.Exists(sDirectoryName & "\image003.gif") Then
                                image3 = attachments.Add(sDirectoryName & "\image003.gif")
                                propertyAccessor = image3.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image3")
                            End If

                            If bImage4 And File.Exists(sDirectoryName & "\image004.gif") Then
                                image4 = attachments.Add(sDirectoryName & "\image004.gif")
                                propertyAccessor = image4.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image4")
                            End If

                            If bImage5 And File.Exists(sDirectoryName & "\image005.gif") Then
                                image5 = attachments.Add(sDirectoryName & "\image005.gif")
                                propertyAccessor = image5.PropertyAccessor
                                propertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "image5")
                            End If

                            If bImage6 And File.Exists(sDirectoryName & "\image006.gif") Then
                                image6 = attachments.Add(sDirectoryName & "\image006.gif")
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

                    If _EmailSubject.Trim <> "" Then
                        _ARSOA_EmailSubject = _EmailSubject.Trim
                        _ARSOA_EmailSubject = _ARSOA_EmailSubject.Replace("<<COMPANYNAME>>", oCompany.CompanyName)
                        _ARSOA_EmailSubject = _ARSOA_EmailSubject.Replace("<<BPCODE>>", _BPCode)
                        _ARSOA_EmailSubject = _ARSOA_EmailSubject.Replace("<<BPNAME>>", _BPName)
                        _ARSOA_EmailSubject = _ARSOA_EmailSubject.Replace("<<TITLE>>", "Statement Of Account")
                        'OutlookMessage.Subject = oCompany.CompanyName & " - " & _BPName & " - Statement Of Account "
                        OutlookMessage.Subject = _ARSOA_EmailSubject
                    Else
                        If _CardOption.Trim.ToUpper = "C" Then
                            _BPName = _BPCode
                        Else
                            _BPName = _BPName
                        End If

                        OutlookMessage.Subject = oCompany.CompanyName & " - " & _BPName & " - Statement Of Account "
                    End If

                    If _CardOption.Trim.ToUpper = "C" Then
                        _BPName = _BPCode
                    Else
                        _BPName = _BPName
                    End If
                    ' ======================================================================
                    ' For LCS - 
                    ' <LCS SOA – APR’22 – {customer full name}> as I had made some editing.
                    ' FEB'22 SOA - ....
                    ' Added since V910.146.2023...
                    ' ======================================================================
                    If oCompany.CompanyDB.ToString.Trim.ToUpper.Contains("LAUCHOYSENG") Or oCompany.CompanyDB.ToString.Trim.ToUpper.Contains("LAU CHOY SENG") Or oCompany.CompanyDB.ToString.Trim.ToUpper.Contains("LCS") Then
                        Dim sMonth As String = ""
                        Dim sYear As String = ""

                        sYear = AsAtDate.Year.ToString.Substring(2, 2)
                        Select Case AsAtDate.Month
                            Case "1", "01"
                                sMonth = "JAN'"
                            Case "2", "02"
                                sMonth = "FEB'"
                            Case "3", "03"
                                sMonth = "MAR'"
                            Case "4", "04"
                                sMonth = "APR'"
                            Case "5", "05"
                                sMonth = "MAY'"
                            Case "6", "06"
                                sMonth = "JUN'"
                            Case "7", "07"
                                sMonth = "JUL'"
                            Case "8", "08"
                                sMonth = "AUG'"
                            Case "9", "09"
                                sMonth = "SEP'"
                            Case "10", "10"
                                sMonth = "OCT'"
                            Case "11", "11"
                                sMonth = "NOV'"
                            Case "12", "12"
                                sMonth = "DEC'"
                        End Select

                        OutlookMessage.Subject = "<LCS SOA - " & sMonth & sYear & " - " & _BPName.Trim.ToUpper & ">" ' HANA
                    End If
                    ' ======================================================================

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
                        sFilePath = Directory.GetCurrentDirectory & "\EmailBody.html"
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

                    If _CardOption.Trim.ToUpper = "C" Then
                        _BPName = _BPCode
                    Else
                        _BPName = _BPName
                    End If

                    'a.Subject = oCompany.CompanyName & " - Statement Of Account - " & _BPName       ' HANA
                    'a.Subject = oCompany.CompanyName & " - " & _BPName & " - Statement Of Account "

                    If _EmailSubject.Trim <> "" Then
                        _ARSOA_EmailSubject = _EmailSubject.Trim
                        _ARSOA_EmailSubject = _ARSOA_EmailSubject.Replace("<<COMPANYNAME>>", oCompany.CompanyName)
                        _ARSOA_EmailSubject = _ARSOA_EmailSubject.Replace("<<BPCODE>>", _BPCode)
                        _ARSOA_EmailSubject = _ARSOA_EmailSubject.Replace("<<BPNAME>>", _BPName)
                        _ARSOA_EmailSubject = _ARSOA_EmailSubject.Replace("<<TITLE>>", "Statement Of Account")
                        'a.Subject = oCompany.CompanyName & " - " & _BPName & " - Statement Of Account"
                        a.Subject = _ARSOA_EmailSubject
                    Else
                        If _CardOption.Trim.ToUpper = "C" Then
                            _BPName = _BPCode
                        Else
                            _BPName = _BPName
                        End If

                        a.Subject = oCompany.CompanyName & " - " & _BPName & " - Statement Of Account"
                    End If

                    If _CardOption.Trim.ToUpper = "C" Then
                        _BPName = _BPCode
                    Else
                        _BPName = _BPName
                    End If

                    ' ====================================================================================
                    ' For LCS - 
                    ' <LCS SOA – APR’22 – {customer full name}> as I had made some editing.
                    ' FEB'22 SOA - ....
                    ' Added since V910.146.2023...
                    ' ====================================================================================
                    If oCompany.CompanyDB.ToString.Trim.ToUpper.Contains("LAUCHOYSENG") Or oCompany.CompanyDB.ToString.Trim.ToUpper.Contains("LAU CHOY SENG") Or oCompany.CompanyDB.ToString.Trim.ToUpper.Contains("LCS") Then
                        Dim sMonth As String = ""
                        Dim sYear As String = ""

                        sYear = AsAtDate.Year.ToString.Substring(2, 2)
                        Select Case AsAtDate.Month
                            Case "1", "01"
                                sMonth = "JAN'"
                            Case "2", "02"
                                sMonth = "FEB'"
                            Case "3", "03"
                                sMonth = "MAR'"
                            Case "4", "04"
                                sMonth = "APR'"
                            Case "5", "05"
                                sMonth = "MAY'"
                            Case "6", "06"
                                sMonth = "JUN'"
                            Case "7", "07"
                                sMonth = "JUL'"
                            Case "8", "08"
                                sMonth = "AUG'"
                            Case "9", "09"
                                sMonth = "SEP'"
                            Case "10", "10"
                                sMonth = "OCT'"
                            Case "11", "11"
                                sMonth = "NOV'"
                            Case "12", "12"
                                sMonth = "DEC'"
                        End Select
                        a.Subject = "<LCS SOA - " & sMonth & sYear & " - " & _BPName.Trim.ToUpper & ">" ' HANA
                    End If
                    ' ====================================================================================

                    bIsHTML = False
                    Select Case _EmailType
                        Case "H"
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

                                Dim fi As New IO.FileInfo(sFilePath)
                                sDirectoryName = fi.DirectoryName

                                If sOutput.Contains("{0}") Then
                                    sOutput2 = sOutput.Replace("{0}", AsAtDate.ToString("dd/MM/yyyy"))
                                Else
                                    sOutput2 = sOutput
                                End If
                            Else
                                sOutput2 = "Please refer to attachment."
                            End If
                        Case Else
                            sOutput2 = _PlainText
                    End Select

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
                                If bImage1 AndAlso IO.File.Exists(sDirectoryName & "\" & "image001.gif") Then
                                    imageRes1 = New System.Net.Mail.LinkedResource(sDirectoryName & "\" & "image001.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                                    imageRes1.ContentId = "image1"
                                    imageRes1.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                                    bodyAltView.LinkedResources.Add(imageRes1)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage2 AndAlso IO.File.Exists(sDirectoryName & "\" & "image002.gif") Then
                                    imageRes2 = New System.Net.Mail.LinkedResource(sDirectoryName & "\" & "image002.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                                    imageRes2.ContentId = "image2"
                                    imageRes2.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                                    bodyAltView.LinkedResources.Add(imageRes2)
                                End If
                            Catch ex As Exception

                            End Try
                            Try
                                If bImage3 AndAlso IO.File.Exists(sDirectoryName & "\" & "image003.gif") Then
                                    imageRes3 = New System.Net.Mail.LinkedResource(sDirectoryName & "\" & "image003.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                                    imageRes3.ContentId = "image3"
                                    imageRes3.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                                    bodyAltView.LinkedResources.Add(imageRes3)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage4 AndAlso IO.File.Exists(sDirectoryName & "\" & "image004.gif") Then
                                    imageRes4 = New System.Net.Mail.LinkedResource(sDirectoryName & "\" & "image004.gif", "image/gif")
                                    imageRes4.ContentId = "image4"
                                    imageRes4.TransferEncoding = Net.Mime.TransferEncoding.Base64
                                    bodyAltView.LinkedResources.Add(imageRes4)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage5 AndAlso IO.File.Exists(sDirectoryName & "\" & "image005.gif") Then
                                    imageRes5 = New System.Net.Mail.LinkedResource(sDirectoryName & "\" & "image005.gif", "image/gif")
                                    imageRes5.ContentId = "image5"
                                    imageRes5.TransferEncoding = Net.Mime.TransferEncoding.Base64
                                    bodyAltView.LinkedResources.Add(imageRes5)

                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage6 AndAlso IO.File.Exists(sDirectoryName & "\" & "image006.gif") Then
                                    imageRes6 = New System.Net.Mail.LinkedResource(sDirectoryName & "\" & "image006.gif", "image/gif")
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
                            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                            c.EnableSsl = True
                        Else
                            If _EnableSSL = "Y" Then
                                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
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

                    If _EmailSubject.Trim.Length > 0 Then
                        _EmailSubject = _EmailSubject.Trim.Replace("<<COMPANYNAME>>", oCompany.CompanyName)
                        _EmailSubject = _EmailSubject.Trim.Replace("<<DOCUMENTNUM>>", DocNum)
                        _EmailSubject = _EmailSubject.Trim.Replace("<<BPNAME>>", _BPName)
                        _EmailSubject = _EmailSubject.Trim.Replace("<<BPCODE>>", _BPCode)
                        _EmailSubject = _EmailSubject.Trim.Replace("<<TITLE>>", "Payment Voucher")

                        OutlookMessage.Subject = _EmailSubject
                    Else
                        OutlookMessage.Subject = "Payment Voucher No. " & DocNum & " - " & _BPName
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

                    'If _EmailSubject.Length > 0 Then
                    '    a.Subject = _EmailSubject & " - Payment Voucher No. " & DocNum & " - " & _BPName
                    'Else
                    '    a.Subject = "Payment Voucher No. " & DocNum & " - " & _BPName       ' HANA
                    'End If

                    If _EmailSubject.Trim.Length > 0 Then
                        _EmailSubject = _EmailSubject.Trim.Replace("<<COMPANYNAME>>", oCompany.CompanyName)
                        _EmailSubject = _EmailSubject.Trim.Replace("<<DOCUMENTNUM>>", DocNum)
                        _EmailSubject = _EmailSubject.Trim.Replace("<<BPNAME>>", _BPName)
                        _EmailSubject = _EmailSubject.Trim.Replace("<<BPCODE>>", _BPCode)
                        _EmailSubject = _EmailSubject.Trim.Replace("<<TITLE>>", "Payment Voucher")

                        a.Subject = _EmailSubject
                    Else
                        a.Subject = "Payment Voucher No. " & DocNum & " - " & _BPName
                    End If


                    bIsHTML = False
                    Select Case _EmailType
                        Case "H"
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
                                    sOutput2 = sOutput.Replace("{0}", DocNum)
                                Else
                                    sOutput2 = sOutput
                                End If
                            Else
                                sOutput2 = "Please refer to attachment."
                            End If
                        Case Else
                            sOutput2 = _PlainText
                    End Select


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
                                If bImage1 AndAlso IO.File.Exists("ImagePV001.gif") Then
                                    imageRes1 = New System.Net.Mail.LinkedResource("ImagePV001.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                                    imageRes1.ContentId = "image1"
                                    imageRes1.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                                    bodyAltView.LinkedResources.Add(imageRes1)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage2 AndAlso IO.File.Exists("ImagePV002.gif") Then
                                    imageRes2 = New System.Net.Mail.LinkedResource("ImagePV002.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                                    imageRes2.ContentId = "image2"
                                    imageRes2.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                                    bodyAltView.LinkedResources.Add(imageRes2)
                                End If
                            Catch ex As Exception

                            End Try
                            Try
                                If bImage3 AndAlso IO.File.Exists("ImagePV003.gif") Then
                                    imageRes3 = New System.Net.Mail.LinkedResource("ImagePV003.gif", "image/gif") ' !*** CHANGE AS NEEDED (image/jpeg, image/gif, etc)
                                    imageRes3.ContentId = "image3"
                                    imageRes3.TransferEncoding = Net.Mime.TransferEncoding.Base64 ' Set the encoding
                                    bodyAltView.LinkedResources.Add(imageRes3)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage4 AndAlso IO.File.Exists("ImagePV004.gif") Then
                                    imageRes4 = New System.Net.Mail.LinkedResource("ImagePV004.gif", "image/gif")
                                    imageRes4.ContentId = "image4"
                                    imageRes4.TransferEncoding = Net.Mime.TransferEncoding.Base64
                                    bodyAltView.LinkedResources.Add(imageRes4)
                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage5 AndAlso IO.File.Exists("ImagePV005.gif") Then
                                    imageRes5 = New System.Net.Mail.LinkedResource("ImagePV005.gif", "image/gif")
                                    imageRes5.ContentId = "image5"
                                    imageRes5.TransferEncoding = Net.Mime.TransferEncoding.Base64
                                    bodyAltView.LinkedResources.Add(imageRes5)

                                End If
                            Catch ex As Exception

                            End Try

                            Try
                                If bImage6 AndAlso IO.File.Exists("ImagePV006.gif") Then
                                    imageRes6 = New System.Net.Mail.LinkedResource("ImagePV006.gif", "image/gif")
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

                    a.Body = sOutput2

                    If _Attachment.Trim.Length > 0 Then
                        Dim b As New System.Net.Mail.Attachment(_Attachment)
                        a.Attachments.Add(b)
                    End If

                    Dim c As New System.Net.Mail.SmtpClient(_SMTP_Server)

                    If _AuthType = "1" Then
                        Dim d As New System.Net.NetworkCredential(_Username, _Password)

                        ' START - added by ES 22.01.2016 for future enhancement.
                        If _LocalIPAddress.Trim <> "" Then
                            c.Host = _LocalIPAddress
                        End If

                        If _IsOffice365 = "Y" Then
                            System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                            c.EnableSsl = True
                        Else
                            If _EnableSSL = "Y" Then
                                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
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


                    c = Nothing
                    a = Nothing

            End Select

            Return True
        Catch ex As Exception
            ErrorMessage = ex.Message
            Return False
        End Try
        Return False
    End Function

End Class
