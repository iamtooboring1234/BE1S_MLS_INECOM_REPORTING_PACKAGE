'' © Copyright © 2007-2019, Inecom Pte Ltd, All rights reserved.
'' =============================================================

Option Strict Off
Option Explicit On 

Imports System.IO
Imports SAPbobsCOM
Imports System.Data.Odbc
Imports System.Data.Common

Public Enum MenuID
    Next_Record = 1288
    Previous_Record = 1289
    First_Record = 1290
    Last_Record = 1291
    Remove_Record = 1283
    Duplicate_Record = 1287
    Delete_Row = 1293
    Add_Row = 1292
    Undo = 769L
    Cut = 771
    Remove = 1283
    Copy = 772
    Paste = 773
    Delete = 774
    Add = 1282
    Find = 1281
    Print = 520
    PrintPreview = 519
    ProductionOrder = 4369
    F5 = 6405
End Enum
Public Enum AgeingType
    ARAgeing = 0
    APAgeing = 1
    ARAgeingSummary = 2
    APAgeingSummary = 3
End Enum
Public Enum ReportCode
    SOA_ByDocDate = 0
    SOA_ByDueDate = 1
End Enum
Public Enum ReportName
    ARAging_Details = 1
    APAging_Details = 2
    ARSoa = 3
    APSoa = 4
    PV = 5
    RA = 6
    GST = 7
    OMARSoa = 8
    SAR_FIFO_SUMMARY = 9
    SAR_FIFO_DETAILS = 10
    SAR_MOVAVG_SUMMARY = 11
    SAR_MOVAVG_DETAILS = 12
    SAR_Audit_Enquiry = 13
    PV_Range = 14
    IRA = 15
    SOA_Email_Config = 17
    ARPayment = 19
    APPayment = 20
    PVDraft = 21    'JN added
    RecpPaym = 22   'JN added V03.28.2007
    SOA_TOS = 23                        'ES added 31.08.2011
    GPA = 24                            'ES added 31.08.2011
    Items_ABC_Analysis = 25             'ES added on 10.11.2011
    ReOrder_Level_Recommendation = 26   'ES added on 10.11.2011
    Weighted_Average_Demand = 27        'ES added on 10.11.2011
    GL_Listing = 28                     'ES
    MRPSupplyDemandReport = 29
    ARAging_Summary = 30
    APAging_Summary = 31
    ARAging6B_Details = 34
    ARAging6B_Summary = 35
    SAR_TM_V1 = 36
    SAR_TM_V2 = 37
    SAR_TM_V3 = 38
    ARAging_Details_Proj = 39
    APAging_Details_Proj = 40
    ARAging_Summary_Proj = 41
    APAging_Summary_Proj = 42
    SO_Detail_Proj = 43
    PO_Detail_Proj = 44
    ARAging7B_Details = 45
    ARAging7B_Summary = 46
    OfficialReceipt = 47
    ARSOA_Landscape = 48
    ARAgeingDetailsCRM = 49
    BankReconciliation = 50
    ARSOA_BY_PROJECT = 51
    PV_Email_Config = 52    'AT added on 04.05.2019
    PV_Mass_Email = 53      'AT added on 04.05.2019
    CHANGE_LOG_AUDIT = 55         'ES added on 03.10.2019
End Enum
Public Enum CompanyCode
    General = 0
    AE = 1
    AMS = 2
    FL = 3
    TME = 4
End Enum
Friend Enum keyID
    Up = 38
    Down = 40
    Left = 37
    Right = 39
    Tab = 9
    Enter = 13
    Delete = 46
End Enum

Module SubMain

#Region "Variables"
    Private oNCM_PO_PROJ As CLS_NCM_PO_PROJ
    Private oNCM_SO_PROJ As CLS_NCM_SO_PROJ
    Private oNCM_GLL As NCM_GLL_V
    Private oNCM_RPT_CONFIG As NCM_RPT_CONFIG_V
    Private oNCM_AR_AGEING_CRM As frmARAgeingCRM
    Private oNCM_BREC As NCM_BREC_V

    Private oARAgeing_Proj As frmARAgeing_Proj
    Private oAPAgeing_Proj As frmAPAgeing_Proj
    Private oARAgeing As frmARAgeing
    Private oARAgeing6B As frmARAgeing6B
    Private oARAgeing7B As frmARAgeing7B
    Private oAPAgeing As frmAPAgeing

    Private oApSOA As frmApSOA
    Private oArSOA As frmArSOA
    Private oARSOAPROJ As frmArSOAProj
    Private oPaymentVoucher As OutgoingPayment
    Private oIncomingPayment As IncomingPayment
    Private oOMArSOA As frmOMArSOA
    Private oFIFO_NonBatch As StockAging_FIFO_NonBatch
    Private oMOV_StockAging As StockAging_MOV
    Private oNCM_gst As NCM_GST
    Private oStockAudit As NCM_StockAuditTrail
    Private oFRM_PaymentVoucher_Range As FRM_PaymentVoucher_Range
    Private oTM_SAR As frmTMSAR
    Private oFrmEmail As frmEmail
    Private oAPPayment As frmAPPayment
    Private oARPayment As frmARPayment
    Private oPaymentDraft As PaymentDraft   'JN added PaymentDraft is selected from selection list
    Private ofrmRecpPaym As frmRecpPaym   'JN added V03.28.2007
    Private oSOA_TOS As frmSOA_TOS  'ES added on 31.08.2011
    Private oGPA As FRM_GPA         'ES added on 31.08.2011
    Private oFrmPVEmail As frmPVEmail   'AT Added on 04.05.2019
    Private oPVEmailParam As NCM_PV_Email_Param

    Friend oStockAudit_Out As NCM_StockAuditTrail_Out
    Friend oFrmSendEmail As frmSendEmail
    Friend oFrmSendEmailProj As frmSendEmailProj
    Friend oNCM_IAR As CLS_NCM_IAR                  'ES added on 10.11.2011
    Friend oNCM_RLR As CLS_NCM_RLR                  'ES added on 10.11.2011
    Friend oNCM_WAR As CLS_NCM_WAR                  'ES added on 10.11.2011
    Friend oNCM_CHG_LOG_AUDIT As NCM_CHG_LOG_AUDIT  'ES added on 03.10.2019
    Friend oFrmPVSendEmail As frmPVSendEmail

    Public SQLDbConnection As System.Data.SqlClient.SqlConnection
    Public HANADbConnection As DbConnection
    Public _DbProviderFactoryObject As DbProviderFactory
    Public ProviderName As String = "System.Data.Odbc"

    Public WithEvents SBO_Application As SAPbouiCOM.Application
    Public oCompany As SAPbobsCOM.Company
    Public TmpThread As New Threading.Thread(AddressOf CloseApp)
    Friend global_DBUsername As String = ""
    Friend global_DBPassword As String = ""
    Friend DBUsername As String = ""
    Friend DBPassword As String = ""
    Friend DBIntegratedSecurity As Boolean = False
    Friend DBConnString As String = ""
    Friend sQuery As String = ""
    Friend connStr As String = ""
    Friend Const PROJ_NAMESPACE As String = "AgeingReport"
    ' LICENSE
    Private Const secretLock As String = "$ql$@@cc"         ' sql sa acc
    Private Const secretKey As String = "%inec0m$ecret!"    ' % inecom secret !

    Private Const dbNCMSBO As String = "NCMSBO"
    Private Const sysProfile As String = "sysProfile"
    Private Const licProfile As String = "licProfile"

    Private sqlServer, adminAccount, adminPassword, dbConnectionString As String
    Private volSerial, macAddress, cpuID, signature As String
    Private integratedSecurity As Boolean

    ' Private oLicLib As New LicenseLib.LicLibrary'kpkp
    Friend Const MNU_ARSOA_PROJ As String = "MNU_ARSOA_PROJ"
    Friend Const FRM_ARSOA_PROJ As String = "NCM_ARSOA_PROJ"
    Friend Const MNU_ARAGEING_7B As String = "MNU_ARAGEING_7B"
    Friend Const FRM_ARAGEING_7B As String = "NCM_ARAGEING_7B"
    Friend Const MNU_NCM_PO_PROJ As String = "MNU_NCM_PO_PROJ"
    Friend Const MNU_NCM_SO_PROJ As String = "MNU_NCM_SO_PROJ"
    Friend Const FRM_NCM_PO_PROJ As String = "NCM_PO_PROJ"
    Friend Const FRM_NCM_SO_PROJ As String = "NCM_SO_PROJ"
    Friend Const FRM_NCM_BREC As String = "NCM_BREC"

    Friend Const MNU_APAGEING_PROJ As String = "MNU_APAGEING_PROJ"
    Friend Const MNU_ARAGEING_PROJ As String = "MNU_ARAGEING_PROJ"
    Friend Const FRM_APAGEING_PROJ As String = "NCM_APAGEING_PROJ"
    Friend Const FRM_ARAGEING_PROJ As String = "NCM_ARAGEING_PROJ"

    Friend Const ACTIONS_ITEM_UID As String = "btnInecom"
    Friend Const MNU_CHG_LOG_AUDIT As String = "MNU_CHG_LOG_AUDIT"
    Friend Const FRM_CHG_LOG_AUDIT As String = "NCM_CHG_LOG_AUDIT"
    Friend Const FILE_CHG_LOG_AUDIT As String = "FIL_CHG_LOG_AUDIT.srf"

    Friend Const MNU_NCM_GLL As String = "MNU_NCM_GLL"
    Friend Const FRM_NCM_GLL As String = "NCM_GLL"

    Friend Const MNU_NCM_BREC As String = "MNU_NCM_BREC"
    Friend Const MNU_HYD As String = "MNU_HYD"
    Friend Const MNU_SUB_IAR As String = "MNU_SUB_IAR"
    Friend Const MNU_SUB_WAR As String = "MNU_SUB_WAR"
    Friend Const MNU_SUB_RLR As String = "MNU_SUB_RLR"
    Friend Const FRM_NCM_IAR As String = "NCM_IAR"
    Friend Const FRM_NCM_WAR As String = "NCM_WAR"
    Friend Const FRM_NCM_RLR As String = "NCM_RLR"

    Friend Const MNU_NCM_FIFO1 As String = "MNU_NCM_FIFO1"
    Friend Const FRM_NCM_FIFO1 As String = "NCM_FIFO1"
    Friend Const MNU_NCM_MOV1 As String = "MNU_NCM_MOV1"
    Friend Const FRM_NCM_MOV1 As String = "NCM_MOV1"
    Friend Const MNU_NCM_SES1 As String = "MNU_NCM_SES1"
    Friend Const FRM_NCM_SES1 As String = "NCM_SES1"
    Friend Const MNU_NCM_SES2 As String = "MNU_NCM_SES2"
    Friend Const FRM_NCM_SES2 As String = "NCM_SES2"
    Friend Const FRM_NCM_PV_RANGE As String = "NCMPAYCHER"
    Friend Const MNU_NCM_PV_RANGE As String = "MNU_PV_RANGE"
    Friend Const MNU_NCM_INECOM_SDK As String = "MNU_NCM_RPT_SDK"
    Friend Const MNU_NCM_TM_STOCK As String = "MNU_NCM_TM_SAR"
    Friend Const MNU_NCM_EMAIL As String = "MNU_NCM_EMAIL"

    Friend Const MNU_NCM_AR_PAYMENT As String = "MNU_NCM_AR_PAYMENT"
    Friend Const FRM_NCM_AR_PAYMENT As String = "NCM_AR_PAYMENT"
    Friend Const MNU_NCM_AP_PAYMENT As String = "MNU_NCM_AP_PAYMENT"
    Friend Const FRM_NCM_AP_PAYMENT As String = "NCM_AP_PAYMENT"
    Friend Const MNU_RPT_CONFIG As String = "MNU_RPT_CONFIG"
    Friend Const FRM_RPT_CONFIG As String = "NCM_RPT_CONFIG"

    Friend Const MNU_ARAGEING_JCS As String = ""
    Friend Const MNU_APAGEING_JCS As String = ""
    Friend Const MNU_ARAGEING_CRM As String = "MNU_AR_AGEING_CRM"

    Friend Const MNU_NCM_PV_EMAIL As String = "MNU_NCM_PV_EMAIL"
    Friend Const MNU_NCM_PV_EMAIL_PARA As String = "MNU_NCM_PV_EMAIL_PARA"

#End Region

#Region "Initialize Application"
    Public Sub Main()
        Try
            ' =========================================================
            SetApplication()
            oCompany = New SAPbobsCOM.Company
            oCompany = SBO_Application.Company.GetDICompany()

#If DEBUG Then
            global_DBUsername = "SYSTEM"
            global_DBPassword = "Hana#sg1"
            connStr = "DRIVER={HDBODBC32};UID=" & global_DBUsername & ";PWD=" & global_DBPassword & ";SERVERNODE=" & oCompany.Server & ";DATABASE=" & oCompany.CompanyDB & ""

            Dim ProviderName As String = "System.Data.Odbc"
            Dim dbConn As DbConnection = Nothing
            Dim _DbProviderFactoryObject As DbProviderFactory
            Dim sQuery As String = ""

            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
            dbConn = _DbProviderFactoryObject.CreateConnection()
            dbConn.ConnectionString = connStr
            dbConn.Open()
#Else
            Dim bCloud As Boolean = False
            If bCloud = False Then
                If CheckLicense() = False Then
                    SBO_Application.MessageBox("Failed to find license for this add-on.")
                End If 'terminating add on
            End If
#End If

            If Not TablesInitialization() Then
                SBO_Application.MessageBox("Failed to initialize tables needed.")
                End
            End If

            If RemoveMenuItems() Then   'if removing success, then add menu items
                AddMenuItems()
            End If

            oNCM_RPT_CONFIG = New NCM_RPT_CONFIG_V
            oNCM_PO_PROJ = New CLS_NCM_PO_PROJ
            oNCM_SO_PROJ = New CLS_NCM_SO_PROJ
            oARAgeing_Proj = New frmARAgeing_Proj
            oAPAgeing_Proj = New frmAPAgeing_Proj
            oNCM_AR_AGEING_CRM = New frmARAgeingCRM
            oNCM_BREC = New NCM_BREC_V
            oNCM_CHG_LOG_AUDIT = New NCM_CHG_LOG_AUDIT

            oSOA_TOS = New frmSOA_TOS
            oGPA = New FRM_GPA
            oNCM_IAR = New CLS_NCM_IAR
            oNCM_RLR = New CLS_NCM_RLR
            oNCM_WAR = New CLS_NCM_WAR
            oARAgeing = New frmARAgeing
            oAPAgeing = New frmAPAgeing
            oApSOA = New frmApSOA
            oArSOA = New frmArSOA
            oARSOAPROJ = New frmArSOAProj
            oOMArSOA = New frmOMArSOA
            oPaymentVoucher = New OutgoingPayment
            oFIFO_NonBatch = New StockAging_FIFO_NonBatch
            oNCM_gst = New NCM_GST
            oStockAudit = New NCM_StockAuditTrail
            oMOV_StockAging = New StockAging_MOV
            oStockAudit_Out = New NCM_StockAuditTrail_Out
            oFRM_PaymentVoucher_Range = New FRM_PaymentVoucher_Range
            oIncomingPayment = New IncomingPayment
            oPaymentDraft = New PaymentDraft    'JN added
            oARAgeing6B = New frmARAgeing6B
            oARAgeing7B = New frmARAgeing7B
            oFrmEmail = New frmEmail
            oFrmSendEmail = New frmSendEmail
            oFrmSendEmailProj = New frmSendEmailProj
            oAPPayment = New frmAPPayment
            oARPayment = New frmARPayment
            ofrmRecpPaym = New frmRecpPaym 'JN added V03.28.2007
            oNCM_GLL = New NCM_GLL_V ' ES added 04.12.2007
            oTM_SAR = New frmTMSAR
            oFrmPVEmail = New frmPVEmail 'AT Added on 04.05.2019
            oFrmPVSendEmail = New frmPVSendEmail
            oPVEmailParam = New NCM_PV_Email_Param

            'MsgBox(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\APSOA\" & oCompany.CompanyDB)

            GetEmbeddedBMP("Inecom_SDK_Reporting_Package.ncmInecom.bmp").Save("ncmInecom.bmp")
            SBO_Application.StatusBar.SetText("Inecom SDK Reporting Package add-on connected.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            System.Windows.Forms.Application.Run()

        Catch ex As Exception
            MsgBox("[Main] : " & vbNewLine & ex.Message)
            System.Environment.Exit(0)
        End Try
    End Sub
    Private Function SetApplication() As Boolean
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String
        Try
            SboGuiApi = New SAPbouiCOM.SboGuiApi
            sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
            SboGuiApi.Connect(sConnectionString)            'connect to a running SBO application
            SBO_Application = SboGuiApi.GetApplication()    'reference to the running SBO application
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function SetConnectionContext() As Long
        Dim sCookie As String
        Dim sConnectionContext As String
        Try
            oCompany = New SAPbobsCOM.Company
            sCookie = oCompany.GetContextCookie                     'acquire the connection context cookie from the DI API
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie)  'retrieve the connection context string from the UI API
            If oCompany.Connected = True Then oCompany.Disconnect() 'ensure the company is not connected
            Return oCompany.SetSboLoginContext(sConnectionContext)  'set the connection context information to the DI API
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function AddMenuItems() As Boolean
        Dim oMenus As SAPbouiCOM.Menus
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oNewMenuItem As SAPbouiCOM.MenuCreationParams
        Dim frmCmdCenter As SAPbouiCOM.Form

        Try
            frmCmdCenter = SBO_Application.Forms.GetFormByTypeAndCount(169, 1)

            ' Reference to Main Menu
            oMenuItem = SBO_Application.Menus.Item("43520")
            oMenus = oMenuItem.SubMenus
            oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_POPUP
            oNewMenuItem.UniqueID = MNU_NCM_INECOM_SDK
            oNewMenuItem.String = "Inecom Reporting Package"
            oNewMenuItem.Image = Application.StartupPath.ToString & "\" & "ncmMenu.bmp"
            oNewMenuItem.Position = oMenus.Count + 1     'append at the end
            oMenus.AddEx(oNewMenuItem)

            ' Reference to SDK System Menu
            oMenuItem = SBO_Application.Menus.Item(MNU_NCM_INECOM_SDK)
            oMenus = oMenuItem.SubMenus

            oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
            oNewMenuItem.UniqueID = MNU_RPT_CONFIG
            oNewMenuItem.String = "Reporting Package Configuration"
            oMenus.AddEx(oNewMenuItem)

            If (IsIncludeModule(ReportName.SOA_Email_Config)) Then
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_EMAIL
                oNewMenuItem.String = "Email Configuration"
                oMenus.AddEx(oNewMenuItem)
            End If

            If (IsIncludeModule(ReportName.PV_Mass_Email)) Then 'AT added on 04.05.2019
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_PV_EMAIL
                oNewMenuItem.String = "PV Email Configuration"
                oMenus.AddEx(oNewMenuItem)
            End If

            If (IsIncludeModule(ReportName.SAR_TM_V1)) Then
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_TM_STOCK
                oNewMenuItem.String = "Stock Valuation Report (TM)"
                oMenus.AddEx(oNewMenuItem)
            End If
            If (IsIncludeModule(ReportName.GPA)) Then   'GP Analysis
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCM_GPA"
                oNewMenuItem.String = "Gross Profit Analysis"
                oMenus.AddEx(oNewMenuItem)
            End If

            If (IsIncludeModule(ReportName.Items_ABC_Analysis)) Then
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_SUB_IAR
                oNewMenuItem.String = "Items ABC Analysis"
                oMenus.AddEx(oNewMenuItem)
            End If

            If (IsIncludeModule(ReportName.ReOrder_Level_Recommendation)) Then
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_SUB_RLR
                oNewMenuItem.String = "ReOrder Level Recommendation"
                oMenus.AddEx(oNewMenuItem)
            End If

            If (IsIncludeModule(ReportName.Weighted_Average_Demand)) Then
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_SUB_WAR
                oNewMenuItem.String = "Weighted Average Demand"
                oMenus.AddEx(oNewMenuItem)
            End If

            If (IsIncludeModule(ReportName.CHANGE_LOG_AUDIT)) Then
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_CHG_LOG_AUDIT
                oNewMenuItem.String = "Change Log Audit Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            If (IsIncludeModule(ReportName.PV_Mass_Email)) Then 'AT added on 04.05.2019
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_PV_EMAIL_PARA
                oNewMenuItem.String = "PV Email"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("12800") ' Sales Report 
            If (IsIncludeModule(ReportName.ARAging_Details_Proj)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCM_ARAgeing_Proj"
                oNewMenuItem.String = "AR Ageing Report with Project"
                oMenus.AddEx(oNewMenuItem)
            End If

            If (IsIncludeModule(ReportName.ARAging_Details)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCM_ARAgeing"
                oNewMenuItem.String = "AR Ageing Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            If (IsIncludeModule(ReportName.ARAgeingDetailsCRM)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_ARAGEING_CRM
                oNewMenuItem.String = "AR Ageing Details Report with CRM Notes"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("12800") ' Sales Report 
            If (IsIncludeModule(ReportName.ARAging6B_Details)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCM_ARAgeing6B"
                oNewMenuItem.String = "AR Ageing (6 Buckets) Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("12800") ' Sales Report 
            If (IsIncludeModule(ReportName.ARAging7B_Details)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_ARAGEING_7B
                oNewMenuItem.String = "AR Ageing (7 Buckets) Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("12800") ' Sales Report 
            If (IsIncludeModule(ReportName.SO_Detail_Proj)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_SO_PROJ
                oNewMenuItem.String = "SO Details By Customer"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("12800") ' Sales Report 
            If (IsIncludeModule(ReportName.ARSoa)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCM_SOA"
                oNewMenuItem.String = "AR Statement Of Account"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("12800") ' Sales Report 
            If (IsIncludeModule(ReportName.ARSOA_BY_PROJECT)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_ARSOA_PROJ
                oNewMenuItem.String = "AR Statement Of Account (Group By Project)"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("12800") ' Sales Report 
            If (IsIncludeModule(ReportName.OMARSoa)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCM_OMARSOA"
                oNewMenuItem.String = "OM - AR Statement Of Account"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("12800") ' Sales Report 
            If (IsIncludeModule(ReportName.SOA_TOS)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCM_SOATOS"
                oNewMenuItem.String = "A/R SOA By Project"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("43534") 'Purchasing Report 
            If (IsIncludeModule(ReportName.APAging_Details_Proj)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCM_APAgeing_Proj"
                oNewMenuItem.String = "AP Ageing Report with Project"
                oMenus.AddEx(oNewMenuItem)
            End If

            If (IsIncludeModule(ReportName.APAging_Details)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCM_APAgeing"
                oNewMenuItem.String = "AP Ageing Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("43534") 'Purchasing Report 
            If (IsIncludeModule(ReportName.PO_Detail_Proj)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_PO_PROJ
                oNewMenuItem.String = "PO Details By Vendor"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("43534") 'Purchasing Report 
            If (IsIncludeModule(ReportName.APSoa)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCM_SOA_AP"
                oNewMenuItem.String = "AP Statement Of Account"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("1760") 'Inventory Report 
            If (IsIncludeModule(ReportName.SAR_FIFO_DETAILS)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_FIFO1
                oNewMenuItem.String = "Stock Ageing for FIFO Items"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("1760") 'Inventory Report 
            If (IsIncludeModule(ReportName.SAR_MOVAVG_DETAILS)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_MOV1
                oNewMenuItem.String = "Stock Ageing Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("1760") 'Inventory Report 
            If (IsIncludeModule(ReportName.SAR_Audit_Enquiry)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_SES1
                oNewMenuItem.String = "Stock Audit Trail Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("43531") 'Financial Report
            If IsIncludeModule(ReportName.GL_Listing) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_GLL
                oNewMenuItem.String = "GL Listing Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("43532") 'Financial Report - Tax
            If (IsIncludeModule(ReportName.GST)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "NCMmnuGST"
                oNewMenuItem.String = "GST Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("43538") 'Banking - Outgoing Payment
            If (IsIncludeModule(ReportName.PV)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_PV_RANGE
                oNewMenuItem.String = "Payment Voucher"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("51197") 'Banking Reports
            If (IsIncludeModule(ReportName.ARPayment)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_AR_PAYMENT
                oNewMenuItem.String = "Incoming Payment Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("51197") 'Banking Reports
            If (IsIncludeModule(ReportName.APPayment)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_AP_PAYMENT
                oNewMenuItem.String = "Outgoing Payment Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("51197") 'Banking Reports
            If (IsIncludeModule(ReportName.RecpPaym)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = "MNU_NCM_RECPPAYM"
                oNewMenuItem.String = "Receipt and Payment List"
                oMenus.AddEx(oNewMenuItem)
            End If

            oMenuItem = SBO_Application.Menus.Item("51197") 'Banking Reports
            If (IsIncludeModule(ReportName.BankReconciliation)) Then
                oMenus = oMenuItem.SubMenus
                oNewMenuItem = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                oNewMenuItem.Type = SAPbouiCOM.BoMenuType.mt_STRING
                oNewMenuItem.UniqueID = MNU_NCM_BREC
                oNewMenuItem.String = "Bank Reconciliation Report"
                oMenus.AddEx(oNewMenuItem)
            End If

            

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Function RemoveMenuItems() As Boolean
        Dim oMenus As SAPbouiCOM.Menus
        Try
            oMenus = SBO_Application.Menus

            If oMenus.Exists("NCM_APAgeing_Proj") Then
                oMenus.RemoveEx("NCM_APAgeing_Proj")
            End If

            If oMenus.Exists("NCM_ARAgeing_Proj") Then
                oMenus.RemoveEx("NCM_ARAgeing_Proj")
            End If

            If oMenus.Exists(MNU_NCM_SO_PROJ) Then
                oMenus.RemoveEx(MNU_NCM_SO_PROJ)
            End If

            If oMenus.Exists(MNU_NCM_PO_PROJ) Then
                oMenus.RemoveEx(MNU_NCM_PO_PROJ)
            End If

            If oMenus.Exists(MNU_NCM_GLL) Then
                oMenus.RemoveEx(MNU_NCM_GLL)
            End If

            If oMenus.Exists(MNU_NCM_BREC) Then
                oMenus.RemoveEx(MNU_NCM_BREC)
            End If

            If oMenus.Exists(MNU_ARAGEING_CRM) Then
                oMenus.RemoveEx(MNU_ARAGEING_CRM)
            End If

            If oMenus.Exists("NCM_ARAgeing") Then
                oMenus.RemoveEx("NCM_ARAgeing")
            End If

            If oMenus.Exists("NCM_ARAgeing6B") Then
                oMenus.RemoveEx("NCM_ARAgeing6B")
            End If

            If oMenus.Exists("NCM_APAgeing") Then
                oMenus.RemoveEx("NCM_APAgeing")
            End If

            If oMenus.Exists("NCM_SOA_AP") Then
                oMenus.RemoveEx("NCM_SOA_AP")
            End If

            If oMenus.Exists("NCM_SOA") Then
                oMenus.RemoveEx("NCM_SOA")
            End If

            If oMenus.Exists("NCM_OMARSOA") Then
                oMenus.RemoveEx("NCM_OMARSOA")
            End If

            If oMenus.Exists(MNU_NCM_FIFO1) Then
                oMenus.RemoveEx(MNU_NCM_FIFO1)
            End If

            If oMenus.Exists(MNU_NCM_MOV1) Then
                oMenus.RemoveEx(MNU_NCM_MOV1)
            End If

            If oMenus.Exists("NCMmnuGST") Then
                oMenus.RemoveEx("NCMmnuGST")
            End If

            If oMenus.Exists(MNU_NCM_SES1) Then
                oMenus.RemoveEx(MNU_NCM_SES1)
            End If

            If oMenus.Exists(MNU_NCM_PV_RANGE) Then
                oMenus.RemoveEx(MNU_NCM_PV_RANGE)
            End If

            If oMenus.Exists(MNU_NCM_AP_PAYMENT) Then
                oMenus.RemoveEx(MNU_NCM_AP_PAYMENT)
            End If

            If oMenus.Exists(MNU_NCM_AR_PAYMENT) Then
                oMenus.RemoveEx(MNU_NCM_AR_PAYMENT)
            End If

            If oMenus.Exists("MNU_NCM_RECPPAYM") Then
                oMenus.RemoveEx("MNU_NCM_RECPPAYM")
            End If

            If oMenus.Exists(MNU_NCM_INECOM_SDK) Then
                oMenus.RemoveEx(MNU_NCM_INECOM_SDK)
            End If

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Sub CloseApp()
        TmpThread.Sleep(20)
        End
    End Sub
    Private Function CheckLicense() As Boolean
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sql As String = ""

            ' HANA
            sql = "  SELECT COUNT(*) FROM SYS.OBJECTS  WHERE ""SCHEMA_NAME""='NCMSBO'"
            oRecordSet.DoQuery(sql)
            If (oRecordSet.RecordCount > 0) Then
                If oRecordSet.Fields.Item(0).Value >= 1 Then
                    sql = "SELECT ""server"", ""passport"", ""lickey"" FROM NCMSBO.""sysProfile"""

                    oRecordSet.DoQuery(sql)
                    If (oRecordSet.RecordCount > 0) Then
                        oRecordSet.MoveFirst()
                        Dim server, passport, lickey As String
                        server = oRecordSet.Fields.Item(0).Value
                        passport = oRecordSet.Fields.Item(1).Value
                        lickey = oRecordSet.Fields.Item(2).Value

                        Dim licAddOn As New LicenseClient.LicClient
                        Dim flLicensed As Boolean

                        flLicensed = licAddOn.IsLicensed(server, passport, lickey, PROJ_NAMESPACE)

                        If flLicensed = True Then
                            global_DBUsername = licAddOn.sqlAdminAccount
                            global_DBPassword = licAddOn.sqlAdminPassword
                            DBIntegratedSecurity = licAddOn.sqlIntegratedSecurity

                            '32 bit connstring
                            connStr = "DRIVER={HDBODBC32};UID=" & global_DBUsername & ";PWD=" & global_DBPassword & ";SERVERNODE=" & oCompany.Server & ";DATABASE=" & oCompany.CompanyDB & ""

                            '64 bit connstring
                            'connStr = "DRIVER={HDBODBC};UID=" & global_DBUsername & ";PWD=" & global_DBPassword & ";SERVERNODE=" & oCompany.Server & ";DATABASE=" & oCompany.CompanyDB & ""

                            _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
                            HANADbConnection = _DbProviderFactoryObject.CreateConnection()
                            HANADbConnection.ConnectionString = connStr
                            HANADbConnection.Open()
                            Return True
                        Else
                            SBO_Application.MessageBox("Failed to find license for this add-on.")
                            Return False
                        End If
                    Else
                        SBO_Application.MessageBox("License database for addon is corrputed.")
                        Return False
                    End If
                Else
                    SBO_Application.MessageBox("Please proceed to setup Inecom License first..")
                    Return False
                End If
            Else
                SBO_Application.MessageBox("Please proceed to setup Inecom License first.")
                Return False
            End If
        Catch ex As Exception
            SBO_Application.MessageBox("[Ncm_Main].[SubMain]" & vbNewLine & ex.Message)
            Return False
        End Try
    End Function
    Private Sub WriteEmail()
        Try
            Dim _SMTP_Server As String = String.Empty
            Dim _EmailFrom As String = String.Empty
            Dim _EmailTo As String = String.Empty
            Dim _AuthType As String = String.Empty
            Dim _Username As String = String.Empty
            Dim _Password As String = String.Empty
            Dim _Attachment As String = String.Empty
            Dim _PortNum As Integer = 0
            Dim _LocalIPAddress As String = ""

            Dim oRecord As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim sQuery As String = "SELECT IFNULL(""U_MailFrom"",''), IFNULL(""U_SMTP"",''), IFNULL(""U_Username"",''), IFNULL(""U_Password"",''), IFNULL(""U_AuthType"",0), IFNULL(""U_PortNum"",0) ""PortNumber"", IFNULL(""U_LocalIP"",'') ""LocalIPAddress"" FROM ""@NCM_EMAIL_CONFIG"" WHERE ""Code"" = 'SOA'"
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                _EmailFrom = oRecord.Fields.Item(0).Value
                _SMTP_Server = oRecord.Fields.Item(1).Value
                _Username = oRecord.Fields.Item(2).Value
                _Password = oRecord.Fields.Item(3).Value
                _AuthType = oRecord.Fields.Item(4).Value
                _PortNum = oRecord.Fields.Item("PortNumber").Value
                _LocalIPAddress = oRecord.Fields.Item("LocalIPAddress").Value.ToString.Trim

            Else
                SBO_Application.MessageBox("[clsEmail].[GetSetting] - Please configure email setting.", 1, "OK", String.Empty, String.Empty)
            End If

            '_EmailFrom = "erwine.lee@gmail.com"
            '_SMTP_Server = "smtp.gmail.com"
            '_Username = "erwine.lee@gmail.com"
            '_Password = "quan12gm"
            '_AuthType = "1"

            _EmailFrom = "sukardye@inecomworld.com"
            _SMTP_Server = "124.6.61.66"
            _Username = "sukardye@inecomworld.com"
            _Password = "Inec0msgp"
            _AuthType = "1"


            Dim tmpMailFr As New System.Net.Mail.MailAddress(_EmailFrom)
            Dim s() As String = "erwine.lee@gmail.com;erwine.sukardy@inecomworld.com".Split(";")
            'Dim s() As String = "erwine.sukardy@inecomworld.com".Split(";")

            Dim a As New System.Net.Mail.MailMessage()
            Dim sOutput As String = String.Empty
            Dim sOutput2 As String = String.Empty
            Dim bIsHTML As Boolean = False

            a.From = tmpMailFr
            For i As Integer = 0 To s.Length - 1
                a.To.Add(s(i))
            Next
            a.Subject = "TEST2 Statement Of Account"
            a.Body = "Please refer to attachment"

            'Dim b As New System.Net.Mail.Attachment(_Attachment)
            'a.Attachments.Add(b)

            Dim c As New System.Net.Mail.SmtpClient(_SMTP_Server)
            Dim strHostName As String = ""
            Dim strIPAddress As String = ""
            strHostName = System.Net.Dns.GetHostName()
            strIPAddress = System.Net.Dns.GetHostEntry(strHostName).AddressList(1).ToString()

            If _AuthType = "1" Then
                Dim d As New System.Net.NetworkCredential(_Username, _Password)
                '
                'c.Host = _SMTP_Server
                'c.Port = "587"
                ' c.EnableSsl = True 'for gmail
                c.DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network
                c.UseDefaultCredentials = False
                c.Credentials = d

                '_LocalIPAddress = "192.168.1.82"

                'If _LocalIPAddress <> "" Then
                '    c.Host = _LocalIPAddress
                'End If
                'If _PortNum > 0 Then
                '    c.Port = _PortNum
                'End If
            End If
            c.Send(a)
            ' b.Dispose()

            c = Nothing
            a = Nothing
            'b = Nothing
            'Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            'ErrorMessage = ex.Message
            'Return False
        End Try
        'Return False
    End Sub
#End Region

#Region "General Function"
    Public Function LoadFromXML(ByRef FileName As String) As Boolean
        Try
            SBO_Application.LoadBatchActions(GetEmbeddedXML(FileName).InnerXml)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function GetEmbeddedXML(ByVal strIdentifier As String) As System.Xml.XmlDocument
        Dim xmlDoc As New System.Xml.XmlDocument
        With New System.IO.StreamReader(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(strIdentifier))
            xmlDoc.Load(.BaseStream)
            .Close()
            Return xmlDoc
        End With
    End Function
    Public Function GetEmbeddedBMP(ByVal strIdentifier As String) As System.Drawing.Bitmap
        Return New System.Drawing.Bitmap(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(strIdentifier))
    End Function
    Friend Function TablesInitialization() As Boolean
        Dim sQuery As String = ""
        Dim sCurrSchema As String = ""
        Dim iCount As Integer = 0
        Dim iCount_NNM1 As Integer = 0
        Dim iCount_NNM2 As Integer = 0
        Dim iCount_NNM3 As Integer = 0
        Dim iCount_NNM4 As Integer = 0
        Dim iCount_NNM5 As Integer = 0
        Dim iCount_NNM6 As Integer = 0
        Dim iCount_NNM7 As Integer = 0

        Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
        Try
            sQuery = " SELECT current_schema FROM DUMMY "
            oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                sCurrSchema = oRec.Fields.Item(0).Value
            End If

            If sCurrSchema.Trim <> "" Then
                sQuery = "  select Count(*) from sys.objects "
                sQuery &= " where ""SCHEMA_NAME"" = '" & sCurrSchema & "' "
                sQuery &= " AND ""OBJECT_TYPE"" = 'TABLE '"
                sQuery &= " AND ""OBJECT_NAME"" = '@NCM_IMAGE' "
                oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                oRec.DoQuery(sQuery)
                If oRec.RecordCount > 0 Then
                    iCount = oRec.Fields.Item(0).Value
                End If

                If iCount <= 0 Then
                    sQuery = " CREATE TABLE ""@NCM_IMAGE"" "
                    sQuery &= " (SrNo         Numeric(9)         NOT NULL,"
                    sQuery &= " Flag         Numeric(9),"
                    sQuery &= " Img1    blob,"
                    sQuery &= " Img2    blob)"
                    oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)

                    sQuery = " INSERT INTO ""@NCM_IMAGE"" VALUES (1,0,NULL,NULL) "
                    oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                End If

                For i As Integer = 1 To 7 Step 1
                    sQuery = "  select Count(*) from sys.objects "
                    sQuery &= " where ""SCHEMA_NAME"" = '" & sCurrSchema & "' "
                    sQuery &= " AND ""OBJECT_TYPE"" = 'VIEW '"
                    sQuery &= " AND ""OBJECT_NAME"" = 'NCM_NNM1_" & i & "' "
                    oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    oRec.DoQuery(sQuery)
                    If oRec.RecordCount > 0 Then
                        iCount_NNM1 = oRec.Fields.Item(0).Value
                    End If

                    If iCount_NNM1 <= 0 Then
                        sQuery = " CREATE VIEW NCM_NNM1_" & i & " AS "
                        sQuery = " SELECT ""ObjectCode"",""Series"",""SeriesName"" FROM ""NNM1"" "
                        sQuery = " UNION ALL "
                        sQuery = " SELECT DISTINCT ""ObjectCode"",-1,'' FROM ""NNM1"" "
                        oRec = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                        oRec.DoQuery(sQuery)
                    End If
                Next
            End If

            oRec = Nothing
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function IsIncludeModule(ByVal RptName As ReportName) As Boolean
        Try
            Dim oRec As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sRptCode As String = GetReportCode(RptName)
            Dim sQry As String = ""
            sQry &= " SELECT IFNULL(""INCLUDED"",'N') FROM ""@NCM_RPT_CONFIG"" WHERE ""RPTCODE"" = '" & sRptCode & "'"
            oRec.DoQuery(sQry)
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                Select Case oRec.Fields.Item(0).Value
                    Case "N", "n", "NO", "No", "no"
                        Return False
                    Case Else
                        Return True
                End Select
            Else
                Return False
            End If
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[IsIncludeModule] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function
    Friend Function GetReportCode(ByVal RptName As ReportName) As String
        Select Case RptName
            Case ReportName.CHANGE_LOG_AUDIT                '55
                Return "CHANGE_LOG_AUDIT"
            Case ReportName.ARSOA_Landscape                 '48
                Return "AR_SOA_LANDSCAPE"
            Case ReportName.GL_Listing                      '1
                Return "GL_LISTING"
            Case ReportName.APSoa                           '2
                Return "AP_SOA"
            Case ReportName.ARSoa                           '3
                Return "AR_SOA"
            Case ReportName.GST                             '4
                Return "RPT_GST"
            Case ReportName.PV                              '5
                Return "PAYMENT_VOUCHER"
            Case ReportName.PV_Range                        '5a
                Return "PAYMENT_VOUCHER_RANGE"
            Case ReportName.RA                              '6
                Return "REMITTANCE_ADVICE"
            Case ReportName.OMARSoa                         '7
                Return "AR_SOA_OM"
            Case ReportName.SAR_FIFO_SUMMARY                '8
                Return "SAR_FIFO_SUMMARY"
            Case ReportName.SAR_MOVAVG_SUMMARY              '9
                Return "SAR_MOVAVG_SUMMARY"
            Case ReportName.SAR_Audit_Enquiry               '10
                Return "SAR_MOVAVG_AUDIT_ENQUIRY"
            Case ReportName.SOA_Email_Config                '11
                Return "AR_SOA_EMAIL"
            Case ReportName.APPayment                       '12
                Return "AP_PAYMENT"
            Case ReportName.ARPayment                       '13
                Return "AR_PAYMENT"
            Case ReportName.PVDraft                         '14
                Return "DRAFT_PAYMENT_VOUCHER"
            Case ReportName.RecpPaym                        '15
                Return "RECEIPT_PAYMENT"
            Case ReportName.SOA_TOS                         '16
                Return "AR_SOA_PROJECT"
            Case ReportName.GPA                             '17
                Return "RPT_GPA"
            Case ReportName.ReOrder_Level_Recommendation    '18
                Return "RPT_RLR"
            Case ReportName.Items_ABC_Analysis              '19
                Return "RPT_IAR"
            Case ReportName.Weighted_Average_Demand         '20
                Return "RPT_WAR"
            Case ReportName.SAR_FIFO_DETAILS                '22
                Return "SAR_FIFO_DETAILS"
            Case ReportName.SAR_MOVAVG_DETAILS              '23
                Return "SAR_MOVAVG_DETAILS"
            Case ReportName.SAR_TM_V1                       '24
                Return "SAR_TM_V1"
            Case ReportName.SAR_TM_V2                       '25
                Return "SAR_TM_V2"
            Case ReportName.SAR_TM_V3                       '26
                Return "SAR_TM_V3"
            Case ReportName.ARAging6B_Details               '27
                Return "AR_AGEING_6B_DETAILS"
            Case ReportName.ARAging6B_Summary               '28
                Return "AR_AGEING_6B_SUMMARY"
            Case ReportName.APAging_Details                 '29
                Return "AP_AGEING_DETAILS"
            Case ReportName.APAging_Summary                 '30
                Return "AP_AGEING_SUMMARY"
            Case ReportName.ARAging_Details                 '31
                Return "AR_AGEING_DETAILS"
            Case ReportName.ARAging_Summary                 '32
                Return "AR_AGEING_SUMMARY"
            Case ReportName.APAging_Details_Proj            '33
                Return "AP_AGEING_PROJ_DETAILS"
            Case ReportName.APAging_Summary_Proj            '34
                Return "AP_AGEING_PROJ_SUMMARY"
            Case ReportName.ARAging_Details_Proj            '35
                Return "AR_AGEING_PROJ_DETAILS"
            Case ReportName.ARAging_Summary_Proj            '36
                Return "AR_AGEING_PROJ_SUMMARY"
            Case ReportName.PO_Detail_Proj                  '44
                Return "PO_DETAILS_BY_CUSTOMER"
            Case ReportName.SO_Detail_Proj                  '43
                Return "SO_DETAILS_BY_CUSTOMER"
            Case ReportName.ARAging7B_Details               '45
                Return "AR_AGEING_7B_DETAILS"
            Case ReportName.ARAging7B_Summary               '46
                Return "AR_AGEING_7B_SUMMARY"
            Case ReportName.OfficialReceipt ' 47
                Return "OFFICIAL_RECEIPT"
            Case ReportName.IRA                             '21
                Return "RPT_IRA"
            Case ReportName.ARAgeingDetailsCRM  '49
                Return "AR_AGEING_DETAILS_CRM"
            Case ReportName.BankReconciliation  '50
                Return "BANK_RECONCILIATION"
            Case ReportName.ARSOA_BY_PROJECT                         '51
                Return "ARSOA_BY_PROJECT"
            Case ReportName.PV_Mass_Email                          '51
                Return "PV_MASS_EMAIL"
        End Select
        Return "NULL"
    End Function
    Friend Function GetSharedFilePath(ByVal sReportName As String) As String
        Try
            Dim oRecord As Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
            Dim sQuery As String = ""
            sQuery &= " SELECT IFNULL(""FILEPATH"",'') "
            sQuery &= " FROM """ & oCompany.CompanyDB & """.""@NCM_RPT_CONFIG"" "
            sQuery &= " WHERE ""RPTCODE"" ='" & GetReportCode(sReportName) & "'"
            oRecord.DoQuery(sQuery)
            If oRecord.RecordCount > 0 Then
                oRecord.MoveFirst()
                Return oRecord.Fields.Item(0).Value
            Else
                Return ""
            End If
            Return ""
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[GetSharedFilePath] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return ""
        End Try
    End Function
    Friend Function IsSharedFilePathExists(ByVal sfilepath As String) As Boolean
        Try
            If File.Exists(sfilepath) Then
                Return True
            End If

            SBO_Application.StatusBar.SetText("Filepath does not exist; the add-on will use default layout.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("[IsSharedFilePathExists] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Friend Function GetDateObject(ByVal sDate As String) As Date
        ' Date must be in yyyyMMdd format
        If sDate.Length = 8 Then
            Dim oDate As Date = New Date(Left(sDate, 4), Mid(sDate, 5, 2), Right(sDate, 2))
            Return oDate
        Else
            Return Nothing
        End If
    End Function
    Friend Function GetCurrentDate() As String
        Try
            Dim sQuery As String = ""
            Dim sCurrDate As String = ""
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(BoObjectTypes.BoRecordset)

            sQuery = " SELECT TO_CHAR(current_timestamp, 'YYYYMMDD') from DUMMY "
            oRec.DoQuery(sQuery)
            If oRec.RecordCount > 0 Then
                sCurrDate = oRec.Fields.Item(0).Value
            End If

            Return sCurrDate
        Catch ex As Exception

            Return ""
        End Try
    End Function
    Friend Function GetEmailCCFromUDT() As String
        Try
            Dim sEmail As String = ""
            Dim sReturn As String = ""
            Dim oRec As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRec.DoQuery("SELECT ""U_EmailAdd"" FROM ""@NCMCCEMAIL"" WHERE IFNULL(""U_EmailAdd"",'') <> '' GROUP BY ""U_EmailAdd"" ")
            If oRec.RecordCount > 0 Then
                oRec.MoveFirst()
                While Not oRec.EoF

                    sEmail = oRec.Fields.Item(0).Value.ToString.Trim
                    If sEmail.Substring(sEmail.Length - 1, 1) = ";" Then
                        sEmail = sEmail.Substring(0, sEmail.Length - 1)
                    End If
                    sEmail = sEmail.Trim

                    sReturn &= sEmail & ";"
                    oRec.MoveNext()
                End While
            End If
            oRec = Nothing
            If sReturn.Length > 0 Then
                sReturn = sReturn.Substring(0, sReturn.Length - 1)
            End If
            Return sReturn
        Catch ex As Exception
            'SBO_Application.StatusBar.SetText("[GetEmailCCFromUDT] : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return ""
        End Try
    End Function
#End Region

#Region "HANA Connection"
    Public Function ExecuteHANACommandToDataAdapter(ByVal sQuery As String) As System.Data.Common.DbDataAdapter
        Dim sFormattedQuery As String = ""
        Dim sCompanyDB As String
        Dim SqlConn As DbConnection = Nothing
        Dim DA As DbDataAdapter = Nothing
        Dim Command As DbCommand = Nothing
        Try

            Dim _DbProviderFactoryObject As DbProviderFactory = Nothing

            SqlConn = GetHANAConnection(_DbProviderFactoryObject)
            SqlConn.Open()

            DA = _DbProviderFactoryObject.CreateDataAdapter()
            Command = SqlConn.CreateCommand()
            sCompanyDB = """" & oCompany.CompanyDB & """."
            sFormattedQuery = String.Format(sQuery, sCompanyDB)
            Command.CommandText = sFormattedQuery
            Command.ExecuteNonQuery()
            DA.SelectCommand = Command
            Return DA
        Finally

        End Try
    End Function

    Public Function GetHANAConnection(ByRef _DbProviderFactoryObject As DbProviderFactory) As DbConnection

        _DbProviderFactoryObject = DbProviderFactories.GetFactory(ProviderName)
        Dim HANADbConnection As DbConnection
        HANADbConnection = _DbProviderFactoryObject.CreateConnection()
        HANADbConnection.ConnectionString = connStr

        Return HANADbConnection
    End Function

    Public Function ExecuteHANACommandToDataTable(ByVal sQuery As String) As System.Data.DataTable
        Dim dt As System.Data.DataTable = New System.Data.DataTable()
        Dim sFormattedQuery As String = ""
        Dim sCompanyDB As String
        Dim SqlConn As DbConnection = Nothing
        Dim DA As DbDataAdapter = Nothing
        Dim Command As DbCommand = Nothing
        Try
            Dim _DbProviderFactoryObject As DbProviderFactory = Nothing

            SqlConn = GetHANAConnection(_DbProviderFactoryObject)
            SqlConn.Open()

            DA = _DbProviderFactoryObject.CreateDataAdapter()
            Command = SqlConn.CreateCommand()

            sCompanyDB = """" & oCompany.CompanyDB & """."
            sFormattedQuery = String.Format(sQuery, sCompanyDB)
            Command.CommandText = sFormattedQuery
            Command.ExecuteNonQuery()
            DA.SelectCommand = Command
            DA.Fill(dt)
            Return dt
        Finally
            If SqlConn IsNot Nothing Then
                If SqlConn.State = ConnectionState.Open Then
                    SqlConn.Close()
                    SqlConn.Dispose()
                End If
            End If
            If DA IsNot Nothing Then
                DA.Dispose()
            End If
            If Command IsNot Nothing Then
                Command.Dispose()
            End If
        End Try
    End Function
#End Region

#Region "Event Handler"
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.ItemEvent
        Try
            Select Case FormUID
                Case FRM_CHG_LOG_AUDIT
                    BubbleEvent = oNCM_CHG_LOG_AUDIT.ItemEvent(pVal)
                Case FRM_NCM_BREC
                    BubbleEvent = oNCM_BREC.ItemEvent(pVal)
                Case FRM_NCM_PO_PROJ
                    BubbleEvent = oNCM_PO_PROJ.ItemEvent(pVal)
                Case FRM_NCM_SO_PROJ
                    BubbleEvent = oNCM_SO_PROJ.ItemEvent(pVal)
                Case FRM_RPT_CONFIG
                    BubbleEvent = oNCM_RPT_CONFIG.ItemEvent(pVal)
                Case "NCM_SOA_TOS"
                    BubbleEvent = oSOA_TOS.ItemEvent(pVal)
                Case "NCM_GPA"
                    BubbleEvent = oGPA.ItemEvent(pVal)
                Case "ncmARAgeingCRM"
                    BubbleEvent = oNCM_AR_AGEING_CRM.ItemEvent(pVal)
                Case "ncmARAgeing_Proj"
                    BubbleEvent = oARAgeing_Proj.ItemEvent(pVal)
                Case "ncmAPAgeing_Proj"
                    BubbleEvent = oAPAgeing_Proj.ItemEvent(pVal)
                Case "ncmARAgeing"
                    BubbleEvent = oARAgeing.ItemEvent(pVal)
                Case "ncmARAging6B"
                    BubbleEvent = oARAgeing6B.ItemEvent(pVal)
                Case FRM_ARAGEING_7B
                    BubbleEvent = oARAgeing7B.ItemEvent(pVal)
                Case "ncmAPAgeing"
                    BubbleEvent = oAPAgeing.ItemEvent(pVal)
                Case "ncmSOA_AP"
                    BubbleEvent = oApSOA.SBO_Application_ItemEvent(pVal)
                Case "NCM_ARSOA"
                    BubbleEvent = oArSOA.SBO_Application_ItemEvent(pVal)
                Case FRM_ARSOA_PROJ
                    BubbleEvent = oARSOAPROJ.SBO_Application_ItemEvent(pVal)
                Case "NCM_OMARSOA"
                    BubbleEvent = oOMArSOA.SBO_Application_ItemEvent(pVal)
                Case FRM_NCM_FIFO1
                    BubbleEvent = oFIFO_NonBatch.ItemEvent(pVal)
                Case FRM_NCM_MOV1
                    BubbleEvent = oMOV_StockAging.ItemEvent(pVal)
                Case "NCM_RPT_GSTT"
                    BubbleEvent = oNCM_gst.ItemEvent(pVal)
                Case FRM_NCM_SES1
                    BubbleEvent = oStockAudit.ItemEvent(pVal)
                Case FRM_NCM_SES2
                    BubbleEvent = oStockAudit_Out.ItemEvent(pVal)
                Case FRM_NCM_PV_RANGE
                    BubbleEvent = oFRM_PaymentVoucher_Range.ItemEvent(pVal)
                Case "ncm_TMSAR"
                    BubbleEvent = oTM_SAR.ItemEvent(pVal)
                Case "ncmSOA_Email"
                    BubbleEvent = oFrmEmail.SBO_Application_ItemEvent(pVal)
                Case "ncmSOA_SendEmail"
                    BubbleEvent = oFrmSendEmail.SBO_Application_ItemEvent(pVal)
                Case "ncmSOA_SendEmailProj"
                    BubbleEvent = oFrmSendEmailProj.SBO_Application_ItemEvent(pVal)
                Case FRM_NCM_AP_PAYMENT
                    BubbleEvent = oAPPayment.ItemEvent(pVal)
                Case FRM_NCM_AR_PAYMENT
                    BubbleEvent = oARPayment.ItemEvent(pVal)
                Case "NCM_RecpPaymList"                       'JN added V03.28.2007
                    BubbleEvent = ofrmRecpPaym.ItemEvent(pVal) 'JN added V03.28.2007
                Case FRM_NCM_GLL
                    BubbleEvent = oNCM_GLL.ItemEvent(pVal) 'ES added V04.12.2007
                Case FRM_NCM_RLR
                    BubbleEvent = oNCM_RLR.ItemEvent(pVal)  'ES added on 10.11.2011
                Case "ncmPV_Email"
                    BubbleEvent = oFrmPVEmail.SBO_Application_ItemEvent(pVal) 'AT added on 04.05.2019
                Case "ncmPV_SendEmail"
                    BubbleEvent = oFrmPVSendEmail.SBO_Application_ItemEvent(pVal)
                Case "ncmPV_Email_Param"
                    BubbleEvent = oPVEmailParam.SBO_Application_ItemEvent(pVal)
                Case Else
                    Select Case pVal.FormTypeEx
                        Case 426
                            BubbleEvent = oPaymentVoucher.ItemEvent(pVal)
                        Case 170
                            BubbleEvent = oIncomingPayment.ItemEvent(pVal)
                        Case 655
                            BubbleEvent = oPaymentDraft.ItemEvent(pVal) 'JN added
                    End Select
            End Select

        Catch ex As Exception
            'SBO_Application.StatusBar.SetText("[ItemEvent]:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction Then
                Dim sFormName As String = ""
                Try
                    sFormName = SBO_Application.Forms.ActiveForm.TypeEx
                Catch ex As Exception
                    ' Throw New System.Exception(ex.Message)
                End Try

                Select Case sFormName
                    Case 426
                        BubbleEvent = oPaymentVoucher.MenuEvent(pVal)
                    Case 170
                        BubbleEvent = oIncomingPayment.MenuEvent(pVal)
                    Case 655
                        BubbleEvent = oPaymentDraft.MenuEvent(pVal) 'JN added
                End Select
            Else
                Select Case pVal.MenuUID
                    Case MNU_CHG_LOG_AUDIT
                        oNCM_CHG_LOG_AUDIT.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_BREC
                        oNCM_BREC.LoadForm()
                        BubbleEvent = False

                    Case MNU_ARAGEING_CRM
                        oNCM_AR_AGEING_CRM.ShowForm()
                        BubbleEvent = False

                    Case MNU_ARAGEING_7B
                        oARAgeing7B.ShowForm()
                        BubbleEvent = False

                    Case MNU_NCM_PO_PROJ
                        oNCM_PO_PROJ.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_SO_PROJ
                        oNCM_SO_PROJ.LoadForm()
                        BubbleEvent = False

                    Case MNU_SUB_IAR
                        SBO_Application.StatusBar.SetText("Viewing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oNCM_IAR.PrintReport()
                        BubbleEvent = False

                    Case MNU_SUB_WAR
                        SBO_Application.StatusBar.SetText("Viewing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        oNCM_WAR.PrintReport()
                        BubbleEvent = False

                    Case MNU_SUB_RLR
                        oNCM_RLR.LoadForm()
                        BubbleEvent = False

                    Case MNU_RPT_CONFIG
                        oNCM_RPT_CONFIG.LoadForm()
                        BubbleEvent = False

                    Case "NCM_GPA"
                        oGPA.LoadForm()
                        BubbleEvent = False

                    Case "NCM_SOATOS"
                        oSOA_TOS.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_GLL
                        oNCM_GLL.LoadForm()
                        BubbleEvent = False

                    Case "NCM_ARAgeing_Proj"
                        oARAgeing_Proj.ShowForm()
                        BubbleEvent = False

                    Case "NCM_APAgeing_Proj"
                        oAPAgeing_Proj.ShowForm()
                        BubbleEvent = False

                    Case "NCM_ARAgeing"
                        oARAgeing.ShowForm()
                        BubbleEvent = False

                    Case "NCM_ARAgeing6B"
                        oARAgeing6B.ShowForm()
                        BubbleEvent = False

                    Case "NCM_APAgeing"
                        oAPAgeing.ShowForm()
                        BubbleEvent = False

                    Case "NCM_SOA_AP"
                        oApSOA.LoadForm()
                        BubbleEvent = False

                    Case "NCM_SOA"
                        oArSOA.LoadForm()
                        BubbleEvent = False

                    Case MNU_ARSOA_PROJ
                        oARSOAPROJ.LoadForm()
                        BubbleEvent = False

                    Case "NCM_OMARSOA"
                        oOMArSOA.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_FIFO1
                        oFIFO_NonBatch.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_MOV1
                        oMOV_StockAging.LoadForm()
                        BubbleEvent = False

                    Case "NCMmnuGST"
                        oNCM_gst.ShowForm()
                        BubbleEvent = False

                    Case MNU_NCM_SES1
                        oStockAudit.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_PV_RANGE
                        oFRM_PaymentVoucher_Range.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_TM_STOCK
                        oTM_SAR.ShowForm()
                        BubbleEvent = False

                    Case MNU_NCM_EMAIL
                        oFrmEmail.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_AP_PAYMENT
                        oAPPayment.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_AR_PAYMENT
                        oARPayment.LoadForm()
                        BubbleEvent = False

                    Case "MNU_NCM_RECPPAYM"
                        ofrmRecpPaym.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_PV_EMAIL
                        oFrmPVEmail.LoadForm()
                        BubbleEvent = False

                    Case MNU_NCM_PV_EMAIL_PARA
                        oPVEmailParam.LoadForm()
                        BubbleEvent = False

                    Case Else
                        Dim sFormName As String = ""
                        Try
                            sFormName = SBO_Application.Forms.ActiveForm.TypeEx
                        Catch ex As Exception
                            ' Throw New System.Exception(ex.Message)
                        End Try
                        Select Case sFormName
                            Case 426
                                BubbleEvent = oPaymentVoucher.MenuEvent(pVal)
                            Case 170
                                BubbleEvent = oIncomingPayment.MenuEvent(pVal)
                            Case 655
                                BubbleEvent = oPaymentDraft.MenuEvent(pVal) 'JN added
                        End Select
                End Select
            End If

        Catch ex As Exception
            'SBO_Application.StatusBar.SetText("[MenuEvent]:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
    Private Sub UnsetConnectionContext()
        If (oCompany.Connected = True) Then oCompany.Disconnect()
        ' UPGRADE_NOTE: Object oCmpany may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        oCompany = Nothing
        TmpThread.Start()
    End Sub
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                    RemoveMenuItems()
                    UnsetConnectionContext()
                Case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition
                    RemoveMenuItems()
                    UnsetConnectionContext()
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    RemoveMenuItems()
                    UnsetConnectionContext()
            End Select

            'RemoveMenuItems()
            ''If SQLDbConnection.State = ConnectionState.Open Then
            ''    SQLDbConnection.Close()
            ''End If

            'If oCompany.Connected = True Then
            '    oCompany.Disconnect()
            '    oCompany = Nothing
            'End If
            'TmpThread.Start()
        Catch ex As Exception
            'SBO_Application.StatusBar.SetText("[AppEvent]:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        End Try
    End Sub
#End Region

End Module