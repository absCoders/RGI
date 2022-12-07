Imports DPayments.InPay
Imports System.Xml.Serialization
Imports System.IO
Imports System.Runtime.Serialization
Imports System.Net
Imports System.Text

' AVS Restrictions for Payeezy
' https://support.payeezy.com/hc/en-us/articles/203730469-Address-Verification-System-AVS-Filters

' Regency login credential to log into Payeezy to see API settings.
' https://globalgatewaye4.firstdata.com/?lang=en
' User id: Anabogo
' Password: Regency100
' Home -> Terminals -> Regency Internaltional-> API Access

'Payeezy Gateway(gwPayeezy)

'Supported Methods :  
'   Sale
'   AuthOnly
'   Capture
'   Refund
'   Credit
'   VoidTransaction
'   Force

'MerchantLogin And MerchantPassword are required properties.

'The TransactionAmount should be represented As dollars And cents With a Decimal point. For example, "1.00".

'The "CurrencyCode" configuration setting Is available for this gateway. This field Is only sent if a value Is explicitly specified.

'This gateway supports sending ThreeDSecure (3DS) verification data by setting the following configs: CAVV, XID, ECI.

'TestMode Is Not supported And when set to "True" an exception will be thrown by the component.

'The "HashSecret" configuration setting Is required by this gateway as an Hmac calculation must be computed And sent in the Authorization header of the request. 
'"HashSecret" must be set to the Hmac Key generated for you by the gateway. Note the component handles the computation of the Hmac.

'The "FDMSKeyId" configuration setting Is also required by this gateway. This configuration setting Is used To specify the Key Id, obtained from FDMS, 
'that corresponds to the HMAC Key (specified via HashSecret) And Is sent within the Authorization header of the request.

'The FullName Is also required To be specified And an exception will be thrown If Not Set.

'AuthCode Is used to perform tagged transactions (transactions that do Not require the Card data to be specified). 
'The value to specify within AuthCode Is the value contained within ApprovalCode after a successful authorization. 
'This transactions are only applicable for Capture, VoidTransaction, And Refund transactions.

<Serializable()> _
Public Class ARCCCARD

#Region "*************** TABLES NEEDED ***************"

    'DROP TABLE ARTCCPA1 CASCADE CONSTRAINTS ; 

    'CREATE TABLE ARTCCPA1 ( 
    '  CCPA_NO                    VARCHAR2 (10), 
    '  CUST_CODE                  VARCHAR2 (10), 
    '  CCPA_STATUS                VARCHAR2 (1), 
    '  CCPA_REASON                VARCHAR2 (1), 
    '  CCPA_NOTE                  VARCHAR2 (60), 
    '  CCPA_AMT                   NUMBER (13,2), 
    '  CCPA_DATE_AUTH             DATE, 
    '  CCPA_DATE_SALE             DATE, 
    '  CCPA_AUTH                  VARCHAR2 (1), 
    '  INIT_OPER                  VARCHAR2 (8), 
    '  INIT_DATE                  DATE, 
    '  LAST_OPER                  VARCHAR2 (8), 
    '  LAST_DATE                  DATE, 
    '  CUST_CREDIT_CARD_NO        VARCHAR2 (100), 
    '  CUST_CREDIT_CARD_EXP_DATE  VARCHAR2 (80), 
    '  CUST_CREDIT_CARD_VER_CODE  VARCHAR2 (80), 
    '  CUST_CREDIT_CARD_NAME      VARCHAR2 (35), 
    '  CUST_CREDIT_CARD_ADDR1     VARCHAR2 (35), 
    '  CUST_CREDIT_CARD_CITY      VARCHAR2 (25), 
    '  CUST_CREDIT_CARD_STATE     VARCHAR2 (2), 
    '  CUST_CREDIT_CARD_ZIP_CODE  VARCHAR2 (10),
    '  CUST_CREDIT_CARD_COUNTRY   VARCHAR2(2),
    '  CUST_CREDIT_CARD_LAST4     VARCHAR2 (4), 
    '  RESPONSE_RETRIEVAL_NO      VARCHAR2 (8), 
    '  RESPONSE_CODE              VARCHAR2 (1), 
    '  RESPONSE_BATCH_NO          VARCHAR2 (15), 
    '  RESPONSE_APPROVAL_CODE     VARCHAR2 (10), 
    '  RESPONSE_TEXT              VARCHAR2 (100), 
    '  CCPA_TYPE                  VARCHAR2 (1), 
    '  ORDR_NO                    VARCHAR2 (10), 
    '  INV_NO                     VARCHAR2 (10), 
    '  CCPA_DATE_VOID             DATE, 
    '  CCPA_REASON_VOID           VARCHAR2 (60), 
    '  STMT_NO                    VARCHAR2 (10), 
    '  CCPA_NO_CREDITED           VARCHAR2 (10), 
    '  CUST_CREDIT_CARD_TYPE      VARCHAR2 (4), 
    '  OPS_YYYYPP                 VARCHAR2 (6), 
    '  LENS_BANK_INV_NO           VARCHAR2 (10), 
    '  CCPA_NO_AUTH               VARCHAR2 (10), 
    '  CCPA_NO_CAPTURE            VARCHAR2 (10), 
    '  CARD_LEVEL_RESULT          VARCHAR2 (2), 
    '  DATAWIRE_RETURN_CODE       VARCHAR2 (4), 
    '  DATAWIRE_STATUS            VARCHAR2 (20), 
    '  ACI_CODE                   VARCHAR2 (2), 
    '  TRANSACTION_DATE           VARCHAR2 (6), 
    '  TRANS_ID                   VARCHAR2 (30), 
    '  TRANS_NUM                  VARCHAR2 (10), 
    '  VALIDATION_CODE            VARCHAR2 (4), 
    '  WEB_PYMT_ID                NUMBER (20),
    'PRIMARY KEY (CCPA_NO));

    'DROP TABLE ARTCCPA2 CASCADE CONSTRAINTS ; 

    'CREATE TABLE ARTCCPA2 ( 
    '  CCPA_NO                 VARCHAR2 (10), 
    '  RESPONSE_TEXT           VARCHAR2 (100), 
    '  RESPONSE_SEQ_NO         VARCHAR2 (6), 
    '  RESPONSE_RETRIEVAL_NO   VARCHAR2 (8), 
    '  RESPONSE_CODE           VARCHAR2 (1), 
    '  RESPONSE_BATCH_NO       VARCHAR2 (6), 
    '  RESPONSE_APPROVAL_CODE  VARCHAR2 (10), 
    '  RESPONSE_DATA           VARCHAR2 (500), 
    '  RESPONSE_AVS            VARCHAR2 (1), 
    '  RESPONSE_AUTH_SOURCE    VARCHAR2 (1), 
    '  INIT_OPER               VARCHAR2 (8), 
    '  INIT_DATE               DATE, 
    '  CCPA_TYPE               VARCHAR2 (1),
    '  PRIMARY KEY (CCPA_NO));

    'DROP TABLE ARTCCPDA CASCADE CONSTRAINTS ; 

    'CREATE TABLE ARTCCPDA ( 
    '  CCPA_NO             VARCHAR2 (10)  NOT NULL, 
    '  DETAIL_AGGREGATE    VARCHAR2 (2000), 
    '  LEVEL2_AGGREGATE    VARCHAR2 (2000), 
    '  LEVEL3_AGGREGATE    VARCHAR2 (3000), 
    '  LEVEL3_AGGREGATE_2  VARCHAR2 (3000), 
    '  LEVEL3_AGGREGATE_3  VARCHAR2 (3000), 
    '  COMM_CARD_TYPE      NUMBER (3), 
    '  LEVEL3_AGGREGATE_4  VARCHAR2 (3000), 
    '  LEVEL3_AGGREGATE_5  VARCHAR2 (3000),
    '  PRIMARY KEY (CCPA_NO)); 

    'DROP TABLE ARTCCPRC CASCADE CONSTRAINTS ; 

    'CREATE TABLE ARTCCPRC ( 
    '  AR_PARM_ACCOUNT            VARCHAR2 (10)  NOT NULL, 
    '  AR_PARM_SERVER             VARCHAR2 (100), 
    '  AR_PARM_MERCHANT_NO        VARCHAR2 (20), 
    '  AR_PARM_TERMINAL_NO        VARCHAR2 (10), 
    '  AR_PARM_CLIENT_NO          VARCHAR2 (4), 
    '  AR_PARM_USER_ID            VARCHAR2 (15), 
    '  AR_PARM_PASSWORD           VARCHAR2 (25), 
    '  AR_PARM_DATAWIRE_ID        VARCHAR2 (30), 
    '  AR_PARM_FDMS_VISA_ID       VARCHAR2 (20), 
    '  AR_PARM_FDMS_MERCH_TAX_ID  VARCHAR2 (10), 
    '  AR_PARM_FDMS_DATAWIRE_URL  VARCHAR2 (30), 
    '  AR_PARM_FDMS_SERVER_ALT    VARCHAR2 (100), 
    '  AR_PARM_AUTH_TRANS_KEY     VARCHAR2 (25), 
    '  AR_PARM_CC_PROC_TYPE       VARCHAR2 (1), 
    '  PRIMARY KEY ( AR_PARM_ACCOUNT ) ) ; 

    'DROP TABLE ARTCUSTC CASCADE CONSTRAINTS ; 

    'CREATE TABLE ARTCUSTC ( 
    '  CUST_CODE                   VARCHAR2 (10)  NOT NULL, 
    '  CUST_CREDIT_CARD_NO         VARCHAR2 (100)  NOT NULL, 
    '  CUST_CREDIT_CARD_EXP_DATE   VARCHAR2 (80), 
    '  CUST_CREDIT_CARD_VER_CODE   VARCHAR2 (80), 
    '  CUST_CREDIT_CARD_NAME       VARCHAR2 (35), 
    '  CUST_CREDIT_CARD_ADDR1      VARCHAR2 (35), 
    '  CUST_CREDIT_CARD_CITY       VARCHAR2 (25), 
    '  CUST_CREDIT_CARD_STATE      VARCHAR2 (2), 
    '  CUST_CREDIT_CARD_ZIP_CODE   VARCHAR2 (10), 
    '  CUST_CREDIT_CARD_COUNTRY    VARCHAR2 (2),     
    '  CUST_CREDIT_CARD_STATUS     VARCHAR2 (1), 
    '  INIT_OPER                   VARCHAR2 (20), 
    '  INIT_DATE                   DATE, 
    '  LAST_OPER                   VARCHAR2 (20), 
    '  LAST_DATE                   DATE, 
    '  CUST_CREDIT_CARD_PREFERRED  VARCHAR2 (1), 
    '  CUST_CREDIT_CARD_LAST4      VARCHAR2 (4), 

    '  PRIMARY KEY ( CUST_CODE, CUST_CREDIT_CARD_NO ) ) ; 

#End Region

#Region "Constants"

    Private Const SSLEnabledProtocols As Int32 = 4032

    Public Enum CreditCardNames
        AmericanExpress = 1
        DinersClubCarteBlanche = 2
        Discover = 3
        JCB = 4
        MasterCard = 5
        Visa = 6
    End Enum

    Public serializeTransactions As Boolean = False
    Private tblARTCCPRC As DataTable = Nothing

    ' Credit Card Authorizations contain a Max Auth period before the Sales of the Auth will generate additional charges.
    Private VisaAuthDays As Int16 = 7
    Private MasterCardAuthDays As Int16 = 7
    Private DiscoveraAuthDays As Int16 = 7
    Private AmexAuthDays As Int16 = 7

#End Region

#Region "Enumerations"

    ' The URL is the gateway's URL not the URL to post transactiosn to.
    Public Enum Gateways
        <EnumMember(Value:="No Gateway")>
        NoGateway = IchargeGateways.gwNoGateway ' (0)

        <EnumMember(Value:="Authorize.Net AIM")>
        AuthorizeNet_AIM = IchargeGateways.gwAuthorizeNet  ' (1)	http://www.authorize.net

        <EnumMember(Value:="eProcessing Transparent Database Engine")>
        eProcessingTransparentDatabaseEngine = IchargeGateways.gwEprocessing  ' (2)	http://www.eProcessingNetwork.com

        <EnumMember(Value:="Intellipay ExpertLink")>
        IntellipayExpertLink = IchargeGateways.gwIntellipay   ' (3)	http://www.intellipay.com

        <EnumMember(Value:="iTransact RediCharge HTML")>
        iTransactRediChargeHTML = IchargeGateways.gwITransact ' (4)	http://www.itransact.com

        <EnumMember(Value:="NetBilling DirectMode")>
        NetBillingDirectMode = IchargeGateways.gwNetBilling ' (5)	http://www.netbilling.com

        <EnumMember(Value:="PayFlow Pro")>
        PayFlowPro = IchargeGateways.gwPayFlowPro ' (6)	https://www.paypal.com/webapps/mpp/payflow-payment-gateway

        <EnumMember(Value:="USA ePay CGI Transaction Gateway")>
        USAePayCGITransactionGateway = IchargeGateways.gwUSAePay ' (7)	http://www.usaepay.com

        <EnumMember(Value:="Plug 'n Pay")>
        PlugnPay = IchargeGateways.gwPlugNPay ' (8)	http://www.plugnpay.com

        <EnumMember(Value:="Planet Payment iPay")>
        PlanetPaymentiPay = IchargeGateways.gwPlanetPayment ' (9)	http://planetpayment.com/

        <EnumMember(Value:="MPCS")>
        MPCS = IchargeGateways.gwMPCS ' (10)	http://merchantcommerce.net/

        <EnumMember(Value:="RTWare")>
        RTWare = IchargeGateways.gwRTWare ' (11)	http://www.rtware.net/

        <EnumMember(Value:="ECX")>
        ECX = IchargeGateways.gwECX ' (12)	http://www.ecx.com

        <EnumMember(Value:="Bank of America (Global Gateway e4)")>
        BankOfAmericaGlobalGatewaye4 = IchargeGateways.gwBankOfAmerica '  (13)	http://bankofamerica.com/merchantservices

        <EnumMember(Value:="Innovative Gateway (PHP)")>
        InnovativeGatewayPHP = IchargeGateways.gwInnovative  '  (14)	http://www.innovativegateway.com

        <EnumMember(Value:="Merchant Anywhere (Transaction Central)")>
        MerchantAnywhereTransactionCentral = IchargeGateways.gwMerchantAnywhere ' (15)	http://www.merchantanywhere.com/

        <EnumMember(Value:="SkipJack")>
        SkipJack = IchargeGateways.gwSkipjack ' (16)	http://www.skipjack.com

        <EnumMember(Value:="3 Delta Systems (3DSI) EC-Linx (18)")>
        DeltaSystems3DSI = IchargeGateways.gw3DSI  ' (18)	http://www.3dsi.com

        <EnumMember(Value:="TrustCommerce API")>
        TrustCommerceAPI = IchargeGateways.gwTrustCommerce ' (19)	http://www.trustcommerce.com

        <EnumMember(Value:="TrustCommerce API")>
        PSIGateHTML = IchargeGateways.gwPSIGate ' (20)	http://www.psigate.com

        <EnumMember(Value:="PayFuse XML (ClearCommerce Engine)")>
        PayFuseXMLClearCommerceEngine = IchargeGateways.gwPayFuse ' (21)	http://www.firstnationalmerchants.com/

        <EnumMember(Value:="LinkPoint")>
        LinkPoint = IchargeGateways.gwLinkPoint ' (24)	http://www.linkpoint.com

        <EnumMember(Value:="Moneris eSelect Plus Canada")>
        MoneriseSelectPlusCanada = IchargeGateways.gwMoneris  ' (25)	http://www.moneris.com

        <EnumMember(Value:="uSight Gateway Post-Auth (")>
        uSightGatewayPostAuth = IchargeGateways.gwMoneris  ' (26)	No Longer in Use

        <EnumMember(Value:="Fast Transact VeloCT (Direct Mode)")>
        FastTransacVeloCTDirectMode = IchargeGateways.gwFastTransact ' (27)	http://www.fasttransact.com/

        <EnumMember(Value:="NetworkMerchants Direct-Post API")>
        NetworkMerchantsDirectPostAPI = IchargeGateways.gwNetworkMerchants ' (28)	http://www.nmi.com/

        <EnumMember(Value:="Ingenico DirectLink / Ogone")>
        IngenicoDirectLinkOgone = IchargeGateways.gwIngenico ' (29)	https://www.ingenico.be/

        <EnumMember(Value:="TransFirst Transaction Central Classic (formerly PRIGate)")>
        TransFirstTransactionCentralClassicformerlyPRIGate = IchargeGateways.gwPRIGate ' (30)	https://www.transfirst.com

        <EnumMember(Value:="Merchant Partners (Transaction Engine)")>
        MerchantPartnersTransactionEngine = IchargeGateways.gwMerchantPartners ' (31)	http://www.merchantpartners.com/

        <EnumMember(Value:="CyberCash")>
        CyberCash = IchargeGateways.gwCyberCash ' (32)	https://www.paypal.com/cybercash

        <EnumMember(Value:="First Data Global Gateway (Linkpoint)")>
        FirstDataGlobalGatewayLinkpoint = IchargeGateways.gwFirstData ' (33)	http://www.firstdata.com

        <EnumMember(Value:="YourPay (Linkpoint)")>
        YourPayLinkpoint = IchargeGateways.gwYourPay '  (34)	http://www.yourpay.com

        <EnumMember(Value:="ACH Payments AGI")>
        ACHPaymentsAGI = IchargeGateways.gwACHPayments ' (35)	http://www.ach-payments.com

        <EnumMember(Value:="Forte AGI / Payments Gateway AGI")>
        ForteAGIPaymentsGatewayAGI = IchargeGateways.gwForte ' (36)	https://www.forte.net/

        <EnumMember(Value:="Cyber Source SOAP API")>
        CyberSourceSOAPAPI = IchargeGateways.gwCyberSource ' (37)	http://www.cybersource.com

        <EnumMember(Value:="eWay XML API (Australia)")>
        eWayXMLAPIAustralia = IchargeGateways.gwCyberSource ' (38)	http://www.eway.com.au/

        <EnumMember(Value:="goEmerchant XML")>
        goEmerchantXML = IchargeGateways.gwGoEMerchant ' (39)	http://www.goemerchant.com/

        <EnumMember(Value:="TransFirst eLink")>
        TransFirsteLink = IchargeGateways.gwTransFirst  ' (40)	http://www.transfirst.com

        <EnumMember(Value:="Chase Merchant Services (Linkpoint)")>
        ChaseMerchantServices_Linkpoint = IchargeGateways.gwChase ' (41)	http://www.chase.com

        <EnumMember(Value:="Thompson Merchant Services NexCommerce (iTransact mode)")>
        ThompsonMerchantServicesNexCommerce_iTransact_mode = IchargeGateways.gwNexCommerce ' (42)	http://www.thompsonmerchant.com

        <EnumMember(Value:="WorldPay Select Junior Invisible")>
        WorldPaySelectJuniorInvisible = IchargeGateways.gwWorldPay  ' (43)	http://www.worldpay.com

        <EnumMember(Value:="TransFirst Transaction Central")>
        TransFirstTransactionCentral = IchargeGateways.gwTransactionCentral ' (44)	http://www.transfirst.com. (This is different from TransFirst eLink, supported above. The TransactionCentral gateway is also used by MerchantAnywhere and PRIGate)

        <EnumMember(Value:="Sterling SPOT XML API (HTTPS POST)")>
        SterlingSPOTXMLAPIHTTPSPOST = IchargeGateways.gwSterling ' (45)	http://www.sterlingpayment.com

        <EnumMember(Value:="PayJunction Trinity Gateway")>
        PayJunctionTrinityGateway = IchargeGateways.gwPayJunction ' (46)	http://www.payjunction.com

        <EnumMember(Value:="SECPay (United Kingdom) API Solution")>
        SECPay_UnitedKingdom_APISolution = IchargeGateways.gwPayJunction ' (47)	http://www.secpay.com

        <EnumMember(Value:="Payment Express PXPost ")>
        PaymentExpressPXPost = IchargeGateways.gwPaymentExpress ' (48)	http://www.paymentexpress.com

        <EnumMember(Value:="Elavon/NOVA/My Virtual Merchant")>
        Elavon_NOVA_MyVirtualMerchant = IchargeGateways.gwMyVirtualMerchant ' (49)	https://support.convergepay.com/s/

        <EnumMember(Value:="Sage Payment Solutions (Bankcard HTTPS Post protocol)")>
        SagePaymentSolutions = IchargeGateways.gwSagePayments ' (50)	http://www.sagepayments.com

        <EnumMember(Value:="SecurePay (Script API/COM Object Interface)")>
        SecurePay = IchargeGateways.gwSagePayments ' (51)	http://securepay.com

        <EnumMember(Value:="Moneris eSelect Plus USA")>
        MoneriseSelectPlusUSA = IchargeGateways.gwMonerisUSA ' (52)	http://www.moneris.com

        <EnumMember(Value:="Bambora / Beanstream Process Transaction API")>
        BamboraBeanstreamProcessTransactionAPI = IchargeGateways.gwBambora ' (53)	https://www.bambora.com/en/ca/

        <EnumMember(Value:="Verifi Direct-Post API")>
        VerifiDirectPostAPI = IchargeGateways.gwVerifi ' (54)	http://www.verifi.com

        <EnumMember(Value:="SagePay Direct (Previously Protx) ")>
        SagePayDirect = IchargeGateways.gwSagePay ' (55)	https://www.opayo.uk/

        <EnumMember(Value:="Merchant E-Solutions Payment Gateway (Trident API)")>
        MerchantESolutionsPaymentGateway_TridentAPI = IchargeGateways.gwMerchantESolutions '  (56)	http://merchante-solutions.com/

        <EnumMember(Value:="PayLeap Web Services API")>
        PayLeapWebServicesAPI = IchargeGateways.gwPayLeap ' (57)	http://www.payleap.com

        <EnumMember(Value:="Worldpay XML (Direct/Invisible)")>
        WorldpayXMLDirectInvisible = IchargeGateways.gwWorldPayXML '  (59)	http://www.worldpay.com

        <EnumMember(Value:="ProPay Merchant Services API")>
        ProPayMerchantServiceAPI = IchargeGateways.gwProPay ' (60)	http://www.propay.com

        '<EnumMember(Value:="Intuit QuickBooks Merchant Services (QBMS)")>
        'IntuitQuickBooksMerchantServices_QBMS = 61 '  (61)	This gateway is no longer in service. It has been replaced by Quickbooks Payments (113).

        <EnumMember(Value:="Heartland POS Gateway")>
        HeartlandPOSGateway = IchargeGateways.gwHeartland ' (62)	http://www.heartlandpaymentsystems.com/

        <EnumMember(Value:="Litle Online Gateway")>
        LitleOnlineGateway = IchargeGateways.gwLitle ' (63)	http://www.litle.com/

        <EnumMember(Value:="BrainTree DirectPost (Server-to-Server Orange) Gateway")>
        BrainTreeDirectPost_ServerToServerOrange_Gateway = IchargeGateways.gwBrainTree ' (64)	http://www.braintreepaymentsolutions.com/

        <EnumMember(Value:="JetPay Gateway")>
        JetPayGateway = IchargeGateways.gwJetPay ' (65)	http://www.jetpay.com/

        <EnumMember(Value:="HSBC XML API (ClearCommerce Engine)")>
        HSBCXMAPIClearCommerceEngine = IchargeGateways.gwHSBC '  (66)	https://www.business.hsbc.uk/en-gb/payments/business-card

        <EnumMember(Value:="BluePay 2.0 Post")>
        BluePay20Post = IchargeGateways.gwBluePay ' (67)	http://www.bluepay.com

        <EnumMember(Value:="PayTrace Payment Gateway")>
        PayTracePaymentGateway = IchargeGateways.gwPayTrace  ' (70)	https://www.paytrace.net/

        <EnumMember(Value:="TransNational Bankcard")>
        TransNationalBankcard = IchargeGateways.gwTransNationalBankcard ' (74)	http://www.tnbci.com/

        <EnumMember(Value:="First Data Global Gateway E4")>
        FirstDataGlobalGatewayE4 = IchargeGateways.gwFirstDataE4 '  (80)	http://www.firstdata.com

        <EnumMember(Value:="Bluefin")>
        Bluefin = IchargeGateways.gwBluefin ' (82)	http://www.bluefin.com/

        <EnumMember(Value:="Payscape")>
        Payscape = IchargeGateways.gwPayscape ' (83)	http://www.payscape.com

        <EnumMember(Value:="Pay Direct (Link2Gov)")>
        PayDirectLink2Gov = IchargeGateways.gwPayDirect '  (84)	https://www.fisglobal.com/solutions/other/government/

        <EnumMember(Value:="WorldPay US Link Gateway")>
        WorldPayUSLinkGateway = IchargeGateways.gwWorldPayLink ' (87)	https://www.worldpay.com/en-us/index

        <EnumMember(Value:="3DSI Payment WorkSuite")>
        i3DSIPaymentWorkSuite = IchargeGateways.gwPaymentWorkSuite ' (88)	http://www.3dsi.com/

        <EnumMember(Value:="First Data PayPoint")>
        FirstDataPayPoint = IchargeGateways.gwFirstDataPayPoint ' (90)	https://www.firstdata.com/en_us/customer-center/financial-institutions/paypoint.html

        <EnumMember(Value:="Converge (formerly MyVirtualMerchant)")>
        ConvergeformerlyMyVirtualMerchant = IchargeGateways.gwConverge '  (93)	https://support.convergepay.com/s/

        <EnumMember(Value:="Payeezy Gateway (formerly First Data E4)")>
        PayeezyGatewayformerlyFirstDataE4 = IchargeGateways.gwPayeezy '  (94)	https://developer.payeezy.com/

        <EnumMember(Value:="Authorize.NET XML")>
        AuthorizeNETXML = IchargeGateways.gwAuthorizeNetXML ' (96)	http://www.authorize.net

        <EnumMember(Value:="PhoeniXGate Gateway")>
        PhoeniXGateGateway = IchargeGateways.gwPhoeniXGate ' (97)	http://www.phoenixmanagednetworks.com/

        <EnumMember(Value:="Repay Gateway")>
        RepayGateway = IchargeGateways.gwRepay ' (98)	https://www.repay.com/

        <EnumMember(Value:="BASYS Gateway")>
        BASYSGateway = IchargeGateways.gwBASYS ' (106)	https://basyspro.com/

        <EnumMember(Value:="Quickbooks Payments")>
        QuickbooksPayments = IchargeGateways.gwQBPayments ' (113)	https://quickbooks.intuit.com/payments/

        <EnumMember(Value:="Shift4")>
        Shift4 = IchargeGateways.gwShift4 ' (114)	https://www.shift4.com/

    End Enum

    Enum CreditCardTypes
        vctAmex = CardvalidatorCardTypes.vctAmex
        vctBankCard = CardvalidatorCardTypes.vctAmex
        vctCUP = CardvalidatorCardTypes.vctCUP
        vctDiners = CardvalidatorCardTypes.vctDiners
        vctDiscover = CardvalidatorCardTypes.vctDiscover
        vctJCB = CardvalidatorCardTypes.vctJCB
        vctLaser = CardvalidatorCardTypes.vctLaser
        vctMaestro = CardvalidatorCardTypes.vctMaestro
        vctMasterCard = CardvalidatorCardTypes.vctMasterCard
        'vctMCardPurchase = CardvalidatorCardTypes.vctMCardPurchase
        'vctSolo = CardvalidatorCardTypes.vctSolo
        'vctSwitch = CardvalidatorCardTypes.vctSwitch
        'vctTempoPayments = CardvalidatorCardTypes.vctTempoPayments
        vctUnknown = CardvalidatorCardTypes.vctUnknown
        vctVisa = CardvalidatorCardTypes.vctVisa
        vctVisaElectron = CardvalidatorCardTypes.vctVisaElectron
        'vctVisaPurchase = CardvalidatorCardTypes.vctVisaPurchase
    End Enum

#End Region

#Region "Class Variables"

    Public Class Address
        Public Address1 As String = String.Empty
        Public Address2 As String = String.Empty
        Public City As String = String.Empty
        Public State As String = String.Empty
        Public ZipCode As String = String.Empty
        Public Country As String = String.Empty
    End Class

    Public Class clsNetworkResponse
        Private clsIcharge As Icharge

        Public Sub New(Icharge As Icharge)
            clsIcharge = Icharge
        End Sub

        ''' <summary>
        ''' Extracts the value of a node from the XML reasponse.
        ''' If the node does not exist or there is a data error then the Empty String is returned.
        ''' </summary>
        ''' <param name="NodeName"></param>
        ''' <returns></returns>
        Public Function ExtractNodeFromResponse(ByVal NodeName As String) As String

            Try
                Dim doc As New XmlDocument
                doc.LoadXml(Data)
                Dim value As String = doc.SelectSingleNode($"TransactionResult/{NodeName}").InnerText
                Return value

            Catch ex As Exception
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' Contains an authorization code for an approved transaction.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ApprovalCode() As String
            Get
                Return clsIcharge.Response.ApprovalCode
            End Get
        End Property

        ''' <summary>
        ''' Indicates whether the transaction was successful (True) or unsuccessful (False).
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Approved() As Boolean
            Get
                Return clsIcharge.Response.Approved
            End Get
        End Property

        ''' <summary>
        ''' The amount approved for the transaction, this is the amount actually charged to the credit card.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ApprovedAmount() As Double
            Get
                Return clsIcharge.Response.ApprovedAmount
            End Get
        End Property

        ''' <summary>
        ''' Contains the Address Verification System result code. Used for fraud detection.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property AVSResult() As String
            Get
                Return clsIcharge.Response.AVSResult
            End Get
        End Property

        ''' <summary>
        ''' Indicates the success or failure of the transaction.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Code As String
            Get
                Return clsIcharge.Response.Code
            End Get
        End Property

        ''' <summary>
        ''' Contains the returned CVV result code if it was requested. Used for fraud detection.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property CVVResult As String
            Get
                Return clsIcharge.Response.CVVResult
            End Get
        End Property

        ''' <summary>
        ''' Contains the raw response from the host.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Data As String
            Get
                Return clsIcharge.Response.Data
            End Get
        End Property

        ''' <summary>
        ''' Additional code returned for declined or failed transactions.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ErrorCode As String
            Get
                Return clsIcharge.Response.ErrorCode
            End Get
        End Property

        ''' <summary>
        ''' Description of the error which occurred.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ErrorText As String
            Get
                Return clsIcharge.Response.ErrorText
            End Get
        End Property

        ''' <summary>
        ''' Merchant-generated invoice number echoed back in the response.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property InvoiceNumber As String
            Get
                Return clsIcharge.Response.InvoiceNumber
            End Get
        End Property

        ''' <summary>
        ''' Return code generated by the processor, or additional gateway response code that may contain more information beyond "Approved" or "Declined".
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ProcessorCode As String
            Get
                Return clsIcharge.Response.ProcessorCode
            End Get
        End Property

        ''' <summary>
        ''' Contains a human-readable message explaining the code.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Text As String
            Get
                Return clsIcharge.Response.Text
            End Get
        End Property

        ''' <summary>
        ''' Host-generated transaction identifier, used for Captures, Credits, or Voids.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property TransactionId As String
            Get
                Return clsIcharge.Response.TransactionId
            End Get
        End Property

        Public ReadOnly DetailAggregate As String = String.Empty
    End Class

    Public NetworkResponse As clsNetworkResponse

    Public Class GatewayData

        Public Sub New()

        End Sub

        Public GateWayName As String = String.Empty
        Public GateWayURL As String = String.Empty
        Public GateWayTestURL As String = String.Empty
    End Class

    Private clsIcharge As Icharge
    Private clsCardValidator As Cardvalidator

    Private clsGateWay As IchargeGateways = IchargeGateways.gwNoGateway
    Private clsGatewayURL As String = String.Empty

    Private clsGatewayData As GatewayData

    Public TransactionAmount As Double = 0
    Public TransactionNumber As String = String.Empty
    Public DetailAggregate As String = String.Empty
    Public BatchNumber As String = String.Empty
    Public BatchStatus As String = String.Empty

    Public CreditCardProcessingNo As String = String.Empty
    Public InternalReference As String = String.Empty

    ' Sub classes Instantiation
    Public MerchantAccount As New Merchant
    Public CustomerCreditCard As New CreditCard
    Public Level2Data As New Level2
    Public Level3Data As New List(Of Level3)

    Private clsLogFileLocation As String = String.Empty
    Private cXmlFileName As String = String.Empty
    Private cXmlDirectory As String = String.Empty

    Private cTestMode As Boolean = False
    Private clsEncryptionClass As Object = Nothing

    Private rawRequest As String = String.Empty
    Private rawResponse As String = String.Empty

    Private clsLastError As String = String.Empty


#End Region

#Region "Class Instantiation"

    ''' <summary>
    ''' Instantiate class
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        clsGateWay = Gateways.NoGateway
        clsLogFileLocation = String.Empty
        InitializeVariables()
    End Sub

    ''' <summary>
    ''' Instantiate class,
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal slogFileLocation As String)
        clsGateWay = Gateways.NoGateway
        clsLogFileLocation = slogFileLocation
        InitializeVariables()
    End Sub

    ''' <summary>
    ''' Instantiate class
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(ByVal slogFileLocation As String, ByVal Gateway As Gateways)
        clsGateWay = Gateway
        clsLogFileLocation = slogFileLocation
        InitializeVariables()
    End Sub

    ''' <summary>
    ''' Instantiate class
    ''' </summary>
    ''' <param name="slogFileLocation">Location to place serialized object</param>
    ''' <param name="encryptionClass">Encryption Class. Must have Functions Ecrypt, Decrypt - each accepting a single string parameter
    ''' and a return value of String</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal slogFileLocation As String, ByVal Gateway As Gateways, encryptionClass As Type)
        clsGateWay = Gateway
        clsLogFileLocation = slogFileLocation
        InitializeVariables()
        clsEncryptionClass = Activator.CreateInstance(encryptionClass)
    End Sub

    ''' <summary>
    ''' Initialize class object and variables
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitializeVariables()

        TransactionAmount = 0
        TransactionNumber = String.Empty
        DetailAggregate = String.Empty
        BatchNumber = String.Empty
        BatchStatus = String.Empty

        ' IpCharge Objects
        clsIcharge = New Icharge
        clsIcharge.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")

        clsCardValidator = New Cardvalidator
        clsCardValidator.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")

        MerchantAccount = New Merchant
        CustomerCreditCard = New CreditCard
        Level2Data = New Level2
        Level3Data = New List(Of Level3)
        Level3Data = New List(Of Level3)

        clsLogFileLocation = clsLogFileLocation.Trim
        If clsLogFileLocation.Length > 0 AndAlso Not My.Computer.FileSystem.DirectoryExists(clsLogFileLocation) Then
            clsLogFileLocation = String.Empty
        End If

        clsEncryptionClass = Nothing

        tblARTCCPRC = ASCDATA1.GetDataTable("SELECT * FROM ARTCCPRC", "ARTCCPRC")

        Dim rowARTCCPRC As DataRow = tblARTCCPRC.Rows.Find({clsGateWay})
        If rowARTCCPRC IsNot Nothing Then
            VisaAuthDays = Val(rowARTCCPRC.Item("VISA_AUTH_MAX_DAYS") & String.Empty)
            MasterCardAuthDays = Val(rowARTCCPRC.Item("MC_AUTH_MAX_DAYS") & String.Empty)
            DiscoveraAuthDays = Val(rowARTCCPRC.Item("DISC_AUTH_MAX_DAYS") & String.Empty)
            AmexAuthDays = Val(rowARTCCPRC.Item("AMEX_AUTH_MAX_DAYS") & String.Empty)
        End If

    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property RawRequestText() As String
        Get
            Return rawRequest
        End Get
    End Property

    Public ReadOnly Property RawResponseText() As String
        Get
            Return rawResponse
        End Get
    End Property


    ''' <summary>
    ''' Indicates the reuslts of the Luhn Check Digit Algorithm
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CreditCardType() As CreditCardTypes
        Get
            Try

                With clsCardValidator
                    ' Since just checking the card type use the current month and year
                    .CardExpMonth = DateAdd(DateInterval.Month, -1, DateTime.Now).Month
                    .CardExpYear = DateTime.Now.Year
                    .CardNumber = CustomerCreditCard.CardNumber
                    .ValidateCard()
                    Return .CardType
                End With
            Catch ex As Exception
                Return CreditCardTypes.vctUnknown
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Get the Location of the XMl file
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property XmlFileName() As String
        Get
            Return cXmlFileName
        End Get
    End Property

    ''' <summary>
    ''' Gets / Set location to place serialized transaction
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property XmlDirectory As String
        Get
            Return cXmlDirectory
        End Get
        Set(value As String)
            cXmlDirectory = value
        End Set
    End Property

    ''' <summary>
    ''' Get set whether in test mode
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TestMode() As Boolean
        Get
            Return cTestMode
        End Get
        Set(ByVal value As Boolean)
            cTestMode = value
        End Set
    End Property

    ''' <summary>
    ''' Get / Set location of logfile to Serialize Class.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LogFileLocation As String
        Get
            Return clsLogFileLocation
        End Get
        Set(value As String)
            clsLogFileLocation = value
        End Set
    End Property

    ''' <summary>
    ''' Last error
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property LastError As String
        Get
            Return clsLastError
        End Get
    End Property

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Used to check the validity of the card without authorizing funds.
    ''' </summary>
    ''' <returns></returns>
    Public Function AVSOnly() As Boolean

        AVSOnly = False
        clsLastError = String.Empty

        Try
            MerchantSetup()

            With clsIcharge
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")
                .Card.ExpMonth = CustomerCreditCard.CardExpMonth

                Dim CardExpYear As String = String.Empty
                CardExpYear = CustomerCreditCard.CardExpYear
                If CardExpYear.Length > 2 Then
                    CardExpYear = CardExpYear.Substring(CardExpYear.Length - 2)
                End If
                .Card.ExpYear = CardExpYear

                Select Case CustomerCreditCard.CardType
                    Case CreditCardTypes.vctAmex : .Card.CardType = TCardTypes.ctAMEX
                    Case CreditCardTypes.vctBankCard
                    Case CreditCardTypes.vctCUP
                    Case CreditCardTypes.vctDiners : .Card.CardType = TCardTypes.ctDiners
                    Case CreditCardTypes.vctDiscover : .Card.CardType = TCardTypes.ctDiscover
                    Case CreditCardTypes.vctJCB : .Card.CardType = TCardTypes.ctJCB
                    Case CreditCardTypes.vctLaser : .Card.CardType = TCardTypes.ctLaser
                    Case CreditCardTypes.vctMaestro : .Card.CardType = TCardTypes.ctMaestro
                    Case CreditCardTypes.vctMasterCard : .Card.CardType = TCardTypes.ctMasterCard
                    Case CreditCardTypes.vctVisa : .Card.CardType = TCardTypes.ctVisa
                    Case CreditCardTypes.vctVisaElectron : .Card.CardType = TCardTypes.ctVisaElectron
                End Select

                .Card.Number = CustomerCreditCard.CardNumber
                .Card.CVVData = CustomerCreditCard.CardCVVData

                '.Customer.FirstName = CustomerCreditCard.CardHolderFirstName
                '.Customer.LastName = CustomerCreditCard.CardHolderLastName
                .Customer.Address = CustomerCreditCard.CardHolderAddress
                .Customer.City = CustomerCreditCard.CardHolderCity
                .Customer.State = CustomerCreditCard.CardHolderState
                .Customer.Zip = CustomerCreditCard.CardHolderZipCode
                .Customer.Country = CustomerCreditCard.CardHolderCountry
                .Customer.Email = CustomerCreditCard.CardHolderEmail
                .Customer.Phone = CustomerCreditCard.CardHolderTelephone

                .AVSOnly()
                AVSOnly = True

                If clsGateWay = IchargeGateways.gwPayeezy AndAlso CustomerCreditCard.TransArmorToken.Length = 0 Then
                    CustomerCreditCard.TransArmorToken = .Config("FDMSTransArmorToken") & String.Empty
                End If

            End With

        Catch exc As InPayIchargeException
            clsLastError = "AVSOnly: " & exc.Message
        Catch ex As Exception
            clsLastError = "AVSOnly: " & ex.Message
        Finally
            rawRequest = clsIcharge.Config("RawRequest")
            rawResponse = clsIcharge.Config("RawResponse")
        End Try

    End Function

    ''' <summary>
    ''' Initiates an authorization-only request transaction.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub AuthOnly()
        AuthOnlySale("AUTH_ONLY")
        ExportSerializedObject()
    End Sub

    ''' <summary>
    ''' Retrieves the current state of the open batch. 
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub BatchInquiry()
        ExportSerializedObject()
    End Sub

    ''' <summary>
    ''' Releases the current batch for settlement.
    ''' </summary>
    ''' <param name="NetDeposit"></param>
    ''' <remarks></remarks>
    Public Sub BatchRelease(ByVal NetDeposit As String)
        ExportSerializedObject()
    End Sub

    ''' <summary>
    ''' Captures a previously authorized transaction.
    ''' 
    ''' </summary>
    ''' <param name="TransactionID"></param>
    ''' <param name="CUST_CREDIT_CARD_NAME"></param>
    ''' <param name="AUTH_CODE"></param>
    ''' <returns></returns>
    Public Function Capture(ByVal TransactionID As String, ByVal CUST_CREDIT_CARD_NAME As String, ByVal AUTH_CODE As String) As Boolean

        Capture = False
        clsLastError = String.Empty

        Try
            MerchantSetup()
            clsIcharge.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")

            clsIcharge.Customer.FullName = CUST_CREDIT_CARD_NAME
            clsIcharge.AuthCode = AUTH_CODE
            clsIcharge.Capture(TransactionID, TransactionAmount)
            Capture = True
            NetworkResponse = New clsNetworkResponse(clsIcharge)

        Catch exc As InPayIchargeException
            clsLastError = ("Capture: Error [" & exc.Code.ToString & "]: " & exc.Message)
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Catch ex As Exception
            clsLastError = ("Capture: " & ex.Message)
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Finally
            rawRequest = clsIcharge.Config("RawRequest")
            rawResponse = clsIcharge.Config("RawResponse")
        End Try

        ExportSerializedObject()
    End Function

    ''' <summary>
    ''' Captures a previously authorized transaction.
    ''' </summary>
    ''' <param name="CreditCardInfo"></param>
    ''' <returns></returns>
    Public Function Capture(ByVal CreditCardInfo As CreditCard) As Boolean

        Capture = False
        clsLastError = String.Empty

        Try
            MerchantSetup()

            ' If Authorization is more than 7 days old then Void the current Authorization and do a Sale.
            If IsDate(CreditCardInfo.AuthorizationDate) Then
                If DateAndTime.DateDiff(DateInterval.Day, CreditCardInfo.AuthorizationDate, DateTime.Now) > 7 Then
                    CreditCardInfo.RefundAmount = CreditCardInfo.CaptureAmount
                    VoidTransaction(CreditCardInfo)
                    TransactionAmount = CreditCardInfo.CaptureAmount
                    Sale()
                    Return True
                End If
            End If

            clsIcharge.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")
            clsIcharge.Customer.FullName = (CreditCardInfo.CardHolderFirstName & " " & CreditCardInfo.CardHolderLastName).Trim
            clsIcharge.AuthCode = CreditCardInfo.ResponseApprovalCode

            With clsIcharge
                .Card.ExpMonth = CreditCardInfo.CardExpMonth
                .Card.ExpYear = CreditCardInfo.CardExpYear

                .Card.Number = CustomerCreditCard.CardNumber
                .Card.CVVData = CustomerCreditCard.CardCVVData

                .Customer.Address = CreditCardInfo.CardHolderAddress
                .Customer.City = CreditCardInfo.CardHolderCity
                .Customer.State = CreditCardInfo.CardHolderState
                .Customer.Zip = CreditCardInfo.CardHolderZipCode
                .Customer.Country = CreditCardInfo.CardHolderCountry
                .Customer.Email = CreditCardInfo.CardHolderEmail
                .Customer.Phone = CreditCardInfo.CardHolderTelephone

                Try
                    ValidateCard()
                Catch ex As Exception

                End Try

                Dim CardLast4Digits As String = StrReverse(StrReverse(CustomerCreditCard.CardNumber).Substring(0, 4))
                .Config($"CardLast4Digits={CardLast4Digits}")

                .InvoiceNumber = CreditCardInfo.InvoiceNumber

                Select Case .Gateway
                    Case IchargeGateways.gwPayeezy
                        Select Case .Card.CardType
                            Case TCardTypes.ctMasterCard, TCardTypes.ctVisa
                                GenerateLevelAggregate(CustomerCreditCard)
                                .Level2Aggregate = GetLevel2Aggregate()
                                .Level3Aggregate = GetLevel3Aggregate()
                        End Select
                End Select

            End With

            clsIcharge.Capture(CreditCardInfo.TransactionID, CreditCardInfo.CaptureAmount)
            Capture = True
            NetworkResponse = New clsNetworkResponse(clsIcharge)

        Catch exc As InPayIchargeException
            clsLastError = ("Capture: Error [" & exc.Code.ToString & "]: " & exc.Message)
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Catch ex As Exception
            clsLastError = ("Capture: " & ex.Message)
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Finally
            rawRequest = clsIcharge.Config("RawRequest")
            rawResponse = clsIcharge.Config("RawResponse")
        End Try

        ExportSerializedObject()
    End Function

    ''' <summary>
    ''' Credits a customer's card.
    ''' </summary>
    ''' <remarks></remarks>
    Public Function Credit() As Boolean

        Credit = False
        clsLastError = String.Empty

        Try
            MerchantSetup()
            clsIcharge.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")

            ' Need to set the Credit Card Number or the last 4 of the Number
            With clsIcharge.Card
                .CVVData = CustomerCreditCard.CardCVVData
                .ExpMonth = CustomerCreditCard.CardExpMonth
                .ExpYear = CustomerCreditCard.CardExpYear
                .Number = CustomerCreditCard.CardNumber

                Dim CardLast4Digits As String = StrReverse(StrReverse(CustomerCreditCard.CardNumber).Substring(0, 4))
                clsIcharge.Config($"CardLast4Digits={CardLast4Digits}")
            End With

            If CustomerCreditCard.TransactionID & String.Empty <> String.Empty Then
                clsIcharge.TransactionId = CustomerCreditCard.TransactionID
            Else
                clsIcharge.TransactionId = TransactionNumber
            End If

            clsIcharge.TransactionAmount = Format(TransactionAmount, "###0.00")
            With clsIcharge.Customer
                .Address = CustomerCreditCard.CardHolderAddress
                '.Address2 = String.Empty
                .City = CustomerCreditCard.CardHolderCity
                .Country = CustomerCreditCard.CardHolderCountry
                .Email = CustomerCreditCard.CardHolderEmail
                '.FirstName = CustomerCreditCard.CardHolderFirstName
                .FullName = (CustomerCreditCard.CardHolderFirstName & " " & CustomerCreditCard.CardHolderLastName).ToString.Trim
                '.LastName = CustomerCreditCard.CardHolderLastName
                .Phone = CustomerCreditCard.CardHolderTelephone
                .State = CustomerCreditCard.CardHolderState
                .Zip = CustomerCreditCard.CardHolderZipCode
            End With

            clsIcharge.Credit()

            Credit = True

            NetworkResponse = New clsNetworkResponse(clsIcharge)

        Catch exc As InPayIchargeException
            clsLastError = ("Credit: " & exc.Message)
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Catch ex As Exception
            clsLastError = ("Credit: " & ex.Message)
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Finally
            rawRequest = clsIcharge.Config("RawRequest")
            rawResponse = clsIcharge.Config("RawResponse")
        End Try

        ExportSerializedObject()
    End Function

    Public Function DateCheckPassed() As Boolean
        Try
            clsCardValidator.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")
            Return clsCardValidator.DateCheckPassed
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function DigitCheckPassed() As Boolean
        Try
            clsCardValidator.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")
            Return clsCardValidator.DigitCheckPassed
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Property Gateway() As Gateways
        Get
            Return clsGateWay
        End Get
        Set(value As Gateways)
            clsGateWay = value
        End Set
    End Property

    ''' <summary>
    ''' Refunds a previously captured transaction.
    ''' </summary>
    ''' <param name="CreditCardInfo"></param>
    ''' <returns></returns>
    Public Function Refund(ByVal CreditCardInfo As CreditCard) As Boolean

        Refund = False
        clsLastError = String.Empty

        Try
            MerchantSetup()
            clsIcharge.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")

            clsIcharge.Customer.FullName = (CreditCardInfo.CardHolderFirstName & " " & CreditCardInfo.CardHolderLastName).Trim
            clsIcharge.AuthCode = CreditCardInfo.ResponseApprovalCode

            With clsIcharge
                .Card.ExpMonth = CreditCardInfo.CardExpMonth
                .Card.ExpYear = CreditCardInfo.CardExpYear

                .Card.Number = CustomerCreditCard.CardNumber
                .Card.CVVData = CustomerCreditCard.CardCVVData

                .Customer.Address = CreditCardInfo.CardHolderAddress
                .Customer.City = CreditCardInfo.CardHolderCity
                .Customer.State = CreditCardInfo.CardHolderState
                .Customer.Zip = CreditCardInfo.CardHolderZipCode
                .Customer.Country = CreditCardInfo.CardHolderCountry
                .Customer.Email = CreditCardInfo.CardHolderEmail
                .Customer.Phone = CreditCardInfo.CardHolderTelephone

                Dim CardLast4Digits As String = StrReverse(StrReverse(CustomerCreditCard.CardNumber).Substring(0, 4))
                .Config($"CardLast4Digits={CardLast4Digits}")
            End With

            clsIcharge.Refund(CreditCardInfo.TransactionID, CreditCardInfo.RefundAmount.ToString("###0.00"))
            Refund = True
            NetworkResponse = New clsNetworkResponse(clsIcharge)

        Catch exc As InPayIchargeException
            clsLastError = ("VoidTransaction: " & exc.Message)
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Catch ex As Exception
            clsLastError = ("VoidTransaction: " & ex.Message)
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Finally
            rawRequest = clsIcharge.Config("RawRequest")
            rawResponse = clsIcharge.Config("RawResponse")
        End Try

        ExportSerializedObject()
    End Function

    ''' <summary>
    ''' Clears all properties to their default values
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset()
        clsGateWay = Gateways.NoGateway
        InitializeVariables()
    End Sub

    ''' <summary>
    ''' Initiates an Sale transaction (authorization and capture).
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Sale()
        AuthOnlySale("SALE")
        ExportSerializedObject()
    End Sub

    'Public Function ValidateAdrress(ByVal inAddress As Address) As Address
    '    Try
    '        Dim ValidtedAddress As New Address

    '        Dim customerAddress As New nsoftware.InShip.AddressDetail
    '        With customerAddress
    '            .Address1 = inAddress.Address1 & String.Empty
    '            .Address2 = inAddress.Address2 & String.Empty
    '            .City = inAddress.City & String.Empty
    '            .CountryCode = inAddress.Country & String.Empty
    '            .State = inAddress.State & String.Empty
    '            .ZipCode = inAddress.ZipCode & String.Empty
    '        End With

    '        Dim EzAddress As New nsoftware.InShip.Ezaddress
    '        EzAddress.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareInship")
    '        EzAddress.Address = customerAddress
    '        EzAddress.ValidateAddress()

    '        With ValidtedAddress
    '            .Address1 = EzAddress.Address.Address1
    '            .Address2 = EzAddress.Address.Address2
    '            .City = EzAddress.Address.City
    '            .State = EzAddress.Address.State
    '            .ZipCode = EzAddress.Address.ZipCode
    '            .Country = EzAddress.Address.CountryCode
    '        End With

    '        Return ValidtedAddress
    '    Catch ex As Exception
    '        Return inAddress
    '    End Try

    'End Function

    Public Function GetCreditCardType() As CardTypes

        Try
            clsCardValidator = New Cardvalidator
            clsCardValidator.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")
            With clsCardValidator
                ' Use next month as the Exp date.
                ' This is used tonly to get the card type
                .CardExpMonth = DateAdd(DateInterval.Month, 1, DateTime.Now).ToString("MM")
                .CardExpYear = DateAdd(DateInterval.Month, 1, DateTime.Now).ToString("yy")
                .CardNumber = CustomerCreditCard.CardNumber
                .ValidateCard()
                Return clsCardValidator.CardType
            End With

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Checks the card number and expiration date for validity.
    ''' </summary>
    ''' <remarks></remarks>
    Public Function ValidateCard() As Boolean

        Try
            clsIcharge.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")
            clsCardValidator = New Cardvalidator
            clsCardValidator.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")

            With clsCardValidator
                If CustomerCreditCard.CardExpMonth.Length > 0 OrElse CustomerCreditCard.CardExpYear.Length > 0 Then
                    .CardExpMonth = CustomerCreditCard.CardExpMonth
                    .CardExpYear = CustomerCreditCard.CardExpYear
                End If
                .CardNumber = CustomerCreditCard.CardNumber
                .ValidateCard()
                CustomerCreditCard.CardType = clsCardValidator.CardType
            End With

            ValidateCard = True

        Catch ex As Exception
            ValidateCard = False
        End Try

    End Function

    ''' <summary>
    '''  Voids a previously authorized transaction.
    ''' </summary>
    ''' <remarks></remarks>
    Public Function VoidTransaction(ByVal CreditCardInfo As CreditCard) As Boolean

        VoidTransaction = False
        clsLastError = String.Empty

        Try
            MerchantSetup()
            clsIcharge.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")

            With clsIcharge
                .Customer.FullName = (CreditCardInfo.CardHolderFirstName & " " & CreditCardInfo.CardHolderLastName).Trim
                .AuthCode = CreditCardInfo.ResponseApprovalCode
                .TransactionAmount = Format(CreditCardInfo.RefundAmount, "###0.00")

                .Card.ExpMonth = CreditCardInfo.CardExpMonth
                .Card.ExpYear = CreditCardInfo.CardExpYear

                .Card.Number = CreditCardInfo.CardNumber
                .Card.CVVData = CreditCardInfo.CardCVVData

                .Customer.Address = CreditCardInfo.CardHolderAddress
                .Customer.City = CreditCardInfo.CardHolderCity
                .Customer.State = CreditCardInfo.CardHolderState
                .Customer.Zip = CreditCardInfo.CardHolderZipCode
                .Customer.Country = CreditCardInfo.CardHolderCountry
                .Customer.Email = CreditCardInfo.CardHolderEmail
                .Customer.Phone = CreditCardInfo.CardHolderTelephone

                Dim CardLast4Digits As String = StrReverse(StrReverse(CreditCardInfo.CardNumber).Substring(0, 4))
                .Config($"CardLast4Digits={CardLast4Digits}")
            End With

            clsIcharge.VoidTransaction(CreditCardInfo.TransactionID)

            VoidTransaction = True
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Catch exc As InPayIchargeException
            clsLastError = ("VoidTransaction: " & exc.Message)
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Catch ex As Exception
            clsLastError = ("VoidTransaction: " & ex.Message)
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Finally
            rawRequest = clsIcharge.Config("RawRequest")
            rawResponse = clsIcharge.Config("RawResponse")
        End Try

        ExportSerializedObject()
    End Function

#End Region

#Region "Private Procedures"

    Private Sub EncodeCreditCardHtmlChars(ByRef creditcardInfo As CreditCard)
        With creditcardInfo
            .CardHolderAddress = EncodeHtmlChars(.CardHolderAddress)
            .CardHolderCity = EncodeHtmlChars(.CardHolderCity)
            .CardHolderCountry = EncodeHtmlChars(.CardHolderCountry)
            .CardHolderFirstName = EncodeHtmlChars(.CardHolderFirstName)
            .CardHolderLastName = EncodeHtmlChars(.CardHolderLastName)
            .CardHolderState = EncodeHtmlChars(.CardHolderState)
            .CardHolderTelephone = EncodeHtmlChars(.CardHolderTelephone)
            .CardHolderZipCode = EncodeHtmlChars(.CardHolderZipCode)
        End With
    End Sub

    Private Function EncodeHtmlChars(ByVal data As String) As String
        Return System.Net.WebUtility.HtmlEncode(data & String.Empty)
    End Function

    Private Sub MerchantSetup()
        clsIcharge.Reset()
        clsIcharge.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")
        EncodeCreditCardHtmlChars(CustomerCreditCard)
        clsIcharge.Config("AllowPartialAuths=False")

        With clsIcharge
            .Gateway = clsGateWay
            ' As per 4D Payments tech support the component will auto choose the production URL for the selected gateway
            '.GatewayURL = MerchantAccount.ProcessingServer

            Select Case clsGateWay
                Case IchargeGateways.gwNoGateway
                    .MerchantLogin = ""
                    .MerchantPassword = ""

                Case IchargeGateways.gwPayeezy, IchargeGateways.gwFirstDataE4
                    ' https://4dpayments.com/kb/epayment-gateway-code-examples/#payeezy

                    ' This is your Gateway ID (ExactID).  This number is of the format Axxxxx-xx
                    .MerchantLogin = MerchantAccount.UserID

                    ' This is your Password
                    .MerchantPassword = MerchantAccount.Password

                    ' The Key Id that corresponds to the HMAC Key for the First Data E4, Payeezy, and Bank Of America gateways.
                    ' This config is used to specify the Key Id, obtained from FDMS, that corresponds to the HMAC Key (specified via HashSecret) and is sent within the Authorization header of the request.

                    ' This is your Key ID
                    .Config($"FDMSKeyId={MerchantAccount.KeyId}")

                    ' This is your HMAC Key
                    .Config($"HashSecret={MerchantAccount.HMACKey}")

                    'HashAlgorithm:      Algorithm used for hashing.
                    'Certain Gateways allow the request to be hashed as an additional authentication mechanism. This configuration setting controls which algorithm Is used for hashing. Valid values are
                    'Value  Algorithm 
                    '0      MD5 (default) 
                    '1      SHA-1 
                    '.Config($"HashAlgorithm=1")

            End Select

        End With
    End Sub

    ''' <summary>
    ''' Returns the Level 2 Aggregate for Settlement Processing
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLevel2Aggregate() As String

        Try
            Dim obJLevel2 As New DPayments.InPay.Level2
            With obJLevel2
                obJLevel2.DutyAmount = Level2Data.DutyAmount
                obJLevel2.FreightAmount = Level2Data.FreightAmount
                .PONumber = Level2Data.PONumber
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")
                .ShipFromZip = Level2Data.ShipFromZip
                .ShipToZip = Level2Data.ShipToZip
                .TaxAmount = Level2Data.TaxAmount
                .TaxExempt = Level2Data.TaxExempt
            End With

            Return obJLevel2.GetAggregate

        Catch ex As Exception
            If ASCMAIN1.USER_ID = "edz" Then
                MessageBox.Show(ex.Message)
            End If
            Return (String.Empty)
        End Try

    End Function

    ''' <summary>
    ''' Returns the Level 3 Aggregate for Settlement Processing
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetLevel3Aggregate() As String


        If Level3Data Is Nothing OrElse Level3Data.Count = 0 Then
            Return String.Empty
        End If

        Try
            If Level3Data Is Nothing OrElse Level3Data.Count = 0 Then
                Return String.Empty
            End If

            Dim objLevel3 As New DPayments.InPay.Level3
            With objLevel3
                .RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")
                .LineItems.Clear()

                For Each level3entry As Level3 In Level3Data
                    Dim lineEntry As New EPLineItem
                    With lineEntry
                        .CommodityCode = level3entry.CommodityCode
                        .Description = level3entry.Description
                        .DiscountAmount = Convert.ToString(Math.Round(level3entry.DiscountAmount, 2) * 100)
                        .DiscountRate = Convert.ToString(Math.Round(level3entry.DiscountRate, 2) * 100)
                        .Name = level3entry.Name
                        .ProductCode = level3entry.ProductCode
                        .Quantity = level3entry.Quantity
                        '.Taxable = ""
                        .TaxRate = Convert.ToString(Math.Round(level3entry.TaxRate, 2) * 100)
                        .TaxAmount = Convert.ToString(Math.Round(level3entry.TaxAmount, 2) * 100)
                        '.TaxType = ""
                        .Total = Convert.ToString(Math.Round(level3entry.Total, 2) * 100)
                        .UnitCost = Convert.ToString(Math.Round(level3entry.UnitCost, 2) * 100)
                        .Units = "each"
                    End With
                    objLevel3.LineItems.Add(lineEntry)
                Next

            End With

            Return objLevel3.GetAggregate

        Catch ex As Exception
            If ASCMAIN1.USER_ID = "edz" Then
                MessageBox.Show(ex.Message)
            End If
            Return String.Empty
        End Try


    End Function

    Private Function AuthOnlySale(ByVal transType As String) As Boolean

        AuthOnlySale = False
        clsLastError = String.Empty

        Try
            MerchantSetup()
            clsIcharge.RuntimeLicense = ASCMAIN1.nSoftwareKeys("4DPayments")

            Dim CardExpYear As String = String.Empty
            CardExpYear = CustomerCreditCard.CardExpYear
            If CardExpYear.Length > 2 Then
                CardExpYear = CardExpYear.Substring(CardExpYear.Length - 2)
            End If

            With clsIcharge
                .InvoiceNumber = TransactionNumber
                .TransactionAmount = Format(TransactionAmount, "###0.00")
                .TransactionId = TransactionNumber
                .TransactionDesc = "Household Items"

                .Card.ExpMonth = CustomerCreditCard.CardExpMonth
                .Card.ExpYear = CardExpYear

                If clsGateWay = IchargeGateways.gwPayeezy AndAlso (CustomerCreditCard.TransArmorToken & String.Empty).Length > 0 Then
                    .Config($"FDMSTransArmorToken={CustomerCreditCard.TransArmorToken}")
                Else
                    .Card.Number = CustomerCreditCard.CardNumber
                    .Card.CVVData = CustomerCreditCard.CardCVVData

                    .Customer.Address = CustomerCreditCard.CardHolderAddress
                    .Customer.City = CustomerCreditCard.CardHolderCity
                    .Customer.State = CustomerCreditCard.CardHolderState
                    .Customer.Zip = CustomerCreditCard.CardHolderZipCode
                    .Customer.Country = CustomerCreditCard.CardHolderCountry
                    .Customer.Email = CustomerCreditCard.CardHolderEmail
                    .Customer.Phone = CustomerCreditCard.CardHolderTelephone

                    Try
                        ValidateCard()
                    Catch ex As Exception

                    End Try

                    Dim CardLast4Digits As String = StrReverse(StrReverse(CustomerCreditCard.CardNumber).Substring(0, 4))
                    .Config($"CardLast4Digits={CardLast4Digits}")
                End If

                Select Case CustomerCreditCard.CardType
                    Case CreditCardTypes.vctAmex : .Card.CardType = TCardTypes.ctAMEX
                    Case CreditCardTypes.vctBankCard
                    Case CreditCardTypes.vctCUP
                    Case CreditCardTypes.vctDiners : .Card.CardType = TCardTypes.ctDiners
                    Case CreditCardTypes.vctDiscover : .Card.CardType = TCardTypes.ctDiscover
                    Case CreditCardTypes.vctJCB : .Card.CardType = TCardTypes.ctJCB
                    Case CreditCardTypes.vctLaser : .Card.CardType = TCardTypes.ctLaser
                    Case CreditCardTypes.vctMaestro : .Card.CardType = TCardTypes.ctMaestro
                    Case CreditCardTypes.vctMasterCard : .Card.CardType = TCardTypes.ctMasterCard
                    Case CreditCardTypes.vctVisa : .Card.CardType = TCardTypes.ctVisa
                    Case CreditCardTypes.vctVisaElectron : .Card.CardType = TCardTypes.ctVisaElectron
                End Select

                .Customer.Id = CustomerCreditCard.CustomerID & String.Empty
                .Customer.FullName = (CustomerCreditCard.CardHolderFirstName & " " & CustomerCreditCard.CardHolderLastName).ToString.Trim

                If transType = "AUTH_ONLY" Then
                    .AuthOnly() ' perform Authorization Only
                Else
                    Select Case .Gateway
                        Case IchargeGateways.gwPayeezy
                            Select Case .Card.CardType
                                Case TCardTypes.ctMasterCard, TCardTypes.ctVisa
                                    GenerateLevelAggregate(CustomerCreditCard)
                                    .Level2Aggregate = GetLevel2Aggregate()
                                    .Level3Aggregate = GetLevel3Aggregate()
                            End Select
                    End Select

                    .Sale() ' perform Sale
                End If

                AuthOnlySale = True

                NetworkResponse = New clsNetworkResponse(clsIcharge)
                If clsGateWay = IchargeGateways.gwPayeezy AndAlso CustomerCreditCard.TransArmorToken.Length = 0 Then
                    CustomerCreditCard.TransArmorToken = .Config("FDMSTransArmorToken") & String.Empty
                End If

            End With

        Catch exc As InPayIchargeException
            clsLastError = "AuthOnlySale: " & exc.Message
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Catch ex As Exception
            clsLastError = "AuthOnlySale: " & ex.Message
            NetworkResponse = New clsNetworkResponse(clsIcharge)
        Finally
            rawRequest = clsIcharge.Config("RawRequest")
            rawResponse = clsIcharge.Config("RawResponse")

            If clsLastError.Length > 0 AndAlso rawResponse.Length > 0 Then
                clsLastError = "AuthOnlySale: " & rawResponse
            End If
        End Try
    End Function

    ''' <summary>
    ''' Credit Card Authorizations contain a Max Auth period before the Sales of the Auth will generate additional charges.
    ''' When an Authorization has expired then we need to Void the Authorization followed by a direct sale.
    ''' </summary>
    ''' <param name="CreditCardType"></param>
    ''' <returns></returns>
    Private Function GetAuthorizationMaxDays(ByVal CreditCardType As CreditCardTypes) As Int16
        Select Case CreditCardType
            Case CreditCardTypes.vctAmex, CreditCardTypes.vctBankCard
                Return AmexAuthDays
            Case CreditCardTypes.vctVisa, CreditCardTypes.vctVisaElectron
                Return VisaAuthDays
            Case CreditCardTypes.vctDiscover
                Return DiscoveraAuthDays
            Case CreditCardTypes.vctMasterCard
                Return MasterCardAuthDays
            Case Else
                Return 30
        End Select
    End Function

    Private Sub GenerateLevelAggregate(ByVal CreditCardInfo As CreditCard)

        Level2Data = New Level2
        Level3Data.Clear()

        Try
            If CreditCardInfo.Level2Data IsNot Nothing Then
                Level2Data = CreditCardInfo.Level2Data
                Level3Data = CreditCardInfo.Level3Data
                Exit Sub
            End If
        Catch ex As Exception

        End Try

        Try
            Dim rowSOTINVH1 As DataRow = CreditCardInfo.invoiceHeaderRow
            If rowSOTINVH1 Is Nothing Then
                Exit Sub
            End If
            Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & String.Empty
            Dim tblSOTINVH2 As DataTable = CreditCardInfo.invoiceDetailsTable

            Dim rowSOTORDR5 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR5 WHERE ORDR_NO = :PARM1 AND CUST_ADDR_TYPE = :PARM2", "VV", {ORDR_NO, "ST"})
            If rowSOTORDR5 Is Nothing Then
                rowSOTORDR5 = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR5 WHERE ORDR_NO = :PARM1", "V", {ORDR_NO})
            End If

            Dim taxrate As Double = 0
            Dim invoiceSales As Double = Val(rowSOTINVH1.Item("INV_SALES") & String.Empty)

            If rowSOTINVH1 IsNot Nothing Then
                Dim WHSE_CODE As String = rowSOTINVH1.Item("WHSE_CODE") & String.Empty
                Dim rowICTWHSE1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTWHSE1 WHERE WHSE_CODE = :PARM1", "V", {WHSE_CODE})

                With Level2Data
                    .DutyAmount = 0
                    .FreightAmount = Val(rowSOTINVH1.Item("INV_FREIGHT") & String.Empty)
                    If rowSOTINVH1.Item("ORDR_CUST_PO") & String.Empty <> String.Empty Then
                        .PONumber = rowSOTINVH1.Item("ORDR_CUST_PO") & String.Empty
                    Else
                        .PONumber = rowSOTINVH1.Item("INV_NO") & String.Empty
                    End If

                    If rowICTWHSE1 IsNot Nothing Then
                        .ShipFromZip = rowICTWHSE1.Item("WHSE_ZIP_CODE") & ""
                    End If

                    If rowSOTORDR5 IsNot Nothing Then
                        .ShipToZip = rowSOTORDR5.Item("CUST_ZIP_CODE") & ""
                    End If

                    .TaxAmount = Val(rowSOTINVH1.Item("INV_STAX") & String.Empty)
                    .TaxExempt = ARCCCARD.Level2.Level2TaxExempts.lFalse

                    If invoiceSales > 0 Then
                        taxrate = .TaxAmount / invoiceSales
                    Else
                        taxrate = 0
                    End If

                End With

                If tblSOTINVH2 IsNot Nothing Then
                    For Each rowSOTINVH2 As DataRow In tblSOTINVH2.Select("ORDR_QTY_SHIP > 0")
                        Dim STYLE_CODE As String = rowSOTINVH2.Item("STYLE_CODE") & String.Empty
                        Dim rowICTSTYL1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTSTYL1 WHERE STYLE_CODE = :PARM1", "V", STYLE_CODE)
                        If rowICTSTYL1 Is Nothing Then
                            Continue For
                        End If

                        With Level3Data
                            Dim level3Item As New TAC.ARCCCARD.Level3
                            With level3Item
                                '.CommodityCode = ""
                                Dim STYLE_DESC As String = ""
                                STYLE_CODE = EncodeHtmlChars(STYLE_CODE)

                                If rowICTSTYL1 IsNot Nothing Then
                                    STYLE_DESC = rowICTSTYL1.Item("STYLE_DESC") & String.Empty
                                Else
                                    STYLE_DESC = STYLE_CODE
                                End If
                                ' remove any HTML tags.
                                STYLE_DESC = EncodeHtmlChars(STYLE_DESC)
                                '.DiscountAmount = 0
                                '.DiscountRate = 0

                                .Description = STYLE_DESC
                                .Name = STYLE_DESC

                                .ProductCode = STYLE_CODE
                                .Quantity = Val(rowSOTINVH2.Item("ORDR_QTY_SHIP") & String.Empty)
                                .UnitCost = Val(rowSOTINVH2.Item("ORDR_UNIT_PRICE") & String.Empty)
                                .Taxable = True
                                .TaxAmount = (.Quantity * .UnitCost) * (taxrate / 100)
                                .TaxRate = Math.Round(taxrate * 100, 0)
                                .Total = Math.Round(((.Quantity * .UnitCost) + .TaxAmount), 2)
                                .Units = "EA"
                            End With
                            .Add(level3Item)
                        End With
                    Next
                End If

            End If
        Catch ex As Exception
            Level2Data = New Level2
            Level3Data.Clear()
        End Try
    End Sub

    Private Function GetTaxInfo(ByVal CUST_CODE As String, ByVal CUST_ADDR_TYPE As String, ByVal CUST_ADDR_CODE As String) As Double

        Dim taxPercent As Double = 0
        Try
            Dim rowARTCUST2 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_ADDR_TYPE = :PARM2 AND CUST_ADDR_CODE = :PARM3", "VVV", New Object() {CUST_CODE, CUST_ADDR_TYPE, CUST_ADDR_CODE})

            If rowARTCUST2 Is Nothing Then
                Return 0
            End If

            Dim STAX_CODE As String = rowARTCUST2.Item("STAX_CODE") & String.Empty

            If STAX_CODE.Length = 0 Then
                Return 0
            End If

            Dim rowARTSTAX1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTSTAX1 WHERE STAX_CODE = :PARM1", "V", New Object() {STAX_CODE})
            If rowARTSTAX1 Is Nothing Then
                Return 0
            End If

            Return Val(rowARTSTAX1.Item("STAX_RATE") & "")
        Catch ex As Exception
            Return 0
        End Try

    End Function

    Private Function TestCardAuthorization() As Boolean

        ' Contains the Gateway to use for Credit Card Processing
        Dim rowSOTPARM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTPARM1 WHERE SO_PARM_KEY = 'Z'")
        Dim SO_PARM_CC_PROC_CODE As String = rowSOTPARM1.Item("SO_PARM_CC_PROC_CODE") & String.Empty
        If SO_PARM_CC_PROC_CODE = "FDMS" Then
            SO_PARM_CC_PROC_CODE = "94"
        End If
        Dim TACMAIN1 As New TAC.TACMAIN1

        ' Contains the Credentials for the Gateway
        Dim rowARTCCPRC As DataRow = ASCDATA1.GetDataRow($"Select * from ARTCCPRC WHERE CC_PROC_CODE = '{SO_PARM_CC_PROC_CODE}'")

        Dim objIcharge As New Icharge
        Dim e4DPaymentsRuntimeLicense As String = TACMAIN1.e4DPayments

        objIcharge.RuntimeLicense = e4DPaymentsRuntimeLicense
        objIcharge.Gateway = Val(rowARTCCPRC.Item("CC_PROC_CODE") & String.Empty)
        objIcharge.MerchantLogin = rowARTCCPRC.Item("CC_PROC_USER_ID") & String.Empty
        objIcharge.MerchantPassword = rowARTCCPRC.Item("CC_PROC_PASSWORD") & String.Empty
        objIcharge.Config($"FDMSKeyId={ rowARTCCPRC.Item("CC_PROC_DATAWIRE_ID") & String.Empty }")
        objIcharge.Config($"HashSecret={ rowARTCCPRC.Item("CC_PROC_AUTH_TRANS_KEY") & String.Empty }")

        objIcharge.Card.ExpMonth = 2
        objIcharge.Card.ExpYear = 11
        objIcharge.Card.Number = "5454545454545454"
        objIcharge.Customer.Address = "1234 Nowhere Ln"
        objIcharge.Customer.Email = "nobody@server.com"

        ' If you have only one name field then use FullName instead of Firstname and Lastname
        ' objIcharge.Customer.FullName = ""
        objIcharge.Customer.FirstName = "John"
        objIcharge.Customer.LastName = "Smith"

        objIcharge.Customer.Zip = "90001"

        ' This must have 2 decimal places. The value of 1 will fail.
        objIcharge.TransactionAmount = "1.00"
        objIcharge.AuthOnly()

        Dim outmsg As String = "Approval Code: " & objIcharge.Response.ApprovalCode & vbCrLf &
       "Response AVS: " & objIcharge.Response.AVSResult & vbCrLf &
       "Response Code: " & objIcharge.Response.Code & vbCrLf &
       "Response CVV2: " & objIcharge.Response.CVVResult & vbCrLf &
       "Invoice Number: " & objIcharge.Response.InvoiceNumber & vbCrLf &
       "Transaction ID: " & objIcharge.Response.TransactionId & vbCrLf &
       "Response Text: " & objIcharge.Response.Text & vbCrLf & vbCrLf

        Return objIcharge.Response.Approved
    End Function


#End Region

#Region "Serialization"

    Private Function SerializeRequest() As String
        Try
            Dim objXS As New XmlSerializer(Me.GetType())
            Dim objSW As New StringWriter

            objXS.Serialize(objSW, Me)

            Dim serializedObject As String = objSW.ToString()

            If clsEncryptionClass IsNot Nothing Then
                serializedObject = clsEncryptionClass.Encrypt(serializedObject)
            End If

            Return serializedObject

        Catch ex As Exception
            Return String.Empty
        End Try

    End Function

    ''' <summary>
    ''' Serial Class to XML
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Serialize() As String

        Try
            Dim objXS As New XmlSerializer(Me.GetType())
            Dim objSW As New StringWriter

            objXS.Serialize(objSW, Me)

            Dim serializedObject As String = objSW.ToString()

            If clsEncryptionClass IsNot Nothing Then
                serializedObject = clsEncryptionClass.Encrypt(serializedObject)
            End If

            Return serializedObject

        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Deserialize Class Object
    ''' </summary>
    ''' <param name="strXML"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeSerialize(ByVal strXML As String) As ARCCCARD

        Try
            Dim objXS As New XmlSerializer(GetType(ARCCCARD))
            Dim objSR As New StringReader(strXML)

            Dim objST As ARCCCARD = CType(objXS.Deserialize(objSR), ARCCCARD)

            Return objST

        Catch ex As Exception
            Return New ARCCCARD(String.Empty)
        End Try
    End Function

    ''' <summary>
    ''' Save contents of object to a log
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ExportSerializedObject()

        If Not serializeTransactions Then
            Exit Sub
        End If

        cXmlFileName = String.Empty

        If (Me.TransactionNumber & String.Empty).ToString.Trim.Length > 0 Then
            cXmlFileName = Me.TransactionNumber.Trim
        Else ' hopefully this will never fire
            cXmlFileName = "D_" & DateTime.Now.ToString("yyyyMMdd_hhmmss")
        End If

        ExportSerializedObject(cXmlFileName, cXmlDirectory)
    End Sub

    ''' <summary>
    ''' Save contents of object to a log
    ''' </summary>
    ''' <param name="filename"></param>
    ''' <remarks></remarks>
    Public Sub ExportSerializedObject(ByRef filename As String, ByVal filePath As String)
        Try

            Dim RandomClass As New Random()
            Dim RandomNumber As Integer

            If filePath.Length = 0 Then
                filePath = clsLogFileLocation.Trim
            End If

            If filePath.Length > 0 AndAlso Not filePath.EndsWith("\") Then
                filePath &= "\"
            End If

            If filePath.Length = 0 OrElse Not My.Computer.FileSystem.DirectoryExists(filePath) Then
                Exit Sub
            End If

            ' Try to get a Unique value that helps identiffy the transaction.
            ' If not available then use the date time
            If filename.Length = 0 Then
                If Me.CreditCardProcessingNo.Trim.Length > 0 Then
                    filename = "C" & Me.CreditCardProcessingNo.Trim
                ElseIf Me.TransactionNumber.Trim.Length > 0 Then
                    filename = "T" & Me.TransactionNumber.Trim
                Else
                    RandomNumber = RandomClass.Next(25000)
                    filename = DateTime.Now.ToString("yyyyMMdd_hhmmss") & "_" & RandomNumber.ToString
                End If
            End If

            filename = filename.ToUpper
            filename = filename.Replace(".XML", String.Empty)

            ' Ensure a unique filename
            Dim randonValue As String = String.Empty
            Dim ictr As Integer = 0
            While My.Computer.FileSystem.FileExists(filename & randonValue & ".xml")
                ictr += 1
                RandomNumber = RandomClass.Next(25000)
                randonValue = "_" & RandomNumber.ToString.Trim
                ' Prevent a run away loop
                If ictr > 25 Then Exit While
            End While

            filename &= randonValue & ".xml"
            Dim serializedData As String = Serialize()

            ' Allow append incase for some reason the file uniqueness fails
            Using writer As StreamWriter = New StreamWriter(filePath & filename, True)
                writer.WriteLine(serializedData)
                writer.Close()
            End Using

        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Deserializes an Object stored on disk
    ''' </summary>
    ''' <param name="filePath">Loaction of object on disk</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ImportSerializedObject(ByVal filePath As String) As ARCCCARD

        If Not My.Computer.FileSystem.FileExists(filePath) Then
            Return Nothing
        End If

        ' Read in the serialized class
        Dim strXML As String
        Dim tReader As IO.TextReader = New IO.StreamReader(filePath)
        strXML = tReader.ReadToEnd

        If clsEncryptionClass IsNot Nothing Then
            strXML = strXML.Trim
            strXML = clsEncryptionClass.Decrypt(strXML)
        End If

        Return DeSerialize(strXML)
    End Function

#End Region

#Region "Other Classes"

    <Serializable()>
    Public Class Merchant
        Public ProcessingServer As String = String.Empty
        Public UserID As String = String.Empty
        Public Password As String = String.Empty
        Public KeyId As String = String.Empty
        Public HMACKey As String = String.Empty
        Public HashSecret As String = String.Empty
    End Class

    <Serializable()>
    Public Class CreditCard

        Public CardNumber As String = String.Empty
        Public CardExpMonth As String = String.Empty
        Public CardExpYear As String = String.Empty
        Public CardCVVData As String = String.Empty
        Public CardType As CreditCardTypes = CreditCardTypes.vctUnknown

        Public CardHolderFirstName As String = String.Empty
        Public CardHolderLastName As String = String.Empty
        Public CardHolderAddress As String = String.Empty
        Public CardHolderCity As String = String.Empty
        Public CardHolderState As String = String.Empty
        Public CardHolderZipCode As String = String.Empty
        Public CardHolderCountry As String = String.Empty
        Public CardHolderEmail As String = String.Empty
        Public CardHolderTelephone As String = String.Empty

        Public CustomerID As String = String.Empty
        Public InvoiceNumber As String = String.Empty
        Public TransArmorToken As String = String.Empty

        Public TransactionID As String = String.Empty
        Public ResponseApprovalCode As String = String.Empty
        Public RefundAmount As Double = 0
        Public CaptureAmount As Double = 0
        Public AuthorizationDate As Date = DateTime.Now

        Public invoiceHeaderRow As DataRow = Nothing
        Public invoiceDetailsTable As DataTable = Nothing

        Public Level2Data As Level2 = Nothing
        Public Level3Data As List(Of Level3)

        Public Sub New()
        End Sub

    End Class

    <Serializable()>
    Public Class Level2

        Public Enum Level2TaxExempts
            lNotProvides = DPayments.InPay.Level2TaxExempts.teNotProvided
            lTrue = DPayments.InPay.Level2TaxExempts.teTrue
            lFalse = DPayments.InPay.Level2TaxExempts.teFalse
        End Enum

        Public DutyAmount As Double = 0
        Public FreightAmount As Double = 0
        Public PONumber As String = String.Empty
        Public ShipFromZip As String = String.Empty
        Public ShipToZip As String = String.Empty
        Public TaxAmount As Double = 0
        Public TaxExempt As Level2TaxExempts = Level2TaxExempts.lNotProvides

        Public Sub Clear()
            DutyAmount = 0
            FreightAmount = 0
            PONumber = String.Empty
            ShipFromZip = String.Empty
            ShipToZip = String.Empty
            TaxAmount = 0
            TaxExempt = Level2TaxExempts.lNotProvides
        End Sub

        Public Sub New()

        End Sub

    End Class

    <Serializable()>
    Public Class Level3

        Public CommodityCode As String = String.Empty
        Public Description As String = String.Empty
        Public DiscountAmount As Double = 0
        Public DiscountRate As Double = 0
        Public Name As String = String.Empty
        Public ProductCode As String = String.Empty
        Public Quantity As Int32 = 0
        Public Taxable = False
        Public TaxAmount As Double = 0
        Public TaxRate As Double = 0 ' 6.00 percent = 600
        Public Total As Double = 0
        Public UnitCost As Double = 0
        Public Units As String = "each"
        Public TaxType As String = String.Empty

        Public Sub New()
        End Sub
    End Class

#End Region

#Region "Payeezy"

    ''' <summary>
    ''' Downloads Payeezy transactions for the provided date range
    ''' </summary>
    ''' <param name="startDate">Search Start Date</param>
    ''' <param name="endDate">Search End Date</param>
    ''' <returns></returns>
    Public Function GetPayeezyTransactions(ByVal startDate As Date, ByVal endDate As Date) As Boolean
        clsLastError = String.Empty

        MerchantSetup()

        ' Return the transactions for the given period and default account
        ' https://api.globalgatewaye4.firstdata.com/transaction/search?start_date=2010-06-01%2000:00:00&end_date=2010-06-01%2023:59:59
        ' curl -i -u wilma:w_pass -H 'Accept: text/search-v3+csv' "https://api.globalgatewaye4.firstdata.com/transaction/search?start_date=2014-03-01&end_date=2014..."

        Try
            Dim myReq As HttpWebRequest
            Dim myResp As HttpWebResponse
            ServicePointManager.SecurityProtocol = DirectCast(4032, SecurityProtocolType)

            Dim url As String = "https://api.globalgatewaye4.firstdata.com/transaction/search"
            Dim myrequest As String = $"?start_date={startDate.ToString("yyyy-MM-dd")}%2000:00:00&end_date={endDate.ToString("yyyy-MM-dd")}%2023:59:59"

            'myReq = HttpWebRequest.Create(url)
            myReq = HttpWebRequest.Create(url & myrequest)

            myReq.Method = "POST"
            myReq.ContentType = "application/text"

            Dim credstring As String = "Anabogo:Regency100"
            credstring = $"{clsIcharge.MerchantLogin}:{clsIcharge.MerchantPassword}"
            Dim authstring As String = Convert.ToBase64String(Encoding.UTF8.GetBytes(credstring))
            myReq.Headers.Add("Authorization", "Basic " & authstring)
            'myReq.GetRequestStream.Write(System.Text.Encoding.UTF8.GetBytes(myrequest), 0, System.Text.Encoding.UTF8.GetBytes(myrequest).Count)
            myResp = myReq.GetResponse
            Dim myreader As New System.IO.StreamReader(myResp.GetResponseStream)
            Dim myText As String
            myText = myreader.ReadToEnd
            Return True

        Catch ex As Exception
            clsLastError = $"GetPayeezyTransactions Error: {ex.Message}"
            Return False
        End Try

    End Function
#End Region

End Class
