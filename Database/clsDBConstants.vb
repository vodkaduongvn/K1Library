#Region " File Information "

'======================================================================
' This class contains all the database constants for Recfind
'======================================================================

#Region " Revision History "

'======================================================================
' Name      Date        Description
'----------------------------------------------------------------------
' KSD       16/01/2007  Implemented.
'======================================================================

#End Region

#End Region

Public Class clsDBConstants

#Region " Enumerations "

    'The Table Class Types
    Public Enum enumTableClass
        SYSTEM_ESSENTIAL = 2            'cant delete table
        SYSTEM_LOCKED = 3               'cant delete table, cant add new field from, cant add/delete data
        APPLICATION = 5                 'user can change structure and data
        LINK_TABLE = 6                  'link table
        LINK_TABLE_ESSENTIAL = 7        'system essential link table, can't delete
    End Enum

    'AutoFill Table FillTypes
    Public Enum enumAutoFillTypes
        TODAY_DATE = 1                  'Today's Date
        LOGGED_IN_PERSON = 2            'Logged in Person
        SELECTED_ITEM_EXTERNALID = 3    'Parent Object's ExternalID
        SELECTED_ITEM_TABLE = 5         'Table of Last Selected Item
        LAST_MOVEMENT = 6               'The Previous Movement
        MONTHS_FROM_DATE = 7            'A Month from Today's date
        LOGGED_PERSON_SECURITYID = 8    'Logged in Person's SecurityID
        LOGGED_IN_USER_PROFILE = 9      'Logged in Person
        AUTO_NUMBER_FORMAT = 10         'AutoNumbered Field
        USER_VALUE = 11                 'User designated Value
        WEEKS_FROM_DATE = 12            'A week from today's date
        DAYS_FROM_DATE = 13             '3 months from today's date
        YEARS_FROM_DATE = 14            'A year from today's date
        [NOTHING] = Integer.MinValue
    End Enum

    'Filter Table FilterTypes 
    Public Enum enumFilterTypes
        [NOTHING] = Integer.MinValue
        TABLE = 1
        IS_NULL = 2
        LOGGED_IN_PERSON = 3
        LOGGED_IN_USERPROFILE = 4
        PARENT_FIELD = 5
        USER_VALUE = 6
        TODAY = 7
        CURRENT_MONTH = 8
        CURRENT_YEAR = 9
        IS_DOCUMENT_PROFILE = 10
        IS_FILE_FOLDER = 11
        IS_ARCHIVE_BOX = 12
        IS_NOT_NULL = 13
    End Enum

    'Date Types
    Public Enum enumDateTypes
        [NOTHING] = Integer.MinValue
        DATE_ONLY = 1
        DATE_AND_TIME = 2
        TIME_ONLY = 3
        LOCAL_DATE_ONLY = 4
        LOCAL_DATE_AND_TIME = 5
        'REGIONAL = 6
    End Enum

    'Buttons
    Public Enum enumButtons
        cMETHOD_SEARCH = 1
        cMETHOD_ADD = 2
        cMETHOD_MODIFY = 3
        cMETHOD_DELETE = 4
        cMETHOD_PROCESS = 5
        cMETHOD_MOVE = 6
        cMETHOD_SORT = 7
        cMETHOD_REQUEST = 8
        cMETHOD_PRINT = 9
        cMETHOD_VIEW = 10
        cMETHOD_CLONE = 11
        cMETHOD_SELECT = 12
        cMETHOD_SAVEAS = 13
        cMETHOD_IMPORT = 14
        cMETHOD_EXPORT = 15
        cMETHOD_EXTERNALID = 16
        cMETHOD_BARCODE = 17
        cMETHOD_METADATA = 18
        cMETHOD_RANGE = 19
        cFRMLOGIN_LOGIN = 20
        cFRMMAIN_LOGOFF = 21
        cFRMSEARCH_FIND = 22
        cBTNOK = 23
        cBTNBACK = 24
        cBTNCANCEL = 25
        cFRMMASK_UPLOAD = 26
        cBTNAPPLY = 27
        cBTNCALENDAR = 28
        cBTNERASER = 29
        cBTNHIDEPANEL = 30
        cBTNSHOWPANEL = 31
        cBTNHIDETABLE = 32
        cBTNSHOWTABLE = 33
        cBTNLINK = 34
        cBTNNEXTPAGE = 35
        cBTNPREVPAGE = 36
        cBTNPREVOBJ = 37
        cBTNVIEW = 38
        cFRMBACK = 39
        cBTN_CAL_OK = 40
        cBTN_CAL_CANCEL = 41
        cBTN_TYPE_SELECT = 42
        cBTN_MESSAGEBOX_OK = 45
        cBTN_MESSAGEBOX_YES = 46
        cBTN_MESSAGEBOX_NO = 47
    End Enum

    ' Language Filters
    Public Enum enumLanguageFilters
        NoFilter = 1
        Buttons
        ErrorMessages
        Fields
        Tables
        FieldLinks
        FieldColumnHeadings
        TypeDependentFields
        TypeDependentFieldLinks
        LabelsAndMessages
        ApplicationMethods
    End Enum

    '[Naing] Ok this is not cool. Why make up your own TypeIds when there are Types already configured in Recfind 6 Types ???!!!
    Public Enum enumMDPTypeCodes
        None = 0
        FileFolder = 1
        DocumentProfile = 2
        ArchiveBox = 3
    End Enum

    Public Enum enumFormatType
        None = 0                'No Formatting
        FileSize                'Display as Bytes, MB, GB, etc.
        URL                     'Display as a clickable url
        Custom                  'User can specify how to format for phone numbers,etc (see mask text)
        Currency                'User specified and can change currency symbol as well as separating by commas
        [Date]                  'User specified and can change look of dates (view only?)
        YesNo                   'Can change text of a bit field from yes no to something else (TRUE|FALSE)
        Email                   'Display a clickable email address (like a mailto on a web page)
        XML                     'Formats XML data
        CorpVocab               'Allows selection 
        FKeyExtraField          'Allows an extra field to be displayed on foreign key fields (ex. ExternalID + part number)
    End Enum

    Public Enum enumApplicationType As Integer
        RecFindActivationKey = -1
        K1ActivationKey
        K1
        RecFind
        Tacit
        Button
        Scan
        Mini_API
        GEM
        RecCapture
        WebClient
        SharePoint
        API
        OneilIntegration
        Archive
        RF6Connector
    End Enum

    Public Enum enumImageTypes
        Tif = 1
        Gif = 2
        bmp = 3
        Jpg = 4
        Png = 5
        PdfOCR = 6
        PdfaOCR = 7
        Pdf = 8
        Pdfa = 9
    End Enum

    'report formats for exporting
    Public Enum ExportFormat
        HTML
        PDF
        RTF
        Text
        XLS
        TIFF
    End Enum

    Public Enum enumTaskType
        UNKNOWN = 0             'The task was created without a specific type
        FLOW = 1                'Standard task - once complete, moves to next task(s)
        DECISION = 2            'Decision needs to be made about which task happens next
        VIRTUAL_AND = 3         'All precedent tasks need to be completed before kicking off the next task
        VIRTUAL_OR = 4          'Only one of the precedent tasks being completed kicks off the next tasks
        [STOP] = 5
    End Enum

    Public Enum enumSessionTimeoutType
        WHEN_LICENCE_NEEDED = 0
        WHEN_INACTIVE = 1
        WHEN_IN_USE = 2
    End Enum

    Public Enum enumPasswordType
        RECFIND_DATABASE = 0
        ACTIVE_DIRECTORY = 1
        AZURE_ACTIVE_DIRECTORY = 2
    End Enum
#End Region

#Region " Constants "

    Public Const cintNULL As Integer = Integer.MinValue
    Public Const cstrNULL As String = ""
    Public Const cMAX_IDS_BEFORE_TEMPTABLE As Integer = 5000
    Public Const cMAX_COLUMN_HEADINGS As Integer = 8
    Public Const cMIN_DATABASE_VERSION As Double = 11.19

#End Region

#Region " Products "

    Public Class Products

        Public Const cK1 As String = "K1"
        Public Const cRECFIND As String = "RecFind 6"
        Public Const cTACIT As String = "Tacit"
        Public Const cBUTTON As String = "Button"
        Public Const cSCAN As String = "RecScan"
        Public Const cMINI_API As String = "MiniAPI"
        Public Const cGEM As String = "GEM"
        Public Const cRECCAPTURE As String = "RecCapture"
        Public Const cDRM As String = "DRM"
        Public Const cWEBCLIENT As String = "Web Client"
        Public Const cSHAREPOINT As String = "SharePoint Integration Tool"
        Public Const cAPI As String = "RecFind 6 SDK"
        Public Const cONEIL As String = "O'Neil Integration"
        Public Const cARCHIVE As String = "Archive"
        Public Const cRF6CONNECTOR As String = "RF6Connector"
    End Class

#End Region

#Region " System Fields "

    Public Class Fields

#Region " Standard "

        Public Const cID As String = "ID"
        Public Const cTYPEID As String = "TypeID"
        Public Const cEXTERNALID As String = "ExternalID"
        Public Const cSECURITYID As String = "SecurityID"
#End Region

#Region " AccessRightField "

        Public Class AccessRightField
            Public Const cAPPLIESTOTYPEID As String = "AppliesToTypeID"
            Public Const cFIELDID As String = "FieldID"
            Public Const cISVISIBLE As String = "isVisible"
            Public Const cISREADONLY As String = "isReadOnly"
            Public Const cSECURITYGROUPID As String = "SecurityGroupID"
        End Class
#End Region

#Region " AccessRightMethod "

        Public Class AccessRightMethod
            Public Const cAPPLIESTOTYPEID As String = "AppliesToTypeID"
            Public Const cSECURITYGROUPID As String = "SecurityGroupID"
            Public Const cTABLEMETHODID As String = "TableMethodID"
        End Class
#End Region

#Region " ApplicationMethod "

        Public Class ApplicationMethod
            Public Const cAPPLIESTOTYPEID As String = "AppliesToTypeID"
            Public Const cEDOCID As String = "EDOCID"
            Public Const cMETHODTYPE As String = "MethodType"
            Public Const cPARENTAPPMETHODID As String = "ParentApplicationMethodID"
            Public Const cSORTORDER As String = "SortOrder"
            Public Const cSTRINGID As String = "StringID"
            Public Const cTABLEID As String = "TableID"
            Public Const cUIID As String = "UIID"
        End Class
#End Region

#Region " AuditTrail "

        Public Class AuditTrail
            Public Const cCREATEDBYAPPLICATION As String = "CreatedByApplication"
            Public Const [cDATE] As String = "Date"
            Public Const cERRORMESSAGEID As String = "ErrorMessageID"
            Public Const cMETHODID As String = "MethodID"
            Public Const cPERSONID As String = "PersonID"
            Public Const cRECORDEXTERNALID As String = "RecordExternalID"
            Public Const cRECORDID As String = "RecordID"
            Public Const cTABLEID As String = "TableID"
            Public Const cRECORDDATA As String = "RecordData"
            Public Const cPROCESSEDCODE As String = "ProcessedCode"
            Public Const cPROCESSEDDATE As String = "ProcessedDate"
            Public Const cFileName As String = "FileName"
            Public Const cImage As String = "Image"
        End Class
#End Region

#Region " AutoNumberFormat "

        Public Class AutoNumberFormat
            Public Const cFORMAT As String = "Format"
            Public Const cMULTIPLESEQUENCEMASK As String = "MultipleSequenceMask"
            Public Const cSEQUENTIALNUMBERLASTVALUE As String = "SequentialNumberLastValue"
            Public Const cSEQUENTIALNUMBERLENGTH As String = "SequentialNumberLength"
            Public Const cSEQUENTIALNUMBERPAD As String = "SequentialNumberPad"
        End Class
#End Region

#Region " AutoNumberFormatMultipleSequence "

        Public Class AutoNumberFormatMultipleSequence
            Public Const cAUTONUMBERFORMATID As String = "AutoNumberFormatID"
            Public Const cMULTIPLESEQUENCELASTVALUE As String = "MultipleSequenceLastValue"
            Public Const cTITLE1 As String = "Title1"
            Public Const cTITLE2 As String = "Title2"
            Public Const cTITLE3 As String = "Title3"
            Public Const cTITLE4 As String = "Title4"
            Public Const cTITLE5 As String = "Title5"
        End Class
#End Region

#Region " Background "

        Public Class Background
            Public Const cEDOCID As String = "EDOCID"
        End Class
#End Region

#Region " Button "

        Public Class Button
            Public Const cDOWNEDOCID As String = "DownEDOCID"
            Public Const cEDOCID As String = "EDOCID"
            Public Const cOVEREDOCID As String = "OverEDOCID"
            Public Const cSTRINGID As String = "StringID"
            Public Const cUIID As String = "UIID"
        End Class
#End Region

#Region " Capture "

#Region " FiltersDirectories "
        Public Const cFilterID As String = "FilterID"
        Public Const cDirectoryID As String = "DirectoryID"
#End Region

#End Region

#Region " Caption "

        Public Class Caption
            Public Const cSTRINGID As String = "StringID"
            Public Const cUIID As String = "UIID"
        End Class
#End Region

#Region " Codes "

        Public Class Codes
            Public Const cCODE As String = "Code"
        End Class
#End Region

#Region " CorporateVocabulary "

        Public Class CorporateVocabulary
            Public Const cCLASSCODEID As String = "ClassCodeID"
        End Class
#End Region

#Region " DefaultSort "

        Public Class DefaultSort
            Public Const cFIELDID As String = "FieldID"
            Public Const cSORTORDER As String = "SortOrder"
            Public Const cISASCENDING As String = "IsAscending"
        End Class
#End Region

#Region " DocumentType "

        Public Class DocumentType
            Public Const cCODE As String = "Code"
            Public Const cNUMBERPERIODS As String = "NumberPeriods"
            Public Const cPERIODID As String = "PeriodID"
            Public Const cWORKFLOWID As String = "WorkFlowID"
        End Class
#End Region

#Region " DRMFunctions "
        Public Class DRMFunctions
            Public Const cUIID As String = "UIID"
        End Class

#End Region

#Region " DRMMethods "
        Public Class DRMMethods
            Public Const cMethodID As String = "MethodID"
            Public Const cDRMFunctionID As String = "DRMFunctionID"
        End Class
#End Region

#Region " EDOC "

        Public Class EDOC
            Public Const cABSTRACT As String = "Abstract"
            Public Const cAUTHOR As String = "Author"
            Public Const cBCCLIST As String = "BCCList"
            Public Const cCCLIST As String = "CCList"
            Public Const cCHECKEDINPERSONID As String = "CheckedInPersonID"
            Public Const cCHECKEDOUT As String = "CheckedOut"
            Public Const cCONTENTTYPE As String = "ContentType"
            Public Const cCREATEDDATE As String = "CreatedDate"
            Public Const cDOCUMENTTYPEID As String = "DocumentTypeID"
            Public Const cEDOCSTATUSID As String = "EDOCStatusID"
            Public Const cFILENAME As String = "FileName"
            Public Const cIMAGE As String = "Image"
            Public Const cISLATESTVERSION As String = "isLatestVersion"
            Public Const cISYSINDEXED As String = "IsysIndexed"
            Public Const cLASTCHECKOUTPERSONID As String = "LastCheckOutPersonID"
            Public Const cLASTCHECKOUTTIME As String = "LastCheckOutTime"
            Public Const cLIFECYCLEACTUALDATE As String = "LifeCycleActualDate"
            Public Const cMETADATAPROFILEID As String = "MetadataProfileID"
            Public Const cORIGINALCOPY As String = "OriginalCopy"
            Public Const cORIGINALPATH As String = "OriginalPath"
            Public Const cPARENTEDOCID As String = "ParentEDOCID"
            Public Const cPREVIOUSEDOCID As String = "PreviousEDOCID"
            Public Const cPUBLISHEDDATE As String = "PublishedDate"
            Public Const cRECEIVEDDATE As String = "ReceivedDate"
            Public Const cRECIPIENT As String = "Recipient"
            Public Const cSENDER As String = "Sender"
            Public Const cSIZE As String = "Size"
            Public Const cSUBJECT As String = "Subject"
            Public Const cSUFFIX As String = "Suffix"
            Public Const cVERSIONNUMBER As String = "VersionNumber"
            Public Const cCREATEDBYAPPLICATION As String = "CreatedByApplication"
            Public Const cCATEGORY As String = "Category"
            Public Const cCOMMENTS As String = "Comments"
            Public Const cKEYWORDS As String = "KeyWords"
            Public Const cTITLE As String = "Title"
            Public Const cTHUMBNAIL As String = "Thumbnail"
            Public Const cHOVERTHUMBNAIL As String = "HoverThumbnail"
        End Class
#End Region

#Region " ReplaceText "

        Public Class [Error]
            Public Const cERROR As String = "Error"
            Public Const cERRORDATE As String = "ErrorDate"
            Public Const cOCCURREDIN As String = "OccurredIn"
            Public Const cPERSONID As String = "PersonID"
            Public Const cSTACKTRACE As String = "StackTrace"
            Public Const cTABLEID As String = "TableID"
        End Class
#End Region

#Region " ErrorMessage "

        Public Class ErrorMessage
            Public Const cSTRINGID As String = "StringID"
            Public Const cUIID As String = "UIID"
        End Class
#End Region

#Region " ExportFormat "

        Public Class ExportFormat
            Public Const cFORMAT As String = "Format"
            Public Const cPERSONID As String = "PersonID"
            Public Const cTABLEID As String = "TableID"
        End Class
#End Region

#Region " Field "

        Public Class Field
            Public Const cALLOWFREETEXTENTRY As String = "AllowFreeTextEntry"
            Public Const cAUTOFILLTYPE As String = "AutoFillType"
            Public Const cAUTOFILLVALUE As String = "AutoFillValue"
            Public Const cAUTONUMBERFORMATID As String = "AutoNumberFormatID"
            Public Const cCAPTIONID As String = "CaptionID"
            Public Const cDATABASENAME As String = "DatabaseName"
            Public Const cDATATYPE As String = "DataType"
            Public Const cDATETYPE As String = "DateType"
            Public Const cDETERMINESMULTIPLESEQUENCE = "DeterminesMultipleSequence"
            Public Const cFILTERVALUE As String = "FilterValue"
            Public Const cFORMATSTRING As String = "FormatString"
            Public Const cFORMATTYPE As String = "FormatType"
            Public Const cISENCRYPTED As String = "isEncrypted"
            Public Const cISEXPANDED As String = "isExpanded"
            Public Const cISMANDATORY As String = "isMandatory"
            Public Const cISMULTILINE As String = "IsMultiLine"
            Public Const cISMULTIPLESEQUENCEFIELD = "isMultipleSequenceField"
            Public Const cISREADONLY As String = "isReadOnly"
            Public Const cISSYSTEMLOCKED As String = "isSystemLocked"
            Public Const cISSYSTEMESSENTIAL As String = "isSystemEssential"
            Public Const cISSYSTEMNULLABLE As String = "isSystemNullable"
            Public Const cISVISIBLE As String = "isVisible"
            Public Const cISWIDTHPERCENTAGE = "isWidthPercentage"
            Public Const cLENGTH As String = "Length"
            Public Const cNUMBEROFLINES As String = "NumberOfLines"
            Public Const cSCALE As String = "Scale"
            Public Const cSORTORDER As String = "SortOrder"
            Public Const cTABLEID As String = "TableID"
        End Class

#End Region

#Region " FieldLink "

        Public Class FieldLink
            Public Const cCAPTIONID As String = "CaptionID"
            Public Const cDISPLAYASDROPDOWN As String = "DisplayAsDropDown"
            Public Const cFOREIGNFIELDID As String = "ForeignFieldID"
            Public Const cISEXPANDED As String = "isExpanded"
            Public Const cISVISIBLE As String = "IsVisible"
            Public Const cPRIMARYFIELDID As String = "PrimaryFieldID"
            Public Const cSORTORDER As String = "SortOrder"
        End Class
#End Region

#Region " Filter "

        Public Class Filter
            Public Const cAPPLIESTOTYPEID As String = "AppliesToTypeID"
            Public Const cFILTERBYFIELDID As String = "FieldID"
            Public Const cFILTERVALUEFIELDID As String = "FilterByFieldID"
            Public Const cFILTERTYPE As String = "FilterType"
            Public Const cFILTERVALUE As String = "FilterValue"
            Public Const cINITFILTERFIELDID As String = "ForeignKeyFieldID"
            Public Const cINITFILTERFIELDLINKID As String = "InitiateFilterFieldLinkID"
            ''--09/10/2015 -- Neelam -- Added for RecCapture 2.7.1
            Public Const cFILTERSTATUS As String = "FilterStatus"
        End Class
#End Region

#Region " Font "

        Public Class Font
            Public Const cCOLORID As String = "ColorID"
            Public Const cDROPSHADOWCOLORID As String = "DropShadowColorID"
            Public Const cDROPSHADOWPIXELOFFSET As String = "DropShadowPixelOffset"
            Public Const cFONTNAME As String = "FontName"
            Public Const cHASDROPSHADOW As String = "hasDropShadow"
            Public Const cISBOLD As String = "isBold"
            Public Const cISITALIC As String = "isItalic"
            Public Const cSIZE As String = "Size"
        End Class
#End Region

#Region " Forms "

        Public Class Forms
            Public Const cFORMEDOCID As String = "FormEDOCID"
            Public Const cTABLEID As String = "TableID"
        End Class
#End Region

#Region " HelpScreen "

        Public Class HelpScreen
            Public Const cHELPFILE As String = "HelpFile"
            Public Const cLANGUAGEID As String = "LanguageID"
        End Class
#End Region

#Region " Icon "

        Public Class Icon
            Public Const cDOWNEDOCID As String = "DownEDOCID"
            Public Const cEDOCID As String = "EDOCID"
            Public Const cOVEREDOCID As String = "OverEDOCID"
        End Class
#End Region

#Region " K1Configuration "

        Public Class K1Configuration
            Public Const cADMINEMAIL As String = "AdminEmail"
            Public Const cAUDITLOGINS As String = "auditLogins"
            Public Const cAUDITLOGOFFS As String = "auditLogoffs"
            Public Const cAUDITUNSUCCESSFULLOGINS As String = "auditUnsuccessfulLogins"
            Public Const cABSTRACTMAXSENTENCE As String = "AbstractMaxSentences"
            Public Const cCHECKOUTLATESTVERSIONONLY As String = "CheckOutLatestVersionOnly"
            Public Const cDATABASEVERSION As String = "DatabaseVersion"
            Public Const cDEFAULTUSERPROFILEID As String = "DefaultUserProfileID"
            Public Const cDISPLAYIMAGECOLINLIST As String = "DisplayImageColumnInList"
            Public Const cDRMDEFAULTSECURITYID As String = "DRMDefaultSecurityID"
            Public Const cINSTALLDATE As String = "InstallDate"
            Public Const cINSTALLDIR As String = "InstallDir"
            Public Const cISSATURDAYNONWORKINGDAY As String = "isSaturdayNonWorkingDay"
            Public Const cISSUNDAYNONWORKINGDAY As String = "isSundayNonWorkingDay"
            Public Const cMAXCAPTIONWIDTH As String = "MaxCaptionWidth"
            Public Const cMINCAPTIONWIDTH As String = "MinCaptionWidth"
            Public Const cRECORDLOCKTIMEOUT As String = "RecordLockTimeout"
            Public Const cRECORDSETLIMIT As String = "RecordSetLimit"
            Public Const cRECORDSPERPAGE As String = "RecordsPerPage"
            Public Const cSESSIONTIMEOUT As String = "SessionTimeout"
            Public Const cSESSIONTIMEOUTSCANNING As String = "SessionTimeoutScanning"
            Public Const cSOLICITATIONCURRENCYPERIODID As String = "SolicitationCurrencyPeriodID"
            Public Const cSOLICITATIONCURRENCYPERIODUNIT As String = "SolicitationCurrencyPeriodUnit"
            Public Const cTACITCURRENCYPERIODID As String = "TacitCurrencyPeriodID"
            Public Const cTACITCURRENCYPERIODUNIT As String = "TacitCurrencyPeriodUnit"
            Public Const cTACITDEFAULTWEIGHTINGID As String = "TacitDefaultWeightingID"
            Public Const cTACITREFRESHINTERVAL As String = "TacitRefreshInterval"
            Public Const cUSEAUTOMATICLOGINS As String = "UseAutomaticLogins"
            Public Const cWEBSERVICESDISABLED As String = "WebServicesDisabled"
            Public Const cSMTPSERVER As String = "SMTPServer"
            '2015-05-27 -- Peter Melisi -- O'Neil Integration
            Public Const cWEBSERVICESURL As String = "WebServicesURL"
            Public Const cTHUMBNAILSIZE As String = "ThumbnailSize"
            Public Const cWEBSITEURL As String = "WebSiteURL"
            Public Const cSESSIONTIMEOUTTYPE As String = "SessionTimeOutType"
            Public Const cACTIVEDIRECTORYNAME As String = "ActiveDirectoryName"
            Public Const cPASSWORDTYPE As String = "PasswordType"
            '2017-02-17 -- Peter Melisi -- New Licensing Model
            Public Const cWEBSERVICELASTCHECKED As String = "WebServiceLastChecked"
            Public Const cUSEPASSWORDSTRENGTH As String = "UsePasswordStrength"
            Public Const cPASSWORDSTRENGTHSETTINGS As String = "PasswordStrengthSettings"
            '2020-11-24 -- Ara Melkonian -- Azure Active Directory
            Public Const cActiveDirectoryTenant = "ActiveDirectoryTenant"
            Public Const cActiveDirectoryClientId = "ActiveDirectoryClientId"
            Public Const cActiveDirectoryAuthority = "ActiveDirectoryAuthority"
            Public Const cActiveDirectoryWebRedirectUri = "ActiveDirectoryWebRedirectUri"
            Public Const cActiveDirectoryAppRedirectUri = "ActiveDirectoryAppRedirectUri"
            '2021-03-31 -- Ara Melkonian -- DocumentHover
            Public Const cHoverEnabled = "HoverEnabled"
        End Class
#End Region

#Region " K1StartPoint "

        Public Class K1StartPoint
            Public Const cMETHODID As String = "MethodID"
            Public Const cTABLEID As String = "TableID"
        End Class
#End Region

#Region " K1Compliance "

        Public Class K1Compliance
            Public Const cVERVersion As String = "VERSVersion"
        End Class
#End Region

#Region " LanguageString "

        Public Class LanguageString
            Public Const cSTRINGID As String = "StringID"
            Public Const cLANGUAGEID As String = "LanguageID"
            Public Const cSTRING As String = "String"
        End Class
#End Region

#Region " LegalHold "

        Public Class LegalHold
            Public Const cLEGALHOLD As String = "LegalHold"
            Public Const cISACTIVE As String = "isActive"
            Public Const cEXTERNALID As String = "ExternalID"
            Public Const cDATECREATED As String = "DateCreated"
            Public Const cPERSONRELEASING As String = "PersonReleasingID"
            Public Const cPERSONCREATING As String = "CreatedBy"
            Public Const cDATERELEASED As String = "DateReleased"
            Public Const cACTUALDURATION As String = "ActualDuration"
            Public Const cEDOCID As String = "EDOCID"
            Public Const cMETADATAID As String = "MetadataProfileID"
            Public Const cLEGALHOLDID As String = "LegalHoldID"
            Public Const cAUTHORITY As String = "Authority"
            Public Const cAUTHORIZED As String = "AuthorizedPersonID"
        End Class
#End Region

#Region " LinkMailListPerson "

        Public Class LinkMailListPerson
            Public Const cMAILLISTID As String = "MailListID"
            Public Const cPERSONID As String = "PersonID"
        End Class
#End Region

#Region " LinkSavedReportSavedReport "

        Public Class LinkSavedReportSavedReport
            Public Const cPARENTREPORTID As String = "ParentReportID"
            Public Const cSUBREPORTID As String = "SubReportID"
        End Class
#End Region

#Region " LinkSecurityGroupTable "

        Public Class LinkSecurityGroupAppMethod
            Public Const cSECURITYGROUPID As String = "SecurityGroupID"
            Public Const cAPPMETHODID As String = "ApplicationMethodID"
        End Class
#End Region

#Region " Link UserProfileDRMMethod "
        Public Class LinkUserProfileDRMMethod
            Public Const CUSERPROFILEID As String = "UserProfileID"
            Public Const cDRMMETHODID As String = "DRMMethodID"
        End Class
#End Region

#Region " Link UserProfileDRMFunction "
        Public Class LinkUserProfileDRMFunction
            Public Const CDRMFUNCTIONID As String = "DRMFunctionID"
            Public Const CUSERPROFILEID As String = "UserProfileID"
        End Class

#End Region

#Region " LinkUserProfileSecurityGroup "
        Public Class LinkUserProfileSecurityGroup
            Public Const cUSERPROFILEID As String = "UserProfileID"
            Public Const cSECURITYGROUPID As String = "SecurityGroupID"
        End Class

#End Region

#Region " LinkSecurityGroupField "

        Public Class LinkSecurityGroupField
            Public Const cFIELDID As String = "FieldID"
            Public Const cSECURITYGROUPID As String = "SecurityGroupID"
        End Class
#End Region

#Region " LinkSecurityGroupReadOnlyField "
        Public Class LinkSecurityGroupReadOnlyField
            Public Const cFIELDID As String = "FieldID"
            Public Const cSECURITYGROUPID As String = "SecurityGroupID"
        End Class
#End Region

#Region " LinkSecurityGroupSecurity "

        Public Class LinkSecurityGroupSecurity
            Public Const cSECURITYGROUPID As String = "SecurityGroupID"
            Public Const cSECURITYID As String = "SecurityID"
        End Class
#End Region

#Region " LinkSecurityGroupTable "

        Public Class LinkSecurityGroupTable
            Public Const cSECURITYGROUPID As String = "SecurityGroupID"
            Public Const cTABLEID As String = "TableID"
        End Class
#End Region

#Region " LinkSecurityGroupTableMethod "

        Public Class LinkSecurityGroupTableMethod
            Public Const cSECURITYGROUPID As String = "SecurityGroupID"
            Public Const cTABLEMETHODID As String = "TableMethodID"
        End Class
#End Region

#Region " LinkSolicitationCorporateVocabulary "

        Public Class LinkSolicitationCorporateVocabulary
            Public Const cCORPORATEVOCABULARYID As String = "CorporateVocabularyID"
            Public Const cSOLICITATIONID As String = "SolicitationID"
        End Class
#End Region

#Region " LinkSolicitationMailList "

        Public Class LinkSolicitationMailList
            Public Const cMAILLISTID As String = "MailListID"
            Public Const cSOLICITATIONID As String = "SolicitationID"
        End Class
#End Region

#Region " LinkTacitCorporateVocabulary "

        Public Class LinkTacitCorporateVocabulary
            Public Const cCORPORATEVOCABULARYID As String = "CorporateVocabularyID"
            Public Const cTACITID As String = "TacitID"
        End Class
#End Region

#Region " LinkTacitEntity "

        Public Class LinkTacitEntity
            Public Const cENTITYID As String = "EntityID"
            Public Const cTACITID As String = "TacitID"
        End Class
#End Region

#Region " LinkTacitPerson "

        Public Class LinkTacitPerson
            Public Const cPERSONID As String = "PersonID"
            Public Const cTACITID As String = "TacitID"
        End Class
#End Region

#Region " LinkTacitProductService "

        Public Class LinkTacitProductService
            Public Const cPRODUCTSERVICEID As String = "ProductServiceID"
            Public Const cTACITID As String = "TacitID"
        End Class
#End Region

#Region " LinkTacitWWW "

        Public Class LinkTacitWWW
            Public Const cTACITID As String = "TacitID"
            Public Const cWWWID As String = "WWWID"
        End Class
#End Region

#Region " LinkTaskTask "

        Public Class LinkTaskTask
            Public Const cPREVTASKID As String = "PrevTaskID"
            Public Const cNEXTTASKID As String = "NextTaskID"
        End Class
#End Region

#Region " ListColumn "

        Public Class ListColumn
            Public Const cAPPLIESTOTYPEID As String = "AppliesToTypeID"
            Public Const cCAPTIONID As String = "CaptionID"
            Public Const cFIELDID As String = "FieldID"
            Public Const cSORTORDER As String = "SortOrder"
            Public Const cWIDTH As String = "Width"
        End Class
#End Region

#Region " Location "

        Public Class Location
            Public Const cBARCODE As String = "Barcode"
            Public Const cENTITYID As String = "EntityID"
            Public Const cSTREETADDRESS1 As String = "StreetAddress1"
            Public Const cSTREETADDRESS2 As String = "StreetAddress2"
            Public Const cSTREETADDRESS3 As String = "StreetAddress3"
            Public Const cSTREETCITYID As String = "StreetCityID"
            Public Const cSTREETSTATEID As String = "StreetStateID"
            Public Const cSTREETPOSTCODE As String = "StreetPostcode"
        End Class

#End Region

#Region " MetadataProfile "

        Public Class MetadataProfile
            Public Const cBARCODE As String = "Barcode"
            Public Const cCLOSEDDATE As String = "ClosedDate"
            Public Const cCREATEDDATE As String = "CreatedDate"
            Public Const cLIFECYCLE1ACTUALDATE As String = "LifeCycle1ActualDate"
            Public Const cLIFECYCLE2ACTUALDATE As String = "LifeCycle2ActualDate"
            Public Const cPARTNUMBER As String = "PartNumber"
            Public Const cRECCATEGORYID As String = "RecordCategoryID"
            Public Const cPERMANENTRECORD As String = "PermanentRecord"
            Public Const cRETENTIONCODE1 As String = "RetentionCodeID"
            Public Const cRETENTIONCODE2 As String = "RetentionCode2ID"
            Public Const cRECORDSFROMDATE As String = "RecordsFromDate"
            Public Const cRECORDSTODATE As String = "RecordsToDate"
            Public Const cSTATUSID As String = "StatusID"
            Public Const cTITLE1 As String = "Title1"
            Public Const cTITLE2 As String = "Title2"
            Public Const cTITLE3 As String = "Title3"
            Public Const cTITLE4 As String = "Title4"
            Public Const cTITLE5 As String = "Title5"
            Public Const cVITALRECORDID As String = "VitalRecordID"
            Public Const cVITALRECORDLASTREVIEWDATE As String = "VitalRecordLastReviewDate"
            Public Const cVITALRECORDNEXTREVIEWDATE As String = "VitalRecordNextReviewDate"
            Public Const cSUBJECT As String = "Subject"
            Public Const cDESCRIPTION As String = "Description"
            Public Const cSERIALNUMBER As String = "SerialNumber"
            Public Const cAUTHORLIST As String = "AuthorList"
            Public Const cRECIPIENTLIST As String = "RecipientList"
            Public Const cCONSIGNMENTNUMBER As String = "ConsignmentNumber"
            Public Const cOLDNUMBER As String = "OldNumber"
            Public Const cTRAPREASON As String = "TrapReason"
            Public Const cCONTENTS As String = "Contents"
            Public Const cRECORDSCENTRENUMBER As String = "RecordsCentreNumber"
            Public Const cCURRENTLOCATION As String = "CurrentLocation"
            Public Const cABSTRACT As String = "Abstract"
            Public Const cSENDER As String = "Sender"
            Public Const cALLOCATEDSPACEID As String = "AllocatedSpaceID"
        End Class
#End Region

#Region " Method "

        Public Class Method
            Public Const cBUTTONID As String = "ButtonID"
            Public Const cUIID As String = "UIID"
        End Class
#End Region

#Region " Movement "

        Public Class Movement
            Public Const cFROMBARCODEREADER As String = "FromBarcodeReader"
            Public Const cISLATESTMOVEMENT As String = "isLatestMovement"
            Public Const cMETADATAPROFILEID As String = "MetadataProfileID"
            Public Const cMOVEDDATE As String = "MovedDate"
            Public Const cMOVERPERSONID As String = "MoverPersonID"
            Public Const cNEWDEPARTMENTDIVISIONID As String = "NewDepartmentDivisionID"
            Public Const cNEWENTITYID As String = "NewEntityID"
            Public Const cNEWLOCATIONID As String = "NewLocationID"
            Public Const cNEWMETADATAPROFILEID As String = "NewMetadataProfileID"
            Public Const cNEWSPACEID As String = "NewSpaceID"
            Public Const cRECIPIENTPERSONID As String = "RecipientPersonID"
            Public Const cKEEPSPACEALLOCATED As String = "KeepSpaceAllocated"
        End Class
#End Region

#Region " Period "

        Public Class Period
            Public Const cCODE As String = "Code"
            Public Const cDESCRIPTION As String = "Description"
        End Class
#End Region

#Region " Person "

        Public Class Person
            Public Const cFIRSTNAME As String = "FirstName"
            Public Const cLASTNAME As String = "LastName"
            Public Const cENTITYID As String = "EntityID"
            Public Const cLOCATIONID As String = "LocationID"
            Public Const cWORKPHONE As String = "WorkPhone"
            Public Const cWORKFAX As String = "WorkFax"
            Public Const cWORKEMAIL As String = "WorkEmail"
            Public Const cHOMEEMAIL As String = "HomeEmail"
            Public Const cBARCODE As String = "Barcode"
            Public Const cISACTIVE As String = "isActive"
        End Class

#End Region

#Region " Process "

        Public Class Process
            Public Const cINPUTTYPE As String = "InputType"
            Public Const cPERSONID As String = "PersonID"
            Public Const cPROCESS As String = "Process"
            Public Const cPROCESSFILEID As String = "ProcessFileID"
            Public Const cPROCESSTYPE As String = "ProcessType"
            Public Const cREQUIRESINPUT As String = "RequiresInput"
            Public Const cTABLEID As String = "TableID"
        End Class
#End Region

#Region " RecordCategory "

        Public Class RecordCategory
            Public Const cRECALCULATE As String = "Recalculate"
            Public Const cRETENTION1ID As String = "RetentionCodeID"
            Public Const cRETENTION2ID As String = "RetentionCode2ID"
        End Class
#End Region

#Region " Retention "

        Public Class RetentionCode
            Public Const cCITATIONRETENTIONCODEID As String = "CitationRetentionCodeID"
            Public Const cRECALCULATE As String = "Recalculate"
        End Class
#End Region

#Region " Request "

        Public Class Request
            Public Const cMETADATAPROFILEID As String = "MetadataProfileID"
            Public Const cMOVEMENTSATISFIEDBYID As String = "MovementSatisfiedByID"
            Public Const cNEWDEPARTMENTDIVISIONID As String = "NewDepartmentDivisionID"
            Public Const cNEWENTITYID As String = "NewEntityID"
            Public Const cNEWLOCATIONID As String = "RequestorLocationID"
            Public Const cNEWMETADATAPROFILEID As String = "NewMetadataProfileID"
            Public Const cNEWSPACEID As String = "NewSpaceID"
            Public Const cNEWPERSONID As String = "NewPersonID"
            Public Const cREQUESTORPERSONID As String = "RequestorPersonID"
            Public Const cREQUIREDDATE As String = "RequiredDate"
            Public Const cSATISFIEDDATE As String = "SatisfiedDate"
        End Class
#End Region

#Region " SavedReport "

        Public Class SavedReport
            Public Const cCREATEDDATE As String = "CreatedDate"
            Public Const cCREATORPERSONID As String = "CreatorPersonID"
            Public Const cEDOCID As String = "EDOCID"
            Public Const cTABLEID As String = "TableID"
            Public Const cMDPTYPE As String = "MDPType"
        End Class
#End Region

#Region " SavedSearch "

        Public Class SavedSearch
            Public Const cCREATEDDATE As String = "CreatedDate"
            Public Const cCREATORPERSONID As String = "CreatorPersonID"
            Public Const cSQL As String = "SQL"
            Public Const cTABLEID As String = "TableID"
        End Class
#End Region

#Region " ScheduledTask "

        Public Class ScheduledTask
            Public Const cBEGINDATE As String = "BeginDate"
            Public Const cENDDATE As String = "EndDate"
            Public Const cISACTIVE As String = "IsActive"
            Public Const cISHTML As String = "IsHTML"
            Public Const cLASTRUNDATE As String = "LastRunDate"
            Public Const cMAILLISTID As String = "MailListID"
            Public Const cPERIODID As String = "PeriodID"
            Public Const cRETURNSMAILLIST As String = "ReturnsMaillist"
            Public Const cSAVEDREPORTFORMAT As String = "SavedReportFormat"
            Public Const cSAVEDREPORTID As String = "SavedReportID"
            Public Const cSAVEDSEARCHID As String = "SavedSearchID"
            Public Const cSENDEREMAIL As String = "SenderEmail"
            Public Const cSENDIFEMPTY As String = "SendIfEmpty"
            Public Const cSMTPSERVER As String = "SMTPServer"
            Public Const cSQL As String = "SQL"
            Public Const cSTOREDPROCEDURENAME As String = "StoredProcedureName"
            Public Const cTIMEDUE As String = "TimeDue"
        End Class
#End Region

#Region " Space "

        Public Class Space
            Public Const cENTITYID As String = "EntityID"
            Public Const cSPACENUMBER As String = "SpaceNumber"
            Public Const cBARCODE As String = "Barcode"
            Public Const cBOXTYPEID As String = "BoxTypeID"
            Public Const cFREECAPACITY As String = "FreeCapcity"
            Public Const cFLOOR As String = "Floor"
            Public Const cROOM As String = "Room"
            Public Const cROW As String = "Row"
            Public Const cBAY As String = "Bay"
            Public Const cSHELF As String = "Shelf"
            Public Const cSLOT As String = "Slot"
            Public Const cIGNOREBOXCAPACITY As String = "IgnoreBoxCapacity"
            Public Const cIGNOREMISMATCHBOXTYPE As String = "IgnoreMismatchBoxType"
            Public Const cCLASSIFICATIONID As String = "ClassificationID"
            Public Const cSPACETYPEID As String = "SpaceTypeID"
            Public Const cCAPACITY As String = "Capacity"
        End Class

#End Region

#Region " Security "
        Public Class Security
            Public Const cISPUBLIC As String = "IsPublic"
        End Class
#End Region

#Region " SecurityGroup "

        '2016-08-02 -- Peter Melisi - Bug fixed for #1600003149
        'Public Class SecurityGroup
        '    Public Const cNOMINATEDSECURITYID As String = "NominatedSecurityID"
        'End Class
#End Region

#Region " Solicitation "

        Public Class Solicitation
            Public Const cCREATEDDATE As String = "CreatedDate"
            Public Const cCREATORPERSONID As String = "CreatorPersonID"
            Public Const cCURRENCYDATE As String = "CurrencyDate"
            Public Const cDETAILS As String = "Details"
            Public Const cISDRAFT As String = "isDraft"
            Public Const cKEYWORDS As String = "Keywords"
            Public Const cVISIBLE As String = "Visible"
        End Class
#End Region

#Region " StoredProcedure "

        Public Class StoredProcedure
            Public Const cDATABASENAME As String = "DatabaseName"
            Public Const cSQL As String = "SQL"
            Public Const cTABLEID As String = "TableID"
        End Class
#End Region

#Region " StoredProcedure "

        Public Class [String]
            Public Const cCAPTIONID As String = "CaptionID"
        End Class

#End Region

#Region " Table "

        Public Class Table
            Public Const cCAPTIONID As String = "CaptionID"
            Public Const cCLASS As String = "Class"
            Public Const cDATABASENAME As String = "DatabaseName"
            Public Const cICONID As String = "IconID"
            Public Const cISTYPEDEPENDENT As String = "isTypeDependent"
            Public Const cSHOWICON As String = "ShowIcon"
        End Class
#End Region

#Region " TableMethod "

        Public Class TableMethod
            Public Const cAUDIT As String = "Audit"
            Public Const cAUDITDATA As String = "AuditData"
            Public Const cMETHODID As String = "MethodID"
            Public Const cTABLEID As String = "TableID"
        End Class
#End Region

#Region " Tacit "

        Public Class Tacit
            Public Const cCREATEDDATE As String = "CreatedDate"
            Public Const cCREATORPERSONID As String = "CreatorPersonID"
            Public Const cCURRENCYDATE As String = "CurrencyDate"
            Public Const cDETAILS As String = "Details"
            Public Const cISDRAFT As String = "isDraft"
            Public Const cKEYWORDS As String = "Keywords"
            Public Const cSOLICITATIONID As String = "SolicitationID"
            Public Const cVISIBLE As String = "Visible"
            Public Const cWEIGHTINGID As String = "WeightingID"
        End Class
#End Region

#Region " Task "

        Public Class Task
            Public Const cWORKFLOWID As String = "WorkFlowID"
            Public Const cSCHEDULEDCOMPLETEDATE As String = "ScheduleCompleteDate"
            Public Const cACTUALCOMPLETEDATE As String = "ActualCompleteDate"
            Public Const cWORKFLOWPROGRESS As String = "Decision"
            Public Const cPRECEDENTTASKID As String = "PrecedentTaskID"
            Public Const cDESCRIPTION As String = "Description"
            Public Const cDELEGATOR As String = "Delegator"
            Public Const cDELEGATE As String = "Delegate"
            Public Const cBUDGETCOMPLETEHOURS As String = "BudgetCompleteHours"
            Public Const cBUDGETCOMPLETECOST As String = "BudgetCompleteCost"
            Public Const cCOMMENTS As String = "Comments"
            Public Const cWORKINGDAYSTOCOMPLETE As String = "WorkingDaysToComplete"
            Public Const cEMAILWHENOVERDUE As String = "EmailWhenOverDue"
            Public Const cPERSONTONOTIFYWHENDUE As String = "PersonToNotifyWhenDue"
            Public Const cPERSONTONOTIFYWHENOVERDUE As String = "PersonToNotifyWhenOverdue"
            Public Const cDAYSNOTICE As String = "DaysNotice"
            Public Const cDAYSGRACE As String = "DaysGrace"
            Public Const cSTOPADVISEPERSONID As String = "StopAdvisePersonID"
            Public Const cISSTARTINGTASK As String = "IsStartingTask"
            Public Const cUIID As String = "UIID"
            Public Const cCREATEDFROMTEMPLATETASKID As String = "CreatedFromTemplateTaskID"
            Public Const cSTATUSID As String = "StatusID"
            Public Const cACTUALCOMPLETEHOURS As String = "ActualCompleteHours"
            Public Const cACTUALCOMPLETECOST As String = "ActualCompleteCost"
            Public Const cTABLEID As String = "TableID"
            Public Const cTASKLINKID As String = "TaskLinkID"
            Public Const cSUSPENDREASON As String = "SuspendReason"
        End Class

#End Region

#Region " TaskLink "

        Public Class TaskLink
            Public Const cSOURCETASKID As String = "SourceTaskID"
            Public Const cDESTINATIONTASKID As String = "DestinationTaskID"
        End Class

#End Region

#Region " Title "

        Public Class Title
            Public Const cLEVEL As String = "Level"
            Public Const cRECALCULATE As String = "Recalculate"
            Public Const cRETENTION1ID As String = "RetentionCodeID"
            Public Const cRETENTION2ID As String = "RetentionCode2ID"
            Public Const cCODE As String = "Code"
            Public Const cPARENTTITLE As String = "ParentTitle"
        End Class
#End Region

#Region " TKQueue "

        Public Class TKQueue
            Public Const cREAD As String = "Read"
            Public Const cRECIPIENTPERSONID As String = "RecipientPersonID"
            Public Const cREPLIED As String = "Replied"
            Public Const cSOLICITATIONID As String = "SolicitationID"
            Public Const cVISIBLE As String = "Visible"
        End Class
#End Region

#Region " Trigger "

        Public Class Trigger
            Public Const cDATABASENAME As String = "DatabaseName"
            Public Const cONDELETE As String = "OnDelete"
            Public Const cONINSERT As String = "OnInsert"
            Public Const cONUPDATE As String = "OnUpdate"
            Public Const cSQL As String = "SQL"
            Public Const cTABLEID As String = "TableID"
            Public Const cTRIGGERACTION As String = "TriggerAction"
        End Class
#End Region

#Region " TriggerColumn "

        Public Class TriggerColumn
            Public Const cFIELDID As String = "FieldID"
            Public Const cTRIGGERID As String = "TriggerID"
        End Class
#End Region

#Region " Type "

        Public Class Type
            Public Const cTABLEID As String = "TableID"
        End Class
#End Region

#Region " TypeFieldInfo "

        Public Class TypeFieldInfo
            Public Const cALLOWFREETEXTENTRY As String = "AllowFreeTextEntry"
            Public Const cAPPLIESTOTYPEID As String = "AppliesToTypeID"
            Public Const cAUTOFILLTYPE As String = "AutoFillType"
            Public Const cAUTOFILLVALUE As String = "AutoFillValue"
            Public Const cAUTONUMBERFORMATID As String = "AutoNumberFormatID"
            Public Const cCAPTIONID As String = "CaptionID"
            Public Const cDETERMINESMULTIPLESEQUENCE = "DeterminesMultipleSequence"
            Public Const cFIELDID As String = "FieldID"
            Public Const cISMANDATORY As String = "isMandatory"
            Public Const cISMULTIPLESEQUENCEFIELD = "isMultipleSequenceField"
            Public Const cISREADONLY As String = "isReadOnly"
            Public Const cISVISIBLE As String = "isVisible"
            Public Const cSORTORDER As String = "SortOrder"
        End Class
#End Region

#Region " TypeFieldLinkInfo "

        Public Class TypeFieldLinkInfo
            Public Const cAPPLIESTOTYPEID As String = "AppliesToTypeID"
            Public Const cCAPTIONID As String = "CaptionID"
            Public Const cFIELDLINKID As String = "FieldLinkID"
            Public Const cISVISIBLE As String = "isVisible"
            Public Const cSORTORDER As String = "SortOrder"
        End Class
#End Region

#Region " UserProfile "

        Public Class UserProfile
            Public Const cBACKGROUNDID As String = "BackgroundID"
            'Public Const cBARCODE As String = "Barcode"
            Public Const cBUTTONFONTID As String = "ButtonFontID"
            Public Const cDISPLAYTABLELIST As String = "DisplayTableList"
            Public Const cFONTID As String = "FontID"
            Public Const cHEADINGFONTID As String = "HeadingFontID"
            Public Const cISADMINISTRATOR As String = "IsAdministrator"
            Public Const cISSUPERADMINISTRATOR As String = "IsSuperAdministrator"
            Public Const cISCAPTUREADMIN As String = "isCaptureAdmin"
            Public Const cISGEMADMIN As String = "isGEMAdmin"
            Public Const cK1STARTPOINTMETHODID As String = "K1StartPointMethodID"
            Public Const cK1STARTPOINTTABLEID As String = "K1StartPointTableID"
            Public Const cLANGUAGEID As String = "LanguageID"
            Public Const cNOMINATEDSECURITYID As String = "NominatedSecurityID"
            Public Const cISPUBLIC As String = "IsPublic"
            Public Const cLASTSIGNOFF As String = "LastSignOff"
            Public Const cLASTSIGNON As String = "LastSignOn"
            Public Const cPASSWORD As String = "Password"
            Public Const cPERSONID As String = "PersonID"
            Public Const cRECIEVESOLICITATIONS As String = "RecieveSolicitations"
            Public Const cSECURITYGROUPID As String = "SecurityGroupID"
            Public Const cSETTINGS As String = "Settings"
            Public Const cUSERID As String = "UserID"
            Public Const cSHAREPOINTUSER As String = "SharePointUser"
            Public Const cTIMEZONE As String = "Timezone"
            Public Const cBIREPORTID As String = "BIReportID"
            Public Const cCONNECTORUSERKEY As String = "ConnectorUSerKey"
        End Class
#End Region

#Region " Warning "

        Public Class Warning
            Public Const cSQL As String = "SQL"
            Public Const cSTOREDPROCEDURENAME As String = "StoredProcedureName"
            Public Const cTABLEID As String = "TableID"
            Public Const cWARNINGTYPE As String = "WarningType"
        End Class
#End Region

#Region " WorkFlow "

        Public Class WorkFlow
            Public Const cWORKINGDAYSTOCOMPLETE As String = "WorkingDaysToComplete"
            Public Const cDESCRIPTION As String = "Description"
            Public Const cBUDGETCOMPLETEHOURS As String = "BudgetCompleteHours"
            Public Const cBUDGETCOMPLETECOST As String = "BudgetCompleteCost"
            Public Const cENTITYID As String = "EntityID"
            Public Const cPERSONID As String = "PersonID"
            Public Const cISADHOC As String = "IsAdHoc"
            Public Const cISSTANDARDREPLY As String = "IsStandardReply"
            Public Const cISTEMPLATE As String = "IsTemplate"
            Public Const cSAVEDTEMPLATE As String = "SavedTemplate"
            Public Const cTEMPLATEWORKFLOWID As String = "TemplateWorkFlowID"
            Public Const cCREATORPERSONID As String = "CreatorPersonID"
            Public Const cSTATUSID As String = "StatusID"
        End Class
#End Region

#Region " K1 Licence File "

        Public Class K1LicenceFile
            Public Const cAPPLICATION_TYPE As String = "ApplicationType"
            Public Const CLICENCEFILE As String = "LicenseFile"
        End Class

#End Region

#Region " K1 Record Lock "

        Public Class K1RecordLock
            Public Const cTABLEID As String = "TableID"
            Public Const cRECORDID As String = "RecordID"
            Public Const cUSERPROFILEID As String = "UserProfileID"
            Public Const cTIMELOCKED As String = "TimeLocked"
            Public Const cUSERNAME As String = "UserName"
        End Class
#End Region

#Region " K1 System Flags "

        Public Class K1SystemFlags
            Public Const cDRM_RUNNING As String = "DRM_Running"
            Public Const cDRM_USERPROFILEID As String = "DRM_UserProfileID"
            Public Const cDRM_CHECK_DATE As String = "DRM_Check_Date"
            Public Const cDRM_REFRESH_DATE As String = "DRM_Refresh_Date"
            Public Const cDRM_FORCED_LOGOFF_DATE As String = "DRM_Forced_Logoff_Date"
            Public Const cDRM_LOCKED As String = "DRM_Locked"
        End Class
#End Region

#Region " K1 Session "

        Public Class K1Session
            Public Const cAPPLICATION_TYPE As String = "ApplicationType"
            Public Const cUSERPROFILEID As String = "UserProfileID"
            Public Const cLAST_UPDATED As String = "LastUpdated"
        End Class
#End Region

#Region " K1 MDP Types "

        Public Class K1MDPTypes
            Public Const cTYPECODE As String = "TypeCode"
        End Class
#End Region

#Region " K1 Pick List "

        Public Class K1PickList
            Public Const cTABLEID As String = "TableID"
            Public Const cRECORDID As String = "RecordID"
            Public Const cUSERID As String = "UserID"
        End Class
#End Region

#Region " K1 Groups "

        Public Class K1Groups
            Public Const cTYPECODE As String = "TypeCode"
            Public Const cDEFAULTTYPEID As String = "DefaultTypeID"
        End Class
#End Region

#Region " K1 Archive "

        Public Class K1Archive
            Public Const cCONNECTIONSTRING As String = "ExternalID"
            Public Const cCURRENT As String = "Current"
        End Class
#End Region

    End Class
#End Region

#Region " System Tables "

    Public Class Tables

#Region " Standard "

        Public Const cACCESSRIGHTFIELD As String = "AccessRightField"
        Public Const cACCESSRIGHTMETHOD As String = "AccessRightMethod"
        Public Const cAPPLICATIONMETHOD As String = "ApplicationMethod"
        Public Const cAUDITTRAIL As String = "AuditTrail"
        Public Const cAUTONUMBERFORMAT As String = "AutoNumberFormat"
        Public Const cAUTONUMBERFORMATMULTIPLESEQUENCE As String = "AutoNumberFormatMultipleSequence"
        Public Const cBACKGROUND As String = "Background"
        Public Const cBOXTYPE As String = "BoxType"
        Public Const cBUTTON As String = "Button"
        Public Const cCAPTION As String = "Caption"
        Public Const cCITY As String = "City"
        Public Const cCLOSETYPE As String = "CloseType"
        Public Const cCODES As String = "Codes"
        Public Const cCOLOR As String = "Color"
        Public Const cCORPORATEVOCABULARY As String = "CorporateVocabulary"
        Public Const cDEFAULTSORT As String = "DefaultSort"
        Public Const cDEPARTMENTDIVISION As String = "DepartmentDivision"
        Public Const cDOCUMENTTYPE As String = "DocumentType"
        Public Const cDRMMethod As String = "DRMMethod"
        Public Const cDRMFunction As String = "DRMFunction"
        Public Const cEDOC As String = "EDOC"
        Public Const cEMAILNOTIFICATION As String = "EmailNotification"
        Public Const cENTITY As String = "Entity"
        Public Const cERROR As String = "Error"
        Public Const cERRORMESSAGE As String = "ErrorMessage"
        Public Const cFIELD As String = "Field"
        Public Const cFIELDLINK As String = "FieldLink"
        Public Const cFILTER As String = "Filter"
        Public Const cFONT As String = "Font"
        Public Const cHELPSCREEN As String = "HelpScreen"
        Public Const cICON As String = "Icon"
        Public Const cINCIDENT As String = "Incident"
        Public Const cK1CONFIGURATION As String = "K1Configuration"
        Public Const cLANGUAGE As String = "Language"
        Public Const cLANGUAGESTRING As String = "LanguageString"
        Public Const cLEGALHOLD As String = "LegalHold"
        Public Const cLICENCEFILE As String = "LicenseFile"
        Public Const cLISTCOLUMN As String = "ListColumn"
        Public Const cLOCATION As String = "Location"
        Public Const cMAILLIST As String = "MailList"
        Public Const cMEDIATYPE As String = "MediaType"
        Public Const cMETADATAPROFILE As String = "MetadataProfile"
        Public Const cMETHOD As String = "Method"
        Public Const cMOVEMENT As String = "Movement"
        Public Const cPERIOD As String = "Period"
        Public Const cPERSON As String = "Person"
        Public Const cPROCESS As String = "Process"
        Public Const cRACTION As String = "RAction"
        Public Const cRECORDCATEGORY As String = "RecordCategory"
        Public Const cREQUEST As String = "Request"
        Public Const cRETENTIONCODE As String = "RetentionCode"
        Public Const cSAVEDREPORT As String = "SavedReport"
        Public Const cSAVEDSEARCH As String = "SavedSearch"
        Public Const cSCHEDULEDTASK As String = "ScheduledTask"
        Public Const cSECURITY As String = "Security"
        Public Const cSECURITYGROUP As String = "SecurityGroup"
        Public Const cSERIES As String = "Series"
        Public Const cSOLICITATION As String = "Solicitation"
        Public Const cSPACE As String = "Space"
        Public Const cSTATE As String = "State"
        Public Const cSTRING As String = "String"
        Public Const cSUPPLEMENTALMARKUP As String = "SupplementalMarkup"
        Public Const cTABLE As String = "Table"
        Public Const cTABLEMETHOD As String = "TableMethod"
        Public Const cTACIT As String = "Tacit"
        Public Const cTASK As String = "Task"
        Public Const cTASKLINK As String = "TaskLink"
        Public Const cTITLE As String = "Title"
        Public Const cTKQUEUE As String = "TKQueue"
        Public Const cTRIGGER As String = "Trigger"
        Public Const [cTYPE] As String = "Type"
        Public Const cTYPEFIELDINFO As String = "TypeFieldInfo"
        Public Const cTYPEFIELDLINKINFO As String = "TypeFieldLinkInfo"
        Public Const cUSERPROFILE As String = "UserProfile"
        Public Const cVITALRECORD As String = "VitalRecord"
        Public Const cWARNING As String = "Warning"
        Public Const cWEIGHTING As String = "Weighting"
        Public Const cWORKFLOW As String = "WorkFlow"
        Public Const cWORKFLOWSUSPENSION As String = "WorkFlowSuspension"
#End Region

#Region " Link Tables "

        Public Const cLINKLEGALHOLDEDOC As String = "LinkLegalHoldEDOC"
        Public Const cLINKLEGALHOLDMETADATAPROFILE As String = "LinkLegalHoldMetadataProfile"
        Public Const cLINKMAILLISTPERSON As String = "LinkMailListPerson"
        Public Const cLINKMETADATAPROFILEWORFLOW As String = "LinkMetadataProfileWorkFlow"
        Public Const cLINKEDOCWORFLOW As String = "LinkEDOCWorkFlow"
        Public Const cLINKSECURITYGROUPAPPMETHOD As String = "LinkSecurityGroupApplicationMethod"
        Public Const cLINKSECURITYGROUPTABLE As String = "LinkSecurityGroupTable"
        Public Const cLINKSECURITYGROUPTABLEMETHOD As String = "LinkSecurityGroupTableMethod"
        Public Const cLINKSECURITYGROUPFIELD As String = "LinkSecurityGroupField"
        Public Const cLINKSECURITYGROUPREADONLYFIELDS = "LinkSecurityGroupReadOnlyFields"
        Public Const cLINKSECURITYGROUPSECURITY As String = "LinkSecurityGroupSecurity"
        Public Const cLINKUSERPROFILEDRMMETHOD As String = "LinkUserProfileDRMMethod"
        Public Const cLINKUSERPROFILEDRMFUNCTION As String = "LinkUserProfileDRMFunction"
        Public Const cLINKUSERPROFILESECURITYGROUP = "LinkUserProfileSecurityGroup"
        Public Const cLINKSAVEDREPORTSAVEDREPORT As String = "LinkSavedReportSavedReport"
        Public Const cLINKTODOPEOPLE As String = "LinkToDoPeople"
        Public Const cLINKTASKTASK As String = "LinkTaskTask"
#End Region

#Region " Xchange "

        Public Const cXCHANGEIMPORT As String = "XChangeImport"
        Public Const cXCHANGEEXPORT As String = "XChangeExport"
#End Region

#Region " Capture "

        Public Const cCAPTUREABSTRACT As String = "Capture_Abstract"
        Public Const cCAPTUREADMINEMAILS As String = "Capture_AdminEmails"
        Public Const cCAPTUREAGENTS As String = "Capture_Agents"
        Public Const cCAPTUREDELETEQUEUE As String = "Capture_DeleteQueue"
        Public Const cCAPTUREEDOCATTRIBUTES As String = "Capture_EDOCAttributes"
        Public Const cCAPTUREEXTENSIONS As String = "Capture_Extensions"
        Public Const cCAPTUREFILTERS As String = "Capture_Filters"
        Public Const cCAPTUREFILTERRULES As String = "Capture_FiltersRules"
        Public Const cCAPTUREKEYWORDS As String = "Capture_Keywords"
        Public Const cCAPTURELINEPARAGRAPH As String = "Capture_LineParagraph"
        Public Const cCAPTURELINKRULETAGS As String = "Capture_LinkRulesTags"
        Public Const cCAPTUREQUEUE As String = "Capture_Queue"
        Public Const cCAPTUREQUEUEWEB As String = "Capture_Queue_Web"
        Public Const cCAPTURERULES As String = "Capture_Rules"
        Public Const cCAPTURERULESSEEDLIST As String = "Capture_RulesSeedList"
        Public Const cCAPTURESEEDLIST As String = "Capture_SeedList"
        Public Const cCAPTURESYSINFO As String = "sysxinfo"
        Public Const cCAPTURESTATS As String = "Capture_Stats"
        Public Const cCAPTURESTOPWORDS As String = "Capture_StopWords"
        Public Const cCAPTURETAGS As String = "Capture_Tags"
        Public Const cCAPTUREGROUPFOLDERS As String = "Capture_GroupFolders"
        Public Const cCAPTUREDIRECTORIES As String = "Capture_Directories"
        Public Const cCAPTURECONFIGURATION As String = "Capture_Configuration"
        Public Const cCAPTURESTATUS As String = "Capture_Status"
        ''--09/10/2015 -- Neelam -- Added for RecCapture 2.7.1
        Public Const cCAPTUREEXCLUSIONS As String = "Capture_ExclusionList"
        Public Const cCAPTUREFILTERSTATUS As String = "Capture_FilterStatus"
        Public Const cCAPTUREFILTERDIRECTORIES As String = "Capture_FiltersDirectories"
        Public Const cCAPTUREFILTERSEXTENSIONS As String = "Capture_FiltersExtensions"
        Public Const cCAPTUREQUEUEPROCESSORS As String = "Capture_QueueProcessors"
#End Region

#Region " System "

        Public Const cK1RECORDLOCK As String = "K1RecordLock"
        Public Const cK1SYSTEMFLAGS As String = "K1SystemFlags"
        Public Const cK1SESSION As String = "K1Session"
        Public Const cK1MDPTYPES As String = "K1MDPTypes"
        Public Const cK1PICKLIST As String = "K1PickList"
        Public Const cK1LICENCEFILE As String = "K1LicenseFile"
        Public Const cK1COMPLIANCE As String = "K1Compliance"
        Public Const cK1GROUPS As String = "K1Groups"
        Public Const cK1ARCHIVE As String = "K1Archive"
#End Region

    End Class
#End Region

#Region " System Stored Procedures "

    Public Class StoredProcedures

#Region " Standard "

        Public Const cGETLIST As String = "_GetList"
        Public Const cGETITEM As String = "_GetItem"
        Public Const cINSERT As String = "_Insert"
        Public Const cUPDATE As String = "_Update"
        Public Const cDELETE As String = "_Delete"

#End Region

#Region " UI Functionality "

        Public Const cUI_K1CONFIG_GETDEFAULT As String = "K1Configuration_GetDefault"
        Public Const cUI_GETREQUESTRANGE As String = "spRequest_GetRange"
        Public Const cUI_CHECK_VERSION As String = "spCheckVersion"

#End Region

#Region " Scheduled Tasks "

        Public Const cST_EMAIL_DELETE As String = "SPTask_EmailAnnouncement_Send_Delete"
        Public Const cST_ACTIVATION_NOTIFICATION As String = "SPTask_K1ActivationNotification"

#End Region

#Region " DRM "

        Public Const cDRM_SYSTEMCHECK As String = "drmSystemCheck"

#End Region

#Region " Capture "

        Public Const cCAPTURE_SYSTEMCHECK As String = "rcSystemCheck"
        Public Const cCAPTURE_GETORPHAN As String = "Capture_Filters_GetOrphan"
        Public Const cCAPTURE_SYNCHRONISE As String = "Capture_Tags_Synchronise"
        Public Const cCAPTURE_STATUS_START As String = "Capture_Status_Start"
#End Region

#Region " System "

        Public Const cK1_RECORDLOCK_INSERT As String = "K1RecordLock_Insert"
        Public Const cK1_RECORDLOCK_UPDATE As String = "K1RecordLock_Update"
#End Region

    End Class
#End Region

#Region " System Files "

    Public Class SystemFiles
        Public Const cstrLICENCE_FILE_RECFIND As String = "LicenceFile.rec"
        Public Const cstrLICENCE_FILE_K1 As String = "LicenceFile.one"
        Public Const cstrLICENCE_FILE_BUTTON As String = "LicenceFile.but"
        Public Const cstrLICENCE_FILE_TACIT As String = "LicenceFile.tac"
        Public Const cstrLICENCE_FILE_GEM As String = "LicenceFile.gem"
        Public Const cstrLICENCE_FILE_CAPTURE As String = "LicenceFile.cap"
        Public Const cstrLICENCE_FILE_HSSM As String = "LicenceFile.rsc"
        Public Const cstrLICENCE_FILE_MINIAPI As String = "LicenceFile.api"
        Public Const cstrLICENCE_FILE_WEBCLIENT As String = "LicenceFile.wcl"
        Public Const cstrLICENCE_FILE_SHAREPOINT As String = "LicenceFile.spt"
        Public Const cstrLICENCE_FILE_API As String = "LicenceFile.sdk"
        Public Const cstrLICENCE_FILE_ONEILINTEGRATION As String = "LicenceFile.onl"
        Public Const cstrLICENCE_FILE_ARCHIVE As String = "LicenceFile.arc"
        Public Const cstrLICENCE_FILE_RF6CONNECTOR As String = "LicenceFile.rfc"
        Public Const cstrINSTALLATION_FILE As String = "InstallationFile.txt"
        Public Const cstrRF6INSTALLATION_FILE As String = "InstallationFile.rf6"
        Public Const cstrACTIVATIONKEY_FILE As String = "ActivationKey.txt"
        Public Const cstrRECFIND6_HELPFILE As String = "RECFIND6HELP.CHM"
        Public Const cstrR6BUTTON_HELPFILE As String = "BUTTONHELP.CHM"
        Public Const cstrRECSCAN_HELPFILE As String = "RECSCANHELP.CHM"
        Public Const cstrR6DRM_HELPFILE As String = "DRMHELP.CHM"
        ''-- 27/11/2015 -- Neelam -- Added for RecCapture 2.7.1
        Public Const cstrRECCAPTURE_HELPFILE As String = "RECCAPTURE.CHM"
    End Class
#End Region

#Region " Codes "

    Public Class clsEdocStatusCodes
        Public Const Destroyed As String = "Destroyed"
        Public Const Draft As String = "Draft"
        Public Const Published As String = "Published"
        Public Const Archived As String = "Archived"
    End Class

    Public Class clsSystemCode

        Public Const cACTIVE As String = "Active"
        Public Const cINACTIVE As String = "Inactive"
        Public Const cDESTROYED As String = "Destroyed"
        Public Const cCLOSED As String = "Closed"
    End Class

    Public Class clsSystemGlobalCodeIDs

        Public Const cACTIVE As Integer = 11
        Public Const cINACTIVE As Integer = 13
        Public Const cCLOSED As Integer = 15
        Public Const cDESTROYED As Integer = 16
        Public Const cLOST As Integer = 17
        Public Const cARCHIVED As Integer = 18
        Public Const cINTERMEDIATE_STORAGE As Integer = 19
        Public Const cOFFSITE_STORAGE As Integer = 20
    End Class
#End Region

#Region " Views "

    Public Class Views
        Public Const cEDOCIndexView As String = "EDOCFieldView"
    End Class

#End Region

End Class
