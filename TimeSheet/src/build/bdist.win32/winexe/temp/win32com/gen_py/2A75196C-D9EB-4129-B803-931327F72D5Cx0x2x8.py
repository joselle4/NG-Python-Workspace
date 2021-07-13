# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 2.7.2 (default, Jun 12 2011, 15:08:59) [MSC v.1500 32 bit (Intel)]
# From type library '{2A75196C-D9EB-4129-B803-931327F72D5C}'
# On Mon Jun 24 15:47:55 2013
'Microsoft ActiveX Data Objects 2.8 Library'
makepy_version = '0.5.01'
python_version = 0x20702f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{2A75196C-D9EB-4129-B803-931327F72D5C}')
MajorVersion = 2
MinorVersion = 8
LibraryFlags = 8
LCID = 0x0

class constants:
	adPriorityAboveNormal         =4          # from enum ADCPROP_ASYNCTHREADPRIORITY_ENUM
	adPriorityBelowNormal         =2          # from enum ADCPROP_ASYNCTHREADPRIORITY_ENUM
	adPriorityHighest             =5          # from enum ADCPROP_ASYNCTHREADPRIORITY_ENUM
	adPriorityLowest              =1          # from enum ADCPROP_ASYNCTHREADPRIORITY_ENUM
	adPriorityNormal              =3          # from enum ADCPROP_ASYNCTHREADPRIORITY_ENUM
	adRecalcAlways                =1          # from enum ADCPROP_AUTORECALC_ENUM
	adRecalcUpFront               =0          # from enum ADCPROP_AUTORECALC_ENUM
	adCriteriaAllCols             =1          # from enum ADCPROP_UPDATECRITERIA_ENUM
	adCriteriaKey                 =0          # from enum ADCPROP_UPDATECRITERIA_ENUM
	adCriteriaTimeStamp           =3          # from enum ADCPROP_UPDATECRITERIA_ENUM
	adCriteriaUpdCols             =2          # from enum ADCPROP_UPDATECRITERIA_ENUM
	adResyncAll                   =15         # from enum ADCPROP_UPDATERESYNC_ENUM
	adResyncAutoIncrement         =1          # from enum ADCPROP_UPDATERESYNC_ENUM
	adResyncConflicts             =2          # from enum ADCPROP_UPDATERESYNC_ENUM
	adResyncInserts               =8          # from enum ADCPROP_UPDATERESYNC_ENUM
	adResyncNone                  =0          # from enum ADCPROP_UPDATERESYNC_ENUM
	adResyncUpdates               =4          # from enum ADCPROP_UPDATERESYNC_ENUM
	adAffectAll                   =3          # from enum AffectEnum
	adAffectAllChapters           =4          # from enum AffectEnum
	adAffectCurrent               =1          # from enum AffectEnum
	adAffectGroup                 =2          # from enum AffectEnum
	adBookmarkCurrent             =0          # from enum BookmarkEnum
	adBookmarkFirst               =1          # from enum BookmarkEnum
	adBookmarkLast                =2          # from enum BookmarkEnum
	adCmdFile                     =256        # from enum CommandTypeEnum
	adCmdStoredProc               =4          # from enum CommandTypeEnum
	adCmdTable                    =2          # from enum CommandTypeEnum
	adCmdTableDirect              =512        # from enum CommandTypeEnum
	adCmdText                     =1          # from enum CommandTypeEnum
	adCmdUnknown                  =8          # from enum CommandTypeEnum
	adCmdUnspecified              =-1         # from enum CommandTypeEnum
	adCompareEqual                =1          # from enum CompareEnum
	adCompareGreaterThan          =2          # from enum CompareEnum
	adCompareLessThan             =0          # from enum CompareEnum
	adCompareNotComparable        =4          # from enum CompareEnum
	adCompareNotEqual             =3          # from enum CompareEnum
	adModeRead                    =1          # from enum ConnectModeEnum
	adModeReadWrite               =3          # from enum ConnectModeEnum
	adModeRecursive               =4194304    # from enum ConnectModeEnum
	adModeShareDenyNone           =16         # from enum ConnectModeEnum
	adModeShareDenyRead           =4          # from enum ConnectModeEnum
	adModeShareDenyWrite          =8          # from enum ConnectModeEnum
	adModeShareExclusive          =12         # from enum ConnectModeEnum
	adModeUnknown                 =0          # from enum ConnectModeEnum
	adModeWrite                   =2          # from enum ConnectModeEnum
	adAsyncConnect                =16         # from enum ConnectOptionEnum
	adConnectUnspecified          =-1         # from enum ConnectOptionEnum
	adPromptAlways                =1          # from enum ConnectPromptEnum
	adPromptComplete              =2          # from enum ConnectPromptEnum
	adPromptCompleteRequired      =3          # from enum ConnectPromptEnum
	adPromptNever                 =4          # from enum ConnectPromptEnum
	adCopyAllowEmulation          =4          # from enum CopyRecordOptionsEnum
	adCopyNonRecursive            =2          # from enum CopyRecordOptionsEnum
	adCopyOverWrite               =1          # from enum CopyRecordOptionsEnum
	adCopyUnspecified             =-1         # from enum CopyRecordOptionsEnum
	adUseClient                   =3          # from enum CursorLocationEnum
	adUseClientBatch              =3          # from enum CursorLocationEnum
	adUseNone                     =1          # from enum CursorLocationEnum
	adUseServer                   =2          # from enum CursorLocationEnum
	adAddNew                      =16778240   # from enum CursorOptionEnum
	adApproxPosition              =16384      # from enum CursorOptionEnum
	adBookmark                    =8192       # from enum CursorOptionEnum
	adDelete                      =16779264   # from enum CursorOptionEnum
	adFind                        =524288     # from enum CursorOptionEnum
	adHoldRecords                 =256        # from enum CursorOptionEnum
	adIndex                       =8388608    # from enum CursorOptionEnum
	adMovePrevious                =512        # from enum CursorOptionEnum
	adNotify                      =262144     # from enum CursorOptionEnum
	adResync                      =131072     # from enum CursorOptionEnum
	adSeek                        =4194304    # from enum CursorOptionEnum
	adUpdate                      =16809984   # from enum CursorOptionEnum
	adUpdateBatch                 =65536      # from enum CursorOptionEnum
	adOpenDynamic                 =2          # from enum CursorTypeEnum
	adOpenForwardOnly             =0          # from enum CursorTypeEnum
	adOpenKeyset                  =1          # from enum CursorTypeEnum
	adOpenStatic                  =3          # from enum CursorTypeEnum
	adOpenUnspecified             =-1         # from enum CursorTypeEnum
	adArray                       =8192       # from enum DataTypeEnum
	adBSTR                        =8          # from enum DataTypeEnum
	adBigInt                      =20         # from enum DataTypeEnum
	adBinary                      =128        # from enum DataTypeEnum
	adBoolean                     =11         # from enum DataTypeEnum
	adChapter                     =136        # from enum DataTypeEnum
	adChar                        =129        # from enum DataTypeEnum
	adCurrency                    =6          # from enum DataTypeEnum
	adDBDate                      =133        # from enum DataTypeEnum
	adDBTime                      =134        # from enum DataTypeEnum
	adDBTimeStamp                 =135        # from enum DataTypeEnum
	adDate                        =7          # from enum DataTypeEnum
	adDecimal                     =14         # from enum DataTypeEnum
	adDouble                      =5          # from enum DataTypeEnum
	adEmpty                       =0          # from enum DataTypeEnum
	adError                       =10         # from enum DataTypeEnum
	adFileTime                    =64         # from enum DataTypeEnum
	adGUID                        =72         # from enum DataTypeEnum
	adIDispatch                   =9          # from enum DataTypeEnum
	adIUnknown                    =13         # from enum DataTypeEnum
	adInteger                     =3          # from enum DataTypeEnum
	adLongVarBinary               =205        # from enum DataTypeEnum
	adLongVarChar                 =201        # from enum DataTypeEnum
	adLongVarWChar                =203        # from enum DataTypeEnum
	adNumeric                     =131        # from enum DataTypeEnum
	adPropVariant                 =138        # from enum DataTypeEnum
	adSingle                      =4          # from enum DataTypeEnum
	adSmallInt                    =2          # from enum DataTypeEnum
	adTinyInt                     =16         # from enum DataTypeEnum
	adUnsignedBigInt              =21         # from enum DataTypeEnum
	adUnsignedInt                 =19         # from enum DataTypeEnum
	adUnsignedSmallInt            =18         # from enum DataTypeEnum
	adUnsignedTinyInt             =17         # from enum DataTypeEnum
	adUserDefined                 =132        # from enum DataTypeEnum
	adVarBinary                   =204        # from enum DataTypeEnum
	adVarChar                     =200        # from enum DataTypeEnum
	adVarNumeric                  =139        # from enum DataTypeEnum
	adVarWChar                    =202        # from enum DataTypeEnum
	adVariant                     =12         # from enum DataTypeEnum
	adWChar                       =130        # from enum DataTypeEnum
	adEditAdd                     =2          # from enum EditModeEnum
	adEditDelete                  =4          # from enum EditModeEnum
	adEditInProgress              =1          # from enum EditModeEnum
	adEditNone                    =0          # from enum EditModeEnum
	adErrBoundToCommand           =3707       # from enum ErrorValueEnum
	adErrCannotComplete           =3732       # from enum ErrorValueEnum
	adErrCantChangeConnection     =3748       # from enum ErrorValueEnum
	adErrCantChangeProvider       =3220       # from enum ErrorValueEnum
	adErrCantConvertvalue         =3724       # from enum ErrorValueEnum
	adErrCantCreate               =3725       # from enum ErrorValueEnum
	adErrCatalogNotSet            =3747       # from enum ErrorValueEnum
	adErrColumnNotOnThisRow       =3726       # from enum ErrorValueEnum
	adErrConnectionStringTooLong  =3754       # from enum ErrorValueEnum
	adErrDataConversion           =3421       # from enum ErrorValueEnum
	adErrDataOverflow             =3721       # from enum ErrorValueEnum
	adErrDelResOutOfScope         =3738       # from enum ErrorValueEnum
	adErrDenyNotSupported         =3750       # from enum ErrorValueEnum
	adErrDenyTypeNotSupported     =3751       # from enum ErrorValueEnum
	adErrFeatureNotAvailable      =3251       # from enum ErrorValueEnum
	adErrFieldsUpdateFailed       =3749       # from enum ErrorValueEnum
	adErrIllegalOperation         =3219       # from enum ErrorValueEnum
	adErrInTransaction            =3246       # from enum ErrorValueEnum
	adErrIntegrityViolation       =3719       # from enum ErrorValueEnum
	adErrInvalidArgument          =3001       # from enum ErrorValueEnum
	adErrInvalidConnection        =3709       # from enum ErrorValueEnum
	adErrInvalidParamInfo         =3708       # from enum ErrorValueEnum
	adErrInvalidTransaction       =3714       # from enum ErrorValueEnum
	adErrInvalidURL               =3729       # from enum ErrorValueEnum
	adErrItemNotFound             =3265       # from enum ErrorValueEnum
	adErrNoCurrentRecord          =3021       # from enum ErrorValueEnum
	adErrNotExecuting             =3715       # from enum ErrorValueEnum
	adErrNotReentrant             =3710       # from enum ErrorValueEnum
	adErrObjectClosed             =3704       # from enum ErrorValueEnum
	adErrObjectInCollection       =3367       # from enum ErrorValueEnum
	adErrObjectNotSet             =3420       # from enum ErrorValueEnum
	adErrObjectOpen               =3705       # from enum ErrorValueEnum
	adErrOpeningFile              =3002       # from enum ErrorValueEnum
	adErrOperationCancelled       =3712       # from enum ErrorValueEnum
	adErrOutOfSpace               =3734       # from enum ErrorValueEnum
	adErrPermissionDenied         =3720       # from enum ErrorValueEnum
	adErrPropConflicting          =3742       # from enum ErrorValueEnum
	adErrPropInvalidColumn        =3739       # from enum ErrorValueEnum
	adErrPropInvalidOption        =3740       # from enum ErrorValueEnum
	adErrPropInvalidValue         =3741       # from enum ErrorValueEnum
	adErrPropNotAllSettable       =3743       # from enum ErrorValueEnum
	adErrPropNotSet               =3744       # from enum ErrorValueEnum
	adErrPropNotSettable          =3745       # from enum ErrorValueEnum
	adErrPropNotSupported         =3746       # from enum ErrorValueEnum
	adErrProviderFailed           =3000       # from enum ErrorValueEnum
	adErrProviderNotFound         =3706       # from enum ErrorValueEnum
	adErrProviderNotSpecified     =3753       # from enum ErrorValueEnum
	adErrReadFile                 =3003       # from enum ErrorValueEnum
	adErrResourceExists           =3731       # from enum ErrorValueEnum
	adErrResourceLocked           =3730       # from enum ErrorValueEnum
	adErrResourceOutOfScope       =3735       # from enum ErrorValueEnum
	adErrSchemaViolation          =3722       # from enum ErrorValueEnum
	adErrSignMismatch             =3723       # from enum ErrorValueEnum
	adErrStillConnecting          =3713       # from enum ErrorValueEnum
	adErrStillExecuting           =3711       # from enum ErrorValueEnum
	adErrTreePermissionDenied     =3728       # from enum ErrorValueEnum
	adErrURLDoesNotExist          =3727       # from enum ErrorValueEnum
	adErrURLNamedRowDoesNotExist  =3737       # from enum ErrorValueEnum
	adErrUnavailable              =3736       # from enum ErrorValueEnum
	adErrUnsafeOperation          =3716       # from enum ErrorValueEnum
	adErrVolumeNotFound           =3733       # from enum ErrorValueEnum
	adErrWriteFile                =3004       # from enum ErrorValueEnum
	adwrnSecurityDialog           =3717       # from enum ErrorValueEnum
	adwrnSecurityDialogHeader     =3718       # from enum ErrorValueEnum
	adRsnAddNew                   =1          # from enum EventReasonEnum
	adRsnClose                    =9          # from enum EventReasonEnum
	adRsnDelete                   =2          # from enum EventReasonEnum
	adRsnFirstChange              =11         # from enum EventReasonEnum
	adRsnMove                     =10         # from enum EventReasonEnum
	adRsnMoveFirst                =12         # from enum EventReasonEnum
	adRsnMoveLast                 =15         # from enum EventReasonEnum
	adRsnMoveNext                 =13         # from enum EventReasonEnum
	adRsnMovePrevious             =14         # from enum EventReasonEnum
	adRsnRequery                  =7          # from enum EventReasonEnum
	adRsnResynch                  =8          # from enum EventReasonEnum
	adRsnUndoAddNew               =5          # from enum EventReasonEnum
	adRsnUndoDelete               =6          # from enum EventReasonEnum
	adRsnUndoUpdate               =4          # from enum EventReasonEnum
	adRsnUpdate                   =3          # from enum EventReasonEnum
	adStatusCancel                =4          # from enum EventStatusEnum
	adStatusCantDeny              =3          # from enum EventStatusEnum
	adStatusErrorsOccurred        =2          # from enum EventStatusEnum
	adStatusOK                    =1          # from enum EventStatusEnum
	adStatusUnwantedEvent         =5          # from enum EventStatusEnum
	adAsyncExecute                =16         # from enum ExecuteOptionEnum
	adAsyncFetch                  =32         # from enum ExecuteOptionEnum
	adAsyncFetchNonBlocking       =64         # from enum ExecuteOptionEnum
	adExecuteNoRecords            =128        # from enum ExecuteOptionEnum
	adExecuteRecord               =2048       # from enum ExecuteOptionEnum
	adExecuteStream               =1024       # from enum ExecuteOptionEnum
	adOptionUnspecified           =-1         # from enum ExecuteOptionEnum
	adFldCacheDeferred            =4096       # from enum FieldAttributeEnum
	adFldFixed                    =16         # from enum FieldAttributeEnum
	adFldIsChapter                =8192       # from enum FieldAttributeEnum
	adFldIsCollection             =262144     # from enum FieldAttributeEnum
	adFldIsDefaultStream          =131072     # from enum FieldAttributeEnum
	adFldIsNullable               =32         # from enum FieldAttributeEnum
	adFldIsRowURL                 =65536      # from enum FieldAttributeEnum
	adFldKeyColumn                =32768      # from enum FieldAttributeEnum
	adFldLong                     =128        # from enum FieldAttributeEnum
	adFldMayBeNull                =64         # from enum FieldAttributeEnum
	adFldMayDefer                 =2          # from enum FieldAttributeEnum
	adFldNegativeScale            =16384      # from enum FieldAttributeEnum
	adFldRowID                    =256        # from enum FieldAttributeEnum
	adFldRowVersion               =512        # from enum FieldAttributeEnum
	adFldUnknownUpdatable         =8          # from enum FieldAttributeEnum
	adFldUnspecified              =-1         # from enum FieldAttributeEnum
	adFldUpdatable                =4          # from enum FieldAttributeEnum
	adDefaultStream               =-1         # from enum FieldEnum
	adRecordURL                   =-2         # from enum FieldEnum
	adFieldAlreadyExists          =26         # from enum FieldStatusEnum
	adFieldBadStatus              =12         # from enum FieldStatusEnum
	adFieldCannotComplete         =20         # from enum FieldStatusEnum
	adFieldCannotDeleteSource     =23         # from enum FieldStatusEnum
	adFieldCantConvertValue       =2          # from enum FieldStatusEnum
	adFieldCantCreate             =7          # from enum FieldStatusEnum
	adFieldDataOverflow           =6          # from enum FieldStatusEnum
	adFieldDefault                =13         # from enum FieldStatusEnum
	adFieldDoesNotExist           =16         # from enum FieldStatusEnum
	adFieldIgnore                 =15         # from enum FieldStatusEnum
	adFieldIntegrityViolation     =10         # from enum FieldStatusEnum
	adFieldInvalidURL             =17         # from enum FieldStatusEnum
	adFieldIsNull                 =3          # from enum FieldStatusEnum
	adFieldOK                     =0          # from enum FieldStatusEnum
	adFieldOutOfSpace             =22         # from enum FieldStatusEnum
	adFieldPendingChange          =262144     # from enum FieldStatusEnum
	adFieldPendingDelete          =131072     # from enum FieldStatusEnum
	adFieldPendingInsert          =65536      # from enum FieldStatusEnum
	adFieldPendingUnknown         =524288     # from enum FieldStatusEnum
	adFieldPendingUnknownDelete   =1048576    # from enum FieldStatusEnum
	adFieldPermissionDenied       =9          # from enum FieldStatusEnum
	adFieldReadOnly               =24         # from enum FieldStatusEnum
	adFieldResourceExists         =19         # from enum FieldStatusEnum
	adFieldResourceLocked         =18         # from enum FieldStatusEnum
	adFieldResourceOutOfScope     =25         # from enum FieldStatusEnum
	adFieldSchemaViolation        =11         # from enum FieldStatusEnum
	adFieldSignMismatch           =5          # from enum FieldStatusEnum
	adFieldTruncated              =4          # from enum FieldStatusEnum
	adFieldUnavailable            =8          # from enum FieldStatusEnum
	adFieldVolumeNotFound         =21         # from enum FieldStatusEnum
	adFilterAffectedRecords       =2          # from enum FilterGroupEnum
	adFilterConflictingRecords    =5          # from enum FilterGroupEnum
	adFilterFetchedRecords        =3          # from enum FilterGroupEnum
	adFilterNone                  =0          # from enum FilterGroupEnum
	adFilterPendingRecords        =1          # from enum FilterGroupEnum
	adFilterPredicate             =4          # from enum FilterGroupEnum
	adGetRowsRest                 =-1         # from enum GetRowsOptionEnum
	adXactBrowse                  =256        # from enum IsolationLevelEnum
	adXactChaos                   =16         # from enum IsolationLevelEnum
	adXactCursorStability         =4096       # from enum IsolationLevelEnum
	adXactIsolated                =1048576    # from enum IsolationLevelEnum
	adXactReadCommitted           =4096       # from enum IsolationLevelEnum
	adXactReadUncommitted         =256        # from enum IsolationLevelEnum
	adXactRepeatableRead          =65536      # from enum IsolationLevelEnum
	adXactSerializable            =1048576    # from enum IsolationLevelEnum
	adXactUnspecified             =-1         # from enum IsolationLevelEnum
	adCR                          =13         # from enum LineSeparatorEnum
	adCRLF                        =-1         # from enum LineSeparatorEnum
	adLF                          =10         # from enum LineSeparatorEnum
	adLockBatchOptimistic         =4          # from enum LockTypeEnum
	adLockOptimistic              =3          # from enum LockTypeEnum
	adLockPessimistic             =2          # from enum LockTypeEnum
	adLockReadOnly                =1          # from enum LockTypeEnum
	adLockUnspecified             =-1         # from enum LockTypeEnum
	adMarshalAll                  =0          # from enum MarshalOptionsEnum
	adMarshalModifiedOnly         =1          # from enum MarshalOptionsEnum
	adMoveAllowEmulation          =4          # from enum MoveRecordOptionsEnum
	adMoveDontUpdateLinks         =2          # from enum MoveRecordOptionsEnum
	adMoveOverWrite               =1          # from enum MoveRecordOptionsEnum
	adMoveUnspecified             =-1         # from enum MoveRecordOptionsEnum
	adStateClosed                 =0          # from enum ObjectStateEnum
	adStateConnecting             =2          # from enum ObjectStateEnum
	adStateExecuting              =4          # from enum ObjectStateEnum
	adStateFetching               =8          # from enum ObjectStateEnum
	adStateOpen                   =1          # from enum ObjectStateEnum
	adParamLong                   =128        # from enum ParameterAttributesEnum
	adParamNullable               =64         # from enum ParameterAttributesEnum
	adParamSigned                 =16         # from enum ParameterAttributesEnum
	adParamInput                  =1          # from enum ParameterDirectionEnum
	adParamInputOutput            =3          # from enum ParameterDirectionEnum
	adParamOutput                 =2          # from enum ParameterDirectionEnum
	adParamReturnValue            =4          # from enum ParameterDirectionEnum
	adParamUnknown                =0          # from enum ParameterDirectionEnum
	adPersistADTG                 =0          # from enum PersistFormatEnum
	adPersistXML                  =1          # from enum PersistFormatEnum
	adPosBOF                      =-2         # from enum PositionEnum
	adPosEOF                      =-3         # from enum PositionEnum
	adPosUnknown                  =-1         # from enum PositionEnum
	adPropNotSupported            =0          # from enum PropertyAttributesEnum
	adPropOptional                =2          # from enum PropertyAttributesEnum
	adPropRead                    =512        # from enum PropertyAttributesEnum
	adPropRequired                =1          # from enum PropertyAttributesEnum
	adPropWrite                   =1024       # from enum PropertyAttributesEnum
	adCreateCollection            =8192       # from enum RecordCreateOptionsEnum
	adCreateNonCollection         =0          # from enum RecordCreateOptionsEnum
	adCreateOverwrite             =67108864   # from enum RecordCreateOptionsEnum
	adCreateStructDoc             =-2147483648 # from enum RecordCreateOptionsEnum
	adFailIfNotExists             =-1         # from enum RecordCreateOptionsEnum
	adOpenIfExists                =33554432   # from enum RecordCreateOptionsEnum
	adDelayFetchFields            =32768      # from enum RecordOpenOptionsEnum
	adDelayFetchStream            =16384      # from enum RecordOpenOptionsEnum
	adOpenAsync                   =4096       # from enum RecordOpenOptionsEnum
	adOpenExecuteCommand          =65536      # from enum RecordOpenOptionsEnum
	adOpenOutput                  =8388608    # from enum RecordOpenOptionsEnum
	adOpenRecordUnspecified       =-1         # from enum RecordOpenOptionsEnum
	adOpenSource                  =8388608    # from enum RecordOpenOptionsEnum
	adRecCanceled                 =256        # from enum RecordStatusEnum
	adRecCantRelease              =1024       # from enum RecordStatusEnum
	adRecConcurrencyViolation     =2048       # from enum RecordStatusEnum
	adRecDBDeleted                =262144     # from enum RecordStatusEnum
	adRecDeleted                  =4          # from enum RecordStatusEnum
	adRecIntegrityViolation       =4096       # from enum RecordStatusEnum
	adRecInvalid                  =16         # from enum RecordStatusEnum
	adRecMaxChangesExceeded       =8192       # from enum RecordStatusEnum
	adRecModified                 =2          # from enum RecordStatusEnum
	adRecMultipleChanges          =64         # from enum RecordStatusEnum
	adRecNew                      =1          # from enum RecordStatusEnum
	adRecOK                       =0          # from enum RecordStatusEnum
	adRecObjectOpen               =16384      # from enum RecordStatusEnum
	adRecOutOfMemory              =32768      # from enum RecordStatusEnum
	adRecPendingChanges           =128        # from enum RecordStatusEnum
	adRecPermissionDenied         =65536      # from enum RecordStatusEnum
	adRecSchemaViolation          =131072     # from enum RecordStatusEnum
	adRecUnmodified               =8          # from enum RecordStatusEnum
	adCollectionRecord            =1          # from enum RecordTypeEnum
	adSimpleRecord                =0          # from enum RecordTypeEnum
	adStructDoc                   =2          # from enum RecordTypeEnum
	adResyncAllValues             =2          # from enum ResyncEnum
	adResyncUnderlyingValues      =1          # from enum ResyncEnum
	adSaveCreateNotExist          =1          # from enum SaveOptionsEnum
	adSaveCreateOverWrite         =2          # from enum SaveOptionsEnum
	adSchemaActions               =41         # from enum SchemaEnum
	adSchemaAsserts               =0          # from enum SchemaEnum
	adSchemaCatalogs              =1          # from enum SchemaEnum
	adSchemaCharacterSets         =2          # from enum SchemaEnum
	adSchemaCheckConstraints      =5          # from enum SchemaEnum
	adSchemaCollations            =3          # from enum SchemaEnum
	adSchemaColumnPrivileges      =13         # from enum SchemaEnum
	adSchemaColumns               =4          # from enum SchemaEnum
	adSchemaColumnsDomainUsage    =11         # from enum SchemaEnum
	adSchemaCommands              =42         # from enum SchemaEnum
	adSchemaConstraintColumnUsage =6          # from enum SchemaEnum
	adSchemaConstraintTableUsage  =7          # from enum SchemaEnum
	adSchemaCubes                 =32         # from enum SchemaEnum
	adSchemaDBInfoKeywords        =30         # from enum SchemaEnum
	adSchemaDBInfoLiterals        =31         # from enum SchemaEnum
	adSchemaDimensions            =33         # from enum SchemaEnum
	adSchemaForeignKeys           =27         # from enum SchemaEnum
	adSchemaFunctions             =40         # from enum SchemaEnum
	adSchemaHierarchies           =34         # from enum SchemaEnum
	adSchemaIndexes               =12         # from enum SchemaEnum
	adSchemaKeyColumnUsage        =8          # from enum SchemaEnum
	adSchemaLevels                =35         # from enum SchemaEnum
	adSchemaMeasures              =36         # from enum SchemaEnum
	adSchemaMembers               =38         # from enum SchemaEnum
	adSchemaPrimaryKeys           =28         # from enum SchemaEnum
	adSchemaProcedureColumns      =29         # from enum SchemaEnum
	adSchemaProcedureParameters   =26         # from enum SchemaEnum
	adSchemaProcedures            =16         # from enum SchemaEnum
	adSchemaProperties            =37         # from enum SchemaEnum
	adSchemaProviderSpecific      =-1         # from enum SchemaEnum
	adSchemaProviderTypes         =22         # from enum SchemaEnum
	adSchemaReferentialConstraints=9          # from enum SchemaEnum
	adSchemaReferentialContraints =9          # from enum SchemaEnum
	adSchemaSQLLanguages          =18         # from enum SchemaEnum
	adSchemaSchemata              =17         # from enum SchemaEnum
	adSchemaSets                  =43         # from enum SchemaEnum
	adSchemaStatistics            =19         # from enum SchemaEnum
	adSchemaTableConstraints      =10         # from enum SchemaEnum
	adSchemaTablePrivileges       =14         # from enum SchemaEnum
	adSchemaTables                =20         # from enum SchemaEnum
	adSchemaTranslations          =21         # from enum SchemaEnum
	adSchemaTrustees              =39         # from enum SchemaEnum
	adSchemaUsagePrivileges       =15         # from enum SchemaEnum
	adSchemaViewColumnUsage       =24         # from enum SchemaEnum
	adSchemaViewTableUsage        =25         # from enum SchemaEnum
	adSchemaViews                 =23         # from enum SchemaEnum
	adSearchBackward              =-1         # from enum SearchDirectionEnum
	adSearchForward               =1          # from enum SearchDirectionEnum
	adSeekAfter                   =8          # from enum SeekEnum
	adSeekAfterEQ                 =4          # from enum SeekEnum
	adSeekBefore                  =32         # from enum SeekEnum
	adSeekBeforeEQ                =16         # from enum SeekEnum
	adSeekFirstEQ                 =1          # from enum SeekEnum
	adSeekLastEQ                  =2          # from enum SeekEnum
	adOpenStreamAsync             =1          # from enum StreamOpenOptionsEnum
	adOpenStreamFromRecord        =4          # from enum StreamOpenOptionsEnum
	adOpenStreamUnspecified       =-1         # from enum StreamOpenOptionsEnum
	adReadAll                     =-1         # from enum StreamReadEnum
	adReadLine                    =-2         # from enum StreamReadEnum
	adTypeBinary                  =1          # from enum StreamTypeEnum
	adTypeText                    =2          # from enum StreamTypeEnum
	adWriteChar                   =0          # from enum StreamWriteEnum
	adWriteLine                   =1          # from enum StreamWriteEnum
	stWriteChar                   =0          # from enum StreamWriteEnum
	stWriteLine                   =1          # from enum StreamWriteEnum
	adClipString                  =2          # from enum StringFormatEnum
	adXactAbortRetaining          =262144     # from enum XactAttributeEnum
	adXactAsyncPhaseOne           =524288     # from enum XactAttributeEnum
	adXactCommitRetaining         =131072     # from enum XactAttributeEnum
	adXactSyncPhaseOne            =1048576    # from enum XactAttributeEnum

from win32com.client import DispatchBaseClass
class ADORecordConstruction(DispatchBaseClass):
	CLSID = IID('{00000567-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	_prop_map_get_ = {
		"Row": (1610743808, 2, (3, 0), ((16397, 10),), "Row", None),
	}
	_prop_map_put_ = {
		"ParentRow": ((1610743810, LCID, 4, 0),()),
		"Row": ((1610743808, LCID, 4, 0),()),
	}

class ADORecordsetConstruction(DispatchBaseClass):
	CLSID = IID('{00000283-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	_prop_map_get_ = {
		"Chapter": (1610743810, 2, (3, 0), ((16387, 10),), "Chapter", None),
		"RowPosition": (1610743812, 2, (3, 0), ((16397, 10),), "RowPosition", None),
		"Rowset": (1610743808, 2, (3, 0), ((16397, 10),), "Rowset", None),
	}
	_prop_map_put_ = {
		"Chapter": ((1610743810, LCID, 4, 0),()),
		"RowPosition": ((1610743812, LCID, 4, 0),()),
		"Rowset": ((1610743808, LCID, 4, 0),()),
	}

class ADOStreamConstruction(DispatchBaseClass):
	CLSID = IID('{00000568-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	_prop_map_get_ = {
		"Stream": (1610743808, 2, (3, 0), ((16397, 10),), "Stream", None),
	}
	_prop_map_put_ = {
		"Stream": ((1610743808, LCID, 4, 0),()),
	}

class Command15(DispatchBaseClass):
	CLSID = IID('{00000508-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	# Result is of type _Parameter
	def CreateParameter(self, Name=u'', Type=0, Direction=1, Size=0
			, Value=defaultNamedOptArg):
		return self._ApplyTypes_(6, 1, (9, 32), ((8, 49), (3, 49), (3, 49), (3, 49), (12, 17)), u'CreateParameter', '{0000050C-0000-0010-8000-00AA006D2EA4}',Name
			, Type, Direction, Size, Value)

	# Result is of type _Recordset
	def Execute(self, RecordsAffected=pythoncom.Missing, Parameters=defaultNamedNotOptArg, Options=-1):
		return self._ApplyTypes_(5, 1, (9, 0), ((16396, 18), (16396, 17), (3, 49)), u'Execute', '{00000556-0000-0010-8000-00AA006D2EA4}',RecordsAffected
			, Parameters, Options)

	# The method SetActiveConnection is actually a property, but must be used as a method to correctly pass the arguments
	def SetActiveConnection(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	_prop_map_get_ = {
		# Method 'ActiveConnection' returns object of type '_Connection'
		"ActiveConnection": (1, 2, (9, 0), (), "ActiveConnection", '{00000550-0000-0010-8000-00AA006D2EA4}'),
		"CommandText": (2, 2, (8, 0), (), "CommandText", None),
		"CommandTimeout": (3, 2, (3, 0), (), "CommandTimeout", None),
		"CommandType": (7, 2, (3, 0), (), "CommandType", None),
		"Name": (8, 2, (8, 0), (), "Name", None),
		# Method 'Parameters' returns object of type 'Parameters'
		"Parameters": (0, 2, (9, 0), (), "Parameters", '{0000050D-0000-0010-8000-00AA006D2EA4}'),
		"Prepared": (4, 2, (11, 0), (), "Prepared", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
	}
	_prop_map_put_ = {
		"ActiveConnection": ((1, LCID, 4, 0),()),
		"CommandText": ((2, LCID, 4, 0),()),
		"CommandTimeout": ((3, LCID, 4, 0),()),
		"CommandType": ((7, LCID, 4, 0),()),
		"Name": ((8, LCID, 4, 0),()),
		"Prepared": ((4, LCID, 4, 0),()),
	}
	# Default property for this class is 'Parameters'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (9, 0), (), "Parameters", '{0000050D-0000-0010-8000-00AA006D2EA4}'))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class Command25(DispatchBaseClass):
	CLSID = IID('{0000054E-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def Cancel(self):
		return self._oleobj_.InvokeTypes(10, LCID, 1, (24, 0), (),)

	# Result is of type _Parameter
	def CreateParameter(self, Name=u'', Type=0, Direction=1, Size=0
			, Value=defaultNamedOptArg):
		return self._ApplyTypes_(6, 1, (9, 32), ((8, 49), (3, 49), (3, 49), (3, 49), (12, 17)), u'CreateParameter', '{0000050C-0000-0010-8000-00AA006D2EA4}',Name
			, Type, Direction, Size, Value)

	# Result is of type _Recordset
	def Execute(self, RecordsAffected=pythoncom.Missing, Parameters=defaultNamedNotOptArg, Options=-1):
		return self._ApplyTypes_(5, 1, (9, 0), ((16396, 18), (16396, 17), (3, 49)), u'Execute', '{00000556-0000-0010-8000-00AA006D2EA4}',RecordsAffected
			, Parameters, Options)

	# The method SetActiveConnection is actually a property, but must be used as a method to correctly pass the arguments
	def SetActiveConnection(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	_prop_map_get_ = {
		# Method 'ActiveConnection' returns object of type '_Connection'
		"ActiveConnection": (1, 2, (9, 0), (), "ActiveConnection", '{00000550-0000-0010-8000-00AA006D2EA4}'),
		"CommandText": (2, 2, (8, 0), (), "CommandText", None),
		"CommandTimeout": (3, 2, (3, 0), (), "CommandTimeout", None),
		"CommandType": (7, 2, (3, 0), (), "CommandType", None),
		"Name": (8, 2, (8, 0), (), "Name", None),
		# Method 'Parameters' returns object of type 'Parameters'
		"Parameters": (0, 2, (9, 0), (), "Parameters", '{0000050D-0000-0010-8000-00AA006D2EA4}'),
		"Prepared": (4, 2, (11, 0), (), "Prepared", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"State": (9, 2, (3, 0), (), "State", None),
	}
	_prop_map_put_ = {
		"ActiveConnection": ((1, LCID, 4, 0),()),
		"CommandText": ((2, LCID, 4, 0),()),
		"CommandTimeout": ((3, LCID, 4, 0),()),
		"CommandType": ((7, LCID, 4, 0),()),
		"Name": ((8, LCID, 4, 0),()),
		"Prepared": ((4, LCID, 4, 0),()),
	}
	# Default property for this class is 'Parameters'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (9, 0), (), "Parameters", '{0000050D-0000-0010-8000-00AA006D2EA4}'))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class Connection15(DispatchBaseClass):
	CLSID = IID('{00000515-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def BeginTrans(self):
		return self._oleobj_.InvokeTypes(7, LCID, 1, (3, 0), (),)

	def Close(self):
		return self._oleobj_.InvokeTypes(5, LCID, 1, (24, 0), (),)

	def CommitTrans(self):
		return self._oleobj_.InvokeTypes(8, LCID, 1, (24, 0), (),)

	# Result is of type _Recordset
	def Execute(self, CommandText=defaultNamedNotOptArg, RecordsAffected=pythoncom.Missing, Options=-1):
		return self._ApplyTypes_(6, 1, (9, 0), ((8, 1), (16396, 18), (3, 49)), u'Execute', '{00000556-0000-0010-8000-00AA006D2EA4}',CommandText
			, RecordsAffected, Options)

	def Open(self, ConnectionString=u'', UserID=u'', Password=u'', Options=-1):
		return self._ApplyTypes_(10, 1, (24, 32), ((8, 49), (8, 49), (8, 49), (3, 49)), u'Open', None,ConnectionString
			, UserID, Password, Options)

	# Result is of type _Recordset
	def OpenSchema(self, Schema=defaultNamedNotOptArg, Restrictions=defaultNamedOptArg, SchemaID=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(19, LCID, 1, (9, 0), ((3, 1), (12, 17), (12, 17)),Schema
			, Restrictions, SchemaID)
		if ret is not None:
			ret = Dispatch(ret, u'OpenSchema', '{00000556-0000-0010-8000-00AA006D2EA4}')
		return ret

	def RollbackTrans(self):
		return self._oleobj_.InvokeTypes(9, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"Attributes": (14, 2, (3, 0), (), "Attributes", None),
		"CommandTimeout": (2, 2, (3, 0), (), "CommandTimeout", None),
		"ConnectionString": (0, 2, (8, 0), (), "ConnectionString", None),
		"ConnectionTimeout": (3, 2, (3, 0), (), "ConnectionTimeout", None),
		"CursorLocation": (15, 2, (3, 0), (), "CursorLocation", None),
		"DefaultDatabase": (12, 2, (8, 0), (), "DefaultDatabase", None),
		# Method 'Errors' returns object of type 'Errors'
		"Errors": (11, 2, (9, 0), (), "Errors", '{00000501-0000-0010-8000-00AA006D2EA4}'),
		"IsolationLevel": (13, 2, (3, 0), (), "IsolationLevel", None),
		"Mode": (16, 2, (3, 0), (), "Mode", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"Provider": (17, 2, (8, 0), (), "Provider", None),
		"State": (18, 2, (3, 0), (), "State", None),
		"Version": (4, 2, (8, 0), (), "Version", None),
	}
	_prop_map_put_ = {
		"Attributes": ((14, LCID, 4, 0),()),
		"CommandTimeout": ((2, LCID, 4, 0),()),
		"ConnectionString": ((0, LCID, 4, 0),()),
		"ConnectionTimeout": ((3, LCID, 4, 0),()),
		"CursorLocation": ((15, LCID, 4, 0),()),
		"DefaultDatabase": ((12, LCID, 4, 0),()),
		"IsolationLevel": ((13, LCID, 4, 0),()),
		"Mode": ((16, LCID, 4, 0),()),
		"Provider": ((17, LCID, 4, 0),()),
	}
	# Default property for this class is 'ConnectionString'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (8, 0), (), "ConnectionString", None))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class ConnectionEvents:
	CLSID = CLSID_Sink = IID('{00000400-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = IID('{00000514-0000-0010-8000-00AA006D2EA4}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		        3 : "OnCommitTransComplete",
		        4 : "OnWillExecute",
		        8 : "OnDisconnect",
		        5 : "OnExecuteComplete",
		        6 : "OnWillConnect",
		        7 : "OnConnectComplete",
		        0 : "OnInfoMessage",
		        1 : "OnBeginTransComplete",
		        2 : "OnRollbackTransComplete",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnCommitTransComplete(self, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pConnection=defaultNamedNotOptArg):
#	def OnWillExecute(self, Source=defaultNamedNotOptArg, CursorType=defaultNamedNotOptArg, LockType=defaultNamedNotOptArg, Options=defaultNamedNotOptArg
#			, adStatus=defaultNamedNotOptArg, pCommand=defaultNamedNotOptArg, pRecordset=defaultNamedNotOptArg, pConnection=defaultNamedNotOptArg):
#	def OnDisconnect(self, adStatus=defaultNamedNotOptArg, pConnection=defaultNamedNotOptArg):
#	def OnExecuteComplete(self, RecordsAffected=defaultNamedNotOptArg, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pCommand=defaultNamedNotOptArg
#			, pRecordset=defaultNamedNotOptArg, pConnection=defaultNamedNotOptArg):
#	def OnWillConnect(self, ConnectionString=defaultNamedNotOptArg, UserID=defaultNamedNotOptArg, Password=defaultNamedNotOptArg, Options=defaultNamedNotOptArg
#			, adStatus=defaultNamedNotOptArg, pConnection=defaultNamedNotOptArg):
#	def OnConnectComplete(self, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pConnection=defaultNamedNotOptArg):
#	def OnInfoMessage(self, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pConnection=defaultNamedNotOptArg):
#	def OnBeginTransComplete(self, TransactionLevel=defaultNamedNotOptArg, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pConnection=defaultNamedNotOptArg):
#	def OnRollbackTransComplete(self, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pConnection=defaultNamedNotOptArg):


class Error(DispatchBaseClass):
	CLSID = IID('{00000500-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	_prop_map_get_ = {
		"Description": (0, 2, (8, 0), (), "Description", None),
		"HelpContext": (4, 2, (3, 0), (), "HelpContext", None),
		"HelpFile": (3, 2, (8, 0), (), "HelpFile", None),
		"NativeError": (6, 2, (3, 0), (), "NativeError", None),
		"Number": (1, 2, (3, 0), (), "Number", None),
		"SQLState": (5, 2, (8, 0), (), "SQLState", None),
		"Source": (2, 2, (8, 0), (), "Source", None),
	}
	_prop_map_put_ = {
	}
	# Default property for this class is 'Description'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (8, 0), (), "Description", None))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class Errors(DispatchBaseClass):
	CLSID = IID('{00000501-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def Clear(self):
		return self._oleobj_.InvokeTypes(1610809345, LCID, 1, (24, 0), (),)

	# Result is of type Error
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, u'Item', '{00000500-0000-0010-8000-00AA006D2EA4}')
		return ret

	def Refresh(self):
		return self._oleobj_.InvokeTypes(2, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"Count": (1, 2, (3, 0), (), "Count", None),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{00000500-0000-0010-8000-00AA006D2EA4}')
		return ret

	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		ob = self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),())
		return win32com.client.util.Iterator(ob, '{00000500-0000-0010-8000-00AA006D2EA4}')
	def _NewEnum(self):
		"Create an enumerator from this object"
		return win32com.client.util.WrapEnum(self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),()),'{00000500-0000-0010-8000-00AA006D2EA4}')
	def __getitem__(self, index):
		"Allow this class to be accessed as a collection"
		if '_enum_' not in self.__dict__:
			self.__dict__['_enum_'] = self._NewEnum()
		return self._enum_.__getitem__(index)
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Field(DispatchBaseClass):
	CLSID = IID('{00000569-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def AppendChunk(self, Data=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1107, LCID, 1, (24, 0), ((12, 1),),Data
			)

	def GetChunk(self, Length=defaultNamedNotOptArg):
		return self._ApplyTypes_(1108, 1, (12, 0), ((3, 1),), u'GetChunk', None,Length
			)

	_prop_map_get_ = {
		"ActualSize": (1109, 2, (3, 0), (), "ActualSize", None),
		"Attributes": (1114, 2, (3, 0), (), "Attributes", None),
		"DataFormat": (1115, 2, (13, 0), (), "DataFormat", None),
		"DefinedSize": (1103, 2, (3, 0), (), "DefinedSize", None),
		"Name": (1100, 2, (8, 0), (), "Name", None),
		"NumericScale": (1113, 2, (17, 0), (), "NumericScale", None),
		"OriginalValue": (1104, 2, (12, 0), (), "OriginalValue", None),
		"Precision": (1112, 2, (17, 0), (), "Precision", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"Status": (1116, 2, (3, 0), (), "Status", None),
		"Type": (1102, 2, (3, 0), (), "Type", None),
		"UnderlyingValue": (1105, 2, (12, 0), (), "UnderlyingValue", None),
		"Value": (0, 2, (12, 0), (), "Value", None),
	}
	_prop_map_put_ = {
		"Attributes": ((1114, LCID, 4, 0),()),
		"DataFormat": ((1115, LCID, 8, 0),()),
		"DefinedSize": ((1103, LCID, 4, 0),()),
		"NumericScale": ((1113, LCID, 4, 0),()),
		"Precision": ((1112, LCID, 4, 0),()),
		"Type": ((1102, LCID, 4, 0),()),
		"Value": ((0, LCID, 4, 0),()),
	}
	# Default property for this class is 'Value'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (12, 0), (), "Value", None))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class Field15(DispatchBaseClass):
	CLSID = IID('{00000505-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def AppendChunk(self, Data=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1107, LCID, 1, (24, 0), ((12, 1),),Data
			)

	def GetChunk(self, Length=defaultNamedNotOptArg):
		return self._ApplyTypes_(1108, 1, (12, 0), ((3, 1),), u'GetChunk', None,Length
			)

	_prop_map_get_ = {
		"ActualSize": (1109, 2, (3, 0), (), "ActualSize", None),
		"Attributes": (1114, 2, (3, 0), (), "Attributes", None),
		"DefinedSize": (1103, 2, (3, 0), (), "DefinedSize", None),
		"Name": (1100, 2, (8, 0), (), "Name", None),
		"NumericScale": (1113, 2, (17, 0), (), "NumericScale", None),
		"OriginalValue": (1104, 2, (12, 0), (), "OriginalValue", None),
		"Precision": (1112, 2, (17, 0), (), "Precision", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"Type": (1102, 2, (3, 0), (), "Type", None),
		"UnderlyingValue": (1105, 2, (12, 0), (), "UnderlyingValue", None),
		"Value": (0, 2, (12, 0), (), "Value", None),
	}
	_prop_map_put_ = {
		"Value": ((0, LCID, 4, 0),()),
	}
	# Default property for this class is 'Value'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (12, 0), (), "Value", None))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class Field20(DispatchBaseClass):
	CLSID = IID('{0000054C-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def AppendChunk(self, Data=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1107, LCID, 1, (24, 0), ((12, 1),),Data
			)

	def GetChunk(self, Length=defaultNamedNotOptArg):
		return self._ApplyTypes_(1108, 1, (12, 0), ((3, 1),), u'GetChunk', None,Length
			)

	_prop_map_get_ = {
		"ActualSize": (1109, 2, (3, 0), (), "ActualSize", None),
		"Attributes": (1114, 2, (3, 0), (), "Attributes", None),
		"DataFormat": (1115, 2, (13, 0), (), "DataFormat", None),
		"DefinedSize": (1103, 2, (3, 0), (), "DefinedSize", None),
		"Name": (1100, 2, (8, 0), (), "Name", None),
		"NumericScale": (1113, 2, (17, 0), (), "NumericScale", None),
		"OriginalValue": (1104, 2, (12, 0), (), "OriginalValue", None),
		"Precision": (1112, 2, (17, 0), (), "Precision", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"Type": (1102, 2, (3, 0), (), "Type", None),
		"UnderlyingValue": (1105, 2, (12, 0), (), "UnderlyingValue", None),
		"Value": (0, 2, (12, 0), (), "Value", None),
	}
	_prop_map_put_ = {
		"Attributes": ((1114, LCID, 4, 0),()),
		"DataFormat": ((1115, LCID, 8, 0),()),
		"DefinedSize": ((1103, LCID, 4, 0),()),
		"NumericScale": ((1113, LCID, 4, 0),()),
		"Precision": ((1112, LCID, 4, 0),()),
		"Type": ((1102, LCID, 4, 0),()),
		"Value": ((0, LCID, 4, 0),()),
	}
	# Default property for this class is 'Value'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (12, 0), (), "Value", None))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class Fields(DispatchBaseClass):
	CLSID = IID('{00000564-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def Append(self, Name=defaultNamedNotOptArg, Type=defaultNamedNotOptArg, DefinedSize=0, Attrib=-1
			, FieldValue=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(3, LCID, 1, (24, 0), ((8, 1), (3, 1), (3, 49), (3, 49), (12, 17)),Name
			, Type, DefinedSize, Attrib, FieldValue)

	def CancelUpdate(self):
		return self._oleobj_.InvokeTypes(7, LCID, 1, (24, 0), (),)

	def Delete(self, Index=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(4, LCID, 1, (24, 0), ((12, 1),),Index
			)

	# Result is of type Field
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, u'Item', '{00000569-0000-0010-8000-00AA006D2EA4}')
		return ret

	def Refresh(self):
		return self._oleobj_.InvokeTypes(2, LCID, 1, (24, 0), (),)

	def Resync(self, ResyncValues=2):
		return self._oleobj_.InvokeTypes(6, LCID, 1, (24, 0), ((3, 49),),ResyncValues
			)

	def Update(self):
		return self._oleobj_.InvokeTypes(5, LCID, 1, (24, 0), (),)

	def _Append(self, Name=defaultNamedNotOptArg, Type=defaultNamedNotOptArg, DefinedSize=0, Attrib=-1):
		return self._oleobj_.InvokeTypes(1610874880, LCID, 1, (24, 0), ((8, 1), (3, 1), (3, 49), (3, 49)),Name
			, Type, DefinedSize, Attrib)

	_prop_map_get_ = {
		"Count": (1, 2, (3, 0), (), "Count", None),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{00000569-0000-0010-8000-00AA006D2EA4}')
		return ret

	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		ob = self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),())
		return win32com.client.util.Iterator(ob, '{00000569-0000-0010-8000-00AA006D2EA4}')
	def _NewEnum(self):
		"Create an enumerator from this object"
		return win32com.client.util.WrapEnum(self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),()),'{00000569-0000-0010-8000-00AA006D2EA4}')
	def __getitem__(self, index):
		"Allow this class to be accessed as a collection"
		if '_enum_' not in self.__dict__:
			self.__dict__['_enum_'] = self._NewEnum()
		return self._enum_.__getitem__(index)
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Fields15(DispatchBaseClass):
	CLSID = IID('{00000506-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	# Result is of type Field
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, u'Item', '{00000569-0000-0010-8000-00AA006D2EA4}')
		return ret

	def Refresh(self):
		return self._oleobj_.InvokeTypes(2, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"Count": (1, 2, (3, 0), (), "Count", None),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{00000569-0000-0010-8000-00AA006D2EA4}')
		return ret

	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		ob = self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),())
		return win32com.client.util.Iterator(ob, '{00000569-0000-0010-8000-00AA006D2EA4}')
	def _NewEnum(self):
		"Create an enumerator from this object"
		return win32com.client.util.WrapEnum(self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),()),'{00000569-0000-0010-8000-00AA006D2EA4}')
	def __getitem__(self, index):
		"Allow this class to be accessed as a collection"
		if '_enum_' not in self.__dict__:
			self.__dict__['_enum_'] = self._NewEnum()
		return self._enum_.__getitem__(index)
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Fields20(DispatchBaseClass):
	CLSID = IID('{0000054D-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def Delete(self, Index=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(4, LCID, 1, (24, 0), ((12, 1),),Index
			)

	# Result is of type Field
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, u'Item', '{00000569-0000-0010-8000-00AA006D2EA4}')
		return ret

	def Refresh(self):
		return self._oleobj_.InvokeTypes(2, LCID, 1, (24, 0), (),)

	def _Append(self, Name=defaultNamedNotOptArg, Type=defaultNamedNotOptArg, DefinedSize=0, Attrib=-1):
		return self._oleobj_.InvokeTypes(1610874880, LCID, 1, (24, 0), ((8, 1), (3, 1), (3, 49), (3, 49)),Name
			, Type, DefinedSize, Attrib)

	_prop_map_get_ = {
		"Count": (1, 2, (3, 0), (), "Count", None),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{00000569-0000-0010-8000-00AA006D2EA4}')
		return ret

	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		ob = self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),())
		return win32com.client.util.Iterator(ob, '{00000569-0000-0010-8000-00AA006D2EA4}')
	def _NewEnum(self):
		"Create an enumerator from this object"
		return win32com.client.util.WrapEnum(self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),()),'{00000569-0000-0010-8000-00AA006D2EA4}')
	def __getitem__(self, index):
		"Allow this class to be accessed as a collection"
		if '_enum_' not in self.__dict__:
			self.__dict__['_enum_'] = self._NewEnum()
		return self._enum_.__getitem__(index)
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Parameters(DispatchBaseClass):
	CLSID = IID('{0000050D-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def Append(self, Object=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1610809344, LCID, 1, (24, 0), ((9, 1),),Object
			)

	def Delete(self, Index=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1610809345, LCID, 1, (24, 0), ((12, 1),),Index
			)

	# Result is of type _Parameter
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, u'Item', '{0000050C-0000-0010-8000-00AA006D2EA4}')
		return ret

	def Refresh(self):
		return self._oleobj_.InvokeTypes(2, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"Count": (1, 2, (3, 0), (), "Count", None),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{0000050C-0000-0010-8000-00AA006D2EA4}')
		return ret

	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		ob = self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),())
		return win32com.client.util.Iterator(ob, '{0000050C-0000-0010-8000-00AA006D2EA4}')
	def _NewEnum(self):
		"Create an enumerator from this object"
		return win32com.client.util.WrapEnum(self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),()),'{0000050C-0000-0010-8000-00AA006D2EA4}')
	def __getitem__(self, index):
		"Allow this class to be accessed as a collection"
		if '_enum_' not in self.__dict__:
			self.__dict__['_enum_'] = self._NewEnum()
		return self._enum_.__getitem__(index)
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Properties(DispatchBaseClass):
	CLSID = IID('{00000504-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	# Result is of type Property
	# The method Item is actually a property, but must be used as a method to correctly pass the arguments
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, u'Item', '{00000503-0000-0010-8000-00AA006D2EA4}')
		return ret

	def Refresh(self):
		return self._oleobj_.InvokeTypes(2, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"Count": (1, 2, (3, 0), (), "Count", None),
	}
	_prop_map_put_ = {
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 2, (9, 0), ((12, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{00000503-0000-0010-8000-00AA006D2EA4}')
		return ret

	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		ob = self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),())
		return win32com.client.util.Iterator(ob, '{00000503-0000-0010-8000-00AA006D2EA4}')
	def _NewEnum(self):
		"Create an enumerator from this object"
		return win32com.client.util.WrapEnum(self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),()),'{00000503-0000-0010-8000-00AA006D2EA4}')
	def __getitem__(self, index):
		"Allow this class to be accessed as a collection"
		if '_enum_' not in self.__dict__:
			self.__dict__['_enum_'] = self._NewEnum()
		return self._enum_.__getitem__(index)
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class Property(DispatchBaseClass):
	CLSID = IID('{00000503-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	_prop_map_get_ = {
		"Attributes": (1610743812, 2, (3, 0), (), "Attributes", None),
		"Name": (1610743810, 2, (8, 0), (), "Name", None),
		"Type": (1610743811, 2, (3, 0), (), "Type", None),
		"Value": (0, 2, (12, 0), (), "Value", None),
	}
	_prop_map_put_ = {
		"Attributes": ((1610743812, LCID, 4, 0),()),
		"Value": ((0, LCID, 4, 0),()),
	}
	# Default property for this class is 'Value'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (12, 0), (), "Value", None))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class Recordset15(DispatchBaseClass):
	CLSID = IID('{0000050E-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def AddNew(self, FieldList=defaultNamedOptArg, Values=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1012, LCID, 1, (24, 0), ((12, 17), (12, 17)),FieldList
			, Values)

	def CancelBatch(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1049, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def CancelUpdate(self):
		return self._oleobj_.InvokeTypes(1013, LCID, 1, (24, 0), (),)

	def Close(self):
		return self._oleobj_.InvokeTypes(1014, LCID, 1, (24, 0), (),)

	# The method Collect is actually a property, but must be used as a method to correctly pass the arguments
	def Collect(self, Index=defaultNamedNotOptArg):
		return self._ApplyTypes_(-8, 2, (12, 0), ((12, 1),), u'Collect', None,Index
			)

	def Delete(self, AffectRecords=1):
		return self._oleobj_.InvokeTypes(1015, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def Find(self, Criteria=defaultNamedNotOptArg, SkipRecords=0, SearchDirection=1, Start=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1058, LCID, 1, (24, 0), ((8, 1), (3, 49), (3, 49), (12, 17)),Criteria
			, SkipRecords, SearchDirection, Start)

	def GetRows(self, Rows=-1, Start=defaultNamedOptArg, Fields=defaultNamedOptArg):
		return self._ApplyTypes_(1016, 1, (12, 0), ((3, 49), (12, 17), (12, 17)), u'GetRows', None,Rows
			, Start, Fields)

	def Move(self, NumRecords=defaultNamedNotOptArg, Start=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1017, LCID, 1, (24, 0), ((3, 1), (12, 17)),NumRecords
			, Start)

	def MoveFirst(self):
		return self._oleobj_.InvokeTypes(1020, LCID, 1, (24, 0), (),)

	def MoveLast(self):
		return self._oleobj_.InvokeTypes(1021, LCID, 1, (24, 0), (),)

	def MoveNext(self):
		return self._oleobj_.InvokeTypes(1018, LCID, 1, (24, 0), (),)

	def MovePrevious(self):
		return self._oleobj_.InvokeTypes(1019, LCID, 1, (24, 0), (),)

	# Result is of type _Recordset
	def NextRecordset(self, RecordsAffected=pythoncom.Missing):
		return self._ApplyTypes_(1052, 1, (9, 0), ((16396, 18),), u'NextRecordset', '{00000556-0000-0010-8000-00AA006D2EA4}',RecordsAffected
			)

	def Open(self, Source=defaultNamedNotOptArg, ActiveConnection=defaultNamedNotOptArg, CursorType=-1, LockType=-1
			, Options=-1):
		return self._oleobj_.InvokeTypes(1022, LCID, 1, (24, 0), ((12, 17), (12, 17), (3, 49), (3, 49), (3, 49)),Source
			, ActiveConnection, CursorType, LockType, Options)

	def Requery(self, Options=-1):
		return self._oleobj_.InvokeTypes(1023, LCID, 1, (24, 0), ((3, 49),),Options
			)

	# The method SetActiveConnection is actually a property, but must be used as a method to correctly pass the arguments
	def SetActiveConnection(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1001, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	# The method SetCollect is actually a property, but must be used as a method to correctly pass the arguments
	def SetCollect(self, Index=defaultNamedNotOptArg, arg1=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(-8, LCID, 4, (24, 0), ((12, 1), (12, 1)),Index
			, arg1)

	# The method SetSource is actually a property, but must be used as a method to correctly pass the arguments
	def SetSource(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1011, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	def Supports(self, CursorOptions=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1036, LCID, 1, (11, 0), ((3, 1),),CursorOptions
			)

	def Update(self, Fields=defaultNamedOptArg, Values=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1025, LCID, 1, (24, 0), ((12, 17), (12, 17)),Fields
			, Values)

	def UpdateBatch(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1035, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	# Result is of type _Recordset
	def _xClone(self):
		ret = self._oleobj_.InvokeTypes(1610809392, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, u'_xClone', '{00000556-0000-0010-8000-00AA006D2EA4}')
		return ret

	def _xResync(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1610809378, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	_prop_map_get_ = {
		"AbsolutePage": (1047, 2, (3, 0), (), "AbsolutePage", None),
		"AbsolutePosition": (1000, 2, (3, 0), (), "AbsolutePosition", None),
		"ActiveConnection": (1001, 2, (12, 0), (), "ActiveConnection", None),
		"BOF": (1002, 2, (11, 0), (), "BOF", None),
		"Bookmark": (1003, 2, (12, 0), (), "Bookmark", None),
		"CacheSize": (1004, 2, (3, 0), (), "CacheSize", None),
		"CursorLocation": (1051, 2, (3, 0), (), "CursorLocation", None),
		"CursorType": (1005, 2, (3, 0), (), "CursorType", None),
		"EOF": (1006, 2, (11, 0), (), "EOF", None),
		"EditMode": (1026, 2, (3, 0), (), "EditMode", None),
		# Method 'Fields' returns object of type 'Fields'
		"Fields": (0, 2, (9, 0), (), "Fields", '{00000564-0000-0010-8000-00AA006D2EA4}'),
		"Filter": (1030, 2, (12, 0), (), "Filter", None),
		"LockType": (1008, 2, (3, 0), (), "LockType", None),
		"MarshalOptions": (1053, 2, (3, 0), (), "MarshalOptions", None),
		"MaxRecords": (1009, 2, (3, 0), (), "MaxRecords", None),
		"PageCount": (1050, 2, (3, 0), (), "PageCount", None),
		"PageSize": (1048, 2, (3, 0), (), "PageSize", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"RecordCount": (1010, 2, (3, 0), (), "RecordCount", None),
		"Sort": (1031, 2, (8, 0), (), "Sort", None),
		"Source": (1011, 2, (12, 0), (), "Source", None),
		"State": (1054, 2, (3, 0), (), "State", None),
		"Status": (1029, 2, (3, 0), (), "Status", None),
	}
	_prop_map_put_ = {
		"AbsolutePage": ((1047, LCID, 4, 0),()),
		"AbsolutePosition": ((1000, LCID, 4, 0),()),
		"ActiveConnection": ((1001, LCID, 4, 0),()),
		"Bookmark": ((1003, LCID, 4, 0),()),
		"CacheSize": ((1004, LCID, 4, 0),()),
		"CursorLocation": ((1051, LCID, 4, 0),()),
		"CursorType": ((1005, LCID, 4, 0),()),
		"Filter": ((1030, LCID, 4, 0),()),
		"LockType": ((1008, LCID, 4, 0),()),
		"MarshalOptions": ((1053, LCID, 4, 0),()),
		"MaxRecords": ((1009, LCID, 4, 0),()),
		"PageSize": ((1048, LCID, 4, 0),()),
		"Sort": ((1031, LCID, 4, 0),()),
		"Source": ((1011, LCID, 4, 0),()),
	}
	# Default property for this class is 'Fields'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (9, 0), (), "Fields", '{00000564-0000-0010-8000-00AA006D2EA4}'))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class Recordset20(DispatchBaseClass):
	CLSID = IID('{0000054F-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def AddNew(self, FieldList=defaultNamedOptArg, Values=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1012, LCID, 1, (24, 0), ((12, 17), (12, 17)),FieldList
			, Values)

	def Cancel(self):
		return self._oleobj_.InvokeTypes(1055, LCID, 1, (24, 0), (),)

	def CancelBatch(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1049, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def CancelUpdate(self):
		return self._oleobj_.InvokeTypes(1013, LCID, 1, (24, 0), (),)

	# Result is of type _Recordset
	def Clone(self, LockType=-1):
		ret = self._oleobj_.InvokeTypes(1034, LCID, 1, (9, 0), ((3, 49),),LockType
			)
		if ret is not None:
			ret = Dispatch(ret, u'Clone', '{00000556-0000-0010-8000-00AA006D2EA4}')
		return ret

	def Close(self):
		return self._oleobj_.InvokeTypes(1014, LCID, 1, (24, 0), (),)

	# The method Collect is actually a property, but must be used as a method to correctly pass the arguments
	def Collect(self, Index=defaultNamedNotOptArg):
		return self._ApplyTypes_(-8, 2, (12, 0), ((12, 1),), u'Collect', None,Index
			)

	def CompareBookmarks(self, Bookmark1=defaultNamedNotOptArg, Bookmark2=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1065, LCID, 1, (3, 0), ((12, 1), (12, 1)),Bookmark1
			, Bookmark2)

	def Delete(self, AffectRecords=1):
		return self._oleobj_.InvokeTypes(1015, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def Find(self, Criteria=defaultNamedNotOptArg, SkipRecords=0, SearchDirection=1, Start=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1058, LCID, 1, (24, 0), ((8, 1), (3, 49), (3, 49), (12, 17)),Criteria
			, SkipRecords, SearchDirection, Start)

	def GetRows(self, Rows=-1, Start=defaultNamedOptArg, Fields=defaultNamedOptArg):
		return self._ApplyTypes_(1016, 1, (12, 0), ((3, 49), (12, 17), (12, 17)), u'GetRows', None,Rows
			, Start, Fields)

	def GetString(self, StringFormat=2, NumRows=-1, ColumnDelimeter=u'', RowDelimeter=u''
			, NullExpr=u''):
		return self._ApplyTypes_(1062, 1, (8, 32), ((3, 49), (3, 49), (8, 49), (8, 49), (8, 49)), u'GetString', None,StringFormat
			, NumRows, ColumnDelimeter, RowDelimeter, NullExpr)

	def Move(self, NumRecords=defaultNamedNotOptArg, Start=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1017, LCID, 1, (24, 0), ((3, 1), (12, 17)),NumRecords
			, Start)

	def MoveFirst(self):
		return self._oleobj_.InvokeTypes(1020, LCID, 1, (24, 0), (),)

	def MoveLast(self):
		return self._oleobj_.InvokeTypes(1021, LCID, 1, (24, 0), (),)

	def MoveNext(self):
		return self._oleobj_.InvokeTypes(1018, LCID, 1, (24, 0), (),)

	def MovePrevious(self):
		return self._oleobj_.InvokeTypes(1019, LCID, 1, (24, 0), (),)

	# Result is of type _Recordset
	def NextRecordset(self, RecordsAffected=pythoncom.Missing):
		return self._ApplyTypes_(1052, 1, (9, 0), ((16396, 18),), u'NextRecordset', '{00000556-0000-0010-8000-00AA006D2EA4}',RecordsAffected
			)

	def Open(self, Source=defaultNamedNotOptArg, ActiveConnection=defaultNamedNotOptArg, CursorType=-1, LockType=-1
			, Options=-1):
		return self._oleobj_.InvokeTypes(1022, LCID, 1, (24, 0), ((12, 17), (12, 17), (3, 49), (3, 49), (3, 49)),Source
			, ActiveConnection, CursorType, LockType, Options)

	def Requery(self, Options=-1):
		return self._oleobj_.InvokeTypes(1023, LCID, 1, (24, 0), ((3, 49),),Options
			)

	def Resync(self, AffectRecords=3, ResyncValues=2):
		return self._oleobj_.InvokeTypes(1024, LCID, 1, (24, 0), ((3, 49), (3, 49)),AffectRecords
			, ResyncValues)

	# The method SetActiveConnection is actually a property, but must be used as a method to correctly pass the arguments
	def SetActiveConnection(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1001, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	# The method SetCollect is actually a property, but must be used as a method to correctly pass the arguments
	def SetCollect(self, Index=defaultNamedNotOptArg, arg1=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(-8, LCID, 4, (24, 0), ((12, 1), (12, 1)),Index
			, arg1)

	# The method SetSource is actually a property, but must be used as a method to correctly pass the arguments
	def SetSource(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1011, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	def Supports(self, CursorOptions=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1036, LCID, 1, (11, 0), ((3, 1),),CursorOptions
			)

	def Update(self, Fields=defaultNamedOptArg, Values=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1025, LCID, 1, (24, 0), ((12, 17), (12, 17)),Fields
			, Values)

	def UpdateBatch(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1035, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	# Result is of type _Recordset
	def _xClone(self):
		ret = self._oleobj_.InvokeTypes(1610809392, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, u'_xClone', '{00000556-0000-0010-8000-00AA006D2EA4}')
		return ret

	def _xResync(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1610809378, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def _xSave(self, FileName=u'', PersistFormat=0):
		return self._ApplyTypes_(1610874883, 1, (24, 32), ((8, 49), (3, 49)), u'_xSave', None,FileName
			, PersistFormat)

	_prop_map_get_ = {
		"AbsolutePage": (1047, 2, (3, 0), (), "AbsolutePage", None),
		"AbsolutePosition": (1000, 2, (3, 0), (), "AbsolutePosition", None),
		"ActiveCommand": (1061, 2, (9, 0), (), "ActiveCommand", None),
		"ActiveConnection": (1001, 2, (12, 0), (), "ActiveConnection", None),
		"BOF": (1002, 2, (11, 0), (), "BOF", None),
		"Bookmark": (1003, 2, (12, 0), (), "Bookmark", None),
		"CacheSize": (1004, 2, (3, 0), (), "CacheSize", None),
		"CursorLocation": (1051, 2, (3, 0), (), "CursorLocation", None),
		"CursorType": (1005, 2, (3, 0), (), "CursorType", None),
		"DataMember": (1064, 2, (8, 0), (), "DataMember", None),
		"DataSource": (1056, 2, (13, 0), (), "DataSource", None),
		"EOF": (1006, 2, (11, 0), (), "EOF", None),
		"EditMode": (1026, 2, (3, 0), (), "EditMode", None),
		# Method 'Fields' returns object of type 'Fields'
		"Fields": (0, 2, (9, 0), (), "Fields", '{00000564-0000-0010-8000-00AA006D2EA4}'),
		"Filter": (1030, 2, (12, 0), (), "Filter", None),
		"LockType": (1008, 2, (3, 0), (), "LockType", None),
		"MarshalOptions": (1053, 2, (3, 0), (), "MarshalOptions", None),
		"MaxRecords": (1009, 2, (3, 0), (), "MaxRecords", None),
		"PageCount": (1050, 2, (3, 0), (), "PageCount", None),
		"PageSize": (1048, 2, (3, 0), (), "PageSize", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"RecordCount": (1010, 2, (3, 0), (), "RecordCount", None),
		"Sort": (1031, 2, (8, 0), (), "Sort", None),
		"Source": (1011, 2, (12, 0), (), "Source", None),
		"State": (1054, 2, (3, 0), (), "State", None),
		"Status": (1029, 2, (3, 0), (), "Status", None),
		"StayInSync": (1063, 2, (11, 0), (), "StayInSync", None),
	}
	_prop_map_put_ = {
		"AbsolutePage": ((1047, LCID, 4, 0),()),
		"AbsolutePosition": ((1000, LCID, 4, 0),()),
		"ActiveConnection": ((1001, LCID, 4, 0),()),
		"Bookmark": ((1003, LCID, 4, 0),()),
		"CacheSize": ((1004, LCID, 4, 0),()),
		"CursorLocation": ((1051, LCID, 4, 0),()),
		"CursorType": ((1005, LCID, 4, 0),()),
		"DataMember": ((1064, LCID, 4, 0),()),
		"DataSource": ((1056, LCID, 8, 0),()),
		"Filter": ((1030, LCID, 4, 0),()),
		"LockType": ((1008, LCID, 4, 0),()),
		"MarshalOptions": ((1053, LCID, 4, 0),()),
		"MaxRecords": ((1009, LCID, 4, 0),()),
		"PageSize": ((1048, LCID, 4, 0),()),
		"Sort": ((1031, LCID, 4, 0),()),
		"Source": ((1011, LCID, 4, 0),()),
		"StayInSync": ((1063, LCID, 4, 0),()),
	}
	# Default property for this class is 'Fields'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (9, 0), (), "Fields", '{00000564-0000-0010-8000-00AA006D2EA4}'))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class Recordset21(DispatchBaseClass):
	CLSID = IID('{00000555-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def AddNew(self, FieldList=defaultNamedOptArg, Values=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1012, LCID, 1, (24, 0), ((12, 17), (12, 17)),FieldList
			, Values)

	def Cancel(self):
		return self._oleobj_.InvokeTypes(1055, LCID, 1, (24, 0), (),)

	def CancelBatch(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1049, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def CancelUpdate(self):
		return self._oleobj_.InvokeTypes(1013, LCID, 1, (24, 0), (),)

	# Result is of type _Recordset
	def Clone(self, LockType=-1):
		ret = self._oleobj_.InvokeTypes(1034, LCID, 1, (9, 0), ((3, 49),),LockType
			)
		if ret is not None:
			ret = Dispatch(ret, u'Clone', '{00000556-0000-0010-8000-00AA006D2EA4}')
		return ret

	def Close(self):
		return self._oleobj_.InvokeTypes(1014, LCID, 1, (24, 0), (),)

	# The method Collect is actually a property, but must be used as a method to correctly pass the arguments
	def Collect(self, Index=defaultNamedNotOptArg):
		return self._ApplyTypes_(-8, 2, (12, 0), ((12, 1),), u'Collect', None,Index
			)

	def CompareBookmarks(self, Bookmark1=defaultNamedNotOptArg, Bookmark2=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1065, LCID, 1, (3, 0), ((12, 1), (12, 1)),Bookmark1
			, Bookmark2)

	def Delete(self, AffectRecords=1):
		return self._oleobj_.InvokeTypes(1015, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def Find(self, Criteria=defaultNamedNotOptArg, SkipRecords=0, SearchDirection=1, Start=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1058, LCID, 1, (24, 0), ((8, 1), (3, 49), (3, 49), (12, 17)),Criteria
			, SkipRecords, SearchDirection, Start)

	def GetRows(self, Rows=-1, Start=defaultNamedOptArg, Fields=defaultNamedOptArg):
		return self._ApplyTypes_(1016, 1, (12, 0), ((3, 49), (12, 17), (12, 17)), u'GetRows', None,Rows
			, Start, Fields)

	def GetString(self, StringFormat=2, NumRows=-1, ColumnDelimeter=u'', RowDelimeter=u''
			, NullExpr=u''):
		return self._ApplyTypes_(1062, 1, (8, 32), ((3, 49), (3, 49), (8, 49), (8, 49), (8, 49)), u'GetString', None,StringFormat
			, NumRows, ColumnDelimeter, RowDelimeter, NullExpr)

	def Move(self, NumRecords=defaultNamedNotOptArg, Start=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1017, LCID, 1, (24, 0), ((3, 1), (12, 17)),NumRecords
			, Start)

	def MoveFirst(self):
		return self._oleobj_.InvokeTypes(1020, LCID, 1, (24, 0), (),)

	def MoveLast(self):
		return self._oleobj_.InvokeTypes(1021, LCID, 1, (24, 0), (),)

	def MoveNext(self):
		return self._oleobj_.InvokeTypes(1018, LCID, 1, (24, 0), (),)

	def MovePrevious(self):
		return self._oleobj_.InvokeTypes(1019, LCID, 1, (24, 0), (),)

	# Result is of type _Recordset
	def NextRecordset(self, RecordsAffected=pythoncom.Missing):
		return self._ApplyTypes_(1052, 1, (9, 0), ((16396, 18),), u'NextRecordset', '{00000556-0000-0010-8000-00AA006D2EA4}',RecordsAffected
			)

	def Open(self, Source=defaultNamedNotOptArg, ActiveConnection=defaultNamedNotOptArg, CursorType=-1, LockType=-1
			, Options=-1):
		return self._oleobj_.InvokeTypes(1022, LCID, 1, (24, 0), ((12, 17), (12, 17), (3, 49), (3, 49), (3, 49)),Source
			, ActiveConnection, CursorType, LockType, Options)

	def Requery(self, Options=-1):
		return self._oleobj_.InvokeTypes(1023, LCID, 1, (24, 0), ((3, 49),),Options
			)

	def Resync(self, AffectRecords=3, ResyncValues=2):
		return self._oleobj_.InvokeTypes(1024, LCID, 1, (24, 0), ((3, 49), (3, 49)),AffectRecords
			, ResyncValues)

	def Seek(self, KeyValues=defaultNamedNotOptArg, SeekOption=1):
		return self._oleobj_.InvokeTypes(1066, LCID, 1, (24, 0), ((12, 1), (3, 49)),KeyValues
			, SeekOption)

	# The method SetActiveConnection is actually a property, but must be used as a method to correctly pass the arguments
	def SetActiveConnection(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1001, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	# The method SetCollect is actually a property, but must be used as a method to correctly pass the arguments
	def SetCollect(self, Index=defaultNamedNotOptArg, arg1=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(-8, LCID, 4, (24, 0), ((12, 1), (12, 1)),Index
			, arg1)

	# The method SetSource is actually a property, but must be used as a method to correctly pass the arguments
	def SetSource(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1011, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	def Supports(self, CursorOptions=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1036, LCID, 1, (11, 0), ((3, 1),),CursorOptions
			)

	def Update(self, Fields=defaultNamedOptArg, Values=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1025, LCID, 1, (24, 0), ((12, 17), (12, 17)),Fields
			, Values)

	def UpdateBatch(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1035, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	# Result is of type _Recordset
	def _xClone(self):
		ret = self._oleobj_.InvokeTypes(1610809392, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, u'_xClone', '{00000556-0000-0010-8000-00AA006D2EA4}')
		return ret

	def _xResync(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1610809378, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def _xSave(self, FileName=u'', PersistFormat=0):
		return self._ApplyTypes_(1610874883, 1, (24, 32), ((8, 49), (3, 49)), u'_xSave', None,FileName
			, PersistFormat)

	_prop_map_get_ = {
		"AbsolutePage": (1047, 2, (3, 0), (), "AbsolutePage", None),
		"AbsolutePosition": (1000, 2, (3, 0), (), "AbsolutePosition", None),
		"ActiveCommand": (1061, 2, (9, 0), (), "ActiveCommand", None),
		"ActiveConnection": (1001, 2, (12, 0), (), "ActiveConnection", None),
		"BOF": (1002, 2, (11, 0), (), "BOF", None),
		"Bookmark": (1003, 2, (12, 0), (), "Bookmark", None),
		"CacheSize": (1004, 2, (3, 0), (), "CacheSize", None),
		"CursorLocation": (1051, 2, (3, 0), (), "CursorLocation", None),
		"CursorType": (1005, 2, (3, 0), (), "CursorType", None),
		"DataMember": (1064, 2, (8, 0), (), "DataMember", None),
		"DataSource": (1056, 2, (13, 0), (), "DataSource", None),
		"EOF": (1006, 2, (11, 0), (), "EOF", None),
		"EditMode": (1026, 2, (3, 0), (), "EditMode", None),
		# Method 'Fields' returns object of type 'Fields'
		"Fields": (0, 2, (9, 0), (), "Fields", '{00000564-0000-0010-8000-00AA006D2EA4}'),
		"Filter": (1030, 2, (12, 0), (), "Filter", None),
		"Index": (1067, 2, (8, 0), (), "Index", None),
		"LockType": (1008, 2, (3, 0), (), "LockType", None),
		"MarshalOptions": (1053, 2, (3, 0), (), "MarshalOptions", None),
		"MaxRecords": (1009, 2, (3, 0), (), "MaxRecords", None),
		"PageCount": (1050, 2, (3, 0), (), "PageCount", None),
		"PageSize": (1048, 2, (3, 0), (), "PageSize", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"RecordCount": (1010, 2, (3, 0), (), "RecordCount", None),
		"Sort": (1031, 2, (8, 0), (), "Sort", None),
		"Source": (1011, 2, (12, 0), (), "Source", None),
		"State": (1054, 2, (3, 0), (), "State", None),
		"Status": (1029, 2, (3, 0), (), "Status", None),
		"StayInSync": (1063, 2, (11, 0), (), "StayInSync", None),
	}
	_prop_map_put_ = {
		"AbsolutePage": ((1047, LCID, 4, 0),()),
		"AbsolutePosition": ((1000, LCID, 4, 0),()),
		"ActiveConnection": ((1001, LCID, 4, 0),()),
		"Bookmark": ((1003, LCID, 4, 0),()),
		"CacheSize": ((1004, LCID, 4, 0),()),
		"CursorLocation": ((1051, LCID, 4, 0),()),
		"CursorType": ((1005, LCID, 4, 0),()),
		"DataMember": ((1064, LCID, 4, 0),()),
		"DataSource": ((1056, LCID, 8, 0),()),
		"Filter": ((1030, LCID, 4, 0),()),
		"Index": ((1067, LCID, 4, 0),()),
		"LockType": ((1008, LCID, 4, 0),()),
		"MarshalOptions": ((1053, LCID, 4, 0),()),
		"MaxRecords": ((1009, LCID, 4, 0),()),
		"PageSize": ((1048, LCID, 4, 0),()),
		"Sort": ((1031, LCID, 4, 0),()),
		"Source": ((1011, LCID, 4, 0),()),
		"StayInSync": ((1063, LCID, 4, 0),()),
	}
	# Default property for this class is 'Fields'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (9, 0), (), "Fields", '{00000564-0000-0010-8000-00AA006D2EA4}'))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class RecordsetEvents:
	CLSID = CLSID_Sink = IID('{00000266-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = IID('{00000535-0000-0010-8000-00AA006D2EA4}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		       16 : "OnMoveComplete",
		       11 : "OnWillChangeRecord",
		       17 : "OnEndOfRecordset",
		       15 : "OnWillMove",
		       10 : "OnFieldChangeComplete",
		       18 : "OnFetchProgress",
		       13 : "OnWillChangeRecordset",
		       19 : "OnFetchComplete",
		       12 : "OnRecordChangeComplete",
		       14 : "OnRecordsetChangeComplete",
		        9 : "OnWillChangeField",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnMoveComplete(self, adReason=defaultNamedNotOptArg, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pRecordset=defaultNamedNotOptArg):
#	def OnWillChangeRecord(self, adReason=defaultNamedNotOptArg, cRecords=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pRecordset=defaultNamedNotOptArg):
#	def OnEndOfRecordset(self, fMoreData=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pRecordset=defaultNamedNotOptArg):
#	def OnWillMove(self, adReason=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pRecordset=defaultNamedNotOptArg):
#	def OnFieldChangeComplete(self, cFields=defaultNamedNotOptArg, Fields=defaultNamedNotOptArg, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg
#			, pRecordset=defaultNamedNotOptArg):
#	def OnFetchProgress(self, Progress=defaultNamedNotOptArg, MaxProgress=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pRecordset=defaultNamedNotOptArg):
#	def OnWillChangeRecordset(self, adReason=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pRecordset=defaultNamedNotOptArg):
#	def OnFetchComplete(self, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pRecordset=defaultNamedNotOptArg):
#	def OnRecordChangeComplete(self, adReason=defaultNamedNotOptArg, cRecords=defaultNamedNotOptArg, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg
#			, pRecordset=defaultNamedNotOptArg):
#	def OnRecordsetChangeComplete(self, adReason=defaultNamedNotOptArg, pError=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pRecordset=defaultNamedNotOptArg):
#	def OnWillChangeField(self, cFields=defaultNamedNotOptArg, Fields=defaultNamedNotOptArg, adStatus=defaultNamedNotOptArg, pRecordset=defaultNamedNotOptArg):


class _ADO(DispatchBaseClass):
	CLSID = IID('{00000534-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	_prop_map_get_ = {
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
	}
	_prop_map_put_ = {
	}

class _Collection(DispatchBaseClass):
	CLSID = IID('{00000512-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def Refresh(self):
		return self._oleobj_.InvokeTypes(2, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"Count": (1, 2, (3, 0), (), "Count", None),
	}
	_prop_map_put_ = {
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		ob = self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),())
		return win32com.client.util.Iterator(ob, None)
	def _NewEnum(self):
		"Create an enumerator from this object"
		return win32com.client.util.WrapEnum(self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),()),None)
	def __getitem__(self, index):
		"Allow this class to be accessed as a collection"
		if '_enum_' not in self.__dict__:
			self.__dict__['_enum_'] = self._NewEnum()
		return self._enum_.__getitem__(index)
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class _Command(DispatchBaseClass):
	CLSID = IID('{B08400BD-F9D1-4D02-B856-71D5DBA123E9}')
	coclass_clsid = IID('{00000507-0000-0010-8000-00AA006D2EA4}')

	def Cancel(self):
		return self._oleobj_.InvokeTypes(10, LCID, 1, (24, 0), (),)

	# Result is of type _Parameter
	def CreateParameter(self, Name=u'', Type=0, Direction=1, Size=0
			, Value=defaultNamedOptArg):
		return self._ApplyTypes_(6, 1, (9, 32), ((8, 49), (3, 49), (3, 49), (3, 49), (12, 17)), u'CreateParameter', '{0000050C-0000-0010-8000-00AA006D2EA4}',Name
			, Type, Direction, Size, Value)

	# Result is of type _Recordset
	def Execute(self, RecordsAffected=pythoncom.Missing, Parameters=defaultNamedNotOptArg, Options=-1):
		return self._ApplyTypes_(5, 1, (9, 0), ((16396, 18), (16396, 17), (3, 49)), u'Execute', '{00000556-0000-0010-8000-00AA006D2EA4}',RecordsAffected
			, Parameters, Options)

	# The method SetActiveConnection is actually a property, but must be used as a method to correctly pass the arguments
	def SetActiveConnection(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	_prop_map_get_ = {
		# Method 'ActiveConnection' returns object of type '_Connection'
		"ActiveConnection": (1, 2, (9, 0), (), "ActiveConnection", '{00000550-0000-0010-8000-00AA006D2EA4}'),
		"CommandStream": (11, 2, (12, 0), (), "CommandStream", None),
		"CommandText": (2, 2, (8, 0), (), "CommandText", None),
		"CommandTimeout": (3, 2, (3, 0), (), "CommandTimeout", None),
		"CommandType": (7, 2, (3, 0), (), "CommandType", None),
		"Dialect": (12, 2, (8, 0), (), "Dialect", None),
		"Name": (8, 2, (8, 0), (), "Name", None),
		"NamedParameters": (13, 2, (11, 0), (), "NamedParameters", None),
		# Method 'Parameters' returns object of type 'Parameters'
		"Parameters": (0, 2, (9, 0), (), "Parameters", '{0000050D-0000-0010-8000-00AA006D2EA4}'),
		"Prepared": (4, 2, (11, 0), (), "Prepared", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"State": (9, 2, (3, 0), (), "State", None),
	}
	_prop_map_put_ = {
		"ActiveConnection": ((1, LCID, 4, 0),()),
		"CommandStream": ((11, LCID, 8, 0),()),
		"CommandText": ((2, LCID, 4, 0),()),
		"CommandTimeout": ((3, LCID, 4, 0),()),
		"CommandType": ((7, LCID, 4, 0),()),
		"Dialect": ((12, LCID, 4, 0),()),
		"Name": ((8, LCID, 4, 0),()),
		"NamedParameters": ((13, LCID, 4, 0),()),
		"Prepared": ((4, LCID, 4, 0),()),
	}
	# Default property for this class is 'Parameters'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (9, 0), (), "Parameters", '{0000050D-0000-0010-8000-00AA006D2EA4}'))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class _Connection(DispatchBaseClass):
	CLSID = IID('{00000550-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = IID('{00000514-0000-0010-8000-00AA006D2EA4}')

	def BeginTrans(self):
		return self._oleobj_.InvokeTypes(7, LCID, 1, (3, 0), (),)

	def Cancel(self):
		return self._oleobj_.InvokeTypes(21, LCID, 1, (24, 0), (),)

	def Close(self):
		return self._oleobj_.InvokeTypes(5, LCID, 1, (24, 0), (),)

	def CommitTrans(self):
		return self._oleobj_.InvokeTypes(8, LCID, 1, (24, 0), (),)

	# Result is of type _Recordset
	def Execute(self, CommandText=defaultNamedNotOptArg, RecordsAffected=pythoncom.Missing, Options=-1):
		return self._ApplyTypes_(6, 1, (9, 0), ((8, 1), (16396, 18), (3, 49)), u'Execute', '{00000556-0000-0010-8000-00AA006D2EA4}',CommandText
			, RecordsAffected, Options)

	def Open(self, ConnectionString=u'', UserID=u'', Password=u'', Options=-1):
		return self._ApplyTypes_(10, 1, (24, 32), ((8, 49), (8, 49), (8, 49), (3, 49)), u'Open', None,ConnectionString
			, UserID, Password, Options)

	# Result is of type _Recordset
	def OpenSchema(self, Schema=defaultNamedNotOptArg, Restrictions=defaultNamedOptArg, SchemaID=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(19, LCID, 1, (9, 0), ((3, 1), (12, 17), (12, 17)),Schema
			, Restrictions, SchemaID)
		if ret is not None:
			ret = Dispatch(ret, u'OpenSchema', '{00000556-0000-0010-8000-00AA006D2EA4}')
		return ret

	def RollbackTrans(self):
		return self._oleobj_.InvokeTypes(9, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"Attributes": (14, 2, (3, 0), (), "Attributes", None),
		"CommandTimeout": (2, 2, (3, 0), (), "CommandTimeout", None),
		"ConnectionString": (0, 2, (8, 0), (), "ConnectionString", None),
		"ConnectionTimeout": (3, 2, (3, 0), (), "ConnectionTimeout", None),
		"CursorLocation": (15, 2, (3, 0), (), "CursorLocation", None),
		"DefaultDatabase": (12, 2, (8, 0), (), "DefaultDatabase", None),
		# Method 'Errors' returns object of type 'Errors'
		"Errors": (11, 2, (9, 0), (), "Errors", '{00000501-0000-0010-8000-00AA006D2EA4}'),
		"IsolationLevel": (13, 2, (3, 0), (), "IsolationLevel", None),
		"Mode": (16, 2, (3, 0), (), "Mode", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"Provider": (17, 2, (8, 0), (), "Provider", None),
		"State": (18, 2, (3, 0), (), "State", None),
		"Version": (4, 2, (8, 0), (), "Version", None),
	}
	_prop_map_put_ = {
		"Attributes": ((14, LCID, 4, 0),()),
		"CommandTimeout": ((2, LCID, 4, 0),()),
		"ConnectionString": ((0, LCID, 4, 0),()),
		"ConnectionTimeout": ((3, LCID, 4, 0),()),
		"CursorLocation": ((15, LCID, 4, 0),()),
		"DefaultDatabase": ((12, LCID, 4, 0),()),
		"IsolationLevel": ((13, LCID, 4, 0),()),
		"Mode": ((16, LCID, 4, 0),()),
		"Provider": ((17, LCID, 4, 0),()),
	}
	# Default property for this class is 'ConnectionString'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (8, 0), (), "ConnectionString", None))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class _DynaCollection(DispatchBaseClass):
	CLSID = IID('{00000513-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = None

	def Append(self, Object=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1610809344, LCID, 1, (24, 0), ((9, 1),),Object
			)

	def Delete(self, Index=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1610809345, LCID, 1, (24, 0), ((12, 1),),Index
			)

	def Refresh(self):
		return self._oleobj_.InvokeTypes(2, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"Count": (1, 2, (3, 0), (), "Count", None),
	}
	_prop_map_put_ = {
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		ob = self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),())
		return win32com.client.util.Iterator(ob, None)
	def _NewEnum(self):
		"Create an enumerator from this object"
		return win32com.client.util.WrapEnum(self._oleobj_.InvokeTypes(-4,LCID,1,(13, 10),()),None)
	def __getitem__(self, index):
		"Allow this class to be accessed as a collection"
		if '_enum_' not in self.__dict__:
			self.__dict__['_enum_'] = self._NewEnum()
		return self._enum_.__getitem__(index)
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(1, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __nonzero__(self):
		return True

class _Parameter(DispatchBaseClass):
	CLSID = IID('{0000050C-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = IID('{0000050B-0000-0010-8000-00AA006D2EA4}')

	def AppendChunk(self, Val=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(7, LCID, 1, (24, 0), ((12, 1),),Val
			)

	_prop_map_get_ = {
		"Attributes": (8, 2, (3, 0), (), "Attributes", None),
		"Direction": (3, 2, (3, 0), (), "Direction", None),
		"Name": (1, 2, (8, 0), (), "Name", None),
		"NumericScale": (5, 2, (17, 0), (), "NumericScale", None),
		"Precision": (4, 2, (17, 0), (), "Precision", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"Size": (6, 2, (3, 0), (), "Size", None),
		"Type": (2, 2, (3, 0), (), "Type", None),
		"Value": (0, 2, (12, 0), (), "Value", None),
	}
	_prop_map_put_ = {
		"Attributes": ((8, LCID, 4, 0),()),
		"Direction": ((3, LCID, 4, 0),()),
		"Name": ((1, LCID, 4, 0),()),
		"NumericScale": ((5, LCID, 4, 0),()),
		"Precision": ((4, LCID, 4, 0),()),
		"Size": ((6, LCID, 4, 0),()),
		"Type": ((2, LCID, 4, 0),()),
		"Value": ((0, LCID, 4, 0),()),
	}
	# Default property for this class is 'Value'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (12, 0), (), "Value", None))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class _Record(DispatchBaseClass):
	CLSID = IID('{00000562-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = IID('{00000560-0000-0010-8000-00AA006D2EA4}')

	def Cancel(self):
		return self._oleobj_.InvokeTypes(13, LCID, 1, (24, 0), (),)

	def Close(self):
		return self._oleobj_.InvokeTypes(10, LCID, 1, (24, 0), (),)

	def CopyRecord(self, Source=u'', Destination=u'', UserName=u'', Password=u''
			, Options=-1, Async=False):
		return self._ApplyTypes_(7, 1, (8, 32), ((8, 49), (8, 49), (8, 49), (8, 49), (3, 49), (11, 49)), u'CopyRecord', None,Source
			, Destination, UserName, Password, Options, Async
			)

	def DeleteRecord(self, Source=u'', Async=False):
		return self._ApplyTypes_(8, 1, (24, 32), ((8, 49), (11, 49)), u'DeleteRecord', None,Source
			, Async)

	# Result is of type _Recordset
	def GetChildren(self):
		ret = self._oleobj_.InvokeTypes(12, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, u'GetChildren', '{00000556-0000-0010-8000-00AA006D2EA4}')
		return ret

	def MoveRecord(self, Source=u'', Destination=u'', UserName=u'', Password=u''
			, Options=-1, Async=False):
		return self._ApplyTypes_(6, 1, (8, 32), ((8, 49), (8, 49), (8, 49), (8, 49), (3, 49), (11, 49)), u'MoveRecord', None,Source
			, Destination, UserName, Password, Options, Async
			)

	def Open(self, Source=defaultNamedNotOptArg, ActiveConnection=defaultNamedNotOptArg, Mode=0, CreateOptions=-1
			, Options=-1, UserName=u'', Password=u''):
		return self._ApplyTypes_(9, 1, (24, 32), ((12, 17), (12, 17), (3, 49), (3, 49), (3, 49), (8, 49), (8, 49)), u'Open', None,Source
			, ActiveConnection, Mode, CreateOptions, Options, UserName
			, Password)

	def SetActiveConnection(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	def SetSource(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(3, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	_prop_map_get_ = {
		"ActiveConnection": (1, 2, (12, 0), (), "ActiveConnection", None),
		# Method 'Fields' returns object of type 'Fields'
		"Fields": (0, 2, (9, 0), (), "Fields", '{00000564-0000-0010-8000-00AA006D2EA4}'),
		"Mode": (4, 2, (3, 0), (), "Mode", None),
		"ParentURL": (5, 2, (8, 0), (), "ParentURL", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"RecordType": (11, 2, (3, 0), (), "RecordType", None),
		"Source": (3, 2, (12, 0), (), "Source", None),
		"State": (2, 2, (3, 0), (), "State", None),
	}
	_prop_map_put_ = {
		"ActiveConnection": ((1, LCID, 4, 0),()),
		"Mode": ((4, LCID, 4, 0),()),
		"Source": ((3, LCID, 4, 0),()),
	}
	# Default property for this class is 'Fields'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (9, 0), (), "Fields", '{00000564-0000-0010-8000-00AA006D2EA4}'))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class _Recordset(DispatchBaseClass):
	CLSID = IID('{00000556-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = IID('{00000535-0000-0010-8000-00AA006D2EA4}')

	def AddNew(self, FieldList=defaultNamedOptArg, Values=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1012, LCID, 1, (24, 0), ((12, 17), (12, 17)),FieldList
			, Values)

	def Cancel(self):
		return self._oleobj_.InvokeTypes(1055, LCID, 1, (24, 0), (),)

	def CancelBatch(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1049, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def CancelUpdate(self):
		return self._oleobj_.InvokeTypes(1013, LCID, 1, (24, 0), (),)

	# Result is of type _Recordset
	def Clone(self, LockType=-1):
		ret = self._oleobj_.InvokeTypes(1034, LCID, 1, (9, 0), ((3, 49),),LockType
			)
		if ret is not None:
			ret = Dispatch(ret, u'Clone', '{00000556-0000-0010-8000-00AA006D2EA4}')
		return ret

	def Close(self):
		return self._oleobj_.InvokeTypes(1014, LCID, 1, (24, 0), (),)

	# The method Collect is actually a property, but must be used as a method to correctly pass the arguments
	def Collect(self, Index=defaultNamedNotOptArg):
		return self._ApplyTypes_(-8, 2, (12, 0), ((12, 1),), u'Collect', None,Index
			)

	def CompareBookmarks(self, Bookmark1=defaultNamedNotOptArg, Bookmark2=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1065, LCID, 1, (3, 0), ((12, 1), (12, 1)),Bookmark1
			, Bookmark2)

	def Delete(self, AffectRecords=1):
		return self._oleobj_.InvokeTypes(1015, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def Find(self, Criteria=defaultNamedNotOptArg, SkipRecords=0, SearchDirection=1, Start=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1058, LCID, 1, (24, 0), ((8, 1), (3, 49), (3, 49), (12, 17)),Criteria
			, SkipRecords, SearchDirection, Start)

	def GetRows(self, Rows=-1, Start=defaultNamedOptArg, Fields=defaultNamedOptArg):
		return self._ApplyTypes_(1016, 1, (12, 0), ((3, 49), (12, 17), (12, 17)), u'GetRows', None,Rows
			, Start, Fields)

	def GetString(self, StringFormat=2, NumRows=-1, ColumnDelimeter=u'', RowDelimeter=u''
			, NullExpr=u''):
		return self._ApplyTypes_(1062, 1, (8, 32), ((3, 49), (3, 49), (8, 49), (8, 49), (8, 49)), u'GetString', None,StringFormat
			, NumRows, ColumnDelimeter, RowDelimeter, NullExpr)

	def Move(self, NumRecords=defaultNamedNotOptArg, Start=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1017, LCID, 1, (24, 0), ((3, 1), (12, 17)),NumRecords
			, Start)

	def MoveFirst(self):
		return self._oleobj_.InvokeTypes(1020, LCID, 1, (24, 0), (),)

	def MoveLast(self):
		return self._oleobj_.InvokeTypes(1021, LCID, 1, (24, 0), (),)

	def MoveNext(self):
		return self._oleobj_.InvokeTypes(1018, LCID, 1, (24, 0), (),)

	def MovePrevious(self):
		return self._oleobj_.InvokeTypes(1019, LCID, 1, (24, 0), (),)

	# Result is of type _Recordset
	def NextRecordset(self, RecordsAffected=pythoncom.Missing):
		return self._ApplyTypes_(1052, 1, (9, 0), ((16396, 18),), u'NextRecordset', '{00000556-0000-0010-8000-00AA006D2EA4}',RecordsAffected
			)

	def Open(self, Source=defaultNamedNotOptArg, ActiveConnection=defaultNamedNotOptArg, CursorType=-1, LockType=-1
			, Options=-1):
		return self._oleobj_.InvokeTypes(1022, LCID, 1, (24, 0), ((12, 17), (12, 17), (3, 49), (3, 49), (3, 49)),Source
			, ActiveConnection, CursorType, LockType, Options)

	def Requery(self, Options=-1):
		return self._oleobj_.InvokeTypes(1023, LCID, 1, (24, 0), ((3, 49),),Options
			)

	def Resync(self, AffectRecords=3, ResyncValues=2):
		return self._oleobj_.InvokeTypes(1024, LCID, 1, (24, 0), ((3, 49), (3, 49)),AffectRecords
			, ResyncValues)

	def Save(self, Destination=defaultNamedNotOptArg, PersistFormat=0):
		return self._oleobj_.InvokeTypes(1057, LCID, 1, (24, 0), ((12, 17), (3, 49)),Destination
			, PersistFormat)

	def Seek(self, KeyValues=defaultNamedNotOptArg, SeekOption=1):
		return self._oleobj_.InvokeTypes(1066, LCID, 1, (24, 0), ((12, 1), (3, 49)),KeyValues
			, SeekOption)

	# The method SetActiveConnection is actually a property, but must be used as a method to correctly pass the arguments
	def SetActiveConnection(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1001, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	# The method SetCollect is actually a property, but must be used as a method to correctly pass the arguments
	def SetCollect(self, Index=defaultNamedNotOptArg, arg1=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(-8, LCID, 4, (24, 0), ((12, 1), (12, 1)),Index
			, arg1)

	# The method SetSource is actually a property, but must be used as a method to correctly pass the arguments
	def SetSource(self, arg0=defaultUnnamedArg):
		return self._oleobj_.InvokeTypes(1011, LCID, 8, (24, 0), ((9, 1),),arg0
			)

	def Supports(self, CursorOptions=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1036, LCID, 1, (11, 0), ((3, 1),),CursorOptions
			)

	def Update(self, Fields=defaultNamedOptArg, Values=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1025, LCID, 1, (24, 0), ((12, 17), (12, 17)),Fields
			, Values)

	def UpdateBatch(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1035, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	# Result is of type _Recordset
	def _xClone(self):
		ret = self._oleobj_.InvokeTypes(1610809392, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, u'_xClone', '{00000556-0000-0010-8000-00AA006D2EA4}')
		return ret

	def _xResync(self, AffectRecords=3):
		return self._oleobj_.InvokeTypes(1610809378, LCID, 1, (24, 0), ((3, 49),),AffectRecords
			)

	def _xSave(self, FileName=u'', PersistFormat=0):
		return self._ApplyTypes_(1610874883, 1, (24, 32), ((8, 49), (3, 49)), u'_xSave', None,FileName
			, PersistFormat)

	_prop_map_get_ = {
		"AbsolutePage": (1047, 2, (3, 0), (), "AbsolutePage", None),
		"AbsolutePosition": (1000, 2, (3, 0), (), "AbsolutePosition", None),
		"ActiveCommand": (1061, 2, (9, 0), (), "ActiveCommand", None),
		"ActiveConnection": (1001, 2, (12, 0), (), "ActiveConnection", None),
		"BOF": (1002, 2, (11, 0), (), "BOF", None),
		"Bookmark": (1003, 2, (12, 0), (), "Bookmark", None),
		"CacheSize": (1004, 2, (3, 0), (), "CacheSize", None),
		"CursorLocation": (1051, 2, (3, 0), (), "CursorLocation", None),
		"CursorType": (1005, 2, (3, 0), (), "CursorType", None),
		"DataMember": (1064, 2, (8, 0), (), "DataMember", None),
		"DataSource": (1056, 2, (13, 0), (), "DataSource", None),
		"EOF": (1006, 2, (11, 0), (), "EOF", None),
		"EditMode": (1026, 2, (3, 0), (), "EditMode", None),
		# Method 'Fields' returns object of type 'Fields'
		"Fields": (0, 2, (9, 0), (), "Fields", '{00000564-0000-0010-8000-00AA006D2EA4}'),
		"Filter": (1030, 2, (12, 0), (), "Filter", None),
		"Index": (1067, 2, (8, 0), (), "Index", None),
		"LockType": (1008, 2, (3, 0), (), "LockType", None),
		"MarshalOptions": (1053, 2, (3, 0), (), "MarshalOptions", None),
		"MaxRecords": (1009, 2, (3, 0), (), "MaxRecords", None),
		"PageCount": (1050, 2, (3, 0), (), "PageCount", None),
		"PageSize": (1048, 2, (3, 0), (), "PageSize", None),
		# Method 'Properties' returns object of type 'Properties'
		"Properties": (500, 2, (9, 0), (), "Properties", '{00000504-0000-0010-8000-00AA006D2EA4}'),
		"RecordCount": (1010, 2, (3, 0), (), "RecordCount", None),
		"Sort": (1031, 2, (8, 0), (), "Sort", None),
		"Source": (1011, 2, (12, 0), (), "Source", None),
		"State": (1054, 2, (3, 0), (), "State", None),
		"Status": (1029, 2, (3, 0), (), "Status", None),
		"StayInSync": (1063, 2, (11, 0), (), "StayInSync", None),
	}
	_prop_map_put_ = {
		"AbsolutePage": ((1047, LCID, 4, 0),()),
		"AbsolutePosition": ((1000, LCID, 4, 0),()),
		"ActiveConnection": ((1001, LCID, 4, 0),()),
		"Bookmark": ((1003, LCID, 4, 0),()),
		"CacheSize": ((1004, LCID, 4, 0),()),
		"CursorLocation": ((1051, LCID, 4, 0),()),
		"CursorType": ((1005, LCID, 4, 0),()),
		"DataMember": ((1064, LCID, 4, 0),()),
		"DataSource": ((1056, LCID, 8, 0),()),
		"Filter": ((1030, LCID, 4, 0),()),
		"Index": ((1067, LCID, 4, 0),()),
		"LockType": ((1008, LCID, 4, 0),()),
		"MarshalOptions": ((1053, LCID, 4, 0),()),
		"MaxRecords": ((1009, LCID, 4, 0),()),
		"PageSize": ((1048, LCID, 4, 0),()),
		"Sort": ((1031, LCID, 4, 0),()),
		"Source": ((1011, LCID, 4, 0),()),
		"StayInSync": ((1063, LCID, 4, 0),()),
	}
	# Default property for this class is 'Fields'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (9, 0), (), "Fields", '{00000564-0000-0010-8000-00AA006D2EA4}'))
	def __unicode__(self, *args):
		try:
			return unicode(self.__call__(*args))
		except pythoncom.com_error:
			return repr(self)
	def __str__(self, *args):
		return str(self.__unicode__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))

class _Stream(DispatchBaseClass):
	CLSID = IID('{00000565-0000-0010-8000-00AA006D2EA4}')
	coclass_clsid = IID('{00000566-0000-0010-8000-00AA006D2EA4}')

	def Cancel(self):
		return self._oleobj_.InvokeTypes(21, LCID, 1, (24, 0), (),)

	def Close(self):
		return self._oleobj_.InvokeTypes(11, LCID, 1, (24, 0), (),)

	def CopyTo(self, DestStream=defaultNamedNotOptArg, CharNumber=-1):
		return self._oleobj_.InvokeTypes(15, LCID, 1, (24, 0), ((9, 1), (3, 49)),DestStream
			, CharNumber)

	def Flush(self):
		return self._oleobj_.InvokeTypes(16, LCID, 1, (24, 0), (),)

	def LoadFromFile(self, FileName=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(18, LCID, 1, (24, 0), ((8, 1),),FileName
			)

	def Open(self, Source=defaultNamedNotOptArg, Mode=0, Options=-1, UserName=u''
			, Password=u''):
		return self._ApplyTypes_(10, 1, (24, 32), ((12, 17), (3, 49), (3, 49), (8, 49), (8, 49)), u'Open', None,Source
			, Mode, Options, UserName, Password)

	def Read(self, NumBytes=-1):
		return self._ApplyTypes_(9, 1, (12, 0), ((3, 49),), u'Read', None,NumBytes
			)

	def ReadText(self, NumChars=-1):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(19, LCID, 1, (8, 0), ((3, 49),),NumChars
			)

	def SaveToFile(self, FileName=defaultNamedNotOptArg, Options=1):
		return self._oleobj_.InvokeTypes(17, LCID, 1, (24, 0), ((8, 1), (3, 49)),FileName
			, Options)

	def SetEOS(self):
		return self._oleobj_.InvokeTypes(14, LCID, 1, (24, 0), (),)

	def SkipLine(self):
		return self._oleobj_.InvokeTypes(12, LCID, 1, (24, 0), (),)

	def Write(self, Buffer=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(13, LCID, 1, (24, 0), ((12, 1),),Buffer
			)

	def WriteText(self, Data=defaultNamedNotOptArg, Options=0):
		return self._oleobj_.InvokeTypes(20, LCID, 1, (24, 0), ((8, 1), (3, 49)),Data
			, Options)

	_prop_map_get_ = {
		"Charset": (8, 2, (8, 0), (), "Charset", None),
		"EOS": (2, 2, (11, 0), (), "EOS", None),
		"LineSeparator": (5, 2, (3, 0), (), "LineSeparator", None),
		"Mode": (7, 2, (3, 0), (), "Mode", None),
		"Position": (3, 2, (3, 0), (), "Position", None),
		"Size": (1, 2, (3, 0), (), "Size", None),
		"State": (6, 2, (3, 0), (), "State", None),
		"Type": (4, 2, (3, 0), (), "Type", None),
	}
	_prop_map_put_ = {
		"Charset": ((8, LCID, 4, 0),()),
		"LineSeparator": ((5, LCID, 4, 0),()),
		"Mode": ((7, LCID, 4, 0),()),
		"Position": ((3, LCID, 4, 0),()),
		"Type": ((4, LCID, 4, 0),()),
	}

from win32com.client import CoClassBaseClass
# This CoClass is known by the name 'ADODB.Command.2.8'
class Command(CoClassBaseClass): # A CoClass
	CLSID = IID('{00000507-0000-0010-8000-00AA006D2EA4}')
	coclass_sources = [
	]
	coclass_interfaces = [
		_Command,
	]
	default_interface = _Command

# This CoClass is known by the name 'ADODB.Connection.2.8'
class Connection(CoClassBaseClass): # A CoClass
	CLSID = IID('{00000514-0000-0010-8000-00AA006D2EA4}')
	coclass_sources = [
		ConnectionEvents,
	]
	default_source = ConnectionEvents
	coclass_interfaces = [
		_Connection,
	]
	default_interface = _Connection

# This CoClass is known by the name 'ADODB.Parameter.2.8'
class Parameter(CoClassBaseClass): # A CoClass
	CLSID = IID('{0000050B-0000-0010-8000-00AA006D2EA4}')
	coclass_sources = [
	]
	coclass_interfaces = [
		_Parameter,
	]
	default_interface = _Parameter

# This CoClass is known by the name 'ADODB.Record.2.8'
class Record(CoClassBaseClass): # A CoClass
	CLSID = IID('{00000560-0000-0010-8000-00AA006D2EA4}')
	coclass_sources = [
	]
	coclass_interfaces = [
		_Record,
	]
	default_interface = _Record

# This CoClass is known by the name 'ADODB.Recordset.2.8'
class Recordset(CoClassBaseClass): # A CoClass
	CLSID = IID('{00000535-0000-0010-8000-00AA006D2EA4}')
	coclass_sources = [
		RecordsetEvents,
	]
	default_source = RecordsetEvents
	coclass_interfaces = [
		_Recordset,
	]
	default_interface = _Recordset

# This CoClass is known by the name 'ADODB.Stream.2.8'
class Stream(CoClassBaseClass): # A CoClass
	CLSID = IID('{00000566-0000-0010-8000-00AA006D2EA4}')
	coclass_sources = [
	]
	coclass_interfaces = [
		_Stream,
	]
	default_interface = _Stream

ADOCommandConstruction_vtables_dispatch_ = 0
ADOCommandConstruction_vtables_ = [
	(( u'OLEDBCommand' , u'ppOLEDBCommand' , ), 1610678272, (1610678272, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 12 , (3, 0, None, None) , 0 , )),
	(( u'OLEDBCommand' , u'ppOLEDBCommand' , ), 1610678272, (1610678272, (), [ (13, 1, None, None) , ], 1 , 4 , 4 , 0 , 16 , (3, 0, None, None) , 0 , )),
]

ADOConnectionConstruction_vtables_dispatch_ = 0
ADOConnectionConstruction_vtables_ = [
]

ADOConnectionConstruction15_vtables_dispatch_ = 0
ADOConnectionConstruction15_vtables_ = [
	(( u'DSO' , u'ppDSO' , ), 1610678272, (1610678272, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 12 , (3, 0, None, None) , 0 , )),
	(( u'Session' , u'ppSession' , ), 1610678273, (1610678273, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 16 , (3, 0, None, None) , 0 , )),
	(( u'WrapDSOandSession' , u'pDSO' , u'pSession' , ), 1610678274, (1610678274, (), [ (13, 1, None, None) , 
			(13, 1, None, None) , ], 1 , 1 , 4 , 0 , 20 , (3, 0, None, None) , 0 , )),
]

Command15_vtables_dispatch_ = 1
Command15_vtables_ = [
	(( u'ActiveConnection' , u'ppvObject' , ), 1, (1, (), [ (16393, 10, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'ActiveConnection' , u'ppvObject' , ), 1, (1, (), [ (9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 8 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'ActiveConnection' , u'ppvObject' , ), 1, (1, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'CommandText' , u'pbstr' , ), 2, (2, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'CommandText' , u'pbstr' , ), 2, (2, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( u'CommandTimeout' , u'pl' , ), 3, (3, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( u'CommandTimeout' , u'pl' , ), 3, (3, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( u'Prepared' , u'pfPrepared' , ), 4, (4, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( u'Prepared' , u'pfPrepared' , ), 4, (4, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( u'Execute' , u'RecordsAffected' , u'Parameters' , u'Options' , u'ppiRs' , 
			), 5, (5, (), [ (16396, 18, None, None) , (16396, 17, None, None) , (3, 49, '-1', None) , (16393, 10, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( u'CreateParameter' , u'Name' , u'Type' , u'Direction' , u'Size' , 
			u'Value' , u'ppiprm' , ), 6, (6, (), [ (8, 49, "u''", None) , (3, 49, '0', None) , 
			(3, 49, '1', None) , (3, 49, '0', None) , (12, 17, None, None) , (16393, 10, None, "IID('{0000050C-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 1 , 72 , (3, 32, None, None) , 0 , )),
	(( u'Parameters' , u'ppvObject' , ), 0, (0, (), [ (16393, 10, None, "IID('{0000050D-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( u'CommandType' , u'plCmdType' , ), 7, (7, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( u'CommandType' , u'plCmdType' , ), 7, (7, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( u'Name' , u'pbstrName' , ), 8, (8, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( u'Name' , u'pbstrName' , ), 8, (8, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
]

Command25_vtables_dispatch_ = 1
Command25_vtables_ = [
	(( u'State' , u'plObjState' , ), 9, (9, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( u'Cancel' , ), 10, (10, (), [ ], 1 , 1 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
]

Connection15_vtables_dispatch_ = 1
Connection15_vtables_ = [
	(( u'ConnectionString' , u'pbstr' , ), 0, (0, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'ConnectionString' , u'pbstr' , ), 0, (0, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'CommandTimeout' , u'plTimeout' , ), 2, (2, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'CommandTimeout' , u'plTimeout' , ), 2, (2, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'ConnectionTimeout' , u'plTimeout' , ), 3, (3, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( u'ConnectionTimeout' , u'plTimeout' , ), 3, (3, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( u'Version' , u'pbstr' , ), 4, (4, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( u'Close' , ), 5, (5, (), [ ], 1 , 1 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( u'Execute' , u'CommandText' , u'RecordsAffected' , u'Options' , u'ppiRset' , 
			), 6, (6, (), [ (8, 1, None, None) , (16396, 18, None, None) , (3, 49, '-1', None) , (16393, 10, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( u'BeginTrans' , u'TransactionLevel' , ), 7, (7, (), [ (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( u'CommitTrans' , ), 8, (8, (), [ ], 1 , 1 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( u'RollbackTrans' , ), 9, (9, (), [ ], 1 , 1 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( u'Open' , u'ConnectionString' , u'UserID' , u'Password' , u'Options' , 
			), 10, (10, (), [ (8, 49, "u''", None) , (8, 49, "u''", None) , (8, 49, "u''", None) , (3, 49, '-1', None) , ], 1 , 1 , 4 , 0 , 80 , (3, 32, None, None) , 0 , )),
	(( u'Errors' , u'ppvObject' , ), 11, (11, (), [ (16393, 10, None, "IID('{00000501-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( u'DefaultDatabase' , u'pbstr' , ), 12, (12, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( u'DefaultDatabase' , u'pbstr' , ), 12, (12, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( u'IsolationLevel' , u'Level' , ), 13, (13, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( u'IsolationLevel' , u'Level' , ), 13, (13, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( u'Attributes' , u'plAttr' , ), 14, (14, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( u'Attributes' , u'plAttr' , ), 14, (14, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( u'CursorLocation' , u'plCursorLoc' , ), 15, (15, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( u'CursorLocation' , u'plCursorLoc' , ), 15, (15, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( u'Mode' , u'plMode' , ), 16, (16, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( u'Mode' , u'plMode' , ), 16, (16, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( u'Provider' , u'pbstr' , ), 17, (17, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( u'Provider' , u'pbstr' , ), 17, (17, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( u'State' , u'plObjState' , ), 18, (18, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( u'OpenSchema' , u'Schema' , u'Restrictions' , u'SchemaID' , u'pprset' , 
			), 19, (19, (), [ (3, 1, None, None) , (12, 17, None, None) , (12, 17, None, None) , (16393, 10, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 2 , 140 , (3, 0, None, None) , 0 , )),
]

ConnectionEventsVt_vtables_dispatch_ = 0
ConnectionEventsVt_vtables_ = [
	(( u'InfoMessage' , u'pError' , u'adStatus' , u'pConnection' , ), 0, (0, (), [ 
			(9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , (16387, 3, None, None) , (9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 12 , (3, 0, None, None) , 0 , )),
	(( u'BeginTransComplete' , u'TransactionLevel' , u'pError' , u'adStatus' , u'pConnection' , 
			), 1, (1, (), [ (3, 1, None, None) , (9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , (16387, 3, None, None) , (9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 16 , (3, 0, None, None) , 0 , )),
	(( u'CommitTransComplete' , u'pError' , u'adStatus' , u'pConnection' , ), 3, (3, (), [ 
			(9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , (16387, 3, None, None) , (9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 20 , (3, 0, None, None) , 0 , )),
	(( u'RollbackTransComplete' , u'pError' , u'adStatus' , u'pConnection' , ), 2, (2, (), [ 
			(9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , (16387, 3, None, None) , (9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 24 , (3, 0, None, None) , 0 , )),
	(( u'WillExecute' , u'Source' , u'CursorType' , u'LockType' , u'Options' , 
			u'adStatus' , u'pCommand' , u'pRecordset' , u'pConnection' , ), 4, (4, (), [ 
			(16392, 3, None, None) , (16387, 3, None, None) , (16387, 3, None, None) , (16387, 3, None, None) , (16387, 3, None, None) , 
			(9, 1, None, "IID('{B08400BD-F9D1-4D02-B856-71D5DBA123E9}')") , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , (9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( u'ExecuteComplete' , u'RecordsAffected' , u'pError' , u'adStatus' , u'pCommand' , 
			u'pRecordset' , u'pConnection' , ), 5, (5, (), [ (3, 1, None, None) , (9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , 
			(16387, 3, None, None) , (9, 1, None, "IID('{B08400BD-F9D1-4D02-B856-71D5DBA123E9}')") , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , (9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'WillConnect' , u'ConnectionString' , u'UserID' , u'Password' , u'Options' , 
			u'adStatus' , u'pConnection' , ), 6, (6, (), [ (16392, 3, None, None) , (16392, 3, None, None) , 
			(16392, 3, None, None) , (16387, 3, None, None) , (16387, 3, None, None) , (9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'ConnectComplete' , u'pError' , u'adStatus' , u'pConnection' , ), 7, (7, (), [ 
			(9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , (16387, 3, None, None) , (9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'Disconnect' , u'adStatus' , u'pConnection' , ), 8, (8, (), [ (16387, 3, None, None) , 
			(9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
]

Error_vtables_dispatch_ = 1
Error_vtables_ = [
	(( u'Number' , u'pl' , ), 1, (1, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( u'Source' , u'pbstr' , ), 2, (2, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'Description' , u'pbstr' , ), 0, (0, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'HelpFile' , u'pbstr' , ), 3, (3, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'HelpContext' , u'pl' , ), 4, (4, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'SQLState' , u'pbstr' , ), 5, (5, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( u'NativeError' , u'pl' , ), 6, (6, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
]

Errors_vtables_dispatch_ = 1
Errors_vtables_ = [
	(( u'Item' , u'Index' , u'ppvObject' , ), 0, (0, (), [ (12, 1, None, None) , 
			(16393, 10, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'Clear' , ), 1610809345, (1610809345, (), [ ], 1 , 1 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
]

Field_vtables_dispatch_ = 1
Field_vtables_ = [
	(( u'Status' , u'pFStatus' , ), 1116, (1116, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
]

Field15_vtables_dispatch_ = 1
Field15_vtables_ = [
	(( u'ActualSize' , u'pl' , ), 1109, (1109, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'Attributes' , u'pl' , ), 1114, (1114, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'DefinedSize' , u'pl' , ), 1103, (1103, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'Name' , u'pbstr' , ), 1100, (1100, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'Type' , u'pDataType' , ), 1102, (1102, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( u'Value' , u'pvar' , ), 0, (0, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( u'Value' , u'pvar' , ), 0, (0, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( u'Precision' , u'pbPrecision' , ), 1112, (1112, (), [ (16401, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( u'NumericScale' , u'pbNumericScale' , ), 1113, (1113, (), [ (16401, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( u'AppendChunk' , u'Data' , ), 1107, (1107, (), [ (12, 1, None, None) , ], 1 , 1 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( u'GetChunk' , u'Length' , u'pvar' , ), 1108, (1108, (), [ (3, 1, None, None) , 
			(16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( u'OriginalValue' , u'pvar' , ), 1104, (1104, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( u'UnderlyingValue' , u'pvar' , ), 1105, (1105, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
]

Field20_vtables_dispatch_ = 1
Field20_vtables_ = [
	(( u'ActualSize' , u'pl' , ), 1109, (1109, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'Attributes' , u'pl' , ), 1114, (1114, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'DefinedSize' , u'pl' , ), 1103, (1103, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'Name' , u'pbstr' , ), 1100, (1100, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'Type' , u'pDataType' , ), 1102, (1102, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( u'Value' , u'pvar' , ), 0, (0, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( u'Value' , u'pvar' , ), 0, (0, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( u'Precision' , u'pbPrecision' , ), 1112, (1112, (), [ (16401, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( u'NumericScale' , u'pbNumericScale' , ), 1113, (1113, (), [ (16401, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( u'AppendChunk' , u'Data' , ), 1107, (1107, (), [ (12, 1, None, None) , ], 1 , 1 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( u'GetChunk' , u'Length' , u'pvar' , ), 1108, (1108, (), [ (3, 1, None, None) , 
			(16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( u'OriginalValue' , u'pvar' , ), 1104, (1104, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( u'UnderlyingValue' , u'pvar' , ), 1105, (1105, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( u'DataFormat' , u'ppiDF' , ), 1115, (1115, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( u'DataFormat' , u'ppiDF' , ), 1115, (1115, (), [ (13, 1, None, None) , ], 1 , 8 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( u'Precision' , u'pbPrecision' , ), 1112, (1112, (), [ (17, 1, None, None) , ], 1 , 4 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( u'NumericScale' , u'pbNumericScale' , ), 1113, (1113, (), [ (17, 1, None, None) , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( u'Type' , u'pDataType' , ), 1102, (1102, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( u'DefinedSize' , u'pl' , ), 1103, (1103, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( u'Attributes' , u'pl' , ), 1114, (1114, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
]

Fields_vtables_dispatch_ = 1
Fields_vtables_ = [
	(( u'Append' , u'Name' , u'Type' , u'DefinedSize' , u'Attrib' , 
			u'FieldValue' , ), 3, (3, (), [ (8, 1, None, None) , (3, 1, None, None) , (3, 49, '0', None) , 
			(3, 49, '-1', None) , (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 52 , (3, 0, None, None) , 0 , )),
	(( u'Update' , ), 5, (5, (), [ ], 1 , 1 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( u'Resync' , u'ResyncValues' , ), 6, (6, (), [ (3, 49, '2', None) , ], 1 , 1 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( u'CancelUpdate' , ), 7, (7, (), [ ], 1 , 1 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
]

Fields15_vtables_dispatch_ = 1
Fields15_vtables_ = [
	(( u'Item' , u'Index' , u'ppvObject' , ), 0, (0, (), [ (12, 1, None, None) , 
			(16393, 10, None, "IID('{00000569-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
]

Fields20_vtables_dispatch_ = 1
Fields20_vtables_ = [
	(( u'_Append' , u'Name' , u'Type' , u'DefinedSize' , u'Attrib' , 
			), 1610874880, (1610874880, (), [ (8, 1, None, None) , (3, 1, None, None) , (3, 49, '0', None) , (3, 49, '-1', None) , ], 1 , 1 , 4 , 0 , 44 , (3, 0, None, None) , 64 , )),
	(( u'Delete' , u'Index' , ), 4, (4, (), [ (12, 1, None, None) , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
]

Parameters_vtables_dispatch_ = 1
Parameters_vtables_ = [
	(( u'Item' , u'Index' , u'ppvObject' , ), 0, (0, (), [ (12, 1, None, None) , 
			(16393, 10, None, "IID('{0000050C-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
]

Properties_vtables_dispatch_ = 1
Properties_vtables_ = [
	(( u'Item' , u'Index' , u'ppvObject' , ), 0, (0, (), [ (12, 1, None, None) , 
			(16393, 10, None, "IID('{00000503-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
]

Property_vtables_dispatch_ = 1
Property_vtables_ = [
	(( u'Value' , u'pval' , ), 0, (0, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( u'Value' , u'pval' , ), 0, (0, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'Name' , u'pbstr' , ), 1610743810, (1610743810, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'Type' , u'ptype' , ), 1610743811, (1610743811, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'Attributes' , u'plAttributes' , ), 1610743812, (1610743812, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'Attributes' , u'plAttributes' , ), 1610743812, (1610743812, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
]

Recordset15_vtables_dispatch_ = 1
Recordset15_vtables_ = [
	(( u'AbsolutePosition' , u'pl' , ), 1000, (1000, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'AbsolutePosition' , u'pl' , ), 1000, (1000, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'ActiveConnection' , u'pvar' , ), 1001, (1001, (), [ (9, 1, None, None) , ], 1 , 8 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'ActiveConnection' , u'pvar' , ), 1001, (1001, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'ActiveConnection' , u'pvar' , ), 1001, (1001, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( u'BOF' , u'pb' , ), 1002, (1002, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( u'Bookmark' , u'pvBookmark' , ), 1003, (1003, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( u'Bookmark' , u'pvBookmark' , ), 1003, (1003, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( u'CacheSize' , u'pl' , ), 1004, (1004, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( u'CacheSize' , u'pl' , ), 1004, (1004, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( u'CursorType' , u'plCursorType' , ), 1005, (1005, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( u'CursorType' , u'plCursorType' , ), 1005, (1005, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( u'EOF' , u'pb' , ), 1006, (1006, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( u'Fields' , u'ppvObject' , ), 0, (0, (), [ (16393, 10, None, "IID('{00000564-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( u'LockType' , u'plLockType' , ), 1008, (1008, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( u'LockType' , u'plLockType' , ), 1008, (1008, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( u'MaxRecords' , u'plMaxRecords' , ), 1009, (1009, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( u'MaxRecords' , u'plMaxRecords' , ), 1009, (1009, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( u'RecordCount' , u'pl' , ), 1010, (1010, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( u'Source' , u'pvSource' , ), 1011, (1011, (), [ (9, 1, None, None) , ], 1 , 8 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( u'Source' , u'pvSource' , ), 1011, (1011, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( u'Source' , u'pvSource' , ), 1011, (1011, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( u'AddNew' , u'FieldList' , u'Values' , ), 1012, (1012, (), [ (12, 17, None, None) , 
			(12, 17, None, None) , ], 1 , 1 , 4 , 2 , 120 , (3, 0, None, None) , 0 , )),
	(( u'CancelUpdate' , ), 1013, (1013, (), [ ], 1 , 1 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( u'Close' , ), 1014, (1014, (), [ ], 1 , 1 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( u'Delete' , u'AffectRecords' , ), 1015, (1015, (), [ (3, 49, '1', None) , ], 1 , 1 , 4 , 0 , 132 , (3, 0, None, None) , 0 , )),
	(( u'GetRows' , u'Rows' , u'Start' , u'Fields' , u'pvar' , 
			), 1016, (1016, (), [ (3, 49, '-1', None) , (12, 17, None, None) , (12, 17, None, None) , (16396, 10, None, None) , ], 1 , 1 , 4 , 2 , 136 , (3, 0, None, None) , 0 , )),
	(( u'Move' , u'NumRecords' , u'Start' , ), 1017, (1017, (), [ (3, 1, None, None) , 
			(12, 17, None, None) , ], 1 , 1 , 4 , 1 , 140 , (3, 0, None, None) , 0 , )),
	(( u'MoveNext' , ), 1018, (1018, (), [ ], 1 , 1 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( u'MovePrevious' , ), 1019, (1019, (), [ ], 1 , 1 , 4 , 0 , 148 , (3, 0, None, None) , 0 , )),
	(( u'MoveFirst' , ), 1020, (1020, (), [ ], 1 , 1 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( u'MoveLast' , ), 1021, (1021, (), [ ], 1 , 1 , 4 , 0 , 156 , (3, 0, None, None) , 0 , )),
	(( u'Open' , u'Source' , u'ActiveConnection' , u'CursorType' , u'LockType' , 
			u'Options' , ), 1022, (1022, (), [ (12, 17, None, None) , (12, 17, None, None) , (3, 49, '-1', None) , 
			(3, 49, '-1', None) , (3, 49, '-1', None) , ], 1 , 1 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( u'Requery' , u'Options' , ), 1023, (1023, (), [ (3, 49, '-1', None) , ], 1 , 1 , 4 , 0 , 164 , (3, 0, None, None) , 0 , )),
	(( u'_xResync' , u'AffectRecords' , ), 1610809378, (1610809378, (), [ (3, 49, '3', None) , ], 1 , 1 , 4 , 0 , 168 , (3, 0, None, None) , 64 , )),
	(( u'Update' , u'Fields' , u'Values' , ), 1025, (1025, (), [ (12, 17, None, None) , 
			(12, 17, None, None) , ], 1 , 1 , 4 , 2 , 172 , (3, 0, None, None) , 0 , )),
	(( u'AbsolutePage' , u'pl' , ), 1047, (1047, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( u'AbsolutePage' , u'pl' , ), 1047, (1047, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 180 , (3, 0, None, None) , 0 , )),
	(( u'EditMode' , u'pl' , ), 1026, (1026, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( u'Filter' , u'Criteria' , ), 1030, (1030, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 188 , (3, 0, None, None) , 0 , )),
	(( u'Filter' , u'Criteria' , ), 1030, (1030, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( u'PageCount' , u'pl' , ), 1050, (1050, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 196 , (3, 0, None, None) , 0 , )),
	(( u'PageSize' , u'pl' , ), 1048, (1048, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( u'PageSize' , u'pl' , ), 1048, (1048, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 204 , (3, 0, None, None) , 0 , )),
	(( u'Sort' , u'Criteria' , ), 1031, (1031, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( u'Sort' , u'Criteria' , ), 1031, (1031, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 212 , (3, 0, None, None) , 0 , )),
	(( u'Status' , u'pl' , ), 1029, (1029, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( u'State' , u'plObjState' , ), 1054, (1054, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 220 , (3, 0, None, None) , 0 , )),
	(( u'_xClone' , u'ppvObject' , ), 1610809392, (1610809392, (), [ (16393, 10, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 224 , (3, 0, None, None) , 64 , )),
	(( u'UpdateBatch' , u'AffectRecords' , ), 1035, (1035, (), [ (3, 49, '3', None) , ], 1 , 1 , 4 , 0 , 228 , (3, 0, None, None) , 0 , )),
	(( u'CancelBatch' , u'AffectRecords' , ), 1049, (1049, (), [ (3, 49, '3', None) , ], 1 , 1 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( u'CursorLocation' , u'plCursorLoc' , ), 1051, (1051, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 236 , (3, 0, None, None) , 0 , )),
	(( u'CursorLocation' , u'plCursorLoc' , ), 1051, (1051, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( u'NextRecordset' , u'RecordsAffected' , u'ppiRs' , ), 1052, (1052, (), [ (16396, 18, None, None) , 
			(16393, 10, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 1 , 244 , (3, 0, None, None) , 0 , )),
	(( u'Supports' , u'CursorOptions' , u'pb' , ), 1036, (1036, (), [ (3, 1, None, None) , 
			(16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( u'Collect' , u'Index' , u'pvar' , ), -8, (-8, (), [ (12, 1, None, None) , 
			(16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 252 , (3, 0, None, None) , 64 , )),
	(( u'Collect' , u'Index' , u'pvar' , ), -8, (-8, (), [ (12, 1, None, None) , 
			(12, 1, None, None) , ], 1 , 4 , 4 , 0 , 256 , (3, 0, None, None) , 64 , )),
	(( u'MarshalOptions' , u'peMarshal' , ), 1053, (1053, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 260 , (3, 0, None, None) , 0 , )),
	(( u'MarshalOptions' , u'peMarshal' , ), 1053, (1053, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( u'Find' , u'Criteria' , u'SkipRecords' , u'SearchDirection' , u'Start' , 
			), 1058, (1058, (), [ (8, 1, None, None) , (3, 49, '0', None) , (3, 49, '1', None) , (12, 17, None, None) , ], 1 , 1 , 4 , 1 , 268 , (3, 0, None, None) , 0 , )),
]

Recordset20_vtables_dispatch_ = 1
Recordset20_vtables_ = [
	(( u'Cancel' , ), 1055, (1055, (), [ ], 1 , 1 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( u'DataSource' , u'ppunkDataSource' , ), 1056, (1056, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 276 , (3, 0, None, None) , 0 , )),
	(( u'DataSource' , u'ppunkDataSource' , ), 1056, (1056, (), [ (13, 1, None, None) , ], 1 , 8 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( u'_xSave' , u'FileName' , u'PersistFormat' , ), 1610874883, (1610874883, (), [ (8, 49, "u''", None) , 
			(3, 49, '0', None) , ], 1 , 1 , 4 , 0 , 284 , (3, 32, None, None) , 64 , )),
	(( u'ActiveCommand' , u'ppCmd' , ), 1061, (1061, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( u'StayInSync' , u'pbStayInSync' , ), 1063, (1063, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 292 , (3, 0, None, None) , 0 , )),
	(( u'StayInSync' , u'pbStayInSync' , ), 1063, (1063, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( u'GetString' , u'StringFormat' , u'NumRows' , u'ColumnDelimeter' , u'RowDelimeter' , 
			u'NullExpr' , u'pRetString' , ), 1062, (1062, (), [ (3, 49, '2', None) , (3, 49, '-1', None) , 
			(8, 49, "u''", None) , (8, 49, "u''", None) , (8, 49, "u''", None) , (16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 300 , (3, 32, None, None) , 0 , )),
	(( u'DataMember' , u'pbstrDataMember' , ), 1064, (1064, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( u'DataMember' , u'pbstrDataMember' , ), 1064, (1064, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 308 , (3, 0, None, None) , 0 , )),
	(( u'CompareBookmarks' , u'Bookmark1' , u'Bookmark2' , u'pCompare' , ), 1065, (1065, (), [ 
			(12, 1, None, None) , (12, 1, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( u'Clone' , u'LockType' , u'ppvObject' , ), 1034, (1034, (), [ (3, 49, '-1', None) , 
			(16393, 10, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 316 , (3, 0, None, None) , 0 , )),
	(( u'Resync' , u'AffectRecords' , u'ResyncValues' , ), 1024, (1024, (), [ (3, 49, '3', None) , 
			(3, 49, '2', None) , ], 1 , 1 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
]

Recordset21_vtables_dispatch_ = 1
Recordset21_vtables_ = [
	(( u'Seek' , u'KeyValues' , u'SeekOption' , ), 1066, (1066, (), [ (12, 1, None, None) , 
			(3, 49, '1', None) , ], 1 , 1 , 4 , 0 , 324 , (3, 0, None, None) , 0 , )),
	(( u'Index' , u'pbstrIndex' , ), 1067, (1067, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( u'Index' , u'pbstrIndex' , ), 1067, (1067, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 332 , (3, 0, None, None) , 0 , )),
]

RecordsetEventsVt_vtables_dispatch_ = 0
RecordsetEventsVt_vtables_ = [
	(( u'WillChangeField' , u'cFields' , u'Fields' , u'adStatus' , u'pRecordset' , 
			), 9, (9, (), [ (3, 1, None, None) , (12, 1, None, None) , (16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 12 , (3, 0, None, None) , 0 , )),
	(( u'FieldChangeComplete' , u'cFields' , u'Fields' , u'pError' , u'adStatus' , 
			u'pRecordset' , ), 10, (10, (), [ (3, 1, None, None) , (12, 1, None, None) , (9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , 
			(16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 16 , (3, 0, None, None) , 0 , )),
	(( u'WillChangeRecord' , u'adReason' , u'cRecords' , u'adStatus' , u'pRecordset' , 
			), 11, (11, (), [ (3, 1, None, None) , (3, 1, None, None) , (16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 20 , (3, 0, None, None) , 0 , )),
	(( u'RecordChangeComplete' , u'adReason' , u'cRecords' , u'pError' , u'adStatus' , 
			u'pRecordset' , ), 12, (12, (), [ (3, 1, None, None) , (3, 1, None, None) , (9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , 
			(16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 24 , (3, 0, None, None) , 0 , )),
	(( u'WillChangeRecordset' , u'adReason' , u'adStatus' , u'pRecordset' , ), 13, (13, (), [ 
			(3, 1, None, None) , (16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( u'RecordsetChangeComplete' , u'adReason' , u'pError' , u'adStatus' , u'pRecordset' , 
			), 14, (14, (), [ (3, 1, None, None) , (9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , (16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'WillMove' , u'adReason' , u'adStatus' , u'pRecordset' , ), 15, (15, (), [ 
			(3, 1, None, None) , (16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'MoveComplete' , u'adReason' , u'pError' , u'adStatus' , u'pRecordset' , 
			), 16, (16, (), [ (3, 1, None, None) , (9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , (16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'EndOfRecordset' , u'fMoreData' , u'adStatus' , u'pRecordset' , ), 17, (17, (), [ 
			(16395, 3, None, None) , (16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'FetchProgress' , u'Progress' , u'MaxProgress' , u'adStatus' , u'pRecordset' , 
			), 18, (18, (), [ (3, 1, None, None) , (3, 1, None, None) , (16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( u'FetchComplete' , u'pError' , u'adStatus' , u'pRecordset' , ), 19, (19, (), [ 
			(9, 1, None, "IID('{00000500-0000-0010-8000-00AA006D2EA4}')") , (16387, 3, None, None) , (9, 1, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
]

_ADO_vtables_dispatch_ = 1
_ADO_vtables_ = [
	(( u'Properties' , u'ppvObject' , ), 500, (500, (), [ (16393, 10, None, "IID('{00000504-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
]

_Collection_vtables_dispatch_ = 1
_Collection_vtables_ = [
	(( u'Count' , u'c' , ), 1, (1, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( u'_NewEnum' , u'ppvObject' , ), -4, (-4, (), [ (16397, 10, None, None) , ], 1 , 1 , 4 , 0 , 32 , (3, 0, None, None) , 1 , )),
	(( u'Refresh' , ), 2, (2, (), [ ], 1 , 1 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
]

_Command_vtables_dispatch_ = 1
_Command_vtables_ = [
	(( u'CommandStream' , u'pvStream' , ), 11, (11, (), [ (13, 1, None, None) , ], 1 , 8 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( u'CommandStream' , u'pvStream' , ), 11, (11, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( u'Dialect' , u'pbstrDialect' , ), 12, (12, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( u'Dialect' , u'pbstrDialect' , ), 12, (12, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( u'NamedParameters' , u'pfNamedParameters' , ), 13, (13, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( u'NamedParameters' , u'pfNamedParameters' , ), 13, (13, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
]

_Connection_vtables_dispatch_ = 1
_Connection_vtables_ = [
	(( u'Cancel' , ), 21, (21, (), [ ], 1 , 1 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
]

_DynaCollection_vtables_dispatch_ = 1
_DynaCollection_vtables_ = [
	(( u'Append' , u'Object' , ), 1610809344, (1610809344, (), [ (9, 1, None, None) , ], 1 , 1 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'Delete' , u'Index' , ), 1610809345, (1610809345, (), [ (12, 1, None, None) , ], 1 , 1 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
]

_Parameter_vtables_dispatch_ = 1
_Parameter_vtables_ = [
	(( u'Name' , u'pbstr' , ), 1, (1, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'Name' , u'pbstr' , ), 1, (1, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'Value' , u'pvar' , ), 0, (0, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'Value' , u'pvar' , ), 0, (0, (), [ (12, 1, None, None) , ], 1 , 4 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'Type' , u'psDataType' , ), 2, (2, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( u'Type' , u'psDataType' , ), 2, (2, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( u'Direction' , u'plParmDirection' , ), 3, (3, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( u'Direction' , u'plParmDirection' , ), 3, (3, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( u'Precision' , u'pbPrecision' , ), 4, (4, (), [ (17, 1, None, None) , ], 1 , 4 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( u'Precision' , u'pbPrecision' , ), 4, (4, (), [ (16401, 10, None, None) , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( u'NumericScale' , u'pbScale' , ), 5, (5, (), [ (17, 1, None, None) , ], 1 , 4 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( u'NumericScale' , u'pbScale' , ), 5, (5, (), [ (16401, 10, None, None) , ], 1 , 2 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( u'Size' , u'pl' , ), 6, (6, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( u'Size' , u'pl' , ), 6, (6, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 84 , (3, 0, None, None) , 0 , )),
	(( u'AppendChunk' , u'Val' , ), 7, (7, (), [ (12, 1, None, None) , ], 1 , 1 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( u'Attributes' , u'plParmAttribs' , ), 8, (8, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( u'Attributes' , u'plParmAttribs' , ), 8, (8, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
]

_Record_vtables_dispatch_ = 1
_Record_vtables_ = [
	(( u'ActiveConnection' , u'pvar' , ), 1, (1, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'ActiveConnection' , u'pvar' , ), 1, (1, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'ActiveConnection' , u'pvar' , ), 1, (1, (), [ (9, 1, None, "IID('{00000550-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 8 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'State' , u'pState' , ), 2, (2, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'Source' , u'pvar' , ), 3, (3, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( u'Source' , u'pvar' , ), 3, (3, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( u'Source' , u'pvar' , ), 3, (3, (), [ (9, 1, None, None) , ], 1 , 8 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( u'Mode' , u'pMode' , ), 4, (4, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( u'Mode' , u'pMode' , ), 4, (4, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( u'ParentURL' , u'pbstrParentURL' , ), 5, (5, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( u'MoveRecord' , u'Source' , u'Destination' , u'UserName' , u'Password' , 
			u'Options' , u'Async' , u'pbstrNewURL' , ), 6, (6, (), [ (8, 49, "u''", None) , 
			(8, 49, "u''", None) , (8, 49, "u''", None) , (8, 49, "u''", None) , (3, 49, '-1', None) , (11, 49, 'False', None) , 
			(16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 72 , (3, 32, None, None) , 0 , )),
	(( u'CopyRecord' , u'Source' , u'Destination' , u'UserName' , u'Password' , 
			u'Options' , u'Async' , u'pbstrNewURL' , ), 7, (7, (), [ (8, 49, "u''", None) , 
			(8, 49, "u''", None) , (8, 49, "u''", None) , (8, 49, "u''", None) , (3, 49, '-1', None) , (11, 49, 'False', None) , 
			(16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 76 , (3, 32, None, None) , 0 , )),
	(( u'DeleteRecord' , u'Source' , u'Async' , ), 8, (8, (), [ (8, 49, "u''", None) , 
			(11, 49, 'False', None) , ], 1 , 1 , 4 , 0 , 80 , (3, 32, None, None) , 0 , )),
	(( u'Open' , u'Source' , u'ActiveConnection' , u'Mode' , u'CreateOptions' , 
			u'Options' , u'UserName' , u'Password' , ), 9, (9, (), [ (12, 17, None, None) , 
			(12, 17, None, None) , (3, 49, '0', None) , (3, 49, '-1', None) , (3, 49, '-1', None) , (8, 49, "u''", None) , 
			(8, 49, "u''", None) , ], 1 , 1 , 4 , 0 , 84 , (3, 32, None, None) , 0 , )),
	(( u'Close' , ), 10, (10, (), [ ], 1 , 1 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( u'Fields' , u'ppFlds' , ), 0, (0, (), [ (16393, 10, None, "IID('{00000564-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 2 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( u'RecordType' , u'ptype' , ), 11, (11, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( u'GetChildren' , u'pprset' , ), 12, (12, (), [ (16393, 10, None, "IID('{00000556-0000-0010-8000-00AA006D2EA4}')") , ], 1 , 1 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( u'Cancel' , ), 13, (13, (), [ ], 1 , 1 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
]

_Recordset_vtables_dispatch_ = 1
_Recordset_vtables_ = [
	(( u'Save' , u'Destination' , u'PersistFormat' , ), 1057, (1057, (), [ (12, 17, None, None) , 
			(3, 49, '0', None) , ], 1 , 1 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
]

_Stream_vtables_dispatch_ = 1
_Stream_vtables_ = [
	(( u'Size' , u'pSize' , ), 1, (1, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 28 , (3, 0, None, None) , 0 , )),
	(( u'EOS' , u'pEOS' , ), 2, (2, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 32 , (3, 0, None, None) , 0 , )),
	(( u'Position' , u'pPos' , ), 3, (3, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 36 , (3, 0, None, None) , 0 , )),
	(( u'Position' , u'pPos' , ), 3, (3, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 40 , (3, 0, None, None) , 0 , )),
	(( u'Type' , u'ptype' , ), 4, (4, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 44 , (3, 0, None, None) , 0 , )),
	(( u'Type' , u'ptype' , ), 4, (4, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 48 , (3, 0, None, None) , 0 , )),
	(( u'LineSeparator' , u'pLS' , ), 5, (5, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 52 , (3, 0, None, None) , 0 , )),
	(( u'LineSeparator' , u'pLS' , ), 5, (5, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( u'State' , u'pState' , ), 6, (6, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 60 , (3, 0, None, None) , 0 , )),
	(( u'Mode' , u'pMode' , ), 7, (7, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( u'Mode' , u'pMode' , ), 7, (7, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 68 , (3, 0, None, None) , 0 , )),
	(( u'Charset' , u'pbstrCharset' , ), 8, (8, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( u'Charset' , u'pbstrCharset' , ), 8, (8, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 76 , (3, 0, None, None) , 0 , )),
	(( u'Read' , u'NumBytes' , u'pval' , ), 9, (9, (), [ (3, 49, '-1', None) , 
			(16396, 10, None, None) , ], 1 , 1 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( u'Open' , u'Source' , u'Mode' , u'Options' , u'UserName' , 
			u'Password' , ), 10, (10, (), [ (12, 17, None, None) , (3, 49, '0', None) , (3, 49, '-1', None) , 
			(8, 49, "u''", None) , (8, 49, "u''", None) , ], 1 , 1 , 4 , 0 , 84 , (3, 32, None, None) , 0 , )),
	(( u'Close' , ), 11, (11, (), [ ], 1 , 1 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( u'SkipLine' , ), 12, (12, (), [ ], 1 , 1 , 4 , 0 , 92 , (3, 0, None, None) , 0 , )),
	(( u'Write' , u'Buffer' , ), 13, (13, (), [ (12, 1, None, None) , ], 1 , 1 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( u'SetEOS' , ), 14, (14, (), [ ], 1 , 1 , 4 , 0 , 100 , (3, 0, None, None) , 0 , )),
	(( u'CopyTo' , u'DestStream' , u'CharNumber' , ), 15, (15, (), [ (9, 1, None, "IID('{00000565-0000-0010-8000-00AA006D2EA4}')") , 
			(3, 49, '-1', None) , ], 1 , 1 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( u'Flush' , ), 16, (16, (), [ ], 1 , 1 , 4 , 0 , 108 , (3, 0, None, None) , 0 , )),
	(( u'SaveToFile' , u'FileName' , u'Options' , ), 17, (17, (), [ (8, 1, None, None) , 
			(3, 49, '1', None) , ], 1 , 1 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( u'LoadFromFile' , u'FileName' , ), 18, (18, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 116 , (3, 0, None, None) , 0 , )),
	(( u'ReadText' , u'NumChars' , u'pbstr' , ), 19, (19, (), [ (3, 49, '-1', None) , 
			(16392, 10, None, None) , ], 1 , 1 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( u'WriteText' , u'Data' , u'Options' , ), 20, (20, (), [ (8, 1, None, None) , 
			(3, 49, '0', None) , ], 1 , 1 , 4 , 0 , 124 , (3, 0, None, None) , 0 , )),
	(( u'Cancel' , ), 21, (21, (), [ ], 1 , 1 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
]

RecordMap = {
}

CLSIDToClassMap = {
	'{00000500-0000-0010-8000-00AA006D2EA4}' : Error,
	'{00000501-0000-0010-8000-00AA006D2EA4}' : Errors,
	'{00000503-0000-0010-8000-00AA006D2EA4}' : Property,
	'{00000504-0000-0010-8000-00AA006D2EA4}' : Properties,
	'{00000505-0000-0010-8000-00AA006D2EA4}' : Field15,
	'{00000506-0000-0010-8000-00AA006D2EA4}' : Fields15,
	'{00000507-0000-0010-8000-00AA006D2EA4}' : Command,
	'{00000508-0000-0010-8000-00AA006D2EA4}' : Command15,
	'{0000050B-0000-0010-8000-00AA006D2EA4}' : Parameter,
	'{0000050C-0000-0010-8000-00AA006D2EA4}' : _Parameter,
	'{0000050D-0000-0010-8000-00AA006D2EA4}' : Parameters,
	'{0000050E-0000-0010-8000-00AA006D2EA4}' : Recordset15,
	'{00000566-0000-0010-8000-00AA006D2EA4}' : Stream,
	'{00000512-0000-0010-8000-00AA006D2EA4}' : _Collection,
	'{00000513-0000-0010-8000-00AA006D2EA4}' : _DynaCollection,
	'{00000514-0000-0010-8000-00AA006D2EA4}' : Connection,
	'{00000515-0000-0010-8000-00AA006D2EA4}' : Connection15,
	'{B08400BD-F9D1-4D02-B856-71D5DBA123E9}' : _Command,
	'{00000534-0000-0010-8000-00AA006D2EA4}' : _ADO,
	'{00000535-0000-0010-8000-00AA006D2EA4}' : Recordset,
	'{0000054C-0000-0010-8000-00AA006D2EA4}' : Field20,
	'{0000054D-0000-0010-8000-00AA006D2EA4}' : Fields20,
	'{0000054E-0000-0010-8000-00AA006D2EA4}' : Command25,
	'{0000054F-0000-0010-8000-00AA006D2EA4}' : Recordset20,
	'{00000550-0000-0010-8000-00AA006D2EA4}' : _Connection,
	'{00000555-0000-0010-8000-00AA006D2EA4}' : Recordset21,
	'{00000556-0000-0010-8000-00AA006D2EA4}' : _Recordset,
	'{00000560-0000-0010-8000-00AA006D2EA4}' : Record,
	'{00000562-0000-0010-8000-00AA006D2EA4}' : _Record,
	'{00000564-0000-0010-8000-00AA006D2EA4}' : Fields,
	'{00000565-0000-0010-8000-00AA006D2EA4}' : _Stream,
	'{00000266-0000-0010-8000-00AA006D2EA4}' : RecordsetEvents,
	'{00000567-0000-0010-8000-00AA006D2EA4}' : ADORecordConstruction,
	'{00000568-0000-0010-8000-00AA006D2EA4}' : ADOStreamConstruction,
	'{00000569-0000-0010-8000-00AA006D2EA4}' : Field,
	'{00000400-0000-0010-8000-00AA006D2EA4}' : ConnectionEvents,
	'{00000283-0000-0010-8000-00AA006D2EA4}' : ADORecordsetConstruction,
}
CLSIDToPackageMap = {}
win32com.client.CLSIDToClass.RegisterCLSIDsFromDict( CLSIDToClassMap )
VTablesToPackageMap = {}
VTablesToClassMap = {
	'{00000500-0000-0010-8000-00AA006D2EA4}' : 'Error',
	'{00000501-0000-0010-8000-00AA006D2EA4}' : 'Errors',
	'{00000402-0000-0010-8000-00AA006D2EA4}' : 'ConnectionEventsVt',
	'{00000503-0000-0010-8000-00AA006D2EA4}' : 'Property',
	'{00000504-0000-0010-8000-00AA006D2EA4}' : 'Properties',
	'{00000505-0000-0010-8000-00AA006D2EA4}' : 'Field15',
	'{00000506-0000-0010-8000-00AA006D2EA4}' : 'Fields15',
	'{00000508-0000-0010-8000-00AA006D2EA4}' : 'Command15',
	'{0000050C-0000-0010-8000-00AA006D2EA4}' : '_Parameter',
	'{0000050D-0000-0010-8000-00AA006D2EA4}' : 'Parameters',
	'{0000050E-0000-0010-8000-00AA006D2EA4}' : 'Recordset15',
	'{00000512-0000-0010-8000-00AA006D2EA4}' : '_Collection',
	'{00000513-0000-0010-8000-00AA006D2EA4}' : '_DynaCollection',
	'{00000515-0000-0010-8000-00AA006D2EA4}' : 'Connection15',
	'{00000516-0000-0010-8000-00AA006D2EA4}' : 'ADOConnectionConstruction15',
	'{00000517-0000-0010-8000-00AA006D2EA4}' : 'ADOCommandConstruction',
	'{B08400BD-F9D1-4D02-B856-71D5DBA123E9}' : '_Command',
	'{00000534-0000-0010-8000-00AA006D2EA4}' : '_ADO',
	'{0000054C-0000-0010-8000-00AA006D2EA4}' : 'Field20',
	'{0000054D-0000-0010-8000-00AA006D2EA4}' : 'Fields20',
	'{0000054E-0000-0010-8000-00AA006D2EA4}' : 'Command25',
	'{0000054F-0000-0010-8000-00AA006D2EA4}' : 'Recordset20',
	'{00000550-0000-0010-8000-00AA006D2EA4}' : '_Connection',
	'{00000551-0000-0010-8000-00AA006D2EA4}' : 'ADOConnectionConstruction',
	'{00000555-0000-0010-8000-00AA006D2EA4}' : 'Recordset21',
	'{00000556-0000-0010-8000-00AA006D2EA4}' : '_Recordset',
	'{00000562-0000-0010-8000-00AA006D2EA4}' : '_Record',
	'{00000564-0000-0010-8000-00AA006D2EA4}' : 'Fields',
	'{00000565-0000-0010-8000-00AA006D2EA4}' : '_Stream',
	'{00000569-0000-0010-8000-00AA006D2EA4}' : 'Field',
	'{00000403-0000-0010-8000-00AA006D2EA4}' : 'RecordsetEventsVt',
}


NamesToIIDMap = {
	'_Connection' : '{00000550-0000-0010-8000-00AA006D2EA4}',
	'Errors' : '{00000501-0000-0010-8000-00AA006D2EA4}',
	'ADOStreamConstruction' : '{00000568-0000-0010-8000-00AA006D2EA4}',
	'Parameters' : '{0000050D-0000-0010-8000-00AA006D2EA4}',
	'Recordset20' : '{0000054F-0000-0010-8000-00AA006D2EA4}',
	'Recordset21' : '{00000555-0000-0010-8000-00AA006D2EA4}',
	'Field' : '{00000569-0000-0010-8000-00AA006D2EA4}',
	'ADORecordConstruction' : '{00000567-0000-0010-8000-00AA006D2EA4}',
	'ADOConnectionConstruction' : '{00000551-0000-0010-8000-00AA006D2EA4}',
	'ConnectionEventsVt' : '{00000402-0000-0010-8000-00AA006D2EA4}',
	'Field20' : '{0000054C-0000-0010-8000-00AA006D2EA4}',
	'_Record' : '{00000562-0000-0010-8000-00AA006D2EA4}',
	'Command15' : '{00000508-0000-0010-8000-00AA006D2EA4}',
	'Fields20' : '{0000054D-0000-0010-8000-00AA006D2EA4}',
	'ADOCommandConstruction' : '{00000517-0000-0010-8000-00AA006D2EA4}',
	'_DynaCollection' : '{00000513-0000-0010-8000-00AA006D2EA4}',
	'RecordsetEvents' : '{00000266-0000-0010-8000-00AA006D2EA4}',
	'_Stream' : '{00000565-0000-0010-8000-00AA006D2EA4}',
	'_Command' : '{B08400BD-F9D1-4D02-B856-71D5DBA123E9}',
	'Properties' : '{00000504-0000-0010-8000-00AA006D2EA4}',
	'RecordsetEventsVt' : '{00000403-0000-0010-8000-00AA006D2EA4}',
	'Recordset15' : '{0000050E-0000-0010-8000-00AA006D2EA4}',
	'Fields' : '{00000564-0000-0010-8000-00AA006D2EA4}',
	'_Recordset' : '{00000556-0000-0010-8000-00AA006D2EA4}',
	'Connection15' : '{00000515-0000-0010-8000-00AA006D2EA4}',
	'Field15' : '{00000505-0000-0010-8000-00AA006D2EA4}',
	'Error' : '{00000500-0000-0010-8000-00AA006D2EA4}',
	'Property' : '{00000503-0000-0010-8000-00AA006D2EA4}',
	'ADOConnectionConstruction15' : '{00000516-0000-0010-8000-00AA006D2EA4}',
	'Command25' : '{0000054E-0000-0010-8000-00AA006D2EA4}',
	'Fields15' : '{00000506-0000-0010-8000-00AA006D2EA4}',
	'_Parameter' : '{0000050C-0000-0010-8000-00AA006D2EA4}',
	'ADORecordsetConstruction' : '{00000283-0000-0010-8000-00AA006D2EA4}',
	'ConnectionEvents' : '{00000400-0000-0010-8000-00AA006D2EA4}',
	'_Collection' : '{00000512-0000-0010-8000-00AA006D2EA4}',
	'_ADO' : '{00000534-0000-0010-8000-00AA006D2EA4}',
}

win32com.client.constants.__dicts__.append(constants.__dict__)

