using System.ComponentModel.DataAnnotations;

namespace KAT.Camelot.Extensibility.Excel.AddIn.ConfigurationExport;

enum DataColumnType
{
	[Display( Name = "Data Field" )]
    DataField,
	[Display( Name = "Include" )]
    Include,
	[Display( Name = "Label" )]
    Label,
	[Display( Name = "Format" )]
    DisplayFormat,
	[Display( Name = "Type" )]
    DataType,
	[Display( Name = "Min" )]
    Min,
	[Display( Name = "Max" )]
    Max,
	[Display( Name = "Sort" )]
    Sort,
	[Display( Name = "Validation Error Type" )]
    ValidationErrorType,
	[Display( Name = "Required Error Type" )]
    RequiredErrorType,
	[Display( Name = "Validation Comments" )]
    ValidationComments,
	[Display( Name = "Skip Audit?" )]
    SkipAudit,
	[Display( Name = "Allowed Audit Variance" )]
    AllowedAuditVariance,
	[Display( Name = "Validation Expression" )]
    ValidationExpression,
	[Display( Name = "Validation Expression Message" )]
    ValidationExpressionMessage,
	[Display( Name = "Is Detail?" )]
    IsDetail,
	[Display( Name = "Display Width" )]
    DisplayWidth,
	[Display( Name = "MH Include" )]
    MadHatterInclude
}

enum ReportColumnType
{
	[Display( Name = "Name" )]
	ReportName,
	[Display( Name = "Report Category" )]
	ReportCategory,
	[Display( Name = "Spec Export ID" )]
	SpecExportID,
	[Display( Name = "Spec Export Custom Process" )]
	SpecExportCustomProcess,
	[Display( Name = "Evolution ID" )]
	EvolutionId,
	[Display( Name = "Description" )]
	Description,
	[Display( Name = "Include" )]
	Include,
	[Display( Name = "History Table Type" )]
	HistoryTableType,
	[Display( Name = "Filter" )]
	Filter,
	[Display( Name = "File Name" )]
	FileName,
	[Display( Name = "AdHoc FolderItem Type" )]
	FolderItemType,
	[Display( Name = "AdHoc Table Name" )]
	ResultTableName,
	[Display( Name = "AdHoc Sort Column" )]
	SortColumn,
	[Display( Name = "AdHoc Sort Direction" )]
	SortDirection,
	[Display( Name = "AdHoc Sort Type" )]
	SortType,
	[Display( Name = "Index Row Count Name" )]
	RowIndexName,
	[Display( Name = "Index Row Count Header" )]
	RowIndexHeader,
	[Display( Name = "Index Column Count Name" )]
	ColumnIndexName,
	[Display( Name = "FTP Url" )]
	FtpUrl,
	[Display( Name = "FTP User Name" )]
	FtpUserName,
	[Display( Name = "FTP Password" )]
	FtpPassword,
	[Display( Name = "FTP Notifications" )]
	FtpNotifications,
	[Display( Name = "Delimiter" )]
	Delimiter,
	[Display( Name = "Extension" )]
	Extension,
	[Display( Name = "Submit Page Path" )]
	SubmitPagePath,
	[Display( Name = "Index Row Lookup" )]
	RowIndexLookup
}

enum CalcInputColumnType
{
	[Display( Name = "Input" )]
    InputName,
	[Display( Name = "Type" )]
    InputType,
	[Display( Name = "Label" )]
    Label,
	[Display( Name = "Help" )]
    Help,
	[Display( Name = "Min" )]
    Min,
	[Display( Name = "Max" )]
    Max,
	[Display( Name = "Min Age" )]
    MinAge,
	[Display( Name = "Max Age" )]
    MaxAge,
	[Display( Name = "Required" )]
    Required,
	[Display( Name = "Css" )]
    Css,
	[Display( Name = "Triggers Calculation" )]
    TriggersCalculation,
	[Display( Name = "Visibility" )]
    Visibility,
	[Display( Name = "Message" )]
    Message,
	[Display( Name = "Default" )]
    DefaultValue,
	[Display( Name = "Compare To" )]
    CompareTo,
	[Display( Name = "Regular Expression" )]
    RegularExpression,
	[Display( Name = "Is Valid" )]
    IsValid
}