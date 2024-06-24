using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Humanizer;
using KAT.Camelot.Data.Repositories;
using KAT.Camelot.Domain;
using KAT.Camelot.Domain.Extensions;
using KAT.Camelot.Extensibility.Excel.AddIn.RBLe.Interop;
using MSExcel = Microsoft.Office.Interop.Excel;

namespace KAT.Camelot.Extensibility.Excel.AddIn.ConfigurationExport;

class Specification
{
	private MSExcel.Workbook workbook = null!;
	private string dateHireName = "date-hire";
	private string dateBirthName = "date-birth";

	public void Export( ExportSpecificationInfo info, MSExcel.Workbook workbook, XElement globalTables )
	{
		this.workbook = workbook;

		var planInfo = ExportPlanInfo();

		var version =
			workbook.GetWorksheet( "Plan Info" )?.RangeOrNull<string>( "Version" ) ??
			workbook.RangeOrNull<string>( "Version" ) ?? "0.1";

		var configProfile =
			new XElement( "Config",
				new XAttribute( "version", version ),
				new XElement( "xDataDef" ),
				new XElement( "ExcelExports" )
			);

		var specReports = InitializeReports( configProfile );

		ExportProfile( configProfile );

		var configLookups = ExportLookups();

		var calcInputLayouts = GetCalcInputLayouts( specReports, configLookups );

		XElement? codeGenLookups = null;

		foreach ( var location in info.Locations.Where( l => l.Selected ) )
		{
			var planInfoPath = Path.Combine( location.Location, "Config-PlanInfo.xml" );
			var planInfoXml = File.Exists( planInfoPath ) ? XElement.Load( planInfoPath ) : new XElement( "PlanInfo", new XElement( "Plan", new XAttribute( "id", "default" ) ) );

			planInfoXml.Elements( "Plan" ).Elements().Where( e => !e.HasElements ).Remove();

			foreach ( var pi in planInfo.Elements() )
			{
				var existing = planInfoXml.Elements( "Plan" ).Elements( pi.Name.LocalName ).FirstOrDefault( e => (string?)e.Attribute( "id" ) == (string?)pi.Attribute( "id" ) );
				if ( existing != null )
				{
					existing.ReplaceWith( pi );
				}
				else
				{
					planInfoXml.Element( "Plan" )!.Add( pi );
				}
			}

			planInfoXml.SaveIndented( planInfoPath );

			if ( configLookups != null )
			{
				var lookupsPath = Path.Combine( location.Location, "Config-Lookups.xml" );
				var configLookupsXml = File.Exists( lookupsPath ) ? XElement.Load( lookupsPath ) : new XElement( "DataTableDefs" );

				foreach ( var l in configLookups.Elements() )
				{
					var existing = configLookupsXml.Elements( "DataTable" ).FirstOrDefault( e => (string?)e.Attribute( "id" ) == (string?)l.Attribute( "id" ) );
					if ( existing != null )
					{
						existing.ReplaceWith( l );
					}
					else
					{
						configLookupsXml.Add( l );
					}
				}
				configLookupsXml.SaveIndented( lookupsPath );

				if ( codeGenLookups == null )
				{
					codeGenLookups = new XElement( configLookupsXml );
					xDSRepository.MergeGlobalTables( codeGenLookups, globalTables );
				}
			}

			configProfile.SaveIndented( Path.Combine( location.Location, "Config-Profile.xml" ) );

			MergeEvolutionReports( location.Location, specReports, calcInputLayouts );

			var isAdmin = location.Location.Contains( @"\Evolution\Websites\Admin\", StringComparison.InvariantCultureIgnoreCase );
			if (
				 ( isAdmin && !location.Location.Contains( @"\Hangfire", StringComparison.InvariantCultureIgnoreCase ) ) ||
				 location.Location.Contains( @"\Evolution\Websites\ESS\", StringComparison.InvariantCultureIgnoreCase )
			)
			{
				var generatedPath = Path.Combine( new DirectoryInfo( location.Location ).Parent!.FullName, "Generated", "xDSModel.generated.cs" );
				GenerateDataModels( generatedPath, configProfile, codeGenLookups! );

				if ( isAdmin )
				{
					generatedPath = Path.Combine( new DirectoryInfo( location.Location ).Parent!.FullName, "Generated", "Calculations.generated.cs" );
					GenerateCalculationModels( generatedPath, configProfile, codeGenLookups! );
					GenerateCalculationViews( location.Location, calcInputLayouts, codeGenLookups! );
				}
			}
		}
	}

	private static void GenerateDataModels( string generatedPath, XElement configProfile, XElement codeGenLookups )
	{
		Directory.CreateDirectory( Path.GetDirectoryName( generatedPath )! );

		var isAdmin = generatedPath.Contains( @"\Admin\", StringComparison.InvariantCultureIgnoreCase );

		var codeGen = new StringBuilder();
		codeGen.AppendLine( "// GENERATED FILE -- DO NOT MANUALLY EDIT THIS FILE" );
		codeGen.AppendLine( $"// Specification Sheet Version: {(string)configProfile.Attribute( "version" )!}" );
		codeGen.AppendLine( "" );

		codeGen.AppendLine( "using BTR.Evolution.Core;" );
		codeGen.AppendLine( "using BTR.Evolution.MadHatter.Resources;" );
		codeGen.AppendLine( "using System;" );
		codeGen.AppendLine( "using System.Collections.Generic;" );
		codeGen.AppendLine( "using System.Data;" );
		codeGen.AppendLine( "using System.Web;" );
		codeGen.AppendLine( "using System.Xml.Linq;" );
		codeGen.AppendLine( "using System.Linq;" );
		codeGen.AppendLine( "" );

		codeGen.AppendLine( isAdmin ? "namespace Administration" : "namespace Modeler" );
		codeGen.AppendLine( "{" );

		codeGen.AppendLine( "\tpublic partial class xDSDataModel" );
		codeGen.AppendLine( "\t{" );
		codeGen.AppendLine( "\t\tXElement profileXml;" );
		codeGen.AppendLine( "\t\tpublic xDSDataModel( XElement profileXml, LookupTableDelegates lookupTableDelegates )" );
		codeGen.AppendLine( "\t\t{" );
		codeGen.AppendLine( "\t\t\tthis.profileXml = profileXml;" );
		codeGen.AppendLine( "\t\t\tProfile = new ProfileModel( profileXml?.Profile() );" );
		codeGen.AppendLine( "\t\t\tLookupTables = new LookupTables( lookupTableDelegates );" );
		codeGen.AppendLine( "\t\t}" );
		codeGen.AppendLine( "" );

		codeGen.AppendLine( "\t\tpublic LookupTables LookupTables { get; }" );
		codeGen.AppendLine( "\t\tpublic ProfileModel Profile { get; }" );
		codeGen.AppendLine( "" );

		var textInfo = new CultureInfo( "en-US" ).TextInfo;

		foreach ( var h in configProfile.Elements( "xDataDef" ).Elements( "HistoryData" ) )
		{
			var historyType = ( (string)h.Attribute( "type" )! ).Replace( " ", "" );
			var properHistoryType = char.ToUpper( historyType[ 0 ] ) + historyType[ 1.. ];

			// codeGen.AppendLine($"\t\tpublic IEnumerable<{historyType}Model> {pluralize(historyType)} {{ get; private set; }}");
			codeGen.AppendLine( $"\t\tpublic {properHistoryType}Model {properHistoryType}( string index ) => new[] {{ profileXml?.GetHistoryRowByPosition( \"{historyType}\", index ) ?? profileXml?.GetHistoryRows( \"{historyType}\", h => (string)h.Attribute( \"hisIndex\" ) == index ).FirstOrDefault() }}.Where( r => r != null ).Select( r => new {properHistoryType}Model( r ) ).FirstOrDefault();" );
			codeGen.AppendLine( $"\t\tpublic IEnumerable<{properHistoryType}Model> {Pluralize( properHistoryType )}" );
			codeGen.AppendLine( "\t\t{" );
			codeGen.AppendLine( "\t\t\tget" );
			codeGen.AppendLine( "\t\t\t{" );
			codeGen.AppendLine( $"\t\t\t\tforeach ( var r in profileXml.HistoryItems( \"{historyType}\" ) ) yield return new {properHistoryType}Model( r );" );
			codeGen.AppendLine( "\t\t\t}" );
			codeGen.AppendLine( "\t\t}" );
			codeGen.AppendLine( "" );
		}

		codeGen.AppendLine( "\t}" );
		codeGen.AppendLine( "" );

		var profileFields = configProfile.Elements( "xDataDef" ).Elements( "Profile" ).Elements().Where( e => e.Name.LocalName != "h" ).ToArray();
		codeGen.AppendLine( "\tpublic partial class ProfileModel" );
		codeGen.AppendLine( "\t{" );
		codeGen.AppendLine( "\t\tXElement profileXml;" );
		codeGen.AppendLine( "" );
		codeGen.AppendLine( "\t\tpublic ProfileModel( XElement profileXml )" );
		codeGen.AppendLine( "\t\t{" );
		codeGen.AppendLine( "\t\t\tthis.profileXml = profileXml;" );
		codeGen.AppendLine( "\t\t}" );
		codeGen.AppendLine( "" );
		codeGen.AppendLine( "\t\tprivate void SetPropertyValue( string fieldName, object value )" );
		codeGen.AppendLine( "\t\t{" );
		codeGen.AppendLine( "\t\t\tif ( profileXml == null ) throw new NullReferenceException( \"profileXml is null\" );" );
		codeGen.AppendLine( "\t\t\tprofileXml.Elements( fieldName ).Remove();" );
		codeGen.AppendLine( "\t\t\tif ( value == null ) return;" );
		codeGen.AppendLine( "\t\t\tprofileXml.Add( new XElement( fieldName, value.GetType() == typeof( DateTime ) ? ( (DateTime)value ).ToString( \"yyyy-MM-dd\" ) : value ) );" );
		codeGen.AppendLine( "\t\t}" );
		codeGen.AppendLine( "" );
		codeGen.AppendLine( "\t\tpublic virtual int? DatabaseKey => (int?)profileXml.Parent.Attribute( \"pKey\" );" );
		codeGen.AppendLine( "\t\tpublic virtual DateTime? DatabaseUpdated => (DateTime?)profileXml.Parent.Attribute( \"date-updated\" );" );
		codeGen.AppendLine( "\t\tpublic virtual DateTime? DatabaseCreated => (DateTime?)profileXml.Parent.Attribute( \"date-created\" );" );
		codeGen.AppendLine( "" );

		void appendFieldProperty( XElement f, string modelName )
		{
			var fieldName = SafeMemberName( f.Name.LocalName );
			var propertyType = GetPropertyType( f );
			var isNullable = propertyType.EndsWith( "?" ) || propertyType == "string";
			var format = (string?)f.Attribute( "format" ) ?? "";

			/*
			codeGen.AppendLine($"\t\t[XmlName( \"{f.Name.LocalName}\" )]");
			if (!string.IsNullOrEmpty(format))
			{
				codeGen.AppendLine($"\t\t[DisplayFormat( DataFormatString = \"{format}\" )]");
			}
			*/

			if ( propertyType.EndsWith( "Lookup" ) || propertyType.EndsWith( "Lookup?" ) )
			{
				codeGen.AppendLine( $"\t\tpublic virtual {propertyType} {fieldName}" );
				codeGen.AppendLine( "\t\t{" );
				codeGen.AppendLine( $"\t\t\tget => EnumKeyAttribute.GetEnumValue<{propertyType}>( (string){modelName}?.Element( \"{f.Name.LocalName}\" ) );" );
				codeGen.AppendLine( $"\t\t\tset => SetPropertyValue( \"{f.Name.LocalName}\", value.Key() );" );
				codeGen.AppendLine( "\t\t}" );
				codeGen.AppendLine( $"\t\tpublic virtual void Set{fieldName}Key( string key ) => {fieldName} = EnumKeyAttribute.GetEnumValue<{propertyType.Replace( "?", "" )}>( key );" );
			}
			else
			{
				codeGen.AppendLine( $"\t\tpublic virtual {propertyType} {fieldName}" );
				codeGen.AppendLine( "\t\t{" );
				codeGen.AppendLine( $"\t\t\tget => ({propertyType}){modelName}?.Element( \"{f.Name.LocalName}\" );" );
				codeGen.AppendLine( $"\t\t\tset => SetPropertyValue( \"{f.Name.LocalName}\", value );" );
				codeGen.AppendLine( "\t\t}" );
			}
		}

		foreach ( var fld in profileFields )
		{
			appendFieldProperty( fld, "profileXml" );
		}

		codeGen.AppendLine( "\t}" );
		codeGen.AppendLine( "" );

		foreach ( var h in configProfile.Elements( "xDataDef" ).Elements( "HistoryData" ) )
		{
			var historyType = ( (string)h.Attribute( "type" )! ).Replace( " ", "" );
			var properHistoryType = char.ToUpper( historyType[ 0 ] ) + historyType[ 1.. ];
			codeGen.AppendLine( $"\tpublic partial class {properHistoryType}Model" );
			codeGen.AppendLine( "\t{" );
			codeGen.AppendLine( "\t\tXElement row;" );
			codeGen.AppendLine( "" );
			codeGen.AppendLine( $"\t\tpublic {properHistoryType}Model( XElement row )" );
			codeGen.AppendLine( "\t\t{" );
			codeGen.AppendLine( "\t\t\tthis.row = row;" );
			codeGen.AppendLine( "\t\t}" );
			codeGen.AppendLine( "" );
			codeGen.AppendLine( "\t\tpublic virtual XElement GetElement() => row;" );
			codeGen.AppendLine( "\t\tpublic virtual int? DatabaseKey => (int?)row.Attribute( \"id\" );" );
			codeGen.AppendLine( "\t\tpublic virtual DateTime? DatabaseUpdated => (DateTime?)row.Attribute( \"hisDateUpdated\" );" );
			codeGen.AppendLine( "\t\tpublic virtual DateTime? DatabaseCreated => (DateTime?)row.Attribute( \"hisDateCreated\" );" );
			codeGen.AppendLine( "" );
			codeGen.AppendLine( "\t\tprivate void SetPropertyValue( string fieldName, object value )" );
			codeGen.AppendLine( "\t\t{" );
			codeGen.AppendLine( "\t\t\tif ( row == null ) throw new NullReferenceException( \"row is null\" );" );
			codeGen.AppendLine( "\t\t\trow.Elements( fieldName ).Remove();" );
			codeGen.AppendLine( "\t\t\tif ( value == null ) return;" );
			codeGen.AppendLine( "\t\t\trow.Add( new XElement( fieldName, value.GetType() == typeof( DateTime ) ? ( (DateTime)value ).ToString( \"yyyy-MM-dd\" ) : typeof( YearHistoryIndex ).IsAssignableFrom( value.GetType() ) ? value.ToString() : value ) );" );
			codeGen.AppendLine( "\t\t}" );
			codeGen.AppendLine( "" );

			foreach ( var fld in h.Elements() )
			{
				appendFieldProperty( fld, "row" );
			}

			codeGen.AppendLine( "\t}" );
			codeGen.AppendLine( "" );
		}

		var lookupTablesUsed =
			configProfile.Elements( "xDataDef" ).Elements( "Profile" ).Elements().Concat(
				configProfile.Elements( "xDataDef" ).Elements( "HistoryData" ).Elements()
			)
			.Select( f => (string?)f.Attribute( "lookuptable" ) )
			.Where( t => !string.IsNullOrEmpty( t ) )
			.Distinct();

		var lookupTablesToProcess = 
			codeGenLookups.Elements( "DataTable" )
				.Where( t => lookupTablesUsed.Contains( (string)t.Attribute( "id" )! ) );

		if ( lookupTablesToProcess.Any() )
		{
			codeGen.AppendLine( $"\tpublic partial class LookupTables" );
			codeGen.AppendLine( "\t{" );

			foreach ( var t in lookupTablesToProcess )
			{
				var name = Pluralize( ( (string)t.Attribute( "id" )! )[ 5.. ] );
				codeGen.AppendLine( $"\t\tpublic {name} {name} {{ get; }}" );
			}
			codeGen.AppendLine( "" );
			codeGen.AppendLine( "\t\tpublic LookupTables( LookupTableDelegates lookupTableDelegates )" );
			codeGen.AppendLine( "\t\t{" );
			foreach ( var t in lookupTablesToProcess )
			{
				var name = Pluralize( ( (string)t.Attribute( "id" )! )[ 5.. ] );
				codeGen.AppendLine( $"\t\t\t{name} = new {name}( lookupTableDelegates );" );
			}
			codeGen.AppendLine( "\t\t}" );
			codeGen.AppendLine( "\t}" );
			codeGen.AppendLine( "" );

			foreach ( var t in lookupTablesToProcess )
			{
				var tableName = (string)t.Attribute( "id" )!;
				var tableType = string.Concat( tableName.AsSpan( 5 ), "Lookup" );
				var name = Pluralize( tableName[ 5.. ] );
				codeGen.AppendLine( $"\tpublic partial class {name} : LookupTableBase" );
				codeGen.AppendLine( "\t{" );
				codeGen.AppendLine( $"\t\tpublic {name}( LookupTableDelegates lookupTableDelegates ) : base( lookupTableDelegates, \"{tableName}\" ) {{ }}" );
				codeGen.AppendLine( $"\t\tpublic string GetColumnText( {tableType} key, string column ) => GetColumnText( key.Key(), column );" );
				codeGen.AppendLine( $"\t\tpublic string GetColumnText( string key, string column ) => LookupTableDelegates.GetConfigLookupText( \"{tableName}\", key, column );" );
				codeGen.AppendLine( "\t}" );
				codeGen.AppendLine( "" );
			}

			codeGen.AppendLine( @"	public partial class LookupTableBase
	{
		string tableName;
		XElement tableXml;
		DataSet tableDataSet;
		DataSet tableDataSetWithBlank;

		protected LookupTableDelegates LookupTableDelegates { get; private set; }

		public LookupTableBase( LookupTableDelegates lookupTableDelegates, string tableName ) 
		{
			this.LookupTableDelegates = lookupTableDelegates;
			this.tableName = tableName;
		}

		public virtual XElement ToXml() => tableXml ?? ( tableXml = LookupTableDelegates.GetConfigLookupTable( tableName ) );
		public virtual DataSet ToDataSet( bool withBlank )
		{
			if ( withBlank ) return tableDataSetWithBlank ?? ( tableDataSetWithBlank = LookupTableDelegates.GetDataTableWithBlank( tableName ) );
			else return tableDataSet ?? ( tableDataSet = LookupTableDelegates.GetDataTable( tableName ) );
		}
	}

	public class LookupTableDelegates
	{
		public Func<string, XElement> GetConfigLookupTable { get; set; }
		public Func<string, DataSet> GetDataTableWithBlank { get; set; }
		public Func<string, DataSet> GetDataTable { get; set; }
		public Func<string, string, string, string> GetConfigLookupText { get; set; }
	}
" );
			codeGen.AppendLine( "" );
		}

		foreach ( var t in lookupTablesToProcess )
		{
			var tableType = string.Concat( ( (string)t.Attribute( "id" )! ).AsSpan( 5 ), "Lookup" );

			codeGen.AppendLine( $"\tpublic enum {tableType}" );
			codeGen.AppendLine( "\t{" );

			var values =
				t.Elements( "Table" ).Elements().Select( f => $"[EnumKey( \"{(string)f.Attribute( "key" )!}\" )] " + SafeMemberName( (string)f.Attribute( "name" )! + "_" + (string)f.Attribute( "key" )! ) );

			codeGen.AppendLine( "\t\t" + string.Join( "," + Environment.NewLine + "\t\t", values ) );

			codeGen.AppendLine( "\t}" );
			codeGen.AppendLine( "" );
		}

		codeGen.AppendLine( "\tpublic static class LookupExtensions" );
		codeGen.AppendLine( "\t{" );

		foreach ( var t in lookupTablesToProcess )
		{
			var tableType = string.Concat( ( (string)t.Attribute( "id" )! ).AsSpan( 5 ), "Lookup" );
			var values = t.Elements( "Table" ).Elements().Select( f => new { Key = (string)f.Attribute( "key" )!, Value = SafeMemberName( (string)f.Attribute( "name" )! + "_" + (string)f.Attribute( "key" )! ) } );
			var texts = t.Elements( "Table" ).Elements().Select( f => new { Key = SafeMemberName( (string)f.Attribute( "name" )! + "_" + (string)f.Attribute( "key" )! ), Value = (string)f.Attribute( "name" )! } );

			codeGen.AppendLine( $"\t\tpublic static string Key ( this {tableType} value )" );
			codeGen.AppendLine( "\t\t{" );

			foreach ( var v in values )
			{
				codeGen.AppendLine( $"\t\t\tif ( value == {tableType}.{v.Value} ) return \"{v.Key}\";" );
			}
			codeGen.AppendLine( "\t\t\tthrow new ArgumentOutOfRangeException();" );
			codeGen.AppendLine( "\t\t}" );
			codeGen.AppendLine( $"\t\tpublic static string Key ( this {tableType}? value ) => value == null ? null : value.Value.Key();" );
			codeGen.AppendLine( "" );

			codeGen.AppendLine( $"\t\tpublic static string Text ( this {tableType} value )" );
			codeGen.AppendLine( "\t\t{" );

			foreach ( var v in texts )
			{
				codeGen.AppendLine( $"\t\t\tif ( value == {tableType}.{v.Key} ) return \"{v.Value}\";" );
			}
			codeGen.AppendLine( "\t\t\tthrow new ArgumentOutOfRangeException();" );
			codeGen.AppendLine( "\t\t}" );
			codeGen.AppendLine( $"\t\tpublic static string Text ( this {tableType}? value ) => value == null ? null : value.Value.Text();" );
			codeGen.AppendLine( "" );
		}

		codeGen.AppendLine( "\t}" );

		codeGen.AppendLine( "}" );

		void generateModelProperty()
		{
			codeGen.AppendLine( @"		public xDSDataModel xDSDataModel
		{
			get
			{
				if ( ProfileXml == null ) return null;
				var model = HttpContext.Current.Items[ nameof( xDSDataModel ) ] as xDSDataModel;

				if ( model == null )
				{
					HttpContext.Current.Items[ nameof( xDSDataModel ) ] = model = 
						new xDSDataModel(
							this.ProfileXml,
							new LookupTableDelegates
							{
								GetConfigLookupTable = this.GetConfigLookupTable,
								GetConfigLookupText = this.GetConfigLookupText,
								GetDataTable = this.GetDataTable,
								GetDataTableWithBlank = this.GetDataTableWithBlank
							}
						);
				}
				return model;
			}
		}" );
		}

		codeGen.AppendLine( "" );
		codeGen.AppendLine( isAdmin ? "namespace Administration" : "namespace Modeler" );
		codeGen.AppendLine( "{" );
		codeGen.AppendLine( isAdmin ? "\tpublic partial class xDSHelper : BTR.Evolution.MadHatter.Administration.Web.xDSHelper" : "\tpublic partial class xDSHelper : BTR.Evolution.MadHatter.Web.xDSHelper" );
		codeGen.AppendLine( "\t{" );
		codeGen.AppendLine( "\t\tpublic xDSHelper( BTR.Evolution.MadHatter.Web.PropertyBag propertyBag, Locale locale ) : base( propertyBag, locale ) { }" );
		codeGen.AppendLine( "\t\tpublic xDSHelper( BTR.Evolution.MadHatter.Web.PropertyBag propertyBag, Locale locale, Action<BTR.Evolution.MadHatter.Web.PropertyBag, XElement> updateProfile, string[] profileDetailTypes ) : base( propertyBag, locale, updateProfile, profileDetailTypes ) { }" );
		codeGen.AppendLine( "" );

		codeGen.AppendLine( "\t\tpublic override void SetProfileXml( XElement profileXml )" );
		codeGen.AppendLine( "\t\t{" );
		codeGen.AppendLine( "\t\t\tbase.SetProfileXml( profileXml );" );
		codeGen.AppendLine( "" );
		codeGen.AppendLine( "\t\t\t// Clear out so it is rebuilt" );
		codeGen.AppendLine( "\t\t\tHttpContext.Current.Items[ nameof( xDSDataModel ) ] = null;" );
		codeGen.AppendLine( "\t\t}" );
		codeGen.AppendLine( "" );

		generateModelProperty();

		codeGen.AppendLine( "\t}" );
		codeGen.AppendLine( "" );

		codeGen.AppendLine( "\tpublic partial class _Default" );
		codeGen.AppendLine( "\t{" );
		codeGen.AppendLine( "\t\tprotected new xDSHelper xDSHelper => base.xDSHelper as xDSHelper;" );
		// generateModelProperty( false );
		codeGen.AppendLine( "\t}" );
		codeGen.AppendLine( "}" );
		codeGen.AppendLine( "" );

		var controlTypes = isAdmin
			? new[] { "EvolutionControl", "RBLeControl", "RBLeServerSubmitControl" }
			: new[] { "EvolutionControl", "SeveranceControl", "RBLeControl", "RBLeServerSubmitControl", "RBLeRequiredDocumentElectionControl", "RBLeSaveInputsControl", "RBLeElectionControl" };

		codeGen.AppendLine( isAdmin ? "namespace Administration.Controls" : "namespace Modeler.Controls" );
		codeGen.AppendLine( "{" );

		foreach ( var ct in controlTypes )
		{
			codeGen.AppendLine( $"\tpublic partial class {ct} : BTR.Evolution.MadHatter{( isAdmin ? ".Administration" : "" )}.Web.{ct}" );
			codeGen.AppendLine( "\t{" );
			codeGen.AppendLine( "\t\tprotected new xDSHelper xDSHelper => base.xDSHelper as xDSHelper;" );
			// generateModelProperty( false );
			codeGen.AppendLine( "\t}" );
		}

		codeGen.AppendLine( "}" );

		using var f = new StreamWriter( generatedPath );
		f.WriteLine( codeGen.ToString() );
	}

	private static void GenerateCalculationModels( string generatedPath, XElement configProfile, XElement codeGenLookups )
	{
		Directory.CreateDirectory( Path.GetDirectoryName( generatedPath )! );

		var adminCalcTypes =
			codeGenLookups
				.Elements( "DataTable" )
				.Where( t => (string)t.Attribute( "id" )! == "TableMHACalculationTypes" )
				.Elements( "Table" )
				.Elements( "TableItem" )
				.ToArray();

		var summaryResultTypes =
			string.Join(
				Environment.NewLine,
				adminCalcTypes.Select( t => $"\t\tpublic int {SafeMemberName( (string)t.Attribute( "name" )! )} {{ get; set; }}" )
			);

		var summaryHeaders =
			string.Join(
				Environment.NewLine,
				adminCalcTypes.Select( t => $"\t\t\tyield return CreateHeader( \"{(string)t.Attribute( "FolderItemType" )!}\", \"{(string)t.Attribute( "name" )!}\", \"text-right\", nameof( CalculationSummaryResult.{SafeMemberName( (string)t.Attribute( "name" )! )} ), false );" )
			);

		var summaryColumns =
			string.Join(
				Environment.NewLine,
				adminCalcTypes.Select( t => $"\t\t\tyield return new HtmlGenericControl( tag, \"text-right\" ) {{ InnerHtml = row.{SafeMemberName( (string)t.Attribute( "name" )! )}.ToString( \"N0\" ) }};" )
			);

		var participantTotals =
			string.Join(
				Environment.NewLine,
				adminCalcTypes.Select( t => $"\t\t\t\t\t\t\t{SafeMemberName( (string)t.Attribute( "name" )! )} = GetParticipantTotal( result, \"{(string)t.Attribute( "FolderItemType" )!}\" )," )
			);

		var calculationTotals =
			string.Join(
				Environment.NewLine,
				adminCalcTypes.Select( t => $"\t\t\t\t\t\t{SafeMemberName( (string)t.Attribute( "name" )! )} = GetCalculationTotal( calculationItems, \"{(string)t.Attribute( "FolderItemType" )!}\" )," )
			);

		var historyAccessors =
			string.Join(
				Environment.NewLine,
				configProfile.Elements( "xDataDef" ).Elements( "HistoryData" )
					.Select( h => ( (string)h.Attribute( "type" )! ).Replace( " ", "" ) )
					.Select( h => char.ToUpper( h[ 0 ] ) + h[ 1.. ] )
					.Select( h =>
						$"\t\tprotected {h}Model {h}( string index ) => xDSHelper.xDSDataModel.{h}( index );" + Environment.NewLine +
						$"\t\tprotected IEnumerable<{h}Model> {Pluralize( h )} => xDSHelper.xDSDataModel.{Pluralize( h )};"
					)
			);

		var csContent =
			Ribbon.ExtractStringResource( "Evolution.Admin.Calculations.cs" )
				.Replace( "{CodeGenHeader}", $"// GENERATED FILE -- DO NOT MANUALLY EDIT THIS FILE{Environment.NewLine}// Specification Sheet Version: {(string)configProfile.Attribute( "version" )!}{Environment.NewLine}" )
				.Replace( "{HistoryAccessors}", historyAccessors )
				.Replace( "{SummaryResultTypes}", summaryResultTypes )
				.Replace( "{SummaryHeaders}", summaryHeaders )
				.Replace( "{SummaryColumns}", summaryColumns )
				.Replace( "{CalculationTypes}", string.Join( ", ", adminCalcTypes.Select( t => $"\"{(string)t.Attribute( "FolderItemType" )!}\"" ) ) )
				.Replace( "{ParticipantTotals}", participantTotals )
				.Replace( "{CalculationTotals}", calculationTotals );

		using var f = new StreamWriter( generatedPath );
		f.WriteLine( csContent );
	}

	private static void GenerateCalculationViews( string currentPath, XElement? calcInputLayouts, XElement codeGenLookups )
	{
		var configurationPath = Path.Combine( currentPath, "Configuration.xml" );

		if ( !( calcInputLayouts?.HasElements ?? false ) || !File.Exists( configurationPath ) ) return;

		var destinationFolder = Path.Combine( new DirectoryInfo( currentPath ).Parent!.FullName, "Generated" );
		Directory.CreateDirectory( destinationFolder );

		XNamespace ns = "http://schemas.benefittech.com/evolution/site";
		var configuration = XElement.Load( configurationPath );

		var ascxTemplate = Ribbon.ExtractStringResource( "Inputs.ascx" );
		var resxTemplate = Ribbon.ExtractStringResource( "Inputs.ascx_resx" );
		var codeBehindTemplate = Ribbon.ExtractStringResource( "Inputs.ascx_cs" );
		var designerTemplate = Ribbon.ExtractStringResource( "Inputs.ascx_designer_cs" );

		var groupTemplate =
@"    <div class=""bs-callout bs-callout-primary{FirstTemplate} vGroup{GroupNumber}"">
        <h5 class=""lGroup{GroupNumber}"">{Label}</h5>
        <div class=""row"">{Inputs}        </div>
    </div>";

		var profileCalculations =
			configuration
				.Elements( ns + "Navigation" )
				.Elements( ns + "Section" ).Where( s => (string?)s.Attribute( "ID" ) == "ParticipantTab" || (string?)s.Attribute( "Name" ) == "Participant" )
				.Elements( ns + "Chapter" )
				.Elements( ns + "Page" ).Where( p => (string?)p.Attribute( "Type" ) == "Listings.Calculations" );

		var calculatedLoads =
			configuration
				.Elements( ns + "Navigation" )
				.Elements( ns + "Section" )
				.Elements( ns + "Chapter" )
				.Elements( ns + "Page" ).Where( p => (string?)p.Attribute( "Type" ) == "Processes.CalculatedDataLoads" );

		var batchCalculations =
			configuration
				.Elements( ns + "Navigation" )
				.Elements( ns + "Section" )
				.Elements( ns + "Chapter" )
				.Elements( ns + "Page" ).Where( p => (string?)p.Attribute( "Type" ) == "Processes.Calculations" );

		var docGenPackages =
			configuration
				.Elements( ns + "Navigation" )
				.Elements( ns + "Section" )
				.Elements( ns + "Chapter" )
				.Elements( ns + "Page" ).Where( p => (string?)p.Attribute( "Type" ) == "Processes.DocumentPackages" );

		var messages = new List<string>();

		void saveFile( string content, string fileName )
		{
			var destination = Path.Combine( destinationFolder, fileName );
			File.WriteAllText( destination, content );
		}

		foreach ( var layout in calcInputLayouts.Elements( "Layout" ) )
		{
			var layoutKey = (string)layout.Attribute( "Key" )!;
			var layoutType = (string)layout.Attribute( "Type" )!;
			var isProfileCalculation = layoutType == "iMHACalcType";
			var layoutTable =
				isProfileCalculation ? new[] { "TableMHACalculationTypes" } :
				layoutType == "iProcessType" ? new[] { "TableRBLProcessTypes", "TableRBLCalculationTypes" } :
				layoutType == "iPackageType" ? new[] { "TableDocTypes" } :
				layoutType == "iReportType" ? null : new[] { "NotSupportedException" };

			if ( layoutTable?[ 0 ] == "NotSupportedException" )
			{
				throw new NotSupportedException( layoutType + " layout type not supported for Calc Inputs." );
			}

			var layoutInfo = layoutTable == null
				? new { Name = (string?)layout.Attribute( "ReportName" ) ?? layoutKey, Type = "iReportType" }
				: codeGenLookups
					.Elements( "DataTable" ).Where( t => layoutTable.Contains( (string?)t.Attribute( "id" ) ) )
					.Elements( "Table" )
					.Elements( "TableItem" )
						.Where( i => ( isProfileCalculation ? (string?)i.Attribute( "FolderItemType" ) : (string?)i.Attribute( "key" ) ) == layoutKey )
						.Select( i => new { Name = (string)i.Attribute( "name" )!, Type = (string)i.Parent!.Parent!.Attribute( "id" )! } )
						.Single();

			var viewContainer =
				layoutInfo.Type == "TableMHACalculationTypes" ? profileCalculations :
				layoutInfo.Type == "TableRBLProcessTypes" ? calculatedLoads :
				layoutInfo.Type == "TableRBLCalculationTypes" ? batchCalculations :
				layoutInfo.Type == "TableDocTypes" ? docGenPackages : null;

			var calculationType = SafeMemberName( layoutInfo.Name );

			var className =
				layoutType == "iMHACalcType" ? $"Calculation{calculationType}" :
				layoutType == "iProcessType" ? $"BatchCalculation{calculationType}" :
				layoutType == "iPackageType" ? $"DocGen{calculationType}" :
				layoutType == "iReportType" ? $"Report{layoutKey}" : $"Modeling{calculationType}";

			if ( layoutTable != null && ( viewContainer?.Any() ?? false ) )
			{
				var view = viewContainer.Elements( ns + "Views" ).Elements( ns + "View" ).FirstOrDefault( v => (string?)v.Attribute( "Type" ) == layoutKey );
				/*
					<Views>
						<View Type="auth" Edit="Generated\BatchCalculationTVLSStatusChange.Generated.ascx" />
					</Views>
				*/

				if ( layout.HasElements )
				{
					// Add View
					var firstContainer = viewContainer.First();
					var views = firstContainer.Elements( ns + "Views" ).FirstOrDefault();
					if ( views == null )
					{
						viewContainer.First().Add( views = new XElement( ns + "Views" ) );
					}

					if ( view == null )
					{
						views.Add( view = new XElement( ns + "View", new XAttribute( "Type", layoutKey ), new XAttribute( "Edit", $@"Generated\{className}.Generated.ascx" ) ) );
						messages.Add( $@"{layoutInfo.Type}.{layoutKey} added to {(string)firstContainer.Parent!.Parent!.Attribute( "Name" )!}\{(string)firstContainer.Parent.Attribute( "Name" )!}\{(string)firstContainer.Parent.Attribute( "Name" )!}" );
					}
					else if (
						( (string)view.Attribute( "Edit" )! ).StartsWith( @"Generated\" ) &&
						(string)view.Attribute( "Edit" )! != $@"Generated\{className}.Generated.ascx"
					)
					{
						view.Attribute( "Edit" )!.Value = $@"Generated\{className}.Generated.ascx";
						messages.Add( $@"{layoutInfo.Type}.{layoutKey} Edit attribute updated for {(string)firstContainer.Parent!.Parent!.Attribute( "Name" )!}\{(string)firstContainer.Parent.Attribute( "Name" )!}\{(string)firstContainer.Parent.Attribute( "Name" )!}" );
					}
				}
				else if ( view != null && ( (string)view.Attribute( "Edit" )! ).StartsWith( @"Generated\" ) )
				{
					// Remove View if starts with Generated
					view.Remove();
					messages.Add( $@"{layoutInfo.Type}.{layoutKey} Edit attribute updated for {(string)view.Parent!.Parent!.Parent!.Parent!.Attribute( "Name" )!}\{(string)view.Parent.Parent.Parent.Attribute( "Name" )!}\{(string)view.Parent.Parent.Parent.Attribute( "Name" )!}" );
				}
			}

			if ( layout.HasElements )
			{
				var useRBLe = (bool?)layout.Attribute( "useRBLe" ) ?? false;
				var sbDesigner = new StringBuilder();
				var sbCodeBehindValidationList = new StringBuilder();
				var sbCodeBehindValidationHelpers = new StringBuilder();
				var sbMarkup = new StringBuilder();

				var firstTemplate = true;
				var currentGroup = 1;

				foreach ( var g in layout.Elements( "Group" ) )
				{
					var sbGroupInputs = new StringBuilder();

					foreach ( var i in g.Elements( "Input" ) )
					{
						var inputId = (string)i.Attribute( "ID" )!;
						var css = (string?)i.Attribute( "Css" ) ?? "col-sm-6";
						var label = (string?)i.Attribute( "Label" );
						var help = (string?)i.Attribute( "Help" );
						var inputType = (string)i.Attribute( "Type" )!;
						var typeParts = inputType.Split( ':' );
						var isDropDown = string.Compare( typeParts[ 0 ], "List", true ) == 0;
						var isCheckBox = string.Compare( inputType, "CheckBox", true ) == 0;
						var isDate = string.Compare( typeParts[ 0 ], "Date", true ) == 0;
						var bootstrapType = "TextBox";
						var listName = isDropDown && typeParts.Length > 1 ? typeParts[ 1 ] : null;
						var isCurrency = string.Compare( typeParts[ 0 ], "Currency", true ) == 0;
						var isPercentage = string.Compare( typeParts[ 0 ], "Percentage", true ) == 0;
						var prefix = isCurrency ? "$" : null;
						var suffix = isPercentage ? "%" : null;
						if ( !( (bool?)i.Attribute( "TriggersCalculation" ) ?? false ) )
						{
							css += " skipRBLe";
						}

						var defaultValue = (string?)i.Attribute( "Default" );
						var visibility = (string?)i.Attribute( "Visibility" );

						//SaveTo
						var attributes = new[]
						{
							!string.IsNullOrEmpty( help ) ? $" HelpContent='<%# {inputId}Help %>'" : null,
							isDate ? " IsDate=\"true\"" : null,
							!string.IsNullOrEmpty( prefix ) ? $" AddOnPrefix=\"{prefix}\"" : null,
							!string.IsNullOrEmpty( suffix ) ? $" AddOnSuffix=\"{suffix}\"" : null,
							!string.IsNullOrEmpty( listName ) ? $" DataSource='<%# xDSHelper.GetDataTableWithBlank( \"Table{listName}\" ) %>'" : null,
							!string.IsNullOrEmpty( defaultValue ) && isDropDown ? $" DefaultValue='<%# {inputId}Default %>'" : null,
							!string.IsNullOrEmpty( defaultValue ) && isCheckBox ? $" DefaultChecked='<%# {inputId}Default %>'" : null,
							!string.IsNullOrEmpty( defaultValue ) && !isDropDown && !isCheckBox ? $" DefaultText='<%# {inputId}Default %>'" : null,
							!string.IsNullOrEmpty( visibility ) ? $" Visible='<%# {inputId}Visible %>'" : null
						};

						if ( isDropDown )
						{
							bootstrapType = "DropDown";
						}
						else if ( isCheckBox )
						{
							bootstrapType = "CheckBox";
						}

						sbGroupInputs.AppendLine( $"\t\t\t<mh:Bootstrap{bootstrapType} CssClass=\"{css}\" Label=\"{label}\" runat=\"server\" ID=\"{inputId}\"{string.Join( "", attributes.Where( a => !string.IsNullOrEmpty( a ) ) )} />" );
						sbDesigner.AppendLine( $"\t\tprotected global::BTR.Evolution.MadHatter.Web.Controls.Bootstrap{bootstrapType} {inputId};" );

						// Validation Processing
						var min = (string?)i.Attribute( "Min" );
						var minAge = (string?)i.Attribute( "MinAge" );
						var max = (string?)i.Attribute( "Max" );
						var maxAge = (string?)i.Attribute( "MaxAge" );
						var required = (string?)i.Attribute( "Required" );
						var regularExpression = (string?)i.Attribute( "RegularExpression" );
						var compareTo = (string?)i.Attribute( "CompareTo" );
						var message = (string?)i.Attribute( "Message" );
						var isValid = (string?)i.Attribute( "IsValid" );

						var hasValidation = new[] { isValid, min, minAge, max, maxAge, required, regularExpression, compareTo }.Any( v => !string.IsNullOrEmpty( v ) );

						var isDouble = isCurrency || isPercentage || string.Compare( typeParts[ 0 ], "Double", true ) == 0;
						var isInteger = string.Compare( typeParts[ 0 ], "Integer", true ) == 0;
						var isNumber = isCurrency || isPercentage || isDouble || isInteger;
						var inputValidationType =
							isDate ? "DateTime" :
							isDouble ? "double" :
							isInteger ? "int" : "string";
						var isRequired = !string.IsNullOrEmpty( required );

						if ( !string.IsNullOrEmpty( defaultValue ) )
						{
							sbCodeBehindValidationHelpers.AppendLine( $"\t\tprotected {( isCheckBox ? "bool" : "string" )} {inputId}Default => {defaultValue};{Environment.NewLine}" );
						}
						if ( !string.IsNullOrEmpty( visibility ) )
						{
							sbCodeBehindValidationHelpers.AppendLine( $"\t\tprotected bool {inputId}Visible => {visibility};{Environment.NewLine}" );
						}
						sbCodeBehindValidationHelpers.AppendLine( $"\t\tprotected bool {inputId}Required => {( isRequired ? required : "false" )};{Environment.NewLine}" );

						var isAgeDate = isDate && ( minAge != null || maxAge != null );

						string getTokenString( string m )
						{
							if ( string.IsNullOrEmpty( m ) ) return m;

							var t =
								m.Replace( "{Label}", "{0}", StringComparison.InvariantCultureIgnoreCase )
									.Replace( "{MinAge", "{3", StringComparison.InvariantCultureIgnoreCase )
									.Replace( "{MaxAge", "{4", StringComparison.InvariantCultureIgnoreCase )
									.Replace( "{Min", "{1", StringComparison.InvariantCultureIgnoreCase )
									.Replace( "{Max", "{2", StringComparison.InvariantCultureIgnoreCase );

							return
								!isNumber && !isDate ? $"GetStringFormat( {t}, GetString( \"{label}\" ) )" :
								isAgeDate ? $"GetStringFormat( {t}, GetString( \"{label}\" ), {inputId}Min, {inputId}Max, {inputId}MinAge, {inputId}MaxAge )" :
								isDouble || isInteger || isDate ? $"GetStringFormat( {t}, GetString( \"{label}\" ), {inputId}Min, {inputId}Max )" :
								"\"TokenMessageNotSupported\"";
						}

						if ( !string.IsNullOrEmpty( help ) )
						{
							help = getTokenString( "\"" + help + "\"" );
							sbCodeBehindValidationHelpers.AppendLine( $"\t\tprotected string {inputId}Help => {help};{Environment.NewLine}" );
						}

						if ( hasValidation )
						{
							sbCodeBehindValidationList.AppendLine( $"\t\t\t\tnew {{ {inputId}.ID, Message = Validate{inputId}() }}," );

							var validationHelpers = "";

							if ( isNumber || isDate )
							{
								var minDefault =
									isDate ? "DateTime.MinValue" :
									isDouble ? "double.MinValue" :
									"int.MinValue";
								var maxDefault =
									isDate ? "DateTime.MaxValue" :
									isDouble ? "double.MaxValue" :
									"int.MaxValue";

								validationHelpers +=
									$"\t\tprotected {inputValidationType} {inputId}Min => {min ?? minDefault};{Environment.NewLine}" +
									$"\t\tprotected {inputValidationType} {inputId}Max => {max ?? maxDefault};{Environment.NewLine + Environment.NewLine}";
							}

							if ( isAgeDate )
							{
								validationHelpers +=
									$"\t\tprotected int {inputId}MinAge => {minAge ?? "int.MinValue"};{Environment.NewLine}" +
									$"\t\tprotected int {inputId}MaxAge => {maxAge ?? "int.MaxValue"};{Environment.NewLine + Environment.NewLine}";
							}

							if ( string.IsNullOrEmpty( message ) )
							{
								// Default message string
								message =
									!isNumber && !isDate ? "\"Exceptions.Calculation.Required.Text\"" :
									isAgeDate ? $"{inputId}Required ? \"Exceptions.Calculation.Required.AgeDate\" : \"Exceptions.Calculation.AgeDate\"" :
									isDouble ? $"{inputId}Required ? \"Exceptions.Calculation.Required.Double\" : \"Exceptions.Calculation.Double\"" :
									isInteger ? $"{inputId}Required ? \"Exceptions.Calculation.Required.Integer\" : \"Exceptions.Calculation.Integer\"" :
									isDate ? $"{inputId}Required ? \"Exceptions.Calculation.Required.Date\" : \"Exceptions.Calculation.Date\"" :
									"\"ValidationMessageNotSupported\"";
							}
							else
							{
								message = "\"" + message + "\"";
							}

							message = getTokenString( message );

							validationHelpers +=
								$"\t\tprotected string Validate{inputId}(){Environment.NewLine}" +
								$"\t\t{{{Environment.NewLine}" +
								$"\t\t\tvar message = {message};{Environment.NewLine}";

							// Required/empty
							validationHelpers +=
								$"\t\t\tif ( string.IsNullOrEmpty( {inputId}.Value ) ) return {inputId}Required ? message : null;{Environment.NewLine}";

							if ( isAgeDate )
							{
								validationHelpers +=
									$"\t\t\tif ( !IsInputValid<DateTime>( {inputId} ) && !IsInputValid<int>( {inputId} ) ) return message;{Environment.NewLine}" +
									$"\t\t\tvar dateValue = TryGetInputValue<DateTime>( {inputId}, out var dv ) ? dv : (DateTime?)null;{Environment.NewLine}" +
									$"\t\t\tvar ageValue = TryGetInputValue<int>( {inputId}, out var iv ) ? iv : (int?)null;{Environment.NewLine}" +
									$"\t\t\tif ( dateValue != null && ( dateValue.Value < {inputId}Min || dateValue.Value > {inputId}Max ) ) return message;{Environment.NewLine}" +
									$"\t\t\tif ( ageValue != null && ( ageValue.Value < {inputId}MinAge || ageValue.Value > {inputId}MaxAge ) ) return message;{Environment.NewLine}";
							}
							else if ( isNumber || isDate )
							{
								validationHelpers +=
									$"\t\t\tif ( !IsInputValid<{inputValidationType}>( {inputId} ) ) return message;{Environment.NewLine}" +
									$"\t\t\tvar value = GetInputValue<{inputValidationType}>( {inputId} );{Environment.NewLine}" +
									$"\t\t\tif ( value < {inputId}Min || value > {inputId}Max ) return message;{Environment.NewLine}";
							}

							if ( !string.IsNullOrEmpty( compareTo ) )
							{
								validationHelpers += $"\t\t\tif ( !new CompareInput<{inputValidationType}> {{ Input = {inputId} }}.{compareTo} ) return message;{Environment.NewLine}";
							}

							if ( !string.IsNullOrEmpty( regularExpression ) )
							{
								validationHelpers +=
									$"\t\t\tif ( !Regex.IsMatch( {inputId}.Text, @\"{regularExpression}\" ) ) return message;{Environment.NewLine}";
							}

							if ( !string.IsNullOrEmpty( isValid ) )
							{
								// Putting !( ) around it in case their validation is ! something or other
								validationHelpers +=
									$"\t\t\tif ( !( {isValid} ) ) return message;{Environment.NewLine}";
							}

							validationHelpers +=
								$"\t\t\treturn null; // No error{Environment.NewLine}" +
								$"\t\t}}{Environment.NewLine}";

							sbCodeBehindValidationHelpers.AppendLine( validationHelpers );
						}
					}

					sbMarkup.AppendLine(
						groupTemplate
							.Replace( "{FirstTemplate}", firstTemplate ? " mt-0" : "" )
							.Replace( "{GroupNumber}", currentGroup.ToString() )
							.Replace( "{Label}", ( !isProfileCalculation ? layoutInfo.Name + " " : "" ) + (string)g.Attribute( "Label" )! )
							.Replace( "{Inputs}", Environment.NewLine + sbGroupInputs.ToString() )
					);

					currentGroup++;
					firstTemplate = false;
				}

				saveFile( resxTemplate, $"{className}.Generated.ascx.resx" );

				saveFile(
					codeBehindTemplate
						.Replace( "{ClassName}", className )
						.Replace( "{BaseClassName}", isProfileCalculation ? "ProfileCalculationBaseControl" : "CalculationBaseControl" )
						.Replace( "{UseRBLe}", useRBLe ? "true" : "false" )
						.Replace( "{InputValidationList}", sbCodeBehindValidationList.ToString() )
						.Replace( "{InputValidationHelpers}", sbCodeBehindValidationHelpers.ToString() ),
					$"{className}.Generated.ascx.cs" );

				saveFile(
					designerTemplate
						.Replace( "{ClassName}", className )
						.Replace( "{Inputs}", sbDesigner.ToString() ),
					$"{className}.Generated.ascx.designer.cs"
				);

				saveFile(
					ascxTemplate
						.Replace( "{ClassName}", className )
						.Replace( "{Inputs}", sbMarkup.ToString() )
						.Replace( " class=\"needsRBLeConfig RBLe\"", useRBLe ? " class=\"needsRBLeConfig RBLe\"" : "" ),
					$"{className}.Generated.ascx"
				);
			}
		}
	}

	private static string GetPropertyType( XElement f )
	{
		var isRequired = f.Element( "Validation" ) != null && f.Element( "Validation" )!.Element( "AllowEmpty" ) == null;
		var lookupTable = (string?)f.Attribute( "lookuptable" );

		string? propertyType;
		if ( f.Element( "Validation" )?.Element( "Date" ) != null )
		{
			propertyType = "DateTime";
		}
		else if ( f.Element( "Validation" )?.Element( "Number" ) != null )
		{
			propertyType = (string)f.Element( "Validation" )!.Element( "Number" )!.Attribute( "dataType" )! == "Integer"
				? "int"
				: "double";
		}
		else if ( (string?)f.Element( "Validation" )?.Element( "HistoryIndex" )?.Attribute( "suffix" ) == "h" )
		{
			return "YearHalfHistoryIndex";
		}
		else if ( (string?)f.Element( "Validation" )?.Element( "HistoryIndex" )?.Attribute( "suffix" ) == "q" )
		{
			return "YearQuarterHistoryIndex";
		}
		else if ( (string?)f.Element( "Validation" )?.Element( "HistoryIndex" )?.Attribute( "suffix" ) == "m" )
		{
			return "YearMonthHistoryIndex";
		}
		else if ( !string.IsNullOrEmpty( lookupTable ) )
		{
			propertyType = string.Concat( lookupTable.AsSpan( 5 ), "Lookup" );
		}
		else
		{
			return "string";
		}

		return isRequired ? propertyType : propertyType + "?";
	}

	private static string SafeMemberName( string text )
	{
		var name =
			RemoveDiacritics(
				string.Concat(
					text.Replace( " ", "-" ).Replace( "–", "-" )
						.Split( '-' ).Where( p => !string.IsNullOrEmpty( p ) )
						.Select( Domain.Extensions.StringExtensions.CamelToPascalCase )
						.ToArray()
				)
			)!
			.Replace( "%", "Pct" )
			.Replace( "$", "Dollars" )
			.Replace( "=", "Equals" )
			.Replace( "&amp;", "And" )
			.Replace( "&", "And" )
			.Replace( "(closed)", "Closed" )
			.Replace( "+", "Plus" )

			.Replace( "|", "_" )
			.Replace( "*", "_" )
			.Replace( "/", "_" )
			.Replace( "|", "_" )
			.Replace( "#", "_" )
			.Replace( "^", "_" )

			// .Replace("-", "_")
			// .Replace("–", "_")
			// .Replace(" ", "")

			.Replace( ".", "" )
			.Replace( "@", "" )
			.Replace( "…", "" )
			.Replace( "...", "" )
			.Replace( ":", "" )
			.Replace( "?", "" )
			.Replace( ",", "" )
			.Replace( "'", "" )
			.Replace( "\"", "" )
			.Replace( "[", "" ).Replace( "]", "" )
			.Replace( "(", "" ).Replace( ")", "" );

		if ( char.IsDigit( name[ 0 ] ) || name.ToLower() == "boolean" )
		{
			name = "_" + name;
		}

		return name;
	}

	private static string? RemoveDiacritics( string? s )
	{
		if ( s == null ) return null;

		var normalizedString = s.Normalize( NormalizationForm.FormD );
		var stringBuilder = new StringBuilder();

		for ( var i = 0; i < normalizedString.Length; i++ )
		{
			var c = normalizedString[ i ];

			if ( CharUnicodeInfo.GetUnicodeCategory( c ) != UnicodeCategory.NonSpacingMark )
			{
				stringBuilder.Append( c );
			}
		}

		return stringBuilder.ToString();
	}

	private static void MergeEvolutionReports( string currentPath, XElement[] specReports, XElement? calcInputLayouts )
	{
		var configurationPath = Path.Combine( currentPath, "Configuration.xml" );

		if ( specReports.Length == 0 || !File.Exists( configurationPath ) || !currentPath.Contains( @"\Evolution\Websites\Admin\", StringComparison.InvariantCultureIgnoreCase ) ) return;

		var webConfigPath = Path.Combine( currentPath, @"..\Web.config" );
		var webConfig = XElement.Load( webConfigPath );
		XNamespace ns = "http://schemas.benefittech.com/evolution/site";
		var configuration = XElement.Load( configurationPath );

		var reportsConfig = configuration.Element( ns + "Reports" );
		if ( reportsConfig == null )
		{
			configuration.Add( reportsConfig = new XElement( ns + "Reports" ) );
		}

		foreach ( var specReport in specReports )
		{
			var specReportId = (string)specReport.Attribute( "ID" )!;
			var reportConfig = reportsConfig.Elements( ns + "Report" ).FirstOrDefault( r => (string)r.Attribute( "ID" )! == specReportId );

			if ( reportConfig == null )
			{
				reportsConfig.Add( reportConfig = new XElement( ns + "Report" ) );
			}

			string? title;
			if ( ( title = (string?)specReport.Attribute( "Title" ) )?.StartsWith( "Evaluate:" ) ?? false )
			{
				var eval = ExpressionEvaluator.ParseAndEvaluateExpression( title );

				var extension = Path.GetExtension( eval )?.EnsureNonBlank();
				var extAttribute = specReport.Attribute( "Extension" );

				if ( extAttribute == null )
				{
					specReport.Add( new XAttribute( "Extension", extension ?? ".csv" ) );
				}
				else
				{
					extAttribute.Value = extension ?? ".csv";
				}

				if ( !string.IsNullOrEmpty( extension ) )
				{
					// Presummably it is at end of expression, so appending " during replace.
					title = title.Replace( extension + "\"", "\"" );
					specReport.Attribute( "Title" )!.Value = title;
				}
			}

			MergeReportAttributes( specReport, reportConfig );

			var hasSubmitPage =
				calcInputLayouts?.Elements( "Layout" )
					.Any( l => (string?)l.Attribute( "Type" ) == "iReportType" && (string?)l.Attribute( "Key" ) == specReportId ) ?? false;

			if ( hasSubmitPage )
			{
				var submitPage = $@"Generated\Report{specReportId}.Generated.ascx";
				if ( reportConfig.Attribute( "SubmitPagePath" ) != null )
				{
					reportConfig.Attribute( "SubmitPagePath" )!.Value = submitPage;
				}
				else
				{
					reportConfig.Add( new XAttribute( "SubmitPagePath", submitPage ) );
				}
			}

			foreach ( var p in specReport.Elements().Where( e => e.Name.LocalName != "reportFTPSettings" ) )
			{
				var configParams = reportConfig.Element( ns + p.Name.LocalName );
				if ( configParams == null )
				{
					reportConfig.Add( configParams = new XElement( ns + p.Name.LocalName ) );
				}
				MergeReportAttributes( p, configParams );
			}
		}

		var validReportKeys = specReports.Select( sr => (string)sr.Attribute( "ID" )! ).ToArray();

		var deletedReports =
			configuration
				.Elements( ns + "Reports" ).Elements( ns + "Report" )
				.Where( r => !validReportKeys.Contains( (string)r.Attribute( "ID" )! ) )
				.ToArray();

		if ( deletedReports.Any() )
		{
			MessageBox.Show(
				$"The following reports were removed from the Spec Sheet:{Environment.NewLine + Environment.NewLine + string.Join( ", ", deletedReports.Select( r => (string)r.Attribute( "ID" )! ) )}",
				"Deleted Reports To Process",
				MessageBoxButtons.OK,
				MessageBoxIcon.Exclamation
			);
			deletedReports.Remove();
		}

		var saveWebConfig = false;

		var badFtp =
			webConfig
				.Elements( "evolutionAdministrationSettings" )
				.Elements( "reportFTPSettings" )
				.Elements( "add" )
				.Where( a => !validReportKeys.Contains( (string)a.Attribute( "key" )! ) );

		if ( badFtp.Any() )
		{
			badFtp.Remove();
			saveWebConfig = true;
		}

		var ftpSettings = specReports.Elements( "reportFTPSettings" ).Elements( "add" );

		if ( ftpSettings.Any() )
		{
			saveWebConfig = true;

			var reportFTPSettings =
				webConfig
					.GetOrCreateElement( "evolutionAdministrationSettings" )
					.GetOrCreateElement( "reportFTPSettings" );

			foreach ( var ftp in ftpSettings )
			{
				var webAdd = reportFTPSettings.Elements( "add" ).FirstOrDefault( a => (string)a.Attribute( "key" )! == (string)ftp.Attribute( "key" )! );

				if ( webAdd == null )
				{
					reportFTPSettings.Add( ftp );
				}
				else
				{
					MergeReportAttributes( ftp, webAdd );
				}
			}
		}

		if ( saveWebConfig )
		{
			webConfig.SaveIndented( webConfigPath );
		}


	}

	private static void MergeReportAttributes( XElement source, XElement dest )
	{
		foreach ( var a in source.Attributes() )
		{
			var configAttr = dest.Attribute( a.Name.LocalName );
			if ( configAttr == null )
			{
				dest.Add( new XAttribute( a.Name.LocalName, (string)a ) );
			}
			else if ( a.Name.LocalName != "Page" || !configAttr.Value.Contains( '|' ) )
			{
				// Don't have a config col in Excel yet, and if they edit, don't want to change if they added 2
				configAttr.Value = (string)a;
			}
			else if ( a.Name.LocalName == "Page" && configAttr.Value.Contains( '|' ) && !configAttr.Value.Contains( (string)a ) )
			{
				var id = (string)source.Attribute( "ID" )!;
				MessageBox.Show( $"Report {id} has Page value of {configAttr.Value}, but should probably contain {id}." );
			}
		}
	}

	private XElement? GetCalcInputLayouts( XElement[] specReports, XElement? configLookups )
	{
		var sheet = workbook.GetWorksheet( "Calc Inputs" );
		var inputs = sheet?.RangeOrNull( "Inputs" );
		var layouts = sheet?.RangeOrNull( "Layouts" );

		if ( sheet == null || inputs == null || layouts == null ) return null;

		var inputColumns = GetColumnConfiguration<CalcInputColumnType>( inputs.Offset[ 1, 0 ] );
		var inputConfig = new Dictionary<string, MSExcel.Range>();

		string? name;
		var row = 0;
		while ( !string.IsNullOrEmpty( name = inputs.Offset[ 2 + row, 0 ].GetText() ) )
		{
			inputConfig[ name ] = inputs.Offset[ 2 + row, 0 ];
			row++;
		}

		if ( inputConfig.Count == 0 ) return null;

		var inputLayouts = new XElement( "CalcInputLayouts" );

		row = 0;

		string? layoutInfo;
		MSExcel.Range layoutRow;
		while ( !string.IsNullOrEmpty( layoutInfo = layouts.Offset[ row, 0 ].GetText() ) )
		{
			var useRble = layoutInfo.HasSwitch( "use-rble" );
			var layoutParts = layoutInfo.Split( '/' )[ 0 ].Split( ':' );
			var layoutType = layoutParts[ 0 ];
			var layoutName = layoutParts[ 1 ];
			var validLayout = layoutType switch
			{
				"iMHACalcType" => configLookups?.Elements( "DataTable" ).Where( t => (string?)t.Attribute( "id" ) == "TableMHACalculationTypes" ).Elements( "Table" ).Elements( "TableItem" ).Any( i => (string?)i.Attribute( "FolderItemType" ) == layoutName ) ?? false,
				"iProcessType" => configLookups?.Elements( "DataTable" ).Where( t => (string?)t.Attribute( "id" ) == "TableRBLProcessTypes" || (string?)t.Attribute( "id" ) == "TableRBLCalculationTypes" ).Elements( "Table" ).Elements( "TableItem" ).Any( i => (string?)i.Attribute( "key" ) == layoutName ) ?? false,
				"iPackageType" => configLookups?.Elements( "DataTable" ).Where( t => (string?)t.Attribute( "id" ) == "TableDocTypes" ).Elements( "Table" ).Elements( "TableItem" ).Any( i => (string?)i.Attribute( "key" ) == layoutName ) ?? false,
				"iReportType" => specReports.Any( r => (string?)r.Attribute( "ID" ) == layoutName ),
				_ => false
			};
			var reportName = validLayout && layoutType == "iReportType"
				? (string?)specReports.First( r => (string?)r.Attribute( "ID" ) == layoutName )!.Attribute( "Name" )
				: null;

			var layoutColumns = GetColumnConfiguration<CalcInputColumnType>( layouts.Offset[ row, 1 ] );

			var inputLayout =
				new XElement( "Layout",
					new XAttribute( "Type", layoutType ),
					new XAttribute( "Key", layoutName ),
					!string.IsNullOrEmpty( reportName ) ? new XAttribute( "ReportName", reportName ) : null
				);

			row++; // Move to inputs...
			string? inputInfo;
			var group = new XElement( "Group", new XAttribute( "Label", "Assumptions" ) );

			while ( !string.IsNullOrEmpty( inputInfo = ( layoutRow = layouts.Offset[ row, 0 ] ).GetText() ) )
			{
				var inputParts = inputInfo.Split( ':' );

				if ( string.Compare( inputParts[ 0 ], "NO INPUTS", true ) != 0 )
				{
					if ( string.Compare( inputParts[ 0 ], "GROUP", true ) == 0 )
					{
						if ( group.HasElements )
						{
							inputLayout.Add( group );
						}
						group = new XElement( "Group", new XAttribute( "Label", inputParts[ 1 ] ) );
					}
					else if ( inputConfig.TryGetValue( inputParts[ 0 ], out var inputRange ) )
					{
						var layoutValues = layoutRow.Offset[ 0, 1 ];

						var label = GetCalcInputValue( CalcInputColumnType.Label, layoutValues, layoutColumns, inputRange, inputColumns );
						var help = GetCalcInputValue( CalcInputColumnType.Help, layoutValues, layoutColumns, inputRange, inputColumns );
						var triggersCalculation = new[] { "Y", "TRUE" }.Contains( GetCalcInputValue( CalcInputColumnType.TriggersCalculation, layoutValues, layoutColumns, inputRange, inputColumns ), StringComparer.OrdinalIgnoreCase );
						var visibility = GetCalcInputValue( CalcInputColumnType.Visibility, layoutValues, layoutColumns, inputRange, inputColumns );
						var defaultValue = GetCalcInputValue( CalcInputColumnType.DefaultValue, layoutValues, layoutColumns, inputRange, inputColumns );
						var min = GetCalcInputValue( CalcInputColumnType.Min, layoutValues, layoutColumns, inputRange, inputColumns );
						var minAge = GetCalcInputValue( CalcInputColumnType.MinAge, layoutValues, layoutColumns, inputRange, inputColumns );
						var max = GetCalcInputValue( CalcInputColumnType.Max, layoutValues, layoutColumns, inputRange, inputColumns );
						var maxAge = GetCalcInputValue( CalcInputColumnType.MaxAge, layoutValues, layoutColumns, inputRange, inputColumns );
						var required = GetCalcInputValue( CalcInputColumnType.Required, layoutValues, layoutColumns, inputRange, inputColumns );
						var regularExpression = GetCalcInputValue( CalcInputColumnType.RegularExpression, layoutValues, layoutColumns, inputRange, inputColumns );
						var compareTo = GetCalcInputValue( CalcInputColumnType.CompareTo, layoutValues, layoutColumns, inputRange, inputColumns );
						var isValid = GetCalcInputValue( CalcInputColumnType.IsValid, layoutValues, layoutColumns, inputRange, inputColumns );
						var message = GetCalcInputValue( CalcInputColumnType.Message, layoutValues, layoutColumns, inputRange, inputColumns );

						if ( triggersCalculation )
						{
							useRble = true;
						}

						group.Add(
							new XElement( "Input",
								new XAttribute( "ID", inputParts[ 0 ] ),
								new XAttribute( "Type", GetCalcInputValue( CalcInputColumnType.InputType, layoutValues, layoutColumns, inputRange, inputColumns, "Text" )! ),
								!string.IsNullOrEmpty( label ) ? new XAttribute( "Label", label ) : null,
								!string.IsNullOrEmpty( help ) ? new XAttribute( "Help", help ) : null,
								new XAttribute( "Css", GetCalcInputValue( CalcInputColumnType.Css, layoutValues, layoutColumns, inputRange, inputColumns, "col-sm-6" )! ),
								triggersCalculation ? new XAttribute( "TriggersCalculation", true ) : null,
								!string.IsNullOrEmpty( visibility ) ? new XAttribute( "Visibility", visibility ) : null,
								!string.IsNullOrEmpty( defaultValue ) ? new XAttribute( "Default", defaultValue ) : null,
								!string.IsNullOrEmpty( min ) ? new XAttribute( "Min", min ) : null,
								!string.IsNullOrEmpty( minAge ) ? new XAttribute( "MinAge", minAge ) : null,
								!string.IsNullOrEmpty( max ) ? new XAttribute( "Max", max ) : null,
								!string.IsNullOrEmpty( maxAge ) ? new XAttribute( "MaxAge", maxAge ) : null,
								!string.IsNullOrEmpty( required ) ? new XAttribute( "Required", required ) : null,
								!string.IsNullOrEmpty( regularExpression ) ? new XAttribute( "RegularExpression", regularExpression ) : null,
								!string.IsNullOrEmpty( compareTo ) ? new XAttribute( "CompareTo", compareTo ) : null,
								!string.IsNullOrEmpty( isValid ) ? new XAttribute( "IsValid", isValid ) : null,
								!string.IsNullOrEmpty( message ) ? new XAttribute( "Message", message ) : null
							)
						);
					}
					else
					{
						throw new ApplicationException( $"Unable to find the input '{inputParts[ 0 ]}' in the Calc Inputs sheet configuration section." );
					}
				}
				row++; // Next input...
			}
			if ( group.HasElements )
			{
				inputLayout.Add( group );
			}

			if ( validLayout )
			{
				if ( useRble )
				{
					inputLayout.Add( new XAttribute( "useRBLe", true ) );
				}

				inputLayouts.Add( inputLayout );
			}

			row++; // Next layout...
		}

		return inputLayouts;
	}

	private static string? GetCalcInputValue(
		CalcInputColumnType column,
		MSExcel.Range layoutRow, Dictionary<CalcInputColumnType, ColumnDefinition> layoutColumns,
		MSExcel.Range inputsRow, Dictionary<CalcInputColumnType, ColumnDefinition> inputColumns,
		string? defaultValue = null
	)
	{
		var value = layoutColumns.GetValue( column, layoutRow );

		if ( value == "{REMOVE}" ) return null;

		return value ??
			inputColumns.GetValue( column, inputsRow ) ??
			defaultValue;
	}

	private XElement? ExportLookups()
	{
		var sheet = workbook.GetWorksheet( "Code Tables" );
		if ( sheet == null ) return null;

		var tableData = sheet.RangeOrNull( "A1" )!;

		while ( tableData.GetText() != "Table" && tableData.Row < 100 )
		{
			tableData = tableData.End[ MSExcel.XlDirection.xlDown ];
		}

		var lookups = new XElement( "DataTableDefs" );

		string? tableName;

		while ( tableData.GetText() == "Table" && !string.IsNullOrEmpty( tableName = tableData.Offset[ 0, 1 ].GetText() ) )
		{
			var includeInfo = tableData.Offset[ 0, 2 ].GetText().EnsureNonBlank() ?? "Y";

			XElement table;

			lookups.Add(
				new XElement( "DataTable",
					new XAttribute( "id", $"Table{tableName}" ),
					new XAttribute( "key", "key" ),
					includeInfo.HasSwitch( "append" ) ? new XAttribute( "append", true ) : null,
					includeInfo.HasSwitch( "remove" ) ? new XAttribute( "remove", true ) : null,
					table = new XElement( "Table" )
				)
			);

			var headers = ( sheet.Range[ tableData.Offset[ 1, 0 ], tableData.Offset[ 1, 0 ].End[ MSExcel.XlDirection.xlToRight ] ] as MSExcel.Range )!.GetValues<string>()!;

			var includeCol = Array.IndexOf( headers, "Include" );
			var keyCol = Array.IndexOf( headers, "Value" );
			var valueCol = Array.IndexOf( headers, "Label" );

			var knownColumns = new[] { includeCol, keyCol, valueCol };

			var rows = tableData.Offset[ 2, 0 ];
			var offset = 0;

			string? include;
			string? key;
			string? value;

			while (
				new[]
				{
					include = includeCol != -1 ? rows.Offset[ offset, includeCol ].GetText().EnsureNonBlank() : null,
					key = keyCol != -1 ? rows.Offset[ offset, keyCol ].GetText().EnsureNonBlank() : null,
					value = valueCol != -1 ? rows.Offset[ offset, valueCol ].GetText().Trim().EnsureNonBlank() : null
				}.Any( s => s != null )
			)
			{
				if ( include?[ 0 ] == 'Y' )
				{
					string? visibleState;
					table.Add(
						new XElement( "TableItem",
							new XAttribute( "key", key ?? "" ),
							value != null ? new XAttribute( "name", value ) : null,
							!string.IsNullOrEmpty( visibleState = include?.GetSwitchValue( "visibleStateID" ) ) ? new XAttribute( "visible", visibleState ) : null,
							headers.Select( ( h, i ) =>
							{
								var value = rows.Offset[ offset, i ].GetText().Trim().EnsureNonBlank();

								return value != null && !knownColumns.Contains( i )
									? new XAttribute( headers[ i ], value )
									: null;
							} )
						)
					);
				}

				offset++;
			}

			tableData = tableData.End[ MSExcel.XlDirection.xlToRight ].End[ MSExcel.XlDirection.xlToRight ];
		}

		return lookups;
	}

	private XElement ExportPlanInfo()
	{
		var sheet = workbook.GetWorksheet( "Plan Info" )!;
		var version = sheet.RangeOrNull<double>( "SheetVersion" );

		if ( version < 2.3 )
		{
			throw new ApplicationException( "You can not use the 'Create Config-PlanInfo.xml' function because the Plan Info Sheet Version is not the correct version.  The version must be 2.3 or greater are needed to use this functionality." );
		}

		var versionHistory = sheet.RangeOrNull( "Version_History" ) ?? sheet.RangeOrNull( "VersionHistory" );
		var planInfo =
			new XElement( "PlanInfo",
				GetPlanInfoItemDefs( "General Info", sheet.RangeOrNull( "General_Information" ), sheet.RangeOrNull( "Search_Indexes" ) ),
				GetPlanInfoItemDefs( "Plan Provisions", sheet.RangeOrNull( "Plan_Provisions" ), versionHistory ),
				GetPlanInfoLinks( "Documents", sheet.RangeOrNull( "Documents" ), sheet.RangeOrNull( "Plan_Provisions" ) ?? versionHistory ),
				GetPlanInfoLinks( "Links", sheet.RangeOrNull( "Links" ), sheet.RangeOrNull( "Documents" ) )
			);
		return planInfo;
	}

	private static XElement? GetPlanInfoLinks( string name, MSExcel.Range? start, MSExcel.Range? end )
	{
		if ( start == null || end == null )
		{
			return null;
		}

		var links = new XElement( name );

		var endAddress = end.Address;
		var offset = 1;

		while ( start.Offset[ offset, 0 ].Address != endAddress )
		{
			string? linkName;
			string? href;
			string? description;
			if ( !string.IsNullOrEmpty( linkName = start.Offset[ offset, 0 ].GetText() ) && !string.IsNullOrEmpty( href = start.Offset[ offset, 1 ].GetText() ) )
			{
				links.Add(
					new XElement( "Link",
						new XElement( "Name", linkName ),
						new XElement( "Href", href ),
						!string.IsNullOrEmpty( description = start.Offset[ offset, 2 ].GetText() ) ? new XElement( "Description", description ) : null
					)
				);
			}
			offset++;
		}

		return links.HasElements ? links : null;
	}

	private static XElement? GetPlanInfoItemDefs( string name, MSExcel.Range? start, MSExcel.Range? end )
	{
		if ( start == null || end == null )
		{
			return null;
		}

		var itemDefs = new XElement( "ItemDefs", new XAttribute( "id", name ) );

		var endAddress = end.Address;
		var offset = 0;

		while ( start.Offset[ offset, 0 ].Address != endAddress )
		{
			string? text;
			string? value;
			if ( !string.IsNullOrEmpty( text = start.Offset[ offset, 0 ].GetText() ) && text != "Severance Variables" && !string.IsNullOrEmpty( value = start.Offset[ offset, 1 ].GetText() ) )
			{
				itemDefs.Add( new XElement( "ItemDef", new XAttribute( "id", text ), value ) );
			}
			offset++;
		}

		return itemDefs.HasElements ? itemDefs : null;
	}

	private void ExportProfile( XElement configProfile )
	{
		var excelExports = configProfile.Elements( "ExcelExports" ).Elements( "ExcelExport" ).ToArray();
		var excelExportColumns = excelExports.Select( e => (string)e.Attribute( "specification-id" )! ).ToArray();

		ProcessFlatDataFields( configProfile, excelExports, excelExportColumns );
		ProcessHistoricalDataFields( configProfile, excelExports, excelExportColumns );

		// Remove temp report attributes
		foreach ( var report in configProfile.Elements( "ExcelExports" ).Elements( "ExcelExport" ) )
		{
			report.Attributes().Where( a => a.Name.LocalName == "specification-id" || a.Name.LocalName == "source" ).Remove();

			var sortedColumns = report.Descendants( "Column" ).Where( e => e.Attribute( "order" ) != null );

			if ( sortedColumns.Any() )
			{
				// If any orders applied, then I need to rearrange things...
				var columns =
					sortedColumns
						.OrderBy( c => (int)c.Attribute( "order" )! )
						.Select( c =>
						{
							return c.Parent!.Name.LocalName == "HistoryData"
								? new XElement( "HistoryData", c.Parent!.Attributes(), c )
								: c;
						} )
						.ToArray();

				for ( var i = 1; i < columns.Length; i++ )
				{
					var item = columns[ i ];
					var previous = columns[ i - 1 ];

					var itemKey = item.Attributes().Select( a => $"{a.Name.LocalName}={a.Value}" ).Join( ", " );
					var previousKey = previous.Attributes().Select( a => $"{a.Name.LocalName}={a.Value}" ).Join( ", " );

					if ( itemKey == previousKey ) // same history items...
					{
						previous.Add( item.Elements() );
						item.RemoveAll();
					}
				}

				sortedColumns.Remove();

				columns =
					columns
						.Where( c => c.Name.LocalName == "Column" || c.HasElements )
						.Concat( report.Elements( "Column" ) )
						.Concat( report.Elements( "HistoryData" ) )
						.ToArray();

				report.Elements().Where( e => new[] { "Column", "HistoryData" }.Contains( e.Name.LocalName ) ).Remove();
				report.Add( columns );
			}
		}
	}

	private void ProcessFlatDataFields( XElement configProfile, XElement[] excelExports, string[] excelExportColumns )
	{
		var flatData = workbook.GetWorksheet( "Flat Data" )!;
		var dataColumns = GetColumnConfiguration<DataColumnType>( flatData.RangeOrNull( "A5" ) );
		var dataFields = flatData.RangeOrNull( "A6" )!;
		var offset = 0;
		var xDataDef = configProfile.Element( "xDataDef" )!;
		string? fieldInfo;
		MSExcel.Range row;
		var group = new XElement( "Profile", new XAttribute( "label", "Default" ), new XAttribute( "id", "Default" ) );
		var hasAuthId = false;
		var exportColumns = GetColumnConfiguration( flatData.RangeOrNull( "A5" ), excelExportColumns );

		while ( !string.IsNullOrEmpty( fieldInfo = dataColumns.GetValue( DataColumnType.DataField, row = dataFields.Offset[ offset, 0 ] ) ) )
		{
			var label = dataColumns.GetValue( DataColumnType.Label, row )!.Trim();

			switch ( fieldInfo )
			{
				case "GROUP":
					if ( group.HasElements )
					{
						xDataDef.Add( group );
					}
					group =
						new XElement( "Profile",
							new XAttribute( "label", label ),
							new XAttribute( "id", label.Replace( " ", "" ).Replace( "&", "" ) )
						);
					break;

				case "HEADER":
					group.Add( new XElement( "h", new XAttribute( "id", label.Replace( " ", "" ).Replace( "&", "" ) ), label ) );
					break;

				default:
					var dataField = GetDataField( configProfile, fieldInfo, dataColumns, row );

					if ( dataField != null )
					{
						if ( dataColumns.GetValue( DataColumnType.DataField, row )!.HasSwitch( "key" ) )
						{
							dataField.Add( new XAttribute( "id-auth", "1" ) );
							hasAuthId = true;
						}

						group.Add( dataField );

						UpdateExcelExportReports(
							excelExports,
							exportColumns,
							dataField.Name.LocalName,
							row
						);
					}
					break;
			}

			offset++;
		}

		if ( group.HasElements )
		{
			xDataDef.Add( group );
		}

		if ( !hasAuthId )
		{
			var ssn = xDataDef.Elements( "Profile" ).Elements( "ssn" ).FirstOrDefault();
			ssn?.Add( new XAttribute( "id-auth", "1" ) );
		}

		// Remove any 'headers' that are empty...
		xDataDef.Elements( "Profile" ).Elements( "h" )
			.Where( h => ( ( h.NextNode as XElement )?.Name.LocalName ?? "h" ) == "h" )
			.Remove();
	}

	private void ProcessHistoricalDataFields( XElement configProfile, XElement[] excelExports, string[] excelExportColumns )
	{
		var sheet = workbook.GetWorksheet( "Historical Data" )!;
		var tables = sheet.RangeOrNull( "A5" )!;
		var xDataDef = configProfile.Element( "xDataDef" )!;

		while ( string.Compare( "Data Type:", tables.GetText(), true ) == 0 )
		{
			var tableOffset = 1;
			string? historyType;

			var dataColumns = GetColumnConfiguration<DataColumnType>( tables.Offset[ 3, 0 ] );
			var exportColumns = GetColumnConfiguration( tables.Offset[ 3, 0 ], excelExportColumns );

			while ( !string.IsNullOrEmpty( historyType = tables.Offset[ 0, tableOffset ].GetText() ) )
			{
				var includeInfo = tables.Offset[ 1, tableOffset ].GetText().EnsureNonBlank() ?? "Y";

				if ( includeInfo[ 0 ] == 'Y' )
				{
					var historyTable = xDataDef.AddElement(
						new XElement( "HistoryData",
							new XAttribute( "type", historyType ),
							new XAttribute( "label", tables.Offset[ 2, tableOffset ].GetText() ),
							includeInfo.HasSwitch( "hide" ) ? new XAttribute( "visible", false ) : null
						)
					);

					var historyFields = tables.Offset[ 4, 0 ];

					string? fieldInfo;
					MSExcel.Range row;
					var fieldOffset = 0;

					while ( !string.IsNullOrEmpty( fieldInfo = dataColumns.GetValue( DataColumnType.DataField, row = historyFields.Offset[ fieldOffset, 0 ] ) ) )
					{
						var dataField = GetDataField( configProfile, fieldInfo, dataColumns, row );

						if ( dataField != null )
						{
							historyTable.Add( dataField );

							UpdateExcelExportReports(
								excelExports,
								exportColumns,
								dataField.Name.LocalName,
								row,
								historyType
							);
						}

						fieldOffset++;
					}
				}

				tableOffset++;
			}

			tables = tables.End[ MSExcel.XlDirection.xlDown ].End[ MSExcel.XlDirection.xlDown ];
		}
	}

	private XElement? GetDataField( XElement configProfile, string fieldInfo, Dictionary<DataColumnType, ColumnDefinition> dataColumns, MSExcel.Range row )
	{
		if ( dataColumns.GetValue( DataColumnType.Include, row ) != "Y" ) return null;

		var format = dataColumns.GetValue( DataColumnType.DisplayFormat, row );
		var dataType = dataColumns.GetValue( DataColumnType.DataType, row ) ?? "string";
		var displayWidth = dataColumns.GetValue( DataColumnType.DisplayWidth, row );
		var sortParts = dataColumns.GetValue( DataColumnType.Sort, row ).EnsureNonBlank()?.Split( ':' );

		var field =
			new XElement( fieldInfo.Split( '/' )[ 0 ],
				fieldInfo.HasSwitch( "hide" ) ? new XAttribute( "visible", false ) : null,
				fieldInfo.HasSwitch( "viewonly" ) ? new XAttribute( "viewonly", true ) : null,
				dataColumns.GetValue( DataColumnType.MadHatterInclude, row ) == "Y" ? new XAttribute( "mh", 1 ) : null,
				new XAttribute( "label", dataColumns.GetValue( DataColumnType.Label, row )!.Trim() ),
				!string.IsNullOrEmpty( format ) ? new XAttribute( "format", format ) : null,
				dataType.StartsWith( "list:" ) ? new XAttribute( "lookuptable", "Table" + dataType[ 5.. ] ) : null,
				string.Compare( dataColumns.GetValue( DataColumnType.SkipAudit, row ), "Y", true ) == 0 ? new XAttribute( "skipAudit", true ) : null,
				string.Compare( dataColumns.GetValue( DataColumnType.IsDetail, row ), "Y", true ) == 0 ? new XAttribute( "isDetail", true ) : null,
				!string.IsNullOrEmpty( displayWidth ) ? new XAttribute( "displayWidth", displayWidth ) : null,
				dataType.StartsWith( "string:" ) ? new XAttribute( "rows", dataType[ 7.. ] ) : null,
				sortParts != null ? new XAttribute( "sortPriority", sortParts[ 0 ] ) : null,
				sortParts != null ? new XAttribute( "sortDirection", sortParts.Length == 1 || string.Compare( ( sortParts[ 1 ] + "a" )[ 0..1 ], "A", true ) == 0 ? "Ascending" : "Descending" ) : null
			);

		dataType = dataType.ToLower().Split( ':' )[ 0 ];

		if ( !new[] { "list", "string", "date", "integer", "double", "year", "yearhalf", "yearquarter", "yearmonth" }.Contains( dataType ) )
		{
			throw new ApplicationException( $"Invalid field type.  {dataType} cannot be processed." );
		}

		// Create Validation if:
		//  1) Date, Integer, Double type or
		//  2) List type but Validation Type is warning or
		//  3) Any type and the item is required
		var validationErrorType = dataColumns.GetValue( DataColumnType.ValidationErrorType, row );
		var requiredErrorType = dataColumns.GetValue( DataColumnType.RequiredErrorType, row );
		var validationExpression = dataColumns.GetValue( DataColumnType.ValidationExpression, row );

		if (
			( dataType != "string" && ( dataType != "list" || validationErrorType == "W" ) ) ||
			!string.IsNullOrEmpty( requiredErrorType ) || !string.IsNullOrEmpty( validationExpression )
		)
		{
			var min = dataColumns.GetValue( DataColumnType.Min, row ).EnsureNonBlank();
			var max = dataColumns.GetValue( DataColumnType.Max, row ).EnsureNonBlank();

			var minValue = dataType switch
			{
				"date" => min ?? "1/1/1800",
				"integer" => min ?? "0",
				"double" => min ?? "-2147483648.00",
				"year" or "yearhalf" or "yearquarter" or "yearmonth" => min ?? "1800",
				_ => null
			};
			var maxValue = dataType switch
			{
				"date" => max ?? "12/31/9999",
				"integer" => max ?? "999999999",
				"double" => max ?? "2147483647.00",
				"year" or "yearhalf" or "yearquarter" or "yearmonth" => max ?? "9999",
				_ => null
			};

			if ( dataType == "year" )
			{
				dataType = "integer";
			}

			field.Add(
				new XElement( "Validation",
					requiredErrorType != "F" ? new XElement( "AllowEmpty", requiredErrorType == "W" ? new XAttribute( "type", "warning" ) : null ) : null,
					minValue == null && !string.IsNullOrEmpty( validationExpression ) ? new XAttribute( "regular-expression", validationExpression ) : null,
					minValue == null && !string.IsNullOrEmpty( validationExpression ) ? new XAttribute( "regular-expression-message", dataColumns.GetValue( DataColumnType.ValidationExpressionMessage, row )! ) : null
				)
			);

			if ( dataType == "double" || dataType == "integer" )
			{
				var auditVariance = dataColumns.GetValue( DataColumnType.AllowedAuditVariance, row );

				field.Element( "Validation" )!.Add(
					new XElement( "Number",
						new XAttribute( "dataType", dataType.ToProperCase() ),
						new XAttribute( "min", GetValidationRange( configProfile, dataType == "double" && !minValue!.Contains( '.' ) ? minValue + ".00" : minValue!, "min", dataType ) ),
						new XAttribute( "max", GetValidationRange( configProfile, dataType == "double" && !maxValue!.Contains( '.' ) ? maxValue + ".00" : maxValue!, "max", dataType ) ),
						validationErrorType == "W" ? new XAttribute( "type", "warning" ) : null,
						!string.IsNullOrEmpty( auditVariance ) ? new XAttribute( "auditVariance", auditVariance ) : null
					)
				);
			}
			else if ( dataType.StartsWith( "year" ) )
			{
				field.Element( "Validation" )!.Add(
					new XElement( "HistoryIndex",
						new XAttribute( "dataType", dataType.ToProperCase() ),
						new XAttribute( "min", GetValidationRange( configProfile, minValue!, "min", dataType ) ),
						new XAttribute( "max", GetValidationRange( configProfile, maxValue!, "max", dataType ) ),
						validationErrorType == "W" ? new XAttribute( "type", "warning" ) : null,
						new XAttribute( "suffix", dataType[ 4 ] )
					)
				);
			}
			else if ( dataType == "date" )
			{
				field.Element( "Validation" )!.Add(
					new XElement( "Date",
						new XAttribute( "min", GetValidationRange( configProfile, dataType == "double" && !minValue!.Contains( '.' ) ? minValue + ".00" : minValue!, "min", dataType ) ),
						new XAttribute( "max", GetValidationRange( configProfile, dataType == "double" && !maxValue!.Contains( '.' ) ? maxValue + ".00" : maxValue!, "max", dataType ) ),
						validationErrorType == "W" ? new XAttribute( "type", "warning" ) : null
					)
				);
			}
			else if ( dataType == "list" && validationErrorType == "W" )
			{
				field.Element( "Validation" )!.Add(
					new XElement( "List", new XAttribute( "type", "warning" ) )
				);
			}
		}

		return field;
	}

	private string GetValidationRange( XElement configProfile, string value, string rangeName, string dataType )
	{
		(string FieldName, string RangeValue) getProfileRange( string fieldName, string prompt, string title, string errorName )
		{
			var attr = configProfile.Elements( "xDataDef" ).Elements( "Profile" ).Elements( fieldName ).Elements( "Validation" ).Elements( "Date" ).FirstOrDefault()?.Attribute( rangeName );

			if ( attr == null )
			{
				var inputResult = InputBox.Show( prompt, title );

				if ( inputResult.ReturnCode != DialogResult.OK )
				{
					throw new ApplicationException( $"Unable to find the {errorName} field used in validation." );
				}

				fieldName = inputResult.Text;
				attr = configProfile.Elements( "xDataDef" ).Elements( "Profile" ).Elements( fieldName ).Elements( "Validation" ).Elements( "Date" ).FirstOrDefault()?.Attribute( rangeName );

				if ( attr == null )
				{
					throw new ApplicationException( $"Unable to find the {errorName} field used in validation." );
				}
			}

			return (fieldName, (string)attr);
		}

		string? dateValue = null;
		if ( value.Contains( "DOB" ) )
		{
			var (fieldName, rangeValue) = getProfileRange(
				dateBirthName,
				"Please enter the Date of Birth element name (date-birth was not found in Flat Data):",
				"Date of Birth Element",
				"DOB"
			);
			dateBirthName = fieldName;
			dateValue = rangeValue;
		}
		else if ( value.Contains( "DOH" ) )
		{
			var (fieldName, rangeValue) = getProfileRange(
				dateHireName,
				"Please enter the Date of Hire element name (date-hire was not found in Flat Data):",
				"Date of Hire Element",
				"DOH"
			);
			dateHireName = fieldName;
			dateValue = rangeValue;
		}

		string getDateValue( string dateValue )
		{
			var isYear = dateValue.EndsWith( ".Year" );
			return isYear
				? dateValue + ".ToString()"
				: dateValue + ".ToString( \"yyyy-MM-dd\" )";
		}

		if ( value == "DOB" || value == "DOH" )
		{
			return dateValue!;
		}
		else if ( value.StartsWith( "DOB." ) || value.StartsWith( "DOH." ) )
		{
			var token = value[ 0..4 ];

			dateValue = dateValue!.Replace( "Evaluate:", "" ).Replace( ".ToString( \"yyyy-MM-dd\" )", "" );

			return DateTime.TryParse( dateValue, out var _ )
				? $"Evaluate:DateTime.{value.Replace( token, $"Parse( \"{dateValue}\" )." )}.ToString( \"yyyy-MM-dd\" )"
				: $"Evaluate:{value.Replace( token, dateValue + "." )}.ToString( \"yyyy-MM-dd\" )";
		}
		else if ( value == "Today" )
		{
			return getDateValue( "Evaluate:DateTime.Today" );
		}
		else if ( value.StartsWith( "Today." ) )
		{
			return getDateValue( $"Evaluate:DateTime.{value}" );
		}
		else if ( value.StartsWith( "Evaluate:" ) )
		{
			return getDateValue( value );
		}
		else if ( value.StartsWith( "new DateTime" ) )
		{
			return getDateValue( $"Evaluate:{value}" );
		}
		else if ( DateTime.TryParse( value, out var date ) && dataType == "date" )
		{
			return date.ToString( "yyyy-MM-dd" );
		}
		else
		{
			return value;
		}
	}

	private XElement[] InitializeReports( XElement configProfile )
	{
		var reports = workbook.GetWorksheet( "Reports" );
		if ( reports == null ) return Array.Empty<XElement>();

		var specReports = new List<XElement>();
		var columnConfiguration = GetColumnConfiguration<ReportColumnType>( reports.RangeOrNull( "A5" ) );
		var reportsData = reports.RangeOrNull( "A6" )!;

		var offset = 0;
		MSExcel.Range row;
		string? reportName;
		while ( !string.IsNullOrEmpty( reportName = columnConfiguration.GetValue( ReportColumnType.ReportName, row = reportsData.Offset[ offset, 0 ] ) ) )
		{
			string? include;
			if ( ( include = columnConfiguration.GetValue( ReportColumnType.Include, row ) ) != "N" )
			{
				var evolutionId = columnConfiguration.GetValue( ReportColumnType.EvolutionId, row );
				var specificationId = columnConfiguration.GetValue( ReportColumnType.SpecExportID, row );

				if ( string.IsNullOrEmpty( evolutionId ) )
				{
					// Old spec format where column headers are ID/Export ID
					evolutionId = row.Offset[ 0, 2 ].GetText();
					specificationId = row.Offset[ 0, 1 ].GetText();
				}

				if ( !string.IsNullOrEmpty( evolutionId ) && !string.IsNullOrEmpty( specificationId ) )
				{
					var excelExport = GetExcelExportInfo( reportName, evolutionId, specificationId, columnConfiguration, row, include );
					configProfile.Element( "ExcelExports" )!.Add( excelExport );
				}

				var report = GetReportConfiguration( reportName, evolutionId, columnConfiguration, row );

				if ( report != null )
				{
					specReports.Add( report );
				}
			}

			offset++;
		}

		return specReports.ToArray();
	}

	private XElement? GetReportConfiguration( string reportName, string evolutionId, Dictionary<ReportColumnType, ColumnDefinition> columnConfiguration, MSExcel.Range row )
	{
		if ( reportName == "Listing" ) return null;

		var id = string.IsNullOrEmpty( evolutionId )
			? reportName.Replace( "/", " " ).Replace( "\\", " " ).ToProperCase().Replace( " ", "" )
			: evolutionId;
		var category = columnConfiguration.GetValue( ReportColumnType.ReportCategory, row )!;
		var filter = columnConfiguration.GetValue( ReportColumnType.Filter, row )!;
		var rowIndexName = columnConfiguration.GetValue( ReportColumnType.RowIndexName, row );
		var columnIndexName = columnConfiguration.GetValue( ReportColumnType.ColumnIndexName, row );

		string? rowIndexLookup;
		string? rowIndexHeader;
		string? adHocResultTableName;
		string? delimiter;

		var adHocSortColumn = columnConfiguration.GetValue( ReportColumnType.SortColumn, row );
		var adHocSortDescending = string.Compare( ( columnConfiguration.GetValue( ReportColumnType.SortColumn, row ) + "A" )[ 0..1 ], "D", true ) == 0;
		var adHocSortType = ( columnConfiguration.GetValue( ReportColumnType.SortType, row ) + "S" )[ 0..1 ];
		var adHocFolderItemType = columnConfiguration.GetValue( ReportColumnType.FolderItemType, row );

		var ftpUrl = columnConfiguration.GetValue( ReportColumnType.FtpUrl, row );
		var ftpUserName = columnConfiguration.GetValue( ReportColumnType.FtpUserName, row );
		var ftpPassword = columnConfiguration.GetValue( ReportColumnType.FtpPassword, row );
		var ftpNotifications = columnConfiguration.GetValue( ReportColumnType.FtpNotifications, row );

		var fileName = columnConfiguration.GetValue( ReportColumnType.FileName, row );
		var extension = columnConfiguration.GetValue( ReportColumnType.Extension, row );
		var isFileNameEvaluate = fileName?.StartsWith( "Evaluate:" ) ?? false;

		var reportType =
			!string.IsNullOrEmpty( rowIndexName ) ? string.IsNullOrEmpty( columnIndexName ) ? "IndexByIndexCount" : "IndexCount" :
			columnConfiguration.GetValue( ReportColumnType.SpecExportCustomProcess, row ) != "Y" ? "SpecificationReport" :
			new[] { "ExcelNotes", "NoteExport" }.Contains( id ) ? "NoteExport" :
			new[] { "ExcelDataAudits", "DataAuditExport" }.Contains( id ) ? "DataAuditExport" :
			new[] { "ExcelCalculationAudit", "CalculationAuditExport" }.Contains( id ) ? "CalculationAuditExport" :
			!string.IsNullOrEmpty( adHocFolderItemType ) ? "AdHocReport" :
			null;

		var submitPagePath = columnConfiguration.GetValue( ReportColumnType.SubmitPagePath, row );

		var clientName =
			workbook.Name.Split( ' ' )[ 0 ] // when downloaded, could be MHA-Spec-Aramark (#).xlsx
				.Split( '-' )[ 2 ] // MHA-Spec-Aramark -> Aramark
				.Split( '.' )[ 0 ]; // In case client has mulitple sites and spec has a .Suffix for site type...

		var report =
			new XElement( "Report",
				new XAttribute( "ID", id ),
				new XAttribute( "Name", reportName ),
				new XAttribute( "Description", columnConfiguration.GetValue( ReportColumnType.Description, row )! ),
				!string.IsNullOrEmpty( rowIndexName ) && !string.IsNullOrEmpty( fileName ) ? new XAttribute( "Title", isFileNameEvaluate ? fileName : Path.GetFileNameWithoutExtension( fileName ) ) : null,
				!string.IsNullOrEmpty( rowIndexName ) && !string.IsNullOrEmpty( fileName ) && !isFileNameEvaluate ? new XAttribute( "Extension", !string.IsNullOrEmpty( extension ) ? extension : Path.GetExtension( fileName )[ 1.. ] ) : null,

				!string.IsNullOrEmpty( category ) ? new XAttribute( "Category", category ) : null,
				!string.IsNullOrEmpty( filter ) ? new XAttribute( "Filter", filter ) : null,

				!string.IsNullOrEmpty( reportType ) ? new XAttribute( "LegacyReportType", reportType ) : null,
				string.IsNullOrEmpty( reportType ) ? new XAttribute( "ClientReportType", $"Administration.Hangfire.{clientName}.Report{id}, Administration.{clientName}" ) : null,

				!string.IsNullOrEmpty( submitPagePath ) ? new XAttribute( "SubmitPagePath", submitPagePath ) : null,
				!string.IsNullOrEmpty( rowIndexName ) ? new XAttribute( "Page", id ) : null,

				!string.IsNullOrEmpty( rowIndexName ) && !string.IsNullOrEmpty( columnIndexName )
					? new XElement( "IndexByIndexCountParameters",
						new XAttribute( "RowIndexName", rowIndexName ),
						new XAttribute( "ColumnIndexName", columnIndexName )
					)
					: null,

				!string.IsNullOrEmpty( rowIndexName ) && string.IsNullOrEmpty( columnIndexName )
					? new XElement( "IndexCountParameters",
						new XAttribute( "IndexField", rowIndexName ),
						!string.IsNullOrEmpty( rowIndexLookup = columnConfiguration.GetValue( ReportColumnType.RowIndexLookup, row ) ) ? new XAttribute( "IndexLookup", rowIndexLookup ) : null,
						!string.IsNullOrEmpty( rowIndexHeader = columnConfiguration.GetValue( ReportColumnType.RowIndexHeader, row ) ) ? new XAttribute( "CountHeader", rowIndexHeader ) : null
					)
					: null,

				!string.IsNullOrEmpty( adHocFolderItemType )
					? new XElement( "AdHocReportParameters",
						new XAttribute( "FolderItemType", adHocFolderItemType ),
						new XAttribute( "CalcEngineTable", !string.IsNullOrEmpty( adHocResultTableName = columnConfiguration.GetValue( ReportColumnType.ResultTableName, row ) ) ? adHocResultTableName : adHocFolderItemType ),
						!string.IsNullOrEmpty( delimiter = columnConfiguration.GetValue( ReportColumnType.Delimiter, row ) ) ? new XAttribute( "Delimiter", delimiter ) : null,
						!string.IsNullOrEmpty( adHocSortColumn ) ? new XAttribute( "SortColumn", adHocSortColumn ) : null,
						!string.IsNullOrEmpty( adHocSortColumn ) ? new XAttribute( "SortDirection", adHocSortDescending ? "Descending" : "Ascending" ) : null,
						!string.IsNullOrEmpty( adHocSortColumn ) && string.Compare( adHocSortType, "N", true ) == 0 ? new XAttribute( "SortType", "Number" ) : null,
						!string.IsNullOrEmpty( adHocSortColumn ) && string.Compare( adHocSortType, "D", true ) == 0 ? new XAttribute( "SortType", "Date" ) : null
					)
					: null,

				// moved to web.config later...
				!string.IsNullOrEmpty( ftpUrl ) && !string.IsNullOrEmpty( ftpUserName ) && !string.IsNullOrEmpty( ftpPassword )
					? new XElement( "reportFTPSettings",
						new XElement( "add",
							new XAttribute( "key", id ),
							new XAttribute( "url", ftpUrl ),
							new XAttribute( "username", ftpUserName ),
							new XAttribute( "password", ftpPassword ),
							!string.IsNullOrEmpty( ftpNotifications ) ? new XAttribute( "emailNotifications", ftpNotifications ) : null
						)
					)
					: null
			);

		return report;
	}

	private static void UpdateExcelExportReports( XElement[] excelExports, Dictionary<string, ColumnDefinition> exportColumns, string fieldName, MSExcel.Range row, string? historyTable = null )
	{
		static string? getIndexRange( string? value )
		{
			return value?[ 0..Math.Min( value.Length, 6 ) ] switch
			{
				"Today" => "Evaluate:DateTime.Today.Year",
				"Today." => $"Evaluate:DateTime.{value}",
				"" => null,
				_ => value
			};
		}

		foreach ( var export in excelExports )
		{
			var exportColumnName = (string)export.Attribute( "specification-id" )!;
			var includeInfo = exportColumns.GetValue( exportColumnName, row ).EnsureNonBlank() ?? "N";

			if ( string.Compare( includeInfo[ 0..1 ], "Y", true ) == 0 )
			{
				var source = (string?)export.Attribute( "source" );

				var searchIndex = includeInfo.GetSwitchValue( "search-index" );
				var specificIndex = getIndexRange( includeInfo.GetSwitchValue( "index" ) );
				var min = getIndexRange( includeInfo.GetSwitchValue( "min" ) );
				var max = getIndexRange( includeInfo.GetSwitchValue( "max" ) );
				var order = includeInfo.GetSwitchValue( "order" );
				var format = includeInfo.GetSwitchValue( "format" );
				var label = includeInfo.GetSwitchValue( "label" );
				var start = getIndexRange( includeInfo.GetSwitchValue( "start" ) );
				var end = getIndexRange( includeInfo.GetSwitchValue( "end" ) );
				var period = includeInfo.GetSwitchValue( "period" );

				var isSearchListing = source == "SearchListing";

				var column =
					new XElement( "Column",
						!string.IsNullOrEmpty( order ) ? new XAttribute( "order", order ) : null,
						!string.IsNullOrEmpty( format ) ? new XAttribute( "format", format ) : null,
						new XAttribute( "source", isSearchListing && !string.IsNullOrEmpty( searchIndex ) ? searchIndex.Replace( " ", "_" ) : fieldName ),
						// If a 'history listing report' and this is 'flat data', put an attribute is-flat="true" on there
						// so the Report Process knows not to append the 'history-type' to this column.
						string.IsNullOrEmpty( historyTable ) && export.Element( "HistoryRowType" ) != null ? new XAttribute( "is-flat", true ) : null,
						label.EnsureNonBlank()
					);

				if ( !string.IsNullOrEmpty( historyTable ) && !isSearchListing )
				{
					string? minIndex = null;
					XElement? historyContainer = null;

					var historyData = export.Elements( "HistoryData" ).Where( h => (string?)h.Attribute( "type" ) == historyTable ); ;

					if ( includeInfo.StartsWith( "Y*", StringComparison.OrdinalIgnoreCase ) )
					{
						historyContainer = historyData.FirstOrDefault( h => (string?)h.Attribute( "min-index" ) == "*" );
						minIndex = "*";
					}
					else if ( !string.IsNullOrEmpty( min ) && !string.IsNullOrEmpty( max ) )
					{
						historyContainer = historyData.FirstOrDefault( h => (string?)h.Attribute( "min-index" ) == min && (string?)h.Attribute( "max-index" ) == max );
						minIndex = min;
					}
					else if ( !string.IsNullOrEmpty( min ) )
					{
						historyContainer = historyData.FirstOrDefault( h => (string?)h.Attribute( "min-index" ) == min );
						minIndex = min;
					}
					else if ( !string.IsNullOrEmpty( searchIndex ) )
					{
						historyContainer = historyData.FirstOrDefault( h => (string?)h.Attribute( "xds-index" ) == searchIndex );
					}
					else if ( !string.IsNullOrEmpty( specificIndex ) )
					{
						historyContainer = historyData.FirstOrDefault( h => (string?)h.Attribute( "history-index" ) == specificIndex );
					}
					else
					{
						historyContainer = historyData.FirstOrDefault( h => (string?)h.Attribute( "min-year" ) == start && (string?)h.Attribute( "max-year" ) == end && ( string.IsNullOrEmpty( period ) || (string?)h.Attribute( "index-suffix" ) == period ) );
					}

					if ( historyContainer == null )
					{
						historyContainer =
							new XElement( "HistoryData",
								new XAttribute( "type", historyTable ),
								!string.IsNullOrEmpty( minIndex ) ? new XAttribute( "min-index", minIndex ) : null,
								!string.IsNullOrEmpty( max ) ? new XAttribute( "max-index", max ) : null,
								!string.IsNullOrEmpty( searchIndex ) ? new XAttribute( "xds-index", searchIndex ) : null,
								!string.IsNullOrEmpty( specificIndex ) ? new XAttribute( "history-index", specificIndex ) : null,
								!string.IsNullOrEmpty( start ) && !string.IsNullOrEmpty( end ) ? new XAttribute( "min-year", start ) : null,
								!string.IsNullOrEmpty( start ) && !string.IsNullOrEmpty( end ) ? new XAttribute( "max-year", end ) : null,
								!string.IsNullOrEmpty( period ) ? new XAttribute( "index-suffix", period ) : null
							);

						// Only append if not 'row' report
						if ( export.Element( "HistoryRowType" ) == null )
						{
							// If only 'type' attribute, none of the conditions were processed correctly...
							// NOTE: Not really following logic for this statement, but leaving here since in MadHatterSpecExport.xlam
							if ( historyContainer.Attributes().Count() == 1 )
							{
								throw new ApplicationException( $"Invalid index setup for {exportColumnName} in {(string?)export.Attribute( "id" )} export." );
							}

							export.Add( historyContainer );
						}
					}

					// If NOT a history table dump, then get the right 'HistoryData' element and put child column in there,
					// otherwise simply put the column inside the export configuration element.
					if ( export.Element( "HistoryRowType" ) == null )
					{
						historyContainer.Add( column );
					}
					else
					{
						export.Add( column );
					}
				}
				else
				{
					export.Add( column );
				}
			}

		}
	}

	private XElement GetExcelExportInfo( string reportName, string evolutionId, string specificationId, Dictionary<ReportColumnType, ColumnDefinition> columnConfiguration, MSExcel.Range row, string? switches )
	{
		// Get reports initialized into configuration, but the actual columns (for Excel Export reports) will be added during field processing
		string? dataSource;
		var excelExport =
			new XElement( "ExcelExport",
				new XAttribute( "id", evolutionId ),
				new XAttribute( "specification-id", specificationId ),
				switches.GetSwitchValue( "export-headers" ) == "false" ? new XAttribute( "export-headers", "false" ) : null,
				!string.IsNullOrEmpty( dataSource = switches.GetSwitchValue( "source" ) ) ? new XAttribute( "source", dataSource ) : null,

				new XElement( "Title",
					switches.GetSwitchValue( "export-title" ) == "false" ? new XAttribute( "export", "false" ) : null,
					$"{workbook.RangeOrNull<string>( "Plan_Sponsor" )} - {reportName}"
				)
			);

		var historyTableInfo = columnConfiguration.GetValue( ReportColumnType.HistoryTableType, row );
		if ( !string.IsNullOrEmpty( historyTableInfo ) )
		{
			var historyTable = historyTableInfo.Split( '/' )[ 0 ];
			string? label;
			string? format;
			string? order;
			string? dateUpdatedLabel;

			excelExport.Add(
				new XElement( "HistoryRowType", historyTable ),
				historyTableInfo.GetSwitchValue( "exportAuthID" ) != "N"
					? new XElement( "Column",
						new XAttribute( "source", "hispAuthID" ),
						!string.IsNullOrEmpty( format = historyTableInfo.GetSwitchValue( "authIDFormat" ) ) ? new XAttribute( "format", format ) : null,
						!string.IsNullOrEmpty( order = historyTableInfo.GetSwitchValue( "authIDOrder" ) ) ? new XAttribute( "order", order ) : null,

						!string.IsNullOrEmpty( label = historyTableInfo.GetSwitchValue( "authIDLabel" ) )
							? label
							: $"AuthID/key/Table:{historyTable}"
					)
					: null,

				!string.IsNullOrEmpty( dateUpdatedLabel = historyTableInfo.GetSwitchValue( "dateUpdatedLabel" ) )
					? new XElement( "Column",
						new XAttribute( "source", "hisDateUpdated" ),
						!string.IsNullOrEmpty( order = historyTableInfo.GetSwitchValue( "dateUpdatedOrder" ) ) ? new XAttribute( "order", order ) : null,
						dateUpdatedLabel
					)
					: null
			);
		}

		return excelExport;
	}

	private static Dictionary<TEnum, ColumnDefinition> GetColumnConfiguration<TEnum>( MSExcel.Range? columnsStart ) where TEnum : struct, Enum
	{
		if ( columnsStart == null ) return new();

		var configuration = new Dictionary<TEnum, ColumnDefinition>();
		var colOffset = 0;
		string? colInfo;
		var enumType = typeof( TEnum );
		var colLookups =
			Enum.GetValues<TEnum>()
				.Select( t => new { Enum = t, Column = enumType.GetMember( t.ToString() ).Last().GetCustomAttribute<DisplayAttribute>()!.Name! } )
				.ToDictionary( t => t.Column, t => t.Enum );

		while ( !string.IsNullOrEmpty( colInfo = columnsStart.Offset[ 0, colOffset ].GetText() ) )
		{
			var colName = colInfo.Split( '/' )[ 0 ];

			// Find enum by matching the Display attribute Name to colName
			if ( colLookups.TryGetValue( colName, out var columnType ) )
			{
				configuration[ columnType ] = GetColumnConfiguration( colName, colOffset );
			}

			colOffset++;
		}

		return configuration;
	}

	private static Dictionary<string, ColumnDefinition> GetColumnConfiguration( MSExcel.Range? columnsStart, string[] columnNames )
	{
		if ( columnsStart == null ) return new();

		var configuration = new Dictionary<string, ColumnDefinition>();
		var colOffset = 0;
		string? colInfo;

		while ( !string.IsNullOrEmpty( colInfo = columnsStart.Offset[ 0, colOffset ].GetText() ) )
		{
			var colName = colInfo.Split( '/' )[ 0 ];

			if ( columnNames.Contains( colName ) )
			{
				configuration[ colName ] = GetColumnConfiguration( colName, colOffset );
			}

			colOffset++;
		}

		return configuration;
	}

	private static ColumnDefinition GetColumnConfiguration( string colName, int colOffset )
	{
		var isListingConfiguration = colName.StartsWith( "Listing:" );
		var IsLabelOverride = colName.StartsWith( "Label." );

		return new ColumnDefinition(
			colName,
			colOffset,
			isListingConfiguration,
			isListingConfiguration ? colName[ ( colName.IndexOf( ":" ) + 1 ).. ] : null,
			IsLabelOverride,
			IsLabelOverride ? colName[ ( colName.IndexOf( "." ) + 1 ).. ] : null
		);
	}

	private static string Pluralize( string name ) =>
		string.Compare( name.Pluralize(), name, true ) == 0
			? name.EndsWith( "s", StringComparison.InvariantCultureIgnoreCase ) ? name + "es" : name + "s"
			: name.Pluralize();
}