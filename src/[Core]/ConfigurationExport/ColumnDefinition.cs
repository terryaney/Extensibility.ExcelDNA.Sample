namespace KAT.Camelot.Extensibility.Excel.AddIn.ConfigurationExport;

record ColumnDefinition( string Name, int Offset, bool IsListingConfiguration, string? ListingId, bool IsLabelOverride, string? LabelArea );