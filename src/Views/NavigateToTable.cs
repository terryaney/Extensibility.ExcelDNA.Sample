using KAT.Camelot.Domain.Extensions;
using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class NavigateToTable : Form
{
	private readonly ListViewColumnSorter lvwColumnSorter = new ();
	private readonly JsonObject windowConfiguration;

	public NavigateToTable( List<NavigationTable> tables, JsonObject? windowConfiguration )
	{
		InitializeComponent();

		availableTables.ListViewItemSorter = lvwColumnSorter;
		availableTables.Items.Clear();

		foreach ( var table in tables.OrderBy( t => t.Name ) )
		{
			var item = new ListViewItem( new[] { table.Name, table.Address, table.Description } ) { Tag = table };
			availableTables.Items.Add( item );
		}

		this.windowConfiguration = windowConfiguration ?? new JsonObject();
	}

	public NavigationInfo? GetTarget()
	{
		var dialogResult = ShowDialog();

		if ( dialogResult != DialogResult.OK )
		{
			return null;
		}

		windowConfiguration[ "state" ] = WindowState.ToString();

		if ( WindowState == FormWindowState.Normal )
		{
			windowConfiguration[ "top" ] = Location.Y;
			windowConfiguration[ "left" ] = Location.X;
			windowConfiguration[ "height" ] = Size.Height;
			windowConfiguration[ "width" ] = Size.Width;
		}

		windowConfiguration[ nameof( availableTables ) ] =
			new JsonObject().AddProperties(
				availableTables.Columns.Cast<ColumnHeader>().Select(
					c => new JsonKeyProperty( ( c.Tag as string )!, c.Width )
				)
			);

		return new()
		{
			Target = ( availableTables.SelectedItems[ 0 ].Tag as NavigationTable )!.Address,
			WindowConfiguration = windowConfiguration
		};
	}

	private void NavigateToTable_Load( object sender, EventArgs e )
	{
		WindowState = Enum.TryParse( (string?)windowConfiguration[ "state" ], out FormWindowState state) ? state : FormWindowState.Normal;
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };

		foreach ( ColumnHeader c in availableTables.Columns )
		{
			c.Width = (int?)windowConfiguration[ nameof( availableTables ) ]?[ ( c.Tag as string )! ] ?? c.Width;
		}
	}

	private void Ok_Click( object sender, EventArgs e )
	{
		if ( availableTables.SelectedItems.Count == 0 )
		{
			MessageBox.Show( "You must select one table to navigate to.", "Select Table" );
			return;
		}

		DialogResult = DialogResult.OK;
		Close();
	}

	private void Cancel_Click( object sender, EventArgs e )
	{
		DialogResult = DialogResult.Cancel;
		Close();
	}

	private void AvailableTables_ColumnClick( object sender, ColumnClickEventArgs e )
	{
		// Determine if clicked column is already the column that is being sorted.
		if ( e.Column == lvwColumnSorter.SortColumn )
		{
			// Reverse the current sort direction for this column.
			if ( lvwColumnSorter.Order == SortOrder.Ascending )
			{
				lvwColumnSorter.Order = SortOrder.Descending;
			}
			else
			{
				lvwColumnSorter.Order = SortOrder.Ascending;
			}
		}
		else
		{
			// Set the column number that is to be sorted; default to ascending.
			lvwColumnSorter.SortColumn = e.Column;
			lvwColumnSorter.Order = SortOrder.Ascending;
		}

		// Perform the sort with these new sort options.
		availableTables.Sort();
	}

	private void AvailableTables_MouseDoubleClick( object sender, MouseEventArgs e )
	{
		var item = availableTables.HitTest( e.X, e.Y );

		if ( null != item.Item )
		{
			Ok_Click( sender, e );
		}
	}
}
