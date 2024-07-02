using KAT.Camelot.Domain.Extensions;
using System.Text.Json.Nodes;

namespace KAT.Camelot.Extensibility.Excel.AddIn;

internal partial class NavigateToTable : Form
{
	private readonly JsonObject windowConfiguration;

	public NavigateToTable( List<NavigationTable> tables, string? currentSheet, JsonObject? windowConfiguration )
	{
		InitializeComponent();

		availableTables.BeginUpdate();
		availableTables.Nodes.Clear();

		var sheetTables = tables
			.Select( t => t.SheetName )
			.Distinct()
			.ToArray();

		TreeNode createTableNode( NavigationTable t ) => new()
		{
			Name = t.Name,
			Text = t.Name,
			ToolTipText = t.Description,
			ImageKey = "Table",
			SelectedImageKey = "Table",
			Tag = $"{t.SheetName}!{t.Address}"
		};

		if ( sheetTables.Length == 1 )
		{
			availableTables.Nodes.AddRange(
				tables
					.OrderBy( t => t.Name )
					.Select( createTableNode )
					.ToArray()
			);
		}
		else
		{
			availableTables.Nodes.AddRange(
				sheetTables
					.OrderBy( t => t )
					.Select( t => new TreeNode
					{
						Name = t,
						Text = t,
						ImageKey = "Sheet",
						SelectedImageKey = "Sheet",
						Tag = $"{t}!StartTables"
					} )
					.ToArray()
			);
		}

		void addTreeNodes( TreeNodeCollection nodes )
		{
			foreach( TreeNode n in nodes )
			{
				if ( n.ImageKey == "Sheet" )
				{
					n.Nodes.AddRange(
						tables
							.Where( t => t.SheetName == n.Name )
							.OrderBy( t => t.Name )
							.Select( createTableNode )
							.ToArray()
					);
					addTreeNodes( n.Nodes );
				}
				else
				{
					n.Nodes.AddRange(
						tables
							.First( t => t.Name == n.Name )
							.Columns.OrderBy( c => c.Name )
							.Select( c => new TreeNode
							{
								Name = c.Name,
								Text = c.Name + " (" + c.Address + ")",
								ImageKey = "Row",
								SelectedImageKey = "Row",
								Tag = $"{(n.Tag as string)!.Split( '!' )[0]}!{c.Address}"
							} )
							.ToArray()
					);
				}
			}
		}
		addTreeNodes( availableTables.Nodes );

		availableTables.ImageList = imageList;
		availableTables.EndUpdate();
		if ( currentSheet != null )
		{
			availableTables.SelectedNode = availableTables.Nodes.Find( currentSheet, false ).First();
			availableTables.SelectedNode.Expand();
			availableTables.SelectedNode.EnsureVisible();
		}

		this.windowConfiguration = windowConfiguration ?? new JsonObject();
	}

	public NavigationInfo? GetInfo()
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
		
		return new()
		{
			Target = (string)availableTables.SelectedNode.Tag,
			WindowConfiguration = windowConfiguration
		};
	}

	private void NavigateToTable_Load( object sender, EventArgs e )
	{
		WindowState = Enum.TryParse( (string?)windowConfiguration[ "state" ], out FormWindowState state) ? state : FormWindowState.Normal;
		Location = new Point { X = (int?)windowConfiguration[ "left" ] ?? Left, Y = (int?)windowConfiguration[ "top" ] ?? Top };
		Size = new Size { Width = (int?)windowConfiguration[ "width" ] ?? Width, Height = (int?)windowConfiguration[ "height" ] ?? Height };
	}

	private void Ok_Click( object sender, EventArgs e )
	{
		if ( availableTables.SelectedNode == null )
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

	private void AvailableTables_NodeMouseDoubleClick( object sender, TreeNodeMouseClickEventArgs e ) => Ok_Click( sender, e );
}
