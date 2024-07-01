using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Excel.AddIn.Setup
{
	public partial class Install : Form
    {
		private (bool IsInstalled, string Version, bool Is32Bit, string Path, string RegKeyName, string RegValueName) excelInfo;
		private string installPath;

		public Install()
        {
            InitializeComponent();
        }

		private void Install_Load( object sender, EventArgs e )
		{
			var assembly = Assembly.GetExecutingAssembly();

			using ( var s = assembly.GetManifestResourceStream( "Excel.AddIn.Setup.Resources.logo.png" ) )
			{
				logo.Image = Image.FromStream( s );
			}

			var isDotNet7Installed = IsDotNet7Installed();
			using ( var s = assembly.GetManifestResourceStream( $"Excel.AddIn.Setup.Resources.{( isDotNet7Installed ? "success" : "failed" )}.png" ) )
			{
				step1Check.Image = Image.FromStream( s );
			}

			excelInfo = GetExcelInstallationInfo();
			installPath = Path.Combine( excelInfo.Path, @"Library\KatTools" );

			ok.Enabled = isDotNet7Installed && excelInfo.IsInstalled;
			step2Label.Enabled = isDotNet7Installed;

			using ( var s = assembly.GetManifestResourceStream( $"Excel.AddIn.Setup.Resources.{( excelInfo.IsInstalled ? "success" : "failed" )}.png" ) )
			{
				step2Check.Image = Image.FromStream( s );
			}

			if ( !isDotNet7Installed )
			{
				step1Label.Text = ".NET 7 runtime is required.  Download now.";
				step1Label.Links.Add( step1Label.Text.Length - 13, 12, "https://dotnet.microsoft.com/en-us/download/dotnet/thank-you/runtime-desktop-7.0.20-windows-x64-installer" );
			}

			step2Label.Text = excelInfo.IsInstalled
				? "Detected Install Path: " + installPath
				: "Unable to detect Microsoft Office Excel installation path.";
		}

		static (bool IsInstalled, string Version, bool Is32Bit, string Path, string RegKeyName, string RegValueName) GetExcelInstallationInfo()
		{
			// Registry paths to check
			var officeKeys = new []
			{
				@"SOFTWARE\Microsoft\Office",
				@"SOFTWARE\WOW6432Node\Microsoft\Office", // For 32-bit applications on 64-bit Windows
			};
			var baseHives = new [] { RegistryHive.LocalMachine, RegistryHive.CurrentUser };

			foreach( var hive in baseHives )
			{
				using ( var baseKey = RegistryKey.OpenBaseKey( hive, RegistryView.Registry64 ) )
				{
					foreach ( var officePath in officeKeys )
					{
						using ( var officeKey = baseKey.OpenSubKey( officePath ) )
						{
							if ( officeKey != null )
							{
								foreach ( var versionKey in officeKey.GetSubKeyNames() )
								{
									// Check for Excel in each version key
									using ( var installRoot = officeKey.OpenSubKey( $@"{versionKey}\Excel\InstallRoot" ) )
									{
										if ( installRoot != null )
										{
											var regPathName = $@"{officePath}\{versionKey}\Excel\Options";

											var path = installRoot.GetValue( "Path" ) as string;
											var is32Bit = officePath.Contains( "WOW6432Node" );
											var regValueName = "OPEN";
											var xllName = is32Bit ? "KAT.Extensibility.Excel.x86.xll" : "KAT.Extensibility.Excel.xll";
											var xllFound = false;

											using ( var currentUser = RegistryKey.OpenBaseKey( RegistryHive.CurrentUser, RegistryView.Registry64 ) )
											using ( var options = currentUser.OpenSubKey( regPathName ) )
											{
												if ( options != null )
												{
													var counter = 0;

													// Check if the value exists and create a new value name if it does
													string openValue = null;
													while ( ( openValue = options.GetValue( regValueName ) as string ) != null )
													{
														if ( openValue.IndexOf( xllName, StringComparison.OrdinalIgnoreCase ) > -1 )
														{
															xllFound = true;
															break;
														}
														counter++;
														regValueName = "OPEN" + counter.ToString();
													}
												}

											}

											return (!string.IsNullOrEmpty( path ), versionKey, is32Bit, path, regPathName, !xllFound ? regValueName : null);
										}
									}
								}
							}
						}
					}
				}
			}


			// Excel not found
			return (false, null, false, null, null, null);
		}

		static bool IsDotNet7Installed()
		{
			try
			{
				var psi = new ProcessStartInfo
				{
					FileName = "dotnet",
					Arguments = "--list-runtimes",
					RedirectStandardOutput = true,
					UseShellExecute = false,
					CreateNoWindow = true,
				};

				using ( var process = Process.Start( psi ) )
				{
					using ( var reader = process.StandardOutput )
					{
						var output = reader.ReadToEnd();
						return output.Contains( "Microsoft.NETCore.App 7." );
					}
				}
			}
			catch ( Exception ex )
			{
				Console.WriteLine( "An error occurred while checking for .NET 7 runtime: " + ex.Message );
				return false;
			}
		}

		private void step1Label_LinkClicked( object sender, LinkLabelLinkClickedEventArgs e )
		{
			Process.Start( e.Link.LinkData as string );
			Close();
		}

		private void cancel_Click( object sender, EventArgs e )
		{
			Close();
		}

		private async void ok_Click( object sender, EventArgs e )
		{
			cancel.Enabled = ok.Enabled = step4Check.Enabled = false;
			step3Label.Enabled = true;

			string currentFile = null;
			var assembly = Assembly.GetExecutingAssembly();
			var updateUserSettings = true;

			void checkStep( PictureBox pb )
			{
				using ( var s = assembly.GetManifestResourceStream( $"Excel.AddIn.Setup.Resources.success.png" ) )
				{
					pb.Image = Image.FromStream( s );
				}
			}

			void failInstall( PictureBox pb, string message, Exception ex = null )
			{
				cancel.Enabled = true;
				using ( var s = assembly.GetManifestResourceStream( $"Excel.AddIn.Setup.Resources.failed.png" ) )
				{
					pb.Image = Image.FromStream( s );
				}
				MessageBox.Show( $"{message} {ex?.Message}" );
			}

			var xllName = excelInfo.Is32Bit ? "KAT.Extensibility.Excel.x86.xll" : "KAT.Extensibility.Excel.xll";

			try
			{
				var currentUser = WindowsIdentity.GetCurrent().Name;
				var fullControlRule = new FileSystemAccessRule( currentUser, FileSystemRights.FullControl, AccessControlType.Allow );

				if ( !Directory.Exists( installPath  ) )
				{
					Directory.CreateDirectory( installPath );
					var ds = Directory.GetAccessControl( installPath );
					ds.AddAccessRule( fullControlRule );
					Directory.SetAccessControl( installPath, ds );
				}

				if ( File.Exists( Path.Combine( installPath, xllName ) ) )
				{
					File.Delete( Path.Combine( installPath, xllName ) );
				}

				var filesToCopy = new[] { xllName, "appsettings.json" }
					.Select( f => Path.Combine( installPath, f ) )
					.Where( f => !File.Exists( f ) )
					.ToArray();

				updateUserSettings = filesToCopy.Any( f => f.EndsWith( "appsettings.json" ) );

				foreach ( var f in filesToCopy )
				{
					using ( var s = assembly.GetManifestResourceStream( $"Excel.AddIn.Setup.Resources.{currentFile = Path.GetFileName( f )}" ) )
					using ( var d = File.OpenWrite( f ) )
					{
						await s.CopyToAsync( d );

						var fs = File.GetAccessControl( f );
						fs.AddAccessRule( fullControlRule );
						File.SetAccessControl( f, fs );
					}
				}

				checkStep( step3Check );
			}
			catch ( Exception ex )
			{
				failInstall( step3Check, $"Unable to copy files. Current file: {currentFile}", ex );
				return;
			}

			if ( updateUserSettings )
			{
				step4Check.Enabled = step4Label.Enabled = true;

				var name = InputBox.Prompt( "Enter your KAT Management Site 'Save History' name in the format of first.last:", "Save History Name?" );
				var email = InputBox.Prompt( "Enter your KAT Management Site email address:", "KAT Management Site Username?", !string.IsNullOrEmpty( name ) ? $"{name}@conduent.com" : null );
				var fixedPath = Environment.GetFolderPath( Environment.SpecialFolder.Desktop );
				var dataExportPath = InputBox.Prompt( "Enter fixed xDS Data Export Path (if blank, exports to same location as file):", "Fixed xDS Export Path", fixedPath );
				var dataExportAppendDateToName = InputBox.Prompt( "Enter 'Y' to append the current date and time to the generated xDS export file name:", "Append Date to xDS Export Name", "N" );

				try
				{
					var settingsPath = Path.Combine( installPath, "appsettings.json" );
					var settingsContent =
						File.ReadAllText( settingsPath )
							.Replace( "{name}", name )
							.Replace( "{email}", email )
							.Replace( "{dataExport.path}", dataExportPath.Replace( "\\", "\\\\" ) )
							.Replace( "\"{dataExport.appendDateToName}\"", string.Compare( "Y", dataExportAppendDateToName, true ) == 0 ? "true" : "false" );

					File.WriteAllText( settingsPath, settingsContent );
				}
				catch ( Exception ex )
				{
					failInstall( step4Check, $"Unable to apply user settings.", ex );
					return;
				}
			}

			checkStep( step4Check );

			try
			{
				if ( !string.IsNullOrEmpty( excelInfo.RegValueName  ) )
				{
					step5Label.Enabled = true;
					using ( var baseKey = RegistryKey.OpenBaseKey( RegistryHive.CurrentUser, RegistryView.Registry64 ) )
					using ( var options = Registry.CurrentUser.OpenSubKey( excelInfo.RegKeyName, true ) )
					{
						if ( options != null )
						{
							options.SetValue( excelInfo.RegValueName, $"/R \"{installPath}\\{xllName}\"" );
						}
						else
						{
							failInstall( step5Check, $"Unable to register {xllName}.  Please use Excel -> Options to manually add the KAT Tools Add-in." );
						}
					}
				}
				checkStep( step5Check );
			}
			catch
			{
				failInstall( step5Check, $"Unable to register {xllName}.  Please use Excel -> Options to manually add the KAT Tools Add-in." );
			}

			MessageBox.Show( "The KAT Tools Add-in was successfully installed." );
			Close();
		}
	}
}
