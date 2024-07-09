using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Principal;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Excel.AddIn.Setup
{
	internal static class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault( false );

			var warnings = new List<string>();
			var excelInfo = GetExcelInstallationInfo();

			if ( !IsRunningAsAdmin() )
			{
				warnings.Add( "- Installation needs to be ran with administrative priveleges." );
			}
			
			if ( IsExcelRunning() )
			{
				warnings.Add( "- Shut down Excel before running this installation." );
			}
			
			if ( !excelInfo.IsInstalled )
			{
				warnings.Add( "- Microsoft Office Excel is not installed." );
			}
			else if ( excelInfo.IsRbleToolsInstalled )
			{
				warnings.Add( "- RBLe Tools / BTR.Extensibility.Excel.xll is still enabled.  Please disable before continuing." );
			}

			if ( warnings.Any() )
			{
				MessageBox.Show( "Please review the following prerequistes and run the installation again." + Environment.NewLine + Environment.NewLine + string.Join( Environment.NewLine + Environment.NewLine, warnings ), "KAT Tools Installation", MessageBoxButtons.OK, MessageBoxIcon.Warning );
				return;
			}

			Application.Run( new Install( excelInfo ) );
		}

		static ExcelInfo GetExcelInstallationInfo()
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
											var path = installRoot.GetValue( "Path" ) as string;
											var is32Bit = officePath.Contains( "WOW6432Node" );
											var regValueName = "OPEN";
											var xllName = is32Bit ? "KAT.Extensibility.Excel.x86.xll" : "KAT.Extensibility.Excel.xll";
											var xllFound = false;
											var rbleToolsInstalled = false;

											var regPathName = $@"{officePath}\{versionKey}\Excel\Options";

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
															
															if ( rbleToolsInstalled )
															{
																break;
															}
														}
														else if ( openValue.IndexOf( "BTR.Extensibility.Excel", StringComparison.OrdinalIgnoreCase ) > -1 )
														{
															rbleToolsInstalled = true;

															if ( xllFound )
															{
																break;
															}
														}
														counter++;
														regValueName = "OPEN" + counter.ToString();
													}
												}
											}

											return new ExcelInfo
											{
												IsInstalled = !string.IsNullOrEmpty( path ),
												Version = versionKey,
												Is32Bit = is32Bit,
												Path = path,
												RegKeyName = regPathName,
												RegValueName = !xllFound ? regValueName : null,
												IsRbleToolsInstalled = rbleToolsInstalled
											};
										}
									}
								}
							}
						}
					}
				}
			}

			return new ExcelInfo
			{
				IsInstalled = false
			};
		}

		static bool IsExcelRunning()
		{
			foreach ( var process in Process.GetProcesses() )
			{
				if ( process.ProcessName.ToUpperInvariant().Equals( "EXCEL" ) )
				{
					return true;
				}
			}
			return false;
		}

		static bool IsRunningAsAdmin()
		{
			using ( var identity = WindowsIdentity.GetCurrent() )
			{
				var principal = new WindowsPrincipal( identity );
				return principal.IsInRole( WindowsBuiltInRole.Administrator );
			}
		}
	}

	public class ExcelInfo
	{
		public bool IsInstalled { get; set; }
		public string Version { get; set; }
		public bool Is32Bit { get; set; }
		public string Path { get; set; }
		public string RegKeyName { get; set; }
		public string RegValueName { get; set; }
		public bool IsRbleToolsInstalled { get; set; }
	}
}