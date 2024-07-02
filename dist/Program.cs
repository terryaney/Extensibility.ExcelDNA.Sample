using System;
using System.Diagnostics;
using System.Security.Principal;
using System.Windows.Forms;

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

			if ( !IsRunningAsAdmin() )
			{
				MessageBox.Show( "Please run this install with administrative priveleges.", "KAT Tools Installation", MessageBoxButtons.OK, MessageBoxIcon.Warning );
				return;
			}
			else if ( IsExcelRunning() )
			{
				MessageBox.Show( "Please shut down Excel before running this installation.", "KAT Tools Installation", MessageBoxButtons.OK, MessageBoxIcon.Warning );
				return;
			}
			Application.Run( new Install() );
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
}