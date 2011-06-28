using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace OABValidate
{
	static class Program
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main(string[] args)
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length < 1)
            {
                Application.Run(new Form1());
            }
            else
            {
                Application.Run(new Form1(args[0]));
            }
		}
	}
}
