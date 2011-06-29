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
                string gcName = args[0];
                string customFilter = null;
                string customPropList = null;
                if (args.Length > 1)
                {
                    customFilter = args[1];
                }

                if (args.Length > 2)
                {
                    customPropList = args[2];
                }

                Application.Run(new Form1(gcName, customFilter, customPropList));
            }
		}
	}
}
