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
                string dclist = null;
                string gcName = null;
                string customFilter = null;
                string customPropList = null;

                try
                {
                    // Are there any named parameters?
                    int namedParamIndex = -1;
                    for (int x = 0; x < args.Length; x++)
                    {
                        if (args[x].StartsWith("-") || args[x].StartsWith("/"))
                        {
                            namedParamIndex = x;
                        }
                    }

                    if (namedParamIndex > -1)
                    {
                        for (int x = namedParamIndex; x < args.Length; x++)
                        {
                            if (args[x].ToLower().TrimStart(new char[] { '-', '/' }) == "dclist")
                            {
                                dclist = args[x + 1];
                            }
                        }
                    }

                    // GC should be the first positional parameter
                    if (namedParamIndex > 0 || namedParamIndex < 0)
                        gcName = args[0];

                    // Filter should be the second positional parameter
                    if (args.Length > 1 && (namedParamIndex > 1 || namedParamIndex < 0))
                        customFilter = args[1];

                    // Prop list should be the third positional parameter
                    if (args.Length > 2 && (namedParamIndex > 2 || namedParamIndex < 0))
                        customPropList = args[2];
                }
                catch
                {
                    throw new Exception("Error parsing command line arguments.");
                }

                Application.Run(new Form1(gcName, customFilter, customPropList, dclist));
            }
		}
	}
}
