﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Windows.Forms;
using System.IO;
using CommonSnappableTypes;

namespace MyExtendableApp
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            LoadSnapin();
            Console.ReadLine();
        }

        static void LoadSnapin()
        {
            // Allow user to select an assembly to load.
            OpenFileDialog dlg = new OpenFileDialog
            {
                //set the initial directory to the path of this project
                InitialDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                Filter = "assemblies (*.dll)|*.dll|All files (*.*)|*.*",
                FilterIndex = 1
            };

            if (dlg.ShowDialog() != DialogResult.OK)
            {
                Console.WriteLine("User cancelled out of the open file dialog.");
                return;
            }

            if (dlg.FileName.Contains("CommonSnappableTypes"))
                Console.WriteLine("CommonSnappableTypes has no snap-ins!");
            else if (!LoadExternalModule(dlg.FileName))
                Console.WriteLine("Nothing implements IAppFunctionality!");
        }


        private static bool LoadExternalModule(string path)
        {
            bool foundSnapIn = false;
            Assembly theSnapInAsm = null;
            try
            {
                // Dynamically load the selected assembly.
                theSnapInAsm = Assembly.LoadFrom(path);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred loading the snapin: {ex.Message}");
                return foundSnapIn;
            }
            // Get all IAppFunctionality compatible classes in assembly.
            var theClassTypes = from t in theSnapInAsm.GetTypes()
                                where t.IsClass && (t.GetInterface("IAppFunctionality") != null)
                                select t;
            // Now, create the object and call DoIt() method.
            foreach (Type t in theClassTypes)
            {
                foundSnapIn = true;
                // Use late binding to create the type.
                IAppFunctionality itfApp = (IAppFunctionality)theSnapInAsm.CreateInstance(t.FullName,
                true);
                itfApp?.DoIt();
                //lstLoadedSnapIns.Items.Add(t.FullName);
                // Show company info.
                DisplayCompanyData(t);
            }
            return foundSnapIn;
        }


        private static void DisplayCompanyData(Type t)
        {
            // Get [CompanyInfo] data.
            var compInfo = from ci in t.GetCustomAttributes(false)
                           where (ci is CompanyInfoAttribute)
                           select ci;
            // Show data.
            foreach (CompanyInfoAttribute c in compInfo)
            {
                Console.WriteLine($"More info about {c.CompanyName} can be found at {c.CompanyUrl}");
            }
        }


    }
}
