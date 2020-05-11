using System;
using System.IO;
using Microsoft.Office.Interop.Access;

namespace ExtractAccess
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            String mdb_path;
            String out_path;
            foreach (var arg in args)
            {
                if (Path.GetExtension(arg) != ".mdb")
                {
                    Console.WriteLine("All arguments must be mdb files.");
                    return; // i.e. exit
                }
                mdb_path = Path.GetFullPath(arg);
                if (!File.Exists(mdb_path))
                {
                    Console.WriteLine("Couldn't find " + mdb_path);
                    return;
                }
            }
            foreach (var mdb in args)
            {
                mdb_path = Path.GetFullPath(mdb);
                out_path = Path.GetDirectoryName(mdb_path) + @"\" + Path.GetFileNameWithoutExtension(mdb_path) + @"\";
                Console.WriteLine("Opening " + mdb_path);
                Directory.CreateDirectory(out_path);

                // after https://stackoverflow.com/questions/50816715/extract-vba-codea-from-access-via-c-sharp
                Application app = new Application();
                app.OpenCurrentDatabase(mdb_path);

                for (int i = 0; i < app.CurrentProject.AllForms.Count; i++)
                {
                    var form = app.CurrentProject.AllForms[i];
                    Console.WriteLine("Saving form: " + form.FullName);
                    app.SaveAsText(AcObjectType.acForm, form.FullName, out_path + form.FullName + ".form.txt");
                }
                for (int i = 0; i < app.CurrentProject.AllMacros.Count; i++)
                {
                    var macro = app.CurrentProject.AllMacros[i];
                    Console.WriteLine("Saving macro: " + macro.FullName);
                    app.SaveAsText(AcObjectType.acMacro, macro.FullName, out_path + macro.FullName + ".macro.txt");
                }
                for (int i = 0; i < app.CurrentProject.AllModules.Count; i++)
                {
                    var module = app.CurrentProject.AllModules[i];
                    Console.WriteLine("Saving module: " + module.FullName);
                    app.SaveAsText(AcObjectType.acModule, module.FullName, out_path + module.FullName + ".module.bas");
                }
                for (int i = 0; i < app.CurrentProject.AllReports.Count; i++)
                {
                    var report = app.CurrentProject.AllReports[i];
                    Console.WriteLine("Saving report: " + report.FullName);
                    app.SaveAsText(AcObjectType.acReport, report.FullName, out_path + report.FullName + ".report.txt");
                }
                
                app.CloseCurrentDatabase();
            }
        }
    }
}