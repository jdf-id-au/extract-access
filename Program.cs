using System;
using System.IO;
using Microsoft.Office.Interop.Access;

namespace ExtractAccess
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            String mdb_path, out_path, forms_path, macros_path, modules_path, reports_path;
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

            Application app = new ApplicationClass();
            foreach (var mdb in args)
            {
                mdb_path = Path.GetFullPath(mdb);
                out_path = Path.GetDirectoryName(mdb_path) + @"\" + Path.GetFileNameWithoutExtension(mdb_path) + @"\";
                forms_path = out_path + @"forms\";
                macros_path = out_path + @"macros\";
                modules_path = out_path + @"modules\";
                reports_path = out_path + @"reports\";
                
                Console.WriteLine("Opening " + mdb_path);
                Directory.CreateDirectory(out_path);

                // after https://stackoverflow.com/questions/50816715/extract-vba-codea-from-access-via-c-sharp
                app.OpenCurrentDatabase(mdb_path);

                if (app.CurrentProject.AllForms.Count > 0) Directory.CreateDirectory(forms_path);
                for (int i = 0; i < app.CurrentProject.AllForms.Count; i++)
                {
                    var form = app.CurrentProject.AllForms[i];
                    Console.WriteLine("Saving form: " + form.FullName);
                    app.SaveAsText(AcObjectType.acForm, form.FullName, forms_path + form.FullName + ".txt");
                }

                if (app.CurrentProject.AllMacros.Count > 0) Directory.CreateDirectory(macros_path);
                for (int i = 0; i < app.CurrentProject.AllMacros.Count; i++)
                {
                    var macro = app.CurrentProject.AllMacros[i];
                    Console.WriteLine("Saving macro: " + macro.FullName);
                    app.SaveAsText(AcObjectType.acMacro, macro.FullName, macros_path + macro.FullName + ".txt");
                }

                if (app.CurrentProject.AllModules.Count > 0) Directory.CreateDirectory(modules_path);
                for (int i = 0; i < app.CurrentProject.AllModules.Count; i++)
                {
                    var module = app.CurrentProject.AllModules[i];
                    Console.WriteLine("Saving module: " + module.FullName);
                    app.SaveAsText(AcObjectType.acModule, module.FullName, modules_path + module.FullName + ".bas");
                }

                if (app.CurrentProject.AllReports.Count > 0) Directory.CreateDirectory(reports_path);
                for (int i = 0; i < app.CurrentProject.AllReports.Count; i++)
                {
                    var report = app.CurrentProject.AllReports[i];
                    Console.WriteLine("Saving report: " + report.FullName);
                    app.SaveAsText(AcObjectType.acReport, report.FullName, reports_path + report.FullName + ".txt");
                }
                
                app.CloseCurrentDatabase();
            }
            app.Quit(AcQuitOption.acQuitSaveNone);
        }
    }
}