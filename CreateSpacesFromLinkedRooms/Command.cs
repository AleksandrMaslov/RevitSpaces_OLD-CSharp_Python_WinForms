using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;
using System.Collections.Generic;

namespace CreateSpacesFromLinkedRooms
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData revit, ref string message, ElementSet elements)
        {
            //App.assemblyPath = typeof(App).Assembly.Location;
            //string sysPaths = Environment.GetEnvironmentVariable("PATH");
            //Environment.SetEnvironmentVariable(sysPaths, App.assemblyPath);

            var options = new Dictionary<string, object> { ["Debug"] = true };
            string pathMain = Path.Combine(App.mainPath, "main.py");

            ScriptEngine engine = Python.CreateEngine(options);
            ScriptScope scope = engine.CreateScope();

            UIApplication ui_app = revit.Application;
            Application app = ui_app.Application;
            UIDocument ui_doc = ui_app?.ActiveUIDocument;
            Document doc = ui_doc?.Document;

            scope.SetVariable("app", app);
            scope.SetVariable("doc", doc);
            engine.ExecuteFile(pathMain, scope);

            return Result.Succeeded;
        }
    }
}