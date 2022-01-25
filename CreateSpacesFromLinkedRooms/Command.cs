using System.IO;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using IronPython.Hosting;
using Microsoft.Scripting.Hosting;

namespace CreateSpacesFromLinkedRooms
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]

    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData revit, ref string message, ElementSet elements)
        {
            UIApplication ui_app = revit.Application;
            Application app = ui_app.Application;
            UIDocument ui_doc = ui_app?.ActiveUIDocument;
            Document doc = ui_doc?.Document;

            string mainModule = Path.Combine(Path.GetDirectoryName(typeof(Command).Assembly.Location), "main.py");
            ScriptEngine engine = Python.CreateEngine();
            ScriptScope scope = engine.CreateScope();

            scope.SetVariable("app", app);
            scope.SetVariable("doc", doc);
            engine.ExecuteFile(mainModule, scope);

            return Result.Succeeded;
        }
    }
}