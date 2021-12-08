using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Windows.Media;
using System.Reflection;

namespace CreateSpacesFromLinkedRooms
{
    [Transaction(TransactionMode.ReadOnly)]

    public class App : IExternalApplication
    {
        public static string assemblyPath = typeof(App).Assembly.Location;
        public static string mainPath = Path.GetDirectoryName(assemblyPath);
        RibbonPanel ribbonPanel;

        public Result OnStartup(UIControlledApplication application)
        {
            string tab_name = "SynSys";
            string panel_name = "Create";
            string button_name = "CreateSpacesFromLinkedRooms";
            string button_text = "Spaces from\nLinked Rooms";
            string button_class_name = button_name + ".Command";
            string button_tool_tip = "Создание пространств и помещений в модели MEP на основе комнат из присоединённых моделей.\n" +
                                    $"v{typeof(App).Assembly.GetName().Version}";

            CreateRibbonTab(application, tab_name);
            CreateRibbonPanel(application, tab_name, panel_name);
            CreateButton(button_name, button_text, button_class_name, button_tool_tip);

            return Result.Succeeded;
        }

        private static void CreateRibbonTab(UIControlledApplication application, string tab_name)
        {
            try
            {
                application.CreateRibbonTab(tab_name);
            }
            catch
            {

            }
        }

        private void CreateRibbonPanel(UIControlledApplication application, string tab_name, string panel_name)
        {
            try
            {
                ribbonPanel = application.CreateRibbonPanel(tab_name, panel_name);
            }
            catch
            {
                List<RibbonPanel> ribbonPanels = application.GetRibbonPanels(tab_name);
                foreach (RibbonPanel rPanel in ribbonPanels)
                {
                    if (rPanel.Name == panel_name)
                    {
                        ribbonPanel = rPanel;
                    }
                }
            }
        }

        private void CreateButton(string button_name, string button_text, string button_class_name, string button_tool_tip)
        {
            PushButtonData buttonData = new PushButtonData(button_name, button_text, assemblyPath, button_class_name);
            PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;
            pushButton.ToolTip = button_tool_tip;
            pushButton.LargeImage = PngImageSource("CreateSpacesFromLinkedRooms.Resources.32x32_SpacesFromLinkedRooms.png");
            //pushButton.Image = PngImageSource("CreateSpacesFromLinkedRooms.Resources.16x16_SpacesFromLinkedRooms.png");
        }

        private ImageSource PngImageSource(string embeddedPathname)
        {
            Stream st = this.GetType().Assembly.GetManifestResourceStream(embeddedPathname);
            PngBitmapDecoder decoder = new PngBitmapDecoder(st, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);
            return decoder.Frames[0];
        }


        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
