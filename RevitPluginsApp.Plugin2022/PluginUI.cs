using Autodesk.Revit.UI;
using RevitPluginsApp.Plugin2022.ClashManagement;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace RevitPluginsApp.Plugin2022
{
    public class PluginUI : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            string assemblyLocation = Assembly.GetExecutingAssembly().Location;
            string iconsDirectoryPath = Path.GetDirectoryName(assemblyLocation) + @"\icons\";

            string tabName = "WildBIM";
            application.CreateRibbonTab(tabName);

            RibbonPanel clashManagementPanel = application.CreateRibbonPanel(tabName, "Управление коллизиями");

            PushButtonData clashIndicatorPlacementButton = new PushButtonData(nameof(ClashIndicatorPlacementCmd), "Размещение индикатора", assemblyLocation, typeof(ClashIndicatorPlacementCmd).FullName)
            {
                LargeImage = new BitmapImage(new Uri(iconsDirectoryPath + "ClashIndicatorPlacementCmd.png"))
            };

            clashManagementPanel.AddItem(clashIndicatorPlacementButton);

            return Result.Succeeded;
        }
    }
}
