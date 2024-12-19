using Autodesk.Revit.UI;
using RevitPluginsApp.Plugin.ClashManagement;
using RevitPluginsApp.Plugin.PinningElements;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;

namespace RevitPluginsApp.Plugin
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
            string iconsDirectoryPath = Path.GetDirectoryName(assemblyLocation) + @"\Icons\";

            string tabName = "WildBIM";
            application.CreateRibbonTab(tabName);

            RibbonPanel commonPanel = application.CreateRibbonPanel(tabName, "Общее");

            PushButtonData clashIndicatorPlacementButton = new PushButtonData(nameof(ClashIndicatorPlacementCmd), "Размещение индикатора", assemblyLocation, typeof(ClashIndicatorPlacementCmd).FullName)
            {
                LargeImage = new BitmapImage(new Uri(iconsDirectoryPath + "ClashIndicatorPlacementCmd.png"))
            };

            commonPanel.AddItem(clashIndicatorPlacementButton);

            PushButtonData PinElementsButton = new PushButtonData(nameof(PinElementsCmd), "Закрепление элементов", assemblyLocation, typeof(PinElementsCmd).FullName)
            {
                LargeImage = new BitmapImage(new Uri(iconsDirectoryPath + "PinElementsCmd.png"))
            };

            commonPanel.AddItem(PinElementsButton);

            return Result.Succeeded;
        }
    }
}
