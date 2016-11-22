#region Namespaces
using System;
using Autodesk.Revit.UI;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.IO;
using Autodesk.Revit.DB.Events;
#endregion

namespace Revit3Drooms
{
    class App : IExternalApplication
    {
        /// <summary>
        /// Add buttons for our command
        /// to the ribbon panel.
        /// </summary>
        /// 
        
        void PopulatePanel(RibbonPanel p)
        {
            string path = Assembly.GetExecutingAssembly().Location;
            PushButton Rooms3D = p.AddItem(new PushButtonData("Revit3Drooms_Create3Drooms", "Rooms 3D", path, "Revit3Drooms.Create3Drooms")) as PushButton;
            Rooms3D.LargeImage = Images.getBitmap(Revit3Drooms.Properties.Resources.view);
            Rooms3D.ToolTip = "Information about the tool";
            
        }


        public class Images
        {
            public static BitmapImage getBitmap(System.Drawing.Bitmap image)
            {
                MemoryStream stream = new MemoryStream();
                image.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                BitmapImage bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.StreamSource = stream;
                bmp.EndInit();
                return bmp;
            }

        }


        public Result OnStartup(UIControlledApplication a)
        {

            a.ControlledApplication.ApplicationInitialized += OnApplicationInitialized;
            
            // Create a custom ribbon tab
            String tabName = "3D Rooms";
            a.CreateRibbonTab(tabName);

            PopulatePanel(a.CreateRibbonPanel(tabName, "Tools"));
            
            return Result.Succeeded;


        }


        void OnApplicationInitialized(object sender, ApplicationInitializedEventArgs e)
        {

        }

        public Result OnShutdown(UIControlledApplication a)
        {
            
            return Result.Succeeded;

        }
    }

    
}
