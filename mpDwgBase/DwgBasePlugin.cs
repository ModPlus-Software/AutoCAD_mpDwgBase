namespace mpDwgBase
{
    using System.IO;
    using Autodesk.AutoCAD.Runtime;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;

    /// <summary>
    /// Точка входа в команду
    /// </summary>
    public class DwgBasePlugin
    {
        /// <summary>
        /// Plugin start
        /// </summary>
        [CommandMethod("ModPlus", "mpDwgBase", CommandFlags.Session)]
        public void Start()
        {
#if !DEBUG
            Statistic.SendCommandStarting(ModPlusConnector.Instance);
#endif
            try
            {
                MoveDwgBase();

                // Директория расположения базы создается при первом обращении к свойству Constants.DwgBaseDirectory!

                var win = new MpDwgBaseMainWindow();
                AcApp.ShowModalWindow(AcApp.MainWindow.Handle, win, false);
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private static void MoveDwgBase()
        {
            var d = Path.Combine(Constants.CurrentDirectory, "Data", "DwgBase");
            if (Directory.Exists(d))
            {
                try
                {
                    // Папку назначения нужно удалить, чтобы не было исключения
                    var dwgBaseDirectory = Constants.DwgBaseDirectory;
                    if (Directory.Exists(dwgBaseDirectory))
                    {
                        Directory.Delete(dwgBaseDirectory, true);
                    }

                    // https://stackoverflow.com/a/38370485/4944499
                    Microsoft.VisualBasic.FileIO.FileSystem.MoveDirectory(d, dwgBaseDirectory);
                }
                catch
                {
                    // ignore
                }
            }
        }
    }
}
