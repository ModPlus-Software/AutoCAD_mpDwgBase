namespace mpDwgBase
{
    using System.IO;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Windows.Data;
    using Models;
    using ModPlusAPI.Windows;
    using Utils;

    public static class ImageCreator
    {
#if !A2013
        public static string ImagePreviewFile(DwgBaseItem selectedItem, string baseFolder)
        {
            try
            {
                var imgFile = DwgBaseHelpers.FindImageFile(selectedItem, baseFolder);

                // Если файл есть
                if (!string.IsNullOrEmpty(imgFile))
                {
                    if (File.Exists(imgFile))
                    {
                        return imgFile;
                    }

                    return CreateImageFile(selectedItem, baseFolder);
                }

                return CreateImageFile(selectedItem, baseFolder);
            }
            catch (System.Exception exception)
            {
                ExceptionBox.Show(exception);
            }

            return string.Empty;
        }

        private static string CreateImageFile(DwgBaseItem selectedItem, string baseFolder)
        {
            string file;
            var sourceDb = new Database(false, true);

            // Read the DWG into a side database
            sourceDb.ReadDwgFile(Path.Combine(baseFolder, selectedItem.SourceFile), FileOpenMode.OpenForReadAndAllShare, true, string.Empty);
            var dwgFile = new FileInfo(Path.Combine(baseFolder, selectedItem.SourceFile));
            
            // ReSharper disable once AssignNullToNotNullAttribute
            var iconPath = Path.Combine(dwgFile.DirectoryName, $"{dwgFile.Name} icons");
            using (var tr = sourceDb.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(sourceDb.BlockTableId, OpenMode.ForRead);

                file = ExtractThumbnails(iconPath, tr, bt[selectedItem.BlockName]);

                tr.Commit();
            }

            return file;
        }

        private static string ExtractThumbnails(string iconPath, Transaction tr, ObjectId id)
        {
            var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);

            if (!Directory.Exists(iconPath))
            {
                Directory.CreateDirectory(iconPath);
            }

            // Save the icon to our out directory

            var imgsrc = CMLContentSearchPreviews.GetBlockTRThumbnail(btr);
            var bmp = ImageSourceToGDI(imgsrc as System.Windows.Media.Imaging.BitmapSource);

            var fname = $"{iconPath}\\{btr.Name}.bmp";

            if (File.Exists(fname))
            {
                File.Delete(fname);
            }

            bmp.Save(fname);

            return fname;
        }

        // Helper function to generate an Image from a BitmapSource
        private static System.Drawing.Image ImageSourceToGDI(System.Windows.Media.Imaging.BitmapSource src)
        {
            var ms = new MemoryStream();
            var encoder = new System.Windows.Media.Imaging.BmpBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(src));
            encoder.Save(ms);
            ms.Flush();
            return System.Drawing.Image.FromStream(ms);
        }
#endif
    }
}