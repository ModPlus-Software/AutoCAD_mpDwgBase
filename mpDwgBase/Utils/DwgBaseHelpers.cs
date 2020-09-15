namespace mpDwgBase.Utils
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Windows.Media.Imaging;
    using System.Xml.Serialization;
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Models;
    using ModPlusAPI;
    using ModPlusAPI.Windows;

    public static class DwgBaseHelpers
    {
        private const string LangItem = "mpDwgBase";
        private static string _customBaseFolder;

        /// <summary>
        /// Сериализация списка элементов в xml
        /// </summary>
        /// <param name="listOfItems">Список элементов</param>
        /// <param name="fileName">Имя файла</param>
        public static void SerializerToXml(List<DwgBaseItem> listOfItems, string fileName)
        {
            var formatter = new XmlSerializer(typeof(List<DwgBaseItem>));
            using (FileStream fs = new FileStream(fileName, FileMode.Create))
            {
                formatter.Serialize(fs, listOfItems);
            }
        }

        public static bool DeserializeFromXml(string fileName, out List<DwgBaseItem> listOfItems)
        {
            try
            {
                var formatter = new XmlSerializer(typeof(List<DwgBaseItem>));
                using (var fs = new FileStream(fileName, FileMode.Open))
                {
                    listOfItems = formatter.Deserialize(fs) as List<DwgBaseItem>;
                    return true;
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                listOfItems = null;
                return false;
            }
        }

        /// <summary>
        /// Возвращает путь к папке пользовательской базы
        /// </summary>
        public static string GetCustomBaseFolder()
        {
            if (!string.IsNullOrEmpty(_customBaseFolder) && Directory.Exists(_customBaseFolder))
                return _customBaseFolder;

            _customBaseFolder = UserConfigFile.GetValue(LangItem, "CustomBaseFolder");
            if (Directory.Exists(_customBaseFolder))
                return _customBaseFolder;

            _customBaseFolder = Constants.DwgBaseDirectory;
            return _customBaseFolder;
        }

        /// <summary>
        /// Сохраняет путь к папке пользовательской базы
        /// </summary>
        /// <param name="folder">Путь к папке пользовательской базы</param>
        public static void SetCustomBaseFolder(string folder)
        {
            _customBaseFolder = folder;
            UserConfigFile.SetValue(LangItem, "CustomBaseFolder", folder, true);
        }

        public static string FindImageFile(DwgBaseItem selectedItem, string baseFolder)
        {
            var file = string.Empty;
            var dwgFile = new FileInfo(Path.Combine(baseFolder, selectedItem.SourceFile));

            if (Directory.Exists(dwgFile.DirectoryName))
            {
                file = Path.Combine(dwgFile.DirectoryName, $"{dwgFile.Name} icons", $"{selectedItem.BlockName}.bmp");
                if (File.Exists(file))
                {
                    return file;
                }

                file = string.Empty;
            }

            return file;
        }

        public static BitmapImage BitmapToImageSource(Bitmap bitmap)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Bmp);
                memory.Position = 0;
                BitmapImage bitmapimage = new BitmapImage();
                bitmapimage.BeginInit();
                bitmapimage.StreamSource = memory;
                bitmapimage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapimage.EndInit();

                return bitmapimage;
            }
        }

        public static bool Is2013DwgVersion(string filename)
        {
            using (var reader = new StreamReader(filename))
            {
                var readLine = reader.ReadLine();
                return readLine != null && readLine.Substring(0, 6).Equals("AC1027");
            }
        }

        public static string TrimStart(string target, string trimString)
        {
            var result = target;
            while (result.StartsWith(trimString))
            {
                result = result.Substring(trimString.Length);
            }

            return result;
        }

        public static bool HasProxyEntities(string file)
        {
            if (!File.Exists(file))
            {
                return false;
            }

            var db = new Database(false, true);
            db.ReadDwgFile(file, FileOpenMode.OpenForReadAndAllShare, true, string.Empty);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                foreach (var objectId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(objectId, OpenMode.ForRead);
                    foreach (var objId in btr)
                    {
                        var ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        if (ent != null && ent.IsAProxy)
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        public static string GetBlkNameForInsertDrawing(string blkName, Database db)
        {
            var returnedBlkName = blkName.Replace(".dwg", string.Empty);
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                while (bt != null && bt.Has(returnedBlkName))
                {
                    returnedBlkName = $"{blkName}_{DateTime.Now.Second}";
                }
            }

            return returnedBlkName;
        }

        /// <summary>
        /// Удаление блоков из БД
        /// </summary>
        public static void RemoveBlocksFromDB(IEnumerable<string> blkNames, Document doc, Database db)
        {
            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var bt = tr.GetObject(db.BlockTableId, OpenMode.ForWrite) as BlockTable;

                foreach (var blkName in blkNames)
                {
                    if (bt != null && bt.Has(blkName))
                    {
                        var btr = tr.GetObject(bt[blkName], OpenMode.ForWrite) as BlockTableRecord;
                        if (btr != null && (btr.GetBlockReferenceIds(false, false).Count == 0 && !btr.IsLayout))
                        {
                            btr.Erase(true);
                        }
                    }
                }

                tr.Commit();
            }
        }
    }
}