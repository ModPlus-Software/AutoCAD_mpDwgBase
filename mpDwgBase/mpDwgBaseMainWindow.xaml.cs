namespace mpDwgBase
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Net;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Forms;
    using System.Windows.Input;
    using System.Windows.Media;
    using System.Windows.Media.Imaging;
    using System.Windows.Threading;
    using System.Xml.Linq;
    using Autodesk.AutoCAD.DatabaseServices;
    using Models;
    using ModPlusAPI;
    using ModPlusAPI.Windows;
    using Utils;
    using Windows;
    using ModPlusStyle.Controls.Dialogs;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using ContextMenu = System.Windows.Controls.ContextMenu;
    using Cursors = System.Windows.Input.Cursors;
    using ListBox = System.Windows.Controls.ListBox;
    using MenuItem = System.Windows.Controls.MenuItem;
    using MessageBoxIcon = ModPlusAPI.Windows.MessageBoxIcon;
    using MouseEventArgs = System.Windows.Input.MouseEventArgs;
    using TextBox = System.Windows.Controls.TextBox;
    using TreeView = System.Windows.Controls.TreeView;
    using Visibility = System.Windows.Visibility;

    public partial class MpDwgBaseMainWindow
    {
        private const string LangItem = "mpDwgBase";

        // Путь к файлу, содержащему описание базы
        private string _dwgBaseFileName;

        // Список значений 
        private List<DwgBaseItem> _dwgBaseItems;

        // Create a Delegate that matches the Signature of the ProgressBar's SetValue method
        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, object value);

        private delegate void UpdateProgressTextDelegate(DependencyProperty dp, object value);

        public MpDwgBaseMainWindow()
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "h9");

            // Zooming and panning image
            MouseWheel += MainWindow_MouseWheel;
            BlkImagePreview.MouseDown += img_MouseDown;
            BlkImagePreview.MouseUp += img_MouseUp;
            BlkImagePreview.MouseMove += image_MouseMove;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SizeToContent = SizeToContent.Manual;

            // startup visibility
            RectangleIs3Dblock.Visibility = Visibility.Collapsed;
            BtUpload.Visibility = Visibility.Collapsed;

            /*
             * Загрузка файла-указателя вынесена в конфигуратор
             * и модуль автообновления, поэтому тут его проверять не будем
             */
            StartUpChecking();
            FillLayers();
            LoadFromSettings();
        }

        private void MpDwgBaseMainWindow_OnClosed(object sender, EventArgs e)
        {
            SaveToSettings();
        }

        private void StartUpChecking()
        {
            if (!File.Exists(Path.Combine(Constants.DwgBaseDirectory, "mpDwgBase.xml")))
            {
                _dwgBaseFileName = string.Empty;
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg1"));
            }
        }

        private void FillLayers()
        {
            CbLayers.Items.Add(ModPlusAPI.Language.GetItem(LangItem, "msg2"));
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            using (var tr = db.TransactionManager.StartTransaction())
            {
                // Open the Layer table for read
                var acLyrTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                if (acLyrTbl != null)
                {
                    foreach (var acObjId in acLyrTbl)
                    {
                        var acLyrTblRec = tr.GetObject(acObjId, OpenMode.ForRead) as LayerTableRecord;

                        if (acLyrTblRec != null && !acLyrTblRec.Name.Contains("|"))
                        {
                            CbLayers.Items.Add(acLyrTblRec.Name);
                        }
                    }
                }
            }

            CbLayers.SelectedIndex = 0;
        }

        private void LoadFromSettings()
        {
            ChkAlphabeticalSort.IsChecked = bool.TryParse(UserConfigFile.GetValue(LangItem, "AlphabeticalSort"), out var b) && b; // false
            ChkAlphabeticalSort.Checked += (sender, args) => FillByBaseType();
            ChkAlphabeticalSort.Unchecked += (sender, args) => FillByBaseType();
            ChkRotate.IsChecked = bool.TryParse(UserConfigFile.GetValue(LangItem, "Rotate"), out b) && b; // false
            ChkCloseAfterInsert.IsChecked = bool.TryParse(UserConfigFile.GetValue(LangItem, "CloseAfterInsert"), out b) && b; // false
            TbCustomBaseFolder.Text = DwgBaseHelpers.GetCustomBaseFolder();
            CbBaseType.SelectedIndex = int.TryParse(UserConfigFile.GetValue(LangItem, "BaseType"), out var i) ? i : 0;
        }

        private void SaveToSettings()
        {
            UserConfigFile.SetValue(LangItem, "Rotate", ChkRotate.IsChecked.ToString(), false);
            UserConfigFile.SetValue(LangItem, "BaseType", CbBaseType.SelectedIndex.ToString(), false);
            UserConfigFile.SetValue(LangItem, "CloseAfterInsert", ChkCloseAfterInsert.IsChecked.ToString(), false);
            UserConfigFile.SetValue(LangItem, "AlphabeticalSort", ChkAlphabeticalSort.IsChecked.ToString(), false);
            if (LbItems.SelectedItem != null)
            {
                UserConfigFile.SetValue(LangItem, "SelectedItemPath", ((DwgBaseItem)LbItems.SelectedItem).Path, false);
                UserConfigFile.SetValue(LangItem, "SelectedItemName", ((DwgBaseItem)LbItems.SelectedItem).Name, false);
            }

            UserConfigFile.SaveConfigFile();
        }

        // Выбор варианта базы
        private void CbBaseType_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FillByBaseType();
        }
        
        private void BtOpenCustomBaseFolder_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var folderBrowserDialog = new FolderBrowserDialog();
                if (folderBrowserDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                    return;
                DwgBaseHelpers.SetCustomBaseFolder(folderBrowserDialog.SelectedPath);
                TbCustomBaseFolder.Text = folderBrowserDialog.SelectedPath;
                FillByBaseType();
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private void BtRestoreCustomBaseFolder_OnClick(object sender, RoutedEventArgs e)
        {
            DwgBaseHelpers.SetCustomBaseFolder(Constants.DwgBaseDirectory);
            TbCustomBaseFolder.Text = Constants.DwgBaseDirectory;
            FillByBaseType();
        }

        // Заполнение окна в зависимости от версии базы
        private async void FillByBaseType()
        {
            // Сначала все очищаем
            ClearControls();

            // Проверка наличия файлов происходит при запуске функции. Поэтому здесь будем только парсить и биндить
            if (CbBaseType.SelectedIndex == 0)
            {
                // Отключаем кнопку загрузки на сервер
                BtUpload.Visibility = Visibility.Collapsed;

                // tree view context menu
                TvGroups.ContextMenu = null;

                // listbox context menu
                LbItems.ContextMenu = null;

                _dwgBaseFileName = Path.Combine(Constants.DwgBaseDirectory, "mpDwgBase.xml");

                // Отключаем контролы и пр.,связанное с локальной базой
                BtAddNewElement.Visibility = Visibility.Collapsed;
                BtUserBaseTools.Visibility = Visibility.Collapsed;
                TbVideoInstruction.Visibility = Visibility.Collapsed;
                TbCustomBaseFolder.Visibility = Visibility.Collapsed;
                BtOpenCustomBaseFolder.Visibility = Visibility.Collapsed;
                BtRestoreCustomBaseFolder.Visibility = Visibility.Collapsed;

                // Если файла нет, то ничего не заполнится
                if (File.Exists(_dwgBaseFileName))
                {
                    TvGroups.ItemsSource = DwgBaseHelpers.DeserializeFromXml(_dwgBaseFileName, out _dwgBaseItems) ? CreateTreeViewModel() : null;
                }
            }
            else if (CbBaseType.SelectedIndex == 1)
            {
                // файл должен лежать в той-же папке
                _dwgBaseFileName = Path.Combine(DwgBaseHelpers.GetCustomBaseFolder(), "UserDwgBase.xml");
                var hasFile = false;
                if (!File.Exists(_dwgBaseFileName))
                {
                    // Пользовательская база отсутствует. Создать?
                    if (await this.ShowMessageAsync(
                        ModPlusAPI.Language.GetItem(LangItem, "msg3"),
                        string.Empty,
                        MessageDialogStyle.AffirmativeAndNegative) == MessageDialogResult.Affirmative)
                    {
                        var newUserFile = new XElement("ArrayOfDwgBaseItem");
                        newUserFile.Save(_dwgBaseFileName);
                        hasFile = true;
                    }
                    else
                    {
                        CbBaseType.SelectedIndex = 0;
                    }
                }
                else
                {
                    hasFile = true;
                }

                if (hasFile)
                {
                    // Включаем контролы и пр.,связанное с локальной базой
                    BtAddNewElement.Visibility = Visibility.Visible;
                    BtUserBaseTools.Visibility = Visibility.Visible;
                    TbVideoInstruction.Visibility = Visibility.Visible;
                    TbCustomBaseFolder.Visibility = Visibility.Visible;
                    BtOpenCustomBaseFolder.Visibility = Visibility.Visible;
                    BtRestoreCustomBaseFolder.Visibility = Visibility.Visible;

                    // tree view context menu
                    TvGroups.ContextMenu = TvGroups.Resources["TvContextMenu"] as ContextMenu;

                    // listbox context menu
                    LbItems.ContextMenu = LbItems.Resources["LbContextMenu"] as ContextMenu;

                    var items = DwgBaseHelpers.DeserializeFromXml(_dwgBaseFileName, out _dwgBaseItems) ? CreateTreeViewModel() : null;

                    // Включаем кнопку загрузки на сервер если есть элементы
                    if (items != null && items.Any())
                    {
                        BtUpload.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        BtUpload.Visibility = Visibility.Collapsed;
                    }

                    TvGroups.ItemsSource = items;
                }
            }

            // После заполнения пробуем открыть элемент, который был открыт последним при прошлом закрытии
            var selectedItemPath = UserConfigFile.GetValue(LangItem, "SelectedItemPath");
            var selectedItemName = UserConfigFile.GetValue(LangItem, "SelectedItemName");
            SearchLastSelectedItem(selectedItemPath, selectedItemName);
        }

        private void ClearControls()
        {
            DgProperties.ItemsSource = null;
            BtInsert.IsEnabled = false;
            ChkInsertAsBlock.IsEnabled = false;
            BlkImagePreview.Source = null;
            TvGroups.ItemsSource = null;
            LbItems.ItemsSource = null;
            BlkImagePreview.Source = null;
            RectangleIs3Dblock.Visibility = Visibility.Collapsed;
        }

        private void ClearControlsWithoutTreeView()
        {
            DgProperties.ItemsSource = null;
            BtInsert.IsEnabled = false;
            ChkInsertAsBlock.IsEnabled = false;
            BlkImagePreview.Source = null;
            LbItems.ItemsSource = null;
            BlkImagePreview.Source = null;
            RectangleIs3Dblock.Visibility = Visibility.Collapsed;
        }

        private void LbItems_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!(sender is ListBox lb)) 
                return;

            if (lb.SelectedIndex != -1)
            {
                if (!(lb.SelectedItem is DwgBaseItem selectedItem))
                {
                    return;
                }

                var baseFolder = GetBaseFolder();

                // Get dwg properties
                DgProperties.ItemsSource = GetProperties(selectedItem);

                RectangleIs3Dblock.Visibility = selectedItem.Is3Dblock ? Visibility.Visible : Visibility.Collapsed;

                BtInsert.IsEnabled = true;
                if (selectedItem.IsBlock)
                {
                    // enabled
                    CbLayers.IsEnabled = true;
                    ChkInsertAsBlock.IsEnabled = false;

                    // Если блок, то нужно создать превью файл
                    // Создание превью только для автокадов выше 2013 версии
#if !A2013
                    if (File.Exists(Path.Combine(baseFolder, selectedItem.SourceFile)))
                    {
                        if (!File.Exists(DwgBaseHelpers.FindImageFile(selectedItem, baseFolder)))
                        {
                            ImageCreator.ImagePreviewFile(selectedItem, baseFolder);
                        }
                    }
#endif
                    // image
                    var imageFile = DwgBaseHelpers.FindImageFile(selectedItem, baseFolder);
                    if (string.IsNullOrEmpty(imageFile))
                    {
                        if (ModPlusAPI.Language.RusWebLanguages.Contains(ModPlusAPI.Language.CurrentLanguageName))
                        {
                            imageFile =
                                $"pack://application:,,,/mpDwgBase_{ModPlusConnector.Instance.AvailProductExternalVersion};component/Resources/NoImage.png";
                        }
                        else
                        {
                            imageFile =
                                $"pack://application:,,,/mpDwgBase_{ModPlusConnector.Instance.AvailProductExternalVersion};component/Resources/NoImageEn.png";
                        }
                    }

                    var bi = new BitmapImage();
                    bi.BeginInit();
                    bi.UriSource = new Uri(imageFile);
                    bi.EndInit();
                    BlkImagePreview.Source = bi;
                }
                else
                {
                    // enabled
                    CbLayers.IsEnabled = false;
                    ChkInsertAsBlock.IsEnabled = true;

                    // Если это чертеж, то нужно просто взять привью из его базы
                    ChkInsertAsBlock.IsEnabled = true;
                    var dwgFile = Path.Combine(baseFolder, selectedItem.SourceFile);
                    if (File.Exists(dwgFile))
                    {
                        var sourceDb = new Database(false, true);

                        // Read the DWG into a side database
                        sourceDb.ReadDwgFile(dwgFile, FileOpenMode.OpenForReadAndAllShare, true, string.Empty);
                        var bi = DwgBaseHelpers.BitmapToImageSource(sourceDb.ThumbnailBitmap);
                        BlkImagePreview.Source = bi;
                    }
                }
            }
            else
            {
                ClearControlsWithoutTreeView();
            }
        }
        
        private void TvGroups_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (sender is TreeView)
            {
                if (e.NewValue is TreeViewModelItem selectedTreeViewItem)
                {
                    var path = selectedTreeViewItem.GetAncestry();

                    // Сначала все очищаем
                    ClearControlsWithoutTreeView();
                    LbItems.ItemsSource = ChkAlphabeticalSort.IsChecked == true ?
                        _dwgBaseItems.OrderBy(i => i.Name).Where(item => item.Path.Equals(path)) :
                        _dwgBaseItems.Where(item => item.Path.Equals(path));
                }
            }
        }

        #region helpers
        private List<TreeViewModelItem> CreateTreeViewModel()
        {
            try
            {
                var ls = _dwgBaseItems.Select(item => item.Path).ToList();

                return BuildTree(ls, null);
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                return null;
            }
        }

        private List<TreeViewModelItem> BuildTree(IEnumerable<string> strings, TreeViewModelItem parent)
        {
            var list = (from s in strings
                        let split = s.Split('/')
                        group s by s.Split('/')[0]
                             into g
                        select new TreeViewModelItem()
                        {
                            Name = g.Key,
                            Parent = parent,
                            Children = g.ToList()
                        }).ToList();

            if (ChkAlphabeticalSort.IsChecked == true)
                list.Sort((i1, i2) => string.Compare(i1.Name, i2.Name, StringComparison.CurrentCulture));

            list.ForEach(x =>
            {
                x.Items = BuildTree(
                    from s in x.Children
                    where s.Length > x.Name.Length + 1
                    select s.Substring(x.Name.Length + 1), x);
            });

            return list;
        }
        #endregion

        #region insertion in dwg
        private void BtInsert_Click(object sender, RoutedEventArgs e)
        {
            InsertSelectedItem();
        }

        private void LbItems_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            InsertSelectedItem();
        }

        private async void InsertSelectedItem()
        {
            try
            {
                if (LbItems.SelectedIndex == -1)
                    return;

                if (!(LbItems.SelectedItem is DwgBaseItem selectedItem))
                    return;

                var baseFolder = GetBaseFolder();

                // Create a new instance of our ProgressBar Delegate that points
                //  to the ProgressBar's SetValue method.
                var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
                var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);
                
                // Проверяем есть ли файл
                var downloaded = false;
                if (!File.Exists(Path.Combine(baseFolder, selectedItem.SourceFile)))
                {
                    // Файл-источник отсутствует. Скачать?
                    if (await this.ShowMessageAsync(
                        ModPlusAPI.Language.GetItem(LangItem, "msg4"),
                        string.Empty,
                        MessageDialogStyle.AffirmativeAndNegative) == MessageDialogResult.Affirmative)
                    {
                        var localPath =
                            new FileInfo(Path.Combine(baseFolder, selectedItem.SourceFile)).DirectoryName;
                        if (!Directory.Exists(localPath))
                        {
                            if (localPath != null)
                            {
                                Directory.CreateDirectory(localPath);
                            }
                        }

                        downloaded = await DownloadSourceFile(
                            selectedItem.SourceFile,
                            $"https://modplus.org/Downloads/DwgBase/{selectedItem.SourceFile}",
                            $"{localPath}\\");
                    }
                }
                else
                {
                    downloaded = true;
                }

                if (downloaded)
                {
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();
                    if (selectedItem.IsBlock)
                    {
                        InsertBlock(selectedItem);
                    }
                    else
                    {
                        InsertDrawing(selectedItem);
                    }
                }

                // clear progress
                Dispatcher.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, string.Empty);
                Dispatcher.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, 0.0);
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private void InsertBlock(DwgBaseItem selectedItem)
        {
            var dm = AcApp.DocumentManager;
            var ed = dm.MdiActiveDocument.Editor;
            var destDb = dm.MdiActiveDocument.Database;
            var sourceDb = new Database(false, true);

            var baseFolder = GetBaseFolder();

            // Read the DWG into a side database
            sourceDb.ReadDwgFile(Path.Combine(baseFolder, selectedItem.SourceFile), FileOpenMode.OpenForReadAndAllShare, true, string.Empty);

            // Create a variable to store the list of block identifiers
            var blockIds = new ObjectIdCollection();
            using (dm.MdiActiveDocument.LockDocument())
            {
                using (var sourceT = sourceDb.TransactionManager.StartTransaction())
                {
                    // Open the block table
                    var bt = (BlockTable)sourceT.GetObject(sourceDb.BlockTableId, OpenMode.ForRead, false);

                    // Check each block in the block table
                    foreach (var btrId in bt)
                    {
                        var btr = (BlockTableRecord)sourceT.GetObject(btrId, OpenMode.ForRead, false);

                        // Only add named & non-layout blocks to the copy list
                        if (btr.Name.Equals(selectedItem.BlockName))
                        {
                            blockIds.Add(btrId);
                            break;
                        }

                        btr.Dispose();
                    }
                }

                // Copy blocks from source to destination database
                var mapping = new IdMapping();
                sourceDb.WblockCloneObjects(
                    blockIds, destDb.BlockTableId, mapping, DuplicateRecordCloning.Replace, false);
                sourceDb.Dispose();

                // Вставка
                Hide();
                using (var tr = destDb.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(destDb.BlockTableId, OpenMode.ForRead, false);

                    var layer = string.Empty;
                    if (CbLayers.SelectedIndex != 0)
                    {
                        layer = CbLayers.SelectedItem.ToString();
                    }

                    var attrs = new List<string>();
                    if (selectedItem.AttributesValues != null && selectedItem.AttributesValues.Any())
                    {
                        foreach (var attributeValue in selectedItem.AttributesValues)
                        {
                            attrs.Add(attributeValue.TextString);
                        }
                    }

                    var blkId = BlockInsertion.InsertBlockRef(
                        ChkRotate.IsChecked != null && ChkRotate.IsChecked.Value ? 1 : 0,
                        tr, destDb, ed, bt[selectedItem.BlockName],
                        attrs, selectedItem.IsAnnotative);

                    var blk = (BlockReference)tr.GetObject(blkId, OpenMode.ForWrite, true, false);
                    if (!string.IsNullOrEmpty(layer))
                    {
                        blk.Layer = layer;
                    }

                    // attributes for specification
                    if (selectedItem.HasAttributesForSpecification)
                    {
                        BlockInsertion.AddAttributesForSpecification(tr, blk, selectedItem);
                    }

                    tr.Commit();
                }

                if (ChkCloseAfterInsert.IsChecked != null && ChkCloseAfterInsert.IsChecked.Value)
                {
                    Close();
                }
                else
                {
                    Show();
                }
            }
        }

        private void InsertDrawing(DwgBaseItem selectedItem)
        {
            var dm = AcApp.DocumentManager;
            var ed = dm.MdiActiveDocument.Editor;
            var destDb = dm.MdiActiveDocument.Database;
            var sourceDb = new Database(false, true);
            Hide();

            var baseFolder = GetBaseFolder();

            using (dm.MdiActiveDocument.LockDocument())
            {
                var sourceFileInfo = new FileInfo(Path.Combine(baseFolder, selectedItem.SourceFile));

                // Read the DWG into a side database
                sourceDb.ReadDwgFile(Path.Combine(baseFolder, selectedItem.SourceFile), FileOpenMode.OpenForReadAndAllShare, true, string.Empty);
                var insertedBlkName = DwgBaseHelpers.GetBlkNameForInsertDrawing(sourceFileInfo.Name, destDb);
                var insertedDrawingId = destDb.Insert(sourceFileInfo.FullName, sourceDb, true);
                sourceDb.Dispose();
                using (var tr = destDb.TransactionManager.StartTransaction())
                {
                    // open block and set name
                    var btr = (BlockTableRecord)tr.GetObject(insertedDrawingId, OpenMode.ForWrite);
                    btr.Name = insertedBlkName;

                    // Вставка
                    var blkId = BlockInsertion.InsertBlockRef(
                        ChkRotate.IsChecked != null && ChkRotate.IsChecked.Value ? 1 : 0,
                        tr, destDb, ed, insertedDrawingId,
                        new List<string>(), selectedItem.IsAnnotative);

                    // if inserted
                    if (blkId != ObjectId.Null)
                    {
                        var blkRef = tr.GetObject(blkId, OpenMode.ForWrite, true, false) as BlockReference;
                        if (blkRef != null)
                        {
                            if (ChkInsertAsBlock.IsChecked != null && !ChkInsertAsBlock.IsChecked.Value)
                            {
                                blkRef.ExplodeToOwnerSpace();
                                DwgBaseHelpers.RemoveBlocksFromDB(new List<string> { insertedBlkName }, dm.MdiActiveDocument, destDb);
                            }
                        }
                    }

                    tr.Commit();
                }
            }// lock

            if (ChkCloseAfterInsert.IsChecked != null && ChkCloseAfterInsert.IsChecked.Value)
            {
                Close();
            }
            else
            {
                Show();
            }
        }

        /// <summary>
        /// Скачивание файла-источника. На сайте путь должен СОВПАДАТЬ!
        /// </summary>
        private async Task<bool> DownloadSourceFile(string sourceFile, string url, string localPath)
        {
            if (await ModPlusAPI.Web.Connection.HasAllConnectionAsync(1))
            {
                // Create a new instance of our ProgressBar Delegate that points
                //  to the ProgressBar's SetValue method.
                var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
                var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);

                // progress text
                Dispatcher.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                    $"{ModPlusAPI.Language.GetItem(LangItem, "msg5")}: {sourceFile}");

                long remoteSize;
                string fullLocalPath;

                try
                {
                    // Get the name of the remote file.
                    var remoteUri = new Uri(url);
                    var fileName = Path.GetFileName(remoteUri.LocalPath);

                    fullLocalPath = Path.GetFileName(localPath).Length == 0
                        ? Path.Combine(localPath, fileName)
                        : localPath;

                    // Have to get size of remote object through the webrequest as not available on remote files,
                    // although it does work on local files.
                    // WebRequest webRequest = WebRequest.Create(URL2.ToString());
                    var webRequest = WebRequest.Create(url);
                    webRequest.Proxy = ModPlusAPI.Web.Proxy.GetWebProxy();
                    using (var response = webRequest.GetResponse())
                    {
                        using (response.GetResponseStream())
                        {
                            remoteSize = response.ContentLength;
                        }
                    }
                }
                catch (Exception exception)
                {
                    ExceptionBox.Show(exception);
                    return false;
                }

                var bytesReadTotal = 0;
                try
                {
                    using (var webClient = new WebClient { Proxy = ModPlusAPI.Web.Proxy.GetWebProxy() })
                    {
                        using (var streamRemote = webClient.OpenRead(new Uri(url)))
                        {
                            using (Stream streamLocal = new FileStream(fullLocalPath, FileMode.Create, FileAccess.Write, FileShare.None))
                            {
                                var byteBuffer = new byte[1024 * 1024 * 2]; // 2 meg buffer although in testing only got to 10k max usage.
                                ProgressBar.Minimum = 0;
                                ProgressBar.Maximum = 100;
                                ProgressBar.Value = 0;
                                var perc = 0;
                                int bytesRead;
                                while ((bytesRead = await streamRemote.ReadAsync(byteBuffer, 0, byteBuffer.Length)) > 0)
                                {
                                    bytesReadTotal += bytesRead;
                                    await streamLocal.WriteAsync(byteBuffer, 0, bytesRead);
                                    var newPerc = (int)(bytesReadTotal / (double)remoteSize * 100);
                                    if (newPerc > perc)
                                    {
                                        perc = newPerc;

                                        // progress bar
                                        Dispatcher.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                                            $"{ModPlusAPI.Language.GetItem(LangItem, "msg5")}: {sourceFile} {perc}%");
                                        Dispatcher.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)perc);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception exception)
                {
                    ExceptionBox.Show(exception);
                    return false;
                }

                return true;
            }

            ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg6"));
            return false;
        }
        #endregion

        #region Zooming and panning image
        private Point _origin;
        private Point _start;

        private void img_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Middle && e.ButtonState == MouseButtonState.Pressed)
            {
                if (BlkImagePreview.IsMouseCaptured)
                {
                    return;
                }

                Cursor = Cursors.Hand;
                BlkImagePreview.CaptureMouse();

                _start = e.GetPosition(ImageBorder);
                _origin.X = BlkImagePreview.RenderTransform.Value.OffsetX;
                _origin.Y = BlkImagePreview.RenderTransform.Value.OffsetY;
            }
        }

        private void img_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Middle && e.ButtonState == MouseButtonState.Released)
            {
                BlkImagePreview.ReleaseMouseCapture();
                Cursor = Cursors.Arrow;
            }
        }

        private void image_MouseMove(object sender, MouseEventArgs e)
        {
            if (!BlkImagePreview.IsMouseCaptured)
            {
                return;
            }

            var p = e.MouseDevice.GetPosition(ImageBorder);

            var m = BlkImagePreview.RenderTransform.Value;
            m.OffsetX = _origin.X + (p.X - _start.X);
            m.OffsetY = _origin.Y + (p.Y - _start.Y);

            BlkImagePreview.RenderTransform = new MatrixTransform(m);
        }

        private void MainWindow_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            var p = e.MouseDevice.GetPosition(BlkImagePreview);

            var m = BlkImagePreview.RenderTransform.Value;
            if (e.Delta > 0)
            {
                m.ScaleAtPrepend(1.2, 1.2, p.X, p.Y);
            }
            else
            {
                m.ScaleAtPrepend(1 / 1.2, 1 / 1.2, p.X, p.Y);
            }

            BlkImagePreview.RenderTransform = new MatrixTransform(m);
        }

        private void BtImageSmall_OnClick(object sender, RoutedEventArgs e)
        {
            var m = BlkImagePreview.RenderTransform.Value;
            m.ScalePrepend(1 / 1.2, 1 / 1.2);
            BlkImagePreview.RenderTransform = new MatrixTransform(m);
        }

        private void BtImageBig_OnClick(object sender, RoutedEventArgs e)
        {
            var m = BlkImagePreview.RenderTransform.Value;
            m.ScalePrepend(1.2, 1.2);
            BlkImagePreview.RenderTransform = new MatrixTransform(m);
        }
        #endregion

        #region block properties
        
        internal class DwgProperties
        {
            public string Name { get; set; }

            public string Value { get; set; }
        }

        private IEnumerable<DwgProperties> GetProperties(DwgBaseItem item)
        {
            var properties = new List<DwgProperties>();

            // document
            properties.Add(new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p1"), Value = item.Document });

            // description
            properties.Add(new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p2"), Value = item.Description });

            // isBlock
            if (item.IsBlock)
            {
                properties.Add(new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p3"), Value = ModPlusAPI.Language.GetItem(LangItem, "block") });
                properties.Add(item.IsAnnotative
                    ? new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p4"), Value = ModPlusAPI.Language.GetItem(LangItem, "yes") }
                    : new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p4"), Value = ModPlusAPI.Language.GetItem(LangItem, "no") });
                properties.Add(item.IsDynamicBlock
                    ? new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p5"), Value = ModPlusAPI.Language.GetItem(LangItem, "yes") }
                    : new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p5"), Value = ModPlusAPI.Language.GetItem(LangItem, "no") });
                properties.Add(item.Is3Dblock
                    ? new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p6"), Value = ModPlusAPI.Language.GetItem(LangItem, "yes") }
                    : new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p6"), Value = ModPlusAPI.Language.GetItem(LangItem, "no") });
                properties.Add(item.HasAttributesForSpecification
                    ? new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p7"), Value = ModPlusAPI.Language.GetItem(LangItem, "yes") }
                    : new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p7"), Value = ModPlusAPI.Language.GetItem(LangItem, "no") });
            }
            else
            {
                properties.Add(new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p3"), Value = ModPlusAPI.Language.GetItem(LangItem, "drawing") });
            }

            // author
            properties.Add(new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p8"), Value = item.Author });

            // source
            properties.Add(new DwgProperties { Name = ModPlusAPI.Language.GetItem(LangItem, "p9"), Value = item.Source });
            return properties;
        }
        #endregion

        #region Работа с локальной базой

        private void BtAddNewElement_OnClick(object sender, RoutedEventArgs e)
        {
            Hide();
            try
            {
                var win = new SelectAddingVariant();

                if (win.ShowDialog() == true)
                {
                    // fill cb with path
                    var pathes = new List<string>();
                    foreach (var dwgBaseItem in _dwgBaseItems)
                    {
                        if (!pathes.Contains(dwgBaseItem.Path))
                        {
                            pathes.Add(dwgBaseItem.Path);
                        }
                    }

                    var baseFolder = GetBaseFolder();

                    var reFill = false;
                    if (win.Variant.Equals("Block"))
                    {
                        // Т.к. новый элемент, то меняем у окна заголовок и создаем пустой элемент DwgBaseItem
                        var newBlock = new BlockWindow(
                            Path.Combine(Constants.DwgBaseDirectory, "mpDwgBase.xml"),
                            Path.Combine(DwgBaseHelpers.GetCustomBaseFolder(), "UserDwgBase.xml"),
                            baseFolder, false)
                        {
                            Title = ModPlusAPI.Language.GetItem(LangItem, "h22"),
                            Item = new DwgBaseItem(),
                            Owner = this,
                            CbPath = { ItemsSource = pathes.Where(x => x.Contains("Блоки/")) }
                        };

                        if (newBlock.ShowDialog() == true)
                        {
                            _dwgBaseItems.Add(newBlock.Item);
                            DwgBaseHelpers.SerializerToXml(_dwgBaseItems, _dwgBaseFileName);
                            reFill = true;
                        }
                    }

                    if (win.Variant.Equals("Drawing"))
                    {
                        var newDrawing = new DrawingWindow(
                            Path.Combine(Constants.DwgBaseDirectory, "mpDwgBase.xml"),
                            Path.Combine(DwgBaseHelpers.GetCustomBaseFolder(), "UserDwgBase.xml"),
                            baseFolder, false)
                        {
                            Title = ModPlusAPI.Language.GetItem(LangItem, "h23"),
                            Item = new DwgBaseItem(),
                            Owner = this,
                            CbPath = { ItemsSource = pathes.Where(x => x.Contains("Чертежи/")) }
                        };
                        if (newDrawing.ShowDialog() == true)
                        {
                            _dwgBaseItems.Add(newDrawing.Item);
                            DwgBaseHelpers.SerializerToXml(_dwgBaseItems, _dwgBaseFileName);
                            reFill = true;
                        }
                    }

                    if (reFill)
                    {
                        // clear before refill
                        ClearControls();
                        TvGroups.ItemsSource = DwgBaseHelpers.DeserializeFromXml(_dwgBaseFileName, out _dwgBaseItems) ? CreateTreeViewModel() : null;

                        // Нужно показать кнопку загрузки на сервер, т.к. ее может не быть
                        BtUpload.Visibility = Visibility.Visible;
                    }
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
            finally
            {
                AcApp.ShowModalWindow(AcApp.MainWindow.Handle, this, false);
            }
        }

        // get userbase tools
        private void BtUserBaseTools_OnClick(object sender, RoutedEventArgs e)
        {
            var selectedIndex = LbItems.SelectedIndex;
            var selectedTvItem = TvGroups.SelectedItem as TreeViewModelItem;

            var win = new UserBaseTools(
                Path.Combine(Constants.DwgBaseDirectory, "mpDwgBase.xml"),
                Path.Combine(DwgBaseHelpers.GetCustomBaseFolder(), "UserDwgBase.xml"),
                GetBaseFolder(), _dwgBaseItems);
            win.ShowDialog();
            if (win.UserBaseChanged)
            {
                TvGroups.ItemsSource = null;

                // rebind
                var items = DwgBaseHelpers.DeserializeFromXml(_dwgBaseFileName, out _dwgBaseItems) ? CreateTreeViewModel() : null;
                TvGroups.ItemsSource = items;

                // reselect
                var listOfItemsToOpen = new List<TreeViewModelItem>();
                GetParentsOfTreeViewModelItem(selectedTvItem, listOfItemsToOpen);
                if (listOfItemsToOpen.Any())
                {
                    listOfItemsToOpen.Reverse();
                    var splittedPath = new List<string>();
                    foreach (var treeViewModelItem in listOfItemsToOpen)
                    {
                        splittedPath.Add(treeViewModelItem.Name);
                    }

                    if (splittedPath.Any())
                    {
                        foreach (TreeViewModelItem tvGroupsItem in TvGroups.Items)
                        {
                            SearchInTreeView(tvGroupsItem, null, splittedPath[0], splittedPath);
                        }
                    }
                }

                if (selectedIndex > 0)
                {
                    if (selectedIndex < LbItems.Items.Count - 1)
                    {
                        LbItems.SelectedIndex = selectedIndex;
                    }
                    else
                    {
                        LbItems.SelectedIndex = LbItems.Items.Count - 1;
                    }
                }
                else
                {
                    LbItems.SelectedIndex = 0;
                }
            }
        }

        // on preview groups context menu
        private void TvGroups_OnContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            // tree view context menu
            if (!(TvGroups.Resources["TvContextMenu"] is ContextMenu cm))
            {
                return;
            }

            var miRenameGroup = LogicalTreeHelper.FindLogicalNode(cm, "MiRenameGroup") as MenuItem;

            if (!(TvGroups.SelectedItem is TreeViewModelItem selectedItem))
            {
                foreach (var cmItem in cm.Items)
                {
                    if (cmItem is MenuItem mi)
                    {
                        mi.IsEnabled = false;
                    }
                }
            }
            else
            {
                foreach (var cmItem in cm.Items)
                {
                    if (cmItem is MenuItem mi)
                    {
                        mi.IsEnabled = true;
                    }
                }

                if (selectedItem.Name.Equals("Блоки") || selectedItem.Name.Equals("Чертежи"))
                {
                    if (miRenameGroup != null)
                    {
                        miRenameGroup.IsEnabled = false;
                    }
                }
                else if (miRenameGroup != null)
                {
                    miRenameGroup.IsEnabled = true;
                }
            }
        }

        // on preview items context menu
        private void LbItems_OnContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            if (!(LbItems.Resources["LbContextMenu"] is ContextMenu cm))
            {
                return;
            }

            if (!(LbItems.SelectedItem is DwgBaseItem))
            {
                foreach (var cmItem in cm.Items)
                {
                    if (cmItem is MenuItem mi)
                    {
                        mi.IsEnabled = false;
                    }
                }
            }
            else
            {
                foreach (var cmItem in cm.Items)
                {
                    if (cmItem is MenuItem mi)
                    {
                        mi.IsEnabled = true;
                    }
                }
            }
        }

        // delete group
        private void TreeViewContextMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            if (!(TvGroups.SelectedItem is TreeViewModelItem selectedItem))
            {
                return;
            }

            if (ModPlusAPI.Windows.MessageBox.ShowYesNo(
                $"{ModPlusAPI.Language.GetItem(LangItem, "msg7")}: {selectedItem.Name} {ModPlusAPI.Language.GetItem(LangItem, "msg8")}!{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg9")}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg9")}: {selectedItem.Name}?", MessageBoxIcon.Question))
            {
                // Нужно построить полный путь. Для этого нужно построить "начало" через родителей (рекурсия)
                // которое будет единственным и построить "окончания" через потомков
                var parentOfSelectedItem = selectedItem.Parent;
                var startPath = string.Empty;
                var listOfPathToDelete = new List<string>();
                GetStartPathForSelectedItem(selectedItem, ref startPath);
                if (selectedItem.Items.Any())
                {
                    foreach (var selectedItemChild in selectedItem.Items)
                    {
                        GetSubItemsPathForSelectedItem(selectedItemChild, startPath, ref listOfPathToDelete);
                    }
                }
                else
                {
                    listOfPathToDelete.Add(startPath.TrimEnd('/'));
                }

                // deleting
                if (listOfPathToDelete.Any())
                {
                    DeleteGroup(listOfPathToDelete);
                }

                // resave
                DwgBaseHelpers.SerializerToXml(_dwgBaseItems, _dwgBaseFileName);

                // Сначала все очищаем
                ClearControls();

                // rebind
                var items = DwgBaseHelpers.DeserializeFromXml(_dwgBaseFileName, out _dwgBaseItems) ? CreateTreeViewModel() : null;

                // Включаем кнопку загрузки на сервер если есть элементы
                if (items != null && items.Any())
                {
                    BtUpload.Visibility = Visibility.Visible;
                }
                else
                {
                    BtUpload.Visibility = Visibility.Collapsed;
                }

                TvGroups.ItemsSource = items;

                // reselect
                if (parentOfSelectedItem != null)
                {
                    var listOfItemsToOpen = new List<TreeViewModelItem>();
                    GetParentsOfTreeViewModelItem(parentOfSelectedItem, listOfItemsToOpen);
                    if (listOfItemsToOpen.Any())
                    {
                        listOfItemsToOpen.Reverse();
                        var splittedPath = new List<string>();
                        foreach (var treeViewModelItem in listOfItemsToOpen)
                        {
                            splittedPath.Add(treeViewModelItem.Name);
                        }

                        if (splittedPath.Any())
                        {
                            foreach (TreeViewModelItem tvGroupsItem in TvGroups.Items)
                            {
                                SearchInTreeView(tvGroupsItem, null, splittedPath[0], splittedPath);
                            }
                        }
                    }
                }
            }
        }

        // rename group
        private void TreeViewContextMenuRenameGroup_OnClick(object sender, RoutedEventArgs e)
        {
            if (!(TvGroups.SelectedItem is TreeViewModelItem selectedItem))
            {
                return;
            }

            // get window
            var win = new RenameGroup();
            win.TbCurrentName.Text = win.TbNewName.Text = selectedItem.Name;
            if (win.ShowDialog() == true)
            {
                // Нужно построить полный путь. Для этого нужно построить "начало" через родителей (рекурсия)
                // которое будет единственным и построить "окончания" через потомков
                var parentOfSelectedItem = selectedItem.Parent;
                var oldStartPath = string.Empty;
                var listOfPathToRename = new List<string>();
                GetStartPathForSelectedItem(selectedItem, ref oldStartPath);

                // После того как получили "начало" пути, нам нужно создать заменяюмую часть этого пути
                var newStartPath = CreateNewStartPath(oldStartPath.TrimEnd('/'), win.TbNewName.Text);

                // Добавим в список "продолжения" пути:
                if (selectedItem.Items.Any())
                {
                    foreach (var selectedItemChild in selectedItem.Items)
                    {
                        GetSubItemsPathForSelectedItem(selectedItemChild, oldStartPath, ref listOfPathToRename);
                    }
                }
                else
                {
                    listOfPathToRename.Add(oldStartPath.TrimEnd('/'));
                }

                // replace
                foreach (var dwgBaseItem in _dwgBaseItems)
                {
                    if (listOfPathToRename.Contains(dwgBaseItem.Path))
                    {
                        var ns = dwgBaseItem.Path.Replace(oldStartPath.TrimEnd('/'), newStartPath.TrimEnd('/'));
                        dwgBaseItem.Path = ns;
                    }
                }

                // resave
                DwgBaseHelpers.SerializerToXml(_dwgBaseItems, _dwgBaseFileName);

                // Сначала все очищаем
                ClearControls();

                // rebind
                var items = DwgBaseHelpers.DeserializeFromXml(_dwgBaseFileName, out _dwgBaseItems) ? CreateTreeViewModel() : null;
                TvGroups.ItemsSource = items;

                // reselect
                if (parentOfSelectedItem != null)
                {
                    var listOfItemsToOpen = new List<TreeViewModelItem>();
                    TreeViewModelItem renamedGroup = null;
                    SearchRenamedGroup(null, win.TbNewName.Text, ref renamedGroup);
                    GetParentsOfTreeViewModelItem(renamedGroup ?? parentOfSelectedItem, listOfItemsToOpen);
                    if (listOfItemsToOpen.Any())
                    {
                        listOfItemsToOpen.Reverse();
                        var splittedPath = new List<string>();
                        foreach (var treeViewModelItem in listOfItemsToOpen)
                        {
                            splittedPath.Add(treeViewModelItem.Name);
                        }

                        if (splittedPath.Any())
                        {
                            foreach (TreeViewModelItem tvGroupsItem in TvGroups.Items)
                            {
                                SearchInTreeView(tvGroupsItem, null, splittedPath[0], splittedPath);
                            }
                        }
                    }
                }
            }
        }

        // Т.к. в "пути" заменяемая часть может повториться, то нам нужно заменить только последний элемент
        private string CreateNewStartPath(string oldStartPath, string newName)
        {
            var newStartPath = string.Empty;
            var lst = oldStartPath.Split('/').ToList();
            lst[lst.Count - 1] = newName;
            foreach (var l in lst)
            {
                newStartPath += $"{l}/";
            }

            return newStartPath;
        }

        /// <summary>
        /// Поиск в дереве переименованного значения
        /// </summary>
        /// <param name="parentItem">Родительский элемент дерева</param>
        /// <param name="newName">Имя элемента для поиска</param>
        /// <param name="treeViewModelItem">Ссылочный элемент для записи найденного элемента</param>
        private void SearchRenamedGroup(TreeViewModelItem parentItem, string newName, ref TreeViewModelItem treeViewModelItem)
        {
            if (parentItem == null)
            {
                foreach (TreeViewModelItem tvGroupsItem in TvGroups.Items)
                {
                    if (tvGroupsItem.Name.Equals(newName))
                    {
                        treeViewModelItem = tvGroupsItem;
                        break;
                    }

                    if (tvGroupsItem.Items.Any())
                    {
                        SearchRenamedGroup(tvGroupsItem, newName, ref treeViewModelItem);
                    }
                }
            }
            else
            {
                foreach (var tvGroupsItem in parentItem.Items)
                {
                    if (tvGroupsItem.Name.Equals(newName))
                    {
                        treeViewModelItem = tvGroupsItem;
                        break;
                    }

                    if (tvGroupsItem.Items.Any())
                    {
                        SearchRenamedGroup(tvGroupsItem, newName, ref treeViewModelItem);
                    }
                }
            }
        }

        // delete item
        private void LbItemsContextMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            var selectedItem = LbItems.SelectedItem as DwgBaseItem;
            var selectedIndex = LbItems.SelectedIndex;
            var selectedTvItem = TvGroups.SelectedItem as TreeViewModelItem;
            if (selectedItem == null)
            {
                return;
            }

            if (selectedTvItem == null)
            {
                return;
            }

            if (ModPlusAPI.Windows.MessageBox.ShowYesNo(
                $"{ModPlusAPI.Language.GetItem(LangItem, "msg11")}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg9")}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg12")}: {selectedItem.Name}?", MessageBoxIcon.Question))
            {
                // deleting
                _dwgBaseItems.Remove(selectedItem);

                // resave
                DwgBaseHelpers.SerializerToXml(_dwgBaseItems, _dwgBaseFileName);

                // Сначала все очищаем
                ClearControls();

                // rebind
                TvGroups.ItemsSource = DwgBaseHelpers.DeserializeFromXml(_dwgBaseFileName, out _dwgBaseItems) ? CreateTreeViewModel() : null;

                // reselect
                var listOfItemsToOpen = new List<TreeViewModelItem>();
                GetParentsOfTreeViewModelItem(selectedTvItem, listOfItemsToOpen);
                if (listOfItemsToOpen.Any())
                {
                    listOfItemsToOpen.Reverse();
                    var splittedPath = new List<string>();
                    foreach (var treeViewModelItem in listOfItemsToOpen)
                    {
                        splittedPath.Add(treeViewModelItem.Name);
                    }

                    if (splittedPath.Any())
                    {
                        foreach (TreeViewModelItem tvGroupsItem in TvGroups.Items)
                        {
                            SearchInTreeView(tvGroupsItem, null, splittedPath[0], splittedPath);
                        }
                    }
                }

                if (selectedIndex > 0)
                {
                    if (selectedIndex < LbItems.Items.Count - 1)
                    {
                        LbItems.SelectedIndex = selectedIndex;
                    }
                    else
                    {
                        LbItems.SelectedIndex = LbItems.Items.Count - 1;
                    }
                }
                else
                {
                    LbItems.SelectedIndex = 0;
                }
            }
        }

        // edit item
        private void LbItemsEditItemMenuItem_OnClick(object sender, RoutedEventArgs e)
        {
            var selectedItem = LbItems.SelectedItem as DwgBaseItem;
            var selectedIndex = LbItems.SelectedIndex;
            var selectedTvItem = TvGroups.SelectedItem as TreeViewModelItem;
            if (selectedItem == null)
            {
                return;
            }

            if (selectedTvItem == null)
            {
                return;
            }

            var needRefill = false;

            // fill cb with path
            var pathes = new List<string>();
            foreach (var dwgBaseItem in _dwgBaseItems)
            {
                if (!pathes.Contains(dwgBaseItem.Path))
                {
                    pathes.Add(dwgBaseItem.Path);
                }
            }

            // if block
            if (selectedItem.IsBlock)
            {
                var editBlock = new BlockWindow(
                            Path.Combine(Constants.DwgBaseDirectory, "mpDwgBase.xml"),
                            Path.Combine(DwgBaseHelpers.GetCustomBaseFolder(), "UserDwgBase.xml"),
                            GetBaseFolder(), true)
                {
                    Title = ModPlusAPI.Language.GetItem(LangItem, "h24"),
                    Item = selectedItem,
                    Owner = this,
                    CbPath = { ItemsSource = pathes.Where(x => x.StartsWith("Блоки/")) }
                };
                editBlock.BtSelectBlock.IsEnabled = false;
                editBlock.TbBlockName.Text = selectedItem.BlockName;
                editBlock.TbIsAnnot.Opacity = selectedItem.IsAnnotative ? 1.0 : 0.5;
                editBlock.TbIsDynamic.Opacity = selectedItem.IsDynamicBlock ? 1.0 : 0.5;

                // Биндим 
                editBlock.GridBlockDetails.DataContext = selectedItem;

                // включаем видимость подробностей
                editBlock.GridBlockDetails.Visibility = Visibility.Visible;

                editBlock.BtAccept.IsEnabled = true;

                editBlock.ChkIs3Dblock.IsChecked = selectedItem.Is3Dblock;

                // other visibility
                editBlock.TbSourceFile.Visibility = editBlock.BtCreateDwgFile.Visibility =
                editBlock.BtSelectDwgFile.Visibility = editBlock.ChkIsCurrentDwgFile.Visibility =
                editBlock.RectangleSourceFile.Visibility = editBlock.BtGetAttrValuesFromBlock.Visibility =
                editBlock.BtLoadLastEnteredData.Visibility = Visibility.Collapsed;
                editBlock.BtSelectBlock.IsEnabled = false;

                if (editBlock.ShowDialog() == true)
                {
                    for (var index = 0; index < _dwgBaseItems.Count; index++)
                    {
                        var dwgBaseItem = _dwgBaseItems[index];
                        if (dwgBaseItem.Equals(selectedItem))
                        {
                            _dwgBaseItems[index] = editBlock.Item;
                            break;
                        }
                    }

                    needRefill = true;
                }
            }

            // is drawing
            else 
            {
                var editDrawing = new DrawingWindow(
                            Path.Combine(Constants.DwgBaseDirectory, "mpDwgBase.xml"),
                            Path.Combine(DwgBaseHelpers.GetCustomBaseFolder(), "UserDwgBase.xml"),
                            GetBaseFolder(), true)
                {
                    Title = ModPlusAPI.Language.GetItem(LangItem, "h25"),
                    Item = selectedItem,
                    Owner = this,
                    CbPath = { ItemsSource = pathes.Where(x => x.StartsWith("Чертежи/")) }
                };

                // Биндим 
                editDrawing.GridDrawingDetails.DataContext = selectedItem;

                editDrawing.BtAccept.IsEnabled = true;

                // other visibility
                editDrawing.TbSourceFile.Visibility = editDrawing.BtCopyDwgFile.Visibility =
                editDrawing.BtSelectDwgFile.Visibility = editDrawing.ChkIsCurrentDwgFile.Visibility =
                editDrawing.RectangleSourceFile.Visibility =
                editDrawing.BtLoadLastInteredData.Visibility = Visibility.Collapsed;
                if (editDrawing.ShowDialog() == true)
                {
                    for (var index = 0; index < _dwgBaseItems.Count; index++)
                    {
                        var dwgBaseItem = _dwgBaseItems[index];
                        if (dwgBaseItem.Equals(selectedItem))
                        {
                            _dwgBaseItems[index] = editDrawing.Item;
                            break;
                        }
                    }

                    needRefill = true;
                }
            }

            if (needRefill)
            {
                // resave
                DwgBaseHelpers.SerializerToXml(_dwgBaseItems, _dwgBaseFileName);

                // Сначала все очищаем
                ClearControls();

                // rebind
                TvGroups.ItemsSource = DwgBaseHelpers.DeserializeFromXml(_dwgBaseFileName, out _dwgBaseItems) ? CreateTreeViewModel() : null;

                // reselect
                var listOfItemsToOpen = new List<TreeViewModelItem>();
                GetParentsOfTreeViewModelItem(selectedTvItem, listOfItemsToOpen);
                if (listOfItemsToOpen.Any())
                {
                    listOfItemsToOpen.Reverse();
                    var splittedPath = new List<string>();
                    foreach (var treeViewModelItem in listOfItemsToOpen)
                    {
                        splittedPath.Add(treeViewModelItem.Name);
                    }

                    if (splittedPath.Any())
                    {
                        foreach (TreeViewModelItem tvGroupsItem in TvGroups.Items)
                        {
                            SearchInTreeView(tvGroupsItem, null, splittedPath[0], splittedPath);
                        }
                    }
                }

                if (selectedIndex > 0)
                {
                    if (selectedIndex < LbItems.Items.Count - 1)
                    {
                        LbItems.SelectedIndex = selectedIndex;
                    }
                    else
                    {
                        LbItems.SelectedIndex = LbItems.Items.Count - 1;
                    }
                }
                else
                {
                    LbItems.SelectedIndex = 0;
                }
            }
        }

        /// <summary>
        /// Построение строки "Путь" от текущего элемента дерева до первого элемента дерева
        /// </summary>
        private void GetStartPathForSelectedItem(TreeViewModelItem treeViewModelItem, ref string str)
        {
            str = $"{treeViewModelItem.Name}/{str}";
            if (treeViewModelItem.Parent != null)
            {
                GetStartPathForSelectedItem(treeViewModelItem.Parent, ref str);
            }
        }

        private void GetSubItemsPathForSelectedItem(
            TreeViewModelItem treeViewModelItem, string path, ref List<string> listOfPathToDelete)
        {
            path += $"{treeViewModelItem.Name}/";
            if (treeViewModelItem.Items.Any())
            {
                foreach (var viewModelItem in treeViewModelItem.Items)
                {
                    GetSubItemsPathForSelectedItem(viewModelItem, path, ref listOfPathToDelete);
                }
            }
            else
            {
                if (!listOfPathToDelete.Contains(path.TrimEnd('/')))
                {
                    listOfPathToDelete.Add(path.TrimEnd('/'));
                }
            }
        }

        /// <summary>
        /// Удаление элементов по их пути. Удаление происходит из коллекции с сохранением и последующим ребиндингом
        /// </summary>
        private void DeleteGroup(List<string> listOfPathToDelete)
        {
            // Т.к. нельзя менять коллекцию при ее переборе, воспользуемся вспомогательной коллекцией
            var tmpList = new List<DwgBaseItem>();
            foreach (var dwgBaseItem in _dwgBaseItems)
            {
                if (listOfPathToDelete.Contains(dwgBaseItem.Path))
                {
                    tmpList.Add(dwgBaseItem);
                }
            }

            if (tmpList.Any())
            {
                foreach (var dwgBaseItem in tmpList)
                {
                    _dwgBaseItems.Remove(dwgBaseItem);
                }
            }
        }

        /// <summary>
        /// Сборка всех "родителей" в список
        /// </summary>
        private void GetParentsOfTreeViewModelItem(TreeViewModelItem treeViewModelItem, List<TreeViewModelItem> list)
        {
            if (!list.Contains(treeViewModelItem))
            {
                list.Add(treeViewModelItem);
            }

            if (treeViewModelItem.Parent != null)
            {
                GetParentsOfTreeViewModelItem(treeViewModelItem.Parent, list);
            }
        }

        // uploading
        private void BtUpload_OnClick(object sender, RoutedEventArgs e)
        {
            if (_dwgBaseItems.Any())
            {
                var win = new BaseUploading(GetBaseFolder(), _dwgBaseItems);
                win.ShowDialog();
            }
        }
        #endregion

        #region Serach

        // Введем временную переменную для определения, что при повторном открытии
        // поля поиска не изменилась база
        private string _tmpDwgBaseFile = string.Empty;

        private void BtSearch_OnClick(object sender, RoutedEventArgs e)
        {
            if (!_tmpDwgBaseFile.Equals(_dwgBaseFileName))
            {
                TbSearchTxt.Text = string.Empty;
                LbSearchResults.ItemsSource = null;
            }

            _tmpDwgBaseFile = _dwgBaseFileName;
            FlyoutSearch.IsOpen = !FlyoutSearch.IsOpen;
        }

        private void TbSearchTxt_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            // Пусть работает от трех символов и выше
            if (sender is TextBox textbox)
            {
                if (textbox.Text.Length >= 3)
                {
                    SearchItemsInBases(textbox.Text);
                }
                else
                {
                    LbSearchResults.ItemsSource = null;
                }
            }
        }

        private void SearchItemsInBases(string searchValue)
        {
            LbSearchResults.ItemsSource = null;
            var serachResult = new List<DwgBaseItem>();
            foreach (var dwgBaseItem in _dwgBaseItems)
            {
                if (dwgBaseItem.Name.ToLower().Contains(searchValue.ToLower()))
                {
                    if (!serachResult.Contains(dwgBaseItem))
                    {
                        serachResult.Add(dwgBaseItem);
                    }
                }

                if (dwgBaseItem.Description.ToLower().Contains(searchValue.ToLower()))
                {
                    if (!serachResult.Contains(dwgBaseItem))
                    {
                        serachResult.Add(dwgBaseItem);
                    }
                }

                if (dwgBaseItem.Path.ToLower().Contains(searchValue.ToLower()))
                {
                    if (!serachResult.Contains(dwgBaseItem))
                    {
                        serachResult.Add(dwgBaseItem);
                    }
                }

                if (dwgBaseItem.Author.ToLower().Contains(searchValue.ToLower()))
                {
                    if (!serachResult.Contains(dwgBaseItem))
                    {
                        serachResult.Add(dwgBaseItem);
                    }
                }

                if (dwgBaseItem.Source.ToLower().Contains(searchValue.ToLower()))
                {
                    if (!serachResult.Contains(dwgBaseItem))
                    {
                        serachResult.Add(dwgBaseItem);
                    }
                }
            }

            LbSearchResults.ItemsSource = serachResult;
        }

        private void LbSearchResults_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var lb = sender as ListBox;
            if (!(lb?.SelectedItem is DwgBaseItem selectedItem))
            {
                return;
            }

            var splittedPath = selectedItem.Path.Split('/').ToList();
            if (splittedPath.Any())
            {
                foreach (TreeViewModelItem tvGroupsItem in TvGroups.Items)
                {
                    SearchInTreeView(tvGroupsItem, null, splittedPath[0], splittedPath);
                }
            }

            // select in listbox
            if (LbItems.Items.Count > 0)
            {
                foreach (var lbItemsItem in LbItems.Items)
                {
                    if (lbItemsItem is DwgBaseItem baseItem && baseItem.Name.Equals(selectedItem.Name))
                    {
                        LbItems.SelectedIndex = LbItems.Items.IndexOf(lbItemsItem);
                        break;
                    }
                }
            }
        }

        private void SearchLastSelectedItem(string selectedPath, string selectedName)
        {
            if (string.IsNullOrEmpty(selectedPath))
            {
                return;
            }

            // try open group
            var splittedPath = selectedPath.Split('/').ToList();
            if (splittedPath.Any())
            {
                foreach (TreeViewModelItem tvGroupsItem in TvGroups.Items)
                {
                    SearchInTreeView(tvGroupsItem, null, splittedPath[0], splittedPath);
                }
            }

            // try select item
            if (!string.IsNullOrEmpty(selectedName))
            {
                if (LbItems.Items.Count > 0)
                {
                    foreach (var lbItemsItem in LbItems.Items)
                    {
                        if (lbItemsItem is DwgBaseItem baseItem && baseItem.Name.Equals(selectedName))
                        {
                            LbItems.SelectedIndex = LbItems.Items.IndexOf(lbItemsItem);
                            break;
                        }
                    }
                }
            }
        }

        private void SearchInTreeView(TreeViewModelItem treeViewModelItem, TreeViewItem parentItem, string toSearch, List<string> splittedPath)
        {
            if (treeViewModelItem.Name.Equals(toSearch))
            {
                if (treeViewModelItem.Items.Any() & splittedPath.Count >= 2)
                {
                    TreeViewItem treeViewItem;
                    if (parentItem == null)
                    {
                        treeViewItem = TvGroups.ItemContainerGenerator.ContainerFromItem(treeViewModelItem) as TreeViewItem;
                    }
                    else
                    {
                        treeViewItem = parentItem.ItemContainerGenerator.ContainerFromItem(treeViewModelItem) as TreeViewItem;
                    }

                    if (treeViewItem != null)
                    {
                        treeViewItem.IsSelected = true;
                        treeViewItem.IsExpanded = true;
                        TvGroups.UpdateLayout();
                    }

                    splittedPath.RemoveAt(0);
                    if (splittedPath.Any())
                    {
                        foreach (var viewModelItem in treeViewModelItem.Items)
                        {
                            SearchInTreeView(viewModelItem, treeViewItem, splittedPath[0], splittedPath);
                        }
                    }
                }
                else
                {
                    TreeViewItem treeViewItem;
                    if (parentItem == null)
                    {
                        treeViewItem = TvGroups.ItemContainerGenerator.ContainerFromItem(treeViewModelItem) as TreeViewItem;
                    }
                    else
                    {
                        treeViewItem = parentItem.ItemContainerGenerator.ContainerFromItem(treeViewModelItem) as TreeViewItem;
                    }

                    if (treeViewItem != null)
                    {
                        treeViewItem.IsSelected = true;
                        if (treeViewModelItem.Items.Any())
                        {
                            treeViewItem.IsExpanded = true;
                        }
                    }
                }
            }
        }
        #endregion

        private void HyperlinkVideoInstruction_OnClick(object sender, RoutedEventArgs e)
        {
            Process.Start("https://youtu.be/5LgxVcM9RsM");
        }

        private string GetBaseFolder()
        {
            return CbBaseType.SelectedIndex == 0 ? Constants.DwgBaseDirectory : DwgBaseHelpers.GetCustomBaseFolder();
        }
    }
}
