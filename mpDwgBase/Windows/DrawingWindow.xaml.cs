namespace mpDwgBase.Windows
{
    using System;
    using System.IO;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Media.Imaging;
    using System.Xml.Linq;
    using Models;
    using ModPlusAPI.Windows;
    using Utils;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using Visibility = System.Windows.Visibility;

    public partial class DrawingWindow
    {
        private const string LangItem = "mpDwgBase";

        private readonly bool _isEdit;
        public DwgBaseItem Item;
        private readonly string _mpDwgBaseFile;
        private readonly string _userDwgBaseFile;

        // Путь к папке с базами dwg плагина
        private readonly string _dwgBaseFolder;

        public DrawingWindow(string mpDwgBaseFile, string userDwgBaseFile, string dwgBaseFolder, bool isEdit)
        {
            InitializeComponent();
            _mpDwgBaseFile = mpDwgBaseFile;
            _userDwgBaseFile = userDwgBaseFile;
            _dwgBaseFolder = dwgBaseFolder;
            _isEdit = isEdit;
            FillHelpImagesToPopUp();
        }

        private void FillHelpImagesToPopUp()
        {
            var uri = 
                new Uri(
                    $"pack://application:,,,/mpDwgBase_{ModPlusConnector.Instance.AvailProductExternalVersion};component/Resources/helpImages/helpImage_1.png",
                    UriKind.RelativeOrAbsolute);
            helpImage_1.Source = BitmapFrame.Create(uri);
        }

        private void DrawingWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            if (!_isEdit)
            {
                // visibility of button to load last data
                BtLoadLastInteredData.Visibility = CheckFileWithLastDataExists()
                    ? Visibility.Visible
                    : Visibility.Collapsed;

                // set first item data
                Item.Path = "Чертежи/";

                // bind
                GridDrawingDetails.DataContext = Item;
            }
        }
        
        private void TbPath_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var cb = sender as ComboBox;
            var tb = cb?.Template.FindName("PART_EditableTextBox", cb) as TextBox;
            if (tb != null && tb.CaretIndex < 8)
            {
                e.Handled = true;
            }

            if (tb != null && !tb.Text.StartsWith("Чертежи/"))
            {
                e.Handled = true;
            }
        }

        private void TbPath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb?.Template.FindName("PART_EditableTextBox", cb) is TextBox tb && !tb.Text.StartsWith("Чертежи/"))
            {
                Dispatcher.BeginInvoke(new Action(() => tb.Undo()));
            }
        }

        private void TbPath_OnTextInput(object sender, TextCompositionEventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb?.Template.FindName("PART_EditableTextBox", cb) is TextBox tb && !tb.Text.StartsWith("Чертежи/"))
            {
                Dispatcher.BeginInvoke(new Action(() => tb.Undo()));
            }
        }

        // Выбрать dwg-файл для добавления 
        private void BtSelectDwgFile_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog(ModPlusAPI.Language.GetItem(LangItem, "msg59"), _dwgBaseFolder, "dwg", "name",
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.NoFtpSites |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.NoShellExtensions |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.NoUrls |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.ForceDefaultFolder |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles);
                var needLoop = true;
                while (needLoop)
                {
                    var ofdresult = ofd.ShowDialog();
                    if (ofdresult == System.Windows.Forms.DialogResult.OK)
                    {
                        var selectedFile = ofd.Filename;
                        if (selectedFile.Contains(_dwgBaseFolder))
                        {
                            if (DwgBaseHelpers.Is2013DwgVersion(selectedFile))
                            {
                                if (!DwgBaseHelpers.HasProxyEntities(selectedFile))
                                {
                                    TbSourceFile.Text =
                                        DwgBaseHelpers.TrimStart(selectedFile, _dwgBaseFolder).TrimStart('\\');
                                    BtAccept.IsEnabled = true;
                                    needLoop = false;
                                }
                                else
                                {
                                    ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg60"));
                                }
                            }
                            else
                            {
                                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg48"));
                            }
                        }
                        else
                        {
                            ModPlusAPI.Windows.MessageBox.Show(
                                $"{ModPlusAPI.Language.GetItem(LangItem, "msg49")} {_dwgBaseFolder}");
                        }
                    }
                    else if (ofdresult == System.Windows.Forms.DialogResult.Cancel)
                    {
                        return;
                    }
                    else
                    {
                        needLoop = false;
                    }
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
            finally
            {
                Focus();
            }
        }

        private void BtCopyDwgFile_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedFile = string.Empty;
                var selectedPath = string.Empty;
                var copiedFile = string.Empty;

                // Сначала нужно выбрать файл, проверив версию его
                var ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog(ModPlusAPI.Language.GetItem(LangItem, "msg59"), _dwgBaseFolder, "dwg", "name",
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.NoFtpSites |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.NoShellExtensions |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.NoUrls |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DefaultIsFolder |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.ForceDefaultFolder |
                    Autodesk.AutoCAD.Windows.OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles);
                var needLoop = true;
                while (needLoop)
                {
                    var ofdresult = ofd.ShowDialog();
                    if (ofdresult == System.Windows.Forms.DialogResult.OK)
                    {
                        selectedFile = ofd.Filename;
                        if (DwgBaseHelpers.Is2013DwgVersion(selectedFile))
                        {
                            if (!DwgBaseHelpers.HasProxyEntities(selectedFile))
                            {
                                needLoop = false;
                            }
                            else
                            {
                                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg60"));
                            }
                        }
                        else
                        {
                            ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg48"));
                        }
                    }
                    else if (ofdresult == System.Windows.Forms.DialogResult.Cancel)
                    {
                        return;
                    }
                    else
                    {
                        needLoop = false;
                    }
                }

                // Теперь нужно указать папку для расположения файла
                var fbd = new System.Windows.Forms.FolderBrowserDialog
                {
                    Description = ModPlusAPI.Language.GetItem(LangItem, "msg61"),
                    SelectedPath = _dwgBaseFolder,
                    ShowNewFolderButton = true
                };
                needLoop = true;
                while (needLoop)
                {
                    var fbdResult = fbd.ShowDialog();
                    if (fbdResult == System.Windows.Forms.DialogResult.OK)
                    {
                        selectedPath = fbd.SelectedPath;
                        if (selectedPath.Contains(_dwgBaseFolder))
                        {
                            if (!selectedPath.Equals(_dwgBaseFolder))
                            {
                                var fi = new FileInfo(selectedFile);
                                copiedFile = System.IO.Path.Combine(selectedPath, fi.Name);
                                if (File.Exists(copiedFile))
                                {
                                    needLoop =
                                        !ModPlusAPI.Windows.MessageBox.ShowYesNo(
                                            $"{ModPlusAPI.Language.GetItem(LangItem, "msg62")} {fi.Name}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg63")}", MessageBoxIcon.Question);
                                }
                                else
                                {
                                    needLoop = false;
                                }
                            }
                            else
                            {
                                ModPlusAPI.Windows.MessageBox.Show(
                                    $"{ModPlusAPI.Language.GetItem(LangItem, "msg64")} {_dwgBaseFolder}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg52")}");
                            }
                        }
                        else
                        {
                            ModPlusAPI.Windows.MessageBox.Show(
                                $"{ModPlusAPI.Language.GetItem(LangItem, "msg65")} {_dwgBaseFolder}");
                        }
                    }
                    else if (fbdResult == System.Windows.Forms.DialogResult.Cancel)
                    {
                        return;
                    }
                    else
                    {
                        needLoop = true;
                    }
                }

                // then copy file
                if (!string.IsNullOrEmpty(selectedFile) & !string.IsNullOrEmpty(selectedPath))
                {
                    if (File.Exists(selectedFile))
                    {
                        File.Copy(selectedFile, copiedFile, true);
                        if (File.Exists(copiedFile))
                        {
                            TbSourceFile.Text = DwgBaseHelpers.TrimStart(copiedFile, _dwgBaseFolder).TrimStart('\\');
                            BtAccept.IsEnabled = true;
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
            finally
            {
                Focus();
            }
        }

        // on accept click
        private void BtAccept_OnClick(object sender, RoutedEventArgs e)
        {
            if (!CheckEmptyData())
            {
                return;
            }

            if (!_isEdit)
            {
                var allGood = true;
                if (CheckInteredItemData())
                {
                    allGood = ModPlusAPI.Windows.MessageBox.ShowYesNo(ModPlusAPI.Language.GetItem(LangItem, "msg66"), MessageBoxIcon.Question);
                }

                if (allGood)
                {
                    SaveInteredData();
                    FillDwgBaseItemData();
                    Closed -= DrawingWindow_OnClosed;
                    DialogResult = true;
                    Close();
                }
            }
            else
            {
                FillDwgBaseItemData();
                Closed -= DrawingWindow_OnClosed;
                DialogResult = true;
                Close();
            }
        }

        // on cancel click
        private void BtCancel_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        /// <summary>
        /// Проверка введенных данных по базе
        /// </summary>
        /// <returns>True - есть такие данные, false - нет таких данных</returns>
        private bool CheckInteredItemData()
        {
            var hasSame = false;

            #region Проверяем по базе плагина

            if (DwgBaseHelpers.DeserializeFromXml(_mpDwgBaseFile, out var mpDwgBaseItems))
            {
                foreach (var mpDwgBaseItem in mpDwgBaseItems)
                {
                    if (!mpDwgBaseItem.IsBlock)
                    {
                        if (mpDwgBaseItem.Name.Equals(Item.Name))
                        {
                            ModPlusAPI.Windows.MessageBox.Show(
                                $"{ModPlusAPI.Language.GetItem(LangItem, "msg67")} {Item.Name}");
                            hasSame = true;
                        }
                    }
                }
            }
            else
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg36"));
                return false;
            }
            #endregion
            
            #region Проверяем по базе пользователя

            if (DwgBaseHelpers.DeserializeFromXml(_userDwgBaseFile, out var userDwgBaseItems))
            {
                foreach (var userDwgBaseItem in userDwgBaseItems)
                {
                    if (!userDwgBaseItem.IsBlock)
                    {
                        if (userDwgBaseItem.Name.Equals(Item.Name))
                        {
                            ModPlusAPI.Windows.MessageBox.Show(
                                $"{ModPlusAPI.Language.GetItem(LangItem, "msg68")} {Item.Name}");
                            hasSame = true;
                        }
                    }
                }
            }
            else
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg38"));
                return false;
            }
            #endregion
            
            return hasSame;
        }

        /// <summary>
        /// Проверка введенных данных на наличие пустых полей
        /// </summary>
        /// <returns>true - все заполнено, false - требуется заполнение</returns>
        private bool CheckEmptyData()
        {
            if (string.IsNullOrEmpty(Item.Name))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg69"));
                TbName.Focus();
                return false;
            }

            if (Item.Path.Equals("Чертежи/"))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg40"));
                CbPath.Focus();
                return false;
            }

            if (!_isEdit)
            {
                var db = AcApp.DocumentManager.MdiActiveDocument.Database;
                if (ChkIsCurrentDwgFile.IsChecked != null && ChkIsCurrentDwgFile.IsChecked.Value)
                {
                    if (!db.Filename.Contains(_dwgBaseFolder))
                    {
                        var fi = new FileInfo(db.Filename);
                        if (File.Exists(db.Filename))
                        {
                            ModPlusAPI.Windows.MessageBox.Show(
                                $"{ModPlusAPI.Language.GetItem(LangItem, "msg41")}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg42")} - {_dwgBaseFolder}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg43")} - {fi.DirectoryName}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg44")}");
                        }
                        else
                        {
                            ModPlusAPI.Windows.MessageBox.Show(
                                ModPlusAPI.Language.GetItem(LangItem, "msg45") +
                                Environment.NewLine + ModPlusAPI.Language.GetItem(LangItem, "msg44"));
                        }

                        return false;
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(Item.SourceFile))
                    {
                        ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg70"));
                        return false;
                    }
                }
            }

            return true;
        }

        // Заполнение текущего элемента данными
        private void FillDwgBaseItemData()
        {
            // Данные о блоке заполнены при выборе блока
            // Проверка введенных данных уже произведена
            Item.IsBlock = false;
            Item.IsAnnotative = false;
            Item.IsDynamicBlock = false;
            Item.Name = TbName.Text.Trim();
            Item.Description = TbDescription.Text.Trim();
            Item.Path = CbPath.Text.Trim().TrimEnd('/');
            Item.SourceFile = TbSourceFile.Text;
            Item.Document = TbDocument.Text.Trim();
            Item.HasAttributesForSpecification = false;
            Item.PositionValue = string.Empty;
            Item.DesignationValue = string.Empty;
            Item.MassValue = string.Empty;
            Item.NoteValue = string.Empty;
            Item.NominationValue = string.Empty;

            Item.Source = TbSource.Text.Trim();
            Item.Author = TbAuthor.Text.Trim();
            Item.Is3Dblock = false;
        }

        private void ChkIsCurrentDwgFile_OnChecked(object sender, RoutedEventArgs e)
        {
            // нужно проверить версию файла
            var db = AcApp.DocumentManager.MdiActiveDocument.Database;
            if (ChkIsCurrentDwgFile.IsChecked != null && ChkIsCurrentDwgFile.IsChecked.Value)
            {
                if (!db.Filename.Contains(_dwgBaseFolder))
                {
                    var fi = new FileInfo(db.Filename);
                    if (File.Exists(db.Filename))
                    {
                        ModPlusAPI.Windows.MessageBox.Show(
                            $"{ModPlusAPI.Language.GetItem(LangItem, "msg55")}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg42")} - {_dwgBaseFolder}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg43")} - {fi.DirectoryName}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg44")}");
                    }
                    else
                    {
                        ModPlusAPI.Windows.MessageBox.Show(
                            ModPlusAPI.Language.GetItem(LangItem, "msg56") +
                            Environment.NewLine + ModPlusAPI.Language.GetItem(LangItem, "msg44"));
                    }

                    ChkIsCurrentDwgFile.IsChecked = false;
                }
                else
                {
                    if (!File.Exists(db.Filename))
                    {
                        ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg56") +
                                                           Environment.NewLine +
                                                           ModPlusAPI.Language.GetItem(LangItem, "msg44"));
                        ChkIsCurrentDwgFile.IsChecked = false;
                    }
                }
            }
        }

        #region last intered data

        public bool CheckFileWithLastDataExists()
        {
            var file = System.IO.Path.Combine(_dwgBaseFolder, "LastInteredUserDataForDrawing.xml");
            return File.Exists(file);
        }

        private void SaveInteredData()
        {
            var file = Path.Combine(_dwgBaseFolder, "LastInteredUserDataForDrawing.xml");
            var xEl = new XElement("LastData");

            // Name
            xEl.SetAttributeValue("Name", TbName.Text);

            // Description
            xEl.SetAttributeValue("Description", TbDescription.Text);

            // document
            xEl.SetAttributeValue("Document", TbDocument.Text);

            // Author
            xEl.SetAttributeValue("Author", TbAuthor.Text);

            // Source
            xEl.SetAttributeValue("Source", TbSource.Text);

            // Path
            xEl.SetAttributeValue("Path", CbPath.Text);

            // Check if current file
            xEl.SetAttributeValue("IsCurrentDwgFile", ChkIsCurrentDwgFile.IsChecked != null && ChkIsCurrentDwgFile.IsChecked.Value);

            // Source file
            xEl.SetAttributeValue("SourceFile", TbSourceFile.Text);

            // save
            xEl.Save(file);
        }

        private void BtLoadLastInteredData_OnClick(object sender, RoutedEventArgs e)
        {
            var file = Path.Combine(_dwgBaseFolder, "LastInteredUserDataForDrawing.xml");
            var xEl = XElement.Load(file);

            // Name 
            TbName.Text = xEl.Attribute("Name")?.Value;

            // Description
            TbDescription.Text = xEl.Attribute("Description")?.Value;

            // Document
            TbDocument.Text = xEl.Attribute("Document")?.Value;

            // Author
            TbAuthor.Text = xEl.Attribute("Author")?.Value;

            // Source
            TbSource.Text = xEl.Attribute("Source")?.Value;

            // Path
            CbPath.Text = xEl.Attribute("Path")?.Value;

            // ChkIsCurrentDwgFile
            ChkIsCurrentDwgFile.IsChecked = bool.TryParse(xEl.Attribute("IsCurrentDwgFile")?.Value, out var b) && b; // false

            // Source File
            // ReSharper disable once AssignNullToNotNullAttribute
            var sf = System.IO.Path.Combine(_dwgBaseFolder, xEl.Attribute("SourceFile")?.Value);
            if (!string.IsNullOrEmpty(sf) && File.Exists(sf))
            {
                TbSourceFile.Text = xEl.Attribute("SourceFile")?.Value;
                BtAccept.IsEnabled = true;
            }
        }
        #endregion

        private void BtRecommend_OnClick(object sender, RoutedEventArgs e)
        {
            var win = new DrawingRecommend();
            win.ShowDialog();
        }

        private void DrawingWindow_OnClosed(object sender, EventArgs e)
        {
            if (_isEdit)
            {
                return;
            }

            if (!string.IsNullOrEmpty(TbName.Text) | !string.IsNullOrEmpty(TbDescription.Text) |
                !string.IsNullOrEmpty(TbAuthor.Text) | !string.IsNullOrEmpty(TbDocument.Text) |
                !string.IsNullOrEmpty(TbSource.Text) | !string.IsNullOrEmpty(TbSourceFile.Text))
            {
                if (ModPlusAPI.Windows.MessageBox.ShowYesNo(ModPlusAPI.Language.GetItem(LangItem, "msg58"), MessageBoxIcon.Question))
                {
                    SaveInteredData();
                }
            }
        }
    }
}
