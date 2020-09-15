namespace mpDwgBase.Windows
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics.CodeAnalysis;
    using System.IO;
    using System.Linq;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Media.Imaging;
    using System.Xml.Linq;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Models;
    using ModPlusAPI.Mvvm;
    using ModPlusAPI.Windows;
    using Utils;
    using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
    using Visibility = System.Windows.Visibility;

    public partial class BlockWindow
    {
        private const string LangItem = "mpDwgBase";

        private readonly bool _isEdit;
        public DwgBaseItem Item;
        private ObjectId _selectedBlockObjectId;
        private readonly string _mpDwgBaseFile;
        private readonly string _userDwgBaseFile;

        // Путь к папке с базами dwg плагина
        private readonly string _dwgBaseFolder;

        public BlockWindow(string mpDwgBaseFile, string userDwgBaseFile, string dwgBaseFolder, bool isEdit)
        {
            InitializeComponent();
            _mpDwgBaseFile = mpDwgBaseFile;
            _userDwgBaseFile = userDwgBaseFile;
            _dwgBaseFolder = dwgBaseFolder;
            _isEdit = isEdit;

            // Отключаем видимость подробностей до выбора блока
            GridBlockDetails.Visibility = Visibility.Collapsed;
            BtLoadLastEnteredData.Visibility = Visibility.Collapsed;
            DgAttributesForSpecifictaion.Visibility = Visibility.Collapsed;

            // block exist visibility
            BtAboutBlockExistAttributes.Visibility = Visibility.Collapsed;
            TbBlkExistHeader.Visibility = Visibility.Collapsed;
            DgBlockExistAttributes.Visibility = Visibility.Collapsed;

            FillHelpImagesToPopUp();
        }

        private void BlockWindow_OnLoaded(object sender, RoutedEventArgs e)
        {
            // if edit - attributes
            if (_isEdit)
            {
                FillAttributesIfIsEdit();
            }
        }

        private void FillAttributesIfIsEdit()
        {
            ChkHasAttributesForSpecification.Checked -= ChkHasAttributesForSpecification_OnChecked;
            ChkHasAttributesForSpecification.IsChecked = Item.HasAttributesForSpecification;
            DgAttributesForSpecifictaion.Visibility = Visibility.Visible;
            ChkHasAttributesForSpecification.Checked += ChkHasAttributesForSpecification_OnChecked;

            // fill
            FillAttributesForSpecification(Item);

            // bind
            DgAttributesForSpecifictaion.ItemsSource = _attributesForSpecification;

            // атрибуты самого блока
            if (Item.AttributesValues.Any())
            {
                var lstToBind = Item.AttributesValues.ToList();
                BtAboutBlockExistAttributes.Visibility = Visibility.Visible;
                TbBlkExistHeader.Visibility = Visibility.Visible;
                DgBlockExistAttributes.Visibility = Visibility.Visible;
                DgBlockExistAttributes.ItemsSource = lstToBind;
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
                    allGood = ModPlusAPI.Windows.MessageBox.ShowYesNo(ModPlusAPI.Language.GetItem(LangItem, "msg32"), MessageBoxIcon.Question);
                }

                if (allGood)
                {
                    SaveInteredData();
                    FillDwgBaseItemData();
                    if (ChkIsCurrentDwgFile.IsChecked != null && ChkIsCurrentDwgFile.IsChecked.Value)
                    {
                        Closed -= BlockWindow_OnClosed;
                        DialogResult = true;
                        Close();
                    }
                    else if (CopySelectedBlockToFile())
                    {
                        Closed -= BlockWindow_OnClosed;
                        DialogResult = true;
                        Close();
                    }
                    else
                    {
                        ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg33"));
                    }
                }
            }
            else
            {
                FillDwgBaseItemData();
                Closed -= BlockWindow_OnClosed;
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

        // Выбор блока
        private void BtSelectBlock_OnClick(object sender, RoutedEventArgs e)
        {
            Hide();

            _selectedBlockObjectId = GetBlock();

            if (_selectedBlockObjectId != ObjectId.Null)
            {
                if (ReadDetailsFromSelectedBlock())
                {
                    TbBlockName.Text = Item.BlockName;
                    TbIsAnnot.Opacity = Item.IsAnnotative ? 1.0 : 0.5;
                    TbIsDynamic.Opacity = Item.IsDynamicBlock ? 1.0 : 0.5;

                    // Биндим 
                    GridBlockDetails.DataContext = Item;

                    // включаем видимость подробностей
                    GridBlockDetails.Visibility = Visibility.Visible;

                    BtAccept.IsEnabled = true;

                    // visibility of button to load last data
                    BtLoadLastEnteredData.Visibility = CheckFileWithLastDataExists() ? Visibility.Visible : Visibility.Collapsed;

                    // move window to center of current screen
                    WindowStartupLocation = WindowStartupLocation.Manual;
                    UpdateLayout();
                    var screen = System.Windows.Forms.Screen.FromHandle(new System.Windows.Interop.WindowInteropHelper(this).Handle);
                    var workArea = screen.WorkingArea;
                    Left = ((workArea.Width - Width) / 2) + workArea.Left;
                    Top = ((workArea.Height - Height) / 2) + workArea.Top;
                }
            }

            ShowDialog();
        }

        // Получение ObjectId для блока
        private static ObjectId GetBlock()
        {
            var doc = AcApp.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;
            try
            {
                var peo = new PromptEntityOptions($"\n{ModPlusAPI.Language.GetItem(LangItem, "msg34")}:");
                peo.SetRejectMessage($"\n{ModPlusAPI.Language.GetItem(LangItem, "wrong")}");
                peo.AllowNone = true;
                peo.AllowObjectOnLockedLayer = true;
                peo.AddAllowedClass(typeof(BlockReference), true);

                var per = ed.GetEntity(peo);
                return per.Status == PromptStatus.OK ? per.ObjectId : ObjectId.Null;
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                return ObjectId.Null;
            }
        }

        // Получение данных из блока и добавление их в текущий DwgBaseItem
        private bool ReadDetailsFromSelectedBlock()
        {
            try
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var selBlkRef = tr.GetObject(_selectedBlockObjectId, OpenMode.ForRead, true) as BlockReference;
                    if (selBlkRef == null)
                    {
                        throw new ArgumentNullException(nameof(selBlkRef));
                    }

                    // С учетом http://adn-cis.org/opredelenie-imeni-bloka-po-vstavke-bloka.html
                    BlockTableRecord selBlock;
                    if (selBlkRef.IsDynamicBlock)
                    {
                        selBlock = tr.GetObject(selBlkRef.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    }
                    else
                    {
                        selBlock = tr.GetObject(selBlkRef.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    }

                    if (selBlock != null)
                    {
                        Item.BlockName = selBlock.Name;
                        Item.IsAnnotative = selBlock.Annotative == AnnotativeStates.True;
                        Item.IsDynamicBlock = selBlock.IsDynamicBlock;

                        // по умолчанию задаем путь
                        Item.Path = "Блоки/";

                        // read attributes
                        if (selBlock.HasAttributeDefinitions)
                        {
                            // Если есть атрибуты в описании блока, то включаем видимость кнопки
                            // и заполняем значения атрибутов из описания блоков
                            BtGetAttrValuesFromBlock.Visibility = Visibility.Visible;
                            GetAttributesFromBlockDefinition(selBlock, tr);
                        }
                        else
                        {
                            BtGetAttrValuesFromBlock.Visibility = Visibility.Collapsed;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                return false;
            }
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
                    if (mpDwgBaseItem.IsBlock)
                    {
                        if (mpDwgBaseItem.Name.Equals(Item.Name))
                        {
                            ModPlusAPI.Windows.MessageBox.Show(
                                $"{ModPlusAPI.Language.GetItem(LangItem, "msg35")}: {Item.Name}");
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
                    if (userDwgBaseItem.IsBlock)
                    {
                        if (userDwgBaseItem.Name.Equals(Item.Name))
                        {
                            ModPlusAPI.Windows.MessageBox.Show(
                                $"{ModPlusAPI.Language.GetItem(LangItem, "msg37")}: {Item.Name}");
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

        /// <summary>Проверка введенных данных на наличие пустых полей</summary>
        /// <returns>true - все заполнено, false - требуется заполнение</returns>
        private bool CheckEmptyData()
        {
            if (string.IsNullOrEmpty(Item.Name))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg39"));
                TbName.Focus();
                return false;
            }

            if (Item.Path.Equals("Блоки/"))
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
                                $"{ModPlusAPI.Language.GetItem(LangItem, "msg41")}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg42")} - {_dwgBaseFolder}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg43")} - {fi.DirectoryName}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg44")}"
                            );
                        }
                        else
                        {
                            ModPlusAPI.Windows.MessageBox.Show(
                                ModPlusAPI.Language.GetItem(LangItem, "msg45") +
                                Environment.NewLine + ModPlusAPI.Language.GetItem(LangItem, "msg44")
                            );
                        }

                        return false;
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(Item.SourceFile))
                    {
                        ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg46"));
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
            Item.IsBlock = true;

            Item.Name = TbName.Text.Trim();
            Item.Description = TbDescription.Text.Trim();
            Item.Path = CbPath.Text.Trim().TrimEnd('/');
            Item.SourceFile = TbSourceFile.Text;
            Item.Document = TbDocument.Text.Trim();
            Item.HasAttributesForSpecification = ChkHasAttributesForSpecification.IsChecked != null && ChkHasAttributesForSpecification.IsChecked.Value;
            if (Item.HasAttributesForSpecification)
            {
                Item.PositionValue = _attributesForSpecification.First(x => x.Name.Equals("Position")).BaseValue;
                Item.DesignationValue = _attributesForSpecification.First(x => x.Name.Equals("Designation")).BaseValue;
                Item.NominationValue = _attributesForSpecification.First(x => x.Name.Equals("Nomination")).BaseValue;
                Item.MassValue = _attributesForSpecification.First(x => x.Name.Equals("Mass")).BaseValue;
                Item.NoteValue = _attributesForSpecification.First(x => x.Name.Equals("Note")).BaseValue;
            }
            else
            {
                Item.PositionValue = string.Empty;
                Item.DesignationValue = string.Empty;
                Item.MassValue = string.Empty;
                Item.NoteValue = string.Empty;
                Item.NominationValue = string.Empty;
            }

            Item.Source = TbSource.Text.Trim();
            Item.Author = TbAuthor.Text.Trim();
            Item.Is3Dblock = ChkIs3Dblock.IsChecked != null && ChkIs3Dblock.IsChecked.Value;

            // attributes values
            if (!_isEdit)
            {
                if (_blockExistAttributes != null && _blockExistAttributes.Any())
                {
                    Item.AttributesValues = new List<AttributeValue>();
                    foreach (var blockExistAttribute in _blockExistAttributes)
                    {
                        // Item.AttributesValues.Add(blockExistAttribute.TextString);
                        Item.AttributesValues.Add(new AttributeValue
                        {
                            Tag = blockExistAttribute.Tag,
                            Prompt = blockExistAttribute.Prompt,
                            TextString = blockExistAttribute.TextString
                        });
                    }
                }
            }
            else
            {
                if (DgBlockExistAttributes.Items.Count > 0)
                {
                    Item.AttributesValues = new List<AttributeValue>();
                    foreach (var item in DgBlockExistAttributes.Items)
                    {
                        var bindedAttr = item as AttributeValue;
                        Item.AttributesValues.Add(bindedAttr);
                    }
                }
            }
        }

        /// <summary>Копирование блока в файл</summary>
        private bool CopySelectedBlockToFile()
        {
            var tempFile = Path.GetTempFileName();
            try
            {
                var doc = AcApp.DocumentManager.MdiActiveDocument;
                var currentDb = doc.Database;
                var file = Path.Combine(_dwgBaseFolder, TbSourceFile.Text);
                if (File.Exists(file))
                {
                    using (var destDb = new Database(false, true))
                    {
                        destDb.ReadDwgFile(file, FileOpenMode.OpenForReadAndWriteNoShare, true, string.Empty);

                        // Create a variable to store the list of block identifiers
                        var blockIds = new ObjectIdCollection { _selectedBlockObjectId };
                        var mapping = new IdMapping();

                        // copy
                        currentDb.WblockCloneObjects(blockIds, destDb.BlockTableId, mapping,
                            DuplicateRecordCloning.Ignore,
                            false);
                        destDb.SaveAs(tempFile, DwgVersion.AC1027);
                    }

                    // now replace
                    File.Copy(tempFile, file, true);
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                return false;
            }
            finally
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }

        private void TbPath_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var cb = sender as ComboBox;
            var tb = cb?.Template.FindName("PART_EditableTextBox", cb) as TextBox;
            if (tb != null && tb.CaretIndex < 6)
            {
                e.Handled = true;
            }

            if (tb != null && !tb.Text.StartsWith("Блоки/"))
            {
                e.Handled = true;
            }
        }

        private void TbPath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb?.Template.FindName("PART_EditableTextBox", cb) is TextBox tb && !tb.Text.StartsWith("Блоки/"))
            {
                Dispatcher?.BeginInvoke(new Action(() => tb.Undo()));
            }
        }

        private void TbPath_OnTextInput(object sender, TextCompositionEventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb?.Template.FindName("PART_EditableTextBox", cb) is TextBox tb && !tb.Text.StartsWith("Блоки/"))
            {
                Dispatcher?.BeginInvoke(new Action(() => tb.Undo()));
            }
        }

        private void FillHelpImagesToPopUp()
        {
            var uri =
                new Uri(
                    $"pack://application:,,,/mpDwgBase_{ModPlusConnector.Instance.AvailProductExternalVersion};component/Resources/helpImages/helpImage_1.png",
                    UriKind.RelativeOrAbsolute);
            helpImage_1.Source = BitmapFrame.Create(uri);
        }

        // Выбрать dwg-файл для добавления в него нового блока
        private void BtSelectDwgFile_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var ofd = new Autodesk.AutoCAD.Windows.OpenFileDialog(
                    ModPlusAPI.Language.GetItem(LangItem, "msg47"), _dwgBaseFolder, "dwg", "name",
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
                                TbSourceFile.Text = DwgBaseHelpers.TrimStart(selectedFile, _dwgBaseFolder).TrimStart('\\');
                                needLoop = false;
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

        // Создать dwg-файл для добавления в него нового блока
        private void BtCreateDwgFile_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var sfd = new Autodesk.AutoCAD.Windows.SaveFileDialog(ModPlusAPI.Language.GetItem(LangItem, "msg50"), _dwgBaseFolder, "dwg", "name",
                    Autodesk.AutoCAD.Windows.SaveFileDialog.SaveFileDialogFlags.DefaultIsFolder |
                    Autodesk.AutoCAD.Windows.SaveFileDialog.SaveFileDialogFlags.DoNotTransferRemoteFiles |
                    Autodesk.AutoCAD.Windows.SaveFileDialog.SaveFileDialogFlags.ForceDefaultFolder |
                    Autodesk.AutoCAD.Windows.SaveFileDialog.SaveFileDialogFlags.NoFtpSites |
                    Autodesk.AutoCAD.Windows.SaveFileDialog.SaveFileDialogFlags.NoShellExtensions |
                    Autodesk.AutoCAD.Windows.SaveFileDialog.SaveFileDialogFlags.NoUrls);
                var needLoop = true;
                while (needLoop)
                {
                    var sfdresult = sfd.ShowDialog();
                    if (sfdresult == System.Windows.Forms.DialogResult.OK)
                    {
                        var selectedFile = sfd.Filename;
                        if (selectedFile.Contains(_dwgBaseFolder))
                        {
                            var fi = new FileInfo(selectedFile);
                            if (fi.DirectoryName != null && !fi.DirectoryName.Equals(_dwgBaseFolder))
                            {
                                try
                                {
                                    using (var db = new Database())
                                    {
                                        db.SaveAs(selectedFile, DwgVersion.AC1027);
                                        TbSourceFile.Text =
                                            DwgBaseHelpers.TrimStart(selectedFile, _dwgBaseFolder).TrimStart('\\');
                                        needLoop = false;
                                    }
                                }
                                catch (Exception exception)
                                {
                                    ExceptionBox.Show(exception);
                                    needLoop = false;
                                }
                            }
                            else
                            {
                                ModPlusAPI.Windows.MessageBox.Show(
                                    $"{ModPlusAPI.Language.GetItem(LangItem, "msg51")} {_dwgBaseFolder}{Environment.NewLine}{ModPlusAPI.Language.GetItem(LangItem, "msg52")}");
                            }
                        }
                        else
                        {
                            ModPlusAPI.Windows.MessageBox.Show(
                                $"{ModPlusAPI.Language.GetItem(LangItem, "msg49")} {_dwgBaseFolder}");
                        }
                    }
                    else if (sfdresult == System.Windows.Forms.DialogResult.Cancel)
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

        #region last intered data

        public bool CheckFileWithLastDataExists()
        {
            var file = Path.Combine(_dwgBaseFolder, "LastInteredUserDataForBlock.xml");
            return File.Exists(file);
        }

        private void SaveInteredData()
        {
            var file = Path.Combine(_dwgBaseFolder, "LastInteredUserDataForBlock.xml");
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

            // attributes for spec
            xEl.SetAttributeValue("HasAttributesForSpecification", ChkHasAttributesForSpecification.IsChecked != null && ChkHasAttributesForSpecification.IsChecked.Value);
            if (_attributesForSpecification != null && _attributesForSpecification.Any())
            {
                var attrEl = new XElement("AttributesForSpecification");
                foreach (AttributeForSpecification attributeForSpecification in _attributesForSpecification)
                {
                    var att = new XElement("AttributeForSpecification");
                    att.SetAttributeValue("BaseValue", attributeForSpecification.BaseValue);
                    att.SetAttributeValue("Name", attributeForSpecification.Name);
                    att.SetAttributeValue("DisplayedName", attributeForSpecification.DisplayedName);
                    attrEl.Add(att);
                }

                xEl.Add(attrEl);
            }

            // save
            xEl.Save(file);
        }

        [SuppressMessage("ReSharper", "AssignNullToNotNullAttribute")]
        private void BtLoadLastEnteredData_OnClick(object sender, RoutedEventArgs e)
        {
            var file = Path.Combine(_dwgBaseFolder, "LastInteredUserDataForBlock.xml");
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
            ChkIsCurrentDwgFile.IsChecked = bool.TryParse(xEl.Attribute("IsCurrentDwgFile")?.Value, out bool b) && b; // false

            // Source File
            // ReSharper disable once AssignNullToNotNullAttribute
            var sf = Path.Combine(_dwgBaseFolder, xEl.Attribute("SourceFile")?.Value);
            if (!string.IsNullOrEmpty(sf) && File.Exists(sf))
            {
                TbSourceFile.Text = xEl.Attribute("SourceFile")?.Value;
            }

            // attributes for spec
            XAttribute attribute = xEl.Attribute("HasAttributesForSpecification");
            if (attribute != null)
            {
                var value = attribute.Value;
                if (bool.TryParse(value, out var attrForSpec))
                {
                    ChkHasAttributesForSpecification.IsChecked = attrForSpec;
                    if (attrForSpec)
                    {
                        var attEl = xEl.Element("AttributesForSpecification");
                        if (attEl != null)
                        {
                            FillAttributesForSpecification(); // на всякий случай
                            DgAttributesForSpecifictaion.ItemsSource = null;
                            foreach (XElement xElement in attEl.Elements("AttributeForSpecification"))
                            {
                                foreach (AttributeForSpecification attributeForSpecification in _attributesForSpecification)
                                {
                                    var xAttribute = xElement.Attribute("Name");
                                    if (xAttribute != null && attributeForSpecification.Name.Equals(xAttribute.Value))
                                    {
                                        attributeForSpecification.BaseValue = xElement.Attribute("BaseValue")?.Value;
                                    }
                                }
                            }

                            DgAttributesForSpecifictaion.ItemsSource = _attributesForSpecification;
                        }
                    }
                }
            }
        }
        #endregion

        #region Attributes for specification

        internal class AttributeForSpecification : VmBase
        {
            private string _baseValue;
            
            public string Name { get; set; }

            public string DisplayedName { get; set; }

            public string BaseValue
            {
                get => _baseValue;
                set
                {
                    _baseValue = value;
                    OnPropertyChanged();
                }
            }
        }

        private List<AttributeForSpecification> _attributesForSpecification;

        private void FillAttributesForSpecification()
        {
            _attributesForSpecification = new List<AttributeForSpecification>();
            _attributesForSpecification.Add(new AttributeForSpecification
            {
                Name = "Position",
                DisplayedName = ModPlusAPI.Language.GetItem(LangItem, "a1"),
                BaseValue = string.Empty
            });
            _attributesForSpecification.Add(new AttributeForSpecification
            {
                Name = "Designation",
                DisplayedName = ModPlusAPI.Language.GetItem(LangItem, "a2"),
                BaseValue = string.Empty
            });
            _attributesForSpecification.Add(new AttributeForSpecification
            {
                Name = "Nomination",
                DisplayedName = ModPlusAPI.Language.GetItem(LangItem, "a3"),
                BaseValue = string.Empty
            });
            _attributesForSpecification.Add(new AttributeForSpecification
            {
                Name = "Mass",
                DisplayedName = ModPlusAPI.Language.GetItem(LangItem, "a4"),
                BaseValue = string.Empty
            });
            _attributesForSpecification.Add(new AttributeForSpecification
            {
                Name = "Note",
                DisplayedName = ModPlusAPI.Language.GetItem(LangItem, "a5"),
                BaseValue = string.Empty
            });
        }

        private void FillAttributesForSpecification(DwgBaseItem editingItem)
        {
            _attributesForSpecification = new List<AttributeForSpecification>();
            _attributesForSpecification.Add(new AttributeForSpecification
            {
                Name = "Position",
                DisplayedName = ModPlusAPI.Language.GetItem(LangItem, "a1"),
                BaseValue = editingItem.PositionValue
            });
            _attributesForSpecification.Add(new AttributeForSpecification
            {
                Name = "Designation",
                DisplayedName = ModPlusAPI.Language.GetItem(LangItem, "a2"),
                BaseValue = editingItem.DesignationValue
            });
            _attributesForSpecification.Add(new AttributeForSpecification
            {
                Name = "Nomination",
                DisplayedName = ModPlusAPI.Language.GetItem(LangItem, "a3"),
                BaseValue = editingItem.NominationValue
            });
            _attributesForSpecification.Add(new AttributeForSpecification
            {
                Name = "Mass",
                DisplayedName = ModPlusAPI.Language.GetItem(LangItem, "a4"),
                BaseValue = editingItem.MassValue
            });
            _attributesForSpecification.Add(new AttributeForSpecification
            {
                Name = "Note",
                DisplayedName = ModPlusAPI.Language.GetItem(LangItem, "a5"),
                BaseValue = editingItem.NoteValue
            });
        }

        private void ChkHasAttributesForSpecification_OnChecked(object sender, RoutedEventArgs e)
        {
            if (Item.IsDynamicBlock)
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "msg53"));
            }

            DgAttributesForSpecifictaion.Visibility = Visibility.Visible;

            // fill
            FillAttributesForSpecification();

            // bind
            DgAttributesForSpecifictaion.ItemsSource = _attributesForSpecification;
        }

        private void ChkHasAttributesForSpecification_OnUnchecked(object sender, RoutedEventArgs e)
        {
            DgAttributesForSpecifictaion.Visibility = Visibility.Collapsed;
        }
        #endregion

        #region block exist attributes

        internal class BlockExistAttribute
        {
            public string Invisible { get; set; }

            public string Constant { get; set; }

            public string Tag { get; set; }

            public string TextString { get; set; }

            public string Prompt { get; set; }
        }

        private List<BlockExistAttribute> _blockExistAttributes;

        // Получение из описания блока атрибутов
        private void GetAttributesFromBlockDefinition(BlockTableRecord blk, Transaction tr)
        {
            if (blk.HasAttributeDefinitions)
            {
                _blockExistAttributes = new List<BlockExistAttribute>();
                foreach (ObjectId objectId in blk)
                {
                    var attrDef = tr.GetObject(objectId, OpenMode.ForRead) as AttributeDefinition;
                    if (attrDef != null)
                    {
                        var blkExistAttr = new BlockExistAttribute
                        {
                            Constant = attrDef.Constant ? ModPlusAPI.Language.GetItem(LangItem, "yes") : ModPlusAPI.Language.GetItem(LangItem, "no"),
                            Invisible = attrDef.Invisible ? ModPlusAPI.Language.GetItem(LangItem, "yes") : ModPlusAPI.Language.GetItem(LangItem, "no"),
                            Prompt = attrDef.Prompt,
                            Tag = attrDef.Tag,
                            TextString = attrDef.TextString
                        };
                        _blockExistAttributes.Add(blkExistAttr);
                    }
                }

                // binding
                if (_blockExistAttributes.Any())
                {
                    BtAboutBlockExistAttributes.Visibility = Visibility.Visible;
                    TbBlkExistHeader.Visibility = Visibility.Visible;
                    DgBlockExistAttributes.Visibility = Visibility.Visible;
                    DgBlockExistAttributes.ItemsSource = _blockExistAttributes;
                }
            }
        }

        // Получение значений атрибутов из вхождения блока
        private void GetAttributesValuesFromBlockReference(BlockReference blkRef, Transaction tr)
        {
            if (blkRef.AttributeCollection.Count > 0 & _blockExistAttributes.Any())
            {
                foreach (ObjectId objectId in blkRef.AttributeCollection)
                {
                    var attrRef = tr.GetObject(objectId, OpenMode.ForRead) as AttributeReference;
                    if (attrRef != null)
                    {
                        foreach (var blockExistAttribute in _blockExistAttributes)
                        {
                            if (blockExistAttribute.Tag.Equals(attrRef.Tag))
                            {
                                blockExistAttribute.TextString = attrRef.TextString;
                            }
                        }
                    }
                }
            }
        }

        // Взять значения атрибутов из блока
        private void BtGetAttrValuesFromBlock_OnClick(object sender, RoutedEventArgs e)
        {
            if (ModPlusAPI.Windows.MessageBox.ShowYesNo(ModPlusAPI.Language.GetItem(LangItem, "msg54"), MessageBoxIcon.Question))
            {
                try
                {
                    DgBlockExistAttributes.ItemsSource = null;

                    var doc = AcApp.DocumentManager.MdiActiveDocument;
                    using (var tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        var blkRef = tr.GetObject(_selectedBlockObjectId, OpenMode.ForRead) as BlockReference;
                        GetAttributesValuesFromBlockReference(blkRef, tr);
                    }

                    DgBlockExistAttributes.ItemsSource = _blockExistAttributes;
                }
                catch (Exception exception)
                {
                    ExceptionBox.Show(exception);
                }
            }
        }

        #endregion

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

        private void BtRecommend_OnClick(object sender, RoutedEventArgs e)
        {
            var win = new BlockRecommend();
            win.ShowDialog();
        }
        
        private void BlockWindow_OnClosed(object sender, EventArgs e)
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
