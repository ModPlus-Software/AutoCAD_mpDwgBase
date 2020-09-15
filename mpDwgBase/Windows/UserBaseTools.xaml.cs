namespace mpDwgBase.Windows
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Threading;
    using Models;
    using ModPlusAPI;
    using ModPlusAPI.Mvvm;
    using ModPlusAPI.Windows;
    using Utils;

    public partial class UserBaseTools
    {
        private const string LangItem = "mpDwgBase";

        // Переменная будет хранить значение были ли внесены изменения в базу
        public bool UserBaseChanged;

        // Путь к файлу, содержащему описание базы
        private readonly string _userDwgBaseFile;

        // Путь к файлу БД плагина
        private readonly string _mpDwgBaseFile;

        // Путь к папке с базами dwg плагина
        private readonly string _dwgBaseFolder;

        // Список значений ПОЛЬЗОВАТЕЛЬСКОЙ базы
        private List<DwgBaseItem> _dwgBaseItems;

        // Create a Delegate that matches the Signature of the ProgressBar's SetValue method
        private delegate void UpdateProgressBarDelegate(DependencyProperty dp, object value);

        private delegate void UpdateProgressTextDelegate(DependencyProperty dp, object value);

        public UserBaseTools(string mpDwgBaseFile, string userDwgBaseFile, string dwgBaseFolder, List<DwgBaseItem> dwgBaseItems)
        {
            InitializeComponent();
            Title = ModPlusAPI.Language.GetItem(LangItem, "u1");
            _mpDwgBaseFile = mpDwgBaseFile;
            _userDwgBaseFile = userDwgBaseFile;
            _dwgBaseFolder = dwgBaseFolder;
            _dwgBaseItems = dwgBaseItems;

            UnusedFiles_DwgBaseFolder.Text = _dwgBaseFolder;
        }

        #region main window

        private void UserBaseTools_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
            }

            // select all
            if (e.Key == Key.A && Keyboard.IsKeyDown(Key.LeftCtrl))
            {
                if (MultiChangePath_Tab.IsSelected)
                {
                    if (MultiChangePath_LvItems.Items.Count > 0)
                    {
                        foreach (DwgBaseItemWithSelector item in MultiChangePath_LvItems.Items)
                        {
                            item.Selected = true;
                        }
                    }
                }

                if (MultiChangeSourceAuthor_Tab.IsSelected)
                {
                    if (MultiChangeSourceAuthor_LvItems.Items.Count > 0)
                    {
                        foreach (DwgBaseItemWithSelector item in MultiChangeSourceAuthor_LvItems.Items)
                        {
                            item.Selected = true;
                        }
                    }
                }

                if (UnusedItems_Tab.IsSelected)
                {
                    if (UnusedItems_LvFiles.Items.Count > 0)
                    {
                        foreach (DwgBaseItemWithSelector item in UnusedItems_LvFiles.Items)
                        {
                            item.Selected = true;
                        }
                    }
                }

                if (UnusedFiles_Tab.IsSelected)
                {
                    if (UnusedFiles_LvFiles.Items.Count > 0)
                    {
                        foreach (UnusedFile item in UnusedFiles_LvFiles.Items)
                        {
                            item.Selected = true;
                        }
                    }
                }
            }
        }

        // main tab selection changed
        private void TabControlTools_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is TabControl)
            {
                // Изменить Путь для нескольких блоков
                if (MultiChangePath_Tab.IsSelected)
                {
                    if (MultiChangePath_CbMainGroup.SelectedIndex == -1)
                    {
                        MultiChangePath_CbMainGroup.SelectedIndex = 0;
                    }
                }

                if (MultiChangeSourceAuthor_Tab.IsSelected)
                {
                    if (MultiChangeSourceAuthor_LvItems.Items.Count == 0)
                    {
                        MultiChangeSourceAuthor_FillData();
                    }
                }

                if (RenameSourceFile_Tab.IsSelected)
                {
                    if (RenameSourceFile_LbFiles.Items.Count == 0)
                    {
                        RenameFiles_FindAllFiles();
                    }
                }
            }
        }
        #endregion

        #region Path events
        private void Blocks_TbPath_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
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

        private void Blocks_TbPath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb?.Template.FindName("PART_EditableTextBox", cb) is TextBox tb && !tb.Text.StartsWith("Блоки/"))
            {
                Dispatcher?.Invoke(new Action(() => tb.Undo()));
            }
        }

        private void Blocks_TbPath_OnTextInput(object sender, TextCompositionEventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb?.Template.FindName("PART_EditableTextBox", cb) is TextBox tb && !tb.Text.StartsWith("Блоки/"))
            {
                Dispatcher?.Invoke(new Action(() => tb.Undo()));
            }
        }

        private void Drawings_TbPath_OnPreviewTextInput(object sender, TextCompositionEventArgs e)
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

        private void Drawings_TbPath_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb?.Template.FindName("PART_EditableTextBox", cb) is TextBox tb && !tb.Text.StartsWith("Чертежи/"))
            {
                Dispatcher?.Invoke(new Action(() => tb.Undo()));
            }
        }

        private void Drawings_TbPath_OnTextInput(object sender, TextCompositionEventArgs e)
        {
            var cb = sender as ComboBox;
            if (cb?.Template.FindName("PART_EditableTextBox", cb) is TextBox tb && !tb.Text.StartsWith("Чертежи/"))
            {
                Dispatcher?.Invoke(new Action(() => tb.Undo()));
            }
        }
        #endregion

        #region MultiChangePath
        private void MultiChangePath_CbMainGroup_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!(sender is ComboBox))
            {
                return;
            }

            MultiChangePath_FillData();
        }

        /// <summary>
        /// Заполнение окна данными. Делается методом для возможности повторного вызова после внесения изменений
        /// </summary>
        private void MultiChangePath_FillData()
        {
            var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
            var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);

            var cb = MultiChangePath_CbMainGroup;
            var tb = MultiChangePath_Path.Template.FindName("PART_EditableTextBox", MultiChangePath_Path) as TextBox;

            // remove events
            MultiChangePath_Path.TextInput -= Blocks_TbPath_OnTextInput;
            MultiChangePath_Path.PreviewTextInput -= Blocks_TbPath_OnPreviewTextInput;
            if (tb != null)
            {
                tb.TextChanged -= Blocks_TbPath_OnTextChanged;
            }

            MultiChangePath_Path.TextInput -= Drawings_TbPath_OnTextInput;
            MultiChangePath_Path.PreviewTextInput -= Drawings_TbPath_OnPreviewTextInput;
            if (tb != null)
            {
                tb.TextChanged -= Drawings_TbPath_OnTextChanged;
            }

            // fill cb with path
            var paths = new List<string>();
            foreach (var dwgBaseItem in _dwgBaseItems)
            {
                if (!paths.Contains(dwgBaseItem.Path))
                {
                    paths.Add(dwgBaseItem.Path);
                }
            }

            if (cb.SelectedIndex == 0) // blocks
            {
                MultiChangePath_Path.Text = "Блоки/";
                MultiChangePath_Path.TextInput += Blocks_TbPath_OnTextInput;
                MultiChangePath_Path.PreviewTextInput += Blocks_TbPath_OnPreviewTextInput;
                if (tb != null)
                {
                    tb.TextChanged += Blocks_TbPath_OnTextChanged;
                }

                MultiChangePath_Path.ItemsSource = paths.Where(x => x.StartsWith("Блоки/"));

                // fill items
                MultiChangePath_LvItems.ItemsSource = null;
                var lst = new List<DwgBaseItemWithSelector>();
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = _dwgBaseItems.Count;
                ProgressBar.Value = 0;
                var index = 1;
                foreach (var dwgBaseItem in _dwgBaseItems)
                {
                    if (dwgBaseItem.IsBlock)
                    {
                        lst.Add(new DwgBaseItemWithSelector { Selected = false, Item = dwgBaseItem });
                    }

                    Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Render, TextBlock.TextProperty,
                        $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{_dwgBaseItems.Count}");
                    Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Render,
                        System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);

                    index++;
                }

                MultiChangePath_LvItems.ItemsSource = lst;
            }

            if (cb.SelectedIndex == 1) // drawings
            {
                MultiChangePath_Path.Text = "Чертежи/";
                MultiChangePath_Path.TextInput += Drawings_TbPath_OnTextInput;
                MultiChangePath_Path.PreviewTextInput += Drawings_TbPath_OnPreviewTextInput;
                if (tb != null)
                {
                    tb.TextChanged += Drawings_TbPath_OnTextChanged;
                }

                MultiChangePath_Path.ItemsSource = paths.Where(x => x.StartsWith("Чертежи/"));

                // fill items
                MultiChangePath_LvItems.ItemsSource = null;
                var lst = new List<DwgBaseItemWithSelector>();
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = _dwgBaseItems.Count;
                ProgressBar.Value = 0;
                var index = 1;
                foreach (var dwgBaseItem in _dwgBaseItems)
                {
                    if (!dwgBaseItem.IsBlock)
                    {
                        lst.Add(new DwgBaseItemWithSelector { Selected = false, Item = dwgBaseItem });
                    }

                    Dispatcher.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                        $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{_dwgBaseItems.Count}");
                    Dispatcher.Invoke(updatePbDelegate, DispatcherPriority.Background,
                        System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);
                    index++;
                }

                MultiChangePath_LvItems.ItemsSource = lst;
            }

            ClearProgress();
        }

        private void MultiChangePath_Accept_OnClick(object sender, RoutedEventArgs e)
        {
            if (MultiChangePath_Path.Text.Equals("Блоки/") | MultiChangePath_Path.Text.Equals("Чертежи/"))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u25"));
                MultiChangePath_Path.Focus();
                return;
            }

            var selectedItems = new List<DwgBaseItem>();
            foreach (DwgBaseItemWithSelector item in MultiChangePath_LvItems.Items)
            {
                if (item.Selected)
                {
                    selectedItems.Add(item.Item);
                }
            }

            if (!selectedItems.Any())
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u26"));
                return;
            }

            if (!ModPlusAPI.Windows.MessageBox.ShowYesNo(
                ModPlusAPI.Language.GetItem(LangItem, "u27"),
                MessageBoxIcon.Question))
            {
                return;
            }

            var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
            var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);
            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = _dwgBaseItems.Count;
            ProgressBar.Value = 0;
            var index = 1;
            foreach (var dwgBaseItem in _dwgBaseItems)
            {
                foreach (var selectedItem in selectedItems)
                {
                    if (dwgBaseItem.Equals(selectedItem))
                    {
                        dwgBaseItem.Path = MultiChangePath_Path.Text;
                    }
                }

                Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                    $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{_dwgBaseItems.Count}");
                Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background,
                    System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);
                index++;
            }

            // was changed
            UserBaseChanged = true;

            // save
            DwgBaseHelpers.SerializerToXml(_dwgBaseItems, _userDwgBaseFile);
            DwgBaseHelpers.DeserializeFromXml(_userDwgBaseFile, out _dwgBaseItems);
            Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, ModPlusAPI.Language.GetItem(LangItem, "u28"));
            ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u28"));

            ClearProgress();

            // refill
            MultiChangePath_FillData();
        }

        private sealed class DwgBaseItemWithSelector : VmBase
        {
            private bool _selected;

            public bool Selected
            {
                get => _selected;
                set
                {
                    _selected = value;
                    OnPropertyChanged();
                }
            }

            public DwgBaseItem Item { get; set; }
        }
        #endregion

        #region Statistic
        private void Statistic_BtGet_OnClick(object sender, RoutedEventArgs e)
        {
            Statistic_TbStat.Text = string.Empty;
            var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
            var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);
            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = _dwgBaseItems.Count;
            ProgressBar.Value = 0;

            var authors = new List<string>();
            var sourceFiles = new List<string>();
            var sourceFilesNotFound = new List<string>();
            var blks = 0;
            var drws = 0;
            var dynBlks = 0;
            var annotBlks = 0;
            var is3dblks = 0;
            var hasSpecAttr = 0;
            var index = 1;
            foreach (var dwgBaseItem in _dwgBaseItems)
            {
                if (dwgBaseItem.IsBlock)
                {
                    blks++;
                    if (dwgBaseItem.IsDynamicBlock)
                    {
                        dynBlks++;
                    }

                    if (dwgBaseItem.Is3Dblock)
                    {
                        is3dblks++;
                    }

                    if (dwgBaseItem.IsAnnotative)
                    {
                        annotBlks++;
                    }

                    if (dwgBaseItem.HasAttributesForSpecification)
                    {
                        hasSpecAttr++;
                    }
                }
                else
                {
                    drws++;
                }

                if (!authors.Contains(dwgBaseItem.Author))
                {
                    authors.Add(dwgBaseItem.Author);
                }

                var sourceFile = Path.Combine(_dwgBaseFolder, dwgBaseItem.SourceFile);
                if (File.Exists(sourceFile))
                {
                    if (!sourceFiles.Contains(dwgBaseItem.SourceFile))
                    {
                        sourceFiles.Add(dwgBaseItem.SourceFile);
                    }
                }
                else
                {
                    if (!sourceFilesNotFound.Contains(dwgBaseItem.SourceFile))
                    {
                        sourceFilesNotFound.Add(dwgBaseItem.SourceFile);
                    }
                }

                // if (index % 10 == 0 | index == _dwgBaseItems.Count - 1)
                // {
                Dispatcher.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                    $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{_dwgBaseItems.Count}");
                Dispatcher.Invoke(updatePbDelegate, DispatcherPriority.Background,
                    System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);

                // }
                index++;
            }

            // result
            var resultStr = new StringBuilder();
            resultStr.AppendLine(
                $"{ModPlusAPI.Language.GetItem(LangItem, "u29")} {_dwgBaseItems.Count} {ModPlusAPI.Language.GetItem(LangItem, "u30")}.");
            resultStr.AppendLine(ModPlusAPI.Language.GetItem(LangItem, "u31"));
            resultStr.AppendLine($"{ModPlusAPI.Language.GetItem(LangItem, "u32")} {blks}");
            resultStr.AppendLine($"{ModPlusAPI.Language.GetItem(LangItem, "u33")} {drws}");
            if (blks > 0)
            {
                resultStr.AppendLine("================================================");
                resultStr.AppendLine(
                    $"{ModPlusAPI.Language.GetItem(LangItem, "u34")} {blks} {ModPlusAPI.Language.GetItem(LangItem, "u35")}");
                if (dynBlks > 0)
                {
                    resultStr.AppendLine($"     {ModPlusAPI.Language.GetItem(LangItem, "u36")} {dynBlks}");
                }

                if (annotBlks > 0)
                {
                    resultStr.AppendLine($"     {ModPlusAPI.Language.GetItem(LangItem, "u37")} {annotBlks}");
                }

                if (is3dblks > 0)
                {
                    resultStr.AppendLine($"     {ModPlusAPI.Language.GetItem(LangItem, "u38")} {is3dblks}");
                }

                if (hasSpecAttr > 0)
                {
                    resultStr.AppendLine($"     {ModPlusAPI.Language.GetItem(LangItem, "u39")} {hasSpecAttr}");
                }
            }

            if (authors.Any())
            {
                resultStr.AppendLine("================================================");
                resultStr.AppendLine(
                    $"{ModPlusAPI.Language.GetItem(LangItem, "u40")} {authors.Count} {ModPlusAPI.Language.GetItem(LangItem, "u41")}");
                foreach (var author in authors)
                {
                    resultStr.AppendLine($"     {author}");
                }
            }

            if (sourceFiles.Any())
            {
                resultStr.AppendLine("================================================");
                resultStr.AppendLine($"{ModPlusAPI.Language.GetItem(LangItem, "u42")} {sourceFiles.Count}");
            }

            if (sourceFilesNotFound.Any())
            {
                resultStr.AppendLine("================================================");
                resultStr.AppendLine(
                    $"{ModPlusAPI.Language.GetItem(LangItem, "u43")} {sourceFilesNotFound.Count} {ModPlusAPI.Language.GetItem(LangItem, "u44")}");
            }

            Statistic_TbStat.Text = resultStr.ToString();

            ClearProgress();
        }
        #endregion

        #region UnusedFiles
        private void UnusedFiles_BtSearch_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
                var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);
                UnusedFiles_LvFiles.ItemsSource = null;

                // Load plugin base
                DwgBaseHelpers.DeserializeFromXml(_mpDwgBaseFile, out var pluginDwgBaseItems);

                // Create full one list
                var allDwgBaseItems = _dwgBaseItems.Concat(pluginDwgBaseItems).ToList();

                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = allDwgBaseItems.Count;
                ProgressBar.Value = 0;
                var index = 1;
                var unusedFiles = new List<UnusedFile>();
                var filesInItems = new List<string>();
                foreach (var baseItem in allDwgBaseItems)
                {
                    var file = Path.Combine(_dwgBaseFolder, baseItem.SourceFile.Replace("/", "\\"));
                    if (!filesInItems.Contains(file))
                    {
                        filesInItems.Add(file);
                    }

                    Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                        $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{allDwgBaseItems.Count}");
                    Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background,
                        System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);
                    index++;
                }

                Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, 0.0);
                var filesInBaseDirectory = Directory.GetFiles(_dwgBaseFolder, "*.dwg", SearchOption.AllDirectories).ToList();
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = filesInBaseDirectory.Count;
                ProgressBar.Value = 0;
                index = 1;
                foreach (var fileInDir in filesInBaseDirectory)
                {
                    if (!filesInItems.Contains(fileInDir))
                    {
                        var fi = new FileInfo(fileInDir);
                        unusedFiles.Add(new UnusedFile
                        {
                            Selected = false,
                            FileName = fi.Name,
                            FullFileName = fi.FullName
                        });
                    }

                    Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                        $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{filesInBaseDirectory.Count}");
                    Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);
                    index++;
                }

                if (unusedFiles.Any())
                {
                    UnusedFiles_BtDelete.Visibility = Visibility.Visible;
                    UnusedFiles_LvFiles.Visibility = Visibility.Visible;
                    UnusedFiles_TbDelInfo.Visibility = Visibility.Visible;
                    UnusedFiles_LvFiles.ItemsSource = unusedFiles;
                }
                else
                {
                    ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u45"));
                }

                ClearProgress();
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                ClearProgress();
            }
        }

        // deleting
        private void UnusedFiles_BtDelete_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedUnusedFiles = new List<UnusedFile>();
                foreach (UnusedFile file in UnusedFiles_LvFiles.Items)
                {
                    if (file.Selected)
                    {
                        selectedUnusedFiles.Add(file);
                    }
                }

                if (selectedUnusedFiles.Any())
                {
                    var count = 0;
                    var filesNotDeleted = new List<UnusedFile>();
                    if (ModPlusAPI.Windows.MessageBox.ShowYesNo(
                        ModPlusAPI.Language.GetItem(LangItem, "u46"),
                        MessageBoxIcon.Question))
                    {
                        foreach (var file in selectedUnusedFiles)
                        {
                            if (File.Exists(file.FullFileName))
                            {
                                try
                                {
                                    File.Delete(file.FullFileName);

                                    // delete bak's and images
                                    var fi = new FileInfo(file.FullFileName);
                                    var images = $"{file.FullFileName} icons";
                                    if (Directory.Exists(images))
                                    {
                                        Directory.Delete(images, true);
                                    }

                                    if (fi.DirectoryName != null)
                                    {
                                        var bak = Path.Combine(fi.DirectoryName,
                                            $"{Path.GetFileNameWithoutExtension(file.FullFileName)}.bak");
                                        try
                                        {
                                            File.Delete(bak);
                                        }
                                        catch
                                        {
                                            // ignored
                                        }
                                    }

                                    count++;
                                }
                                catch (IOException)
                                {
                                    filesNotDeleted.Add(file);
                                }
                            }
                        }

                        ModPlusAPI.Windows.MessageBox.Show($"{ModPlusAPI.Language.GetItem(LangItem, "u47")}: {count}");
                    }

                    if (filesNotDeleted.Any())
                    {
                        var filesStr = string.Empty;
                        foreach (var file in filesNotDeleted)
                        {
                            filesStr += file.FullFileName + Environment.NewLine;
                        }

                        ModPlusAPI.Windows.MessageBox.Show(
                            ModPlusAPI.Language.GetItem(LangItem, "u48") + Environment.NewLine + filesStr +
                            ModPlusAPI.Language.GetItem(LangItem, "u49"));
                    }

                    // visibility
                    UnusedFiles_BtDelete.Visibility = Visibility.Collapsed;
                    UnusedFiles_LvFiles.Visibility = Visibility.Collapsed;
                    UnusedFiles_TbDelInfo.Visibility = Visibility.Collapsed;
                    UnusedFiles_LvFiles.ItemsSource = null;
                }
                else
                {
                    ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u26"));
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }

        private sealed class UnusedFile : VmBase
        {
            private bool _selected;

            public bool Selected
            {
                get => _selected;
                set
                {
                    _selected = value;
                    OnPropertyChanged();
                }
            }

            public string FileName { get; set; }

            public string FullFileName { get; set; }
        }
        #endregion

        #region UnusedItems
        private void UnusedItems_BtSearch_OnClick(object sender, RoutedEventArgs e)
        {
            var unusedItems = new List<DwgBaseItemWithSelector>();
            try
            {
                var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
                var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);
                UnusedItems_LvFiles.ItemsSource = null;

                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = _dwgBaseItems.Count;
                ProgressBar.Value = 0;
                var index = 1;

                foreach (var baseItem in _dwgBaseItems)
                {
                    var file = Path.Combine(_dwgBaseFolder, baseItem.SourceFile.Replace("/", "\\"));
                    if (!File.Exists(file))
                    {
                        unusedItems.Add(new DwgBaseItemWithSelector
                        {
                            Selected = false,
                            Item = baseItem
                        });
                    }

                    Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                        $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{_dwgBaseItems.Count}");
                    Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background,
                        System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);
                    index++;
                }

                if (unusedItems.Any())
                {
                    UnusedItems_BtDelete.Visibility = Visibility.Visible;
                    UnusedItems_LvFiles.Visibility = Visibility.Visible;
                    UnusedItems_LvFiles.ItemsSource = unusedItems;
                }
                else
                {
                    ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u50"));
                }

                ClearProgress();
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                ClearProgress();
            }
        }

        private void UnusedItems_BtDelete_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedUnusedItems = new List<DwgBaseItem>();
                foreach (DwgBaseItemWithSelector item in UnusedItems_LvFiles.Items)
                {
                    if (item.Selected)
                    {
                        selectedUnusedItems.Add(item.Item);
                    }
                }

                if (selectedUnusedItems.Any())
                {
                    // var count = 0;
                    if (ModPlusAPI.Windows.MessageBox.ShowYesNo(
                        ModPlusAPI.Language.GetItem(LangItem, "u51"),
                        MessageBoxIcon.Question))
                    {
                        var removed = _dwgBaseItems.RemoveAll(x => selectedUnusedItems.Contains(x));

                        // resave
                        DwgBaseHelpers.SerializerToXml(_dwgBaseItems, _userDwgBaseFile);
                        DwgBaseHelpers.DeserializeFromXml(_userDwgBaseFile, out _dwgBaseItems);
                        ModPlusAPI.Windows.MessageBox.Show($"{ModPlusAPI.Language.GetItem(LangItem, "u52")} {removed}");
                    }

                    // visibility
                    UnusedItems_BtDelete.Visibility = Visibility.Collapsed;
                    UnusedItems_LvFiles.Visibility = Visibility.Collapsed;
                    UnusedItems_LvFiles.ItemsSource = null;
                }
                else
                {
                    ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u26"));
                }
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
            }
        }
        #endregion

        #region MultiChangeSourceAuthor

        private void MultiChangeSourceAuthor_FillData()
        {
            var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
            var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);

            // fill items
            MultiChangeSourceAuthor_LvItems.ItemsSource = null;
            var lst = new List<DwgBaseItemWithSelector>();
            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = _dwgBaseItems.Count;
            ProgressBar.Value = 0;
            var index = 1;
            foreach (var dwgBaseItem in _dwgBaseItems)
            {
                lst.Add(new DwgBaseItemWithSelector { Selected = false, Item = dwgBaseItem });
                Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                    $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{_dwgBaseItems.Count}");
                Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background,
                    System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);
                index++;
            }

            MultiChangeSourceAuthor_LvItems.ItemsSource = lst;

            ClearProgress();
        }

        // on accept
        private void MultiChangeSourceAuthor_BtAccept_OnClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(MultiChangeSourceAuthor_TbSource.Text) &
                string.IsNullOrEmpty(MultiChangeSourceAuthor_TbAuthor.Text))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u53"));
                return;
            }

            var selectedItems = new List<DwgBaseItem>();
            foreach (DwgBaseItemWithSelector item in MultiChangeSourceAuthor_LvItems.Items)
            {
                if (item.Selected)
                {
                    selectedItems.Add(item.Item);
                }
            }

            if (!selectedItems.Any())
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u26"));
                return;
            }

            if (!ModPlusAPI.Windows.MessageBox.ShowYesNo(
                ModPlusAPI.Language.GetItem(LangItem, "u27"),
                MessageBoxIcon.Question))
            {
                return;
            }

            var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
            var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);
            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = _dwgBaseItems.Count;
            ProgressBar.Value = 0;
            var index = 1;
            foreach (var dwgBaseItem in _dwgBaseItems)
            {
                foreach (var selectedItem in selectedItems)
                {
                    if (dwgBaseItem.Equals(selectedItem))
                    {
                        if (!string.IsNullOrEmpty(MultiChangeSourceAuthor_TbSource.Text))
                        {
                            dwgBaseItem.Source = MultiChangeSourceAuthor_TbSource.Text;
                        }

                        if (!string.IsNullOrEmpty(MultiChangeSourceAuthor_TbAuthor.Text))
                        {
                            dwgBaseItem.Author = MultiChangeSourceAuthor_TbAuthor.Text;
                        }
                    }
                }

                Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                    $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{_dwgBaseItems.Count}");
                Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background, System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);
                index++;
            }

            // was changed
            UserBaseChanged = true;

            // save
            DwgBaseHelpers.SerializerToXml(_dwgBaseItems, _userDwgBaseFile);
            DwgBaseHelpers.DeserializeFromXml(_userDwgBaseFile, out _dwgBaseItems);
            Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, ModPlusAPI.Language.GetItem(LangItem, "u28"));
            ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u28"));

            ClearProgress();

            // refill
            MultiChangeSourceAuthor_FillData();
        }
        #endregion

        #region Rename source file

        private void RenameFiles_FindAllFiles()
        {
            try
            {
                var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
                var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);
                RenameSourceFile_LbFiles.ItemsSource = null;

                // Load plugin base
                DwgBaseHelpers.DeserializeFromXml(_mpDwgBaseFile, out _);

                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = _dwgBaseItems.Count;
                ProgressBar.Value = 0;
                var index = 1;
                var files = new List<DbFile>();
                var filesInItems = new List<string>();
                foreach (var baseItem in _dwgBaseItems)
                {
                    var file = Path.Combine(_dwgBaseFolder, baseItem.SourceFile.Replace("/", "\\"));
                    if (!filesInItems.Contains(file) && File.Exists(file))
                    {
                        filesInItems.Add(file);
                    }

                    Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                        $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{_dwgBaseItems.Count}");
                    Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background,
                        System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);
                    index++;
                }

                foreach (var filesInItem in filesInItems)
                {
                    var fi = new FileInfo(filesInItem);
                    files.Add(new DbFile()
                    {
                        FileName = fi.Name,
                        FullFileName = fi.FullName
                    });
                }

                if (files.Any())
                {
                    RenameSourceFile_LbFiles.ItemsSource = files;
                }

                ClearProgress();
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                ClearProgress();
            }
        }

        private void RenameSourceFile_BtRename_OnClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(RenameSourceFile_NewFileName.Text))
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u58"));
                return;
            }

            if (RenameSourceFile_LbFiles.SelectedIndex == -1)
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u26"));
                return;
            }

            var symbolsLst = new[] { "\\", "/", ":", "*", "?", "\"", "<", ">", "|" };
            var symbols = symbolsLst.Aggregate(string.Empty, (current, s) => current + (s + " "));
            foreach (var s in symbolsLst)
            {
                if (RenameSourceFile_NewFileName.Text.Contains(s))
                {
                    ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u59") +
                        Environment.NewLine + symbols);
                }
            }

            try
            {
                var selectedFile = (DbFile)RenameSourceFile_LbFiles.SelectedItem;
                var fi = new FileInfo(selectedFile.FullFileName);
                var newFullFileName = Path.Combine(fi.DirectoryName, $"{RenameSourceFile_NewFileName.Text}.dwg");

                File.Move(selectedFile.FullFileName, newFullFileName);

                foreach (var file in Directory.GetFiles(DwgBaseHelpers.GetCustomBaseFolder(), "*.dwg", SearchOption.AllDirectories))
                {
                    var efi = new FileInfo(file);
                    if (efi.FullName.Equals(newFullFileName))
                    {
                        ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u61") + newFullFileName);
                        return;
                    }
                }

                var updatePbDelegate = new UpdateProgressBarDelegate(ProgressBar.SetValue);
                var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = _dwgBaseItems.Count;
                ProgressBar.Value = 0;
                var index = 1;
                foreach (var dwgBaseItem in _dwgBaseItems)
                {
                    if (Path.Combine(DwgBaseHelpers.GetCustomBaseFolder(), dwgBaseItem.SourceFile).Equals(selectedFile.FullFileName))
                    {
                        dwgBaseItem.SourceFile = newFullFileName.TrimStart(DwgBaseHelpers.GetCustomBaseFolder().ToCharArray());
                    }

                    Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                        $"{ModPlusAPI.Language.GetItem(LangItem, "u24")}: {index}/{_dwgBaseItems.Count}");
                    Dispatcher?.Invoke(updatePbDelegate, DispatcherPriority.Background,
                        System.Windows.Controls.Primitives.RangeBase.ValueProperty, (double)index);
                    index++;
                }

                // was changed
                UserBaseChanged = true;

                // save
                DwgBaseHelpers.SerializerToXml(_dwgBaseItems, _userDwgBaseFile);
                DwgBaseHelpers.DeserializeFromXml(_userDwgBaseFile, out _dwgBaseItems);
                Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty,
                    ModPlusAPI.Language.GetItem(LangItem, "u28"));
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u28"));

                ClearProgress();

                // refill
                RenameFiles_FindAllFiles();
            }
            catch (UnauthorizedAccessException)
            {
                ModPlusAPI.Windows.MessageBox.Show(ModPlusAPI.Language.GetItem(LangItem, "u60"));
            }
            catch (Exception exception)
            {
                ExceptionBox.Show(exception);
                ClearProgress();
            }
        }

        internal class DbFile
        {
            public string FileName { get; set; }

            public string FullFileName { get; set; }
        }
        #endregion

        private void ClearProgress()
        {
            var updatePtDelegate = new UpdateProgressTextDelegate(ProgressText.SetValue);
            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = 1;
            ProgressBar.Value = 0;
            Dispatcher?.Invoke(updatePtDelegate, DispatcherPriority.Background, TextBlock.TextProperty, string.Empty);
        }
    }
}
