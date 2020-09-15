namespace mpDwgBase.Models
{
    using System.Collections.Generic;
    using ModPlusAPI.Mvvm;

    /// <summary>
    /// Элемент дерева
    /// </summary>
    public class TreeViewModelItem : VmBase
    {
        private string _name;
        private TreeViewModelItem _parent;

        /// <summary>
        /// Базовый конструктор
        /// </summary>
        public TreeViewModelItem()
        {
            Items = new List<TreeViewModelItem>();
        }

        /// <summary>
        /// Элементы
        /// </summary>
        public List<TreeViewModelItem> Items { get; set; }

        /// <summary>
        /// Имя
        /// </summary>
        public string Name
        {
            get => _name;
            set
            {
                if (_name == value)
                {
                    return;
                }

                _name = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Родительский элемент дерева
        /// </summary>
        public TreeViewModelItem Parent
        {
            get => _parent;
            set
            {
                if (_parent == value)
                {
                    return;
                }

                _parent = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Дети
        /// </summary>
        public List<string> Children { get; set; }

        /// <summary>
        /// Получить родословную
        /// </summary>
        public string GetAncestry()
        {
            var x = string.Empty;
            if (Parent != null)
            {
                x += $"{Parent.GetAncestry()}/";
            }

            x += Name;
            return x;
        }
    }
}