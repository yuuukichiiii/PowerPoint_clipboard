using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ShapePalette.Engine;
using ShapePalette.Storage;

namespace ShapePalette.UI
{
    public partial class PaletteView : UserControl
    {
        private ShapeStore _store;
        private readonly ObservableCollection<TileVM> _tiles = new ObservableCollection<TileVM>();

        // タイル幅（px）。スライダーとタイルの幅・サムネ高さがこれにバインドする。
        public static readonly DependencyProperty TileWidthProperty =
            DependencyProperty.Register("TileWidth", typeof(double), typeof(PaletteView),
                new PropertyMetadata(92.0));

        public double TileWidth
        {
            get { return (double)GetValue(TileWidthProperty); }
            set { SetValue(TileWidthProperty, value); }
        }

        public PaletteView()
        {
            InitializeComponent();
            TileItems.ItemsSource = _tiles;
        }

        public void Initialize(ShapeStore store)
        {
            _store = store;
            TileWidth = store.Data.Zoom;
            BuildTabs();
            LoadTiles();
            SetStatus("");
        }

        private void ZoomSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_store == null) return;
            _store.Data.Zoom = e.NewValue;
            _store.Data.Save();
        }

        // ===== タブ =====
        private void BuildTabs()
        {
            TabStrip.Children.Clear();
            if (_store == null) return;
            string active = _store.Data.ActiveTabObj()?.Id;
            foreach (var tab in _store.Data.Tabs)
            {
                var tb = new ToggleButton
                {
                    Content = tab.Name,
                    Tag = tab.Id,
                    Style = (Style)Resources["TabPill"],
                    IsChecked = tab.Id == active
                };
                tb.Click += Tab_Click;
                tb.MouseRightButtonUp += Tab_RightClick;
                TabStrip.Children.Add(tb);
            }
        }

        private void Tab_Click(object sender, RoutedEventArgs e)
        {
            var id = (string)((ToggleButton)sender).Tag;
            _store.Data.ActiveTab = id;
            _store.Data.Save();
            BuildTabs();
            LoadTiles();
        }

        private void Tab_RightClick(object sender, MouseButtonEventArgs e)
        {
            var id = (string)((ToggleButton)sender).Tag;
            var tab = _store.Data.Tabs.Find(t => t.Id == id);
            if (tab == null) return;
            var menu = new ContextMenu();
            var rename = new MenuItem { Header = "タブ名を変更" };
            rename.Click += (s, a) =>
            {
                string n = Prompt("タブ名", tab.Name);
                if (!string.IsNullOrWhiteSpace(n)) { tab.Name = n.Trim(); _store.Data.Save(); BuildTabs(); }
            };
            var del = new MenuItem { Header = "タブを削除" };
            del.Click += (s, a) =>
            {
                if (_store.Data.Tabs.Count <= 1) { SetStatus("最後のタブは削除できません"); return; }
                if (MessageBox.Show("「" + tab.Name + "」を削除しますか？\n中の図形も削除されます。", "確認",
                        MessageBoxButton.OKCancel, MessageBoxImage.Warning) != MessageBoxResult.OK) return;
                foreach (var it in tab.Items.ToArray()) _store.DeleteItem(it);
                _store.Data.Tabs.Remove(tab);
                if (_store.Data.ActiveTab == tab.Id) _store.Data.ActiveTab = _store.Data.Tabs[0].Id;
                _store.Data.Save();
                BuildTabs(); LoadTiles();
            };
            menu.Items.Add(rename);
            menu.Items.Add(del);
            menu.IsOpen = true;
            e.Handled = true;
        }

        private void BtnAddTab_Click(object sender, RoutedEventArgs e)
        {
            string name = Prompt("新しいタブ名", "");
            if (string.IsNullOrWhiteSpace(name)) return;
            var tab = new PaletteTab { Name = name.Trim() };
            _store.Data.Tabs.Add(tab);
            _store.Data.ActiveTab = tab.Id;
            _store.Data.Save();
            BuildTabs(); LoadTiles();
            // 追加したタブが見えるよう右端へスクロール
            TabScroll.ScrollToRightEnd();
        }

        // マウスホイールでタブを左右スクロール
        private void TabScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            TabScroll.ScrollToHorizontalOffset(TabScroll.HorizontalOffset - e.Delta);
            e.Handled = true;
        }

        // ===== タイル =====
        private void LoadTiles()
        {
            _tiles.Clear();
            var tab = _store?.Data.ActiveTabObj();
            if (tab == null) return;
            foreach (var it in tab.Items)
                _tiles.Add(new TileVM(it, LoadThumb(_store.ThumbPathOf(it))));
        }

        private static ImageSource LoadThumb(string path)
        {
            if (string.IsNullOrEmpty(path)) return null;
            try
            {
                var bi = new BitmapImage();
                bi.BeginInit();
                bi.CacheOption = BitmapCacheOption.OnLoad;       // ファイルをロックしない
                bi.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                bi.UriSource = new Uri(path);
                bi.EndInit();
                bi.Freeze();
                return bi;
            }
            catch { return null; }
        }

        private void BtnCapture_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var item = _store.CaptureSelection();
                LoadTiles();
                SetStatus("取り込みました: " + item.Name);
            }
            catch (Exception ex) { SetStatus(ex.Message); }
        }

        private void Tile_Click(object sender, RoutedEventArgs e)
        {
            var vm = ((FrameworkElement)sender).Tag as TileVM;
            InsertTile(vm);
        }

        private void InsertTile(TileVM vm)
        {
            if (vm == null) return;
            try { _store.InsertItem(vm.Item); SetStatus("挿入しました: " + vm.Name); }
            catch (Exception ex) { SetStatus(ex.Message); }
        }

        private TileVM CtxTarget(object sender)
        {
            var mi = sender as MenuItem;
            var cm = mi?.Parent as ContextMenu;
            var btn = cm?.PlacementTarget as FrameworkElement;
            return btn?.Tag as TileVM;
        }

        private void Ctx_Insert(object sender, RoutedEventArgs e) => InsertTile(CtxTarget(sender));

        private void Ctx_Rename(object sender, RoutedEventArgs e)
        {
            var vm = CtxTarget(sender);
            if (vm == null) return;
            string n = Prompt("名前を変更", vm.Name);
            if (!string.IsNullOrWhiteSpace(n)) { vm.Item.Name = n.Trim(); _store.Data.Save(); LoadTiles(); }
        }

        private void Ctx_Delete(object sender, RoutedEventArgs e)
        {
            var vm = CtxTarget(sender);
            if (vm == null) return;
            _store.DeleteItem(vm.Item);
            LoadTiles();
            SetStatus("削除しました");
        }

        // ===== ユーティリティ =====
        private void SetStatus(string text) { Status.Text = text; }

        private static string Prompt(string caption, string def)
        {
            return Microsoft.VisualBasic.Interaction.InputBox(caption, "Shape Palette", def);
        }
    }

    /// <summary>値 × パラメータ（比率）を返すコンバータ。サムネ高さ＝タイル幅×比率 に使う。</summary>
    public class RatioConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            double v = value is double ? (double)value : 0;
            double r = 0.62;
            if (parameter != null) double.TryParse(parameter.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out r);
            return v * r;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }

    /// <summary>タイル1個分の表示モデル。</summary>
    public class TileVM
    {
        public PaletteItem Item { get; }
        public string Name => Item.Name;
        public ImageSource Image { get; }

        public TileVM(PaletteItem item, ImageSource image)
        {
            Item = item;
            Image = image;
        }
    }
}
