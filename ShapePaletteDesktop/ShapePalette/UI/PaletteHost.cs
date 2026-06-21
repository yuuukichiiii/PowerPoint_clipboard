using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using ShapePalette.Engine;
using ShapePalette.Storage;

namespace ShapePalette.UI
{
    /// <summary>
    /// カスタムタスクペインに載せるコンテナ。
    /// Phase 2 では保存エンジン検証用の簡易 UI（取り込み/挿入/削除）。Phase 3 で WPF に置換する。
    /// </summary>
    [ComVisible(true)]
    [Guid("7F4C3A87-FB96-43D8-8B52-B9D92EF2A81B")]
    [ProgId("ShapePalette.PaletteHost")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class PaletteHost : UserControl
    {
        private PowerPoint.Application _app;
        private ShapeStore _store;

        private readonly Button _btnCapture = new Button();
        private readonly Button _btnInsert = new Button();
        private readonly Button _btnDelete = new Button();
        private readonly ListBox _list = new ListBox();
        private readonly PictureBox _thumb = new PictureBox();
        private readonly Label _status = new Label();

        public PaletteHost()
        {
            BackColor = Color.FromArgb(244, 245, 248);
            Padding = new Padding(10);

            _btnCapture.Text = "選択した図形を取り込む";
            _btnCapture.Dock = DockStyle.Top;
            _btnCapture.Height = 40;
            _btnCapture.Click += (s, e) => Safe(DoCapture);

            var row = new Panel { Dock = DockStyle.Top, Height = 34 };
            _btnInsert.Text = "挿入";
            _btnInsert.Width = 150; _btnInsert.Dock = DockStyle.Left; _btnInsert.Height = 32;
            _btnInsert.Click += (s, e) => Safe(DoInsert);
            _btnDelete.Text = "削除";
            _btnDelete.Width = 150; _btnDelete.Dock = DockStyle.Right; _btnDelete.Height = 32;
            _btnDelete.Click += (s, e) => Safe(DoDelete);
            row.Controls.Add(_btnInsert);
            row.Controls.Add(_btnDelete);

            _list.Dock = DockStyle.Fill;
            _list.IntegralHeight = false;
            _list.SelectedIndexChanged += (s, e) => ShowThumb();
            _list.DoubleClick += (s, e) => Safe(DoInsert);

            _thumb.Dock = DockStyle.Bottom;
            _thumb.Height = 140;
            _thumb.SizeMode = PictureBoxSizeMode.Zoom;
            _thumb.BackColor = Color.White;

            _status.Dock = DockStyle.Bottom;
            _status.Height = 22;
            _status.ForeColor = Color.FromArgb(118, 125, 138);
            _status.Text = "（接続待ち）";

            // 追加順に注意（Dock の積み重なり）
            Controls.Add(_list);
            Controls.Add(row);
            Controls.Add(_btnCapture);
            Controls.Add(_thumb);
            Controls.Add(_status);
        }

        public void Initialize(PowerPoint.Application app)
        {
            _app = app;
            try
            {
                _store = new ShapeStore(app);
                RefreshList();
                _status.Text = "準備OK（Phase 2 検証）";
            }
            catch (Exception ex)
            {
                _status.Text = "init 失敗: " + ex.Message;
                Log.Write("PaletteHost.Initialize EX: " + ex);
            }
        }

        private PaletteTab Tab => _store?.Data.ActiveTabObj();

        private void RefreshList()
        {
            _list.Items.Clear();
            var tab = Tab;
            if (tab != null)
                foreach (var it in tab.Items) _list.Items.Add(it.Name);
            ShowThumb();
        }

        private PaletteItem SelectedItem()
        {
            var tab = Tab;
            int i = _list.SelectedIndex;
            if (tab == null || i < 0 || i >= tab.Items.Count) return null;
            return tab.Items[i];
        }

        private void ShowThumb()
        {
            if (_thumb.Image != null) { _thumb.Image.Dispose(); _thumb.Image = null; }
            var item = SelectedItem();
            if (item == null) return;
            var path = _store.ThumbPathOf(item);
            if (path != null)
            {
                try { using (var fs = System.IO.File.OpenRead(path)) _thumb.Image = Image.FromStream(fs); }
                catch { }
            }
        }

        private void DoCapture()
        {
            var item = _store.CaptureSelection();
            RefreshList();
            _list.SelectedIndex = _list.Items.Count - 1;
            _status.Text = "取り込みました: " + item.Name;
        }

        private void DoInsert()
        {
            var item = SelectedItem();
            if (item == null) { _status.Text = "挿入する図形を選んでください"; return; }
            _store.InsertItem(item);
            _status.Text = "挿入しました: " + item.Name;
        }

        private void DoDelete()
        {
            var item = SelectedItem();
            if (item == null) return;
            _store.DeleteItem(item);
            RefreshList();
            _status.Text = "削除しました";
        }

        private void Safe(Action a)
        {
            try { a(); }
            catch (Exception ex)
            {
                _status.Text = "エラー: " + ex.Message;
                Log.Write("UI action EX: " + ex);
            }
        }
    }
}
