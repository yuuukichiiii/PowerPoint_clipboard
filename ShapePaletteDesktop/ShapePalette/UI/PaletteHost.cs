using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using ShapePalette.Engine;

namespace ShapePalette.UI
{
    /// <summary>
    /// カスタムタスクペインに載せる WinForms コンテナ。
    /// 中身は ElementHost 経由で WPF の PaletteView を表示する。
    /// </summary>
    [ComVisible(true)]
    [Guid("7F4C3A87-FB96-43D8-8B52-B9D92EF2A81B")]
    [ProgId("ShapePalette.PaletteHost")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class PaletteHost : UserControl
    {
        private ShapeStore _store;
        private ElementHost _host;
        private PaletteView _view;

        public PaletteHost()
        {
            BackColor = System.Drawing.Color.FromArgb(247, 248, 250);
        }

        private void PositionHost()
        {
            if (_host == null) return;
            _host.Bounds = new System.Drawing.Rectangle(0, -1, ClientSize.Width, ClientSize.Height + 1);
        }

        public void Initialize(PowerPoint.Application app)
        {
            try
            {
                _store = new ShapeStore(app);
                _view = new PaletteView();
                _view.Initialize(_store);

                _host = new ElementHost
                {
                    Child = _view,
                    BackColor = BackColor   // 未描画部分が黒線にならないよう地色に合わせる
                };
                Controls.Clear();
                Controls.Add(_host);
                // ElementHost 上端に出る 1px の継ぎ目を見出しバー下へ逃がす（親がクリップ）
                PositionHost();
                Resize += (s, e) => PositionHost();
            }
            catch (Exception ex)
            {
                Log.Write("PaletteHost.Initialize EX: " + ex);
                Controls.Clear();
                Controls.Add(new Label
                {
                    Dock = DockStyle.Fill,
                    Text = "初期化に失敗しました:\r\n" + ex.Message,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                });
            }
        }
    }
}
