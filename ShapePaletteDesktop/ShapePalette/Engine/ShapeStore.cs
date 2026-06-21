using System;
using System.IO;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using ShapePalette.Storage;

namespace ShapePalette.Engine
{
    /// <summary>
    /// 図形の保存エンジン。選択図形を PowerPoint 自身のコピー＆ペーストで
    /// バックグラウンドのライブラリ pptx に退避し、サムネ PNG を生成する。
    /// 挿入はライブラリからコピペで完全再現する（線・グループ・グラデ等もそのまま）。
    /// </summary>
    public class ShapeStore
    {
        private readonly PowerPoint.Application _app;
        private PowerPoint.Presentation _lib;

        public PaletteData Data { get; private set; }

        public ShapeStore(PowerPoint.Application app)
        {
            _app = app;
            Data = PaletteData.Load();
        }

        #region ライブラリ pptx

        private PowerPoint.Presentation EnsureLibrary()
        {
            if (_lib != null)
            {
                try { var _ = _lib.Name; return _lib; }   // 生存確認
                catch { _lib = null; }
            }

            string path = PaletteData.LibraryPath;
            if (File.Exists(path))
            {
                _lib = _app.Presentations.Open(path,
                    MsoTriState.msoFalse,   // ReadOnly
                    MsoTriState.msoFalse,   // Untitled
                    MsoTriState.msoFalse);  // WithWindow（非表示で開く）
            }
            else
            {
                _lib = _app.Presentations.Add(MsoTriState.msoFalse);
                _lib.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            }
            Log.Write("Library ready: " + path);
            return _lib;
        }

        public void CloseLibrary()
        {
            try { if (_lib != null) { _lib.Save(); _lib.Close(); } }
            catch { }
            _lib = null;
        }

        #endregion

        /// <summary>選択中の図形を取り込む。取り込んだ PaletteItem を返す。</summary>
        public PaletteItem CaptureSelection(string tabId = null)
        {
            var win = _app.ActiveWindow;
            if (win == null) throw new InvalidOperationException("ウィンドウがありません。");

            var sel = win.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                throw new InvalidOperationException("スライドで図形を選択してください。");

            var srcRange = sel.ShapeRange;
            int count = srcRange.Count;
            srcRange.Copy();

            var lib = EnsureLibrary();
            var slide = lib.Slides.Add(lib.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);

            var pasted = slide.Shapes.Paste();
            PowerPoint.Shape thumbShape = (pasted.Count > 1) ? pasted.Group() : pasted[1];

            // サムネ PNG を書き出し
            string thumbFile = Guid.NewGuid().ToString("N") + ".png";
            string thumbPath = Path.Combine(PaletteData.ThumbsDir, thumbFile);
            try { thumbShape.Export(thumbPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 0, 0, PowerPoint.PpExportMode.ppRelativeToSlide); }
            catch (Exception ex) { Log.Write("Export thumb EX: " + ex.Message); thumbFile = null; }

            int slideId = slide.SlideID;
            lib.Save();

            var tab = tabId != null ? Data.Tabs.Find(t => t.Id == tabId) : Data.ActiveTabObj();
            if (tab == null) tab = Data.ActiveTabObj();

            var item = new PaletteItem
            {
                Name = "図形 " + (tab.Items.Count + 1),
                SlideId = slideId,
                ThumbFile = thumbFile
            };
            tab.Items.Add(item);
            Data.Save();
            Log.Write("Captured: id=" + item.Id + " slideId=" + slideId + " count=" + count);
            return item;
        }

        /// <summary>アイテムの図形をクリップボードへコピーする（ドラッグ&ドロップ挿入の元）。</summary>
        public void CopyItemToClipboard(PaletteItem item)
        {
            var lib = EnsureLibrary();
            PowerPoint.Slide libSlide;
            try { libSlide = lib.Slides.FindBySlideID(item.SlideId); }
            catch { throw new InvalidOperationException("ライブラリに該当図形が見つかりません。"); }

            if (libSlide.Shapes.Count == 0) throw new InvalidOperationException("図形が空です。");
            libSlide.Shapes.Range().Copy();
        }

        /// <summary>アイテムを現在のスライドの中央に挿入する（完全再現）。</summary>
        public void InsertItem(PaletteItem item)
        {
            CopyItemToClipboard(item);

            var win = _app.ActiveWindow;
            if (win == null) throw new InvalidOperationException("挿入先のウィンドウがありません。");
            var activeSlide = win.View.Slide as PowerPoint.Slide;
            if (activeSlide == null) throw new InvalidOperationException("挿入先のスライドがありません。");

            var pasted = activeSlide.Shapes.Paste();
            try
            {
                // スライド中央へ配置（表示領域の中心座標は API で取得できないため）
                var ps = _app.ActivePresentation.PageSetup;
                float sw = ps.SlideWidth, sh = ps.SlideHeight;
                if (pasted.Count >= 1)
                {
                    var shp = pasted[1];
                    shp.Left = (sw - shp.Width) / 2f;
                    shp.Top = (sh - shp.Height) / 2f;
                }
                pasted.Select();
            }
            catch (Exception ex) { Log.Write("InsertItem center EX: " + ex.Message); }
            Log.Write("Inserted: id=" + item.Id);
        }

        /// <summary>
        /// ドラッグ&ドロップ後に、落とした画面座標がスライド内なら自前で図形を挿入する。
        /// ドラッグは PowerPoint が認識しない形式で行うため、PowerPoint 側はドロップを処理しない
        /// （プレースホルダがテキスト編集に入らない）。挿入は完全にこちらで制御する。
        /// </summary>
        public void DropInsert(PaletteItem item, double screenX, double screenY)
        {
            try
            {
                var win = _app.ActiveWindow;
                if (win == null) return;
                var slide = win.View.Slide as PowerPoint.Slide;
                if (slide == null) return;

                if (!TryScreenToSlide(win, screenX, screenY, out double slideX, out double slideY)) return;

                var ps = _app.ActivePresentation.PageSetup;
                bool onSlide = slideX >= 0 && slideY >= 0 && slideX <= ps.SlideWidth && slideY <= ps.SlideHeight;
                if (!onSlide) return;   // スライドの外に落とした場合は何もしない

                try { win.Selection.Unselect(); } catch { }   // 念のためテキスト編集状態を解除
                // クリップボードへはドラッグ開始時にコピー済み（ここでの再コピーは省いて高速化）
                var pasted = slide.Shapes.Paste();
                if (pasted.Count >= 1)
                {
                    var shp = pasted[1];
                    shp.Left = (float)(slideX - shp.Width / 2.0);
                    shp.Top = (float)(slideY - shp.Height / 2.0);
                }
                try { pasted.Select(); } catch { }
            }
            catch (Exception ex) { Log.Write("DropInsert EX: " + ex.Message); }
        }

        // 画面ピクセル -> スライド座標(pt) への逆変換（2点サンプル）
        private bool TryScreenToSlide(PowerPoint.DocumentWindow win, double screenX, double screenY,
                                      out double slideX, out double slideY)
        {
            slideX = slideY = 0;
            double x0 = win.PointsToScreenPixelsX(0), x100 = win.PointsToScreenPixelsX(100);
            double y0 = win.PointsToScreenPixelsY(0), y100 = win.PointsToScreenPixelsY(100);
            double sx = (x100 - x0) / 100.0, sy = (y100 - y0) / 100.0;
            if (sx == 0 || sy == 0) return false;
            slideX = (screenX - x0) / sx;
            slideY = (screenY - y0) / sy;
            return true;
        }

        public string ThumbPathOf(PaletteItem item)
        {
            if (string.IsNullOrEmpty(item.ThumbFile)) return null;
            var p = Path.Combine(PaletteData.ThumbsDir, item.ThumbFile);
            return File.Exists(p) ? p : null;
        }

        public void DeleteItem(PaletteItem item)
        {
            foreach (var tab in Data.Tabs)
            {
                if (tab.Items.Remove(item))
                {
                    try
                    {
                        var lib = EnsureLibrary();
                        var s = lib.Slides.FindBySlideID(item.SlideId);
                        s.Delete();
                        lib.Save();
                    }
                    catch (Exception ex) { Log.Write("DeleteItem lib EX: " + ex.Message); }
                    var tp = ThumbPathOf(item);
                    if (tp != null) { try { File.Delete(tp); } catch { } }
                    Data.Save();
                    break;
                }
            }
        }
    }
}
