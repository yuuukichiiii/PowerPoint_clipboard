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

        /// <summary>アイテムを現在のスライドに挿入する（完全再現）。</summary>
        public void InsertItem(PaletteItem item)
        {
            var lib = EnsureLibrary();
            PowerPoint.Slide libSlide;
            try { libSlide = lib.Slides.FindBySlideID(item.SlideId); }
            catch { throw new InvalidOperationException("ライブラリに該当図形が見つかりません。"); }

            if (libSlide.Shapes.Count == 0) throw new InvalidOperationException("図形が空です。");

            libSlide.Shapes.Range().Copy();

            var win = _app.ActiveWindow;
            if (win == null) throw new InvalidOperationException("挿入先のウィンドウがありません。");
            var activeSlide = win.View.Slide as PowerPoint.Slide;
            if (activeSlide == null) throw new InvalidOperationException("挿入先のスライドがありません。");

            var pasted = activeSlide.Shapes.Paste();
            try { pasted.Select(); } catch { }
            Log.Write("Inserted: id=" + item.Id);
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
