using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointUsefulTools
{
    internal class CellStyleInfo
    {
        public Office.MsoFillType FillType { get; set; }
        public int FillForeColorRGB { get; set; }
        public float FillTransparency { get; set; }

        public string FontName { get; set; }
        public float FontSize { get; set; }
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
        public int FontColorRGB { get; set; }

        public float MarginTop { get; set; }
        public float MarginBottom { get; set; }
        public float MarginLeft { get; set; }
        public float MarginRight { get; set; }
    }

    internal class TableLayoutInfo
    {
        public CellStyleInfo HeaderStyle { get; set; }
        public CellStyleInfo BodyStyle { get; set; }
    }

    internal static class TableLayoutManager
    {
        public static TableLayoutInfo CopyLayout(PowerPoint.Table table)
        {
            var info = new TableLayoutInfo();

            if (table.Rows.Count >= 1 && table.Columns.Count >= 1)
                info.HeaderStyle = CopyCellStyle(table.Cell(1, 1));

            if (table.Rows.Count >= 2 && table.Columns.Count >= 1)
                info.BodyStyle = CopyCellStyle(table.Cell(2, 1));
            else
                info.BodyStyle = info.HeaderStyle;

            return info;
        }

        private static CellStyleInfo CopyCellStyle(PowerPoint.Cell cell)
        {
            var style = new CellStyleInfo();

            var fill = cell.Shape.Fill;
            try { style.FillType = fill.Type; } catch { style.FillType = Office.MsoFillType.msoFillBackground; }
            try { style.FillForeColorRGB = fill.ForeColor.RGB; } catch { }
            try { style.FillTransparency = fill.Transparency; } catch { }

            var font = cell.Shape.TextFrame.TextRange.Font;
            try { style.FontName = font.Name; } catch { }
            try { style.FontSize = font.Size; } catch { }
            try { style.FontBold = font.Bold == Office.MsoTriState.msoTrue; } catch { }
            try { style.FontItalic = font.Italic == Office.MsoTriState.msoTrue; } catch { }
            try { style.FontColorRGB = font.Color.RGB; } catch { }

            var tf = cell.Shape.TextFrame;
            try { style.MarginTop = tf.MarginTop; } catch { }
            try { style.MarginBottom = tf.MarginBottom; } catch { }
            try { style.MarginLeft = tf.MarginLeft; } catch { }
            try { style.MarginRight = tf.MarginRight; } catch { }

            return style;
        }

        public static void ApplyLayout(PowerPoint.Table table, TableLayoutInfo layout)
        {
            if (layout == null) return;

            int rows = table.Rows.Count;
            int cols = table.Columns.Count;

            for (int r = 1; r <= rows; r++)
            {
                var style = (r == 1) ? layout.HeaderStyle : layout.BodyStyle;
                if (style == null) continue;

                for (int c = 1; c <= cols; c++)
                {
                    try { ApplyCellStyle(table.Cell(r, c), style); } catch { }
                }
            }
        }

        private static void ApplyCellStyle(PowerPoint.Cell cell, CellStyleInfo style)
        {
            var fill = cell.Shape.Fill;
            try
            {
                switch (style.FillType)
                {
                    case Office.MsoFillType.msoFillSolid:
                        fill.Solid();
                        fill.ForeColor.RGB = style.FillForeColorRGB;
                        fill.Transparency = style.FillTransparency;
                        break;
                    case Office.MsoFillType.msoFillBackground:
                        fill.Background();
                        break;
                    default:
                        fill.Solid();
                        fill.ForeColor.RGB = style.FillForeColorRGB;
                        fill.Transparency = style.FillTransparency;
                        break;
                }
            }
            catch { }

            var font = cell.Shape.TextFrame.TextRange.Font;
            try { if (!string.IsNullOrEmpty(style.FontName)) font.Name = style.FontName; } catch { }
            try { if (style.FontSize > 0) font.Size = style.FontSize; } catch { }
            try { font.Bold = style.FontBold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; } catch { }
            try { font.Italic = style.FontItalic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; } catch { }
            try { font.Color.RGB = style.FontColorRGB; } catch { }

            var tf = cell.Shape.TextFrame;
            try { tf.MarginTop = style.MarginTop; } catch { }
            try { tf.MarginBottom = style.MarginBottom; } catch { }
            try { tf.MarginLeft = style.MarginLeft; } catch { }
            try { tf.MarginRight = style.MarginRight; } catch { }
        }
    }
}
