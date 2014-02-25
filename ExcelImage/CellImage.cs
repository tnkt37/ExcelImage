using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using System.Windows.Forms;
using TNKTLib;

namespace TNKTLib.ExcelDNA
{
    /// <summary>
    /// セル画像を表すクラス
    /// </summary>
    public class CellImage
    {
        /// <summary>
        /// コピー元の情報
        /// </summary>
        public CellRectangle Rectangle { get; private set; }

        public int Row { get { return Rectangle.Row; } }
        public int Column { get { return Rectangle.Column; } }
        public int Width { get { return Rectangle.Width; } }
        public int Height { get { return Rectangle.Height; } }

        public dynamic Sheet { get; set; }

        /// <summary>
        /// 取得されたセル画像
        /// </summary>
        public dynamic Image { get; private set; }

        public CellImage(dynamic sheet, CellRectangle rect, bool isDeletedBoarder = true)
        {
            Rectangle = rect;
            Image = sheet.Range[sheet.Cells[Row, Column], sheet.Cells[Row + Height - 1, Column + Width - 1]];
            if (isDeletedBoarder)
                Image.Borders.LineStyle = 0;
        }

        /// <summary>
        /// セル画像を描画
        /// </summary>
        /// <param name="sheet">描画先のワークシート</param>
        /// <param name="row">描画先のRow</param>
        /// <param name="col">描画先のColumn</param>
        public void Draw(dynamic sheet, int row, int col)
        {
            Image.Copy(sheet.Range[sheet.Cells[row, col], sheet.Cells[row + Height - 1, col + Width - 1]]);
        }

        //ひきすうおおすぎ
        /// <summary>
        /// 拡大描画します
        /// 遅いです
        /// </summary>
        /// <param name="sheet">描画先のワークシート</param>
        /// <param name="image">描画元セル画像</param>
        /// <param name="magnification">倍率</param>
        /// <param name="row">描画を始めるRow</param>
        /// <param name="column">描画を始めるColumn</param>
        /// <param name="color">描画色</param>
        public void DrawExtend(dynamic sheet, int magnification, int row = 1, int column = 1, int color = 0)
        {
            for (int i = 0; i < Height; i++)
            {
                for (int j = 0; j < Width; j++)
                {
                    var from = Image.Cells[i + 1, j + 1];
                    if (from.Interior.Color == 0xFFFFFF)
                        continue;
                    DrawExpandedPixel(sheet, from, magnification
                        , row + i * magnification, column + j * magnification, color);
                }
            }
        }

        //引数多すぎてつらい
        void DrawExpandedPixel(dynamic sheet, dynamic from, int magnification, int row, int column, int color)
        {
            for (int k = 0; k < magnification; k++)
            {
                for (int l = 0; l < magnification; l++)
                {
                    var to = sheet.Cells[row + k, column + l];
                    if (from.Interior.Color != 0xFFFFFF)
                        to.Interior.Color = color == 0 ? from.Interior.Color : color;
                }
            }
        }
    }
}
