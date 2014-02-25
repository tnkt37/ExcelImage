using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TNKTLib.ExcelDNA
{
    /// <summary>
    /// セル画像の場所と大きさを表すクラス
    /// </summary>
    public class CellRectangle
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }

        /// <summary>
        /// 決まった形式の画像情報を読み取ります
        /// </summary>
        /// <param name="sheet">読み取るワークシート</param>
        /// <param name="currentLine">現在読んでいる行</param>
        /// <returns></returns>
        public static CellRectangle LoadImageInfo(dynamic sheet, int currentLine)
        {
            //データ読み取り
            var row = sheet.Cells[currentLine, 2].Text;
            var column = sheet.Cells[currentLine, 3].Text;
            var width = sheet.Cells[currentLine, 4].Text;
            var height = sheet.Cells[currentLine, 5].Text;

            return new CellRectangle()
            {
                Row = Int32.Parse(row),
                Column = Int32.Parse(column),
                Width = Int32.Parse(width),
                Height = Int32.Parse(height),
            };
        }
    }
}
