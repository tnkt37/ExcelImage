using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TNKTLib.ExcelDNA
{/// <summary>
    /// セル画像のアニメーションを表すクラス
    /// </summary>
    public class CellAnimation
    {
        /// <summary>
        /// セル画像の配列
        /// </summary>
        public CellImage[] Images { get; private set; }

        public bool LoopEnded { get { return controller.LoopEnded; } }

        AnimationIndexController controller;

        public CellAnimation(params CellImage[] images)
        {
            Images = images;
            controller = new AnimationIndexController(1, images.Length)
            {
                Loop = false,
            };
        }
        int counter = 1;

        /// <summary>
        /// セルアニメーションを描画し一コマ進める
        /// </summary>
        /// <param name="sheet">描画先のワークシート</param>
        /// <param name="row">描画先のRow</param>
        /// <param name="col">描画先のColumn</param>
        public void Draw(dynamic sheet, int row, int col)
        {
            Images[controller.CurrentIndex].Draw(sheet, row, col);
            controller.AdvanceFrame(++counter);
        }

        /// <summary>
        /// アニメーションのフレームをリセット
        /// </summary>
        public void Reset()
        {
            controller.Reset();
        }

        /// <summary>
        /// アニメーションをロードします
        /// </summary>
        /// <param name="sheet">取得元のワークシート</param>
        /// <param name="rect">一つのフレームの大きさ</param>
        /// <param name="frameCount">フレーム数</param>
        /// <returns>取得されたアニメーション</returns>
        public static CellAnimation Load(dynamic sheet, CellRectangle rect, int frameCount)
        {
            var imageList = new List<CellImage>();
            for (int i = 0; i < frameCount; i++)
            {
                var rec = new CellRectangle()
                {
                    Row = rect.Row + rect.Height * i,
                    Column = rect.Column,
                    Width = rect.Width,
                    Height = rect.Height,
                };
                var img = new CellImage(sheet, rec);
                imageList.Add(img);
            }
            imageList.Reverse();
            return new CellAnimation(imageList.ToArray());
        }
    }
}
