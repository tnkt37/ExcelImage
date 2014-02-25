using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Windows.Forms;

namespace TNKTLib.ExcelDNA
{
    public class InitialFunctions
    {
#if DEBUG
        
        [ExcelCommand(MenuName = "テスト", MenuText = "メソッドテスト", ShortCut = "^T")]
        public static void Test()
        {
            //Ctrl + Shift + T
            //デバッグ用

            //ここにテストしたいコードを書く
        }

        [ExcelCommand(MenuName = "素材用", MenuText = "素材拡大", ShortCut = "^Q")]
        public static void ShowAnime()
        {
            //素材を拡大して色を変えるメソッド
            //Sheet1のR1C1に素材名R1C2に拡大率(整数)R1C3に色(0なら色はそのまま)
            //Ctrl + Shift + Q
            dynamic app = ExcelDnaUtil.Application;
            dynamic ws1 = app.Worksheets["Sheet1"];
            string name = ws1.Cells[1, 1].Text;
            CellImage[] images = null;
            if (!ImageContainer.Instance.Animations.ContainsKey(name))
            {
                if (!ImageContainer.Instance.Images.ContainsKey(name))
                {
                    MessageBox.Show("No Image");
                    return;
                }
                images = new CellImage[1];
                images[0] = ImageContainer.Instance.Images[name];
            }
            images = ImageContainer.Instance.Animations[name].Images;
            var mag = Convert.ToInt32(ws1.Cells[1, 2].Text);
            var color = Convert.ToInt32((ws1.Cells[1, 3].Text), 16);
            for (int i = 0; i < images.Length; i++)
            {
                var image = images[images.Length - i - 1];
                image.DrawExtend(ws1, mag, 1 + i * image.Height * mag, 4, color);
                var range = ws1.Range[ws1.Cells[1 + i * image.Height * mag, 4], ws1.Cells[1 + (i + 1) * image.Height * mag - 1, 3 + image.Width * mag]];
                range.BorderAround2(LineStyle: 1);//XlDataBarBorderType.xlDataBarBorderSolid
            }
        }


        [ExcelCommand(MenuName = "素材用", MenuText = "アニメーションテスト", ShortCut = "^W")]
        public static void PlayAnime()
        {
            //アニメーション再生テスト用メソッド
            //R1C1に素材名
            //Ctrl + Shift + W でアニメーション実行
            var animes = ImageContainer.Instance.Animations;
            dynamic ws = (ExcelDnaUtil.Application as dynamic).Sheets["Sheet1"];
            string name = ws.Cells[1, 1].Text;
            if (!animes.ContainsKey(name))
            {
                MessageBox.Show("No Image");
                return;
            }
            var anime = animes[name];
            while (!anime.LoopEnded)
            {
                anime.Draw(ws, 1, 1);
                System.Threading.Thread.Sleep(10);
            }
            anime.Reset();
        }
#endif
    }
}
