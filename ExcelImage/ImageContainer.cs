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
    /// シングルトンパターン（）
    /// 画像を格納する辞書配列を保持するクラス
    /// </summary>
    public sealed class ImageContainer
    {

        //シングルトンパターンのようでシングルトンパターンになりきれなかったクラス
        //Loadを使用前に呼び出さないとnullが返るしスレッドセーフじゃないしそもそもシングルトン自体色々アレ


        private IDictionary<string, CellImage> images;

        /// <summary>
        /// 画像を格納するテーブル
        /// </summary>
        public IDictionary<string, CellImage> Images
        {
            get { return images; }
        }

        private IDictionary<string, CellAnimation> animations;

        /// <summary>
        /// アニメーションを格納するテーブル
        /// </summary>
        public IDictionary<string, CellAnimation> Animations
        {
            get { return animations; }
        }


        private static ImageContainer instance;

        /// <summary>
        /// インスタンスを取得します
        /// nullである可能性があることに注意して下さい
        /// </summary>
        public static ImageContainer Instance { get { return instance; } }
        private ImageContainer(dynamic sheet)
        {
            LoadImages(sheet);
        }

        public static void Load(dynamic sheet)
        {
            if (instance == null)
            {
                instance = new ImageContainer(sheet);
            }
        }

        /// <summary>
        /// ワークシートからセル画像をロードします
        /// </summary>
        /// <param name="sheet">ロード元の画像</param>
        /// <param name="images">画像と名前のテーブル</param>
        /// <param name="animes">アニメーションと名前のテーブル</param>
        void LoadImages(dynamic sheet)
        {
            var imagePairs = new Dictionary<string, CellImage>();
            var animePairs = new Dictionary<string, CellAnimation>();
            int currentLine = 2;
            while (!String.IsNullOrEmpty((sheet.Cells[currentLine, 1].Text)))
            {
                var name = sheet.Cells[currentLine, 1].Text;
                var rect = CellRectangle.LoadImageInfo(sheet, currentLine);
                var frame = sheet.Cells[currentLine, 6].Text;
                var image = new CellImage(sheet, rect);
                imagePairs.Add(name, image);

                var frameCount = Int32.Parse(frame);
                //一枚以上あるなら
                if (frameCount > 1)
                {
                    //アニメーションを読み込み
                    animePairs.Add(name, CellAnimation.Load(sheet, rect, frameCount));
                }

                currentLine++;
            }
            images = imagePairs;
            animations = animePairs;
        }
    }
}
