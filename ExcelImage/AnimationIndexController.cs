using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TNKTLib
{
    /// <summary>
    /// アニメーション用の添字を管理するクラス
    /// </summary>
    public class AnimationIndexController
    {
        public int Length { get; set; }
        /// <summary>
        /// 現在の画像の添え字
        /// </summary>
        public int CurrentIndex { get; private set; }
        /// <summary>
        /// 次の画像に入れ替えるまでのフレーム数
        /// </summary>
        public int FrameChangeInterval
        {
            get { return frameChangeInterval; }
            set
            {
                frameChangeInterval = value > 0 ? value : frameChangeInterval;//必ず1以上になるように
            }
        }
        public bool Loop { get; set; }
        public bool LoopEnded { get; set; }
        public bool IsReversed { get; set; }
        int frameChangeInterval = 1;
        //前回更新したフレーム
        int previousFrame = 1;
        public AnimationIndexController(int frameChangeInterval, int length, int initialFrame = 0)
        {
            Length = length;
            FrameChangeInterval = frameChangeInterval;
            Loop = true;
            LoopEnded = false;
            IsReversed = false;
            CurrentIndex = initialFrame;
        }

        public void InitializeFrame(int initialFrame)
        {
            previousFrame = initialFrame;
            Reset();
        }

        bool reversing = false;
        public int AdvanceFrame(int currentFrame)
        {
            var advancedFrame = (currentFrame - previousFrame) / FrameChangeInterval;//前回画像更新時から進んだフレーム / 画像を変更する間隔
            if (advancedFrame != 0)//変更する間隔より進んでいたら
            {
                CurrentIndex += reversing ? -advancedFrame : advancedFrame;//画像を更新する
                previousFrame = currentFrame;//フレームも更新する
            }
            if (CurrentIndex >= Length)//添え字が範囲外なら
            {
                CurrentIndex = Loop ? 0 : Length - 1;//0に戻すか最大値にする
                LoopEnded = true;
                if (IsReversed)//逆向きにするなら
                {
                    reversing = true;
                    CurrentIndex = Length - 1;
                }
            }
            if (IsReversed && CurrentIndex < 0)//reverseの関係で0以下になったら
            {
                reversing = false;
                CurrentIndex = 0;
            }
            return CurrentIndex = Math.Max(CurrentIndex, 0);//インデックスが配列の境界外にならないように調整して画像を更新(Math.MaxはadabancedFrameが負になった時の対策)
        }

        public void Reset()
        {
            CurrentIndex = 0;
            LoopEnded = false;
        }
    }
}
