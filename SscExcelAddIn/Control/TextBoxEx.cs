using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace SscExcelAddIn.Control
{
    /// <summary>
    /// <see href="http://www.madeinclinic.jp/c/20180421/"/>
    /// </summary>
    public class TextBoxEx : TextBox
    {
        /// <summary>
        /// IME利用中かどうか判定するフラグ
        /// </summary>
        private bool isImeOnConv = false;
        /// <summary>
        /// IMEでの変換決定のEnterキーに反応させないためのバッファ
        /// </summary>
        private int EnterKeyBuffer;
        /// <summary>
        /// IME入力以外のEnterキー押下イベントをEnterKeyUpのみで受け取るようにするか
        /// </summary>
        public bool TakeOverEnterKeyUp { get; set; }

        /// <summary>
        /// IME入力以外のEnterキー押下イベント
        /// <see href="https://blog.okazuki.jp/entry/2014/08/22/211021"/>
        /// <see href="https://wpf.2000things.com/2012/08/07/619-event-sequence-for-the-key-updown-events/"/>
        /// </summary>
        public static RoutedEvent EnterKeyUpEvent = EventManager.RegisterRoutedEvent(
            "EnterKeyUp", // イベント名
            RoutingStrategy.Bubble, // イベントタイプ
            typeof(KeyEventHandler), // イベントハンドラの型
            typeof(TextBoxEx)); // イベントのオーナー

        /// <summary>
        /// IME入力以外のEnterキー押下イベント
        /// </summary>
        public event KeyEventHandler EnterKeyUp
        {
            add
            {
                AddHandler(EnterKeyUpEvent, value, handledEventsToo: false);
            }
            remove
            {
                RemoveHandler(EnterKeyUpEvent, value);
            }
        }

        /// <summary>
        /// IME不使用中のEnterキー押下を判定可能なテキストボックス
        /// </summary>
        public TextBoxEx() : base()
        {
            TextCompositionManager.AddPreviewTextInputHandler(this, OnPreviewTextInput);
            TextCompositionManager.AddPreviewTextInputUpdateHandler(this, OnPreviewTextInputUpdate);
            this.KeyUp += OnKeyUp;
        }

        /*
        public bool IsDirectInput(KeyEventArgs e)
        {
            if (e.Source == this && e.RoutedEvent.Name == "KeyUp")
            {
                return !isImeOnConv;
            }
            throw new NotSupportedException("IsDirectInput is supported only in KeyUp event.");
        }
        */

        private void OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (isImeOnConv)
            {
                EnterKeyBuffer = 1;
            }
            else
            {
                EnterKeyBuffer = 0;
            }
            isImeOnConv = false;
        }

        private void OnPreviewTextInputUpdate(object sender, TextCompositionEventArgs e)
        {
            if (e.TextComposition.CompositionText.Length == 0)
            {
                isImeOnConv = false;
            }
            else
            {
                isImeOnConv = true;
            }
        }

        private void OnKeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (isImeOnConv == false && e.Key == Key.Enter && EnterKeyBuffer == 1)
            {
                EnterKeyBuffer = 0;
            }
            else if (isImeOnConv == false && e.Key == Key.Enter && EnterKeyBuffer == 0)
            {
                if (TakeOverEnterKeyUp)
                {
                    e.Handled = true;
                }
                KeyEventArgs kea = new KeyEventArgs(e.KeyboardDevice,
                    Keyboard.PrimaryDevice.ActiveSource, e.Timestamp, e.Key)
                {
                    RoutedEvent = EnterKeyUpEvent
                };
                this.RaiseEvent(kea);
            }
        }
    }
}
