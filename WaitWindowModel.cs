using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace WpfSample
{
    /// <summary>
    /// 処理待ちウィンドウのバックストアを構成します。
    /// </summary>
    public class WaitWindowModel : INotifyPropertyChanged
    {
        /// <summary>
        /// プロパティの値が変更された時に発生します。 
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// プロパティの値が変更された時に呼び出されます。
        /// </summary>
        /// <param name="name">プロパティ名。</param>
        protected virtual void OnPropertyChanged(string name)
        {
            var handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

        #region Text

        private string _Text = "処理が終了するまでしばらくお待ちください。";

        /// <summary>
        /// ウィンドウに表示するテキストを取得または設定します。
        /// </summary>
        public string Text
        {
            get { return _Text; }
            set
            {
                if (_Text != value)
                {
                    _Text = value;
                    OnPropertyChanged(nameof(Text));
                }
            }
        }

        #endregion // Text

        #region ButtonVisibility

        private Visibility _ButtonVisibility = Visibility.Hidden;

        /// <summary>
        /// ボタンの表示状態を取得または設定します。
        /// </summary>
        public Visibility ButtonVisibility
        {
            get { return _ButtonVisibility; }
            set
            {
                if (_ButtonVisibility != value)
                {
                    _ButtonVisibility = value;
                    OnPropertyChanged(nameof(ButtonVisibility));
                }
            }
        }

        #endregion // ButtonVisibility
    }
}
