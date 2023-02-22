using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace WpfSample
{
    /// <summary>
    /// WaitWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class WaitWindow : Window
    {
        #region Win32

        private const int GWL_STYLE = -16;
        private const int WS_SYSMENU = 0x80000;

        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            var hWnd = (new WindowInteropHelper(this)).Handle;
            var style = GetWindowLong(hWnd, GWL_STYLE) & ~WS_SYSMENU;
            SetWindowLong(hWnd, GWL_STYLE, style);
        }

        #endregion // Win32

        #region ViewModel

        private WaitWindowModel _ViewModel = new WaitWindowModel();

        public WaitWindowModel ViewModel
        {
            get { return _ViewModel; }
        }

        #endregion // ViewModel

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public WaitWindow()
        {
            InitializeComponent();

            this.DataContext = _ViewModel;
        }

        /// <summary>
        /// OKボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
