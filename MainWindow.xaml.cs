using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Text.Json;
using Microsoft.Win32;
using AI.Talk.Editor.Api;

namespace WpfSample
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private TtsControl _ttsControl;     // TTS APIの呼び出し用オブジェクト
        private MainWindowModel _vm = new MainWindowModel();    // メインウィンドウのビューモデル

        #region コンストラクタ

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();

            this.DataContext = _vm;

            _ttsControl = new TtsControl();
        }

        #endregion // コンストラクタ

        #region メッセージ

        /// <summary>
        /// ステータスバーにメッセージを表示します。
        /// </summary>
        /// <param name="message">表示メッセージ。</param>
        private void ShowStatus(string message)
        {
            _vm.StatusText = message;
        }

        /// <summary>
        /// 情報メッセージダイアログを表示します。
        /// </summary>
        /// <param name="message">表示メッセージ。</param>
        /// <param name="owner">オーナーウィンドウ。</param>
        private void ShowInformation(Window owner, string message)
        {
            MessageBox.Show(owner, message, "情報", MessageBoxButton.OK, MessageBoxImage.Information);
            this.ShowStatus(message);
        }

        /// <summary>
        /// 情報メッセージダイアログを表示します。
        /// </summary>
        /// <param name="message">表示メッセージ。</param>
        private void ShowInformation(string message)
        {
            this.ShowInformation(this, message);
        }

        /// <summary>
        /// エラーメッセージダイアログを表示します。
        /// </summary>
        /// <param name="message">表示メッセージ。</param>
        /// <param name="owner">オーナーウィンドウ。</param>
        private void ShowError(Window owner, string message)
        {
            MessageBox.Show(owner, message, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            this.ShowStatus(message);
        }

        /// <summary>
        /// エラーメッセージダイアログを表示します。
        /// </summary>
        /// <param name="message">表示メッセージ。</param>
        private void ShowError(string message)
        {
            this.ShowError(this, message);
        }

        #endregion // メッセージ

        #region ホストへの接続・切断

        /// <summary>
        /// ホストへの接続を開始します。
        /// </summary>
        private void Startup()
        {
            try
            {
                if (_ttsControl.Status == HostStatus.NotRunning)
                {
                    // ホストプログラムを起動する
                    _ttsControl.StartHost();
                }

                // ホストプログラムに接続する
                _ttsControl.Connect();

                _vm.DisplayVersion = "ホストバージョン: " + _ttsControl.Version;

                this.RefreshVoiceNames();
                this.RefreshVoicePresetNames();

                TextBox.Focus();

                this.ShowStatus("ホストへの接続を開始しました。");
            }
            catch (Exception ex)
            {
                this.ShowError("ホストへの接続に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// ホストへの接続を終了します。
        /// </summary>
        private void Disconnect()
        {
            try
            {
                // ホストプログラムとの接続を解除する
                _ttsControl.Disconnect();

                this.ShowStatus("ホストへの接続を終了しました。");
            }
            catch (Exception ex)
            {
                this.ShowError("ホストへの接続の終了に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        #endregion // ホストへの接続・切断

        #region ボイス

        /// <summary>
        /// ボイス名の一覧を更新します。
        /// </summary>
        private void RefreshVoiceNames()
        {
            try
            {
                // ボイス名の一覧を取得する
                _vm.VoiceNames = _ttsControl.VoiceNames;

                this.ShowStatus("ボイス一覧を更新しました。");
            }
            catch (Exception ex)
            {

                this.ShowError("ボイス一覧の更新に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// ボイスプリセット名の一覧を更新します。
        /// </summary>
        private void RefreshVoicePresetNames()
        {
            try
            {
                // ボイスプリセット名の配列を取得する
                _vm.VoicePresetNames = _ttsControl.VoicePresetNames;

                // 現在のボイスプリセットを選択する
                for (int i = 0; i < _vm.VoicePresetNames.Length; ++i)
                {
                    if (_ttsControl.CurrentVoicePresetName == _vm.VoicePresetNames[i])
                    {
                        _vm.SelectedVoicePresetIndex = i;
                        break;
                    }
                }

                this.ShowStatus("ボイスプリセット一覧を更新しました。");
            }
            catch (Exception ex)
            {
                this.ShowError("ボイスプリセット一覧の更新に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        #endregion // ボイス

        #region マスターコントロール値の変換

        /// <summary>
        /// 現在のマスターコントロールの各値を JSON 形式の文字列に変換します。
        /// </summary>
        /// <returns>JSON 形式のマスターコントロール値</returns>
        string MasterContorolToJson()
        {
            return "{ \"Volume\" : " + _vm.Volume + ", " +
                "\"Pitch\" : " + _vm.Pitch + ", " +
                "\"Speed\" : " + _vm.Speed + ", " +
                "\"PitchRange\" : " + _vm.PitchRange + ", " +
                "\"MiddlePause\" : " + _vm.MiddlePause + ", " +
                "\"LongPause\" : " + _vm.LongPause + ", " +
                "\"SentencePause\" : " + _vm.SentencePause + " }"; 
        }

        /// <summary>
        /// JSON 形式のマスターコントロール値を反映します。
        /// </summary>
        /// <param name="json">JSON 形式のマスターコントロール値</param>
        void JsonToMasterControl(string json)
        {
            // { "Volume":1.00, "Speed":1.00, "Pitch":1.00, ... }

            var jsonDocument = System.Text.Json.JsonDocument.Parse(json);

            if (jsonDocument != null)
            {
                _vm.Volume = jsonDocument.RootElement.GetProperty("Volume").GetDouble();
                _vm.Speed = jsonDocument.RootElement.GetProperty("Speed").GetDouble();
                _vm.Pitch = jsonDocument.RootElement.GetProperty("Pitch").GetDouble();
                _vm.PitchRange = jsonDocument.RootElement.GetProperty("PitchRange").GetDouble();
                _vm.MiddlePause = jsonDocument.RootElement.GetProperty("MiddlePause").GetInt32();
                _vm.LongPause = jsonDocument.RootElement.GetProperty("LongPause").GetInt32();
                _vm.SentencePause = jsonDocument.RootElement.GetProperty("SentencePause").GetInt32();
            }
        }

        #endregion // マスターコントロール値の変換

        #region イベントハンドラ

        /// <summary>
        /// メインウィンドウがロードされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // 接続可能なホストの一覧を取得する
            _vm.AvailableHosts = _ttsControl.GetAvailableHostNames();

            if (_vm.AvailableHosts.Count() > 0)
            {
                _vm.CurrentHost = _vm.AvailableHosts[0];
            }
        }

        /// <summary>
        /// メインウィンドウが閉じられたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void Window_Closed(object sender, EventArgs e)
        {
            // ホストプログラムとの接続を解除する
            this.Disconnect();

            if (_vm.ShutdownHost)
            {
                // ホストプログラムを終了する
                Task.Factory.StartNew(() =>
                {
                    _ttsControl.TerminateHost();
                });
            }
        }

        /// <summary>
        /// 「接続」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonStartup_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // APIを初期化する
                _ttsControl.Initialize(_vm.CurrentHost);

                // ホストと接続する
                this.Startup();
            }
            catch (Exception ex)
            {
                this.ShowError("接続処理でエラーが発生しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// 「切断」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonDisconnect_Click(object sender, RoutedEventArgs e)
        {
            // ホストプログラムとの接続を解除する
            this.Disconnect();
        }

        /// <summary>
        /// テキスト形式の「ホストから取得」ボタンがクリックされた時に呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonGetText_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // テキスト
                _vm.Text = _ttsControl.Text;

                // 選択開始位置
                TextBox.SelectionStart = _ttsControl.TextSelectionStart;

                // 選択文字数
                TextBox.SelectionLength = _ttsControl.TextSelectionLength;

                TextBox.Focus();

                this.ShowStatus("テキスト形式のテキストを取得しました。");
            }
            catch (Exception ex)
            {
                this.ShowError("テキスト形式のテキストの取得に失敗しました。" + Environment.NewLine + ex.Message);
            }

        }

        /// <summary>
        /// テキスト形式の「ホストへ設定」ボタンがクリックされた時に呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonSetText_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // テキスト
                _ttsControl.Text = _vm.Text;

                // 選択開始位置
                _ttsControl.TextSelectionStart = TextBox.SelectionStart;

                // 選択文字数
                _ttsControl.TextSelectionLength = TextBox.SelectionLength;

                this.ShowStatus("テキスト形式のテキストを設定しました。");
            }
            catch (Exception ex)
            {
                this.ShowError("テキスト形式のテキストの設定に失敗しました。" + Environment.NewLine + ex.Message);
            }

        }

        /// <summary>
        /// ボイス名一覧の「更新」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonRefreshVoiceNames_Click(object sender, RoutedEventArgs e)
        {
            this.RefreshVoiceNames();
        }

        /// <summary>
        /// ボイスプリセット名一覧の「更新」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonRefreshVoicePresetNames_Click(object sender, RoutedEventArgs e)
        {
            this.RefreshVoicePresetNames();
        }

        /// <summary>
        /// ボイスプリセット名一覧の「表示」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonShowVoicePresetInfo_Click(object sender, RoutedEventArgs e)
        {
            // 選択されているボイスプリセットの情報を JSON 形式のテキストとして取得
            var jsonVoicePreset = _ttsControl.GetVoicePreset(_ttsControl.CurrentVoicePresetName);

            // メッセージダイアログで表示
            this.ShowInformation(jsonVoicePreset);

            this.ShowStatus("ボイスプリセットの情報を取得しました。");
        }

        /// <summary>
        /// 「再生」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonPlay_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // テキスト
                _ttsControl.Text = _vm.Text;

                // 選択開始位置
                _ttsControl.TextSelectionStart = TextBox.SelectionStart;

                // 選択文字数
                _ttsControl.TextSelectionLength = TextBox.SelectionLength;

                this.ShowStatus("テキスト形式のテキストを設定しました。");
            }
            catch (Exception ex)
            {
                this.ShowError("テキスト形式のテキストの設定に失敗しました。" + Environment.NewLine + ex.Message);
            }
            try
            {
                // 再生
                _ttsControl.Play();

                this.ShowStatus("音声を再生しました。");
            }
            catch (Exception ex)
            {
                this.ShowError("音声の再生に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// 「停止」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonStop_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 停止
                _ttsControl.Stop();

                this.ShowStatus("音声を停止しました。");
            }
            catch (Exception ex)
            {
                this.ShowError("音声の停止に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        private string _fileName = "";

        /// <summary>
        /// 「音声保存」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonSaveAudioToFile_Click(object sender, RoutedEventArgs e)
        {
            string path = null;
            try
            {
                // テキスト
                _ttsControl.Text = _vm.Text;

                // 選択開始位置
                _ttsControl.TextSelectionStart = TextBox.SelectionStart;

                // 選択文字数
                _ttsControl.TextSelectionLength = TextBox.SelectionLength;

                this.ShowStatus("テキスト形式のテキストを設定しました。");
            }
            catch (Exception ex)
            {
                this.ShowError("テキスト形式のテキストの設定に失敗しました。" + Environment.NewLine + ex.Message);
            }
            try
            {
                // 保存先ファイルパスを取得する

                var sfd = new SaveFileDialog()
                {
                    AddExtension = true,
                    CheckPathExists = true,
                    DefaultExt = "",
                    FileName = _fileName,
                    Filter = "すべてのファイル (*.*)|*.*|WAVEファイル (*.wav)|*.wav|MP3ファイル (*.mp3)|*.mp3|WMAファイル (*.wma)|*.wma",
                    FilterIndex = 1,
                    OverwritePrompt = true,
                    Title = "名前を付けて保存",
                    ValidateNames = true,
                };

                if (!(bool)sfd.ShowDialog(this))
                {
                    return;
                }

                _fileName = System.IO.Path.GetFileName(sfd.FileName);

                path = sfd.FileName;
            }
            catch (Exception ex)
            {
                this.ShowError("音声の保存に失敗しました。" + Environment.NewLine + ex.Message);
                return;
            }

            var context = System.Threading.SynchronizationContext.Current;

            // 待機画面を作成
            /*var dialog = new WaitWindow()
            {
                Title = "音声ファイル保存",
                Owner = this,
            };*/

            // 保存処理を別スレッドで実行
            var task = Task.Factory.StartNew(() =>
            {
                try
                {
                    // 合成結果をファイルに保存する
                    _ttsControl.SaveAudioToFile(path);

                    // 保存処理が終了したら待機画面のメッセージを変更、「OK」ボタンを押下可能にする
                    /*var msg = "音声をファイルに保存しました。";
                    this.ShowStatus(msg);
                    dialog.ViewModel.Text = msg;
                    dialog.ViewModel.ButtonVisibility = Visibility.Visible;*/
                }
                catch (Exception ex)
                {
                    /*context.Send(_ =>
                    {
                        this.ShowError(dialog, "音声の保存に失敗しました。" + Environment.NewLine + ex.Message);
                        dialog.Close();
                    }, null);*/
                }
            });

            // 待機画面を表示
            /*dialog.ShowDialog();*/
        }

        /// <summary>
        /// 「再生時間」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonGetPlayTime_Click(object sender, RoutedEventArgs e)
        {
            var context = System.Threading.SynchronizationContext.Current;
            var dialog = new WaitWindow()
            {
                Title = "再生時間計測",
                Owner = this,
            };

            var task = Task.Factory.StartNew(() =>
            {
                try
                {
                    var playTimeMilliseconds = _ttsControl.GetPlayTime();

                    var playTimeSeconds = playTimeMilliseconds / 1000D;     // 再生時間（単位：秒）
                    var playTimeMinutesPart = playTimeMilliseconds / (60 * 1000);   // 分部分の整数表現
                    var playTimeSecondsAndMilliseconds = playTimeMilliseconds - playTimeMinutesPart * (60 * 1000);  // 秒＋ミリ秒部分の整数表現
                    var playTimeSecondsPart = playTimeSecondsAndMilliseconds / 1000;    // 秒部分の整数表現
                    var playTimeMillisecondsPart = playTimeSecondsAndMilliseconds % 1000;   // ミリ秒部分の整数表現

                    this.ShowStatus("再生時間の計測を完了しました。");
                    dialog.ViewModel.Text = $"再生時間は {playTimeSeconds} 秒（{playTimeMinutesPart}:{playTimeSecondsPart,2:D2}.{playTimeMillisecondsPart,3:D3}）です。"
                                                + Environment.NewLine + Environment.NewLine + "※再生時間に開始ポーズと終了ポーズの長さは含まれません。";
                    dialog.ViewModel.ButtonVisibility = Visibility.Visible;
                }
                catch (Exception ex)
                {
                    context.Send(_ =>
                    {
                        this.ShowError(dialog, "再生時間の計測に失敗しました。" + Environment.NewLine + ex.Message);
                        dialog.Close();
                    }, null);
                }

            });

            dialog.ShowDialog();
        }

        /// <summary>
        /// マスターコントロールの「初期値」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonSetDefaultValues_Click(object sender, RoutedEventArgs e)
        {
            _vm.Volume = 1.0D;
            _vm.Speed = 1.0D;
            _vm.Pitch = 1.0D;
            _vm.PitchRange = 1.0D;
            _vm.MiddlePause = 150;
            _vm.LongPause = 370;
            _vm.SentencePause = 800;
        }

        /// <summary>
        /// マスターコントロールの「ホストから取得」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonGetMasterControl_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                JsonToMasterControl(_ttsControl.MasterControl);

                this.ShowStatus("マスターコントロールを取得しました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "マスターコントロールの取得に失敗しました。" + Environment.NewLine + ex.Message);
            }

        }

        /// <summary>
        /// マスターコントロールの「ホストへ設定」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonSetMasterControl_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _ttsControl.MasterControl = MasterContorolToJson();

                this.ShowStatus("マスターコントロールを設定しました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "マスターコントロールの設定に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// ボイスプリセットの「更新」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonReloadVoicePresets_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _ttsControl.ReloadVoicePresets();

                this.ShowStatus("ボイスプリセットを再読み込みしました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "ボイスプリセットの再読み込みに失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// フレーズ辞書の「更新」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonReloadPhraseDic_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _ttsControl.ReloadPhraseDictionary();

                this.ShowStatus("フレーズ辞書を再読み込みしました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "フレーズ辞書の再読み込みに失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// 単語辞書の「更新」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonReloadWordDic_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _ttsControl.ReloadWordDictionary();

                this.ShowStatus("単語辞書を再読み込みしました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "単語辞書の再読み込みに失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// 記号ポーズ辞書の「更新」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonReloadSymbolDic_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _ttsControl.ReloadSymbolDictionary();

                this.ShowStatus("記号ポーズ辞書を再読み込みしました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "記号ポーズ辞書の再読み込みに失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// リスト形式の選択行の「取得」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonGetListSelection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var indices = _ttsControl.GetListSelectionIndices();
                var sb = new StringBuilder();
                for (int i = 0; i < indices.Length; i++)
                {
                    if (i != 0)
                    {
                        sb.Append(", ");
                    }
                    sb.Append(indices[i].ToString());
                }

                _vm.ListSelectionIndices = sb.ToString();

                this.ShowStatus("リスト形式の選択行を取得しました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "リスト形式の選択行の取得に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// リスト形式の選択行の「設定」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonSetListSelection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var indices = new List<int>();
                foreach (var strIndex in _vm.ListSelectionIndices.Split(','))
                {
                    if (int.TryParse(strIndex, out int index))
                    {
                        indices.Add(index);
                    }
                }

                if (indices.Count > 0)
                {
                    _ttsControl.SetListSelectionIndices(indices.ToArray());
                }
                this.ShowStatus("リスト形式の選択行を設定しました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "リスト形式の選択行の設定に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// リスト形式の選択行に対する「挿入」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonInsertListItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_vm.VoicePresetNames.Contains(_vm.ListVoicePreset))
                {
                    _ttsControl.InsertListItem(_vm.ListVoicePreset, _vm.ListText);
                    this.ShowStatus("リスト形式の要素を挿入しました。");
                }
                else
                {
                    this.ShowError(this, "存在しないボイスプリセットが指定されました。");
                }
            }
            catch (Exception ex)
            {
                this.ShowError(this, "リスト形式の要素の挿入に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// リスト形式の選択行に対する「削除」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonRemoveListSelectedItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _ttsControl.RemoveListItem();

                this.ShowStatus("リスト形式の選択要素を削除しました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "リスト形式の選択要素の削除に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// リスト形式に対する「追加」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonAddListItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_vm.VoicePresetNames.Contains(_vm.ListVoicePreset))
                {
                    _ttsControl.AddListItem(_vm.ListVoicePreset, _vm.ListText);
                    this.ShowStatus("リスト形式の要素を追加しました。");
                }
                else
                {
                    this.ShowError(this, "存在しないボイスプリセットが指定されました。");
                }
            }
            catch (Exception ex)
            {
                this.ShowError(this, "リスト形式の要素の追加に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// リスト形式の選択行に対するセンテンスの「取得」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonGetListSentence_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _vm.ListText = _ttsControl.GetListSentence();

                this.ShowStatus("リスト形式の選択行のセンテンスを取得しました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "リスト形式の選択行のセンテンスの取得に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// リスト形式の選択行に対するセンテンスの「設定」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonSetListSentence_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _ttsControl.SetListSentence(_vm.ListText, _vm.DoAnalysis);

                this.ShowStatus("リスト形式の選択行のセンテンスを設定しました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "リスト形式の選択行のセンテンスの設定に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// リスト形式の選択行に対するボイスプリセットの「取得」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonGetListVoicePreset_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _vm.ListVoicePreset = _ttsControl.GetListVoicePreset();

                this.ShowStatus("リスト形式の選択行のボイスプリセットを取得しました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "リスト形式の選択行のボイスプリセットの取得に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// リスト形式の選択行に対するボイスプリセットの「設定」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonSetListVoicePreset_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_vm.VoicePresetNames.Contains(_vm.ListVoicePreset))
                {
                    _ttsControl.SetListVoicePreset(_vm.ListVoicePreset);
                    this.ShowStatus("リスト形式の選択行のボイスプリセットを設定しました。");
                }
                else
                {
                    this.ShowError(this, "存在しないボイスプリセットが指定されました。");
                }
            }
            catch (Exception ex)
            {
                this.ShowError(this, "リスト形式の選択行のボイスプリセットの設定に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// リスト形式の「クリア」ボタンがクリックされたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void ButtonClearList_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _ttsControl.ClearListItems();

                this.ShowStatus("リスト形式の要素をすべて削除しました。");
            }
            catch (Exception ex)
            {
                this.ShowError(this, "リスト形式の要素の削除に失敗しました。" + Environment.NewLine + ex.Message);
            }
        }

        /// <summary>
        /// ボイスプリセット一覧の選択状態が変更されたときに呼び出されます。
        /// </summary>
        /// <param name="sender">イベント ハンドラーが添付されるオブジェクト。</param>
        /// <param name="e">イベントのデータ。</param>
        private void VoicePresetList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (0 <= _vm.SelectedVoicePresetIndex && _vm.SelectedVoicePresetIndex < _vm.VoicePresetNames.Length)
            {
                // ボイスプリセット一覧で選択されたプリセットをホストのボイスプリセットとして設定する
                _ttsControl.CurrentVoicePresetName = _vm.VoicePresetNames[_vm.SelectedVoicePresetIndex];
            }
        }

        #endregion // イベントハンドラ
    }
}
