using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfSample
{
    /// <summary>
    /// メインウィンドウのバックストアを構成します。
    /// </summary>
    class MainWindowModel : INotifyPropertyChanged
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

        #region DisplayVersion

        private string _DisplayVersion = "ホストバージョン: N/A";

        /// <summary>
        /// 画面表示用のバージョンを取得または設定します。
        /// </summary>
        public string DisplayVersion
        {
            get { return _DisplayVersion; }
            set
            {
                if (_DisplayVersion != value)
                {
                    _DisplayVersion = value;
                    OnPropertyChanged(nameof(DisplayVersion));
                }
            }
        }

        #endregion // DisplayVersion

        #region VoiceNames

        private string[] _VoiceNames = new string[0];

        /// <summary>
        /// ボイス名の配列を取得または設定します。
        /// </summary>
        public string[] VoiceNames
        {
            get { return _VoiceNames; }
            set
            {
                if (_VoiceNames != value)
                {
                    _VoiceNames = value;
                    OnPropertyChanged(nameof(VoiceNames));
                }
            }
        }

        #endregion // VoiceNames

        #region VoicePresetNames

        private string[] _VoicePresetNames = new string[0];

        /// <summary>
        /// ボイスプリセット名の配列を取得または設定します。
        /// </summary>
        public string[] VoicePresetNames
        {
            get { return _VoicePresetNames; }
            set
            {
                if (_VoicePresetNames != value)
                {
                    _VoicePresetNames = value;
                    OnPropertyChanged(nameof(VoicePresetNames));
                }
            }
        }

        #endregion // VoicePresetNames

        #region SelectedVoicePresetIndex

        private int _SelectedVoicePresetIndex = -1;

        /// <summary>
        /// 現在選択されているボイスプリセットのインデックスを取得または設定します。
        /// </summary>
        public int SelectedVoicePresetIndex
        {
            get { return _SelectedVoicePresetIndex; }
            set
            {
                if (_SelectedVoicePresetIndex != value)
                {
                    _SelectedVoicePresetIndex = value;
                    OnPropertyChanged(nameof(SelectedVoicePresetIndex));
                }
            }
        }

        #endregion // SelectedVoicePresetIndex

        #region Volume

        /// <summary>
        /// 音量を取得または設定します。
        /// </summary>
        private double _Volume = 1.0D;

        public double Volume
        {
            get { return _Volume; }
            set
            {
                if (_Volume != value)
                {
                    _Volume = Math.Round(value * 100) / 100D;
                    OnPropertyChanged(nameof(Volume));
                }
            }
        }

        #endregion // Volume

        #region Speed

        private double _Speed = 1.0D;

        /// <summary>
        /// 話速を取得または設定します。
        /// </summary>
        public double Speed
        {
            get { return _Speed; }
            set
            {
                if (_Speed != value)
                {
                    _Speed = Math.Round(value * 100) / 100D;
                    OnPropertyChanged(nameof(Speed));
                }
            }
        }

        #endregion // Speed

        #region Pitch

        private double _Pitch = 1.0D;

        /// <summary>
        /// 声の高さを取得または設定します。
        /// </summary>
        public double Pitch
        {
            get { return _Pitch; }
            set
            {
                if (_Pitch != value)
                {
                    _Pitch = Math.Round(value * 100) / 100D;
                    OnPropertyChanged(nameof(Pitch));
                }
            }
        }

        #endregion // Pitch

        #region PitchRange

        private double _PitchRange = 1.0D;

        /// <summary>
        /// 抑揚を取得または設定します。
        /// </summary>
        public double PitchRange
        {
            get { return _PitchRange; }
            set
            {
                if (_PitchRange != value)
                {
                    _PitchRange = Math.Round(value * 100) / 100D;
                    OnPropertyChanged(nameof(PitchRange));
                }
            }
        }

        #endregion // PitchRange

        #region MiddlePause

        private int _MiddlePause = 150;

        /// <summary>
        /// 短ポーズを取得または設定します。
        /// </summary>
        public int MiddlePause
        {
            get { return _MiddlePause; }
            set
            {
                if (_MiddlePause != value)
                {
                    _MiddlePause = value;
                    OnPropertyChanged(nameof(MiddlePause));
                }
            }
        }

        #endregion // MiddlePause

        #region LongPause

        private int _LongPause = 370;

        /// <summary>
        /// 長ポーズを取得または設定します。
        /// </summary>
        public int LongPause
        {
            get { return _LongPause; }
            set
            {
                if (_LongPause != value)
                {
                    _LongPause = value;
                    OnPropertyChanged(nameof(LongPause));
                }
            }
        }

        #endregion // LongPause

        #region SentencePause

        private int _SentencePause = 800;

        /// <summary>
        /// 文末ポーズを取得または設定します。
        /// </summary>
        public int SentencePause
        {
            get { return _SentencePause; }
            set
            {
                if (_SentencePause != value)
                {
                    _SentencePause = value;
                    OnPropertyChanged(nameof(SentencePause));
                }
            }
        }

        #endregion // SentencePause

        #region Text

        private string _Text = "";

        /// <summary>
        /// 入力テキストを取得または設定します。
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

        #region ListText

        private string _ListText = "";

        /// <summary>
        /// 入力テキストを取得または設定します。
        /// </summary>
        public string ListText
        {
            get { return _ListText; }
            set
            {
                if (_ListText != value)
                {
                    _ListText = value;
                    OnPropertyChanged(nameof(ListText));
                }
            }
        }

        #endregion // ListText

        #region ListVoicePreset

        private string _ListVoicePreset;

        public string ListVoicePreset
        {
            get { return _ListVoicePreset; }
            set
            {
                if (_ListVoicePreset != value)
                {
                    _ListVoicePreset = value;
                    OnPropertyChanged(nameof(ListVoicePreset));
                }
            }
        }

        #endregion // ListVoicePreset

        #region DoAnalysis

        private bool _DoAnalysis;

        public bool DoAnalysis
        {
            get { return _DoAnalysis; }
            set
            {
                if (_DoAnalysis != value)
                {
                    _DoAnalysis = value;
                    OnPropertyChanged(nameof(DoAnalysis));
                }
            }
        }

        #endregion // DoAnalysis

        #region TargetListIndex

        private string _TargetListIndex;

        public string TargetListIndex
        {
            get { return _TargetListIndex; }
            set
            {
                if (_TargetListIndex != value)
                {
                    _TargetListIndex = value;
                    OnPropertyChanged(nameof(TargetListIndex));
                }
            }
        }

        #endregion // TargetListIndex

        #region ListSelectionIndices

        private string _ListSelectionIndices;

        public string ListSelectionIndices
        {
            get { return _ListSelectionIndices; }
            set
            {
                if (_ListSelectionIndices != value)
                {
                    _ListSelectionIndices = value;
                    OnPropertyChanged(nameof(ListSelectionIndices));
                }
            }
        }

        #endregion // ListSelectionIndices

        #region ListSelectionStart

        private int _ListSelectionStart;

        public int ListSelectionStart
        {
            get { return _ListSelectionStart; }
            set
            {
                if (_ListSelectionStart != value)
                {
                    _ListSelectionStart = value;
                    OnPropertyChanged(nameof(ListSelectionStart));
                }
            }
        }

        #endregion // ListSelectionStart

        #region ListSelectionLength

        private int _ListSelectionLength;

        public int ListSelectionLength
        {
            get { return _ListSelectionLength; }
            set
            {
                if (_ListSelectionLength != value)
                {
                    _ListSelectionLength = value;
                    OnPropertyChanged(nameof(ListSelectionLength));
                }
            }
        }

        #endregion // ListSelectionLength

        #region StatusText

        private string _StatusText = "";

        /// <summary>
        /// ステータスバーに状態を表示します。
        /// </summary>
        public string StatusText
        {
            get { return _StatusText; }
            set
            {
                if (_StatusText != value)
                {
                    _StatusText = value;
                    OnPropertyChanged(nameof(StatusText));
                }
            }
        }

        #endregion // StatusText

        #region ShutdownHost

        private bool _ShutdownHost = true;

        public bool ShutdownHost
        {
            get { return _ShutdownHost; }
            set
            {
                if (_ShutdownHost != value)
                {
                    _ShutdownHost = value;
                    OnPropertyChanged(nameof(ShutdownHost));
                }
            }
        }

        #endregion // ShutdownHost

        #region AvailableHosts

        private string[] _AvailableHosts = new string[0];

        public string[] AvailableHosts
        {
            get { return _AvailableHosts; }
            set
            {
                if (_AvailableHosts != value)
                {
                    _AvailableHosts = value;
                    OnPropertyChanged(nameof(AvailableHosts));
                }
            }
        }

        #endregion // AvailableHosts

        #region CurrentHost

        private string _CurrentHost = "";

        public string CurrentHost
        {
            get { return _CurrentHost; }
            set
            {
                if (_CurrentHost != value)
                {
                    _CurrentHost = value;
                    OnPropertyChanged(nameof(CurrentHost));
                }
            }
        }

        #endregion // CurrentHost
    }
}
