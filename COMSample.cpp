#include "stdafx.h"

#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "ai.talk.editor.api.lib")
#include <Windows.h>
#include <comutil.h>
#include <objbase.h>
#include <iostream>
#include <time.h>
#include <stdlib.h>
#include <shlobj.h>
#include <direct.h>
#include <string>

using namespace std;

// A.I.VOICE Editor API タイプライブラリのインポート
//#import "libid:5edbd481-4f61-4dc1-b23b-f3b318aa5533" rename_namespace("AIVoiceEditorApi")
//using namespace AIVoiceEditorApi;

int main()
{
	std::wcout.imbue(std::locale(""));
	std::wcout << _T("Sample program start.") << endl;

	int ret = 0;

	// COM初期化
	HRESULT hr = ::CoInitialize(0);
	if (FAILED(hr))
	{
		std::wcout << _T("::CoInitialize() Failed.") << endl;
		return -1;
	}

	std::wcout << _T("::CoInitialize() Succeeded.") << endl;

	try
	{
		ITtsControlPtr pTtsControl(__uuidof(TtsControl));
		if (pTtsControl == NULL)
		{
			std::wcout << _T("Construct ITtsControl Failed.") << endl;
			ret = -2;
		}
		else
		{
			std::wcout << _T("Construct ITtsControl Succeeded.") << endl;

			//////////////////////////////////////////////////
			// 使用可能なホストプログラム名の一覧を取得します。
			//////////////////////////////////////////////////
			SAFEARRAY* hosts = pTtsControl->GetAvailableHostNames();

			std::wcout << _T("ITtsControl::GetAvailableHostNames() :") << endl;

			long lb;
			long ub;
			SafeArrayGetLBound(hosts, 1, &lb);
			SafeArrayGetUBound(hosts, 1, &ub);
			if (lb <= ub)
			{
				_bstr_t initialhost;
				for (long i = lb; i <= ub; i++)
				{
					_bstr_t hostname;
					SafeArrayGetElement(hosts, &i, (void**)hostname.GetAddress());
					std::wcout << _T("  ") << i + 1 << _T(". ") << hostname.GetBSTR() << endl;
					if (i == 0)
					{
						// 例として最初に見つかったホストを使用する
						initialhost = hostname.copy();
					}
				}

				//////////////////////////////////////////////////
				// A.I.VOICE Editor APIを初期化します。
				//////////////////////////////////////////////////
				hr = pTtsControl->Initialize(initialhost);
			}
			else
			{
				std::wcout << _T("Host not found.") << endl;
				hr = E_FAIL;
			}
			SafeArrayDestroy(hosts);

			if (FAILED(hr))
			{
				std::wcout << _T("ITtsControl::Initialize() Failed.") << endl;
				ret = -3;
			}
			else
			{
				//////////////////////////////////////////////////
				// APIが初期化されているかどうかを取得します。
				//////////////////////////////////////////////////
				VARIANT_BOOL initialized = pTtsControl->IsInitialized;

				std::wcout << _T("ITtsControl::IsInitialized : ") << (initialized ? _T("true") : _T("false")) << endl;

				if (initialized == false)
				{
					std::wcout << _T("ITtsControl::IsInitialized = false.") << endl;
					ret = -4;
				}
				else
				{
					std::wcout << _T("ITtsControl::Initialize() Succeeded.") << endl;

					if (pTtsControl->Status == HostStatus::HostStatus_NotRunning)
					{
						//////////////////////////////////////////////////
						// ホストプログラムを起動します。
						//////////////////////////////////////////////////
						hr = pTtsControl->StartHost();
						if (FAILED(hr))
						{
							std::wcout << _T("ITtsControl::StartHost() Failed.") << endl;
							ret = -5;
						}
						else
						{
							std::wcout << _T("ITtsControl::StartHost() Succeeded.") << endl;
						}
					}

					if (pTtsControl->Status == HostStatus::HostStatus_NotConnected)
					{
						//////////////////////////////////////////////////
						// ホストプログラムと接続します。
						//////////////////////////////////////////////////
						hr = pTtsControl->Connect();
						if (FAILED(hr))
						{
							std::wcout << _T("ITtsControl::Connect() Failed.") << endl;
							ret = -6;
						}
						else
						{
							std::wcout << _T("ITtsControl::Connect() Succeeded.") << endl;
						}
					}

					if (SUCCEEDED(hr))
					{
						long textSelectionStart = 0;
						long textSelectionLength = 0;
						_bstr_t currentVoicePresetName;

						// 取得
						{
							{
								//////////////////////////////////////////////////
								// ホストプログラムのバージョンを取得します。
								//////////////////////////////////////////////////
								_bstr_t version = pTtsControl->Version;

								std::wcout << _T("ITtsControl::Version : ") << version << endl;
							}

							{
								//////////////////////////////////////////////////
								// 利用可能なボイス名を取得します。
								//////////////////////////////////////////////////
								SAFEARRAY* voices = pTtsControl->VoiceNames;

								std::wcout << _T("ITtsControl::VoiceNames :") << endl;

								SafeArrayGetLBound(voices, 1, &lb);
								SafeArrayGetUBound(voices, 1, &ub);
								for (long i = lb; i <= ub; i++)
								{
									_bstr_t voice;
									SafeArrayGetElement(voices, &i, (void**)voice.GetAddress());
									std::wcout << _T("  ") << i + 1 << _T(". ") << voice.GetBSTR() << endl;
								}
								SafeArrayDestroy(voices);
							}

							{
								//////////////////////////////////////////////////
								// 登録されているボイスプリセット名を取得します。
								//////////////////////////////////////////////////
								SAFEARRAY* presets = pTtsControl->VoicePresetNames;

								std::wcout << _T("ITtsControl::VoicePresetNames :") << endl;

								SafeArrayGetLBound(presets, 1, &lb);
								SafeArrayGetUBound(presets, 1, &ub);
								for (long i = lb; i <= ub; i++)
								{
									_bstr_t preset;
									SafeArrayGetElement(presets, &i, (void**)preset.GetAddress());
									std::wcout << _T("  ") << i + 1 << _T(". ") << preset.GetBSTR() << endl;
								}
								SafeArrayDestroy(presets);
							}

							{
								//////////////////////////////////////////////////
								// ホストプログラムの状態を取得します。
								//////////////////////////////////////////////////
								HostStatus status = pTtsControl->Status;

								std::wcout << _T("ITtsControl::Status : ") <<
									(status == HostStatus_Busy ? _T("Busy") :
										status == HostStatus_Idle ? _T("Idle") :
										status == HostStatus_NotConnected ? _T("NotConnected") :
										_T("NotRunning")) << endl;
							}

							{
								//////////////////////////////////////////////////
								// 現在のボイスプリセット名を取得します。
								//////////////////////////////////////////////////
								currentVoicePresetName = pTtsControl->CurrentVoicePresetName;

								std::wcout << _T("ITtsControl::CurrentVoicePresetName : ") << currentVoicePresetName.GetBSTR() << endl;
							}

							{
								//////////////////////////////////////////////////
								// 入力テキストを取得します。
								//////////////////////////////////////////////////
								_bstr_t text = pTtsControl->Text;

								std::wcout << _T("ITtsControl::Text : ") << text.GetBSTR() << endl;
							}

							{
								//////////////////////////////////////////////////
								// 入力テキストの選択開始位置を取得します。
								//////////////////////////////////////////////////
								textSelectionStart = pTtsControl->TextSelectionStart;

								std::wcout << _T("ITtsControl::TextSelectionStart : ") << textSelectionStart << endl;
							}

							{
								//////////////////////////////////////////////////
								// 入力テキストの選択文字数を取得します。
								//////////////////////////////////////////////////
								textSelectionLength = pTtsControl->TextSelectionLength;

								std::wcout << _T("ITtsControl::TextSelectionLength : ") << textSelectionLength << endl;
							}

							{
								//////////////////////////////////////////////////
								// マスターコントロールを取得します。
								//////////////////////////////////////////////////
								_bstr_t mc = pTtsControl->MasterControl;

								std::wcout << _T("ITtsControl::MasterControl : ") << mc.GetBSTR() << endl;
							}
						}

						// ボイスプリセット、ユーザー辞書の再読み込み
						{
							{
								//////////////////////////////////////////////////
								// ボイスプリセットを再読込みします。
								//////////////////////////////////////////////////
								hr = pTtsControl->ReloadVoicePresets();

								if (SUCCEEDED(hr))
								{
									std::wcout << _T("ITtsControl::ReloadVoicePresets() Succeeded.") << endl;
								}
								else
								{
									std::wcout << _T("ITtsControl::ReloadVoicePresets() Failed.") << endl;
								}
							}

							{
								//////////////////////////////////////////////////
								// フレーズ辞書を再読込みします。
								//////////////////////////////////////////////////
								hr = pTtsControl->ReloadPhraseDictionary();

								if (SUCCEEDED(hr))
								{
									std::wcout << _T("ITtsControl::ReloadPhraseDictionary() Succeeded.") << endl;
								}
								else
								{
									std::wcout << _T("ITtsControl::ReloadPhraseDictionary() Failed.") << endl;
								}
							}

							{
								//////////////////////////////////////////////////
								// 記号ポーズ辞書を再読込みします。
								//////////////////////////////////////////////////
								hr = pTtsControl->ReloadSymbolDictionary();

								if (SUCCEEDED(hr))
								{
									std::wcout << _T("ITtsControl::ReloadSymbolDictionary() Succeeded.") << endl;
								}
								else
								{
									std::wcout << _T("ITtsControl::ReloadSymbolDictionary() Failed.") << endl;
								}
							}

							{
								//////////////////////////////////////////////////
								// 単語辞書を再読込みします。
								//////////////////////////////////////////////////
								hr = pTtsControl->ReloadWordDictionary();

								if (SUCCEEDED(hr))
								{
									std::wcout << _T("ITtsControl::ReloadWordDictionary() Succeeded.") << endl;
								}
								else
								{
									std::wcout << _T("ITtsControl::ReloadWordDictionary() Failed.") << endl;
								}
							}
						}

						// 設定
						{
							{
								//////////////////////////////////////////////////
								// 現在のボイスプリセット名を設定します。
								//////////////////////////////////////////////////
								pTtsControl->CurrentVoicePresetName = currentVoicePresetName;

								std::wcout << _T("ITtsControl::CurrentVoicePresetName = \"") << pTtsControl->CurrentVoicePresetName.GetBSTR() << _T("\"") << endl;
							}

							{
								//////////////////////////////////////////////////
								// 入力テキストを設定します。
								//////////////////////////////////////////////////
								pTtsControl->Text = _T("メロスは激怒した。必ず、かの邪智暴虐の王を除かなければならぬと決意した。メロスには政治がわからぬ。メロスは、村の牧人である。笛を吹き、羊と遊んで暮して来た。けれども邪悪に対しては、人一倍に敏感であった。");

								std::wcout << _T("ITtsControl::Text = \"") << pTtsControl->Text.GetBSTR() << _T("\"") << endl;
							}

							{
								//////////////////////////////////////////////////
								// 入力テキスト中の位置を設定します。
								//////////////////////////////////////////////////
								pTtsControl->TextSelectionStart = textSelectionStart;

								std::wcout << _T("ITtsControl::TextSelectionStart = ") << pTtsControl->TextSelectionStart << endl;
							}

							{
								//////////////////////////////////////////////////
								// 入力テキストの長さを設定します。
								//////////////////////////////////////////////////
								pTtsControl->TextSelectionLength = textSelectionLength;

								std::wcout << _T("ITtsControl::TextSelectionLength = ") << pTtsControl->TextSelectionLength << endl;
							}

							{
								//////////////////////////////////////////////////
								// マスターコントロールを設定します。
								//////////////////////////////////////////////////

								// 初期値を基準にランダムで値を変更する
								srand((unsigned)time(NULL));
								double volume = 1.0 * (rand() % 2 == 0) ? 1.2 : 0.8;
								double speed = 1.0 * (rand() % 2 == 0) ? 1.2 : 0.8;
								double pitch = 1.0 * (rand() % 2 == 0) ? 1.2 : 0.8;
								double pitchRange = 1.0 * (rand() % 2 == 0) ? 1.2 : 0.8;
								int middlePause = (int)(150.0 * ((rand() % 2 == 0) ? 1.2 : 0.8));
								int longPause = (int)(370.0 * ((rand() % 2 == 0) ? 1.2 : 0.8));
								int sentencePause = (int)(800.0 * ((rand() % 2 == 0) ? 1.2 : 0.8));

								std::string jsondoc = "{ \"Volume\" : " + std::to_string(volume) + ", " +
									"\"Speed\" : " + std::to_string(speed) + ", " +
									"\"Pitch\" : " + std::to_string(pitch) + ", " +
									"\"PitchRange\" : " + std::to_string(pitchRange) + ", " +
									"\"MiddlePause\" : " + std::to_string(middlePause) + ", " +
									"\"LongPause\" : " + std::to_string(longPause) + ", " +
									"\"SentencePause\" : " + std::to_string(sentencePause) + " }";

								pTtsControl->MasterControl = jsondoc.c_str();

								std::wcout << _T("ITtsControl::MasterControl = ") << pTtsControl->MasterControl.GetBSTR() << endl;
							}
						}

						// 操作
						{
							{
								//////////////////////////////////////////////////
								// 音声の再生を開始します。
								//////////////////////////////////////////////////
								hr = pTtsControl->Play();

								if (SUCCEEDED(hr))
								{
									std::wcout << _T("ITtsControl::Play() Succeeded.") << endl;
								}
								else
								{
									std::wcout << _T("ITtsControl::Play() Failed.") << endl;
									ret = -5;
								}
							}

							Sleep(10000);

							{
								//////////////////////////////////////////////////
								// 再生中の音声を停止します。
								//////////////////////////////////////////////////
								hr = pTtsControl->Stop();

								if (SUCCEEDED(hr))
								{
									std::wcout << _T("ITtsControl::Stop() Succeeded.") << endl;
								}
								else
								{
									std::wcout << _T("ITtsControl::Stop() Failed.") << endl;
								}
							}

							Sleep(1000);

							{
								// パス取得バッファ
								TCHAR path[MAX_PATH];

								// デスクトップのパスを取得
								SHGetSpecialFolderPath(NULL, path, CSIDL_DESKTOP, 0);

								_bstr_t file(_T("\\test.wav"));
								_tcscat_s(path, MAX_PATH, file);

								//////////////////////////////////////////////////
								// 音声合成の結果を指定されたファイルに書込みます。
								//////////////////////////////////////////////////
								hr = pTtsControl->SaveAudioToFile(path);

								if (SUCCEEDED(hr))
								{
									std::wcout << _T("ITtsControl::SaveAudioToFile(\"") << path << _T("\") Succeeded.") << endl;
								}
								else
								{
									std::wcout << _T("ITtsControl::SaveAudioToFile(\"") << path << _T("\") Failed.") << endl;
								}
							}
						}
					}
				}

				//////////////////////////////////////////////////
				// ホストプログラムとの接続を解除します。
				//////////////////////////////////////////////////
				hr = pTtsControl->Disconnect();

				if (SUCCEEDED(hr))
				{
					std::wcout << _T("ITtsControl::Disconnect() Succeeded.") << endl;
				}
				else
				{
					std::wcout << _T("ITtsControl::Disconnect() Failed.") << endl;
				}
			}
		}
	}
	catch (_com_error& e)
	{
		std::wcout << _T("COM Error : ") << e.Error() << " : " << e.ErrorMessage() << endl;
		ret = -99;
	}
	catch (...)
	{
		std::wcout << _T("Error") << endl;
		ret = -999;
	}

	// COM解放
	::CoUninitialize();

	std::wcout << _T("Sample program end. : ") << ret << endl;

	std::wcout << _T("Press Enter") << endl;
	cin.ignore();

	return ret;
}
