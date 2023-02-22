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

// A.I.VOICE Editor API �^�C�v���C�u�����̃C���|�[�g
//#import "libid:5edbd481-4f61-4dc1-b23b-f3b318aa5533" rename_namespace("AIVoiceEditorApi")
//using namespace AIVoiceEditorApi;

int main()
{
	std::wcout.imbue(std::locale(""));
	std::wcout << _T("Sample program start.") << endl;

	int ret = 0;

	// COM������
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
			// �g�p�\�ȃz�X�g�v���O�������̈ꗗ���擾���܂��B
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
						// ��Ƃ��čŏ��Ɍ��������z�X�g���g�p����
						initialhost = hostname.copy();
					}
				}

				//////////////////////////////////////////////////
				// A.I.VOICE Editor API�����������܂��B
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
				// API������������Ă��邩�ǂ������擾���܂��B
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
						// �z�X�g�v���O�������N�����܂��B
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
						// �z�X�g�v���O�����Ɛڑ����܂��B
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

						// �擾
						{
							{
								//////////////////////////////////////////////////
								// �z�X�g�v���O�����̃o�[�W�������擾���܂��B
								//////////////////////////////////////////////////
								_bstr_t version = pTtsControl->Version;

								std::wcout << _T("ITtsControl::Version : ") << version << endl;
							}

							{
								//////////////////////////////////////////////////
								// ���p�\�ȃ{�C�X�����擾���܂��B
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
								// �o�^����Ă���{�C�X�v���Z�b�g�����擾���܂��B
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
								// �z�X�g�v���O�����̏�Ԃ��擾���܂��B
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
								// ���݂̃{�C�X�v���Z�b�g�����擾���܂��B
								//////////////////////////////////////////////////
								currentVoicePresetName = pTtsControl->CurrentVoicePresetName;

								std::wcout << _T("ITtsControl::CurrentVoicePresetName : ") << currentVoicePresetName.GetBSTR() << endl;
							}

							{
								//////////////////////////////////////////////////
								// ���̓e�L�X�g���擾���܂��B
								//////////////////////////////////////////////////
								_bstr_t text = pTtsControl->Text;

								std::wcout << _T("ITtsControl::Text : ") << text.GetBSTR() << endl;
							}

							{
								//////////////////////////////////////////////////
								// ���̓e�L�X�g�̑I���J�n�ʒu���擾���܂��B
								//////////////////////////////////////////////////
								textSelectionStart = pTtsControl->TextSelectionStart;

								std::wcout << _T("ITtsControl::TextSelectionStart : ") << textSelectionStart << endl;
							}

							{
								//////////////////////////////////////////////////
								// ���̓e�L�X�g�̑I�𕶎������擾���܂��B
								//////////////////////////////////////////////////
								textSelectionLength = pTtsControl->TextSelectionLength;

								std::wcout << _T("ITtsControl::TextSelectionLength : ") << textSelectionLength << endl;
							}

							{
								//////////////////////////////////////////////////
								// �}�X�^�[�R���g���[�����擾���܂��B
								//////////////////////////////////////////////////
								_bstr_t mc = pTtsControl->MasterControl;

								std::wcout << _T("ITtsControl::MasterControl : ") << mc.GetBSTR() << endl;
							}
						}

						// �{�C�X�v���Z�b�g�A���[�U�[�����̍ēǂݍ���
						{
							{
								//////////////////////////////////////////////////
								// �{�C�X�v���Z�b�g���ēǍ��݂��܂��B
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
								// �t���[�Y�������ēǍ��݂��܂��B
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
								// �L���|�[�Y�������ēǍ��݂��܂��B
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
								// �P�ꎫ�����ēǍ��݂��܂��B
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

						// �ݒ�
						{
							{
								//////////////////////////////////////////////////
								// ���݂̃{�C�X�v���Z�b�g����ݒ肵�܂��B
								//////////////////////////////////////////////////
								pTtsControl->CurrentVoicePresetName = currentVoicePresetName;

								std::wcout << _T("ITtsControl::CurrentVoicePresetName = \"") << pTtsControl->CurrentVoicePresetName.GetBSTR() << _T("\"") << endl;
							}

							{
								//////////////////////////////////////////////////
								// ���̓e�L�X�g��ݒ肵�܂��B
								//////////////////////////////////////////////////
								pTtsControl->Text = _T("�����X�͌��{�����B�K���A���̎גq�\�s�̉��������Ȃ���΂Ȃ�ʂƌ��ӂ����B�����X�ɂ͐������킩��ʁB�����X�́A���̖q�l�ł���B�J�𐁂��A�r�ƗV��ŕ邵�ė����B����ǂ��׈��ɑ΂��ẮA�l��{�ɕq���ł������B");

								std::wcout << _T("ITtsControl::Text = \"") << pTtsControl->Text.GetBSTR() << _T("\"") << endl;
							}

							{
								//////////////////////////////////////////////////
								// ���̓e�L�X�g���̈ʒu��ݒ肵�܂��B
								//////////////////////////////////////////////////
								pTtsControl->TextSelectionStart = textSelectionStart;

								std::wcout << _T("ITtsControl::TextSelectionStart = ") << pTtsControl->TextSelectionStart << endl;
							}

							{
								//////////////////////////////////////////////////
								// ���̓e�L�X�g�̒�����ݒ肵�܂��B
								//////////////////////////////////////////////////
								pTtsControl->TextSelectionLength = textSelectionLength;

								std::wcout << _T("ITtsControl::TextSelectionLength = ") << pTtsControl->TextSelectionLength << endl;
							}

							{
								//////////////////////////////////////////////////
								// �}�X�^�[�R���g���[����ݒ肵�܂��B
								//////////////////////////////////////////////////

								// �����l����Ƀ����_���Œl��ύX����
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

						// ����
						{
							{
								//////////////////////////////////////////////////
								// �����̍Đ����J�n���܂��B
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
								// �Đ����̉������~���܂��B
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
								// �p�X�擾�o�b�t�@
								TCHAR path[MAX_PATH];

								// �f�X�N�g�b�v�̃p�X���擾
								SHGetSpecialFolderPath(NULL, path, CSIDL_DESKTOP, 0);

								_bstr_t file(_T("\\test.wav"));
								_tcscat_s(path, MAX_PATH, file);

								//////////////////////////////////////////////////
								// ���������̌��ʂ��w�肳�ꂽ�t�@�C���ɏ����݂܂��B
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
				// �z�X�g�v���O�����Ƃ̐ڑ����������܂��B
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

	// COM���
	::CoUninitialize();

	std::wcout << _T("Sample program end. : ") << ret << endl;

	std::wcout << _T("Press Enter") << endl;
	cin.ignore();

	return ret;
}
