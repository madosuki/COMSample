// COMSample.cpp : このファイルには 'main' 関数が含まれています。プログラム実行の開始と終了がそこで行われます。
//

#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "comsuppw.lib")

#include <iostream>
#include <Windows.h>
#include <comutil.h>
#include <string>
#include <vector>
#include <memory>

#import "D:\AI\AIVoice\AIVoiceEditor\AI.Talk.Editor.Api.tlb" rename_namespace("AIVoiceEditorApi")

class AIVoiceEditorApiClass
{
private:
	AIVoiceEditorApi::ITtsControlPtr pTtsControl = NULL;
public:
	AIVoiceEditorApiClass();
	~AIVoiceEditorApiClass()
	{
		if (this->pTtsControl != NULL)
		{
			this->DisConnect();
			::CoUninitialize();
		}
	}

	int Init()
	{
		int result = 0;
		auto hr = ::CoInitialize(0);
		if (FAILED(hr))
		{
			return -1;
		}

		try
		{
			AIVoiceEditorApi::ITtsControlPtr pTtsControl(__uuidof(AIVoiceEditorApi::TtsControl));
			this->pTtsControl = pTtsControl;

			if (pTtsControl == NULL)
			{
				result = -2;
				goto failed;
			}

			auto host_list = pTtsControl->GetAvailableHostNames();
			long lb = 0;
			long ub = 0;
			SafeArrayGetLBound(host_list, 1, &lb);
			SafeArrayGetUBound(host_list, 1, &ub);
			if (lb <= ub)
			{
				_bstr_t initial_host;
				for (auto i = lb; i <= ub; ++i)
				{
					_bstr_t hostname;
					SafeArrayGetElement(host_list, &i, (void**)hostname.GetAddress());
					std::wcout << i << ": " << hostname.GetBSTR() << std::endl;
					if (i == 0)
					{
						initial_host = hostname.copy();
					}
				}

				hr = pTtsControl->Initialize(initial_host);
			}
			else
			{
				hr = E_FAIL;
			}
			SafeArrayDestroy(host_list);

			if (FAILED(hr))
			{
				result = -3;
				goto failed;
			}

			VARIANT_BOOL is_initialized = pTtsControl->IsInitialized;
			if (!is_initialized)
			{
				result = -4;
			}
			goto end;
		}
		catch (_com_error& e)
		{
			std::wcout << "COM Error : " << e.Error() << " : " << e.ErrorMessage() << std::endl;
			goto failed;
		}
		catch (...)
		{
			std::wcout << "Error" << std::endl;
			goto failed;
		}

	end:
		return result;

	failed:
		::CoUninitialize();
		return result;
	}

	HRESULT StartHost()
	{
		if (this->pTtsControl == NULL)
		{
			return 1;
		}

		if (this->pTtsControl->Status == AIVoiceEditorApi::HostStatus_NotRunning)
		{
			return pTtsControl->StartHost();
		}

		return 0;
	}

	HRESULT Connect()
	{
		if (this->pTtsControl == NULL)
		{
			return 1;
		}

		HRESULT hr = this->StartHost();
		if (FAILED(hr))
		{
			return hr;
		}

		if (pTtsControl->Status == AIVoiceEditorApi::HostStatus::HostStatus_NotConnected)
		{
			hr = pTtsControl->Connect();
		}
		return hr;
	}


	HRESULT DisConnect()
	{
		if (this->pTtsControl == NULL)
		{
			return 1;
		}

		HRESULT hr = 0;
		if (this->pTtsControl->Status != AIVoiceEditorApi::HostStatus_NotRunning &&
			this->pTtsControl->Status != AIVoiceEditorApi::HostStatus_NotConnected)
		{

			hr = pTtsControl->Disconnect();
		}
		return hr;
	}

	std::wstring GetVersion()
	{
		if (this->pTtsControl == NULL)
		{
			return L"";
		}

		_bstr_t version = pTtsControl->GetVersion();
		auto result = std::wstring(version.GetBSTR(), SysStringLen(version));
		return result;
	}

	std::vector<std::wstring> GetVoiceList()
	{
		auto result = std::vector<std::wstring>();
		if (this->pTtsControl == NULL)
		{
			return result;
		}

		long lb = 0;
		long ub = 0;
		SAFEARRAY* voice_list = pTtsControl->VoiceNames;

		SafeArrayGetLBound(voice_list, 1, &lb);
		SafeArrayGetUBound(voice_list, 1, &ub);
		for (auto i = lb; i <= ub; ++i)
		{
			_bstr_t voice;
			SafeArrayGetElement(voice_list, &i, (void**)voice.GetAddress());
			result.push_back(std::wstring(voice, SysStringLen(voice)));
		}
		SafeArrayDestroy(voice_list);

		return result;
	}

	std::wstring GetCurrentMasterControl()
	{
		if (this->pTtsControl == NULL)
		{
			return L"";
		}

		_bstr_t master_control = this->pTtsControl->MasterControl;
		return std::wstring(master_control, SysStringLen(master_control));
	}

	std::wstring GetCurrentVoicePresetInfo()
	{
		if (this->pTtsControl == NULL)
		{
			return L"";
		}


		_bstr_t voice_preset = this->pTtsControl->CurrentVoicePresetName;
		_bstr_t voice_preset_json = this->pTtsControl->GetVoicePreset(voice_preset);
		return std::wstring(voice_preset_json, SysStringLen(voice_preset_json));
	}

	void SetVoice(const std::wstring &name)
	{
		if (this->pTtsControl == NULL)
		{
			return;
		}

		_bstr_t target_voice_preset = this->pTtsControl->GetVoicePreset(name.c_str());
		this->pTtsControl->SetVoicePreset(target_voice_preset);
	}

	HRESULT Play(std::wstring text)
	{
		if (this->pTtsControl == NULL)
		{
			return 1;
		}

		pTtsControl->Text = text.c_str();
		pTtsControl->TextSelectionStart = 0;
		pTtsControl->TextSelectionLength = text.size();

		return pTtsControl->Play();
	}

};

AIVoiceEditorApiClass::AIVoiceEditorApiClass()
{
}

int main()
{
	std::wcout.imbue(std::locale(""));

	auto i = std::make_unique<AIVoiceEditorApiClass>();
	i.get()->Init();
	i.get()->Connect();
	std::wcout << i.get()->GetVersion() << std::endl;
	std::wcout << "Current Voice Preset Info: " << i.get()->GetCurrentVoicePresetInfo() << std::endl;
	std::wcout << "Current Master Control: " << i.get()->GetCurrentMasterControl() << std::endl;
	i.get()->Play(L"テストです");


	return 0;
}

