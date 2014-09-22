#include "stdafx.h"
#include <string>
#include <cmath>
#include <tlhelp32.h>
#include "iTunesCOMInterface.h"
#include "iTunesCOMInterface_.h"
#include "hexchat-plugin.h"
using namespace std;

#define PNAME "iTunes"
#define PDESC "iTunes"
#define PVERSION "0.1"

static hexchat_plugin *ph;      /* plugin handle */

string ConvertWCSToMBS(const wchar_t* pstr, long wslen)
{
	int len = WideCharToMultiByte(CP_UTF8, 0, pstr, wslen, NULL, 0, NULL, NULL);

	string dblstr(len, '\0');
	len = WideCharToMultiByte(CP_UTF8, 0 /* no flags */,
		pstr, wslen /* not necessary NULL-terminated */,
		&dblstr[0], len,
		NULL, NULL /* no default char */);

	return dblstr;
}

string ConvertBSTRToMBS(BSTR bstr)
{
	int wslen = SysStringLen(bstr);
	return ConvertWCSToMBS((wchar_t*)bstr, wslen);
}

string smartsize(long bytes)
{
	double sizek = bytes / 1024.0;
	double sizem = sizek / 1024.0;
	double sizeg = sizem / 1024.0;
	double result;
	char *size;
	if (sizek > 0.8)
	{
		if (sizem > 0.8)
		{
			if (sizeg > 0.8)
			{
				result = sizeg;
				size = "GB";
			}
			else
			{
				result = sizem;
				size = "MB";
			}
		}
		else
		{
			result = sizek;
			size = "KB";
		}
	}
	else
	{
		result = bytes;
		size = "B";
	}
	char buf[256];
	memset(buf, 0, 256);
	_snprintf_s(buf, 256, "%.2f%s", result, size);
	return buf;
}

bool isiTunesStarted()
{
	PROCESSENTRY32 entry;
	entry.dwSize = sizeof(PROCESSENTRY32);
	bool result = false;

	HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, NULL);

	if (Process32First(snapshot, &entry) == TRUE)
		while (Process32Next(snapshot, &entry) == TRUE)
			if (_wcsicmp(entry.szExeFile, L"itunes.exe") == 0)
			{
				result = true;
				break;
			}

	CloseHandle(snapshot);
	return result;
}

static int itunes_cb (char *word[], char *word_eol[], void *userdata)
{
	if (!isiTunesStarted())
	{
		hexchat_print(ph, "iTunes is not running.\n");
		return HEXCHAT_EAT_ALL;
	}

	IiTunes *piTunes;

	if (FAILED(CoCreateInstance(CLSID_iTunesApp, NULL, CLSCTX_LOCAL_SERVER, IID_IiTunes, (void **)&piTunes)))
	{
		hexchat_print(ph, "Failed to create iTunes COM object.\n");
		return HEXCHAT_EAT_ALL;
	}

	ITPlayerState state;
	piTunes->get_PlayerState(&state);
	if (state == ITPlayerStateStopped)
	{
		hexchat_print(ph, "iTunes is stopped.\n");
		piTunes->Release(); //release iTunes object
		return HEXCHAT_EAT_ALL;
	}
	IITTrack *track;
	piTunes->get_CurrentTrack(&track);
	BSTR tmpstr;
	track->get_Name(&tmpstr);
	string msg = "me is listening to: \x1F" + ConvertBSTRToMBS(tmpstr) + "\x1F";
	if (track->get_Artist(&tmpstr) != E_POINTER && SysStringLen(tmpstr) > 0)
		msg += " by \x1F" + ConvertBSTRToMBS(tmpstr) + "\x1F";
	track->get_Album(&tmpstr);
	msg += " from \x1F" + ConvertBSTRToMBS(tmpstr) + "\x1F";
	char buf[256];
	long year;
	if (track->get_Year(&year) != E_POINTER && year > 0)
	{
		msg += " (";
		memset(buf, 0, 256);
		_snprintf_s(buf, 256, "%04d", year);
		msg += buf;
		msg += ")";
	}
	msg += " (";
	long pos;
	piTunes->get_PlayerPosition(&pos);
	memset(buf, 0, 256);
	_snprintf_s(buf, 256, "%d:%02d", pos / 60, pos % 60);
	msg += buf;
	track->get_Time(&tmpstr);
	msg += "/" + ConvertBSTRToMBS(tmpstr) + ")";
	long size;
	track->get_Size(&size);
	msg += " " + smartsize(size);
	long rate;
	track->get_BitRate(&rate);
	msg += " ";
	memset(buf, 0, 256);
	_snprintf_s(buf, 256, "%d", rate);
	msg += buf;
	msg += "kbps \x03";
	msg += "08";
	long rating;
	track->get_Rating(&rating);
	rating /= 20;
	for (int i = 0; i < rating; i++)
		msg += "\xE2\x98\x85";
	msg += "\x03";
	hexchat_command(ph, msg.c_str());
	track->Release();
	piTunes->Release(); //release iTunes object
	return HEXCHAT_EAT_ALL;     /* eat this command so HexChat and other plugins can't process it */
}

extern "C"
{
	__declspec(dllexport) void hexchat_plugin_get_info (char **name, char **desc, char **version, void **reserved)
	{
		*name = PNAME;
		*desc = PDESC;
		*version = PVERSION;
	}

	__declspec(dllexport) int hexchat_plugin_init (hexchat_plugin *plugin_handle, char **plugin_name, char **plugin_desc, char **plugin_version, char *arg)
	{
		/* we need to save this for use with any hexchat_* functions */
		ph = plugin_handle;

		/* tell HexChat our info */
		*plugin_name = PNAME;
		*plugin_desc = PDESC;
		*plugin_version = PVERSION;

		CoInitialize(0);

		hexchat_hook_command (ph, "iTunes", HEXCHAT_PRI_NORM, itunes_cb, "", 0);

		hexchat_print (ph, "iTunes plugin loaded successfully!\n");

		return 1;       /* return 1 for success */
	}

	__declspec(dllexport) int hexchat_plugin_deinit(void)
	{
		CoUninitialize();
		hexchat_print (ph, "iTunes plugin unloaded\n");
		return 1;
	}
}