// BAWM.cpp : 定义控制台应用程序的入口点。
//

#include "stdafx.h"

unsigned int const TOTAL_FRAME = 3636;
unsigned int const TOTAL_TIME = 218000; 

HANDLE hFile, hJob, hProcess[6], hThread[6];
HWND hWnd[6];
BOOL bExit;

BOOL CALLBACK EnumThreadWndProc(
	_In_ HWND   hwnd,
	_In_ LPARAM lParam
	)
{
	hWnd[lParam] = hwnd;
	return FALSE;
}

DWORD WINAPI ThreadProc(LPVOID lpParameter) {
	WaitForMultipleObjects(6, hProcess, FALSE, INFINITE);
	bExit = TRUE;
	return 0;
}

int main()
{
	HANDLE hMonThread = NULL;
	HMODULE hModule;
	DWORD ret;
	int id;
	int i, j, k;
	int size;
	int x, y, w, h;
	unsigned int timeinit, framecount;
	int SleepValue;
	RECT rect;
	PVOID ptr;
	HRSRC hRes;
	HMENU hMenu, hSubMenu;
	FILE *file = NULL;
	STARTUPINFO si;
	PROCESS_INFORMATION pi;
	JOBOBJECT_EXTENDED_LIMIT_INFORMATION eli;
	wchar_t tmpdir[1024], wmpath[1024], respath[1024], cmd[1024], txtpath[1024], mucmd[1024];
	char lines[64][256], buf[512];

	hModule = GetModuleHandle(NULL);
	ret = GetTempPath(1024, tmpdir);

	wcscpy_s(wmpath, tmpdir);
	wcscpy_s(respath, tmpdir);
	wcscat_s(wmpath, L"winmine.exe");
	wcscat_s(respath, L"res.cab");

	wcscpy_s(cmd, L"expand \"");
	wcscat_s(cmd, respath);
	wcscat_s(cmd, L"\" -F:* ");
	wcscat_s(cmd, tmpdir);

	wcscpy_s(txtpath, tmpdir);
	wcscat_s(txtpath, L"output.txt");

	wcscpy_s(mucmd, L"open \"");
	wcscat_s(mucmd, tmpdir);
	wcscat_s(mucmd, L"BadApple.mp3\" alias music");

	hFile = INVALID_HANDLE_VALUE;

	printf("Extracting Winmine...\n");
	hRes = FindResource(hModule, MAKEINTRESOURCE(IDR_RES2), L"RES");
	ptr = LoadResource(hModule, hRes);
	size = SizeofResource(hModule, hRes);
	hFile = CreateFile(wmpath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		printf("CreateFile failed with last error 0x%08x\n", GetLastError());
		goto _goexit;
	}
	WriteFile(hFile, ptr, size, &ret, NULL);
	CloseHandle(hFile);

	printf("Extracting Resources...\n");
	hRes = FindResource(hModule, MAKEINTRESOURCE(IDR_RES1), L"RES");
	ptr = LoadResource(hModule, hRes);
	size = SizeofResource(hModule, hRes);
	hFile = CreateFile(respath, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile == INVALID_HANDLE_VALUE) {
		printf("CreateFile failed with last error 0x%08x\n", GetLastError());
		goto _goexit;
	}
	WriteFile(hFile, ptr, size, &ret, NULL);
	CloseHandle(hFile);

	ZeroMemory(&si, sizeof(STARTUPINFO));
	ret = CreateProcess(NULL, cmd, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);
	if (!ret) {
		printf("CreateProcess failed with last error 0x%08x\n", GetLastError());
		goto _goexit;
	}
	WaitForSingleObject(pi.hProcess, INFINITE);
	CloseHandle(pi.hProcess);
	CloseHandle(pi.hThread);

	hJob = CreateJobObject(NULL, NULL);
	if (!hJob) {
		printf("CreateJobObject failed with last error 0x%08x\n", GetLastError());
		goto _goexit;
	}
	ZeroMemory(&eli, sizeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
	eli.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
	ret = SetInformationJobObject(hJob, JobObjectExtendedLimitInformation, &eli, sizeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION));
	if (!hJob) {
		printf("SetInformationJobObjec failed with last error 0x%08x\n", GetLastError());
		goto _goexit;
	}

	ZeroMemory(&hProcess, sizeof(hProcess));
	ZeroMemory(&hThread, sizeof(hThread));
	x = 0;
	y = 0;
	for (i = 0; i < 6; i++) {
		ret = CreateProcess(NULL, wmpath, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);
		if (!ret) {
			printf("CreateProcess failed with last error 0x%08x\n", GetLastError());
			goto _goexit;
		}
		hProcess[i] = pi.hProcess;
		hThread[i] = pi.hThread;
		ret = AssignProcessToJobObject(hJob, pi.hProcess);
		if (!ret) {
			printf("AssignProcessToJobObject failed with last error 0x%08x\n", GetLastError());
			goto _goexit;
		}
		Sleep(50);
		ret = EnumThreadWindows(pi.dwThreadId, &EnumThreadWndProc, i);
		if (!(ret || hWnd[i])) {
			printf("EnumThreadWindows failed with last error 0x%08x\n", GetLastError());
			goto _goexit;
		}
		hMenu = GetMenu(hWnd[i]);
		hSubMenu = GetSubMenu(hMenu, 0);
		id = GetMenuItemID(hSubMenu, (i >= 4 ? 3 : 4));
		SendMessage(hWnd[i], WM_COMMAND, id, 0);
		DeleteMenu(hMenu, 0, MF_BYPOSITION);
		DeleteMenu(hMenu, 0, MF_BYPOSITION);
		DestroyMenu(hMenu);
		InvalidateRect(hWnd[i], NULL, TRUE);
		GetWindowRect(hWnd[i], &rect);
		SetWindowLong(hWnd[i], GWL_STYLE, GetWindowLong(hWnd[i], GWL_STYLE) & ~(WS_CAPTION/*WS_BORDER*/));
		w = rect.right - rect.left;
		h = rect.bottom - rect.top - GetSystemMetrics(SM_CYCAPTION) - GetSystemMetrics(SM_CYMENU);
		SetWindowPos(hWnd[i], HWND_TOPMOST, x, y, w, h, 0);
		EnableWindow(hWnd[i], FALSE);
		if (i % 2 == 0) {
			y = h;
		}
		else {
			y = 0;
			x += w;
		}
		Sleep(50);
	}

	if (_wfopen_s(&file, txtpath, L"r")) {
		printf("open output.txt failed\n");
		goto _goexit;
	}

	bExit = FALSE;
	hMonThread = CreateThread(NULL, 0, ThreadProc, NULL, 0, NULL);
	if (!hThread) {
		printf("create monitor thread failed with last error 0x%08x\n", GetLastError());
		goto _goexit;
	}

	ret = mciSendString(mucmd, NULL, 0, NULL);
	if (ret) {
		printf("open BadApple.mp3 failed with mci status %d\n", ret);
		goto _goexit;
	}

	ret = mciSendString(L"play music", NULL, 0, NULL);
	if (ret) {
		printf("play BadApple.mp3 failed with mci status %d\n", ret);
		goto _goexit;
	}

	SleepValue = 60;
	framecount = 0;
	timeinit = GetTickCount();
	while (!(feof(file) || bExit)) {
		if (framecount % 10 == 9) {
			if ((TOTAL_FRAME - framecount) * TOTAL_TIME > TOTAL_FRAME * (TOTAL_TIME + timeinit - GetTickCount()))
				SleepValue--;
			else
				SleepValue++;
		}
		for (i = 0; i <= 60; i++) {
			fgets(lines[i], 256, file);
		}
		for (i = 0; i < 6; i++) {
			ZeroMemory(&buf, sizeof(buf));
			x = (i / 2) * 64;
			y = (i % 2) * 30;
			for (j = 0; j < (i >= 4 ? 16 : 30); j++) {
				for (k = 0; k < 16; k++) {
					buf[k * 32 + j] = lines[y + k * 7 / 4][x + j * 2] - 48;
				}
				;
			}
			WriteProcessMemory(hProcess[i], (LPVOID)0x01005361L, buf, 512, NULL);
		}
		for (i = 0; i < 6; i++) {
			InvalidateRect(hWnd[i], NULL, FALSE);
		}
		Sleep(SleepValue);
		framecount++;
	}

_goexit:
	mciSendString(L"close music", NULL, 0, NULL);
	if (hMonThread) CloseHandle(hMonThread);
	if (file) fclose(file);
	if (hJob) CloseHandle(hJob);
	for (i = 0; i < 4; i++) {
		if (hProcess[i]) CloseHandle(hProcess[i]);
		if (hThread[i]) CloseHandle(hThread[i]);
	}
    return 0;
}