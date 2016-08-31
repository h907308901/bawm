// BAWM.cpp

#include "stdafx.h"

// total frames and total time of the demo
unsigned int const TOTAL_FRAME = 3636;
unsigned int const TOTAL_TIME = 218000; 

// variables accessed in multiple procedures/threads
HANDLE hFile, hJob, hProcess[6], hThread[6];
HWND hWnd[6];
BOOL bExit; // exit flag

// winmine window enumeration procedure
BOOL CALLBACK EnumThreadWndProc(
	_In_ HWND   hwnd,
	_In_ LPARAM lParam
	)
{
	hWnd[lParam] = hwnd;
	return FALSE; // we only acquire one window for one thread, so we break the enumeration
}

// monitor thread procedure
DWORD WINAPI ThreadProc(LPVOID lpParameter) {
	WaitForMultipleObjects(6, hProcess, FALSE, INFINITE); //  we stop on termination of any one of the winmines 
	bExit = TRUE;
	return 0;
}

// main procedure
int main()
{
	HANDLE hMonThread = NULL;
	HMODULE hModule; // executable image instance handle
	DWORD ret; // used to receive return values
	int id; // sub menu id
	int i, j, k; // for enumeration
	int size; // resource size
	int x, y, w, h; // position and size of windows, also used in mapping
	unsigned int timeinit, framecount; // initial time and frame counter
	int SleepValue; // milliseconds to Sleep after each frame passed
	RECT rect;
	PVOID ptr; // resource pointer
	HRSRC hRes;
	HMENU hMenu, hSubMenu;
	FILE *file = NULL; // the record file
	STARTUPINFO si;
	PROCESS_INFORMATION pi;
	JOBOBJECT_EXTENDED_LIMIT_INFORMATION eli;
	wchar_t tmpdir[1024], wmpath[1024], respath[1024], cmd[1024], txtpath[1024], mucmd[1024];
	// tmpdir:	temporary directory path
	// wmpath:	winmine path (winmine.exe)
	// respath:	resource package path
	// cmd:		unpack command line
	// txtpath: record file path (output.txt)
	// mucmd:	music open command
	char lines[64][256], buf[512];
	// lines:	record file buffer
	// buf:		buffer for data to be written
	// in fact the buffers needn't be so large, but we have to align

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

	// extract winmine.exe
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

	// extract res.cab
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

	// unpack res.cab
	ZeroMemory(&si, sizeof(STARTUPINFO));
	si.dwFlags = STARTF_USESHOWWINDOW;
	si.wShowWindow = 0; // befor everything is ready we don't want the window to appear
	ret = CreateProcess(NULL, cmd, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);
	if (!ret) {
		printf("CreateProcess failed with last error 0x%08x\n", GetLastError());
		goto _goexit;
	}
	WaitForSingleObject(pi.hProcess, INFINITE);
	CloseHandle(pi.hProcess);
	CloseHandle(pi.hThread);

	// job object to kill all winmines on close
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

	// create winmines
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
		ret = EnumThreadWindows(pi.dwThreadId, &EnumThreadWndProc, i); // acquire the window (only one needed)
		if (!(ret || hWnd[i])) { // we test hWnd[i] since EnumThreadWindows returns FALSE on EnumThreadWndProc returning FALSE
			printf("EnumThreadWindows failed with last error 0x%08x\n", GetLastError());
			goto _goexit;
		}
		hMenu = GetMenu(hWnd[i]);
		hSubMenu = GetSubMenu(hMenu, 0);
		id = GetMenuItemID(hSubMenu, (i >= 4 ? 3 : 4)); // middle-3 senior-4
		SendMessage(hWnd[i], WM_COMMAND, id, 0); // imitate click action
		DeleteMenu(hMenu, 0, MF_BYPOSITION); // delete all menus
		DeleteMenu(hMenu, 0, MF_BYPOSITION);
		DestroyMenu(hMenu);
		SetWindowLong(hWnd[i], GWL_STYLE, GetWindowLong(hWnd[i], GWL_STYLE) & ~(WS_CAPTION/*WS_BORDER*/));
		InvalidateRect(hWnd[i], NULL, TRUE); // redraw the window otherwise it will be ugly
		GetWindowRect(hWnd[i], &rect);
		w = rect.right - rect.left;
		h = rect.bottom - rect.top - GetSystemMetrics(SM_CYCAPTION) - GetSystemMetrics(SM_CYMENU); // the height needs shortening since we cut the title bar and the menu bar
		SetWindowPos(hWnd[i], HWND_TOPMOST, x, y, w, h, 0); // topmost
		EnableWindow(hWnd[i], FALSE); // disable window in case user click on it and thus it crashes
		if (i % 2 == 0) { // for upper row
			y = h; // let the next one be in down row
		}
		else { // for down row
			y = 0; // let the next one be in upper row
			x += w; // let the next one be in new column
		}
		Sleep(50);
	}
	for (i = 0; i < 6; i++) {
		ShowWindow(hWnd[i], 1); // now show them
	}

	// open record file
	if (_wfopen_s(&file, txtpath, L"r")) {
		printf("open output.txt failed\n");
		goto _goexit;
	}

	// create monitor thread
	bExit = FALSE;
	hMonThread = CreateThread(NULL, 0, ThreadProc, NULL, 0, NULL);
	if (!hThread) {
		printf("create monitor thread failed with last error 0x%08x\n", GetLastError());
		goto _goexit;
	}

	// open BadApple.mp3
	ret = mciSendString(mucmd, NULL, 0, NULL);
	if (ret) {
		printf("open BadApple.mp3 failed with mci status %d\n", ret);
		goto _goexit;
	}

	// play BadApple.mp3
	ret = mciSendString(L"play music", NULL, 0, NULL);
	if (ret) {
		printf("play BadApple.mp3 failed with mci status %d\n", ret);
		goto _goexit;
	}

	// now begin
	SleepValue = 60;
	framecount = 0;
	timeinit = GetTickCount();
	while (!(feof(file) || bExit)) {
		// a stupid fps stablizer which runs after each 10 frames passed
		// we compare the fps for the rest and the desired fps to determine SleepValue
		if (framecount % 10 == 9) {
			if ((TOTAL_FRAME - framecount) * TOTAL_TIME > TOTAL_FRAME * (TOTAL_TIME + timeinit - GetTickCount()))
				SleepValue--;
			else
				SleepValue++;
		}
		// read data of a frame (60 lines + 1 blank line, 160 chars per line)
		for (i = 0; i <= 60; i++) {
			fgets(lines[i], 256, file);
		}
		// map record data to each winmine
		/*
		|   1~60   |  65~124  |129~160|
		|----------|----------|-------|-
		|    0     |    2     |   4   |1~28
		|----------|----------|-------|-
		|    1     |    3     |   5   |31~58
		|----------|----------|-------|-
		*/
		// middle: 16x16 senoir: 30x16
		for (i = 0; i < 6; i++) {
			ZeroMemory(&buf, sizeof(buf));
			// x, y bias in lines
			x = (i / 2) * 64;
			y = (i % 2) * 30;
			for (j = 0; j < (i >= 4 ? 16 : 30); j++) {
				for (k = 0; k < 16; k++) {
					buf[k * 32 + j] = lines[y + k * 7 / 4][x + j * 2] - 48; // scale: x: 2:1 y: 7:4
				}
				;
			}
			WriteProcessMemory(hProcess[i], (LPVOID)0x01005361L, buf, 512, NULL); // 0x01105362 is a hard code
		}
		for (i = 0; i < 6; i++) {
			InvalidateRect(hWnd[i], NULL, FALSE); // refresh
		}
		Sleep(SleepValue);
		framecount++;
	}

_goexit:
	//cleanup work
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