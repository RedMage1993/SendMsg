//*******************************************************************
// Author: Fritz Ammon
// Date: 5 February 2013
// Description: Uses several Windows APIs to repeatedly send strings
// to the window that has focus.
//
// Updated 1 November 2013
// -- improveSleepAcc should be turned on/off only when calling
//	the Sleep API.
// Updated 2 April 2016
// -- Comments about virtual key codes in msgProc and the RNG
//  in use for random.
//*******************************************************************

#pragma comment(lib, "Winmm")

#include <Windows.h>
#include <WinBase.h>
#include <iostream>
#include <cstdlib>
#include <ctime>
#include <cstring>
#include <string>

using std::cout;
using std::cin;
using std::endl;
using std::string;
using std::getline;
using std::srand;
using std::rand;
using std::time;
using std::memset;

//*********************
// Function Prototypes

bool improveSleepAcc(bool activate = true);
int random(int min, int max);
void showOptions();
void closeThread(HANDLE& hThread);
DWORD WINAPI killProc(LPVOID lpParameter);
DWORD WINAPI msgProc(LPVOID lpParameter);

//******************
// Global Variables

DWORD msWaitPerMsg = 100;
DWORD msWaitPerKey = 20;
string msgToSend = "Hello, world!";
HANDLE hmsgProc = 0;
unsigned int repetitions = 2;
bool randomizeMsg = false;

int main()
{
	short option;
	bool done = false;
	HANDLE hKillProc;

	srand(static_cast<unsigned int> (time(reinterpret_cast<time_t*> (NULL))));

	cout << "A keystroke simulator by Fritz Ammon.\n\n";
	cout << std::boolalpha;

	while (!done)
	{
		showOptions();
		cin >> option;
		cout << endl;

		cin.clear();
		cin.ignore(200, '\n');

		switch (option)
		{
		case 0:
			cout << "Enter string: ";
			getline(cin, msgToSend);
			cout << endl;

			randomizeMsg = false;
			break;
		case 1:
			randomizeMsg = true;
			break;
		case 2:
			cout << "Enter new msg delay: ";
			cin >> msWaitPerMsg;
			cout << endl;
			break;
		case 3:
			cout << "Enter new key delay: ";
			cin >> msWaitPerKey;
			cout << endl;
			break;
		case 4:
			cout << "Enter repetitions: ";
			cin >> repetitions;
			cout << endl;
			break;
		case 5:
			if (!hmsgProc)
			{ // Activate procedure.
				hmsgProc = CreateThread(
					NULL,
					0,
					msgProc,
					NULL,
					0,
					NULL);

				hKillProc = CreateThread(
					NULL,
					0,
					killProc,
					NULL,
					0,
					NULL);
			}
			else
			{ // Deactivate procedure.
				closeThread(hmsgProc);
				closeThread(hKillProc);
			}

			break;
		case 6:
			if (hmsgProc)
			{
				closeThread(hmsgProc);
				closeThread(hKillProc);
			}

			done = true;
			continue;
		}
	}

	return 0;
}

//***********************************************************
// improveSleepAcc
// activate: Tells whether or not to turn on the effect.
//
// It returns bool, but for this purpose it is not necessary
// to check. Prepared for any future use.
//
// This function improves the accuracy of the Sleep API
// by reducing the time resolution to the minimum supported
// value of the system. It may be dangerous to have multiple
// calls active by different processes and, therefore, it
// is necessary to deactivate the change when it is not
// needed.

bool improveSleepAcc(bool activate)
{
	TIMECAPS tc;
	MMRESULT mmr;

	// Fill the TIMECAPS structure.
	if (timeGetDevCaps(&tc, sizeof(tc)) != MMSYSERR_NOERROR)
		return false;

	if (activate)
		mmr = timeBeginPeriod(tc.wPeriodMin);
	else
		mmr = timeEndPeriod(tc.wPeriodMin);

	if (mmr != TIMERR_NOERROR)
		return false;

	return true;
}

// end improveSleepAcc

//*****************************************************************
// showOptions
//
// This function displays multiple options available to the user
// during runtime, and makes sure that the proper third option
// is presented depending on whether or not a procedure is already
// active.

void showOptions()
{
	cout << "Choose option:\n";
	cout << "  0: Set Message (currently \""
		 << msgToSend << "\")\n";
	cout << "  1: Randomize Message (currently "
		 << randomizeMsg << ")\n";
	cout << "  2: Set Wait Per Message (currently "
		 << msWaitPerMsg << " ms)\n";
	cout << "  3: Set Wait Per Key (currently "
		 << msWaitPerKey << " ms)\n";
	cout << "  4: Set Reptitions (currently "
		 << repetitions << ")\n";

	if (!hmsgProc)
		cout << "  5: Activate Procedure (F2 to start, F4 to stop)\n";
	else
		cout << "  5: Deactivate Procedure\n";

	cout << "  6: Quit\n";
	cout << "Your selection is: ";
}

// end showOptions

//************************************************************
// random
// min: The smallest integer possible.
// max: The largest integer possible.
//
// Returns a random number in the range specified, inclusive.
// Of course, this may be improved with a better RNG function.
// This one is not so uniform.

int random(int min, int max)
{
	return min + rand() % (max - min + 1);
}

// end random

//********************************************************************
// closeThread
// hThread: A handle to the thread to be terminated.
//
// Terminates the thread, closes its handle, removing it from
// the system, and finally sets the handle to 0 for further
// processing.

void closeThread(HANDLE& hThread)
{
	DWORD dwExitCode;

	GetExitCodeThread(hThread, &dwExitCode);
	TerminateThread(hThread, dwExitCode);
	CloseHandle(hThread);
	hThread = 0;
}

// end closeThread

//**********************************************************
// msgProc
// lpParameter: Data passed into the procedure.
//
// This thread/procedure will handle the repetition of
// message sending. It is created so that the main loop
// that focuses on user input is able to terminate the
// procedure itself with a call to the TerminateThread API.

DWORD WINAPI msgProc(LPVOID lpParameter)
{
	DWORD dwExtraInfo = GetMessageExtraInfo();
	bool active = true;
	INPUT inKeys[4];
	WORD noKeys;
	short vKey = 0;
	size_t msgLen = 0;

	UNREFERENCED_PARAMETER(lpParameter);

	// Initialize constant members of structures.
	memset(inKeys, 0, sizeof(INPUT) * 4);
	for (int key = 0; key < 4; key++)
	{
		inKeys[key].type = 1;
		inKeys[key].ki.dwExtraInfo = dwExtraInfo;
	}

	// Preparing what will remain static throughout the life of this
	// function.
	// The first and last key will be VK_SHIFT if necessary.
	inKeys[0].ki.dwFlags = KEYEVENTF_EXTENDEDKEY;
	inKeys[0].ki.wVk = inKeys[3].ki.wVk = VK_SHIFT;
	inKeys[2].ki.dwFlags = KEYEVENTF_KEYUP;
	inKeys[3].ki.dwFlags = KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP;

	while (active)
	{
		// Wait for hotkey to be pressed for further execution.
		while (!(GetAsyncKeyState(VK_F2) & 0x8000))
		{
			improveSleepAcc(true);
			Sleep(40);
			improveSleepAcc(false);
		}

		if (!randomizeMsg)
			msgLen = msgToSend.size();

		for (unsigned int rep = 0; rep < repetitions; rep++)
		{
			if (randomizeMsg)
				msgLen = random(0x20, 0x7E); // Displayable characters.

			for (size_t ch = 0; ch < msgLen; ch++)
			{
				if (randomizeMsg)
					vKey = VkKeyScan(static_cast<WCHAR> (random(0x20, 0x7E)));
				else
					vKey = VkKeyScan(msgToSend[ch]);
				
				// Check if extended key (shift) required.
				// VkKeyScan returns a 2-byte value, including a 0x100 metadata
				// if the character was uppercase.
				if ((vKey & 0x100) == 0x100)
					noKeys = 4;
				else
					noKeys = 2;

				// SendInput does not expect an uppercase (0x100) form of the virtual key
				// since it will use a VK_SHIFT should it be required.
				// This will remove the 0x100 and keep only the rightmost byte.
				inKeys[1].ki.wVk = inKeys[2].ki.wVk = vKey & 0xFF;
				
				// Notice the trick here is to begin either at the start of inKeys or at
				// the 2nd element (key) to ignore VK_SHIFT altogether.
				// This is possible thanks to the noKeys parameter, which will result in the
				// key being ignored should noKeys == 2.
				SendInput(noKeys, noKeys == 4 ? inKeys : &inKeys[1], sizeof(INPUT));

				improveSleepAcc(true);
				Sleep(msWaitPerKey);
				improveSleepAcc(false);
			}

			// Press Enter to send message.
			inKeys[1].ki.wVk = inKeys[2].ki.wVk = VK_RETURN;

			SendInput(2, &inKeys[1], sizeof(INPUT));

			improveSleepAcc(true);
			Sleep(msWaitPerMsg);
			improveSleepAcc(false);
		}
	}

	return 0;
}

// end spamProc

//************************************************************
// killProc
//
// This procedure is responsible for terminating the spamProc
// and restarting it.

DWORD WINAPI killProc(LPVOID lpParameter)
{
	bool active = true;

	UNREFERENCED_PARAMETER(lpParameter);

	while (active)
	{
		if ((GetAsyncKeyState(VK_F4) & 0x8000) && hmsgProc)
		{
			closeThread(hmsgProc);

			hmsgProc = CreateThread(
				NULL,
				0,
				msgProc,
				NULL,
				0,
				NULL);
		}

		improveSleepAcc(true);
		Sleep(40);
		improveSleepAcc(false);
	}

	return 0;
}

// end killProc
