/***********************************************************************
 * CybWinLock - Subset of functions found in WinLockDLL by A. Miguel   *
 *              Feijão for use by cybera. the original code can be     *
 *              found at:                                              *
 *              http://www.codeproject.com/win32/AntonioWinLock.asp    *
 ***********************************************************************/


 /***********************************************************************
 * WinLockDLL - Library containing functions to Lock Desktop, Taskbar, *
 *              and disable Ctrl+Alt+Del and task switching keys.      *
 *                                                                     *
 * (c) A. Miguel Feijão, 1/12/2004                                     *
 ***********************************************************************/

#define     WIN32_LEAN_AND_MEAN
#define     _WIN32_WINNT 0x0400

#include    <windows.h>
#include    <stdlib.h>
#include    "CybWinLock.h"
#include    "inject.h"

#define     PROGRAM_MANAGER "Program Manager"	// Program manager window name
#define     TASKBAR         "Shell_TrayWnd"		// Taskbar class name

HHOOK		hHook;          // Mouse hook
HINSTANCE	hInst;		    // Instance handle

/*************************
 * Mouse Hook procedure. *
 * (Win 9x)              *
 *************************/
LRESULT CALLBACK MouseHookProc(int nCode, WORD wParam, DWORD lParam) 
{
	if(nCode >= 0)
	{
        if (wParam == WM_LBUTTONDBLCLK)
        {
            if (((MOUSEHOOKSTRUCT *)lParam)->hwnd == GetDesktopWindow())
            {
                return 1;
            }
        }
	}

    return CallNextHookEx(hHook, nCode, wParam, lParam); 
} 


/***********************************
 * Low Level Mouse Hook procedure. *
 * (Win NT4SP3+)                   *
 ***********************************/
LRESULT CALLBACK LowLevelMouseHookProc(int nCode, WORD wParam, DWORD lParam) 
{
	PMSLLHOOKSTRUCT p = (PMSLLHOOKSTRUCT)lParam;
	HWND hWnd = WindowFromPoint(p->pt);

	if(nCode >= 0)
	{
        if ((wParam == WM_LBUTTONDOWN || wParam == WM_RBUTTONDOWN) && hWnd == GetDesktopWindow())
        {
			return 1;
        }
	}

    return CallNextHookEx(hHook, nCode, wParam, lParam); 
} 


/*************************
 * Hide/Show Desktop.    *
 * TRUE=Show, FALSE=Hide * 
 * (Win9x/NT/2K/XP)      *
 *************************/
int DLL_EXP_IMP WINAPI Desktop_Show_Hide(BOOL bShowHide)
{
	OSVERSIONINFO	osvi;
	BOOL			bIsWindowsNT4SP3orLater;
	int				iServicePack;
	char           *p;

	// Determine the current windows version
	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
	GetVersionEx(&osvi);
	for (p = osvi.szCSDVersion; *p  && !isdigit(*p); *p++);
	iServicePack = atoi(p); 
	bIsWindowsNT4SP3orLater = (osvi.dwPlatformId == VER_PLATFORM_WIN32_NT) && 
		                      (((osvi.dwMajorVersion == 4) && (iServicePack >= 3)) || 
						      (osvi.dwMajorVersion > 4));

    if (!bShowHide)
    {
        if (!hHook)
		{
            hHook  = SetWindowsHookEx(bIsWindowsNT4SP3orLater ? WH_MOUSE_LL : WH_MOUSE, 
									  bIsWindowsNT4SP3orLater ? (HOOKPROC)LowLevelMouseHookProc : (HOOKPROC)MouseHookProc, 
                                      hInst, 
                                      0);
			if (!hHook)
				return 0;
		}
    }
    else
    {
        UnhookWindowsHookEx(hHook);
        hHook = NULL;
    }

    return SetWindowPos(FindWindow(NULL, PROGRAM_MANAGER),
                        NULL,
                        0, 0, 0, 0,
                        bShowHide ? SWP_SHOWWINDOW : SWP_HIDEWINDOW | SWP_NOSIZE | SWP_NOMOVE | SWP_NOZORDER);
}



/*************************
 * Hide/Show Taskbar.    *
 * TRUE=Show, FALSE=Hide * 
 * (Win9x/NT/2K/XP)      *
 *************************/
int DLL_EXP_IMP WINAPI Taskbar_Show_Hide(BOOL bShowHide)
{
    HWND    hWnd;

    hWnd = FindWindow(TASKBAR, NULL);
    if (hWnd == NULL)
        return 0;

    ShowWindow(hWnd, bShowHide ? SW_SHOW : SW_HIDE);
    UpdateWindow(hWnd);

    return 1;
}


/**************************************
 * Low Level Keyboard Hook procedure. *
 * (Win NT4SP3+)                      *
 **************************************/

HHOOK hKeyboardHook;  // Old low level keyboard hook 

LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) 
{
    PKBDLLHOOKSTRUCT p;

    if (nCode == HC_ACTION) 
    {
        p = (PKBDLLHOOKSTRUCT) lParam;

        if (
            // WIN key (for Start Menu)
            ((p->vkCode == VK_LWIN) || (p->vkCode == VK_RWIN)) ||       
            // ALT+TAB
            (p->vkCode == VK_TAB && p->flags & LLKHF_ALTDOWN) ||       
            // ALT+ESC
            (p->vkCode == VK_ESCAPE && p->flags & LLKHF_ALTDOWN) ||    
            // CTRL+ESC
            ((p->vkCode == VK_ESCAPE) && ((GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0)) ||
            // CTRL+SHIFT+ESC
            ((p->vkCode == VK_ESCAPE) && ((GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0) && ((GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0)) ||
			// CTRL+ALT+DEL (Unfortunately doesn't work !)
            ((p->vkCode == VK_DELETE) && ( (p->flags & LLKHF_ALTDOWN) != 0 ) && ( (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0))
            )
            return 1;
   }
                    
   return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
}


/*****************************************************
 * Enable/Disable task switching keys.               *
 * (Alt+Tab, Alt+Esc, Ctrl+Esc, Win, Ctrl+Shift+Esc) *
 * TRUE=Enable, FALSE=Disable                        *
 * (Win NT4SP3+, Win2K).                             *
 *****************************************************/
int DLL_EXP_IMP WINAPI TaskSwitching_Enable_Disable(BOOL bEnableDisable)
{
    if (!bEnableDisable)
    {
        if (!hKeyboardHook)
		{
            // Install the low-level keyboard hook
            hKeyboardHook  = SetWindowsHookEx(WH_KEYBOARD_LL, 
                                              LowLevelKeyboardProc, 
                                              hInst, 
                                              0);
			if (!hKeyboardHook)
				return 0;
		}
    }
    else
    {
        UnhookWindowsHookEx(hKeyboardHook);
        hKeyboardHook = NULL;
    }

    return 1;
}


/*****************************************************************
 * Enable/Disable Ctrl+Alt+Del and Ctrl+Shift+Esc key sequences. *
 * TRUE=Enable, FALSE=Disable                                    *
 * (Win 2K).                                                     *
 *****************************************************************/
int DLL_EXP_IMP WINAPI CtrlAltDel_Enable_Disable(BOOL bEnableDisable)
{
	static BOOL bInjected = FALSE;

    if (!bEnableDisable)
    {
		if (!bInjected)
		{
			bInjected = Inject();
			return bInjected;
		}
    }
    else
    {
		if (bInjected)
		{
			bInjected = !Eject();
			return !bInjected;
		}
    }

    return 0;
}


/************************
 * Library entry point. *
 ************************/
BOOL WINAPI DllMain(HINSTANCE hInstDll, DWORD fdwReason, LPVOID lpReserved)
{
    hInst = hInstDll;

	return TRUE;
}
