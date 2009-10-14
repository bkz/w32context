///////////////////////////////////////////////////////////////////////////////////////////////////
//
//  Manually launch Explorer context menu for a file/folder.
//
//  Inspired from code in Jeff Prosise column "Wicked Code" (WSJ April 1997).
//
//  Babar k. Zafar (babar.zafar@gmail.com)
//
///////////////////////////////////////////////////////////////////////////////////////////////////

#include <windows.h>
#include <shlobj.h>

///////////////////////////////////////////////////////////////////////////////////////////////////
// Global data.
///////////////////////////////////////////////////////////////////////////////////////////////////

static WNDPROC        g_pOldWndProc;
static LPCONTEXTMENU2 g_pIContext2or3;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Utilities.
///////////////////////////////////////////////////////////////////////////////////////////////////

LPITEMIDLIST GetNextItem(LPITEMIDLIST pidl)
{
    USHORT nLen;

    if ((nLen = pidl->mkid.cb) == 0)
    {
        return NULL;
    }

    return (LPITEMIDLIST)(((LPBYTE) pidl) + nLen);
}

UINT GetItemCount(LPITEMIDLIST pidl)
{
    USHORT nLen;
    UINT nCount;

    nCount = 0;
    while ((nLen = pidl->mkid.cb) != 0) 
    {
        pidl = GetNextItem(pidl);
        nCount++;
    }

    return nCount;
}

LPITEMIDLIST DuplicateItem(LPMALLOC pMalloc, LPITEMIDLIST pidl)
{
    USHORT nLen;
    LPITEMIDLIST pidlNew;

    nLen = pidl->mkid.cb;
    if (nLen == 0)
    {
        return NULL;
    }

    pidlNew = (LPITEMIDLIST) pMalloc->Alloc(nLen + sizeof(USHORT));
    if (pidlNew == NULL)
    {
        return NULL;
    }

    ::CopyMemory(pidlNew, pidl, nLen);
    *((USHORT*)(((LPBYTE) pidlNew) + nLen)) = 0;

    return pidlNew;
}


///////////////////////////////////////////////////////////////////////////////////////////////////
// Render context menu.
///////////////////////////////////////////////////////////////////////////////////////////////////

LRESULT CALLBACK HookWndProc(HWND hWnd, UINT msg, WPARAM wp, LPARAM lp)
{
    switch (msg) 
    { 
    case WM_DRAWITEM:
    case WM_MEASUREITEM:
    case WM_INITMENUPOPUP:
        g_pIContext2or3->HandleMenuMsg(msg, wp, lp);
        return msg == (WM_INITMENUPOPUP ? 0 : TRUE);
    default:
        break;
    }

    return CallWindowProc(g_pOldWndProc, hWnd, msg, wp, lp);
}

BOOL PopupExplorerMenu(HWND hWnd, LPWSTR pwszPath, POINT point)
{
    BOOL bResult = FALSE;

    LPMALLOC		pMalloc			= NULL;
    LPSHELLFOLDER	psfFolder		= NULL;
    LPSHELLFOLDER	psfNextFolder	= NULL;

    if (!SUCCEEDED(SHGetMalloc(&pMalloc)))
        return bResult;

    if (!SUCCEEDED(SHGetDesktopFolder(&psfFolder))) 
    {
        pMalloc->Release();
        return bResult;
    }

    //
    // Convert the path name into a pointer to an item ID list (pidl).
    //

    LPITEMIDLIST	pidlMain;
    ULONG			ulCount;
    ULONG			ulAttr;
    UINT			nCount;

    if (SUCCEEDED(psfFolder->ParseDisplayName(hWnd, NULL, pwszPath, &ulCount, &pidlMain, &ulAttr)) && (pidlMain != NULL)) 
    {
        if (nCount = GetItemCount(pidlMain)) 
        {
            //
            // Initialize psfFolder with a pointer to the IShellFolder
            // interface of the folder that contains the item whose context
            // menu we're after, and initialize pidlItem with a pointer to
            // the item's item ID. If nCount > 1, this requires us to walk
            // the list of item IDs stored in pidlMain and bind to each
            // subfolder referenced in the list.
            //

            LPITEMIDLIST pidlItem = pidlMain;

            while (--nCount) 
            {
                //
                // Create a 1-item item ID list for the next item in pidlMain.
                //

                LPITEMIDLIST pidlNextItem = DuplicateItem(pMalloc, pidlItem);

                if (pidlNextItem == NULL) 
                {
                    pMalloc->Free(pidlMain);
                    psfFolder->Release();
                    pMalloc->Release();
                    return bResult;
                }

                //
                // Bind to the folder specified in the new item ID list.
                //

                if (!SUCCEEDED(psfFolder->BindToObject(pidlNextItem, NULL, IID_IShellFolder, (void**)&psfNextFolder))) 
                {
                    pMalloc->Free(pidlNextItem);
                    pMalloc->Free(pidlMain);
                    psfFolder->Release();
                    pMalloc->Release();
                    return bResult;
                }

                //
                // Release the IShellFolder pointer to the parent folder
                // and set psfFolder equal to the IShellFolder pointer for
                // the current folder.
                //

                psfFolder->Release();
                psfFolder = psfNextFolder;

                //
                // Release the storage for the 1-item item ID list we created
                // just a moment ago and initialize pidlItem so that it points
                // to the next item in pidlMain.
                //

                pMalloc->Free(pidlNextItem);
                pidlItem = GetNextItem(pidlItem);
            }

            LPITEMIDLIST* ppidl = &pidlItem;

            //
            // Get a pointer to the item's IContextMenu interface and call
            // IContextMenu::QueryContextMenu to initialize a context menu.
            //

            LPCONTEXTMENU pContextMenu = NULL;

            if (SUCCEEDED(psfFolder->GetUIObjectOf(hWnd, 1, (LPCITEMIDLIST*)(ppidl), IID_IContextMenu, NULL, (void**)&pContextMenu))) 
            {
                // 
                // Try to see if we can upgrade to an IContextMenu2/3 interface. If so we need 
                // to hook into the message handler to invoke IContextMenu2::HandleMenuMsg().
                //

                bool hook = false;

                void *pInterface = NULL;

                if (pContextMenu->QueryInterface(IID_IContextMenu3, &pInterface) == NOERROR)
                {
                    pContextMenu->Release();
                    pContextMenu = (LPCONTEXTMENU)pInterface;
                    hook = true;
                }
                else if (pContextMenu->QueryInterface(IID_IContextMenu2, &pInterface) == NOERROR)
                {
                    pContextMenu->Release();
                    pContextMenu = (LPCONTEXTMENU)pInterface;
                    hook = true;
                }

                const int MIN_ID = 1;
                const int MAX_ID = 0x7FFF;

                HMENU hMenu = ::CreatePopupMenu();

                if (SUCCEEDED(pContextMenu->QueryContextMenu(hMenu, 0, MIN_ID, MAX_ID, CMF_EXPLORE))) 
                {
                    ::ClientToScreen(hWnd, &point);

                    if (hook) 
                    {
                        g_pOldWndProc   = (WNDPROC)::SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)HookWndProc);
                        g_pIContext2or3 = (LPCONTEXTMENU2)pContextMenu;
                    }
                    else 
                    {
                        g_pOldWndProc   = 0;
                        g_pIContext2or3 = NULL;
                    }

                    //
                    // Display the context menu.
                    //

                    UINT nCmd = ::TrackPopupMenu(
                        hMenu, 
                        TPM_LEFTALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON | TPM_RETURNCMD,
                        point.x, 
                        point.y, 
                        0, 
                        hWnd, 
                        NULL);

                    //
                    // Restore our hook into the message handler.
                    //

                    if (g_pOldWndProc) 
                    {
                        ::SetWindowLongPtr(hWnd, GWLP_WNDPROC, (LONG_PTR)g_pOldWndProc);
                    }

                    //
                    // If a command was selected from the menu, execute it.
                    //

                    if (nCmd) 
                    {
                        CMINVOKECOMMANDINFO ici;

                        ici.cbSize          = sizeof(CMINVOKECOMMANDINFO);
                        ici.fMask           = 0;
                        ici.hwnd            = hWnd;
                        ici.lpVerb          = MAKEINTRESOURCEA(nCmd - MIN_ID);
                        ici.lpParameters    = NULL;
                        ici.lpDirectory     = NULL;
                        ici.nShow           = SW_SHOWNORMAL;
                        ici.dwHotKey        = 0;
                        ici.hIcon           = NULL;

                        if (SUCCEEDED(pContextMenu->InvokeCommand(&ici)))
                        {
                            bResult = TRUE;
                        }
                    }
                }

                ::DestroyMenu(hMenu);

                pContextMenu->Release();
            }
        }

        pMalloc->Free(pidlMain);
    }

    psfFolder->Release();
    pMalloc->Release();

    return bResult;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Enum process child windows (property dialogs).
///////////////////////////////////////////////////////////////////////////////////////////////////

BOOL IsPropertyDialog(HWND hWnd)
{
    wchar_t wszClassName[256] = {0};

    ::GetClassName(hWnd, wszClassName, _countof(wszClassName));

    // Property dialogs use the special "#32770 (Dialog)" class.

    return !wcscmp(wszClassName, L"#32770");
}

BOOL CALLBACK EnumWindowCallback(HWND hWnd, LPARAM lParam)
{
    DWORD dwPID;

    ::GetWindowThreadProcessId(hWnd, &dwPID);

    if (dwPID == ::GetCurrentProcessId())
    {	
        BOOL* pbFoundChild = (BOOL*)(lParam);

        if (IsPropertyDialog(hWnd))
        {
            *pbFoundChild = TRUE;

            return FALSE;
        }
    }

    return TRUE;
}

BOOL HaveChildDialogs()
{
    BOOL bFoundChild = FALSE;

    ::EnumWindows(EnumWindowCallback, (LPARAM)(&bFoundChild));

    return bFoundChild;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Application entrypoint.
///////////////////////////////////////////////////////////////////////////////////////////////////

const int QUIT_TIMER_EVENT = 1;
const int QUIT_TIMER_SLEEP = 200;

LRESULT CALLBACK MainWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if (uMsg == WM_TIMER && wParam == QUIT_TIMER_EVENT)
    {
        if (HaveChildDialogs())
        {
            ::SetTimer(hWnd, QUIT_TIMER_EVENT, QUIT_TIMER_SLEEP, NULL);
        }
        else
        {
            ::PostQuitMessage(EXIT_SUCCESS);
        }
    }

    return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpszCmdLine, int nCmdShow)
{
    int argc = 0;

    LPWSTR *argv = ::CommandLineToArgvW(::GetCommandLine(), &argc);

    //
    // Don't bother if we weren't passed a valid file or folder argument.
    //

    if (argc < 2)
    {
        return EXIT_FAILURE;
    }

    //
    // We can't use a hidden window since that would break our popup menus.
    //

    static wchar_t szWndClassName[] = L"9534921E-82FD-4d3a-B073-FDE61CDEFE25";
    static wchar_t szWndTitleName[] = L"AF150A73-1756-49f0-961A-1D13C4DD13D0";

    WNDCLASS wc;
    wc.style			= 0;
    wc.lpfnWndProc		= MainWndProc;
    wc.cbClsExtra		= 0;
    wc.cbWndExtra		= 0;
    wc.hInstance		= hInstance;
    wc.hIcon			= ::LoadIcon(NULL, IDI_WINLOGO);
    wc.hCursor			= ::LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground	= (HBRUSH)::GetStockObject(BLACK_BRUSH);
    wc.lpszMenuName		= NULL;
    wc.lpszClassName	= szWndClassName;

    ::RegisterClass(&wc);

    HWND hWnd = ::CreateWindow(
        szWndClassName, 
        szWndTitleName,
        WS_POPUP, 
        0, 0,
        1, 1,
        HWND_DESKTOP, 
        NULL, 
        hInstance, 
        NULL);

    ::ShowWindow(hWnd, SW_SHOW);
    ::UpdateWindow(hWnd);

    POINT pt;
    ::GetCursorPos(&pt);

    PopupExplorerMenu(hWnd, argv[1], pt);

    //
    // Property dialogs are owned by this process so we have to wait until 
    // all child windows are destroyed before we can safely exit. We'll use
    // timer to keep the message pump while we wait for the user to finish 
    // interacting with the property dialogs.
    //

    ::SetTimer(hWnd, QUIT_TIMER_EVENT, QUIT_TIMER_SLEEP, NULL);

    MSG msg;

    while (::GetMessage(&msg, NULL, 0, 0)) 
    {
        ::TranslateMessage(&msg);
        ::DispatchMessage(&msg);
    }

    return EXIT_SUCCESS;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// The End.
///////////////////////////////////////////////////////////////////////////////////////////////////
