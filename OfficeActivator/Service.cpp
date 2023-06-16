#include "pch.h"
#include "framework.h"
#include <winsvc.h>
#include "Service.h"
#include "PeFile.h"
#include "Helps.h"
#include "Mutex.h"
#pragma comment(lib, "advapi32.lib")

SERVICE_STATUS          gSvcStatus;
SERVICE_STATUS_HANDLE   gSvcStatusHandle;
HANDLE                  ghSvcStopEvent;
HANDLE                  ghThread;
HANDLE                  ghMutex;
LPCTSTR                 glpMSO;

VOID SvcWriteLog(LPTSTR szModule, LPTSTR szText) {
    UNREFERENCED_PARAMETER(szModule);
    UNREFERENCED_PARAMETER(szText);
}

VOID PatchMSO() {
    WaitForSingleObject(ghMutex, INFINITE);

    if (IsMsoPatched(glpMSO) == 0 && VerifyEmbeddedSignature(glpMSO)) {
        CString msoPath = CString(glpMSO) + _T(".bak");
        DeleteFile(msoPath);

        if (MoveFileW(glpMSO, msoPath)) {
            PatchMsoFile(msoPath, glpMSO);
        }
    }

    ReleaseMutex(ghMutex);
}

DWORD WINAPI ThreadProc(PVOID) {

#ifdef _DEBUG
    while (!IsDebuggerPresent())Sleep(1000);
#endif

    CString msop(glpMSO);
    int index = msop.ReverseFind(_T('\\'));
    if (index == -1) {
        return -1;
    }
    msop = msop.Left(index);

    PatchMSO();

    HANDLE handles[] = {
        FindFirstChangeNotification(msop, FALSE, FILE_NOTIFY_CHANGE_SIZE | FILE_NOTIFY_CHANGE_LAST_WRITE),
        ghSvcStopEvent
    };

    if (INVALID_HANDLE_VALUE == handles[0] || !handles[0]) {
        SvcWriteLog(TEXT("ThreadProc"), TEXT("FindFirstChangeNotification failed"));
        return GetLastError();
    }

    while (1) {
        switch (WaitForMultipleObjects(sizeof(handles) / sizeof(HANDLE), handles, FALSE, INFINITE)) {
        case WAIT_OBJECT_0:
            
            if (!FindNextChangeNotification(handles[0])) {
                SvcWriteLog(TEXT("ThreadProc"), TEXT("FindNextChangeNotification failed"));
                FindCloseChangeNotification(handles[0]);
                return GetLastError();
            }

            PatchMSO();

            break;

        case WAIT_OBJECT_0 + 1:
            FindCloseChangeNotification(handles[0]);
            return 0;

        default:
            SvcWriteLog(TEXT("ThreadProc"), TEXT("WaitForMultipleObjects failed"));
            return GetLastError();
        }
    }
}

VOID ReportSvcStatus(
    DWORD dwCurrentState,
    DWORD dwWin32ExitCode,
    DWORD dwWaitHint) {
    static DWORD dwCheckPoint = 1;

    gSvcStatus.dwCurrentState = dwCurrentState;
    gSvcStatus.dwWin32ExitCode = dwWin32ExitCode;
    gSvcStatus.dwWaitHint = dwWaitHint;

    if (dwCurrentState == SERVICE_START_PENDING)
        gSvcStatus.dwControlsAccepted = 0;
    else gSvcStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP;

    if ((dwCurrentState == SERVICE_RUNNING) ||
        (dwCurrentState == SERVICE_STOPPED))
        gSvcStatus.dwCheckPoint = 0;
    else gSvcStatus.dwCheckPoint = dwCheckPoint++;

    SetServiceStatus(gSvcStatusHandle, &gSvcStatus);
}

VOID WINAPI SvcCtrlHandler(DWORD dwCtrl)
{
    // Handle the requested control code. 

    switch (dwCtrl)
    {
    case SERVICE_CONTROL_STOP:
        ReportSvcStatus(SERVICE_STOP_PENDING, NO_ERROR, 0);

        // Signal the service to stop.

        SetEvent(ghSvcStopEvent);
        ReportSvcStatus(gSvcStatus.dwCurrentState, NO_ERROR, 0);

        return;

    case SERVICE_CONTROL_INTERROGATE:
        break;

    default:
        break;
    }

}

VOID WINAPI SvcMain(DWORD dwArgc, LPTSTR* lpszArgv) {

    int argc;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);

    do {
        gSvcStatusHandle = RegisterServiceCtrlHandler(SVCNAME, SvcCtrlHandler);
        if (!gSvcStatusHandle) {
            SvcWriteLog(TEXT("SvcMain"), TEXT("RegisterServiceCtrlHandler failed"));
            break;
        }

        gSvcStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
        gSvcStatus.dwServiceSpecificExitCode = 0;

        ReportSvcStatus(SERVICE_START_PENDING, NO_ERROR, 3000);

        if (argc != 3 || _tcscmp(_T("-mso"), argv[1])) {
            ReportSvcStatus(SERVICE_STOPPED, ERROR_INVALID_PARAMETER, 0);
            break;
        }

        ghSvcStopEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
        if (ghSvcStopEvent == NULL) {
            ReportSvcStatus(SERVICE_STOPPED, GetLastError(), 0);
            break;
        }

        ghMutex = CreateGlobalMutex();
        if (ghMutex == NULL) {
            ReportSvcStatus(SERVICE_STOPPED, GetLastError(), 0);
            break;
        }

        glpMSO = argv[2];
        ghThread = CreateThread(NULL, 0, ThreadProc, NULL, 0, NULL);
        if (ghThread == NULL) {
            ReportSvcStatus(SERVICE_STOPPED, GetLastError(), 0);
            break;
        }

        ReportSvcStatus(SERVICE_RUNNING, NO_ERROR, 0);
        WaitForSingleObject(ghThread, INFINITE);
        ReportSvcStatus(SERVICE_STOPPED, NO_ERROR, 0);
    } while (false);

}

BOOL WINAPI SvcRun() {
    SERVICE_TABLE_ENTRY DispatchTable[] =
    {
        { SVCNAME, SvcMain },
        { NULL, NULL }
    };

    return StartServiceCtrlDispatcher(DispatchTable);
}
