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
tm                      gLogTime;
FILE*                   gLogFile;

#ifdef UNICODE
#define _tftime wcsftime
#else
#define _tftime strftime
#endif

BOOL InitializeLog() {
    CString app;
    AfxGetModuleFileName(AfxGetInstanceHandle(), app);

    int index = app.ReverseFind(_T('\\'));
    if (index == -1) return FALSE;

    app = app.Left(index) + _T("\\logs");

    time_t t = time(nullptr);
    if (0 != localtime_s(&gLogTime, &t))return FALSE;

    TCHAR buffer[26];
    _tftime(buffer, 26, _T("\\log_%Y_%m_%d.txt"), &gLogTime);

    if (!PathFileExists(app) && !CreateDirectory(app, nullptr)) {
        return FALSE;
    }

    app += buffer;

    if (gLogFile) fclose(gLogFile);
    gLogFile = _tfsopen(app, _T("a"), SH_DENYWR);
    return 0 != gLogFile;
}

VOID SvcWriteLog(LPCTSTR szModule, LPCTSTR szText) {
    time_t timer;
    TCHAR buffer[26];
    tm tm_info;

    timer = time(NULL);
    localtime_s(&tm_info, &timer);

    if (gLogTime.tm_year != tm_info.tm_year ||
        gLogTime.tm_mon != tm_info.tm_mon ||
        gLogTime.tm_mday != tm_info.tm_mday) {
        if (!InitializeLog()) {
            ASSERT(FALSE);
        }
    }

    _tftime(buffer, 26, _T("%Y-%m-%d %H:%M:%S"), &tm_info);
    _ftprintf(gLogFile, _T("[%s] "), buffer);
    _ftprintf(gLogFile, _T("[%s]: %s\n"), szModule, szText);
    fflush(gLogFile);
}

VOID PatchMSO() {
    WaitForSingleObject(ghMutex, INFINITE);

    switch (IsMsoPatched(glpMSO)) {
    case 0:
        if (VerifyEmbeddedSignature(glpMSO)) {
            CString msoPath = CString(glpMSO) + _T(".bak");
            DeleteFile(msoPath);

            if (MoveFileW(glpMSO, msoPath)) {
                switch (PatchMsoFile(msoPath, glpMSO)) {
                case 0:
                    SvcWriteLog(_T("PatchMSO"), _T("The patch was applied successfully."));
                    break;

                case 1:
                    SvcWriteLog(_T("PatchMSO"), _T("Patch failed: Cannot read source file."));
                    break;

                case 2:
                    SvcWriteLog(_T("PatchMSO"), _T("Patch failed: Cannot create new file."));
                    break;

                case 3:
                    SvcWriteLog(_T("PatchMSO"), _T("Patch failed: Invalid file format."));
                    break;

                case 4:
                case 5:
                    SvcWriteLog(_T("PatchMSO"), _T("Patch failed: Unsupported file."));
                    break;

                case 6:
                    SvcWriteLog(_T("PatchMSO"), _T("Patch failed: Insufficient memory resource."));
                    break;

                case 7:
                    SvcWriteLog(_T("PatchMSO"), _T("Patch failed: Write file failed."));
                    break;

                case -1:
                    SvcWriteLog(_T("PatchMSO"), _T("Patch failed."));
                    break;
                }
            }
            else {
                CString msg;
                msg.Format(_T("Backup MSO.DLL failed: %d."), GetLastError());

                SvcWriteLog(_T("PatchMSO"), msg);
            }
        }
        else {
            SvcWriteLog(_T("PatchMSO"), _T("The status of MSO.DLL is invalid (Signature)."));
        }
        break;

    case 1:
        SvcWriteLog(_T("PatchMSO"), _T("MSO.DLL has been patched."));
        break;

    case -1:
        SvcWriteLog(_T("PatchMSO"), _T("The status of MSO.DLL is invalid."));
        break;
    }

    ReleaseMutex(ghMutex);
}

DWORD WINAPI ThreadProc(PVOID) {
    CString msop(glpMSO);
    int index = msop.ReverseFind(_T('\\'));
    if (index == -1) {
        return -1;
    }
    msop = msop.Left(index);

    PatchMSO();

    HANDLE handles[] = {
        FindFirstChangeNotification(msop, FALSE, FILE_NOTIFY_CHANGE_SIZE | FILE_NOTIFY_CHANGE_LAST_WRITE | FILE_NOTIFY_CHANGE_FILE_NAME),
        ghSvcStopEvent
    };

    if (INVALID_HANDLE_VALUE == handles[0] || !handles[0]) {
        SvcWriteLog(_T("ThreadProc"), _T("FindFirstChangeNotification failed"));
        return GetLastError();
    }

    while (1) {
        SvcWriteLog(_T("ThreadProc"), _T("Wait for notifications..."));

        switch (WaitForMultipleObjects(sizeof(handles) / sizeof(HANDLE), handles, FALSE, INFINITE)) {
        case WAIT_OBJECT_0:
            
            if (!FindNextChangeNotification(handles[0])) {
                SvcWriteLog(_T("ThreadProc"), _T("FindNextChangeNotification failed"));
                FindCloseChangeNotification(handles[0]);
                return GetLastError();
            }

            PatchMSO();

            break;

        case WAIT_OBJECT_0 + 1:
            SvcWriteLog(_T("ThreadProc"), _T("ghSvcStopEvent"));
            FindCloseChangeNotification(handles[0]);
            return 0;

        default:
            SvcWriteLog(_T("ThreadProc"), _T("WaitForMultipleObjects failed"));
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

    gSvcStatusHandle = RegisterServiceCtrlHandler(SVCNAME, SvcCtrlHandler);
    if (!gSvcStatusHandle) {
        SvcWriteLog(_T("SvcMain"), _T("RegisterServiceCtrlHandler failed"));
        return;
    }

    int argc;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    DWORD status = NO_ERROR;

    do {
        gSvcStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
        gSvcStatus.dwServiceSpecificExitCode = 0;
        ReportSvcStatus(SERVICE_START_PENDING, NO_ERROR, 3000);

        if (argc != 3 || _tcscmp(_T("-mso"), argv[1])) {
            status = ERROR_INVALID_PARAMETER;
            break;
        }

        ghSvcStopEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
        if (ghSvcStopEvent == NULL) {
            status = GetLastError();
            break;
        }

        ghMutex = CreateGlobalMutex();
        if (ghMutex == NULL) {
            status = GetLastError();
            break;
        }

        glpMSO = argv[2];
        ghThread = CreateThread(NULL, 0, ThreadProc, NULL, 0, NULL);
        if (ghThread == NULL) {
            status = GetLastError();
            break;
        }

        ReportSvcStatus(SERVICE_RUNNING, NO_ERROR, 0);
        WaitForSingleObject(ghThread, INFINITE);

        GetExitCodeThread(ghThread, &status);
        CloseHandle(ghThread);
    } while (false);

    ReportSvcStatus(SERVICE_STOPPED, status, 0);
}

BOOL WINAPI SvcRun() {
    SERVICE_TABLE_ENTRY DispatchTable[] =
    {
        { SVCNAME, SvcMain },
        { NULL, NULL }
    };

    return StartServiceCtrlDispatcher(DispatchTable);
}
