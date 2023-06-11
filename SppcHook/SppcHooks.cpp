#include "stdafx.h"
#pragma comment(lib,"slc.lib")
#pragma comment(lib,"Rpcrt4.lib")

DECLARE_ORIGIN(SLGetLicensingStatusInformation);
DECLARE_ORIGIN(SLGetApplicationPolicy);
DECLARE_ORIGIN(SLGetSLIDList);
DECLARE_ORIGIN(SLGetPolicyInformation);
DECLARE_ORIGIN(SLGetProductSkuInformation);
DECLARE_ORIGIN(SLSetAuthenticationData);
DECLARE_ORIGIN(SLGetAuthenticationResult);

//#ifndef _WIN64
//#pragma warning(disable:4273)
//SIZE_T NTAPI RtlCompareMemory(
//    _In_ const VOID* Source1,
//    _In_ const VOID* Source2,
//    _In_ SIZE_T Length) {
//    return decltype(&RtlCompareMemory)(GetProcAddress(GetModuleHandleW(L"ntdll.dll"), "RtlCompareMemory"))(Source1, Source2, Length);
//}
//#pragma warning(default:4273)
//#endif

#ifndef HELPERS
VOID InfoHelper(
    _In_ HRESULT hr,
    _In_ PCWSTR pwszValueName,
    _Out_opt_ SLDATATYPE* peDataType,
    _Out_ UINT* pcbValue,
    _Outptr_result_bytebuffer_(*pcbValue) PBYTE* ppbValue) {
    printf("%08X\n", hr);

    if (hr == ERROR_SUCCESS) {
        SLDATATYPE local = SLDATATYPE::SL_DATA_NONE;
        if (!peDataType) peDataType = &local;

        switch (*peDataType) {
        case SLDATATYPE::SL_DATA_DWORD:
            printf("\t%d\n", *PDWORD(*ppbValue));
            break;

        case SLDATATYPE::SL_DATA_SZ:
            printf("\t%ls\n", PCWSTR(*ppbValue));
            break;

        case SLDATATYPE::SL_DATA_MULTI_SZ:
        {
            PCWSTR str = (PCWSTR)*ppbValue;
            while (true) {
                int count = printf("\t%ls\n", str);
                if (count == 2)break;
                str += count - 1;
            }

            break;
        }

        default:
        {
            PBYTE ptr = *ppbValue;
            putchar('\t');
            for (DWORD i = 0; i < *pcbValue; ++i) {
                printf("%02X", (unsigned int)ptr[i]);
                if (0 == i + 1 % 10)printf("\n\t");
            }
            putchar('\n');

            break;
        }
        }
    }
}

BOOL PolicyValueModified(
    _Inout_ HRESULT* hr,
    _In_ PCWSTR pwszValueName,
    _Inout_opt_ SLDATATYPE* peDataType,
    _Inout_ UINT* pcbValue,
    _Inout_updates_bytes_(*pcbValue) PBYTE* ppbValue) {
    BOOL result = FALSE;

    //if (wcscmp(L"office-3CFF5AB2-9B16-4A31-BC3F-FAD761D92780", pwszValueName) == 0) {
    //    if (*hr == ERROR_SUCCESS) {
    //        if (*PDWORD(*ppbValue) != 1) {
    //            *PDWORD(*ppbValue) = 1;
    //            result = TRUE;
    //        }
    //    }
    //    else {
    //        *hr = ERROR_SUCCESS;
    //        if (peDataType) {
    //            *peDataType = SLDATATYPE::SL_DATA_DWORD;
    //            *pcbValue = sizeof(DWORD);
    //            *ppbValue = (PBYTE)LocalAlloc(0, sizeof(DWORD));
    //            assert(*ppbValue);

    //            *PDWORD(*ppbValue) = 1;
    //            result = TRUE;
    //        }
    //    }
    //}

    if (result) printf("%ls policy modified.\n", pwszValueName);

    return result;
}

PSPPC_AUTH_RESULT PolicyCreateAuthHashLockHeld(
    _In_ HCRYPTKEY hKey,
    _In_ PCWSTR pwszValueName,
    _In_ UINT pcbValue,
    _In_reads_bytes_(pcbValue) PBYTE ppbValue) {
    assert(pcbValue == sizeof(DWORD));

    DWORD len = sizeof(WCHAR) * (DWORD)wcslen(pwszValueName) + sizeof(DWORD) * 3;
    auto hashData = (PSPPC_AUTH_HASH)LocalAlloc(0, len);
    assert(hashData);

    hashData->dwUnknown10000 = 0x10000;
    hashData->dwDataSize = sizeof(DWORD);
    hashData->dwDataValue = *PDWORD(ppbValue);
    RtlCopyMemory(hashData->PolicyName, pwszValueName, len - 3 * sizeof(DWORD));

    HCRYPTHASH hh;
    HMAC_INFO hi{};

    hi.HashAlgid = CALG_SHA;
    assert(CryptCreateHash(g_CryptProv, CALG_HMAC, hKey, 0, &hh));
    assert(CryptSetHashParam(hh, HP_HMAC_INFO, (PBYTE)&hi, 0));
    assert(CryptHashData(hh, (PBYTE)hashData, len, 0));

    DWORD hashLen = 0;
    PBYTE hash = nullptr;
    assert(CryptGetHashParam(hh, HP_HASHVAL, hash, &hashLen, 0));
    hash = (PBYTE)LocalAlloc(0, hashLen);
    assert(hash);
    assert(CryptGetHashParam(hh, HP_HASHVAL, hash, &hashLen, 0));
    CryptDestroyHash(hh);

    DWORD resultLen = hashLen + sizeof(DWORD) * 4;
    PSPPC_AUTH_RESULT result = (PSPPC_AUTH_RESULT)LocalAlloc(LMEM_ZEROINIT, resultLen);
    assert(result);

    result->dwSize = resultLen;
    result->dwUnknown2 = 2;
    result->dwHashSize = hashLen;
    RtlCopyMemory(result->Hashs, hash, hashLen);

    LocalFree(hash);
    LocalFree(hashData);
    return result;
}

VOID LookupPolicy(_In_ HSLC hSLC, _In_ HCRYPTKEY hk) {
    UINT u1, u0;
    PBYTE p1 = nullptr, p0 = nullptr;
    SLID pid;

    CLSIDFromString(L"{0ff1ce15-a989-479d-af46-f275c6370663}", &pid);

    HRESULT hr = SLConsumeRight(hSLC, &pid, nullptr, nullptr, nullptr);
    if (hr == ERROR_SUCCESS) {
        printf("SLConsumeRight success\n");
    }
    else {
        printf("SLConsumeRight fail 0x%08X\n", hr);
    }

    auto name = L"office-DC5CCACD-A7AC-4FD3-9F70-9454B5DE5161";

    HRESULT hr0 = OriginSLGetPolicyInformation(hSLC, name, nullptr, &u0, &p0);
    if (hr0 == ERROR_SUCCESS) {
        printf("query success:%d\n", *PDWORD(p0));
    }
    else {
        printf("query fail 0x%08X\n", hr0);
        assert(false);
    }

    HRESULT hr1 = OriginSLGetAuthenticationResult(hSLC, &u1, &p1);
    if (ERROR_SUCCESS == hr1) {
        printf("get result success.\n");
    }
    else {
        printf("Get result failed.0x%08X\n", hr1);
        assert(false);
    }

    auto ar = (PSPPC_AUTH_RESULT)p1;
    auto result = PolicyCreateAuthHashLockHeld(hk, name, u0, p0);

    if ((u1 == result->dwSize) && (RtlCompareMemory(result->Hashs, ar->Hashs, result->dwHashSize) == result->dwHashSize)) {
        printf("compare success.\n");
    }
    else {
        printf("compare fail:u1=%d,r=%d.\n", u1, result->dwSize);
        assert(false);
    }

    LocalFree(p0);
    LocalFree(p1);
    LocalFree(result);
}
#endif

HRESULT WINAPI HookSLGetLicensingStatusInformation(
    _In_ HSLC hSLC,
    _In_opt_ CONST SLID* pAppID,
    _In_opt_ CONST SLID* pProductSkuId,
    _In_opt_ PCWSTR pwszRightName,
    _Out_ UINT* pnStatusCount,
    _Outptr_result_buffer_(*pnStatusCount) SL_LICENSING_STATUS** ppLicensingStatus) {
    HRESULT hr = OriginSLGetLicensingStatusInformation(
        hSLC, pAppID, pProductSkuId, pwszRightName, pnStatusCount, ppLicensingStatus
    );

    printf("HookSLGetLicensingStatusInformation\n");

    if (hr == ERROR_SUCCESS) {
        printf("%ls\t%d\n", pwszRightName, *pnStatusCount);

        for (UINT i = *pnStatusCount; i > 0; --i) {
            SL_LICENSING_STATUS* pp = *ppLicensingStatus;
            printf("%d\t%d\t%d\t%08x\t%lld\n\n", pp[i - 1].eStatus, pp[i - 1].dwGraceTime, pp[i - 1].dwTotalGraceDays, pp[i - 1].hrReason, pp[i - 1].qwValidityExpiration);

            if (pp[i - 1].hrReason != SL_E_PKEY_NOT_INSTALLED) {
                pp[i - 1].eStatus = SLLICENSINGSTATUS::SL_LICENSING_STATUS_LICENSED;
                pp[i - 1].dwGraceTime = 259200;
                pp[i - 1].dwTotalGraceDays = 0;
                pp[i - 1].hrReason = ERROR_SUCCESS;
                pp[i - 1].qwValidityExpiration = 0;
            }
        }
    }

    return hr;
}

HRESULT WINAPI HookSLGetApplicationPolicy(
    _In_ HSLP hPolicyContext,
    _In_ PCWSTR pwszValueName,
    _Out_opt_ SLDATATYPE* peDataType,
    _Out_ UINT* pcbValue,
    _Outptr_result_bytebuffer_(*pcbValue) PBYTE* ppbValue) {

    printf("**HookSLGetApplicationPolicy\t%ls\t", pwszValueName);
    HRESULT hr = OriginSLGetApplicationPolicy(
        hPolicyContext,
        pwszValueName,
        peDataType,
        pcbValue,
        ppbValue
    );

    InfoHelper(hr, pwszValueName, peDataType, pcbValue, ppbValue);

    return hr;
}

HRESULT WINAPI HookSLGetPolicyInformation(
    _In_ HSLC hSLC,
    _In_ PCWSTR pwszValueName,
    _Out_opt_ SLDATATYPE* peDataType,
    _Out_ UINT* pcbValue,
    _Outptr_result_bytebuffer_(*pcbValue) PBYTE* ppbValue) {

    printf("**HookSLGetPolicyInformation\t%ls\t", pwszValueName);
    HRESULT hr = OriginSLGetPolicyInformation(
        hSLC,
        pwszValueName,
        peDataType,
        pcbValue,
        ppbValue
    );

    InfoHelper(hr, pwszValueName, peDataType, pcbValue, ppbValue);

    EnterCriticalSection(&SlcAuthMapLock);
    auto item = SlcAuthMap.find(hSLC);
    if (item != SlcAuthMap.end()) {
        if (item->second.second != nullptr) {
            LocalFree(item->second.second);
            item->second.second = nullptr;
        }

        if (PolicyValueModified(&hr, pwszValueName, peDataType, pcbValue, ppbValue) && hr == ERROR_SUCCESS) {
            item->second.second = PolicyCreateAuthHashLockHeld(item->second.first, pwszValueName, *pcbValue, *ppbValue);
        }
    }
    LeaveCriticalSection(&SlcAuthMapLock);

    return hr;
}

HRESULT WINAPI HookSLGetProductSkuInformation(
    _In_ HSLC hSLC,
    _In_ CONST SLID* pProductSkuId,
    _In_ PCWSTR pwszValueName,
    _Out_opt_ SLDATATYPE* peDataType,
    _Out_ UINT* pcbValue,
    _Outptr_result_bytebuffer_(*pcbValue) PBYTE* ppbValue) {

    printf("**HookSLGetProductSkuInformation\t%ls\t", pwszValueName);
    HRESULT hr = OriginSLGetProductSkuInformation(
        hSLC,
        pProductSkuId,
        pwszValueName,
        peDataType,
        pcbValue,
        ppbValue
    );

    InfoHelper(hr, pwszValueName, peDataType, pcbValue, ppbValue);

    return hr;
}

HRESULT WINAPI HookSLGetSLIDList(
    _In_ HSLC hSLC,
    _In_ SLIDTYPE eQueryIdType,
    _In_opt_ CONST SLID* pQueryId,
    _In_ SLIDTYPE eReturnIdType,
    _Out_ UINT* pnReturnIds,
    _Outptr_result_buffer_(*pnReturnIds) SLID** ppReturnIds) {

    HRESULT hr = OriginSLGetSLIDList(
        hSLC,
        eQueryIdType,
        pQueryId,
        eReturnIdType,
        pnReturnIds,
        ppReturnIds
    );

    RPC_CSTR qid = nullptr;
    if (pQueryId) {
        UuidToStringA(pQueryId, &qid);
    }

    printf("%d\t%s\t%d\n", eQueryIdType, qid, *pnReturnIds);
    if (qid)RpcStringFreeA(&qid);

    if (hr == ERROR_SUCCESS) {

        for (DWORD i = 0; i < *pnReturnIds; ++i) {
            RPC_CSTR rid;
            UuidToStringA(ppReturnIds[i], &rid);
            printf("\t%s\n", rid);
            RpcStringFreeA(&rid);
        }

        if (*pnReturnIds == 1) {
            CLSIDFromString(L"59c3e7f6-69f7-a0b9-5637-62e4eadbe59f", ppReturnIds[0]);
        }
    }

    return hr;
}

HRESULT WINAPI HookSLSetAuthenticationData(
    _In_ HSLC hSLC,
    _In_opt_ UINT cbValue,
    _In_reads_bytes_opt_(cbValue) CONST BYTE* pbValue) {

    HRESULT hr = SL_E_NOT_SUPPORTED;

    if (cbValue && pbValue) {
        auto data = PSPPC_AUTH_DATA(pbValue);
        if (data->dwSize == cbValue &&
            data->dwUnknown1 == 1 &&
            data->dwUnknown10000 == 0x10000 &&
            data->dwKeySize > 0 &&
            data->dwKeySize < data->dwSize - 0x10) {

            printf("HookSLSetAuthenticationData\n");

            EnterCriticalSection(&SlcAuthMapLock);

            HCRYPTKEY hk;
            if (CryptImportKey(g_CryptProv, data->pbKeys, data->dwKeySize, g_ProxyPvk, CRYPT_EXPORTABLE, &hk)) {

                cbValue = 0;
                pbValue = nullptr;
                CryptExportKey(hk, g_SppsvcPubKey, SIMPLEBLOB, 0, (PBYTE)pbValue, (PDWORD)&cbValue);
                if (cbValue != 0) {
                    pbValue = (PBYTE)LocalAlloc(0, cbValue + 0x14);
                    assert(pbValue);

                    data = (PSPPC_AUTH_DATA)pbValue;
                    data->dwSize = cbValue + 0x14;
                    data->dwKeySize = cbValue;
                    data->dwUnknown1 = 1;
                    data->dwUnknown10000 = 0x10000;

                    if (CryptExportKey(hk, g_SppsvcPubKey, SIMPLEBLOB, 0, data->pbKeys, &data->dwKeySize)) {
                        hr = OriginSLSetAuthenticationData(
                            hSLC,
                            cbValue + 0x14,
                            pbValue
                        );
                        if (hr == ERROR_SUCCESS) {

                            //LookupPolicy(hSLC, hk);

                            printf("all success\n");
                        }
                        else {
                            printf("4 fail 0x%08X\n", hr);
                        }
                    }
                    else {
                        printf("3 fail 0x%08X\n", GetLastError());
                    }

                    LocalFree((PVOID)pbValue);
                }
                else {
                    printf("2 fail 0x%08X\n", GetLastError());
                }

                if (hr == ERROR_SUCCESS) {
                    auto pair = std::make_pair(hk, nullptr);
                    auto item = SlcAuthMap.find(hSLC);
                    if (item != SlcAuthMap.end()) {
                        CryptDestroyKey(item->second.first);
                        item->second = pair;
                    }
                    else {
                        SlcAuthMap.insert(std::make_pair(hSLC, pair));
                    }
                }
                else {
                    CryptDestroyKey(hk);
                }

            }
            else {
                printf("1 fail 0x%08X\n", GetLastError());
            }

            LeaveCriticalSection(&SlcAuthMapLock);
        }
    }

    return hr;
}

HRESULT WINAPI HookSLGetAuthenticationResult(
    _In_ HSLC hSLC,
    _Out_ UINT* pcbValue,
    _Outptr_result_bytebuffer_(*pcbValue) PBYTE* ppbValue) {

    HRESULT hr;

    EnterCriticalSection(&SlcAuthMapLock);

    auto item = SlcAuthMap.find(hSLC);
    if (item != SlcAuthMap.end() && item->second.second != nullptr) {
        auto result = item->second.second;
        *pcbValue = result->dwSize;
        *ppbValue = (PBYTE)LocalAlloc(0, result->dwSize);
        assert(*ppbValue);

        RtlCopyMemory(
            *ppbValue,
            result,
            result->dwSize
        );

        hr = ERROR_SUCCESS;
    }
    else {
        hr = OriginSLGetAuthenticationResult(
            hSLC,
            pcbValue,
            ppbValue
        );
    }

    LeaveCriticalSection(&SlcAuthMapLock);

    return hr;
}
