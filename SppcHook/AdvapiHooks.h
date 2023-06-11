#pragma once

DECLARE_EXTERN_ORIGIN(CryptImportKey);

typedef struct _SPPC_AUTH_DATA {
    DWORD dwSize;
    DWORD dwUnknown1;
    DWORD dwUnknown10000;
    DWORD dwKeySize;
    BYTE pbKeys[ANYSIZE_ARRAY];
}SPPC_AUTH_DATA, * PSPPC_AUTH_DATA;

typedef struct _SPPC_AUTH_RESULT {
    DWORD dwSize;
    DWORD dwUnknown2;
    DWORD dwHashSize;
    BYTE Hashs[ANYSIZE_ARRAY];
}SPPC_AUTH_RESULT, * PSPPC_AUTH_RESULT;

typedef struct _SPPC_AUTH_HASH {
    DWORD dwUnknown10000;
    DWORD dwDataSize;
    DWORD dwDataValue;
    BYTE PolicyName[ANYSIZE_ARRAY];
}SPPC_AUTH_HASH, * PSPPC_AUTH_HASH;

extern std::map<HSLC, std::pair<HCRYPTKEY, PSPPC_AUTH_RESULT>> SlcAuthMap;
extern CRITICAL_SECTION SlcAuthMapLock;

extern HCRYPTPROV g_CryptProv;
extern HCRYPTKEY g_SppsvcPubKey;
extern HCRYPTKEY g_ProxyPvk;

extern BYTE SppsvcPubKey[276];
extern BYTE ProxyPubKey[148];
extern BYTE ProxyPvk[596];

BOOL WINAPI HookCryptImportKey(
    _In_ HCRYPTPROV  hProv,
    _In_reads_bytes_(dwDataLen)  CONST BYTE* pbData,
    _In_ DWORD dwDataLen,
    _In_ HCRYPTKEY hPubKey,
    _In_ DWORD dwFlags,
    _Out_ HCRYPTKEY* phKey
);

VOID InitMITM();
