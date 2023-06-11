#pragma once

DECLARE_EXTERN_ORIGIN(SLGetLicensingStatusInformation);
DECLARE_EXTERN_ORIGIN(SLGetApplicationPolicy);
DECLARE_EXTERN_ORIGIN(SLGetSLIDList);
DECLARE_EXTERN_ORIGIN(SLGetPolicyInformation);
DECLARE_EXTERN_ORIGIN(SLGetProductSkuInformation);
DECLARE_EXTERN_ORIGIN(SLSetAuthenticationData);
DECLARE_EXTERN_ORIGIN(SLGetAuthenticationResult);

HRESULT WINAPI HookSLGetLicensingStatusInformation(
    _In_ HSLC hSLC,
    _In_opt_ CONST SLID* pAppID,
    _In_opt_ CONST SLID* pProductSkuId,
    _In_opt_ PCWSTR pwszRightName,
    _Out_ UINT* pnStatusCount,
    _Outptr_result_buffer_(*pnStatusCount) SL_LICENSING_STATUS** ppLicensingStatus
);

HRESULT WINAPI HookSLGetApplicationPolicy(
    _In_ HSLP hPolicyContext,
    _In_ PCWSTR pwszValueName,
    _Out_opt_ SLDATATYPE* peDataType,
    _Out_ UINT* pcbValue,
    _Outptr_result_bytebuffer_(*pcbValue) PBYTE* ppbValue
);

HRESULT WINAPI HookSLGetPolicyInformation(
    _In_ HSLC hSLC,
    _In_ PCWSTR pwszValueName,
    _Out_opt_ SLDATATYPE* peDataType,
    _Out_ UINT* pcbValue,
    _Outptr_result_bytebuffer_(*pcbValue) PBYTE* ppbValue
);

HRESULT WINAPI HookSLGetProductSkuInformation(
    _In_ HSLC hSLC,
    _In_ CONST SLID* pProductSkuId,
    _In_ PCWSTR pwszValueName,
    _Out_opt_ SLDATATYPE* peDataType,
    _Out_ UINT* pcbValue,
    _Outptr_result_bytebuffer_(*pcbValue) PBYTE* ppbValue
);

HRESULT WINAPI HookSLGetSLIDList(
    _In_ HSLC hSLC,
    _In_ SLIDTYPE eQueryIdType,
    _In_opt_ CONST SLID* pQueryId,
    _In_ SLIDTYPE eReturnIdType,
    _Out_ UINT* pnReturnIds,
    _Outptr_result_buffer_(*pnReturnIds) SLID** ppReturnIds
);

HRESULT WINAPI HookSLSetAuthenticationData(
    _In_ HSLC hSLC,
    _In_opt_ UINT cbValue,
    _In_reads_bytes_opt_(cbValue) CONST BYTE* pbValue
);

HRESULT WINAPI HookSLGetAuthenticationResult(
    _In_ HSLC hSLC,
    _Out_ UINT* pcbValue,
    _Outptr_result_bytebuffer_(*pcbValue) PBYTE* ppbValue
);
