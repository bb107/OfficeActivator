#include "pch.h"
#include "PeFile.h"
#include "Helps.h"
#pragma comment(lib,"ntdll.lib")
#pragma warning(disable:4996)

#ifdef _WIN64
#define HOST_MACHINE IMAGE_FILE_MACHINE_AMD64
#else
#define HOST_MACHINE IMAGE_FILE_MACHINE_I386
#endif

#define AlignValueUp(value, alignment) ((size_t(value) + size_t(alignment) + 1) & ~(size_t(alignment) - 1))

DWORD VA2RVA(
	_In_ PIMAGE_SECTION_HEADER section,
	_In_ DWORD va) {
	return section->PointerToRawData + (ULONG_PTR(va) - section->VirtualAddress);
}

DWORD RVA2VA(
	_In_ PIMAGE_SECTION_HEADER section,
	_In_ DWORD rva) {
	return section->VirtualAddress + (ULONG_PTR(rva) - section->PointerToRawData);
}

extern "C" {
	NTSYSAPI PIMAGE_NT_HEADERS NTAPI RtlImageNtHeader(IN PVOID ModuleAddress);
}

INT IsMsoPatched(_In_ LPCTSTR lpMsoPath) {

	INT result = -1;
	PVOID Base = MapFileReadOnly(lpMsoPath);

	if (Base) {
		PIMAGE_NT_HEADERS headers = RtlImageNtHeader(Base);
		if (headers != nullptr && headers->FileHeader.Machine == HOST_MACHINE) {
			PIMAGE_SECTION_HEADER section = IMAGE_FIRST_SECTION(headers);

			result = 0;
			for (DWORD i = 0; i < headers->FileHeader.NumberOfSections; ++i, ++section) {
				if (memcmp(section->Name, ".1hooks1", 8) == 0) {
					result = 1;
					break;
				}
			}

		}

		UnmapViewOfFile(Base);
	}

	return result;
}

INT PatchMsoFile(
	_In_ LPCTSTR lpMsoPath,
	_In_ LPCTSTR lpNewMsoPath) {

	INT result = -1;
	DWORD dwFileSize;
	PVOID OldData = MapFileReadOnly(lpMsoPath, &dwFileSize);
	PVOID NewData = nullptr;
	HANDLE hNewFile = CreateFile(
		lpNewMsoPath,
		GENERIC_WRITE,
		0,
		nullptr,
		CREATE_NEW,
		0,
		nullptr
	);

	do {
		if (!OldData) {
			result = 1;
			break;
		}

		if (hNewFile == INVALID_HANDLE_VALUE) {
			result = 2;
			break;
		}

		auto headers = RtlImageNtHeader(OldData);
		if (!headers) {
			result = 3;
			break;
		}

		PIMAGE_DATA_DIRECTORY importDir = &headers->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
		PIMAGE_DATA_DIRECTORY secDir = &headers->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_SECURITY];
		PIMAGE_SECTION_HEADER section = IMAGE_FIRST_SECTION(headers);
		PIMAGE_SECTION_HEADER importSection = nullptr;

		if (!importDir->VirtualAddress || !importDir->Size || importDir->Size % sizeof(IMAGE_IMPORT_DESCRIPTOR) != 0) {
			result = 4;
			break;
		}

		BOOL success = TRUE;
		PIMAGE_SECTION_HEADER NewSection = section + headers->FileHeader.NumberOfSections;
		ULONG_PTR NewSectionHeaderFilePtr = ULONG_PTR(NewSection) - ULONG_PTR(OldData);

		DWORD LastSectionVA = 0;
		DWORD LastSectionSize = 0;
		DWORD LastSectionPA = 0;
		DWORD LastSectionRawSize = 0;

		for (DWORD i = 0; i < headers->FileHeader.NumberOfSections; ++i, ++section) {
			if (section->VirtualAddress > LastSectionVA) {
				LastSectionVA = section->VirtualAddress;
				LastSectionSize = section->Misc.VirtualSize;
			}

			if (section->PointerToRawData > LastSectionPA) {
				LastSectionPA = section->PointerToRawData;
				LastSectionRawSize = section->SizeOfRawData;
			}

			if ((section->VirtualAddress <= importDir->VirtualAddress) && ((importDir->VirtualAddress + importDir->Size) <= (section->VirtualAddress + section->Misc.VirtualSize))) {
				ASSERT(importSection == nullptr);
				importSection = section;
			}

			if (NewSectionHeaderFilePtr + sizeof(IMAGE_SECTION_HEADER) > (ULONG_PTR)section->PointerToRawData) {
				//
				// fail
				//
				success = FALSE;
				break;
			}
		}

		if (!success || !importSection) {
			result = 5;
			break;
		}

		//
		// Allocate memory so we can change it.
		//
		DWORD NewImportSize = importDir->Size + sizeof(IMAGE_IMPORT_DESCRIPTOR) + sizeof(PVOID) * 4 + 14;//"SppcHook.dll\0"
		DWORD dwOffsetOfLastSectionEnd = LastSectionPA + LastSectionRawSize;
		DWORD dwSizeofNewSection = AlignValueUp(NewImportSize, headers->OptionalHeader.FileAlignment);
		DWORD dwSizeOfNewFile = dwFileSize - secDir->Size + dwSizeofNewSection;
		NewData = VirtualAlloc(nullptr, dwSizeOfNewFile, MEM_COMMIT, PAGE_READWRITE);
		if (!NewData) {
			result = 6;
			break;
		}

		RtlCopyMemory(NewData, OldData, dwFileSize - secDir->Size);

		headers = RtlImageNtHeader(NewData);
		importDir = &headers->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_IMPORT];
		secDir = &headers->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_SECURITY];
		importSection = PIMAGE_SECTION_HEADER(PBYTE(NewData) + (PBYTE(importSection) - PBYTE(OldData)));
		NewSection = PIMAGE_SECTION_HEADER(PBYTE(NewData) + (PBYTE(NewSection) - PBYTE(OldData)));

		secDir->Size = 0;
		secDir->VirtualAddress = 0;

		RtlCopyMemory(NewSection->Name, ".1hooks1", 8);
		NewSection->Misc.VirtualSize = NewImportSize;
		NewSection->VirtualAddress = AlignValueUp(LastSectionVA + LastSectionSize, headers->OptionalHeader.SectionAlignment);
		NewSection->SizeOfRawData = dwSizeofNewSection;
		NewSection->PointerToRawData = dwOffsetOfLastSectionEnd;
		NewSection->PointerToRelocations = 0;
		NewSection->PointerToLinenumbers = 0;
		NewSection->NumberOfRelocations = 0;
		NewSection->NumberOfLinenumbers = 0;
		NewSection->Characteristics = IMAGE_SCN_MEM_READ | IMAGE_SCN_CNT_INITIALIZED_DATA | IMAGE_SCN_MEM_WRITE;

		headers->OptionalHeader.SizeOfHeaders += sizeof(IMAGE_SECTION_HEADER);
		headers->OptionalHeader.SizeOfImage += AlignValueUp(NewImportSize, headers->OptionalHeader.SectionAlignment);
		++headers->FileHeader.NumberOfSections;

		PIMAGE_IMPORT_DESCRIPTOR iid = PIMAGE_IMPORT_DESCRIPTOR(PBYTE(NewData) + NewSection->PointerToRawData);
		PIMAGE_IMPORT_DESCRIPTOR old = PIMAGE_IMPORT_DESCRIPTOR(PBYTE(NewData) + VA2RVA(importSection, importDir->VirtualAddress));
		DWORD numberOfIIDs = importDir->Size / sizeof(IMAGE_IMPORT_DESCRIPTOR);
		PBYTE dataTable = (PBYTE)&iid[numberOfIIDs + 1];

		RtlCopyMemory(iid, old, importDir->Size);
		importDir->VirtualAddress = RVA2VA(NewSection, DWORD(PBYTE(iid) - PBYTE(NewData)));
		importDir->Size += sizeof(IMAGE_IMPORT_DESCRIPTOR);

		iid = &iid[numberOfIIDs - 1];
		iid->TimeDateStamp = 0;
		iid->ForwarderChain = 0;

		PULONG_PTR lpNewSnap = (PULONG_PTR)dataTable;
		iid->OriginalFirstThunk = RVA2VA(NewSection, DWORD(PBYTE(lpNewSnap) - PBYTE(NewData)));
		dataTable += sizeof(PVOID) * 2;
		lpNewSnap[0] = 1 | IMAGE_ORDINAL_FLAG;
		lpNewSnap[1] = 0;

		PULONG_PTR lpThunk = (PULONG_PTR)dataTable;
		iid->FirstThunk = RVA2VA(NewSection, DWORD(PBYTE(lpThunk) - PBYTE(NewData)));
		dataTable += sizeof(PVOID) * 2;
		lpThunk[0] = 1 | IMAGE_ORDINAL_FLAG;
		lpThunk[1] = 0;

		iid->Name = RVA2VA(NewSection, DWORD(PBYTE(dataTable) - PBYTE(NewData)));
		strcpy((LPSTR)dataTable, "SppcHook.dll");
		dataTable += 14;

		DWORD dwSizeWritten;
		success = WriteFile(
			hNewFile,
			NewData,
			dwSizeOfNewFile,
			&dwSizeWritten,
			nullptr
		) && dwSizeOfNewFile == dwSizeWritten;

		result = success ? 0 : 7;
	} while (FALSE);

	if (OldData)UnmapViewOfFile(OldData);
	if (NewData)VirtualFree(NewData, 0, MEM_RELEASE);
	if (hNewFile != INVALID_HANDLE_VALUE) {
		CloseHandle(hNewFile);

		//
		// If we failed, delete the file.
		//
		if (0 != result)DeleteFile(lpNewMsoPath);
	}

	return result;
}

