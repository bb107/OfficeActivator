#pragma once

INT IsMsoPatched(_In_ LPCTSTR lpMsoPath);

INT PatchMsoFile(
	_In_ LPCTSTR lpMsoPath,
	_In_ LPCTSTR lpNewMsoPath
);
