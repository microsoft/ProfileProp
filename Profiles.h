#pragma once
#include <string>

// Profile Section GUIDs
#define pbGlobalProfileSectionGuid	"\x13\xDB\xB0\xC8\xAA\x05\x10\x1A\x9B\xB0\x00\xAA\x00\x2F\xC4\x5A"

// Flags
#define MAPI_FORCE_ACCESS 0x00080000

std::string GetDefaultProfileName();
void BackupProfile(LPSTR lpszProfileName);

STDMETHODIMP SvcAdminOpenProfileSection(LPSERVICEADMIN lpSvcAdmin,
										LPMAPIUID lpUID,
										LPCIID lpInterface,
										ULONG ulFlags,
										LPPROFSECT FAR * lppProfSect);

STDMETHODIMP ProvAdminOpenProfileSection(LPPROVIDERADMIN lpProvAdmin,
										 LPMAPIUID lpUID,
										 LPCIID lpInterface,
										 ULONG ulFlags,
										 LPPROFSECT FAR * lppProfSect);

STDMETHODIMP ExamineProperty(LPSTR lpszProfileName, ULONG ulPropTag, BOOL bDeleteProperty);