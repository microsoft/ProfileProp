#include <mapix.h>
#include <MAPIUtil.h>
#include <edkmdb.h>
#include <time.h>
#include "Profiles.h"
#include "stdio.h"

typedef std::basic_string<TCHAR> tstring;

#define PR_EMSMDB_SECTION_UID PROP_TAG( PT_BINARY, 0x3d15)

STDMETHODIMP ExamineProperty(LPSERVICEADMIN lpSvcAdmin, LPMAPIUID pEmsmdbUID, ULONG ulPropTag, BOOL bDeleteProperty);

std::string GetDefaultProfileName()
{
	std::string szDefaultProfile;
	LPPROFADMIN lpProfileAdmin = NULL;
	HRESULT hRes = MAPIAdminProfiles(0, &lpProfileAdmin);
	if (FAILED(hRes) || lpProfileAdmin == NULL)
	{
		return "";
	}

	LPMAPITABLE lpProfileTable = NULL;
	hRes = lpProfileAdmin->GetProfileTable(0, &lpProfileTable);
	if (FAILED(hRes) || lpProfileTable == NULL)
	{
		return "";
	}

	SPropValue spvDefaultProfile;
	spvDefaultProfile.ulPropTag = PR_DEFAULT_PROFILE;
	spvDefaultProfile.Value.b = TRUE;

	SRestriction srProfile;
	srProfile.rt = RES_PROPERTY;
	srProfile.res.resProperty.relop = RELOP_EQ;
	srProfile.res.resProperty.ulPropTag = PR_DEFAULT_PROFILE;
	srProfile.res.resProperty.lpProp = &spvDefaultProfile;

	enum { iDispName, iDefaultProfile, cptaProfile };
	SizedSPropTagArray(cptaProfile, sptCols) = { cptaProfile, PR_DISPLAY_NAME_A, PR_DEFAULT_PROFILE };

	LPSRowSet lpRowSet = NULL;
	hRes = HrQueryAllRows(
		lpProfileTable,
		(LPSPropTagArray)&sptCols,
		&srProfile,
		NULL,
		0,
		&lpRowSet);
	if (SUCCEEDED(hRes) && lpRowSet != NULL && lpRowSet->cRows == 1 && lpRowSet->aRow[0].lpProps[iDefaultProfile].ulPropTag == PR_DEFAULT_PROFILE)
	{
		szDefaultProfile = lpRowSet->aRow[0].lpProps[iDispName].Value.lpszA;
	}

	FreeProws(lpRowSet);
	return szDefaultProfile;
}

std::string GetTime()
{
	time_t timeObj;
	time(&timeObj);
	tm *pTime = localtime(&timeObj);
	char buffer[100];
	sprintf(buffer, "%d-%02d-%02d-%02d:%02d:%02d",
		pTime->tm_year+1900, pTime->tm_mon+1, pTime->tm_mday,
		pTime->tm_hour, pTime->tm_min, pTime->tm_sec);
	return buffer;
}

void BackupProfile(LPSTR lpszProfileName)
{
	printf("Backing up profile %s\n", lpszProfileName);
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;

	hRes = MAPIAdminProfiles(0, &lpProfAdmin);
	if (SUCCEEDED(hRes) && lpProfAdmin)
	{
		tstring szNewName = lpszProfileName;
		szNewName.append("-ProfilePropBackup-");
		szNewName.append(GetTime());

		hRes = lpProfAdmin->CopyProfile(
			lpszProfileName,
			NULL,
			(LPTSTR)szNewName.c_str(),
			NULL,
			0);
		if (FAILED(hRes))
		{
			printf("Profile backup failed 0x%08X\n", hRes);
		}
		else
		{
			printf("Profile backed up to %s\n", szNewName.c_str());
		}
	}

	if (lpProfAdmin) lpProfAdmin->Release();
}

STDMETHODIMP SvcAdminOpenProfileSection(LPSERVICEADMIN lpSvcAdmin,
	LPMAPIUID lpUID,
	LPCIID lpInterface,
	ULONG ulFlags,
	LPPROFSECT FAR * lppProfSect)
{
	HRESULT hRes = S_OK;

	// Note: We have to open the profile section with full access.
	// MAPI discriminates who can modify profiles, especially
	// in certain sections.  The way to force access has changed in
	// different versions of Outlook. Therefore, there are two methods.  See KB article 822977
	// for more information.

	// First, let us try the easier method of passing the MAPI_FORCE_ACCESS flag
	// to OpenProfileSection.  This method is available only in Outlook 2003 and in later versions of Outlook.

	hRes = lpSvcAdmin->OpenProfileSection(lpUID,
		lpInterface,
		ulFlags | MAPI_FORCE_ACCESS,
		lppProfSect);
	if (FAILED(hRes))
	{
		// If this does not succeed, it may be because you are using an earlier version of Outlook.
		// In this case, use the sample code
		// from KB article 228736 for more information.  Note: This information was compiled from that sample.

		///////////////////////////////////////////////////////////////////
		// MAPI will always return E_ACCESSDENIED
		// when we open a profile section on the service if we are a client.  The workaround
		// is to call into one of MAPI's internal functions that bypasses
		// the security check.  We build an interface to it, and then point to it from our
		// offset of 0x48.  USE THIS METHOD AT YOUR OWN RISK! THIS METHOD IS NOT SUPPORTED!
		interface IOpenSectionHack : public IUnknown
		{
		public:
			virtual HRESULT STDMETHODCALLTYPE OpenSection(LPMAPIUID, ULONG, LPPROFSECT*) = 0;
		};

		IOpenSectionHack** ppProfile = (IOpenSectionHack**)((((BYTE*)lpSvcAdmin) + 0x48));

		// Now, we want to open the Services Profile Section and store that
		// interface with the Object
		hRes = (*ppProfile)->OpenSection(lpUID,
			ulFlags,
			lppProfSect);

		//
		///////////////////////////////////////////////////////////////////
	}

	return hRes;
}

STDMETHODIMP ProvAdminOpenProfileSection(LPPROVIDERADMIN lpProvAdmin,
	LPMAPIUID lpUID,
	LPCIID lpInterface,
	ULONG ulFlags,
	LPPROFSECT FAR * lppProfSect)
{
	HRESULT hRes = S_OK;

	hRes = lpProvAdmin->OpenProfileSection(lpUID,
		lpInterface,
		ulFlags | MAPI_FORCE_ACCESS,
		lppProfSect);

	if ((FAILED(hRes)) && (MAPI_E_UNKNOWN_FLAGS == hRes))
	{
		// The MAPI_FORCE_ACCESS flag is implemented only in Outlook 2002 and in later versions of Outlook.

		// Makes MAPI think we are a service and not a client.
		// MAPI grants us Service Access.  This makes it all possible.
		*(((BYTE*)lpProvAdmin) + 0x60) = 0x2;  // USE THIS METHOD AT YOUR OWN RISK! THIS METHOD IS NOT SUPPORTED!

		hRes = lpProvAdmin->OpenProfileSection(lpUID,
			lpInterface,
			ulFlags,
			lppProfSect);
	}

	return hRes;
}

STDMETHODIMP ExamineProperty(LPSTR lpszProfileName, ULONG ulPropTag, BOOL bDeleteProperty)
{
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;

	hRes = MAPIAdminProfiles(0, &lpProfAdmin);
	if (SUCCEEDED(hRes) && lpProfAdmin)
	{
		LPSERVICEADMIN lpSvcAdmin = NULL;

		hRes = lpProfAdmin->AdminServices(
			(LPTSTR)lpszProfileName,
			NULL,
			NULL,
			0,
			&lpSvcAdmin);
		if (SUCCEEDED(hRes) && lpSvcAdmin)
		{
			LPMAPITABLE lpMsgSvcTable = NULL;

			hRes = lpSvcAdmin->GetMsgServiceTable(0, &lpMsgSvcTable);
			if (SUCCEEDED(hRes) && lpMsgSvcTable)
			{
				LPSRowSet lpSvcRows = NULL;

				enum { iSvcName, iDispName, iSvcUID, cptaSvc };
				SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_SERVICE_NAME_A, PR_DISPLAY_NAME_A, PR_EMSMDB_SECTION_UID };

				hRes = HrQueryAllRows(
					lpMsgSvcTable,
					(LPSPropTagArray)&sptCols,
					NULL,
					NULL,
					0,
					&lpSvcRows);
				if (SUCCEEDED(hRes) && NULL != lpSvcRows)
				{
					ULONG i = 0;
					for (i = 0; i < lpSvcRows->cRows; i++)
					{
						if (lpSvcRows->aRow[i].lpProps[iSvcName].ulPropTag == PR_SERVICE_NAME_A)
						{
							if (0 != strcmp("MSEMS", lpSvcRows->aRow[i].lpProps[iSvcName].Value.lpszA))
							{
								continue;
							}
						}

						if (lpSvcRows->aRow[i].lpProps[iDispName].ulPropTag == PR_DISPLAY_NAME_A)
						{
							printf("Examining account %s\n", lpSvcRows->aRow[i].lpProps[iDispName].Value.lpszA);
						}

						if (lpSvcRows->aRow[i].lpProps[iSvcUID].ulPropTag == PR_EMSMDB_SECTION_UID)
						{
							ExamineProperty(lpSvcAdmin, (LPMAPIUID)lpSvcRows->aRow[i].lpProps[iSvcUID].Value.bin.lpb, ulPropTag, bDeleteProperty);
						}
						else
						{
							printf("\tDoes not appear to be an Exchange account.\n");
						}
					}
				}

				FreeProws(lpSvcRows);
			}

			if (lpMsgSvcTable) lpMsgSvcTable->Release();
		}

		if (lpSvcAdmin) lpSvcAdmin->Release();
	}

	if (lpProfAdmin) lpProfAdmin->Release();

	return hRes;
}

std::wstring format(const LPWSTR fmt, ...)
{
	LPWSTR buffer = NULL;
	va_list vl;
	va_start(vl, fmt);
	int len = _vscwprintf(fmt, vl);
	if (0 != len)
	{
		len++;
		buffer = new wchar_t[len];
		memset(buffer, 0, sizeof(wchar_t)* len);
		(void)_vsnwprintf_s(buffer, len, len, fmt, vl);
	}

	std::wstring ret(buffer);
	va_end(vl);
	delete[] buffer;
	return ret;
}

void PrintProperty(LPSPropValue lpProp)
{
	if (!lpProp) return;

	printf("\t\tProperty tag: 0x%08X\n", lpProp->ulPropTag);

	ULONG iMVCount = 0;
	if (MV_FLAG & PROP_TYPE(lpProp->ulPropTag))
	{
		return;
	}
	else
	{
		std::wstring szTmp;
		std::wstring szAltTmp;
		switch (PROP_TYPE(lpProp->ulPropTag))
		{
		case PT_I2:
			szTmp = format(L"%d", lpProp->Value.i);
			szAltTmp = format(L"0x%X", lpProp->Value.i);
			break;
		case PT_LONG:
			szTmp = format(L"%d", lpProp->Value.l);
			szAltTmp = format(L"0x%X", lpProp->Value.l);
			break;
		case PT_R4:
			szTmp = format(L"%f", lpProp->Value.flt);
			break;
		case PT_DOUBLE:
			szTmp = format(L"%f", lpProp->Value.dbl);
			break;
		case PT_CURRENCY:
			szTmp = format(L"%05I64d", lpProp->Value.cur.int64);
			if (szTmp.length() > 4)
			{
				szTmp.insert(szTmp.length() - 4, L".");
			}

			szAltTmp = format(L"0x%08X:0x%08X", (int)(lpProp->Value.cur.Hi), (int)lpProp->Value.cur.Lo);
			break;
		case PT_APPTIME:
			szTmp = format(L"%f", lpProp->Value.at);
			break;
		case PT_ERROR:
			szTmp = format(L"0x%08X", lpProp->Value.err);
			break;
		case PT_BOOLEAN:
			if (lpProp->Value.b)
				szTmp = L"true";
			else
				szTmp = L"false";
			break;
		case PT_OBJECT:
			szTmp = L"object";
			break;
		case PT_I8: // LARGE_INTEGER
			szTmp = format(L"0x%08X:0x%08X", (int)(lpProp->Value.li.HighPart), (int)lpProp->Value.li.LowPart);
			szAltTmp = format(L"%I64d", lpProp->Value.li.QuadPart);
			break;
		case PT_STRING8:
			szTmp = format(L"%hs", lpProp->Value.lpszA);
			break;
		case PT_UNICODE:
			szTmp = format(L"%ws", lpProp->Value.lpszA);
			break;
		case PT_SYSTIME:
			break;
		case PT_CLSID:
			break;
		case PT_BINARY:
			szTmp = format(L"Size: %d bytes", lpProp->Value.bin.cb);
			break;
		default:
			break;
		}

		if (!szTmp.empty())
		{
			printf("\t\tValue: %ws\n", szTmp.c_str());
		}

		if (!szAltTmp.empty())
		{
			printf("\t\tAlternate value: %ws\n", szAltTmp.c_str());
		}
	}
}

STDMETHODIMP ExamineProperty(LPSERVICEADMIN lpSvcAdmin, LPMAPIUID pEmsmdbUID, ULONG ulPropTag, BOOL bDeleteProperty)
{
	HRESULT hRes = S_OK;

	LPPROFSECT lpProfileSection = NULL;
	hRes = SvcAdminOpenProfileSection(
		lpSvcAdmin,
		pEmsmdbUID,
		NULL,
		MAPI_MODIFY,
		&lpProfileSection);
	if (SUCCEEDED(hRes) && lpProfileSection)
	{
		SizedSPropTagArray(1, sptTags) = { 1, ulPropTag };

		LPSPropValue pPropArray = 0;
		ULONG ulValues = 0;

		hRes = lpProfileSection->GetProps((LPSPropTagArray)&sptTags, 0, &ulValues, &pPropArray);

		if (SUCCEEDED(hRes) && MAPI_W_ERRORS_RETURNED != hRes && ulPropTag == pPropArray[0].ulPropTag)
		{
			printf("\tLocated property\n");
			PrintProperty(&pPropArray[0]);

			if (bDeleteProperty)
			{
				printf("\n");
				printf("\tDeleting Property\n");
				hRes = lpProfileSection->DeleteProps((LPSPropTagArray)&sptTags, NULL);

				if (SUCCEEDED(hRes))
				{
					printf("\t\tSucceeded in deleting property.\n");
				}
				else
				{
					printf("\t\tFailed to delete property.\n");
				}
			}
		}
		else
		{
			if (MAPI_W_ERRORS_RETURNED != hRes)
			{
				printf("Error (0x%x): Could not locate property.\n", hRes);
			}
			else
			{
				if (pPropArray != NULL && PROP_TYPE(pPropArray[0].ulPropTag) == PT_ERROR && pPropArray[0].Value.err != MAPI_E_NOT_FOUND)
				{
					printf("\tCould not locate property. Error (0x%x)\n", pPropArray[0].Value.err);
				}
				else
				{
					printf("\tThis property was not found in this account.\n");
				}
			}
		}
	}

	if (lpProfileSection) lpProfileSection->Release();

	return hRes;
}