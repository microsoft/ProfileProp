#include <mapix.h>
#include <mapiutil.h>
#include <stdio.h>
#include "Profiles.h"
#include <string>

struct MYOPTIONS
{
	LPSTR lpszProfile;
	BOOL bDeleteProperty;
	ULONG ulPropNum;
	std::string lpszPropName;
};

void DisplayUsage()
{
	printf("ProfileProp - Profile Property Examination Tool\n");
	printf("   Locates and optionally deletes a property from the Exchange Global Profile section of a profile.\n");
	printf("   In the case of multiple Exchange accounts, will locate the property for each account.\n");
	printf("\n");
	printf("Usage:  ProfileProp [-?] [-p profile] [-d] <property tag number>\n");
	printf("\n");
	printf("Options:\n");
	printf("        -p profile Name of new profile to examine.\n");
	printf("                      Default profile will be used if -p is not used.\n");
	printf("\n");
	printf("        -d         Delete the property (otherwise just locate it)\n");
	printf("\n");
	printf("        -?         Displays this usage information.\n");
}

#define _NAN 0xFFFFFFFF
// scans an arg and returns the string or hex number that it represents
void GetArg(char* szArgIn, char** lpszArgOut, ULONG* ulArgOut)
{
	ULONG ulArg = NULL;
	char* szArg = NULL;
	LPSTR szEndPtr = NULL;
	ulArg = strtoul(szArgIn, &szEndPtr, 16);

	// if szEndPtr is pointing to something other than NULL, this must be a string
	if (!szEndPtr || *szEndPtr)
	{
		ulArg = _NAN;
		szArg = szArgIn;
	}

	if (lpszArgOut) *lpszArgOut = szArg;
	if (ulArgOut)   *ulArgOut = ulArg;
}

BOOL ParseArgs(int argc, char * argv[], MYOPTIONS * pRunOpts)
{
	if (!pRunOpts) return FALSE;

	ZeroMemory(pRunOpts, sizeof(MYOPTIONS));

	for (int i = 1; i < argc; i++)
	{
		switch (argv[i][0])
		{
		case '-':
		case '/':
		case '\\':
			if (0 == argv[i][1])
			{
				// Bad argument - get out of here
				return false;
			}

			switch (tolower(argv[i][1]))
			{
			case 'p':
				if (i + 1 < argc)
				{
					pRunOpts->lpszProfile = argv[i + 1];
					i++;
				}
				else return false;
				break;
			case 'd':
				pRunOpts->bDeleteProperty = true;
				break;
			case '?':
			default:
				// display help
				return false;
				break;
			}

			break;
		default:
			// Naked option without a flag, must be a property name or number
			if (!pRunOpts->lpszPropName.empty()) return false; // He's already got one, you see.
			pRunOpts->lpszPropName = argv[i];
			break;
		}
	}

	if (!pRunOpts->lpszPropName.empty())
	{
		ULONG ulArg = NULL;
		LPSTR szEndPtr = NULL;
		ulArg = strtoul(pRunOpts->lpszPropName.c_str(), &szEndPtr, 16);

		// If szEndPtr is pointing to something other than NULL, this must be a string
		if (!szEndPtr || *szEndPtr)
		{
			ulArg = NULL;
		}

		pRunOpts->ulPropNum = ulArg;
	}

	// Validate that we have bare minimum to run
	if (pRunOpts->lpszPropName.empty()) return false;

	// Didn't fail - return true
	return true;
}

void main(int argc, char* argv[])
{
	MYOPTIONS ProgOpts = { 0 };

	if (!ParseArgs(argc, argv, &ProgOpts))
	{
		DisplayUsage();
		return;
	}

	printf("Profile Property Tool\n");
	HRESULT hRes = S_OK;

	hRes = MAPIInitialize(NULL);
	if (SUCCEEDED(hRes))
	{
		std::string szDefaultProfile;
		if (ProgOpts.lpszProfile == NULL)
		{
			szDefaultProfile = GetDefaultProfileName();
			if (!szDefaultProfile.empty())
			{
				ProgOpts.lpszProfile = (LPSTR)szDefaultProfile.c_str();
			}
		}

		if (ProgOpts.lpszProfile != NULL)
		{
			printf("Profile: %s\n", ProgOpts.lpszProfile);
			printf("Property: 0x%08X\n", ProgOpts.ulPropNum);
			if (ProgOpts.bDeleteProperty)
			{
				BackupProfile(ProgOpts.lpszProfile);
				printf("\tDeleting Property\n");
			}

			printf("\n");

			hRes = ExamineProperty(ProgOpts.lpszProfile, ProgOpts.ulPropNum, ProgOpts.bDeleteProperty);
		}
		else
		{
			printf("No profile found\n");
		}

		MAPIUninitialize();
	}
	else
	{
		printf("Error initializing MAPI. HRESULT = 0x%08X\n", hRes);
	}
}