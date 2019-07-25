# ProfileProp
Migrated from my [blog](https://blogs.msdn.microsoft.com/stephen_griffin/2015/07/18/profileprop-examine-profile-properties/).

I’ve had some requests lately to write a MAPI sample that shows how to access profile properties programmatically, so I threw this together from some bits and pieces of code I had laying around.

I had the chance to incorporate a few cool features in this code:
1. [MultiEx](https://msdn.microsoft.com/en-us/library/office/ff522797.aspx): It loops over the profile services looking for services of type MSEMS, grabbing the [PR_EMSMDB_SECTION_UID](https://msdn.microsoft.com/en-us/library/office/ff625289.aspx) property for each. It uses this to open the global profile section for each account.
2. [MAPI Stub Library](http://mapistublibrary.codeplex.com): This sample uses the MAPI Stub Library, demonstrating yet again how easy this library is to incorporate. The version I’m using here is cribbed from the [MFCMAPI](https://github.com/stephenegriffin/mfcmapi) source.
3. The profile name is optional: You can specify any profile you want, or if you leave it out, it will look up and use the default profile.
4. Deletion: If you have a property you want to delete, you can pass the –d switch. Be careful with this! Deleting random properties could corrupt your profile and leave you in a state where the only fix is to recreate the profile or restore from backup. Make sure you know what your doing before you delete a property!
5. Backup: If you do use the deletion switch, this sample will create a backup of the profile using [CopyProfile](https://msdn.microsoft.com/en-us/library/cc815840(v=office.15).aspx).

Here’s the help, showing how to use the sample:
```
C:\>ProfileProp.exe -?
ProfileProp - Profile Property Examination Tool
Locates and optionally deletes a property from the Exchange Global Profile section of a profile.
In the case of multiple Exchange accounts, will locate the property for each account.

Usage:  ProfileProp [-?] [-p profile] [-d] <property tag number>

Options:
-p profile Name of new profile to examine.
Default profile will be used if -p is not used.
-d         Delete the property (otherwise just locate it)
-?         Displays this usage information.</td>
```

Suppose you wanted to use this to output the display name of each account in the profile. You could do this (note that 0x3001001F is [PR_DISPLAY_NAME](https://msdn.microsoft.com/en-us/library/office/cc842383.aspx">PR_DISPLAY_NAME)):
```
C:\>ProfileProp.exe 0x3001001F
Profile Property Tool
Profile: Outlook
Property: 0x3001001F
Examining account Microsoft Exchange
Located property
Property tag: 0x3001001F
Value: sgriffin@example.com
Examining account Microsoft Exchange
Located property
Property tag: 0x3001001F
Value: sgriffin@example.example.com
```
