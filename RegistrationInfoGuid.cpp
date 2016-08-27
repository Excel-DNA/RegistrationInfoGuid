#include "stdafx.h"
#include "GuidUtil.h"

int main()
{
	// pXllPath should be the full path to the xll

	wchar_t* pXllPath = L"C:\\Temp\\Test.xll";	// Result for "C:\\Temp\\Test.xll" should be: f3fe5a6e63fc5c18b9bb5ae7b3d12632
	wchar_t guidStringBuffer[33];
	GuidStringFromXllPath(pXllPath, guidStringBuffer);

	// UDF name is then RegistrationInfo_f3fe5a6e63fc5c18b9bb5ae7b3d12632

	wprintf(L"Registration UDF for ");
	wprintf(pXllPath);
	wprintf(L" will be \n");
	wprintf(L"RegistrationInfo_");
	wprintf(guidStringBuffer);

	printf("\n");
	getchar();
}

