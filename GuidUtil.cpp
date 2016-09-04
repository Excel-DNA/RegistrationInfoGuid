#include "stdafx.h"
#include "GuidUtil.h"
#include "sha1\\sha1.h"

#define byte unsigned char

#ifdef __GNUC__
#define PACKED __attribute__((packed))
#else
#define PACKED
#pragma pack(1)
#endif

typedef struct _dna_guid_t {
	unsigned __int32 time_low;
	unsigned __int16 time_mid;
	unsigned __int16 time_hi_and_version;
	unsigned __int8  clock_seq_hi_and_reserved;
	unsigned __int8  clock_seq_low;
	unsigned __int8  node[6];
} PACKED dna_guid_t;

dna_guid_t NameSpace_ExcelDna = { /* 306D016E-CCE8-4861-9DA1-51A27CBE341A */
	0x306D016E,
	0xCCE8,
	0x4861,
	0x9D, 0xA1, 0x51, 0xA2, 0x7C, 0xBE, 0x34, 0x1A,
};

void SwapBytes(byte* guid, int left, int right)
{
	byte temp = guid[left];
	guid[left] = guid[right];
	guid[right] = temp;
}

// Converts a GUID (expressed as a byte array) to/from network order (MSB-first).
void SwapByteOrder(byte* guid)
{
	SwapBytes(guid, 0, 3);
	SwapBytes(guid, 1, 2);
	SwapBytes(guid, 4, 5);
	SwapBytes(guid, 6, 7);
}

void CopyIntToBytes(int n, byte* pBytes, int offset)
{
	pBytes[offset] = (n >> 24) & 0xFF;
	pBytes[offset + 1] = (n >> 16) & 0xFF;
	pBytes[offset + 2] = (n >> 8) & 0xFF;
	pBytes[offset + 3] = n & 0xFF;
}

void CreateGuid(byte* nameSpaceGuid, byte* name, unsigned int nameLength, _dna_guid_t* guid)
{
	SwapByteOrder(nameSpaceGuid);

	unsigned int hash[5];

	SHA1 c;
	c.Reset();
	c.Input(nameSpaceGuid, 16);
	c.Input(name, nameLength);
	int resultOK = c.Result(hash);

	byte* newGuid = reinterpret_cast<byte*>(guid);
	for (int i = 0; i < 4; i++)
		CopyIntToBytes(hash[i], newGuid, i * 4);

	// set the four most significant bits (bits 12 through 15) of the time_hi_and_version field to the appropriate 4-bit version number
	int version = 5; // for SHA - 1 hashing
	newGuid[6] = (byte)((newGuid[6] & 0x0F) | (version << 4));
	
	// set the two most significant bits (bits 6 and 7) of the clock_seq_hi_and_reserved to zero and one, respectively
	newGuid[8] = (byte)((newGuid[8] & 0x3F) | 0x80);
	
	// convert the resulting UUID to local byte order
	SwapByteOrder(newGuid);
}

void GuidFromXllPath(const wchar_t* xllPath, _dna_guid_t* guid)
{
	// Convert ToUpper and to UTF8 byte array

	wchar_t xllPathUpper[512];

	int i = 0;
	while (xllPath[i])
	{
		xllPathUpper[i] = towupper(xllPath[i]);
		i++;
	}
	xllPathUpper[i] = L'\0';

	const int MAX_LENGTH = 1024;
	char name[MAX_LENGTH];
	int nameLength = WideCharToMultiByte(CP_UTF8, 0, xllPathUpper, -1, name, MAX_LENGTH, NULL, NULL);
	nameLength--; /* ignore null terminator */

	CreateGuid((byte*)&NameSpace_ExcelDna, (byte*)name, nameLength, guid);
}

void GuidStringFromXllPath(const wchar_t* xllPath, wchar_t* guidStringBuffer)
{
	_dna_guid_t guid;
	GuidFromXllPath(xllPath, &guid);

	// Equivalent of the "N" Guid format in .NET
	wsprintfW(guidStringBuffer, L"%8.8x%4.4x%4.4x%2.2x%2.2x%2.2x%2.2x%2.2x%2.2x%2.2x%2.2x\0",
		guid.time_low, guid.time_mid,
		guid.time_hi_and_version, guid.clock_seq_hi_and_reserved,
		guid.clock_seq_low,
		guid.node[0], guid.node[1], guid.node[2], guid.node[3], guid.node[4], guid.node[5]);
}
