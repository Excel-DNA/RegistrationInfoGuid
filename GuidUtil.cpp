#include "stdafx.h"
#include "sha1.h"
#include "GuidUtil.h"

#define byte unsigned char

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

byte* CreateGuid(byte* nameSpaceGuid, byte* name, unsigned int nameLength)
{
	SwapByteOrder(nameSpaceGuid);

	SHA1Context c;
	SHA1Reset(&c);
	SHA1Input(&c, nameSpaceGuid, 16);
	SHA1Input(&c, name, nameLength);
	int resultOK = SHA1Result(&c);
	unsigned int * hash = c.Message_Digest;	// uint[5]

	byte newGuid[16];
	for (int i = 0; i < 4; i++)
		CopyIntToBytes(c.Message_Digest[i], newGuid, i * 4);

	// set the four most significant bits (bits 12 through 15) of the time_hi_and_version field to the appropriate 4-bit version number
	int version = 5; // for SHA - 1 hashing
	newGuid[6] = (byte)((newGuid[6] & 0x0F) | (version << 4));
	
	// set the two most significant bits (bits 6 and 7) of the clock_seq_hi_and_reserved to zero and one, respectively
	newGuid[8] = (byte)((newGuid[8] & 0x3F) | 0x80);
	
	// convert the resulting UUID to local byte order
	SwapByteOrder(newGuid);
	return (byte*)newGuid;
}

typedef struct _dna_guid_t {
	unsigned __int32 time_low;
	unsigned __int16 time_mid;
	unsigned __int16 time_hi_and_version;
	unsigned __int8  clock_seq_hi_and_reserved;
	unsigned __int8  clock_seq_low;
	unsigned __int8  node[6];
} dna_guid_t;

dna_guid_t NameSpace_ExcelDna = { /* 306D016E-CCE8-4861-9DA1-51A27CBE341A */
	0x306D016E,
	0xCCE8,
	0x4861,
	0x9D, 0xA1, 0x51, 0xA2, 0x7C, 0xBE, 0x34, 0x1A,
};

dna_guid_t* GuidFromXllPath(const wchar_t* xllPath)
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

	return (dna_guid_t*)CreateGuid((byte*)&NameSpace_ExcelDna, (byte*)name, nameLength);
}

void GuidStringFromXllPath(const wchar_t* xllPath, wchar_t* guidStringBuffer)
{
	dna_guid_t* pu = GuidFromXllPath(xllPath);

	// Equivalent of the "N" Guid format in .NET
	wsprintf(guidStringBuffer, L"%8.8x%4.4x%4.4x%2.2x%2.2x%2.2x%2.2x%2.2x%2.2x%2.2x%2.2x\0",
		pu->time_low, pu->time_mid,
		pu->time_hi_and_version, pu->clock_seq_hi_and_reserved,
		pu->clock_seq_low,
		pu->node[0], pu->node[1], pu->node[2], pu->node[3], pu->node[4], pu->node[5]);
}
