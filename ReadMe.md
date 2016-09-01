# RegistrationInfo Guid Sample

The Excel-DNA IntelliSense extension will attempt to find registration info for any .xll add-in loaded.
Discovery of the registration info is by calling a hidden function called "RegistrationInfo_\<XllGuid\>", where "\<XllGuid\>" is a guid computed from the full path of the xll.
This project shows how to compute this Guid.

## RegistrationInfo Guid calculation

The Guid assocated with the xll is computed as the version 5 (SHA1) UUID (in accordance with section 4.3 of RFC 4122) using:
* the special Excel-DNA name space ID 306D016E-CCE8-4861-9DA1-51A27CBE341A,
* the SHA-1 hash algorithm of the full path of the xll, converted to upper-case and UTF-8 encoding.

An implementation of the Guid calculation in C# is [included in Excel-DNA](https://github.com/Excel-DNA/ExcelDna/blob/master/Source/ExcelDna.Integration/GuidUtility.cs).

This project provides a small C program that computes the same Guid.

As a test, the path "C:\Temp\Test.xll" should evaluate to the Guid string: f3fe5a6e63fc5c18b9bb5ae7b3d12632.

The registration function for the add-in "C:\Temp\Test.xll" will thus be called "RegistrationInfo_f3fe5a6e63fc5c18b9bb5ae7b3d12632".

## RegistrationInfo implementation

The RegistrationInfo function should be registered with type "QQ", and may be a hidden function.

The parameter is interpreted as a version number.
To allow polling, we return as the first row a (double) version which can be passed to short-circuit the call if nothing has changed.

The RegistrationInfo function should check the first parameter - if it is the same as the last returned version, the RegistrationInfo function can return null (`xlEmpty`?).

Otherwise the IntelliSense extension expects an array with 255 columns, one row for the version and one row for each function.
The first row should contain:
* the full path to the xll, and
* the info version number.

Every registration row should contain exactly the information passed to the `xlfRegister` function.
Any missing entries in the array should be returned as `xlEmpty`.

## Acknowledgements

The original GuidUtility code used in Excel-DNA to generate the version 5 Guid is from https://github.com/LogosBible/Logos.Utility/blob/master/src/Logos.Utility/GuidUtility.cs

The SHA-1 code in this project is from https://www.packetizer.com/security/sha1/

