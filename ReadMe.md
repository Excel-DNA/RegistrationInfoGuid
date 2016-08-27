# RegistrationInfo Guid Sample

The Excel-DNA IntelliSense extension will attempt to find registration info for any .xll add-in loaded.
Discovery of the registration info is by calling a hidden function called "RegistrationInfo_<XllGuid>", where "<XllGuid>" is a guid computed from the full path of the xll.
This project shows how to compute this Guid.

## RegistrationInfo registration

The RegistrationInfo function should be registered with type "QQ", and may be a hidden function.

## RegistrationInfo implementation

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

The GuidUtility code used in Excel-DNA to generate the version 5 Guid is from https://github.com/LogosBible/Logos.Utility/blob/master/src/Logos.Utility/GuidUtility.cs

The SHA1 code in this project is from https://github.com/Spl3en/Crypto

