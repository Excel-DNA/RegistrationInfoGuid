# IntelliSense Refresh Control Macro

The IntelliSense server registers a control macro that can be called from VBA or an .xll to force a refresh of the IntelliSense macro information.

The control macro is called `IntelliSenseServerControl_{serverId}`, where `{serverId}` is the server ID for the active IntelliSense server.
It is registered with the C API type string `"QQ"` and macro type `2`.

## Determining the Active Server ID

Information about the active server is set in the environment variable `EXCELDNA_INTELLISENSE_ACTIVE_SERVER`.

* If this environment variable does not exist or is empty, there is no active IntelliSense server.

* Otherwise the environment variable shoudl have a three-part format, delimited by `','`s. The three parts are the XLL path, the Server Id and the server version of the active server.
(Future versions of the IntelliSense might add extra parts to this string.)

* E.g. the environment variable might have the value `C:\Temp\ExcelDna.IntelliSense.Host-AddIn.xll,02e87607ece24348ab5deb1dd705b6da,0.1.7`.
For the control macro, we use the second part, to give use the control macro name `IntelliSenseServerControl_02e87607ece24348ab5deb1dd705b6da`.

## Calling the control macro

The control macro takes one parameter, indicating the control command to run.
To refresh the server, the command parameter should have the string value `"REFRESH"`.

The control macro will return:
* `true` if the refresh command was successful, 
* `false` if the server called is not the active server,
* the `#VALUE` error value if there is some internal error.

The control macro can be called via the C API (`xlcRun`) or the COM object model (`Application.Run`).
