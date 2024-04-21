## TODO

1. Ribbon_GetContent - is hacked to get it working
1. Propmt for DebugCalcEngines credentials and store in preferences.json - `var newHash = ( new PasswordHasher<string>().HashPassword( "terry.aney", "password" ) ).Dump();`
```
// Encrypt (hash MacAddress(length64))(encrypted(password))
// Decrypt, verify Mac with hash
// Decrypt, decrypted password and pass it along
// Cryptography3DES.DefaultEncryptAsync
```
1. Save History - Load/Save window position in preferences
1. Implement Ribbon handlers
	1. exportMappedxDSData - Rename this better after you figure out what it is doing
1. [Custom Intellisense/Help Info](https://github.com/Excel-DNA/IntelliSense/issues/21) - read this and linked topic to see what's possible
	1. https://github.com/Excel-DNA/Tutorials/blob/master/SpecialTopics/IntelliSenseForVBAFunctions/README.md
1. [Possible Async Information](https://github.com/Excel-DNA/Samples/blob/master/Registration.Sample/AsyncFunctionExamples.cs)
	1. https://excel-dna.net/docs/guides-advanced/performing-asynchronous-work
	1. https://excel-dna.net/docs/tips-n-tricks/creating-a-threaded-modal-dialog
	1. Excel-DNA/Samples/Archive/Async/NativeAsyncCancellation.dna
1. BTR* functions...
	1. https://excel-dna.net/docs/guides-advanced/dynamic-delegate-registration - possible dynamic creation instead of having to create functions for each item and passing through?  Reference the SSG assembly detect custom functions.
1. Look for email/thread about excel not shutting down properly
1. Readme
	1. Badge count on ribbon image
	1. Dynamic menus
