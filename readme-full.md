
VBA Helpers is a collection of useful VBA functions.

I developed them for use in MS Access (and I developed and tested them **in** MS Access, too), but most of them should work in other MS Office programs as well.

---

## Links

- [Download vba-helpers.bas directly](https://bitbucket.org/christianspecht/vba-helpers/raw/tip/vba-helpers.bas) (see "Setup" below)
- [Found a bug?](https://bitbucket.org/christianspecht/vba-helpers/issues/new)
- [Main project page on Bitbucket](https://bitbucket.org/christianspecht/vba-helpers)

---

## Setup

VBA Helpers consists of a single VBA file, [vba-helpers.bas](https://bitbucket.org/christianspecht/vba-helpers/raw/tip/vba-helpers.bas), which you can just import into your application.  
Right-click and "Save as" to save it on your machine, then import it from the VBA editor if you are using it for the first time.  

To replace an older version of VBA Helpers, put the new file in the same folder as your Access database and run the `VBAHelpers_Import` function.

---

## Reference

Most of the functions are named after (and do the same like) useful functions in .NET that I missed in VBA.  
Here is a short summary of the available functions and what they do in a nutshell:

- `File_ReadAllLines`  
Reads a text file and returns a string array, each array item containing a line from the file.
- `File_ReadAllText`  
Reads a text file and returns the content in a string variable.
- `File_WriteAllLines`  
Writes the content of a string array into a text file, each array item into a new line.
- `File_WriteAllText`  
Writes the content of a string variable into a text file.
- `Path_Combine`  
Combines several strings into a path and takes care of directory separators, i.e. `path_combine("c:\","\foo","bar")` will return `c:\foo\bar`
- `Path_GetCurrentDirectory`  
Returns the directory of the current Access database.
- `Path_GetDirectoryName`  
Receives a complete path, returns only the directory.
- `Path_GetFileName`  
Receives a complete path, returns only the file name.
- `Path_GetFileNameWithoutExtension`  
Receives a complete path, returns only the file name without extension.
- `String_EndsWith`  
Returns `True` if the second parameter matches the end of the first parameter.
- `String_Format`  
Replaces numbered placeholders (`{0}`, `{1}`, ...) in the first parameter by the corresponding value from the additional parameter list.
- `String_StartsWith`  
Returns `True` if the second parameter matches the beginning of the first parameter.
- `VBAHelpers_Export`  
Exports the VBA Helpers module to the current directory and increases the version number.  
Useful for development only (see "Committing changes" below)
- `VBAHelpers_Import`  
Imports a new version of the VBA Helpers module from the current directory.  
Useful for updating to a newer version (see "Setup" above)

---

### Development

##### Coding Guidelines

All variable names must be in lower case, to avoid [VBA changing the case automatically when mixing upper/lower case](http://stackoverflow.com/q/4852735) - this is **very** annoying when using source control.

##### Committing changes

Run `VBAHelpers_Export` to export the actual VBA Helpers module and the module with the tests to the current directory, and commit them from there.

---

### Acknowledgements

VBA Helpers makes use of the following open source projects:

- [AccUnit](http://accunit.access-codelib.net/) (which uses [SimplyVBUnit](http://sourceforge.net/projects/simplyvbunit/))

---

### License

VBA Helpers is licensed under the MIT License. See [License.txt](https://bitbucket.org/christianspecht/vba-helpers/raw/tip/license.txt) for details.