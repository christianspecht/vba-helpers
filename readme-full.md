![logo](https://bitbucket.org/christianspecht/vba-helpers/raw/tip/img/logo128x128.png)

VBA Helpers is a collection of useful VBA functions.

I developed them for use in MS Access (and I developed and tested them **in** MS Access, too), but most of them should work in other MS Office programs as well.

---

## Links

- [Download vba-helpers.bas directly](https://bitbucket.org/christianspecht/vba-helpers/raw/tip/vba-helpers.bas) (see "Setup" below)
- [Report a bug](https://bitbucket.org/christianspecht/vba-helpers/issues/new)
- [Main project page on Bitbucket](https://bitbucket.org/christianspecht/vba-helpers)

---

## Setup

VBA Helpers consists of a single VBA file, [vba-helpers.bas](https://bitbucket.org/christianspecht/vba-helpers/raw/tip/vba-helpers.bas), which you can just import into your application.  
Right-click and "Save as" to save it on your machine. To import:

- if you are using VBA Helpers for the first time, import the downloaded file from the VBA editor
- if your project already contains VBA Helpers and you want to replace it with the newer version, you can put the new file in the same folder as your Access database and run the `VBAHelpers_Update` function.

---

## Reference

Most of the functions are named after (and do the same like) useful functions in .NET that I missed in VBA.  
Here is a short summary of the available functions and what they do in a nutshell:

- **`Directory_Exists`**  
Returns True if the specified directory exists.
- **`File_Delete`**  
Deletes a file. If the file does not exist, nothing happens.
- **`File_Exists`**  
Returns True if the specified file exists.
- **`File_ReadAllLines`**  
Reads a text file and returns a string array, each array item containing a line from the file.
- **`File_ReadAllText`**  
Reads a text file and returns the content in a string variable.
- **`File_WriteAllLines`**  
Writes the content of a string array into a text file, each array item into a new line.
- **`File_WriteAllText`**  
Writes the content of a string variable into a text file.
- **`Path_Combine`**  
Combines several strings into a path and takes care of directory separators, i.e. `path_combine("c:\","\foo","bar")` will return `c:\foo\bar`
- **`Path_GetCurrentDirectory`**  
Returns the directory of the current Access database.
- **`Path_GetDirectoryName`**  
Receives a complete path, returns only the directory.
- **`Path_GetFileName`**  
Receives a complete path, returns only the file name.
- **`Path_GetFileNameWithoutExtension`**  
Receives a complete path, returns only the file name without extension.
- **`String_Contains`**  
Returns `True` if the second parameter occurs within the first parameter.
- **`String_EndsWith`**  
Returns `True` if the second parameter matches the end of the first parameter.
- **`String_Format`**  
Replaces numbered placeholders (`{0}`, `{1}`, ...) in the first parameter by the corresponding value from the additional parameter list.
- **`String_PadLeft`**  
Right-aligns the first string parameter by padding it on the left with the second string parameter, up to the total specified width.  
Example: `String_PadLeft("foo",5,"a")` will return `aafoo`
- **`String_PadRight`**  
Left-aligns the first string parameter by padding it on the right with the second string parameter, up to the total specified width.  
Example: `String_PadRight("foo",5,"a")` will return `fooaa`
- **`String_StartsWith`**  
Returns `True` if the second parameter matches the beginning of the first parameter.
- **`VBAHelpers_Update`**  
Updates VBA Helpers to newer version by importing a downloaded file (file must be in same folder as current Access database). See "Setup" above for more information.

---

## Development

### Coding Guidelines

Unfortunately, [VBA globally changes the case of variable names when you mix upper/lower case](http://stackoverflow.com/q/4852735).  
This is **very** annoying when using source control.  
It's even worse when VBA Helpers is imported into another VBA project, and the case of *the variables in this project* is changed because some of them happened to have the same names like some of the VBA Helpers variables.

To avoid this, all variable names in VBA Helpers must adhere to the following guidelines:

- lower case only
- suffixed by `_vbah` (for "VBA Helpers"), e.g. `foo_vbah`


### Committing changes

Run `VBAHelpers_Export` to export all modules to the current directory, and commit them from there.

---

### Acknowledgements

VBA Helpers makes use of the following open source projects:

- [AccUnit](http://accunit.access-codelib.net/) (which uses [SimplyVBUnit](http://sourceforge.net/projects/simplyvbunit/))

---

### License

VBA Helpers is licensed under the MIT License. See [License.txt](https://bitbucket.org/christianspecht/vba-helpers/raw/tip/license.txt) for details.

---

### Project Info

<script type="text/javascript" src="http://www.ohloh.net/p/603791/widgets/project_basic_stats.js"></script>
