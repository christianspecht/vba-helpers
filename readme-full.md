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
*(if Access asks you for a module name, name it `VBAHelpers`)*
- if your project already contains VBA Helpers and you want to replace it with the newer version, you can put the new file in the same folder as your Access database and run the `VBAHelpers_Update` function.

---

## Reference

Most of the functions are named after (and do the same like) useful functions in .NET that I missed in VBA.  
Here is a short summary of the available functions and what they do in a nutshell:

- **`Directory_Exists`**  
Returns `True` if the specified directory exists.

- **`Environment_AccessVersion`**  
Returns an `Enum` which contains the version of the current `msaccess.exe`.

- **`Environment_MachineName`**  
Returns the name of the local computer.

- **`File_Delete`**  
Deletes a file. If the file does not exist, nothing happens.

- **`File_Exists`**  
Returns `True` if the specified file exists.

- **`File_ReadAllLines`**  
Reads a text file and returns a string array, each array item containing a line from the file.

- **`File_ReadAllText`**  
Reads a text file and returns the content in a string variable.

- **`File_WriteAllLines`**  
Writes the content of a string array into a text file, each array item into a new line.

- **`File_WriteAllText`**  
Writes the content of a string variable into a text file.

- **`InputBox_PressedCancel`**  
Receives the return value of an `InputBox`, returns `True` when the input was canceled.  
Normally you can't distinguish whether you cancelled the input or submitted an empty string - the `InputBox` returns an empty string in both cases.  
Example: `InputBox_PressedCancel(InputBox("foo"))` returns `True` when you press Cancel, and `False` when you press OK without entering a value.

- **`Path_Combine`**  
Combines several strings into a path and takes care of directory separators.  
Example: `path_combine("c:\","\foo","bar")` will return `c:\foo\bar`

- **`Path_GetCurrentDirectory`**  
Returns the directory of the current Access database.

- **`Path_GetDirectoryName`**  
Receives a complete path, returns only the directory.

- **`Path_GetFileName`**  
Receives a complete path, returns only the file name.

- **`Path_GetFileNameWithoutExtension`**  
Receives a complete path, returns only the file name without extension.

- **`Path_GetTempPath`**  
Returns the current user's temp folder.

- **`Process_Start`**  
Executes a file. If the file itself is not an application, it will be started with the default application *(as if you double-clicked it in Windows Explorer)*.  
Use the optional parameters to supply command-line arguments to the executed file, and to open the file hidden (without a visible window - useful for executing command-line tools)

- **`String_Contains`**  
Returns `True` if the second parameter occurs within the first parameter.  
Example: `String_Contains("abc", "ab")` will return `True`

- **`String_EndsWith`**  
Returns `True` if the second parameter matches the end of the first parameter.  
Example: `String_EndsWith("abc", "bc")` will return `True`

- **`String_Format`**  
Replaces numbered placeholders (`{0}`, `{1}`, ...) in the first parameter by the corresponding value from the additional parameter list.  
Example: `String_Format("Hello {0}", "world")` will return `Hello world`

- **`String_IsNullOrEmpty`**  
Returns True when the input is either Null or an empty string.  
*(note: a VBA string can't be Null, but the function is called `String_` anyway to keep the naming consistent)*

- **`String_IsNullOrWhiteSpace`**  
Returns True when the input is either Null, an empty string or consists of whitespace characters (blanks) only.  
*(note: a VBA string can't be Null, but the function is called `String_` anyway to keep the naming consistent)*

- **`String_PadLeft`**  
Right-aligns the first string parameter by padding it on the left with the second string parameter, up to the total specified width.  
Example: `String_PadLeft("foo",5,"a")` will return `aafoo`

- **`String_PadRight`**  
Left-aligns the first string parameter by padding it on the right with the second string parameter, up to the total specified width.  
Example: `String_PadRight("foo",5,"a")` will return `fooaa`

- **`String_StartsWith`**  
Returns `True` if the second parameter matches the beginning of the first parameter.  
Example: `String_StartsWith("abc", "ab")` will return `True`

- **`VBAHelpers_Update`**  
Updates VBA Helpers to newer version by importing a downloaded file (file must be in same folder as current Access database). See "Setup" above for more information.

---

## Development

### Importing the modules into a `.mdb` for the first time

1. If you are starting from scratch, you need to import the three `.bas` files from the repository's into an Access database first.  
**Important:** `VBAHelpersTests` needs to be a **class module**. The other two need to be **"regular" modules**.
2. Add references to `AccUnit Access/VBA TestSuite` and `SimplyVBUnit Framework 3.0`.


### Coding Guidelines

Unfortunately, [VBA globally changes the case of variable names when you mix upper/lower case](http://stackoverflow.com/q/4852735/6884).  
This is **very** annoying when using source control.  
It's even worse when VBA Helpers is imported into another VBA project, and the case of *the variables in this project* is changed because some of them happened to have the same names like some of the VBA Helpers variables.

To avoid this, all variable names in VBA Helpers must adhere to the following guidelines:

- lower case only
- suffixed by `_vbah` (for "VBA Helpers"), e.g. `foo_vbah`

This should minimize the chance of variable names from VBA Helpers colliding with variable names in your application.
 

### Committing changes

Run `VBAHelpers_Export` to export all modules to the current directory, and commit them from there.

---

### Acknowledgements

VBA Helpers makes use of the following open source projects:

- [AccUnit](http://accunit.access-codelib.net/) (which uses [SimplyVBUnit](http://sourceforge.net/projects/simplyvbunit/))

<a name="license"></a>

---

### License

VBA Helpers is licensed under the MIT License. See [License.txt](https://bitbucket.org/christianspecht/vba-helpers/raw/tip/license.txt) for details.

---

### Project Info

<script type="text/javascript" src="http://www.ohloh.net/p/603791/widgets/project_basic_stats.js"></script>
