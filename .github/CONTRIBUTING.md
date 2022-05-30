# How to contribute to PEMEditor

## Bug report?
- Please check  [issues](https://github.com/VFPX/PEMEditor/issues) if the bug is reported
- If you're unable to find an open issue addressing the problem, open a new one. Be sure to include a title and clear description, as much relevant information as possible, and a code sample or an executable test case demonstrating the expected behavior that is not occurring.

### Did you write a patch that fixes a bug?
- Open a new GitHub pull request with the patch.
- Ensure the PR description clearly describes the problem and solution.
  - Include the relevant version number if applicable.
- See [New version](#new-version) for additional tasks

## Coding conventions
Start reading our code and you'll get the hang of it. We optimize for readability:

- Beautification is done like:
  - Keywords: Mixed case 
  - Symbols: First occurence
  - Indentation Tabs, 1
  - Indent anything then Comments
- Please do not run BeautifyX with mDots insertion against the code. 
- We ALWAYS put spaces after list items and method parameters (`[1, 2, 3]`, not `[1,2,3]`), around operators (`x = 1`, not `x=1`).
- This is open source software. Consider the people who will read your code, and make it look nice for them. It's sort of like driving a car: Perhaps you love doing donuts when you're alone, but with passengers the goal is to make the ride as smooth as possible.
- Please kindly add comments where and what you change

## Important
Why ever, the PEMEditor repository is not self contained. The code is scattered over 3 repositories.
- The code doing the work: https://github.com/VFPX/PEMEditor
- Settings: https://github.com/VFPX/ThorRepository (different files for different parts of PEMEditor)
- Version of ThorRepository: File _Thor_Update_Thor_Repository.prg_ of https://github.com/VFPX/Thor/blob/master/ThorUpdater/Updates.zip
### Most important
Compile the _downloads\Source\pemeditor.pjx_ to _downloads\pemeditor.app_ **Do not copy, set the folder in compile!**   

## New version
Please note, there are some tasks to set up a new version.
Stuff is a bit scattered, so this is where to look up.
1. Please create a fork at github
   - See this [guide](https://www.dataschool.io/how-to-contribute-on-github/) for setting up and using a fork
1. Make whatever changes are necessary.
2. Run FoxBin2Prg to create the text files.
   - `DO foxbin2prg.prg WITH 'BIN2PRG','downloads\source'`
2. In _downloads\source\pemeditorversion.h_ set up the version number as well as the date of version. Please change all three DEFINES.   
3. Alter version in _README.md_, add description of the change.
4. On major changes add a description to the _docs_ folder, see _Documentation/pemeditor_thor_changes_7_40.md_ as example.
5. If a change to the config is made, update https://github.com/VFPX/ThorRepository and file _Thor_Update_Thor_Repository.prg_ of https://github.com/VFPX/Thor/blob/master/ThorUpdater/Updates.zip
   - see https://github.com/VFPX/ThorRepository for contribution.
   - see https://github.com/VFPX/Thor for contribution.
   - see that PEMEDitor, ThorRepository and Thor are pushed as close as possible. The order how to publish is on the maintainers of Thor.
6. Please alter the footer of \*.md files touched to recent date.
7. **Repeat:** Compile the _downloads\Source\pemeditor.pjx_ to _downloads\pemeditor.app_ **Do not copy, set the folder in compile!**
8. Change Thor ([see below](#thor-conventions))
8. commit
9. push to your fork
0. create a pull request

## Thor conventions
This project is part of [VFPX](https://vfpx.github.io/) and published via [Thor](https://github.com/VFPX/Thor).   
There are some considerations to make to add a new version to Thor.   
In Folder _downloads\ThorUpdater_
- _PEMEditorVersionFile.txt: set up version information and brief list of changes
- PEM Editor Version.txt: set up version information
- PEMEditor_Source.zip
  - copy PEM Editor Version.txt and pemeditor.app
  - replace the folders touched (since this is a binary repository all files of _source_ folder are changing on compile ...)

Thanks

----
Last changed: _2022/04/03_ ![Picture](../docs/pictures/vfpxpoweredby_alternative.gif)
