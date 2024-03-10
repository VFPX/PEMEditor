# How to contribute to PEMEditor

## Bug report?
- Please check [issues](https://github.com/VFPX/PEMEditor/issues) to see if the bug has already been reported.
- If you're unable to find an open issue addressing the problem, open a new one. Be sure to include a title and clear description, as much relevant information as possible, and a code sample or an executable test case demonstrating the expected behavior that is not occurring.

## Fix a bug or add an enhancement

### Coding conventions
Start reading our code and you'll get the hang of it. We optimize for readability:

- Beautification is done like:
  - Keywords: Mixed case
  - Symbols: First occurrence
  - Indentation: Tabs, 1
  - Indent anything then Comments

- Please do not run BeautifyX with mDots insertion against the code. 

- We ALWAYS put spaces after list items and method parameters (`[1, 2, 3]`, not `[1,2,3]`), around operators (`x = 1`, not `x=1`).

- This is open source software. Consider the people who will read your code, and make it look nice for them. It's sort of like driving a car: Perhaps you love doing donuts when you're alone, but with passengers the goal is to make the ride as smooth as possible.

- Please kindly add comments where and what you change.

### Important
The PEMEditor repository is not self contained. The code is scattered over a couple of repositories:

- PEM Editor itself: https://github.com/VFPX/PEMEditor

- Settings: https://github.com/VFPX/ThorRepository (different files for different parts of PEMEditor)

### New version

- Fork the project: see this [guide](https://www.dataschool.io/how-to-contribute-on-github/) for setting up and using a fork.

- Make whatever changes are necessary. The source code is in Downloads\Source, including PEMEditor.pjx.

- Edit the **Version** setting in *BuildProcess\ProjectSettings.txt*.

- Describe the changes at the top of _ChangeLog.md_.

- For major changes, you may wish to add a Markdown file documenting the new features to the _Documentation_ folder. See _Documentation/pemeditor_thor_changes_7_40.md_ as an example.

- If a change to the configuration setting is made, update https://github.com/VFPX/ThorRepository and _Thor_Update_Thor_Repository.prg_ of https://github.com/VFPX/Thor/blob/master/ThorUpdater/Updates.zip
   - see https://github.com/VFPX/ThorRepository for contribution.
   - see https://github.com/VFPX/Thor for contribution.
   - see that PEMEditor, ThorRepository and Thor are pushed as close as possible. The order how to publish is on the maintainers of Thor.

- If you haven't already done so, install VFPX Deployment: choose Check for Updates from the Thor menu, turn on the checkbox for VFPX Deployment, and click Install.

- Start VFP 9 (not VFP Advanced) and CD to the PEMEditor root folder (not the Downloads folder). **Do not run with PJX open, it needs at least a CLOSE ALL**

- Run the VFPX Deployment tool to create the installation files: choose VFPX Project Deployment from the Thor Tools, Application menu. Alternately, execute ```EXECSCRIPT(_screen.cThorDispatcher, 'Thor_Tool_DeployVFPXProject')```.

- Commit the changes.

- Push to your fork.

- Create a pull request; ensure the description clearly describes the problem and solution or the enhancement.

----
Last changed: _2023/06/07_ ![Picture](../docs/pictures/vfpxpoweredby_alternative.gif)