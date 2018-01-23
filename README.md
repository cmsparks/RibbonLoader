# RibbonLoader
Do not use this project unless you want massive migranes.
Also the template document bricks your Word Version.
Loads custom .officeUI files and configures templates in Microsoft Office 2007+.
## Notes
### Config Files
Config files are a simple zip file with up to 4 components:
* a Normal.dotm file containing macros
* a customizations.xml file containing keybinds for macros
* a styles.xml file containing heading styles
* a Word.officeUI file containing XML for the customUI
The python script will load these files into the correct places (once the script is done)

The Test_Configuration is the modified verbatim configuration.
The Normal_Configuration is the default configuration to revert to word's original state.

### Source and build files
The src file is where the rewritten verbatim source can be viewed. The Normal.dotm file is a prebuilt OLE thing containing the macros which can be loaded into Word.

Unfortunately Microsoft has a terribly obfuscated file format for containing macros so files must be manually copied and pasted in to the VBA IDE if you want to build it on your own. Also the VBA Extensibility Library 5.3 is completely broken so you cannot dynamically load plaintext .bas files. There is no way to automatically build files unless you want to create a OLE compiler.