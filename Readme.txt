====================================
SaveMacro for Microsoft Word Readme
====================================

This is a macro for Microsoft Word to save documents with incremented revision numbers in order to prevent loss of work and to easily revert changes.

When saving this will save the document with the original filename.
Everytime it is saved using this macro it will append the revision number to the filename.
E.g. Dissertation - Revision 011 - Jan 6 2011.doc

===================================================
Installation instructions for Microsoft Word 2010:
===================================================

First, we need to enable macros:
Click File -> Options
Click Customize Ribbon on the left.
In the right column, check the 'Developer' check box.
Click OK.

Click the new Developer tab in Word.
Click Macros.
Type in the Macro name "SaveMacro" then click Create.
Delete everything from Sub SaveMacro() to End Sub.

Copy and paste the contents of SaveMacro.txt

Now click on the down arrow from the Quick Access Toolbar, then click More Commands...
Now under where it says 'Choose commands from:" change the drop down box to Macros.
Click on Normal.NewMacros.SaveMacro then click Add.
It will now appear in the right column.
If you want to change the icon, click on the Macro, then click Modify below.
Once done, click OK.