Hi guys I want to share my first Add-In with you: Document Map v1.7

Name: ScreenshotV1.3.jpg Views: 619 Size: 42.6 KB

It display graph colored representation of the current code window wich can be clicked to navigate from
one point of the code to another instantly
It shows what part of the code you're seeing in the editor.
Its flicker free!
It has checks to not redraw when there is no code/views changes (its fast and not cpu intensive)
Default Colors:
-Grey: Normal code
-Light Grey: Begin of sub/function
-Green: Comments


Left Click jump to that part of code.
Right click brings a menu to add/remove marks:
Cause I dont know how to access the Flags/Breakpoints I added a few special delimiters
that draw visual clues on the Document Map. (Must be the first chars in the line, except TODO)


'*- = Full purple line
'*1 = Red mark on the right
'*2 = Yellow mark on the right
'*3 = Cyan mark on the right
'TODO: = Shorter green mark on the right


(Marks are part of the code so, unlike bookmarks, they are saved)



IMPORTANT: While debuggin AddIns from the IDE I lost functionality of the right mouse (Definition,Last position,etc.) so before
compiling do a backup of your registry key HKCU\Software\Microsoft\Visual Basic\6.0\UI
If this happens to you deleting that key and reopening vb6 restore the lost functionality.
This does not happen if you compile the DLL and use the AddIn normally.


DOWLOAD: vb6_Addin_DocumentMapv1.7b.zip



Changelog
=======
v1.7 (07/06/21)
+Added Codemap default region header/footer
+Improve some colors a little. (more contrast on Function/Sub)
+Fixed End Sub/Function draw with a line (should have been working before)
+No more dependency on Office library
+Selection is highlighted on the document map


v1.6 (06/12/19)
+Added new button to clear inmediate window


v1.5 (20/08/19)
+Right mouse menu integration on VB Ide (to add marks)
+New Toolbar Icon

v1.4 (06/08/19)
+Added a refresh to the UserDocument after redrawing wich make it more responsive using a compiled dll while
moving the mouse with the left button down.
+Recicle AddIn toolbar menu (only useful if youre editing the addin and stop it without removing it from
another proyect.

v1.3 (04/08/19)
+Better drawing representation (with spaces, no more just lines)
+No more double height lines on short documents (double line space instead)


v1.2 (28/07/19)
+Integrated in VB6 Ide (No more a floating on-top form)
+Full analisys for marks (before was skipping lines if code was too big)
+Reworked some logics for faster code


v1.1 (26/07/19)
+Change focus back to project so mousewheel works after clicking document map
+Moving while keeping the mouse pressed allow smooth scrolling through the document map
+Right mouse menu give option to set / remove marks
+Codeview are now centered on the click
+While mouse button held it show the name of Sub / Fuction in that part of code
+Support for CodeSmart bookmarks


v1.0 (24/07/19)
+Basic version as proof of concept (it just 'works') and is my first AddIn. There is a lot
of room for improvement.