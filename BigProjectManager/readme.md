<pre>

Treeview based project explorer addin for vb6 IDE
 
Video: https://youtu.be/2CQTjyeyCUA 
 
I want a more powerful tree view project explorer than the built in vb6 IDE one
 
ColinE66 has a great custom drawn one, I cribbed heavily from it thanks Colin!
    https://www.vbforums.com/showthread.php?890617-Add-In-Large-Project-Organiser-(alternative-Project-Explorer)-No-sub-classing!&highlight=
	 
Features:
	- dockable UserDocument mirrors vb6 project treeview on start
    - auto shows for projects which had it open on close
	- only supports a single project, no project groups (rare for me)
	- regroup nodes by drag drop and add arbitrary new folders
    - explicitly control child node layout/position 	
	- auto synced with IDE events adding/renaming/removing files 

	- can save and restore last view to disk
	    - reloaded trees will diff against current IDE files
	    - add files its missing from IDE to default folder for type
	    - mark files removed from the IDE with ? icon
	    - keeps your groupings, position, node state etc (expanded/collapsed)
		
	- add source files & folders to treeview by browse or drag/drop from explorer
    - view file in IDE on double click in tree or find, or right click (source or designer)
    - find dialog allows you to quickly search by name, group, or file name
	- find form allows drag/drop to main tree to reorder
	
</pre>
		
![screenshot](https://github.com/dzzie/addins/blob/master/BigProjectManager/sample.png)
		