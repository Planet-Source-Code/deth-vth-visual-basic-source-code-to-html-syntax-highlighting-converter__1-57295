			*****************
			*     VTH       *
			*****************

The Visual Basic 6 To Html Source Code Syntax Highlighting Converter

- By Lewis Miller -

#About
VTH came about because i needed to create small, short, syntax highlighted, code snippets with a colored background for a help file. After the trial on the shareware version i was using expired, i searched in vain for something to do the job i needed. Either they want to much money, or it didnt do a good job of conversion. So, i decided to write my own. VTH is the result of about two days of solid coding and some debugging. I hope you find it as useful as i do...


#Required support files
1) the visual basic 6 runtime files which you can download here: http://download.microsoft.com/download/5/a/d/5ad868a0-8ecd-4bb0-a882-fe53eb7ef348/VB6.0-KB290887-X86.exe

2) VTH requires you to have Internet Explorer installed, (at least 4.0 i think)

Thats it, VTH has no other dependencies (ocx's etc)


#Support
VTH is free, thus there is little incentive to provide expensive support options, you may however, email me dethbomb at hotmail dot com ( sorry no gmail for me :( ), and i will try to reply as soon as I can.


#How To Convert VB Source Code
Open VTH, Paste the source code into the top window provided and click convert. The results will appear in the bottom window. Click on 'copy source' to copy the html source code to the clipboard.


#How To Convert A VB File
In VTH select the 'Convert File' option on the left and browse for the file you wish to convert in the area provided. Below it, you can enter or select the output file to convert to. The source file must be a valid visual basic source code file. The output file should be text ('.txt' for html source) or html ('.html' for html page). Click on convert. Once completed you can choose to open the folder that the destination file was saved to, or you can preview the source or html page. You can also optionally drop a file on VTH's icon to have it auto convert the file to html. Remember, make sure that the input file is vb source.


#How To Convert A VB Project
Converting an entire project is simple. Click on the 'Convert Project' button in VTH and browse for the project file. Choose an output folder in the area provided and click on 'convert'. Once that is completed, you can create a master index page with links to each converted file by simply clicking on 'create master index'.


#Notes
While i commonly use best coding practices for *speed*, this application has not yet been fine tuned and optimized to get the best speed possible. It will convert about 500 lines of code a second on a pIII 450 with 512 ram. Also, while it will convert everything ive thrown at it, I realize that different coding styles may affect VTH and it may not be able to recognize your own unique code. If that should happen please email me the lines of code that cause failure and I will fix it. Most generally failure will happen if parsing vb file attributes, to resolve the problem, try to covert files without converting the attributes (the default option). Your code has to be syntacticly correct (actually compile) before using VTH. 
 
Currently VTH does a minimal command line operation (more is planned) where, if passed a file path it will attempt to convert it to a file in the same folder as the original file, with the extension ".html". In the future, a full command line capability will be added to include project files.

The VTH source code was a good exercise for me in parsing 'basic' code, something ive been wanting to do for a while. The source is chock full of helper functions, tricks and snippets of vb code, that will make nice additions to anyones 'toolbox' :) VTH is basically a tokenizer, in that it converts symbols it recognizes into named tokens. Each kind of token has a color associated with it and it is a simple matter of spitting out the html once the tokenizer has done its work. modCode contains all the code related to tokenizing and to the app gui. modFile, modIni, and modBrowse all contain file and folder manipulation code. modMain contains the app startup code. I tried to comment as i felt was needed for an intermediate/advanced programmer to be able to follow along, but im sure more commenting wouldnt have hurt... :)

enjoy...


