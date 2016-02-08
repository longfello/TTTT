Well, Here it is. Привет! И мы не лыком шиты.

Steps to follow
1) Copy the core and TutorialDatabase folders to you C:
2) Copy the website folder (or just its contents) to your
   Inetpub/wwwroot folder
3) In IIS create a virtual directory named 'core' (without the single quotes)
   and link it to your C:\core directory.
4) Right click the TutorialDatabase folder, Select Properties,
   Click the Security tab, add the IUSR_xxx (where xxx is your computers name)
   to the list (if it isn't already there) and give it write permissions
   on this folder (do the same for your Temp directory, located under
   C:\WINDOWS or C:\WINNT)
5) Right click the Tutorial.mdb database and make sure the IUSR_xxx
   account also has write permissions.


This sample website is meant as a learning tool for new or inexperienced
ASP programmers and web developers. It is not meant to be used commercially.

This website is designed using templates. A precursor to ASP.NET's 
code-behind design strategy.

If you have any concerns or questions about this sample site,
or if you would like additional functionality included,
feel free to contact me at sjensen1207@hotmail.com

I would like to acknowledge Valerio Santinelli for allowing
the use of his aspTemplate class.





