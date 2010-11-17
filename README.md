<img src="http://github.com/impulzia/kooltuo/raw/master/artwork/kooltou_logo.png" width="150px" height="150px" />

Kooltuo:
--------------------

*** Description ***

This is a Openerp - Outlook syncronizer. Sincronize contacts, task and appointments

How to install
--------------------

To install you need:

1. Install the openerp module (kooltuo_module)
2. Install de adding in Outlook


*** Install the Module *** 

To install the module its simple: 

1. Copy the kooltuo_module in the openerp-server addon folder (or use a symlink)
2. Use the module manager to update you module database 
3. Use the openerp installer system

*** Install the Addin *** 

To install the addin you need to install in your windows

1. [Python 2.6.6](http://www.python.org/ftp/python/2.6.6/python-2.6.6.msi) 32Bit.
2. [pywin32](http://sourceforge.net/projects/pywin32/files/pywin32/Build%20214/pywin32-214.win32-py2.6.exe/download).
3. [wxPython](http://downloads.sourceforge.net/wxpython/wxPython2.8-win32-unicode-2.8.11.0-py26.exe).

After this copy the addin folder in a system folder and use the command:

<pre><code>
c:\> python addin.py
</code></pre>

**Note:** We are working in a py2exe installer to bypass this step

*** Uninstall the Addin *** 

To uninstall the adding in the shell go to the addin folder and execute

<pre><code>
c:\> python addin.py --unregister
</code></pre>

Developing
----------------

We allways appreciate patchs and bugs info. If you want to contribute you can post your bugs in 

[GitHub Issues](https://github.com/impulzia/kooltuo/issues)

or in our group list

[google-groups](http://groups.google.com/group/kooltuo)


*** Debugging de adding ***
To debug the app in the registration time use the --debug option as:

<pre><code>
c:\> python addin.py --debug
</code></pre>

Once made the install use the win32traceutil.py (cames with the win32com) as a listener to see the debug.

If you find a bug please post it with the dump.
