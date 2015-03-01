from bbfreeze import Freezer
 
f = Freezer(r"c:\Pyscripts\migrationLB\patch-copy.exe")
f.addScript(r"c:\Pyscripts\migrationLB\patch-copy.py")
f()