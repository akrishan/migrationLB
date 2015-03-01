
import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD
import shutil
import sys

def do_version_comparison(myPath, myPath2, outfile):    
    writer = open(outfile, 'w')
    writer.writelines("Source Filename,Source File Version,Target Filename,Target File Version\n")
    for root, dirs, files in os.walk(myPath):
        for f in files:
            #print f
            if os.path.isfile(os.path.join(root,f)):
                f = f.lower() # Convert .EXE to .exe so next line wo
                if (f.count('.exe') or f.count('.dll')): # Check only exe or dll files
                    fullPathToFile  = os.path.join(root,f)
                    print f 
                    major,minor,subminor,revision= get_version_info(fullPathToFile)
                    for root2, dirs2, files2 in os.walk(myPath2):
                        for f2 in files2:
                            f2 = f2.lower() #Convert .EXE to .exe so next line works
                    #if (file.count('.exe') or file.count('.dll')): # Check only exe or dll files
                            if (f2.count('.exe') or f2.count('.dll')):
                                fullPathToFile2  = os.path.join(root2,f2)
                                major2,minor2,subminor2,revision2= get_version_info(fullPathToFile2)
                                if f == f2:
                                    writer.writelines("%s,%s.%s.%s.%s,%s,%s.%s.%s.%s\n" % (fullPathToFile,major,minor,subminor,revision,fullPathToFile2,major2,minor2,subminor2,revision2))
                                    print ("%s,%s.%s.%s.%s,%s,%s.%s.%s.%s\n" % (fullPathToFile,major,minor,subminor,revision,fullPathToFile2,major2,minor2,subminor2,revision2))    
                                    sourceversion = map(int, [major,minor,subminor,revision])
                                    targetversion = map(int,[major2,minor2,subminor2,revision2])
                                    print ("This is a sourceversion %s " % sourceversion)
                                    print ("This is a targetversion %s " % targetversion)                                    
                                    if (major > major2):
                                        print "patch-source major version is greater than current target version, apply patch to target"
                                        copy_patch_version()
                                    elif (major == major2 and minor > minor2):
                                        print "patch-source minor version is greater than current target version, apply patch to target"
                                        copy_patch_version()
                                    elif (major == major2 and minor == minor2 and subminor > subminor2):                                        
                                        print "patch-source subminor version is greater than current target version, apply patch to target"
                                        copy_patch_version()
                                    elif (major == major2 and minor == minor2 and subminor == subminor2 and revision > revision2):
                                        print "patch-source revision version is greater than current target version, apply patch to target"
                                        copy_patch_version()
                                else:
                                    print "There is no latest patch release available for %s this file" % f
    writer.close()

def copy_patch_version():
    userprompt = "Would you like to copy %s to %s " % fullPathToFile, fullPathToFile2
    if userprompt == "Y" or userprompt == "y":
        shutil.copy2(fullPathToFile,fullPathToFile2)
    else:
        print "You have selected to exit the program"
        sys.exit() 
    

#Helper function
def get_version_info(filename):
    try:
    	info = GetFileVersionInfo(filename, "\\")
    	ms = info['FileVersionMS']
    	ls = info['FileVersionLS']
    	return HIWORD (ms), LOWORD (ms), HIWORD (ls), LOWORD (ls)
    except:
    	return 0,0,0,0   


if __name__ == '__main__':
    versionFile = r'c:\output\VersionComparison-test.csv'
    sourceDir = r'c:\patch-store'
    targetDir = r'c:\LiberateDev'
    #print (sourceDir,targetDir,resultFile)
    do_version_comparison(sourceDir,targetDir,versionFile)

