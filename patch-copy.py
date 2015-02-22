
import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD 

def do_version_comparison(myPath, myPath2, outfile):    
    writer = open(outfile, 'w')
    writer.writelines("Source Filename,Source File Version,Target Filename,Target File Version\n")
    for root, dirs, files in os.walk(myPath):
        for f in files:
            print f
            if os.path.isfile(os.path.join(root,f)):
                f = f.lower() # Convert .EXE to .exe so next line wo
                if (f.count('.exe') or f.count('.dll')): # Check only exe or dll files
                    fullPathToFile  = os.path.join(root,f)
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
                                    sourceversion = [major,minor,subminor,revision]
                                    targetversion = [major2,minor2,subminor2,revision2]
                                    print ("This is a sourceversion %s " % sourceversion)
                                    print ("This is a targetversion %s " % targetversion)
                                    if major > major2:
                                        print "patch-source major version is greater than current target version, apply patch to target"
                                    elif major == major2 and minor > minor2:
                                        print "patch-source minor version is greater than current target version, apply patch to target"
                                    elif major == major2 and minor == minor2 and subminor > subminor2:
                                        print "patch-source subminor version is greater than current target version, apply patch to target"
                                    elif revision > revision2:
                                        print "patch-source major version is greater than current target version, apply patch to target"
                                    else:
                                        "There is no latest patch release available %s for this file" % f
    writer.close()
    
#Helper function to get_version_number function
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

