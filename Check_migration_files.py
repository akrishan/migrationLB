##Two directories comparison script with version listing into two output files
## version 0.1

import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD  #Referred to VB API from MSDN to retrieve version attributes for dll and exe files 
import csv

def find_folders_difference(sourceDir,targetDir, resultFile):
    f =[]
    d= []
    f2 =[]
    d2 = []
    #Walk through migration-from(source) directory by calling helper function and store in a list
    f,d = get_filenames_dirnames(sourceDir)
    #Walk through migration-To (target) directory    
    f2,d2 = get_filenames_dirnames(targetDir)
    f = set(f)         #convert list to set so 
    d = set(d)    
    f2 = set(f2)
    d2 = set(d2)
    fdiff = f.difference(f2)
    fdiff2 = f2.difference(f)
    writer = open(resultFile, 'w')
    if len(fdiff) > 0:
        writer.writelines("Exceptions Found:\n\n")
        writer.writelines("Note below files are in %s but are not in the %s:\n" % (sourceDir, targetDir))        
        for item in fdiff:
            writer.writelines('%s\n' % item)    
        writer.writelines("\n================================================================\n\n")
    writer.flush()
    if len(fdiff2) > 0:
        writer.writelines("Exceptions Found:\n\n")
        writer.writelines("Note below files are in %s but are not in the %s:\n" % (targetDir,  sourceDir))   
        for item in fdiff2:
            writer.writelines('%s\n' % item)
        writer.writelines("\n================================================================\n\n")        
        writer.close()

#Helper function to retrieve filenames and folder names and save it in lists
def get_filenames_dirnames(pathname):
    fname = []
    dname = []
    for roots,dirs,files in os.walk(pathname):
        for f in files:       ##iterate over all the files in the dir
            fname.append(f)
        for d in dirs:
            dname.append(d)
    return fname,dname   

def do_version_comparison(myPath, myPath2, outfile):    
    writer = open(outfile, 'w')
    writer.writelines("Source Filename,Source File Version,Target Filename,Target File Version\n")
    for root, dirs, files in os.walk(myPath):
        for f in files:
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
    resultFile = r'c:\output\Difference.txt'
    versionFile = r'c:\output\VersionComparison.csv'
    sourceDir = raw_input("Please enter the full path of source dir:")
    targetDir = raw_input("Please enter the full path of source dir:")
    find_folders_difference(sourceDir,targetDir,resultFile)
    do_version_comparison(sourceDir,targetDir,versionFile)



