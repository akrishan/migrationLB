import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD  #Referred to VB API from MSDN to retrieve version attributes for dll and exe files 
import csv

def findDirDifference(sourceDir,targetDir, resultFile):
    f=[]
    d= []
    f2 =[]
    d2 = []
    for roots,dirs,files in os.walk(sourceDir):
        for file in files:
            f.append(file)        
        for dir in dirs:
            d.append(dir)            
    for roots2,dirs2,files2 in os.walk(targetDir):
        for file2 in files2:
            f2.append(file2)        
        for dir2 in dirs2:
            d2.append(dir2)
    f = set(f)
    d = set(d)    
    f2 = set(f2)
    d2 = set(d2)
    fdiff = f.difference(f2)
    print d.difference(d2)
    fdiff2 = f2.difference(f)
    wtf = open(resultFile, 'w')
    if len(fdiff) > 0:
        wtf.writelines("Exceptions Found:\n\n")
        wtf.writelines("Note below files are in %s but are not in the %s:\n" % (sourceDir, targetDir))        
        for item in fdiff:
            wtf.writelines('%s\n' % item)    
        wtf.writelines("\n================================================================\n\n")
    wtf.flush()
    if len(fdiff2) > 0:
        wtf.writelines("Exceptions Found:\n\n")
        wtf.writelines("Note below files are in %s but are not in the %s:\n" % (targetDir,  sourceDir))   
        for item in fdiff2:
            wtf.writelines('%s\n' % item)
        wtf.writelines("\n================================================================\n\n")        
        wtf.close()

def get_version_number(myPath, myPath2, outfile):    
    wtf = open(outfile, 'w')
    fullPathToFile1,filename  = split_root_and_files(myPath) #Call refactored helper function for directory1
    fullPathToFile2,filename2  = split_root_and_files(myPath2) #Call refactored helper function for directory2
    major,minor,subminor,revision= get_version_info(fullPathToFile1) #Call helper function to get actual version info
    major2,minor2,subminor2,revision2= get_version_info(fullPathToFile2) #Call helper function to get actual version info
    
    wtf.writelines("%s \t %s.%s.%s.%s" % (filename,major,minor,subminor,revision))
    wtf.close()
    
def split_root_and_files(path):
    for root, dirs, files in os.walk(Path):
        for file in files:
            file = file.lower() # Convert .EXE to .exe so next line works
            if (file.count('.exe') or file.count('.dll')): # Check only exe or dll files
                AbsPathToFile  = os.path.join(root,file)
                return AbsPathToFile, file

##Helper function to retrieve version info and return to get_version_number() function
def get_version_info(filename):
    try:
    	info = GetFileVersionInfo(filename, "\\")
    	ms = info['FileVersionMS']
    	ls = info['FileVersionLS']
    	return HIWORD (ms), LOWORD (ms), HIWORD (ls), LOWORD (ls)
    except:
    	return 0,0,0,0        
 

if __name__ == '__main__':
    resultFile = r'c:\output\ResultFile.txt'
    versionFile = r'c:\output\VersionComparison.txt'
    sourceDir = raw_input("Please enter the full path of source dir:")
    targetDir = raw_input("Please enter the full path of source dir:")
    findDirDifference(sourceDir,targetDir, resultFile)
    get_version_number(sourceDir,versionFile)



