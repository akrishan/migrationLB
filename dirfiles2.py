##Two directories comparison script with version listing into two output files - ResultFile.txt and VersionComparison.csv
##
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
    wtf.writelines("File 1 Name,File 1 Version,File 2 Name,File 2 Version\n")
    for root, dirs, files in os.walk(myPath):
        for f in files:
            if os.path.isfile(f):
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
                                    wtf.writelines("%s,%s.%s.%s.%s,%s,%s.%s.%s.%s\n" % (fullPathToFile,major,minor,subminor,revision,fullPathToFile2,major2,minor2,subminor2,revision2))
    wtf.close()
    
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
    resultFile = r'c:\output\ResultFile.txt'
    versionFile = r'c:\output\VersionComparison.csv'
    sourceDir = raw_input("Please enter the full path of source dir:")
    targetDir = raw_input("Please enter the full path of source dir:")
    findDirDifference(sourceDir,targetDir,resultFile)
    get_version_number(sourceDir,targetDir,versionFile)



