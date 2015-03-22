import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD  #Referred to VB API from MSDN to retrieve version attributes for dll and exe files 
#import csv

sourceDir = r'c:\NewRelease'
targetDir = r'c:\LiberateDev\AutoUpdates'
targetDir2 = r'c:\LiberateDev'
dst_workbook = r'c:\test\output\out.xlsx'

def main():
    print "test"
    compare_files()


def relative_files(path):
    """Generate filenames with pathnames relative to the initial path."""
    for root, dirnames, files in os.walk(path):
        relroot = os.path.relpath(root, path)
        #print relroot
        for filename in files:
            fullpath = os.path.join(root,filename)
            #print filename
            yield os.path.join(relroot, filename), fullpath
  
def target_relative_files(path):
    """Generate filenames with pathnames relative to the initial path."""
    for root, dirnames, files in os.walk(path):
        relroot = os.path.relpath(root, path)
        #print relroot
        for filename in files:
            fullpath = os.path.join(root,filename)
            #print filename
            yield os.path.join(relroot, filename), fullpath     
            
            
            
def compare_files():
    
    
    #for f in relative_files(r'c:\release'):
        #print f
    set1 = set(relative_files(sourceDir))
    file1 = set(x[0] for x in set1)
    file1fullpath = ([x[1] for x in set1])
    
    #print file1
    #print file1fullpath[0]
    
    set2 = set(relative_files(targetDir))
    file2 = set(x[0] for x in set2)
    file2fullpath = [x[1] for x in set2]
    

    #print file2
    #print file2fullpath
    #print file1path
    for same in file1.intersection(set(x[0] for x in target_relative_files(targetDir))):
        #print same
        srcVersion = print_version(same)
        #for f_in_target in file2.intersection(set):
            
        
        for same2 in file2.intersection(file1):
            #if same == same2:
                #targetVersion = print_version(file2fullpath[0])
        #dstVersion = print
                print "%s,%s" % (same, srcVersion)

def print_version(sourcePath):
    if sourcePath.lower().endswith('.exe'):
                #print "TRUE"
        major,minor,subminor,revision = get_version_info(sourcePath)
        ver = ("%s.%s.%s.%s" % (major,minor,subminor,revision))
        #print ver
        return ver
    return         
    


def get_version_info(filename):
    try:
    	info = GetFileVersionInfo(filename, "\\")
    	ms = info['FileVersionMS']
    	ls = info['FileVersionLS']
    	return HIWORD (ms), LOWORD (ms), HIWORD (ls), LOWORD (ls)
    except:
    	return 0,0,0,0
    
    
if __name__ == '__main__':
    main()