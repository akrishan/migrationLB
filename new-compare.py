import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD  #Referred to VB API from MSDN to retrieve version attributes for dll and exe files 
#import csv
from openpyxl import Workbook

sourceDir = r'c:\NewRelease'
targetDir = r'c:\release'
#targetDir = r'c:\LiberateDev\AutoUpdates'
targetDir2 = r'c:\LiberateDev'
dst_workbook = r'c:\test\output\out.xlsx'
wb = Workbook()
ws1 = wb.active
ws1.title = "Files common and version"
ws1.append(("Filename","New Release Version","Current Liberate Version","File Extension"))


def main():
    print "test"
    compare_files()


def relative_files(path):
    """Generate filenames with pathnames relative to the initial path."""
    for root, dirnames, files in os.walk(path):
        relroot = os.path.relpath(root, path)
        for filename in files:
            yield os.path.join(relroot, filename)#fullpath
        
         
def compare_files():          #File comparison function, it calls generator relative_files to  iterate over files found after converting it to sets 
    file1 = set(relative_files(sourceDir))
    for same in file1.intersection(set(relative_files(targetDir))):    #Compare using set operation intersaction and only perform function on matched files
        srcfullpath = os.path.join(sourceDir,same)
        trgfullpath = os.path.join(targetDir,same)
        if os.path.isfile(srcfullpath) and os.path.isfile(trgfullpath):
            srcVersion = print_version(srcfullpath)
            trgVersion = print_version(trgfullpath)
            
            fname,file_extension = os.path.splitext(srcfullpath)   #Get file extensions from filename
            file_extension = file_extension.replace(".","")   #example: Replace .txt to txt
            
            ws1.append((same,srcVersion,trgVersion, file_extension))
    wb.save(filename = dst_workbook)

def print_version(sourcePath):  #returns version number to caller
    if sourcePath.lower().endswith('.exe'):
        major,minor,subminor,revision = get_version_info(sourcePath)
        ver = ("%s.%s.%s.%s" % (major,minor,subminor,revision))
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