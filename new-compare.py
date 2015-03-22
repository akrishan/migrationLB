import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD  #Referred to VB API from MSDN to retrieve version attributes for dll and exe files 
#import csv
from openpyxl import Workbook

sourceDir = r'c:\NewRelease'
#targetDir = r'c:\release'
targetDir = r'c:\LiberateDev\AutoUpdates'
targetDir2 = r'c:\LiberateDev'
dst_workbook = r'c:\test\output\out.xlsx'
wb = Workbook()             #Create a workbook to store version comparison data in an output
ws1 = wb.active
ws1.title = "Versions Comparison"
ws1.append(("Filename","New Release Version","Current Liberate Version","File Extension","New Release Filesize","Current Liberate Filesize","KB/MB"))
ws2 = wb.create_sheet(title="New Files")
ws2.append(("Filename","New Release Version","File Extension","New Release Filesize","KB/MB"))

def main():
    print "test"
    compare_files_and_versions()


def relative_files(path):
    #Generate filenames with pathnames relative to the initial path.
    for root, dirnames, files in os.walk(path):
        relroot = os.path.relpath(root, path)
        for filename in files:
            yield os.path.join(relroot, filename)
        
         
def compare_files_and_versions():          #File comparison function, it calls generator relative_files to  iterate over files found after converting it to sets 
    file1 = set(relative_files(sourceDir))
    for same in file1.intersection(set(relative_files(targetDir))):    #Compare using set operation intersaction and only perform function on matched files
        srcfullpath = os.path.join(sourceDir,same)
        trgfullpath = os.path.join(targetDir,same)
        if os.path.isfile(srcfullpath) and os.path.isfile(trgfullpath):
            srcVersion = print_version(srcfullpath) #get version for New Release file
            trgVersion = print_version(trgfullpath)  #get version for Liberate file          
            fname,file_extension = os.path.splitext(srcfullpath)   #Get file extensions from filename
            file_extension = file_extension.replace(".","")   #example: Replace .txt to txt
            src_file_size = round(float(os.stat(srcfullpath).st_size/1024.00),2)  #get file size in KB
            trg_file_size = round(float(os.stat(trgfullpath).st_size/1024.00),2)
            ws1.append((same,srcVersion,trgVersion,file_extension,src_file_size,trg_file_size,"KB"))
            adjust_column_width(ws1) #Adjust columns width for content to be visible upon workbook opening            
    wb.save(filename = dst_workbook)    
    #Find unmatched files in New release after comparing against the current liberate dir
    for diff in file1.difference(set(relative_files(targetDir))):
        #print diff
        srcfullpath = os.path.join(sourceDir,diff)        
        if os.path.isfile(srcfullpath):
            print diff
            srcVersion = print_version(srcfullpath) #get version for New Release file        
            fname,file_extension = os.path.splitext(srcfullpath)   #Get file extensions from filename
            file_extension = file_extension.replace(".","")   #example: Replace .txt to txt
            src_file_size = round(float(os.stat(srcfullpath).st_size/1024.00),2)  #get file size in KB
            ws2.append((diff,srcVersion,file_extension,src_file_size,"KB"))
            adjust_column_width(ws2)
    wb.save(filename = dst_workbook)
    

def adjust_column_width(sheetname):
    sheetname.column_dimensions["A"].width = 60.0
    sheetname.column_dimensions["B"].width = 20.0
    sheetname.column_dimensions["C"].width = 23.0
    return


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