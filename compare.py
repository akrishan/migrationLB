##Two directories comparison script with version listing into two output files
## version 0.1

#import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD  #Referred to VB API from MSDN to retrieve version attributes for dll and exe files 
#import csv
from filecmp import dircmp
import os
import fnmatch
from openpyxl import Workbook



sourceDir = r'c:\release'
targetDir = r'c:\LiberateDev\AutoUpdates'
targetDir2 = r'c:\LiberateDev'
dst_workbook = r'c:\test\output\out.xlsx'


def main():
    #resultFile = r'c:\output\Difference.txt'
    #versionFile = r'c:\output\VersionComparison.csv'

    #find_folders_difference(sourceDir,targetDir,resultFile)
    #do_version_comparison(sourceDir,targetDir,versionFile)
    dcmp = dircmp(sourceDir, targetDir)
    #files_not_in_target(dcmp)
    full_report(dcmp)
    #find_app_version_in_release(dcmp)
    #find_version_in_release_vs_LiberateSE(dircmp(sourceDir,targetDir2))

#Find files which exists in new release but are not in Liberate destination folder 
def files_not_in_target(dcmp):
    for name in dcmp.left_only:
        sourcePath = os.path.join(dcmp.left,name)
        if os.path.isdir(sourcePath):
            print ('"%s" folder found in "%s" and not in "%s"\n\n' % (name, dcmp.left, dcmp.right))
        else:
            print ('"%s" file found in "%s" and not in "%s"\n\n' % (name, dcmp.left, dcmp.right))
            print_version(sourcePath,name)
    for sub_dir in dcmp.subdirs.values():
        files_not_in_target(sub_dir)
'''       
#Find out new version of EXEs in new release
def find_app_version_in_release(dcmp):
    for name in dcmp.left_list:
        sourcePath = os.path.join(dcmp.left,name)
        foldername = os.listdir(sourcePath)
        if os.path.isdir(sourcePath):
            if foldername == 'Liberate SE':
                print foldername
            #pass
            #print
            #print ('"%s" folder in "%s"\n\n' % (name, dcmp.left))
        else:
            #pass
            #print
            #print ('"%s" file in "%s"\n\n' % (name, dcmp.left))
            #print sourcePath
            print_version(sourcePath,name)
            #if sourcePath.lower().endswith('.exe'):
                #print "TRUE"
                #major,minor,subminor,revision = get_version_info(sourcePath)
                #print ("%s - %s.%s.%s.%s" % (name, major,minor,subminor,revision))          
    for sub_dir in dcmp.subdirs.values():
        find_app_version_in_release(sub_dir)
'''

def find_version_in_release_vs_LiberateSE(dcmp):
    for name in dcmp.common:
        sourcePath = os.path.join(dcmp.left,name)
        #foldername = dcmp.left.split("\\")[-1]
        #print foldername
        if os.path.isdir(sourcePath):   
            print name     
        else:
            #pass
            #print
            #print ('"%s" file in "%s"\n\n' % (name, dcmp.left))
            #print sourcePath
            print_version(sourcePath,name)
            if sourcePath.lower().endswith('.exe'):
                #print "TRUE"
                major,minor,subminor,revision = get_version_info(sourcePath)
                print ("%s - %s.%s.%s.%s" % (name, major,minor,subminor,revision))          
    for sub_dir in dcmp.subdirs.values():
        find_version_in_release_vs_LiberateSE((sub_dir))
    return


 #Find files present in both and get their versions     
def find_app_version_in_release(dcmp):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Versions in both"
    ws1.append(("Filename","New Release Version","Liberate Version"))
    for name in dcmp.funny_files:
        sourcePath = os.path.join(dcmp.left,name)
        #foldername = dcmp.left.split("\\")[-1]
        #print foldername
        if os.path.isdir(sourcePath):
            
            folder_list = ['Liberate SE', 'Liberate']
            if name in folder_list:
                continue #skip folders in folder_list               print name               
                #find_version_in_release_vs_LiberateSE(dircmp(sourceDir,targetDir2))
                #continue    
            else:
                
                print name     
        else:
            #pass
            #print
            #print ('"%s" file in "%s"\n\n' % (name, dcmp.left))
            #print sourcePath
            #export_list = []
            targetPath = os.path.join(dcmp.right,name)
            srcfileversion = print_version(sourcePath,name)
            dstfileversion = print_version(targetPath,name)
            export_list = [name,srcfileversion,dstfileversion]
            
            #export_list.append(srcfileversion)
            #export_list.append(dstfileversion[-1])
            ws1.append(export_list)
            #if sourcePath.lower().endswith('.exe'):
                #print "TRUE"
                #major,minor,subminor,revision = get_version_info(sourcePath)
                #print ("%s - %s.%s.%s.%s" % (name, major,minor,subminor,revision))          
    for sub_dir in dcmp.subdirs.values():
        find_app_version_in_release(sub_dir)
    wb.save(filename = dst_workbook)
    return    
        
def print_version(sourcePath, fname):
    if sourcePath.lower().endswith('.exe'):
                #print "TRUE"
        major,minor,subminor,revision = get_version_info(sourcePath)
        ver = ("%s.%s.%s.%s" % (major,minor,subminor,revision))
        return ver
    return 

def full_report(dcmp):
    with open(r'C:\test\output\out.txt','w') as f:
        f.write(dcmp.report_full_closure())
    
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