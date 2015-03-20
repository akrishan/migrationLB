##Two directories comparison script with version listing into two output files
## version 0.1

#import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD  #Referred to VB API from MSDN to retrieve version attributes for dll and exe files 
#import csv
from filecmp import dircmp
import os
import fnmatch

    
def main():
    #resultFile = r'c:\output\Difference.txt'
    #versionFile = r'c:\output\VersionComparison.csv'
    sourceDir = r'c:\release'
    targetDir = r'c:\LiberateDev'
    #find_folders_difference(sourceDir,targetDir,resultFile)
    #do_version_comparison(sourceDir,targetDir,versionFile)
    dcmp = dircmp(sourceDir, targetDir)
    #files_not_in_target(dcmp)
    #full_report(dcmp)
    find_app_version_in_release(dcmp)

#Find files which exists in new release but are not in Liberate destination folder 
def files_not_in_target(dcmp):
    for name in dcmp.left_only:
        soucePath = os.path.join(dcmp.left,name)
        if os.path.isdir(soucePath):
            print ('"%s" folder found in "%s" and not in "%s"\n\n' % (name, dcmp.left, dcmp.right))
        else:
            print ('"%s" file found in "%s" and not in "%s"\n\n' % (name, dcmp.left, dcmp.right))
            print_version(sourcePath,name)
    for sub_dir in dcmp.subdirs.values():
        files_not_in_target(sub_dir)
        
#Find out new version of EXEs in new release
def find_app_version_in_release(dcmp):
    for name in dcmp.left_list:
        sourcePath = os.path.join(dcmp.left,name)
        if os.path.isdir(sourcePath):
            pass
            #print
            #print ('"%s" folder in "%s"\n\n' % (name, dcmp.left))
        else:
            pass
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
        
 #Find files present in both and get their versions     
def find_app_version_in_release(dcmp):
    for name in dcmp.common:
        sourcePath = os.path.join(dcmp.left,name)
        if os.path.isdir(sourcePath):
            pass
            #print
            #print ('"%s" folder in "%s"\n\n' % (name, dcmp.left))
        else:
            pass
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
        
        
def print_version(sourcePath, fname):
    if sourcePath.lower().endswith('.exe'):
                #print "TRUE"
        major,minor,subminor,revision = get_version_info(sourcePath)
        print ("%s - %s.%s.%s.%s" % (fname, major,minor,subminor,revision))    
    return

def full_report(dcmp):
    print dcmp.report_full_closure()
    
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