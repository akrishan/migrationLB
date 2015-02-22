import os
import fnmatch
import shutil

      
def main():    
    outlook_pattern = 'outlook*'
    excel_pattern = 'excel*'
    release_full_path = r'c:\release'
    dev_full_path = r'c:\LiberateDev'
    releasepath,releaseroot,releasefoldername = get_paths_based_on_pattern(release_full_path,outlook_pattern)
    devpath, devroot,devdir = get_paths_based_on_pattern(dev_full_path,outlook_pattern)
    match_and_rename_release(releasepath,releaseroot,devdir)

def get_paths_based_on_pattern(abspathname,pattern):
    for root,dirs,files in os.walk(abspathname,topdown=False):
            for f in files:
                for filename in fnmatch.filter(files,pattern):
                    if len(fnmatch.filter(files,pattern)) > 0:                        
                        splitpath = (os.path.join(root,filename).split('\\'))[:-1]
                        print splitpath
                        rootpath = (root)
                        foldername = splitpath[-1]
                        break
    return splitpath, rootpath, foldername
        
def match_and_rename_release(releasepath,releaseroot,sourcedir):
    releasepath[-1] = sourcedir #replace directory name from development in release folder
    newpath = releasepath[0] + os.sep + os.path.relpath(os.path.join(*releasepath))
    try:
        if not os.path.isdir(newpath):
            shutil.copytree(releaseroot,newpath)
            shutil.rmtree(releaseroot)
        elif os.path.isdir(newpath):
            print ("Directory already exists,copying correct addin files in this %s directory from %s" % (newpath,releaseroot))
            copytree(releaseroot,newpath)
    except shutil.Error, exc:
        print exc.args[0]
                    
        
#helper function
def copytree(src, dst, symlinks=False, ignore=None):
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)
        
if __name__ == "__main__":
    main()