
import os
import fnmatch
import shutil

def copy_files():
    pattern = 'outlook*.txt'
    for root,dirs,files in os.walk('c:\\release'):
         #print root
        for f in files:                
                
                #print (os.listdir())
                #dirlist.append(os.listdir(root))
                if len(fnmatch.filter(files,pattern)) > 0:
                    testpath = (os.path.join(root,f).split('\\'))[:-1]
                    releasepath = (root)
                    #tdir = os.path.dirname(os.path.join(root,f))
                    #print root
                    
                    #print "this is dir %s" % tdir
                    #relitem = testpath[-2]
                    #relfullpath = (root)
                    #print f
                    break
                    #reloutlookdir = (test[-2])
    #print testpath
    #print reldir                #prinst reloutlookdir
    #print root
    for root,dirs,files in os.walk('c:\\LiberateDev'):
         #print root
        for f in files:                
                
                #print (os.listdir())
                #dirlist.append(os.listdir(root))
                if len(fnmatch.filter(files,pattern)) > 0:
                    devpath =  (os.path.join(root,f).split('\\'))[:-1]
                    devdir = devpath[-1]
                    devfullpath = (root)
                    break
                    #reloutlookdir = (test[-2])
        
    #print devdir                #prinst reloutlookdir
    print devfullpath
    print releasepath
    testpath[-1] = devdir
    #testpath[-2] = reldir
    print devpath
    newpath = testpath[0] + os.sep + os.path.relpath(os.path.join(*testpath))
    print newpath
    if not os.path.isdir(newpath):
        shutil.copytree(releasepath,newpath)
        shutil.rmtree(releasepath)
    else:
        print "directory already exists"
if __name__ == "__main__":
    copy_files()