# -*- coding: utf-8 -*-
import os, sys, glob, ntpath
from PIL import Image
from shutil import copyfile
import logging
import os.path
#from pprint import pprint

logging.basicConfig(filename='assetupload.log',format='%(asctime)s %(message)s', level=logging.INFO)
logger=logging.getLogger(__name__)

def main():    
    size = 64, 45
    cwd = os.getcwd()
    filelist = glob.glob(os.path.join(cwd, '*.jpg'))
    out_list=[]
    for infile in sorted(filelist):
        #print('infile=%s' % infile)
        filename = ntpath.basename(os.path.splitext(infile)[0])
        assetname = filename.split('_')[0]  #handle one asset with multi pic case, filename list 560123_1.jpg, 560123_2.jpg
        #print('assetname=%s' % assetname)
        if not os.path.exists(assetname):
            os.mkdir(assetname)
        outfile = assetname + ".thumb.jpg"        
        try:
            im = Image.open(infile)
            im.thumbnail(size)            
            path = os.path.join(cwd, assetname,outfile)
            im.save(path, "JPEG")
            dst = os.path.join(cwd, assetname,filename+'.jpg')
            #print('dst=%s' % dst)
            copyfile(infile,dst)
            out_list.append([assetname,dst])
            logger.info('%s,%s' %(assetname,dst))
        except IOError:
            print("cannot create thumbnail for", infile)
    
    #pprint(out_list)
            
if __name__ == "__main__":
    main()              
