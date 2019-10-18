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
    files = sorted(filelist)
    asset_count = len(files)    
    for idx, infile in enumerate(files):
        #print('infile=%s' % infile)
        filename = ntpath.basename(os.path.splitext(infile)[0])        
        if len(filename)<12:
            continue
        assetname = filename[:12]  #handle one asset with multi pic case, filename list 560123_1.jpg, 560123_2.jpg
        next_assetname = ntpath.basename(os.path.splitext(files[idx+1])[0])[:12] if idx < asset_count - 1 else ''
        prev_assetname = ntpath.basename(os.path.splitext(files[idx-1])[0])[:12] if idx >0 else ''
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
            #combine same assets' pic using ; concatenation
            output_path = dst if assetname != prev_assetname else ';'.join([output_path,filename+'.jpg'])
            if assetname != next_assetname:
                out_list.append([assetname,output_path])
                logger.info('%s,%s' %(assetname,output_path))
        except IOError:
            print("cannot create thumbnail for", infile)
    
    #pprint(out_list)
            
if __name__ == "__main__":
    main()   
