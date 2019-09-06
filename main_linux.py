import os
import shutil
import sys
import Rabbit
from zipfile import ZipFile
 
def extract():
    with ZipFile(imp_file, 'r') as zipObj:
        listOfFileNames = zipObj.namelist()
        for fileName in listOfFileNames:
            zipObj.extract(fileName, tmp_dir)
       
def insert():
    os.chdir(tmp_dir)
    dirName = './'
    # create a ZipFile object
    with ZipFile(doc_name, 'w') as zipObj:
    # Iterate over all the files in directory
        for folderName, subfolders, filenames in os.walk(dirName):
            for filename in filenames:
                if not doc_name in filename:
                     
    #create complete filepath of file in directory
                    filePath = os.path.join(folderName, filename)
    # Add file to zip
                    zipObj.write(filePath)
                     
    os.chdir(sys.path[0])
    source = tmp_dir
    dest = out_dir
    files = os.listdir(source)

    for f in files:
        if '.docx' in f:
            if os.path.exists(dest+f):
                os.remove(dest+f)
                shutil.move(source+f, dest)
            else:
                shutil.move(source+f, dest)
            
        elif '.xlsx' in f:
            if os.path.exists(dest+f):
                os.remove(dest+f)
                shutil.move(source+f, dest)
            else:
                shutil.move(source+f, dest)
        
        elif '.pptx' in f:
            if os.path.exists(dest+f):
                os.remove(dest+f)
                shutil.move(source+f, dest)
            else:
                shutil.move(source+f, dest)
        
    shutil.rmtree(source)
        
       
def convert():
    files = []
    for r, d, f in os.walk(tmp_dir):
        for doc in f:
            if '.xml' in doc:
                doc_name = doc
                f_locate = str(r+'/'+doc)
                #file read
                r_file = open(f_locate, "r", encoding='utf-8')
                text = (r_file.read())
                #file convert
                # unicode_text = Rabbit.uni2zg(text)
                zawgyi_text = Rabbit.zg2uni(text)
                #file write
                w_file = open(f_locate, 'w+', encoding='utf-8')
                w_file.write (zawgyi_text)
                r_file.close()


in_dir = "Zawgyi_Files/"
out_dir = "Converted_Files/"
tmp_dir = "_tmp/"
files = []
# r=root, d=directories, f = files
for r, d, f in os.walk(in_dir):
    for file in f:
        if '.docx' in file:
            doc_name = file
            imp_file = (r+'/'+file)
            print ('>>> Converting the',doc_name)
            extract()
            convert()
            insert()
            
        elif '.xlsx' in file:
            doc_name = file
            imp_file = (r+'/'+file)
            print ('>>> Converting the',doc_name)
            extract()
            convert()
            insert()
            
        elif '.pptx' in file:
            doc_name = file
            imp_file = (r+'/'+file)
            print ('>>> Converting the',doc_name)
            extract()
            convert()
            insert()
 
print (' All files converted!!!','\n----------------------')