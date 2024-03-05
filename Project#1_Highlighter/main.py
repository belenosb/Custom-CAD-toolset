from resources.findAttributeACAD import findInACAD
import shutil
import os
import glob


#Delete output files to ensure clean start
dir_path = 'output'
try:
    shutil.rmtree(dir_path)
except OSError as e:
    print("Error: %s : %s" % (dir_path, e.strerror))
# Create the directory 
os.mkdir(os.path.join(os.path.abspath(os.path.dirname(__file__)), "output"))

#Main Routine
#Keeps track of items not found. Is printed at the EOF
neverFound = []      
#Iterates script for all .dwg files in input directory                                   
for filepath in glob.iglob('input/*.dwg'):
    neverFound = findInACAD(os.path.basename(os.path.normpath(filepath)), neverFound)

#Open output file to present if all items were found
f= open("output/iFound.txt","a", encoding="utf-8")
if neverFound == []:
    print("\n-----All items were Found-----\n")
    f.write("\n-----All items were Found-----\n")
else:
    print("\n---Never Found---")
    f.write("\n---Never Found---")
    for entity in neverFound:
        print(entity + " Was Never Found")
        f.write("\n%s Was Never Found\n" % (entity))

#Closing recently opened output file
f.close()
    
#Remove .bak clutter files
dir = os.path.join(os.path.abspath(os.path.dirname(__file__)), "output")
for zippath in glob.iglob(os.path.join(dir, '*.bak')):
    os.remove(zippath)