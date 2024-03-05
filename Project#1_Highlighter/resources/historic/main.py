# -*- coding: utf-8 -*-
"""
Created on Mon Sep 25 08:47:18 2023

@author: belenosbol
"""
# Note boundary case of multi line "AAA \n & BBB" Cannot be found

#Import Libraries
import win32com.client
import pandas as pd
import os

#Define Autocad Application to use
acad = win32com.client.Dispatch("AutoCAD.Application")

doc = acad.ActiveDocument   # Document object


#Open output file to store tree in .txt
#with open(fname, "w", encoding="utf-8") as f:
f= open("output/iFound.txt","w+", encoding="utf-8")

#Import Excel Column with data to look for
df = pd.read_excel('input/equipmentData.xlsx', sheet_name=0) # can also index sheet by name or fetch all sheets
mylist = df['Equipment'].tolist()

wasFound = []

#Colors:
#White=7
#Red=1
#Yellow=2
#Green=3
#Blue=5
#Magenta=6
#Black=250

myColor=6

# iterate trough all objects (entities) in the currently opened drawing
# and if its a BlockReference, display its attributes and some other things.
for entity in acad.ActiveDocument.ModelSpace:
    name = entity.EntityName
    #Check if its block reference
    if name == 'AcDbBlockReference':
        HasAttributes = entity.HasAttributes
        if HasAttributes:
            #Display general properties, not necessary
            #print(entity.Name)
            #print(entity.Layer)
            #print(entity.ObjectID)
            #Iterate through all attributes
            for attrib in entity.GetAttributes():
                #print("  {}: {}".format(attrib.TagString, attrib.TextString))
                #Iterate through values provided by Excel file checking for value match
                #If found display which value was found and change color
                for item in mylist:
                    if attrib.TextString == item:
                        wasFound.append(item)
                        attrib.color = myColor
                # update text
                #attrib.TextString = '101'
                #attrib.Update()
    #Check if its regular text
    elif name == 'AcDbMText':
        for item in mylist:
            if entity.TextString == item:
                wasFound.append(item)
                entity.color = myColor
            #Filter Underscores
            elif entity.TextString == "%%U"+item:
                wasFound.append(item)
                entity.color = myColor
            #Filter Underscores
            elif entity.TextString == "%%u"+item:
                wasFound.append(item)
                entity.color = myColor
    elif name == 'AcDbText':
        for item in mylist:
            if entity.TextString == item:
                wasFound.append(item)
                entity.color = myColor
            #Filter Underscores
            elif entity.TextString == "%%U"+item:
                wasFound.append(item)
                entity.color = myColor
            #Filter Underscores
            elif entity.TextString == "%%u"+item:
                wasFound.append(item)
                entity.color = myColor

notFound = list(set(mylist).difference(wasFound))
print("Was Found\n") 
f.write("Was Found\n")  
for entity in wasFound:
    print(entity)
    f.write("%s\n" % (entity))          

print("\n\nWas Not Found\n")   
f.write("\nWas Not Found\n")  
for entity in notFound:
    print(entity)
    f.write("%s\n" % (entity))  

#Closing output file
f.close()

# Save the documet
doc.SaveAs(os.path.join(os.path.abspath(os.path.dirname(__file__)), "output", "iFound.dwg"))

### Adjust dwg ###
doc.Save()


        

