import win32com.client
import pandas as pd
import os
import time

# Note boundary case of multi line "AAA \n & BBB" Cannot be found by function

#Creates list of intersection between two lists
def intersection(lst1, lst2):
    lst3 = [value for value in lst1 if value in lst2]
    return lst3

#Essential function
def findInACAD(myFile, neverFound):
    #Define Autocad Application to use
    acad = win32com.client.Dispatch("AutoCAD.Application")
    #acad.Visible = True
    doc = acad.Documents.Open(os.path.join(os.path.abspath(os.path.dirname(__file__)), "..", "input", myFile))
    #Extremely necesary because pull request are quicker than AUTOCAD boot
    #Troubleshooting Achiles Heel
    time.sleep(.2)          #Give time to fetch and open files
    
    #Open output file to store tree in .txt
    #with open(fname, "w", encoding="utf-8") as f:
    f= open("output/iFound.txt","a", encoding="utf-8")
    
    #Import Excel Column with data to look for
    df = pd.read_excel('input/equipmentData.xlsx', sheet_name=0) # can also index sheet by name or fetch all sheets
    mylist = df['Equipment'].tolist()
    
    #List for Equipment tracking
    wasFound = []
    
    #Colors:
    #Red=1, Yellow=2, Green=3, Blue=5
    #Magenta=6, #White=7, Black=250
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
        else:
            continue
    #I have been repeatedly told
    #that my work and these tools are
    #worthless and a waste of time by
    #our documentation lead.
    #Never let clueless people 
    #influence the quality of your passion
    #Keep going, you are doing great
    #Finds disjunction between two lists
    notFound = list(set(mylist).difference(wasFound))
    
    #Initialize NeverFound
    if neverFound == []:
        neverFound = notFound
    
    #Store and print data
    if wasFound != []:
        print("\n---In " + myFile + "---\n" + "*Was Found*") 
        f.write("\n---In " + myFile + "---\n" + "*Was Found*\n")  
        for entity in wasFound:
            print(entity)
            f.write("%s\n" % (entity))          
        
        #Uncomment in case output of not found per file is desired
        #print("*Was Not Found*")   
        #f.write("*Was Not Found*\n")  
        #for entity in notFound:
            #print(entity)
            #f.write("%s\n" % (entity))  
    
        # Save the updated .dwg document to prevent modifying original file
        doc.SaveAs(os.path.join(os.path.abspath(os.path.dirname(__file__)), "..", "output", "mod_" + myFile))
        
        ### Adjust .dwg ###
        doc.Save()
    #Print headsup for nonUseful .dwg file
    else:
        print("\n---In " + myFile + "---\n" + "-Nothing Was Found-") 
        f.write("\n---In " + myFile + "---\n" + "-Nothing Was Found-\n")  
    
    #Closing iFound.txt output file
    f.close()
    
    #close modified AutoCAD file (original was changed to modified when Saved As)
    #No need to close original
    doc.Close()
    
    #return updated neverFound list
    return intersection(neverFound, notFound)