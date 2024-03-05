#// ===============================
#// AUTHOR     :Bolivar Beleno
#// CREATE DATE     :9/25/2023
#// PURPOSE     :Plot Hierarchy tree for Equipment
#// SPECIAL NOTES:
#// ===============================
#// Change History:
#//
#//==================================
###


#Load libraries: Pandas and anytree
import pandas as pd
from anytree import Node, RenderTree

#Open output file to store tree in .txt
#with open(fname, "w", encoding="utf-8") as f:
f= open("output/myTree.txt","w+", encoding="utf-8")

# Read data with subset of columns 
df = pd.read_excel("input/equipmentData.xlsx", usecols = ['Equipment','Source'])

#Check
#print(df)

#Drop NaN rows
df2=df.dropna()
df2=df.dropna(axis=0)

#Check
#print(df2)

def add_nodes(nodes, parent, child):
    if parent not in nodes:
        nodes[parent] = Node(parent)  
    if child not in nodes:
        nodes[child] = Node(child)
    nodes[child].parent = nodes[parent]

data = pd.DataFrame(columns=["Source","Equipment"], data=df2)
nodes = {}  # store references to created nodes 
# data.apply(lambda x: add_nodes(nodes, x["Item"], x["Source"]), axis=1)  # 1-liner
for parent, child in zip(data["Source"],data["Equipment"]):
    add_nodes(nodes, parent, child)

roots = list(data[~data["Source"].isin(data["Equipment"])]["Source"].unique())
for root in roots:         # you can skip this for roots[0], if there is no forest and just 1 tree
    for pre, _, node in RenderTree(nodes[root]):
        print("%s%s" % (pre, node.name))
        f.write("%s%s\n" % (pre, node.name))          #Printing into output myTree.txt file


#RenderTreeGraph(root).to_picture("tree.png")

#To update for a specific Root
#root = 'SB-1' # change according to usecase
#for pre, _, node in RenderTree(nodes[root]):
   #print("%s%s" % (pre, node.name))

#Closing output file
f.close()

   
