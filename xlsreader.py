# import openpyxl
# wb = openpyxl.load_workbook('Network.xlsx')
# my_sheetnames = wb.sheetnames
# sheet = my_sheetnames[1]
# worksheet = wb[sheet]
# print(sheet)
# print(my_sheetnames)
# for value in worksheet.iter_columns(values_only=True):
#    print(value)
class Network():

    def __init__(self, network_graph=None):
        if network_graph == None:
            network_graph = {}
        self._network_graph = network_graph

    def AllNetworkNodes(self):
        return list(set(self._network_graph.keys()))

    def AllDeviceNodes(self):
        Device = []
        for Node in self._network_graph.keys():
            for Neighbour in self._network_graph[Node]:
                if Neighbour not in Device:
                    Device.append(Neighbour)
        return (Device)

    def AllLinks(self):
        link = []
        for Node in self._network_graph.keys():
            for Neighbour in self._network_graph[Node]:
                if {Node, Neighbour} not in link:
                    link.append({Node, Neighbour})
        # print(link)
        for i in range(len(link)):
            link[i] = tuple(link[i])
        return (link)

    def NodePresense(self, Node):
        if Node in self._network_graph:
            return "Node present"
        else:
            return "Node NOT present"

    def NodeLinks(self, Node):
        if Node in self._network_graph:
            return self._network_graph[Node]
        else:
            return "Node NOT present"

    def find_path(self, StartNode, EndNode, Path=None):
        if Path == None:
            Path = []
        Path.append(StartNode)
        if EndNode in self._network_graph[StartNode]:
            Path.append(EndNode)
            return Path
        if StartNode == EndNode:
            return Path
        if StartNode in self._network_graph.keys():
            for nextNode in self._network_graph[StartNode]:
                if nextNode not in Path:
                    extended_path = self.find_path(nextNode, EndNode, Path)
                    if extended_path:
                        return extended_path
            return None

    def find_all_paths(self, start_vertex, end_vertex, path=[]):
        """ find all paths from start_vertex to
            end_vertex in graph """
        graph = self._network_graph
        path = path + [start_vertex]
        if start_vertex == end_vertex:
            return [path]
        if start_vertex not in graph:
            return []
        paths = []
        for vertex in graph[start_vertex]:
            if vertex not in path:
                extended_paths = self.find_all_paths(vertex,
                                                     end_vertex,
                                                     path)
                for p in extended_paths:
                    paths.append(p)
        return paths

    def __iter__(self):
        self._iterObject = iter(self._network_graph)
        return self._iterObject

    def __next__(self):
        return next(self._iterObject)


def sheetAddition(AllNetworkNodes,AllDeviceNodes,AllLinks):
    NodeDict = {}
    LinkDict = {}
    AxisDict = {"XAxis": "", "YAxis": ""}
    AttDict = {"Attribute": ""}

    #Sheet data creation for Node Position
    for i in AllNetworkNodes:
        NodeDict[i] = AxisDict
    for i in AllDeviceNodes:
        NodeDict[i] = AxisDict
    dfPosition = pd.DataFrame(NodeDict)

    #Sheet data creation fro link
    for i in AllLinks:
        LinkDict[str(i)] = AttDict
    dfLinkAtt = pd.DataFrame(LinkDict)

    #Writing to Excel
    with pd.ExcelWriter('Network.xls') as writer:
        df.to_excel(writer, sheet_name='data', index=False)
        dfPosition.to_excel(writer, sheet_name='Position')
        dfLinkAtt.to_excel(writer, sheet_name='LinkAtt')

def remove_null(dlist):
    flist = []
    for i in dlist:
      j = i.replace(' ','')
      j = j.lower()
      if j != 'na':
        flist.append(j)
    return flist

import networkx as nx
import matplotlib.pyplot as plt
import pandas as pd


# def remove_null(dlist):
#     flist = []
#     for i in dlist:
#         if i != 'NA':
#             flist.append(i)
#     return flist
#
#
# df = pd.read_excel('Network.xls', sheet_name='data')
# df.fillna('NA', inplace=True)
# df_dict = df.to_dict('list')
# df_dict.pop("Network")
# for i in df_dict:
#     df_dict[i] = remove_null(df_dict[i])


df = pd.read_excel('Network.xls', sheet_name='data')
dfPos = pd.read_excel('Network.xls', sheet_name='Position')
dfDtype = pd.read_excel('Network.xls', sheet_name='Device_Type')

#fill all blacks with NA
df.fillna('NA', inplace=True)
dfPos.fillna(50, inplace=True)
dfDtype.fillna('NA', inplace=True)

#Convert excel to list dictionary (vlaues in lict) , column head as key
df_dict = df.to_dict('list')
# dfPos_dict = dfPos.to_dict('list')
# dfPos_dict = {x:tuple(y) for x, y in dfPos_dict.items()}
# # print(dfPos_dict)

dfPos_dict = dfPos.to_dict('list')
pos = {}
for po in range (len(dfPos_dict['Device'])):
    x, y = 10, 10
    pos[dfPos_dict['Device'][po]] = ((x*dfPos_dict['Position from East'][po],y*dfPos_dict['Position from South'][po]))
    # x += 5
    # y += 5
print(pos)

#Remove Network Column
# dfPos_dict.pop('Unnamed: 0')
df_dict.pop("Network")
dfDtype_dict = dfDtype.to_dict('list')
# print(dfDtype_dict)

#Relace blankc spaces from key and Convert keys to lower case
df_dict = {x.replace(' ', '') : y for x, y in df_dict.items()}
df_dict = {x.lower() : y for x, y in df_dict.items()}

#Remove balanks and convert all to lower case from values
for i in df_dict:
    df_dict[i] = remove_null(df_dict[i])

#Print dictioanry which is input to graph
# print(df_dict)

#Convert G into "Network" class object
G = Network(df_dict)
#Get Nodes and Edges from dictionary and add colour and device type attribute
AllNetworkNodes = G.AllNetworkNodes()
for i in range(len(AllNetworkNodes)):
    AllNetworkNodes[i] = (AllNetworkNodes[i], {"Type":"Network"})
AllDeviceNodes = G.AllDeviceNodes()
for i in range(len(AllDeviceNodes)):
     AllDeviceNodes[i] = (AllDeviceNodes[i], {"Type":"Device"})
AllLinks = G.AllLinks()

# sheetAddition(AllNetworkNodes,AllDeviceNodes,AllLinks)

# create Gnx blank graph and ad nodes and edges to Gragh
Gnx = nx.Graph()
Gnx.add_nodes_from(AllNetworkNodes)
Gnx.add_nodes_from(AllDeviceNodes)
Gnx.add_edges_from(AllLinks)


# # Print data like edges,nodes,degree,edge data,node data,adjacency data,neighbour data
# print ("\n Gnx.edges".ljust(15)+ ":", Gnx.edges())
# print ("\n Gnx.nodes".ljust(15)+ ":",Gnx.nodes())
# print ("\n Gnx.degree".ljust(15)+ ":",Gnx.degree)
# print ("\n Gnx.edges.data".ljust(15)+ ":",Gnx.edges.data())
# print ("\n Gnx.nodes.data".ljust(15)+ ":",Gnx.nodes.data())
# print ("\n Gnx.adjacency".ljust(15)+ ":",Gnx.adj)
# for nbr in Gnx.neighbors('10.0.0.0/24'):
#     print ("\n Gnx.neighbors".ljust(15)+ ":",nbr)
# # nx.draw_networkx_nodes()
# total_pos = len(AllNetworkNodes) + len(AllDeviceNodes) + len(AllLinks)
# pos = {}
# for po in range (total_pos):
#     x, y = 10, 10
#     pos[po] = (x,y)
#     x += 5
#     y += 5
# print (pos)
# print (Gnx.nodes['20.0.0.0/24']["Type"])
node_sizes = []
node_sizes = [1000 if Gnx.nodes[i]["Type"] == 'Device' else 100 for i in Gnx.nodes]
# print(node_sizes)
node_Colours = []
for i in Gnx.nodes:
    if i in dfDtype_dict["Firewall"]:
      node_Colours.append("r")
    elif i in dfDtype_dict["Switch"]:
      node_Colours.append("c")
    elif i in dfDtype_dict["Router"]:
       node_Colours.append("m")
    elif i in dfDtype_dict["Loadbalancer"]:
       node_Colours.append("g")
    else:
      node_Colours.append("b")
# print(node_Colours)
path = G.find_all_paths("a", "h")
print(path)

#Draw graph
nx.draw(Gnx,with_labels=True, pos=pos,node_shape='s',node_size=node_sizes, font_size=4, node_color=node_Colours)
#Display Graph
plt.show()