import re
import csv
import pandas as pd
from igraph import *
import leidenalg as la # for leiden algorithm
#import infomap as la    for infomap
#import louvain as la    for louvain
import xlsxwriter
import networkx as nx
import itertools
from igraph.drawing import vertex
from matplotlib import pyplot as plt
from operator import itemgetter
import networkx as nx
import numpy as np
def intersection(lst1, lst2):
    lst3 = [value for value in lst1 if value in lst2]
    return lst3
row = 0
offo=0
column = 0
bas=0
max=0
mylist =[]
mylist1=[]
mylist2=[]
yourlist =[]
yourlist1 =[]
Edges=[]
rjj=0
uff=0
temp=0
temp1 =[]
position=0
jaclist=[]
jnrlist=[]
lnrlist=[]
for chalpa in range(6): #make time changes (1 num more)
    mylist.append([])
    mylist1.append([])
    mylist2.append([])
    jaclist.append([])
    jnrlist.append([])
    lnrlist.append([])
G = Graph(directed = True)
df=pd.read_csv('COLLMSGLINK1.csv')#Edges data
f1 = open("PALS.txt", "a")
f2 = open("SIM.txt", "a")
for k in range(6): #change in time (1 num more)
#      fil2 = open('Jaccard.csv', 'a+')
 #     jnr1 = open('JNR.csv', 'a+')
  #    lnr1 = open('LNR.csv', 'a+')
      newdf = df[(df.TIME == k)]
      nodekag =newdf[['SRC']].values.tolist()
      nodekag1 =newdf[['TGT']].values.tolist()
      nodekag2 = nodekag + nodekag1
      nodekag3=np.array(nodekag2)
      nodekag4=np.unique(nodekag3)
      nodekag5=nodekag4.tolist()
      print(G)
      G.add_vertices(nodekag5)
      kaggle = newdf[['SRC', 'TGT']].values.tolist()
      for line in kaggle:
          edge1=G.vs['name'].index(line[0])
          edge2 = G.vs['name'].index(line[1])
          Edges.append([str(edge1),str(edge2)])
          G.add_edges([(edge1,edge2)])
      total = len(set(newdf.index))
      print('Time'+str(k)+':')
      f1.write('Time'+str(k)+':')
      partition = la.find_partition(G, la.ModularityVertexPartition) # for leiden and louvain
      # partition = G.community_infomap()      for infomap
      modularity_dict = {}
      for i,c in enumerate(partition):
        for name in c:
            modularity_dict[name] = i
            G.vs['modularity'] = (modularity_dict)
      for i,c in enumerate(partition):
         print('Community '+str(i)+':', list(c))
         print(len(c))
         z=0
         for t in list(c):
            for u in list(c):
                 mytuple = [str(t), str(u)]
                 if mytuple in Edges:
                     z = z + 1
         numerator = z * 2
         rjj = rjj+z
         if (len(c) !=1):
             denominator = len(c) * (len(c)-1)
         if (len(c) == 1): len(c)
         density = numerator/denominator
         relative=z/len(Edges)
         print(density)
         print(relative)
         mylist[bas].append(density)
         mylist1[bas].append(relative)
         cluster = G.subgraph(c)
         Clustering = cluster.transitivity_undirected(mode='nan')
         mylist2[bas].append(Clustering)
         yourlist1.append(list(c))
         if k==0:
             jaclist[bas].append(0)
             jnrlist[bas].append(0)
             lnrlist[bas].append(0)
         if(k != 0):
          for simmi, items in enumerate(yourlist):
                temp = len(intersection(items, list(c)))
                jacden= len(set(items+list(c)))
                jac=temp/jacden
                jnr=(len(list(c))-temp)/len(list(c))
                lnr=(len(items)-temp)/len(items)
                jaclist[bas].append(jac)
                jnrlist[bas].append(jnr)
                lnrlist[bas].append(lnr)
          jaclist[bas][-1]=jac+100
          jnrlist[bas][-1] = jnr + 100
          lnrlist[bas][-1] = lnr + 100
         f1.write('Community'+str(i)+':\n'+str(len(c))+'           \n')
         for x in list(c):
           f1.write(str(x)+" \n")
           f2.write('Time' + str(k) + ',')
           f2.write(str(x)+", "+str(i)+"\n")
      f1.write('Communities count: '+str(i+1)+'\n')
      output = plot(partition)
      output.save(str(k) + 'graph.png')
      mylist.append([])
      mylist1.append([])
      mylist2.append([])
      bas = bas + 1
      yourlist=yourlist1.copy()
      yy=len(yourlist1)
#      with fil2:
 #         write0 = csv.writer(fil2)
  #        write0.writerows(jaclist)
   #   with jnr1:
    #      writejnr = csv.writer(jnr1)
     #     writejnr.writerows(jnrlist)
      #with lnr1:
       #   writelnr = csv.writer(lnr1)
        #  writelnr.writerows(lnrlist)
      #jaclist.clear()
      #jnrlist.clear()
      #lnrlist.clear()
      yourlist1.clear()
      uff = 0
      print("not in any community")
      print(len(Edges)-rjj)
      rjj=0
      Edges.clear()
      #print(rjj)
      G.clear()
file = open('Density.csv', 'a+')
with file:
          write= csv.writer(file)
          write.writerows(mylist)
file_relative = open('relative.csv', 'a+')
with file_relative:
    write1 = csv.writer(file_relative)
    write1.writerows(mylist1)
file_cluster = open('cluster.csv', 'a+')
with file_cluster:
    write2= csv.writer(file_cluster)
    write2.writerows(mylist2)
fil2 = open('Jaccard.csv', 'a+')
with fil2:
    write0 = csv.writer(fil2)
    write0.writerows(jaclist)
jnr1 = open('JNR.csv', 'a+')
with jnr1:
    writejnr = csv.writer(jnr1)
    writejnr.writerows(jnrlist)
lnr1 = open('LNR.csv', 'a+')
with lnr1:
    writelnr = csv.writer(lnr1)
    writelnr.writerows(lnrlist)
f1.close()
f2.close()
