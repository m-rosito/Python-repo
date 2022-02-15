###############################################################################



############################### INIT USR ######################################

#path = 'Y:/Mijn documenten/@FM/Python/Scripts/data' # locatie xlsx-files
#path = r'C:\Users\Michele\Documents\Python Scripts\Data'
path = r'C:\Users\PMMROONT\Documents\Python Scripts\Data'

#geef kolomnamen op van velden die NIET moeten worden meegenomen in matching:
cols_del='DWH_'

############################### INIT DEV ######################################

f_names   = [] # lijst namen xlsx-files in 'path' + verkorte werknamen
f_list    = [] # lijst verkorte werknamen tbv eenvoudige for-loops
f_data    = {} # dict van dataframes met xlsx-files
f_missing = {} # lijst met aantal records per file met missing values
f_matches = {} # matrix met resulltaat compare van elke file-combinatie

import os
import numpy as np
import pandas as pd # pd.__version__ = '0.20.3'
import matplotlib.pyplot as plt

############################## READ FILES #####################################

#inlezen f_names:
i=0
for file in os.listdir(path):
    if file.endswith(".xlsx") and file[0:2]!='~$': #exclude temp-files
        f_names.append(['File_'+str(i+1),file])
        f_list.append(  'File_'+str(i+1))
        i+=1

#inlezen f_data:
i=0
for file in f_names:
    f_data[f_names[i][0]]=pd.read_excel(os.path.join(path,f_names[i][1]))
    i+=1
    
#en verwijder kolommen die evt. niet meegenomen moeten worden in matching:
for file in f_list:
    cols = f_data[file].columns.tolist()
    cols = [i for i in cols if cols_del not in i]
    f_data[file] = f_data[file].loc[:, cols]

##################### COUNT RECORDS w. MISSING DATA ###########################

#bepaal f_missing:
f_missing=pd.DataFrame(index=f_list,columns=['recs_with_missing'])
for file in f_list:
    f_missing.loc[file] = f_data[file].isnull().any(axis=1).sum()

############################ COMPARE FILES ####################################

#init lege f_matches met f_names in rij- en kolomnamen :
f_matches=pd.DataFrame(index=f_list, columns=f_list)

#vervang alle missing values in f_data met '?' tbv een correcte match:
for file in f_list:
    _=f_data[file].fillna('?',inplace=True)

#bepaal match-resultaat in matrix voor elke combinatie van Excel-bestanden:
for row in f_list:
    for col in f_list:
        f_matches.loc[row,col] = \
        (f_data[row]!=f_data[col]).any().any()

#vervang booleans met [0,1], met 1=gelijk, 0=ongelijk, tbv opmaak tabel-output
f_matches=f_matches.astype(int).replace([0,1],[1,0])

############################ PRINT RESULTS ###################################

#print tabel met resultaat van match & legenda:
fig,axs=plt.subplots(2,2,figsize=(12,7))
fig.suptitle('Summary Automated Checks Risk Files', fontsize=20, y=0.98)

# verwiijder alle x- en y- ticks:
r=0; c=0
for r in range(axs.shape[0]):
    for c in range(axs.shape[1]):
        axs[r,c].set_xticks([],[])
        axs[r,c].set_yticks([],[])
        c+=1
    r+=1

#print tabel met legenda voor gebruikte short names:
ax=axs[0,0]
ax.set_title('Short → Orig. Names Compare Files', fontsize=14)
df      = pd.DataFrame(f_names, columns=['Short Name','Orig. Name'])
rlabels = df.index
clabels = df.columns
vals    = df.values
table2=ax.table(cellText  = df.values, 
                cellLoc   ='left',
                colLabels = clabels,
                loc       = 'left',
                bbox      =  [0.1, 0.2, 0.8, 0.6])
table2.auto_set_column_width(col=list(range(len(df.columns))))

#print matrix met resultaat compare
ax=axs[0,1]
ax.set_title('Match Result Compare Files', fontsize=14)
ax.text(0.4,0.05,' 1=Match, 0=No-match')
df      = f_matches
labels  = f_list
vals    = df.values
normal  = plt.Normalize(vals.min()-0.5, vals.max()+0.5)
colours = plt.cm.RdYlGn(normal(vals))
table1=ax.table(cellText = vals,
                cellLoc     ='center',
                rowLabels   = labels, 
                colLabels   = labels, 
                loc         = 'center',
                cellColours = colours,
                bbox        =  [0.25, 0.15, 0.6, 0.7])
table1.auto_set_column_width(col=list(range(len(df.columns))))

#print tabel met aantal records met missing values:
ax=axs[1,0]
ax.set_title('Aantal records met missing values', fontsize=14)
df      = f_missing
rlabels = df.index
clabels = df.columns
vals    = df.values
table3=ax.table(cellText  = df.values, 
                cellLoc   ='center',
                rowLabels = rlabels, 
                colLabels = clabels,
                loc       = 'left',
                bbox      =  [0.1, 0.2, 0.8, 0.6])
#table3.auto_set_column_width(col=list(range(len(df.columns))))

#instellen fontsize voor alle tables in output:
[t.set_fontsize(10) for t in [table1, table2, table3]]

#opslaan table prints in jpg-format in dir input files:
fig.savefig(os.path.join(path,'f_matrix.jpg'), 
            transparent=False, 
            dpi=96, 
            bbox_inches="tight")

############################### THE END #######################################
