# -*- coding: utf-8 -*-
"""
Created on Mon Mar 18 18:05:04 2019

@author: Dipankar
"""

import shutil
import os
from difflib import SequenceMatcher
import xlsxwriter
from openpyxl import load_workbook
import xlrd

def getdirname(d1,d2,sheetnm):
    dlist1=list()
    dlist2=list()
    dlistsr1=list()
    dlistsr2=list()
    totcntlist=list()
    changelist=list()
    mninetylist=list()
    addlist=list()
    matchlist=list()
    chngfllist=list()
    addfllist=list()
    chngNfllist=list()
    samelist = list()
    print(d1)
    print(d2)
    filevel=0
    
    for (dirpath, dirnames, filenames) in os.walk(d1):
        for d in dirnames:
            dlist1.append(os.path.join(dirpath,d))
            dlistsr1.append(d)
            filevel = filevel + 1
            
        if filevel > 0:
            break

    filevel =0
    for (dirpath, dirnames, filenames) in os.walk(d2):
        for d in dirnames:
            dlist2.append(os.path.join(dirpath,d))
            dlistsr2.append(d)
            filevel = filevel +1 
            
        if filevel > 0:
            break
    

    
    skipcnt=0
    cnt=0
    flag=False
    for i in range(0,len(dlist1)):
        emptyTuple = ()
        skipcnt=0
       # totflcnt=0
        #totflcntlist=list()

        pcflenth = len('C:\GW\workspace\v10_conversion\ClaimCenter\modules\configuration\config\web\\')
        valupath=dlist1[i]
        abspath=valupath[pcflenth:]
        
        print(dlist1[i])
        totflcntlist= ()
        totflcntlistfl= list()
        addflist=list()
 #       totflcntlistdir=list()
        addtp=()
        for j in range(0,len(dlist2)):
            
            if cnt != skipcnt:
                skipcnt = skipcnt+1
                flag = True
                continue
 
            if dlistsr1[i] == dlistsr2[j]:
                
                emptyTuple = getChangeCount(dlist1[i],dlist2[j])
                cnt= cnt + 1
                totcntlist.append(emptyTuple[0])
                matchlist.append(emptyTuple[1])
                changelist.append(emptyTuple[2])
                mninetylist.append(emptyTuple[3])
                addlist.append(emptyTuple[4])
                if len(emptyTuple[5]) > 0:
                    chngNfllist.append(emptyTuple[5])
                if len(emptyTuple[6]) > 0:
                    chngfllist.append(emptyTuple[6])
                if len(emptyTuple[7]) > 0:
                    addfllist.append(emptyTuple[7])
                if len(emptyTuple[8]) > 0:
                    samelist.append(emptyTuple[8])
                flag = False
                
                break
            else:
                totflcntlist= getFileCnt(dlist1[i])
                dr3=getImmidiatedir(dlist1[i])
                
                for nm in totflcntlist:
                    addtp=()
                    pathfll= os.path.join(dr3, nm) 
                    addtp=(abspath,pathfll)
                    addflist.append(addtp)
                         
               # changelist.append(totflcnt)
                addlist.append(len(totflcntlist))
                if len(addflist) > 0:
                    addfllist.append(addflist)
                             
                
#                if len (totflcntlist) > 0:
#                    addfllist.append(totflcntlist)
                flag = False
                break
        
        if flag:
            totflcntlist= getFileCnt(dlist1[i])
            
            
            #changelist.append(totflcnt)
            addlist.append(len(totflcntlist))
            dr3=getImmidiatedir(dlist1[i])  
            for nm in totflcntlist:
                addtp=()
                pathfll= os.path.join(dr3, nm) 
                addtp = (abspath,pathfll)
                addflist.append(addtp)
            
            if len(addflist) > 0:
                addfllist.append(addflist)
#            if len(totflcntlist) > 0 :
#                addfllist.append(totflcntlist)
            
     
    
    totpcffilecnt=0
    changepcffilecnt=0  
    cngnity=0    
    addfile=0 
    matchfl=0    
    for  i in range(0,len(totcntlist)):
        totpcffilecnt = totpcffilecnt + totcntlist[i]
        
        
    for  b in range(0,len(matchlist)):
        matchfl = matchfl + matchlist[b]
        
    
    
    for  i in range(0,len(changelist)):
        changepcffilecnt = changepcffilecnt + changelist[i]
        
    
    for k in range (0,len(mninetylist)):
        cngnity = cngnity + mninetylist[k]
        
    for j in range (0,len(addlist)):
        addfile = addfile + addlist[j]
        

    print('totfile ',totpcffilecnt)
    print('change ',changepcffilecnt)
    print('nieetyper ',cngnity)
    print('addfile ',addfile)
    print('matchfile ',matchfl)
    
    percetange=  (changepcffilecnt + addfile) * 100 / totpcffilecnt
    exceltp=(addfllist,samelist,chngNfllist,chngfllist,totpcffilecnt,matchfl,changepcffilecnt,cngnity,addfile,percetange)
   
    
    writExcelFile(exceltp,sheetnm)
    return percetange       

def writExcelFile(exceltp,sheetnm):
    addfllist=exceltp[0]
    samelist=exceltp[1]
    chngNfllist = exceltp[2]
    chngfllist = exceltp[3]
    exists = os.path.isfile('C:\\TestTmp\\StatFile.xlsx')
    #flag=False
    if not exists:
        workbook  = xlsxwriter.Workbook('C:\\TestTmp\\StatFile.xlsx')
        bold = workbook.add_format({'bold': 1})
        
    worksheet1 = workbook.add_worksheet(sheetnm)
    row = 0
    col = 0
    if len(addfllist) > 0:
        worksheet1.write('A1', 'DirectoryNameOfAddFile', bold)
        worksheet1.write('B1', 'FileNameWithimmidateDir', bold)
        for adfl in addfllist:
            for fl in adfl:
                row +=1
                worksheet1.write_string(row,col,fl[0])
                worksheet1.write_string(row,col+1,fl[1])
            
            
            
    if row > 0:
        row=0
        col = 3
        
        
    if (len(chngNfllist) > 0 or len(chngfllist) > 0 or len(samelist) > 0):
        worksheet1.write('D1', 'DirectoryPath', bold)
        worksheet1.write('E1', 'FilenameAndPath', bold)
        worksheet1.write('F1', 'Percentage_of_Similarity', bold)
        
        if len(samelist) > 0:
            for smlist in samelist:
                for fl in smlist:
                    row +=1
                    worksheet1.write_string(row,col,fl[0])
                    worksheet1.write_string(row,col+1,fl[1])
                    worksheet1.write_string(row,col+2,fl[2])
        
        if len(chngNfllist) > 0:
            for tplist in chngNfllist:
                for tp in tplist:
                    row +=1
                    worksheet1.write_string(row,col,tp[0])
                    worksheet1.write_string(row,col+1,tp[1])
                    worksheet1.write_string(row,col+2,tp[2])
                    
                    
                    
        if len(chngfllist) > 0:
            for tpchlist in chngfllist:
                for tpfl in tpchlist:
                    row +=1             
                    worksheet1.write_string(row,col,tpfl[0])
                    worksheet1.  write_string(row,col+1,tpfl[1])
                    worksheet1.  write_string(row,col+2,tpfl[2])
                    
                  
          
                        
              
    row = 0
    col = 7
    worksheet1.write('H1', 'Statistic', bold)
    worksheet1.write('I1', 'Count/Percentage', bold)
    row +=1             
    worksheet1.write_string(row,col,'TotalFileCount')
   # worksheet1.write_string(row,col+1,exceltp[3])
    worksheet1.write_number(row,col+1,exceltp[4])
    row +=1             
    worksheet1.write_string(row,col,'MatchFileCount')
    #worksheet1.write_string(row,col+1,exceltp[4])
    worksheet1.write_number(row,col+1,exceltp[5])
    row +=1             
    worksheet1.write_string(row,col,'TotalChangeCount')
   # worksheet1.write_string(row,col+1,exceltp[5])
    worksheet1.write_number(row,col+1,exceltp[6])
    row +=1             
    worksheet1.write_string(row,col,'90% Or More similar')
   # worksheet1.write_string(row,col+1,exceltp[6])
    worksheet1.write_number(row,col+1,exceltp[7])
    row +=1             
    worksheet1.write_string(row,col,'AddFileCount')
    #worksheet1.write_string(row,col+1,exceltp[7])
    worksheet1.write_number(row,col+1,exceltp[8])
    row +=1             
    worksheet1.write_string(row,col,'PercentageOfChange')
    #worksheet1.write_string(row,col+1,exceltp[8])
    worksheet1.write_number(row,col+1,exceltp[9])
#    if flag:
#        wb2.save('C:\\TestTmp\\StatFile.xlsx')      
        
    
    workbook.close()
 
     
def getChangeCount(dr1,dr2):
    
    emptyTp =()
    emptydir = ()
    listOfFiles1 = list()
    for (dirpath, dirnames, filenames) in os.walk(dr1):
        listOfFiles1 += [os.path.join(dirpath, file) for file in filenames]
     

     
    listOfFiles2 = list()
    for (dirpath, dirnames, filenames) in os.walk(dr2):
        listOfFiles2 += [os.path.join(dirpath, file) for file in filenames]
        
    
    print (listOfFiles1)
    print(listOfFiles2)
    emptydir = createandCopyDir(listOfFiles1,listOfFiles2)
    dirval1 = emptydir[0]
    dirval2 = emptydir[1]
    emptyTp = getcntfiles(dirval1,dirval2,listOfFiles1) 
    shutil.rmtree(emptydir[0], ignore_errors=True)
    shutil.rmtree(emptydir[1], ignore_errors=True)
    return emptyTp
    
def dirandFilecompre(fn1,fn2,sheetnm):
    chngNfllist=list()
    chngfllist=list()
    addfllist=list()
    
    newentytp= getChangeCount(fn1,fn2)
    print('totfile',newentytp[0])
    print('match',newentytp[1])
    print('change',newentytp[2])
    print('ninepercemat',newentytp[3])
    print('add',newentytp[4])
    if len(newentytp[5]) > 0:
        chngNfllist.append(newentytp[5])
    if len(newentytp[6]) > 0:
        chngfllist.append(newentytp[6])
    if len(newentytp[7]) > 0:
        addfllist.append(newentytp[7])
        
    
    percetange=(newentytp[2] + newentytp[4])  * 100 / newentytp[0]
    exceltp=(addfllist,chngNfllist,chngfllist,newentytp[0],newentytp[1],newentytp[2],newentytp[3],newentytp[4],percetange)
    writExcelFile(exceltp,sheetnm)
    return percetange

def createandCopyDir(l1,l2):
    
    dirName1 = 'C:\\testtemp\\tempDiror'
    dirName2 = 'C:\\testtemp\\tempDirmod'
    emptydir = ()
    try:
        # Create target Directory
        
        
        os.mkdir(dirName1)
      
        print("Directory " , dirName1 , dirName2, " Created ") 
    except FileExistsError:
        print("Directory " , dirName1 ,  "already exists")
        
    try:
        os.mkdir(dirName2)
        print("Directory " , dirName1 , dirName2, " Created ") 
    except FileExistsError:
        print("Directory " , dirName1 ,  "already exists")
     
        
    for f in l1:
       # newext= f.replace('.pcf','.xml')
        shutil.copy(f, dirName1)
    for ff in l2:
     #   newext1= ff.replace('.pcf','.xml')
        shutil.copy(ff, dirName2)
    
    emptydir = (dirName1,dirName2)
    
    #shutil.rmtree(dirName1, ignore_errors=True)
   # shutil.rmtree(dirName2, ignore_errors=True)
    
    return emptydir


def getcntfiles(dr1,dr2,l1):
    emptyTp = ()
    
    #filelist = filecmp.dircmp(dr1,dr2).diff_files
  #  totflcnt = 0
   # changflcnt = 0
    print ('dr1',dr1)
    print ('dr2',dr2)
   # os.chdir(curdir)
    d1_contents = list(os.listdir(dr1))
    d2_contents = list(os.listdir(dr2))
    
   # print ('d1_contents',d1_contents)
  #  print ('d2_contents',d2_contents)
    
    mcnt=0
    miscnt = 0
    flag = False
    mninety=0
    addfile=0
    totfile=0
    changefl = 0
    changelist = list()
    addlist=list()
    changeNlist=list()
    samelist=list()

    for i in range(0,len(d1_contents)):
        skipcnt=0
        changetp=()
        addtp=()
        for pathfl in l1:
            if d1_contents[i]  in pathfl:
                flpth=pathfl
                break
        lenth = len(d1_contents[i])
        originalpth = flpth[0:len(flpth)-lenth-1]    
        dr3=getImmidiatedir(originalpth)  
        pcflenth = len('C:\GW\workspace\v10_conversion\ClaimCenter\modules\configuration\config\web\\')
        abspathh=originalpth[pcflenth:]

        for j in range(0,len(d2_contents)):
            
            if totfile != skipcnt:
                skipcnt = skipcnt+1
                flag = True
                continue
            
            #size=0       
            print('mod',d1_contents[i])
            print('chnage',d2_contents[j])
            if d1_contents[i] == d2_contents[j]:
                print(d1_contents[i])
                print(d2_contents[j])
                totfile = totfile + 1
                filepathor= open(os.path.join(dr1,d1_contents[i]),encoding='utf-8').read()
                filepathmod =open(os.path.join(dr2,d2_contents[j]),encoding='utf-8').read()
                m = SequenceMatcher(None, filepathor, filepathmod)
                changePer = m.ratio()*100
                if str(m.ratio()*100)[:1] == '1':
                    mcnt= mcnt + 1
              #  totcnt = totcnt +1
                    flag = False
                    #dr3=getImmidiatedir(dr1)
                    flnm = os.path.join(dr3, d1_contents[i])
                    sametp = (abspathh,flnm,'Same')
                    samelist.append(sametp)
                    break
                elif str(m.ratio()*100)[:2] >= '90':
                    #mcnt= mcnt + 1
                    mninety = mninety + 1
                   # dr3=getImmidiatedir(dr1)
                    flnm = os.path.join(dr3, d1_contents[i])
                    changeNtp = (abspathh,flnm,str(changePer))
                    changeNlist.append(changeNtp)
                    flag = False
                    break
                    
                else:
                    #mcnt= mcnt + 1
                    changefl= changefl+1
                  #  dr3=getImmidiatedir(dr1)
                    flnm = os.path.join(dr3, d1_contents[i])
                    changetp = (abspathh,flnm,str(changePer))
                    changelist.append(changetp)
                    #changelist.append(d1_contents[i])
              #  totcnt = totcnt +1
                    flag = False
                    break
            else:
                flag = False
                miscnt= miscnt+1
                addfile = addfile + 1
             #   dr3=getImmidiatedir(dr1)
                flnm = os.path.join(dr3, d1_contents[i])
                addtp=(abspathh,flnm)
                addlist.append(addtp)
                
               # changelist.append(d1_contents[i])
                break
        
        
        
        if flag:
            miscnt= miscnt+1
            #if d2_contents[j] !='order.txt':
            addfile = addfile + 1
          #  dr3=getImmidiatedir(dr1)
            flnm = os.path.join(dr3, d1_contents[i])
            #filepathorpathh  = os.path.join(dr1,d1_contents[i])
            addtp=(abspathh,flnm)
            addlist.append(addtp)
  
   
    emptyTp = (totfile,mcnt,changefl,mninety,addfile,changeNlist,changelist,addlist,samelist)
    
    return emptyTp

        
def getImmidiatedir(dr3):
    words2 = dr3.split("\\")
    for dirn in words2:
        fndir= dirn
    
    return fndir
    
def getFileCnt(dr1):
    listOfFiles1 = list()
   # listofdirwithFile=list()
    for (dirpath, dirnames, filenames) in os.walk(dr1):
        listOfFiles1 += [os.path.join(file) for file in filenames]
        print('listOfFiles1extra', listOfFiles1)
        return listOfFiles1
 
def main():
    mainDir = 'C:\\testtemp'
   
    if not os.path.exists(mainDir):
        os.mkdir(mainDir)
        print("Directory " , mainDir , " Created ") 
    else:
        print("Directory " , mainDir ,  "already exists")
        
    
    
    d1='C:\\GW\\workspace\\v10_conversion\\ClaimCenter\\modules\\configuration\\config\\web\\pcf'
    d2='D:\\ClaimCenter10\\modules\\configuration\\config\\web\\pcf'
    sheetnm='PCF'
    pcfper = getdirname(d1,d2,sheetnm)
    print(pcfper) 
    

if __name__== "__main__":
  main()