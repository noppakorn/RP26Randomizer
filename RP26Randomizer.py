# By Noppakorn Jiravaranun
import pandas as pd
import random

def Diff(li1, li2): return list(set(li1).difference(set(li2)))
def Select(In,L,Out=[]):
    if L-len(Out) != 0 and len(In) > L-len(Out): Out.extend(random.sample(In, L-len(Out)))
    elif L-len(Out) == 0 : pass
    else: Out.extend(In)
    return Out
NApp = 10
NDoc = 10
NSpon = 9
NPR = 10
NGra = 5
NVid = 2
NPho = 4
NChap = 36
NChapM = 12
NChapF = 12
NAct = 53
NRec = 32
NRecD = 3
NAud = 4
NLog = 13

df = pd.read_excel('RP26_Responses_1.xlsx',engine='openpyxl')
App_1 = list(df[df['First'] == 'ฝ่ายรับสมัคร'].loc[:,'Name'])
Doc_1 = list(df[df['First'] == 'ฝ่ายเอกสารทั่วไป'].loc[:,'Name'])
Spon_1= list(df[df['First'] == 'ฝ่าย Sponsor'].loc[:,'Name'])
PR_1 = list(df[df['First'] == 'ฝ่ายประชาสัมพันธ์'].loc[:,'Name'])
Gra_1 = list(df[(df['First'] == 'ฝ่าย Graphic') & (df['Graphic'] == 'ได้')].loc[:,'Name'])
Act_1 = list(df[df['First'] == 'ฝ่ายกิจกรรมทั่วไป'].loc[:,'Name'])
Aud_1 = list(df[df['First'] == 'ฝ่ายโสต'].loc[:,'Name'])
Log_1 = list(df[df['First'] == 'ฝ่าย Logist'].loc[:,'Name'])
Chap_1 = list(df[df['First'] == 'ฝ่ายพี่กลุ่ม'].loc[:,'Name'])
ChapM_1 = list(df[(df['First'] == 'ฝ่ายพี่กลุ่ม') & (df['Gender'] == 'Male')].loc[:,'Name'])
ChapF_1 = list(df[(df['First'] == 'ฝ่ายพี่กลุ่ม') & (df['Gender'] == 'Female')].loc[:,'Name'])
Pho_1 = list(df[(df['First'] == 'ฝ่าย Photo&Video') & (df['Camera'] == 'ได้')].loc[:,'Name'])
Vid_1 = list(df[(df['First'] == 'ฝ่าย Photo&Video') & (df['Editing'] == 'ได้')].loc[:,'Name'])
Rec_1 = list(df[df['First'] == 'ฝ่ายสันทนาการ'].loc[:,'Name'])
RecD_1 = list(df[(df['First'] == 'ฝ่ายสันทนาการ') & (df['Drum'] == 'ได้')].loc[:,'Name'])

Application = Select(App_1,NApp,[])
Document = Select(Doc_1,NDoc,[])
Sponsor = Select(Spon_1,NSpon,[])
PR = Select(PR_1,NPR,[])
Graphic = Select(Gra_1,NGra,[])
Activities = Select(Act_1,NAct,[])
Audio = Select(Aud_1,NAud,[])
Logist = Select(Log_1,NLog,[])

ChapM = Select(ChapM_1,NChapM,[])
ChapF = Select(ChapF_1,NChapF,[])
Chap = Select(Diff(Chap_1,ChapM+ChapF),NChap-NChapM-NChapF,[])

Video = Select(Vid_1,NVid,[])
Photo = Select(Diff(Pho_1,Video),NPho,[])

RecD = Select(RecD_1,NRecD,[])
Rec = Select(Diff(Rec_1,RecD),NRec-NRecD,[])
# End of Round 1
df2 = df.set_index('Name')
df2 = df2.drop([i for i in Application+Document+Sponsor+PR+Graphic+Activities+Audio+Logist+ChapM+ChapF+Chap+Video+Photo+RecD+Rec],axis=0)
df2 = df2.reset_index()
App_2 = list(df2[df2['Second'] == 'ฝ่ายรับสมัคร'].loc[:,'Name'])
Doc_2 = list(df2[df2['Second'] == 'ฝ่ายเอกสารทั่วไป'].loc[:,'Name'])
Spon_2= list(df2[df2['Second'] == 'ฝ่าย Sponsor'].loc[:,'Name'])
PR_2 = list(df2[df2['Second'] == 'ฝ่ายประชาสัมพันธ์'].loc[:,'Name'])
Gra_2 = list(df2[(df2['Second'] == 'ฝ่าย Graphic') & (df2['Graphic'] == 'ได้')].loc[:,'Name'])
Act_2 = list(df2[df2['Second'] == 'ฝ่ายกิจกรรมทั่วไป'].loc[:,'Name'])
Aud_2 = list(df2[df2['Second'] == 'ฝ่ายโสต'].loc[:,'Name'])
Log_2 = list(df2[df2['Second'] == 'ฝ่าย Logist'].loc[:,'Name'])
Chap_2 = list(df2[df2['Second'] == 'ฝ่ายพี่กลุ่ม'].loc[:,'Name'])
ChapM_2 = list(df2[(df2['Second'] == 'ฝ่ายพี่กลุ่ม') & (df2['Gender'] == 'Male')].loc[:,'Name'])
ChapF_2 = list(df2[(df2['Second'] == 'ฝ่ายพี่กลุ่ม') & (df2['Gender'] == 'Female')].loc[:,'Name'])
Pho_2 = list(df2[(df2['Second'] == 'ฝ่าย Photo&Video') & (df2['Camera'] == 'ได้')].loc[:,'Name'])
Vid_2 = list(df2[(df2['Second'] == 'ฝ่าย Photo&Video') & (df2['Editing'] == 'ได้')].loc[:,'Name'])
Rec_2 = list(df2[df2['Second'] == 'ฝ่ายสันทนาการ'].loc[:,'Name'])
RecD_2 = list(df2[(df2['Second'] == 'ฝ่ายสันทนาการ') & (df2['Drum'] == 'ได้')].loc[:,'Name'])

Application = Select(App_2,NApp,Application)
Document = Select(Doc_2,NDoc,Document)
Sponsor = Select(Spon_2,NSpon,Sponsor)
PR = Select(PR_2,NPR,PR)
Graphic = Select(Gra_2,NGra,Graphic)
Activities = Select(Act_2,NAct,Activities)
Audio = Select(Aud_2,NAud,Audio)
Logist = Select(Log_2,NLog,Logist)

ChapM = Select(ChapM_2,NChapM,ChapM)
ChapF = Select(ChapF_2,NChapF,ChapF)
Chap = Select(Diff(Chap_2,ChapM+ChapF),NChap-NChapM-NChapF,Chap)

Video = Select(Vid_2,NVid,Video)
Photo = Select(Diff(Pho_2,Video),NPho,Photo)

RecD = Select(RecD_2,NRecD,RecD)
Rec = Select(Diff(Rec_2,RecD),NRec-NRecD,Rec)
# End of Round 2
df3 = df.set_index('Name')
df3 = df3.drop([i for i in Application+Document+Sponsor+PR+Graphic+Activities+Audio+Logist+ChapM+ChapF+Chap+Video+Photo+RecD+Rec],axis=0)
df3 = df3.reset_index()
App_3 = list(df3[df3['Third'] == 'ฝ่ายรับสมัคร'].loc[:,'Name'])
Doc_3 = list(df3[df3['Third'] == 'ฝ่ายเอกสารทั่วไป'].loc[:,'Name'])
Spon_3= list(df3[df3['Third'] == 'ฝ่าย Sponsor'].loc[:,'Name'])
PR_3 = list(df3[df3['Third'] == 'ฝ่ายประชาสัมพันธ์'].loc[:,'Name'])
Gra_3 = list(df3[(df3['Third'] == 'ฝ่าย Graphic') & (df3['Graphic'] == 'ได้')].loc[:,'Name'])
Act_3 = list(df3[df3['Third'] == 'ฝ่ายกิจกรรมทั่วไป'].loc[:,'Name'])
Aud_3 = list(df3[df3['Third'] == 'ฝ่ายโสต'].loc[:,'Name'])
Log_3 = list(df3[df3['Third'] == 'ฝ่าย Logist'].loc[:,'Name'])
Chap_3 = list(df3[df3['Third'] == 'ฝ่ายพี่กลุ่ม'].loc[:,'Name'])
ChapM_3 = list(df3[(df3['Third'] == 'ฝ่ายพี่กลุ่ม') & (df3['Gender'] == 'Male')].loc[:,'Name'])
ChapF_3 = list(df3[(df3['Third'] == 'ฝ่ายพี่กลุ่ม') & (df3['Gender'] == 'Female')].loc[:,'Name'])
Pho_3 = list(df3[(df3['Third'] == 'ฝ่าย Photo&Video') & (df3['Camera'] == 'ได้')].loc[:,'Name'])
Vid_3 = list(df3[(df3['Third'] == 'ฝ่าย Photo&Video') & (df3['Editing'] == 'ได้')].loc[:,'Name'])
Rec_3 = list(df3[df3['Third'] == 'ฝ่ายสันทนาการ'].loc[:,'Name'])
RecD_3 = list(df3[(df3['Third'] == 'ฝ่ายสันทนาการ') & (df3['Drum'] == 'ได้')].loc[:,'Name'])

Application = Select(App_3,NApp,Application)
Document = Select(Doc_3,NDoc,Document)
Sponsor = Select(Spon_3,NSpon,Sponsor)
PR = Select(PR_3,NPR,PR)
Graphic = Select(Gra_3,NGra,Graphic)
Activities = Select(Act_3,NAct,Activities)
Audio = Select(Aud_3,NAud,Audio)
Logist = Select(Log_3,NLog,Logist)

ChapM = Select(ChapM_3,NChapM,ChapM)
ChapF = Select(ChapF_3,NChapF,ChapF)
Chap = Select(Diff(Chap_3,ChapM+ChapF),NChap-NChapM-NChapF,Chap)

Video = Select(Vid_3,NVid,Video)
Photo = Select(Diff(Pho_3,Video),NPho,Photo)

RecD = Select(RecD_3,NRecD,RecD)
Rec = Select(Diff(Rec_3,RecD),NRec-NRecD,Rec)
# End of Round 3
df4 = df.set_index('Name')
df4 = df4.drop([i for i in Application+Document+Sponsor+PR+Graphic+Activities+Audio+Logist+ChapM+ChapF+Chap+Video+Photo+RecD+Rec],axis=0)
df4 = df4.reset_index()
App_4 = list(df4[df4['Forth'] == 'ฝ่ายรับสมัคร'].loc[:,'Name'])
Doc_4 = list(df4[df4['Forth'] == 'ฝ่ายเอกสารทั่วไป'].loc[:,'Name'])
Spon_4= list(df4[df4['Forth'] == 'ฝ่าย Sponsor'].loc[:,'Name'])
PR_4 = list(df4[df4['Forth'] == 'ฝ่ายประชาสัมพันธ์'].loc[:,'Name'])
Gra_4 = list(df4[(df4['Forth'] == 'ฝ่าย Graphic') & (df4['Graphic'] == 'ได้')].loc[:,'Name'])
Act_4 = list(df4[df4['Forth'] == 'ฝ่ายกิจกรรมทั่วไป'].loc[:,'Name'])
Aud_4 = list(df4[df4['Forth'] == 'ฝ่ายโสต'].loc[:,'Name'])
Log_4 = list(df4[df4['Forth'] == 'ฝ่าย Logist'].loc[:,'Name'])
Chap_4 = list(df4[df4['Forth'] == 'ฝ่ายพี่กลุ่ม'].loc[:,'Name'])
ChapM_4 = list(df4[(df4['Forth'] == 'ฝ่ายพี่กลุ่ม') & (df4['Gender'] == 'Male')].loc[:,'Name'])
ChapF_4 = list(df4[(df4['Forth'] == 'ฝ่ายพี่กลุ่ม') & (df4['Gender'] == 'Female')].loc[:,'Name'])
Pho_4 = list(df4[(df4['Forth'] == 'ฝ่าย Photo&Video') & (df4['Camera'] == 'ได้')].loc[:,'Name'])
Vid_4 = list(df4[(df4['Forth'] == 'ฝ่าย Photo&Video') & (df4['Editing'] == 'ได้')].loc[:,'Name'])
Rec_4 = list(df4[df4['Forth'] == 'ฝ่ายสันทนาการ'].loc[:,'Name'])
RecD_4 = list(df4[(df4['Forth'] == 'ฝ่ายสันทนาการ') & (df4['Drum'] == 'ได้')].loc[:,'Name'])

Application = Select(App_4,NApp,Application)
Document = Select(Doc_4,NDoc,Document)
Sponsor = Select(Spon_4,NSpon,Sponsor)
PR = Select(PR_4,NPR,PR)
Graphic = Select(Gra_4,NGra,Graphic)
Activities = Select(Act_4,NAct,Activities)
Audio = Select(Aud_4,NAud,Audio)
Logist = Select(Log_4,NLog,Logist)

ChapM = Select(ChapM_4,NChapM,ChapM)
ChapF = Select(ChapF_4,NChapF,ChapF)
Chap = Select(Diff(Chap_4,ChapM+ChapF),NChap-NChapM-NChapF,Chap)

Video = Select(Vid_4,NVid,Video)
Photo = Select(Diff(Pho_4,Video),NPho,Photo)

RecD = Select(RecD_4,NRecD,RecD)
Rec = Select(Diff(Rec_4,RecD),NRec-NRecD,Rec)
# End of Round 4
df5 = df.set_index('Name')
df5 = df5.drop([i for i in Application+Document+Sponsor+PR+Graphic+Activities+Audio+Logist+ChapM+ChapF+Chap+Video+Photo+RecD+Rec],axis=0)
df5 = df5.reset_index()
App_5 = list(df5[df5['Fifth'] == 'ฝ่ายรับสมัคร'].loc[:,'Name'])
Doc_5 = list(df5[df5['Fifth'] == 'ฝ่ายเอกสารทั่วไป'].loc[:,'Name'])
Spon_5= list(df5[df5['Fifth'] == 'ฝ่าย Sponsor'].loc[:,'Name'])
PR_5 = list(df5[df5['Fifth'] == 'ฝ่ายประชาสัมพันธ์'].loc[:,'Name'])
Gra_5 = list(df5[(df5['Fifth'] == 'ฝ่าย Graphic') & (df5['Graphic'] == 'ได้')].loc[:,'Name'])
Act_5 = list(df5[df5['Fifth'] == 'ฝ่ายกิจกรรมทั่วไป'].loc[:,'Name'])
Aud_5 = list(df5[df5['Fifth'] == 'ฝ่ายโสต'].loc[:,'Name'])
Log_5 = list(df5[df5['Fifth'] == 'ฝ่าย Logist'].loc[:,'Name'])
Chap_5 = list(df5[df5['Fifth'] == 'ฝ่ายพี่กลุ่ม'].loc[:,'Name'])
ChapM_5 = list(df5[(df5['Fifth'] == 'ฝ่ายพี่กลุ่ม') & (df5['Gender'] == 'Male')].loc[:,'Name'])
ChapF_5 = list(df5[(df5['Fifth'] == 'ฝ่ายพี่กลุ่ม') & (df5['Gender'] == 'Female')].loc[:,'Name'])
Pho_5 = list(df5[(df5['Fifth'] == 'ฝ่าย Photo&Video') & (df5['Camera'] == 'ได้')].loc[:,'Name'])
Vid_5 = list(df5[(df5['Fifth'] == 'ฝ่าย Photo&Video') & (df5['Editing'] == 'ได้')].loc[:,'Name'])
Rec_5 = list(df5[df5['Fifth'] == 'ฝ่ายสันทนาการ'].loc[:,'Name'])
RecD_5 = list(df5[(df5['Fifth'] == 'ฝ่ายสันทนาการ') & (df5['Drum'] == 'ได้')].loc[:,'Name'])

Application = Select(App_5,NApp,Application)
Document = Select(Doc_5,NDoc,Document)
Sponsor = Select(Spon_5,NSpon,Sponsor)
PR = Select(PR_5,NPR,PR)
Graphic = Select(Gra_5,NGra,Graphic)
Activities = Select(Act_5,NAct,Activities)
Audio = Select(Aud_5,NAud,Audio)
Logist = Select(Log_5,NLog,Logist)

ChapM = Select(ChapM_5,NChapM,ChapM)
ChapF = Select(ChapF_5,NChapF,ChapF)
Chap = Select(Diff(Chap_5,ChapM+ChapF),NChap-NChapM-NChapF,Chap)

Video = Select(Vid_5,NVid,Video)
Photo = Select(Diff(Pho_5,Video),NPho,Photo)

RecD = Select(RecD_5,NRecD,RecD)
Rec = Select(Diff(Rec_5,RecD),NRec-NRecD,Rec)
# End of Round 5
dff = df.set_index('Name')
dff = dff.drop([i for i in Application+Document+Sponsor+PR+Graphic+Activities+Audio+Logist+ChapM+ChapF+Chap+Video+Photo+RecD+Rec],axis=0)
dff = dff.reset_index()

print(Application)
print(Document)
print(Sponsor)
print(PR)
print(Graphic)
print(Activities)
print(Audio)
print(Logist)
print('---')
print(ChapM)
print(ChapF)
print(Chap)
print('---')
print(Video)
print(Photo)
print('---')
print(Rec)
print(RecD)

print('Round 0',len(df))
print('Round 1',len(df2))
print('Round 2',len(df3))
print('Round 3',len(df4))
print('Round 4',len(df5))
print('Round 5',len(dff))