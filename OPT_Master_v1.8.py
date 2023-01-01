import pandas as pd
import glob
import os
import io
import numpy as np
from dateutil.parser import parse
from tika import parser
import re
import sqlite3
import shutil
import win32com.client
pathwin32=win32com.__gen_path__
directory_contents = os.listdir(pathwin32)
for item in directory_contents:    
    if os.path.isdir(os.path.join(pathwin32, item)):        
        shutil.rmtree(os.path.join(pathwin32, item))
import patoolib
import random
import PyPDF2
from datetime import datetime
from datetime import date
import tkinter as tk  
from tkinter import filedialog
from tkinter import messagebox

root = tk.Tk()
root.withdraw()

now=datetime.strftime(datetime.now(), "%Y-%m-%d %H:%M:%S")
today=datetime.strftime(datetime.now(), "%Y-%m-%d")
todayasvers=datetime.strftime(datetime.now(), "%d_%m_%Y")


currpath= os.getcwd()
odpath=currpath.replace("phyton\db_python\OPT","")
userpath=currpath.replace("TURISTIK HAVA TASIMACILIK A.S\Gökmen Düzgören - FOE_2019\phyton\db_python\OPT", "")
shutil.copy(os.path.join(odpath , "airport", "Current Airport DB", "airport.sdb"), os.path.join(odpath ,  "airport", "Current Airport DB", today + "_airport.sdb"))
conn=sqlite3.connect(os.path.join(odpath ,  "airport", "Current Airport DB", today + "_airport.sdb"))
os.environ['TIKA_SERVER_JAR'] = os.path.join(userpath, "AppData","Local","Programs","Python","Python37", "tika-server-1.24.jar")

def randomgen():
    checktrue=True
    while checktrue==True:
        rand_int=random.randint(150000000, 300000000)
        crand=conn.cursor()
        crand.execute("SELECT * from ObstInfo WHERE Number=?", [rand_int])
        ftcrand=crand.fetchone()
        if ftcrand:
            checktrue=True
            
        else:
            checktrue=False
    return rand_int


def Exceptions():
   
    #### LTAS #####
    
    #EOP18="Follow SID. "
    EOP36="Special SID procedure for all conditions: Climb on RWY track – crossing R080 CAY LT follow D2.1 CAY Arc – crossing R043 CAY RT 350 to follow the river northbound."
    cexc=conn.cursor()
    cexc.execute("DELETE from EOProced WHERE Code='LTAS' AND ID='18TMP'")
    cexc.execute("DELETE from EOProced WHERE Code='LTAS' AND ID='36TMP'")
    #cexc.execute("UPDATE EOProced SET EOProc=? WHERE Code='LTAS' AND ID='18'", [EOP18])
    cexc.execute("UPDATE EOProced SET EOProc=? WHERE Code='LTAS' AND ID='36'", [EOP36])
    cexc.execute("DELETE from GAProced WHERE Code='LTAS' AND ID='18TMP'")
    cexc.execute("DELETE from GAProced WHERE Code='LTAS' AND ID='36TMP'")
    cexc.execute("DELETE from NOTAMinfo WHERE Code='LTAS' AND ID='18TMP'")
    cexc.execute("DELETE from NOTAMinfo WHERE Code='LTAS' AND ID='36TMP'")
    cexc.execute("DELETE from ObstInfo WHERE Code=?", ['LTAS'])
    random_int1=randomgen()
    cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
    VALUES('LTAS','18',"",?,3100,330,0)""",[random_int1])
    conn.commit()
    random_int2=randomgen()
    cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
    VALUES('LTAS','36',"",?,2800,386,0)""",[random_int2])
    cexc.execute("DELETE from RwyInfo WHERE Code='LTAS' AND ID='18TMP'")
    cexc.execute("DELETE from RwyInfo WHERE Code='LTAS' AND ID='36TMP'")
    conn.commit()

    #### LTAU #####

    EOP07="NON-STD. At D3.0 KSR turn LEFT to KSR HP. D116.3 KSR HP: Inbound 250° RIGHT turn."
    EOP25="NON-STD. At D8.0 KSR turn RIGHT to KSR HP. D116.3 KSR HP: Inbound 250° RIGHT turn."
    
    cexc.execute("UPDATE EOProced SET EOProc=? WHERE Code='LTAU' AND ID='07'", [EOP07])
    cexc.execute("UPDATE EOProced SET EOProc=? WHERE Code='LTAU' AND ID='25'", [EOP25])
    cexc.execute("DELETE from ObstInfo WHERE Code=?", ['LTAU'])
    dist07=(403, 499,517)
    ht07=(12, 15, 25)
    dist25=(4286, 4263,4249, 4255, 4454)
    ht25=(187, 186, 183, 182, 252)
    for cntex1 in range(len(dist07)):
        random_int3=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('LTAU','07',"",?,?,?,0)""",[random_int3, dist07[cntex1], ht07[cntex1]])
        conn.commit()

    for cntex2 in range(len(dist25)):
        random_int4=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('LTAU','25',"",?,?,?,0)""",[random_int4, dist25[cntex2], ht25[cntex2]])
        conn.commit()


    #### GLRB #####

    
    cexc.execute("DELETE from ObstInfo WHERE Code=? AND ID=?", ['GLRB', '22'])
   
    dist22=(670, 3200)
    ht22=(70, 195)
    for cntex1 in range(len(dist22)):
        random_int4=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('GLRB','22',"",?,?,?,0)""",[random_int4, dist22[cntex1], ht22[cntex1]])
        conn.commit()



    #### DFFD #####
    cexc.execute("DELETE from ObstInfo WHERE Code=?", ['DFFD'])
   
    dist04=(286, 856, 1483, 1547)
    ht04=(13, 75, 36, 69)
    for cntex5 in range(len(dist04)):
        random_int55=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('DFFD','04',"",?,?,?,0)""",[random_int55, dist04[cntex5], ht04[cntex5]])
        conn.commit()


    
    dist22=(1006, 1096, 1116, 1156)
    ht22=(46, 30,89,92)
    for cntex6 in range(len(dist22)):
        random_int66=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('DFFD','22',"",?,?,?,0)""",[random_int66, dist22[cntex6], ht22[cntex6]])
        conn.commit()

    


    #### WSSS ####
    cexc.execute("DELETE from ObstInfo WHERE Code=?", ['WSSS'])
    cexc.execute("DELETE from IntersectInfo WHERE Code=?", ['WSSS'])
    
    cexc.execute("DELETE from RwyInfo WHERE Code=? AND ID=?", ['WSSS', '20L'])
    cexc.execute("DELETE from RwyInfo WHERE Code=? AND ID=?", ['WSSS', '02R'])
    cexc.execute("DELETE from RwyInfo WHERE Code=? AND ID=?", ['WSSS', '20R'])

    cexc.execute("UPDATE RwyInfo SET TODA=?, ASDA=?, SlopeTOD=?, SlopeASD=?, SlopeLDA=? WHERE Code=? AND ID=?", [0,0,0, 0, 0, 'WSSS', '02C'])
    cexc.execute("UPDATE RwyInfo SET TODA=?, ASDA=?, SlopeTOD=?, SlopeASD=?, SlopeLDA=? WHERE Code=? AND ID=?", [0,0,1.99, 1.99, 1.99, 'WSSS', '20C'])
    cexc.execute("UPDATE RwyInfo SET TODA=?, ASDA=?,  SlopeTOD=?, SlopeASD=?, SlopeLDA=? WHERE Code=? AND ID=?", [0,0,-1.99, -1.99, -1.99, 'WSSS', '02L'])
    cexc.execute("UPDATE AptInfo SET Elevation=? WHERE Code=?", [0, 'WSSS'])
    conn.commit()


    ##### WMKL  ####
    
    cexc.execute("DELETE from ObstInfo WHERE Code=?", ['WMKL'])
    cexc.execute("DELETE from IntersectInfo WHERE Code=?", ['WMKL'])


    cexc.execute("UPDATE RwyInfo SET TORA=?, XLDA=?, SlopeTOD=?, SlopeASD=?, SlopeLDA=? WHERE Code=? AND ID=?", [4000, 4000 ,0, 0, 0, 'WMKL', '03'])
    cexc.execute("UPDATE RwyInfo SET TORA=?, XLDA=?, SlopeTOD=?, SlopeASD=?, SlopeLDA=? WHERE Code=? AND ID=?", [4000, 4000 ,1.99, 1.99, 1.99, 'WMKL', '21']) 
    cexc.execute("UPDATE AptInfo SET Elevation=? WHERE Code=?", [1000, 'WMKL'])
    conn.commit()

    ##### LPMA  ####
    
    EOP05="NON-STD. At 400 or D2.0 FUN, whichever first, turn RIGHT (15° bank angle) to 074°. At 2000 PROCEED to FUN HP. Maintain V2 TKOF flaps and max 161 KIAS during first turn. D112.2 FUN HP: Inbound 199°, LEFT turn."
    EOP23="NON-STD. At 400 or D5.5 FUN, whichever first, turn LEFT to 089°. (15° bank angle) At 2000 PROCEED to FUN HP. Maintain V2 TKOF flaps and max 161 KIAS during first turn. D112.2 FUN HP: Inbound 199°, LEFT turn."
    
    cexc.execute("UPDATE EOProced SET EOProc=? WHERE Code='LPMA' AND ID='05'", [EOP05])
    cexc.execute("UPDATE EOProced SET EOProc=? WHERE Code='LPMA' AND ID='23'", [EOP23])
    conn.commit()

    

    #### VOBM #####
       
    cexc.execute("""INSERT INTO AptInfo (Code, IATA, Name, City, Country, Elevation,RwyMeasType, DistUnit, HtUnit,ObDistRef, ObHtRef, CrDate, UpdDate,MagVar)
    VALUES("VOBM", "IXG", "BELAGAVI", "BELAGAVI", "INDIA", 2489, 1,0,1,1,1, ?,?, 0)""", [now, now])
    conn.commit()

    EOP08="After takeoff, maintain runway heading, contact ATC."
    EOP26="After takeoff, maintain runway heading, contact ATC."
    cexc.execute("""INSERT INTO EOProced (Code, ID, EOProc, MinV2, MaxV2, SpeedUnits)
    VALUES("VOBM", "08", ?, 0, 0, "KTAS")""", [EOP08])
    cexc.execute("""INSERT INTO EOProced (Code, ID, EOProc, MinV2, MaxV2, SpeedUnits)
    VALUES("VOBM", "26", ?, 0, 0, "KTAS")""", [EOP26])
    conn.commit()

    cexc.execute("""INSERT INTO GAProced (Code, ID, gaGradient, DecisionHt, DeltaHt)
    VALUES("VOBM", "08", 2.5, 0, 0)""")
    cexc.execute("""INSERT INTO GAProced (Code, ID, gaGradient, DecisionHt, DeltaHt)
    VALUES("VOBM", "26", 2.5, 0, 0)""")
    conn.commit()

    dist08=(68, 336, 2260, 2462, 3863, 4046, 4136, 19523, 20538, 21146)
    ht08=(6, 12, 30, 44, 78, 94, 98, 358, 393, 410)
    for cntex6 in range(len(dist08)):
        random_int77=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('VOBM','08',"",?,?,?,0)""",[random_int77, dist08[cntex6]*.3048, ht08[cntex6]])
        conn.commit()
    
    dist26=(126, 11642, 13965, 38853)
    ht26=(7, 145, 257, 547)
    for cntex6 in range(len(dist26)):
        random_int66=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('VOBM','26',"",?,?,?,0)""",[random_int66, dist26[cntex6]*.3048, ht26[cntex6]])
        conn.commit()

    cexc.execute("""INSERT INTO RwyInfo (Code,ID,TORA,TODA,ASDA,XLDA,Width,Surface,SlopeTOD,SlopeASD,SlopeLDA,CrDate,UpdDate,CrTime,UpTime,eoLabel,lineupAngle,
    elevStartTORA,elevEndTORA,MagHdg,gaLabel,mfrh,useForTakeoff,useForLanding,mfrhType)
    VALUES("VOBM", "08", 2300, 2300, 2300, 2300, 45, 0, -0.2, -0.2,-0.2, ?,?,?,?, "Engine Failure Procedure", 180, 0, 0, 81, "Go-Around Procedure", 0,1,1,0 )""", [now, now,now,now])
    cexc.execute("""INSERT INTO RwyInfo (Code,ID,TORA,TODA,ASDA,XLDA,Width,Surface,SlopeTOD,SlopeASD,SlopeLDA,CrDate,UpdDate,CrTime,UpTime,eoLabel,lineupAngle,
    elevStartTORA,elevEndTORA,MagHdg,gaLabel,mfrh,useForTakeoff,useForLanding,mfrhType)
    VALUES("VOBM", "26", 2300, 2300, 2300, 2300, 45, 0, 0.2, 0.2,0.2, ?,?,?,?, "Engine Failure Procedure", 180, 0, 0, 261, "Go-Around Procedure", 0,1,1,0 )""", [now, now,now,now])
    
    conn.commit()


    
    #### VOVZ #####
        
    cexc.execute("""INSERT INTO AptInfo (Code, IATA, Name, City, Country, Elevation,RwyMeasType, DistUnit, HtUnit,ObDistRef, ObHtRef, CrDate, UpdDate,MagVar)
    VALUES("VOVZ", "VTZ", "VISHAKHAPATNAM", "VISHAKHAPATNAM", "INDIA", 10, 1,0,1,1,1, ?,?, 0)""", [now, now])
    conn.commit()

    EOP10="After takeoff, maintain runway heading, contact ATC."
    EOP28="REFER Jeppesen VIZAG Chart 10-7"
    cexc.execute("""INSERT INTO EOProced (Code, ID, EOProc, MinV2, MaxV2, SpeedUnits)
    VALUES("VOVZ", "10", ?, 0, 0, "KTAS")""", [EOP10])
    cexc.execute("""INSERT INTO EOProced (Code, ID, EOProc, MinV2, MaxV2, SpeedUnits)
    VALUES("VOVZ", "28", ?, 0, 0, "KTAS")""", [EOP28])
    conn.commit()

    cexc.execute("""INSERT INTO GAProced (Code, ID, gaGradient, DecisionHt, DeltaHt)
    VALUES("VOVZ", "10", 2.5, 0, 0)""")
    cexc.execute("""INSERT INTO GAProced (Code, ID, gaGradient, DecisionHt, DeltaHt)
    VALUES("VOVZ", "28", 2.5, 0, 0)""")
    conn.commit()

    cexc.execute("""INSERT INTO IntersectInfo (Code,ID,Name,deltaFL,deltaRef,elevStartTORA,lineupAngle,slopeTOD,slopeASD)
    VALUES("VOVZ", "10", "10N4", 307,0, 0, 90, -0.06, -0.06)""")
    cexc.execute("""INSERT INTO IntersectInfo (Code,ID,Name,deltaFL,deltaRef,elevStartTORA,lineupAngle,slopeTOD,slopeASD)
    VALUES("VOVZ", "10", "10N3", 612,0, 0, 90, -0.06, -0.06)""")
   
    conn.commit()
        
    dist28=(13458, 14019, 14562, 16540, 16597, 18276, 24484, 24774, 27447, 33826)
    ht28=(169, 176, 179, 209, 274, 366, 406, 412, 432, 616)
    for cntex6 in range(len(dist28)):
        random_int66=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('VOVZ','28',"",?,?,?,0)""",[random_int66, dist28[cntex6]*.3048, ht28[cntex6]])
        conn.commit()

    cexc.execute("""INSERT INTO RwyInfo (Code,ID,TORA,TODA,ASDA,XLDA,Width,Surface,SlopeTOD,SlopeASD,SlopeLDA,CrDate,UpdDate,CrTime,UpTime,eoLabel,lineupAngle,
    elevStartTORA,elevEndTORA,MagHdg,gaLabel,mfrh,useForTakeoff,useForLanding,mfrhType)
    VALUES("VOVZ", "10", 3050, 3050, 3050, 0, 45, 0, -0.06, -0.06,-0.06, ?,?,?,?, "Engine Failure Procedure", 90, 0, 0, 100, "Go-Around Procedure", 0,1,0,0 )""", [now, now,now,now])
    cexc.execute("""INSERT INTO RwyInfo (Code,ID,TORA,TODA,ASDA,XLDA,Width,Surface,SlopeTOD,SlopeASD,SlopeLDA,CrDate,UpdDate,CrTime,UpTime,eoLabel,lineupAngle,
    elevStartTORA,elevEndTORA,MagHdg,gaLabel,mfrh,useForTakeoff,useForLanding,mfrhType)
    VALUES("VOVZ", "28", 3050, 3050, 3050, 3050, 45, 0, 0.06, 0.06,0.06, ?,?,?,?, "Engine Failure Procedure", 90, 0, 0, 280, "Go-Around Procedure", 0,1,1,0 )""", [now, now,now,now])
    
    conn.commit()

    
     #### VEBD #####
    

    
    cexc.execute("""INSERT INTO AptInfo (Code, IATA, Name, City, Country, Elevation,RwyMeasType, DistUnit, HtUnit,ObDistRef, ObHtRef, CrDate, UpdDate,MagVar)
    VALUES("VEBD", "IXB", "BAGDOGRA", "BAGDOGRA", "INDIA", 414, 1,0,1,1,1, ?,?, 0)""", [now, now])
    conn.commit()

    EOP18="H At D3.0 BGD RIGHT turn to 200."
    EOP36="H At D3.0 BGD LEFT turn to 200."
    cexc.execute("""INSERT INTO EOProced (Code, ID, EOProc, MinV2, MaxV2, SpeedUnits)
    VALUES("VEBD", "18", ?, 0, 0, "KTAS")""", [EOP18])
    cexc.execute("""INSERT INTO EOProced (Code, ID, EOProc, MinV2, MaxV2, SpeedUnits)
    VALUES("VEBD", "36", ?, 0, 0, "KTAS")""", [EOP36])
    conn.commit()

    cexc.execute("""INSERT INTO GAProced (Code, ID, gaGradient, DecisionHt, DeltaHt)
    VALUES("VEBD", "18", 2.5, 0, 0)""")
    cexc.execute("""INSERT INTO GAProced (Code, ID, gaGradient, DecisionHt, DeltaHt)
    VALUES("VEBD", "36", 2.5, 0, 0)""")
    conn.commit()

    
    dist36=(29387,30123,30263,30552,33001,33030,33136,34021,34464,36976,41473,41904)
    ht36=(240,252,266,270,288,292,299,311,317,368,428,436)
    for cntex6 in range(len(dist36)):
        random_int66=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('VEBD','36',"",?,?,?,0)""",[random_int66, dist36[cntex6]*.3048, ht36[cntex6]])
        conn.commit()

    cexc.execute("""INSERT INTO RwyInfo (Code,ID,TORA,TODA,ASDA,XLDA,Width,Surface,SlopeTOD,SlopeASD,SlopeLDA,CrDate,UpdDate,CrTime,UpTime,eoLabel,lineupAngle,
    elevStartTORA,elevEndTORA,MagHdg,gaLabel,mfrh,useForTakeoff,useForLanding,mfrhType)
    VALUES("VEBD", "18", 2743, 3116, 3012, 2743, 45, 0, -0.37, -0.37,-0.37, ?,?,?,?, "Engine Failure Procedure", 90, 0, 0, 182, "Go-Around Procedure", 0,1,1,0 )""", [now, now,now,now])
    cexc.execute("""INSERT INTO RwyInfo (Code,ID,TORA,TODA,ASDA,XLDA,Width,Surface,SlopeTOD,SlopeASD,SlopeLDA,CrDate,UpdDate,CrTime,UpTime,eoLabel,lineupAngle,
    elevStartTORA,elevEndTORA,MagHdg,gaLabel,mfrh,useForTakeoff,useForLanding,mfrhType)
    VALUES("VEBD", "36", 2743, 3087, 3025, 2743, 45, 0, 0.37, 0.37,0.37, ?,?,?,?, "Engine Failure Procedure", 90, 0, 0, 2, "Go-Around Procedure", 0,1,1,0 )""", [now, now,now,now])
    
    conn.commit()


    #### VEJH #####
    
    
    cexc.execute("""INSERT INTO AptInfo (Code, IATA, Name, City, Country, Elevation,RwyMeasType, DistUnit, HtUnit,ObDistRef, ObHtRef, CrDate, UpdDate,MagVar)
    VALUES("VEJH", "JRG", "VEER SURENDRA SAI", "JHARSUGUDA", "INDIA", 757, 1,0,1,1,1, ?,?, 0)""", [now, now])
    conn.commit()

    EOP06="After takeoff, maintain runway heading, contact ATC."
    EOP24="After takeoff, maintain runway heading, contact ATC."
    cexc.execute("""INSERT INTO EOProced (Code, ID, EOProc, MinV2, MaxV2, SpeedUnits)
    VALUES("VEJH", "06", ?, 0, 0, "KTAS")""", [EOP06])
    cexc.execute("""INSERT INTO EOProced (Code, ID, EOProc, MinV2, MaxV2, SpeedUnits)
    VALUES("VEJH", "24", ?, 0, 0, "KTAS")""", [EOP24])
    conn.commit()

    cexc.execute("""INSERT INTO GAProced (Code, ID, gaGradient, DecisionHt, DeltaHt)
    VALUES("VEJH", "06", 2.5, 0, 0)""")
    cexc.execute("""INSERT INTO GAProced (Code, ID, gaGradient, DecisionHt, DeltaHt)
    VALUES("VEJH", "24", 2.5, 0, 0)""")
    conn.commit()

    dist08=(1271,1321,1635,1662,2075,12323,155506, 155883, 156787, 164226, 164302)
    ht08=(16,33,39,47,70,209,422,454,484,575,702,722)
    for cntex6 in range(len(dist08)):
        random_int77=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('VEJH','06',"",?,?,?,0)""",[random_int77, dist08[cntex6]*.3048, ht08[cntex6]])
        conn.commit()
    
    dist26=(1055,1066,3807)
    ht26=(7, 145, 257, 547)
    for cntex6 in range(len(dist26)):
        random_int66=randomgen()
        cexc.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
        VALUES('VEJH','24',"",?,?,?,0)""",[random_int66, dist26[cntex6]*.3048, ht26[cntex6]])
        conn.commit()

    cexc.execute("""INSERT INTO RwyInfo (Code,ID,TORA,TODA,ASDA,XLDA,Width,Surface,SlopeTOD,SlopeASD,SlopeLDA,CrDate,UpdDate,CrTime,UpTime,eoLabel,lineupAngle,
    elevStartTORA,elevEndTORA,MagHdg,gaLabel,mfrh,useForTakeoff,useForLanding,mfrhType)
    VALUES("VEJH", "06", 2390, 2390, 2390, 2390, 45, 0, 0.05, 0.05,0.05, ?,?,?,?, "Engine Failure Procedure", 180, 0, 0, 63, "Go-Around Procedure", 0,1,1,0 )""", [now, now,now,now])
    cexc.execute("""INSERT INTO RwyInfo (Code,ID,TORA,TODA,ASDA,XLDA,Width,Surface,SlopeTOD,SlopeASD,SlopeLDA,CrDate,UpdDate,CrTime,UpTime,eoLabel,lineupAngle,
    elevStartTORA,elevEndTORA,MagHdg,gaLabel,mfrh,useForTakeoff,useForLanding,mfrhType)
    VALUES("VEJH", "24", 2390, 2390, 2390, 2390, 45, 0, -0.05, -0.05,-0.05, ?,?,?,?, "Engine Failure Procedure", 180, 0, 0, 243, "Go-Around Procedure", 0,1,1,0 )""", [now, now,now,now])
    
    conn.commit()
########################################################################################################################
#################################################   SQL DEFINITIONS   ##################################################



Aerodromes=pd.read_excel(r"\\10.1.0.51\Safety DB\OneDrive - TURISTIK HAVA TASIMACILIK A.S\Statistics&Analysis\SERA İstatistik\SMS Database\SERA_REPORTS/Aerodromes.xlsx")               

def AptInfo():
    c2=conn.cursor()
    c2.execute("DELETE FROM AptInfo")
    conn.commit()
    for cnt1 in range(len(airport.index)):
        
        AptInfo_Code=airport.iat[cnt1, 0].replace('"','')
        AptInfo_IATA=airport.iat[cnt1, 1].replace('"','')
        AptInfo_Elevation=int(airport.iat[cnt1, 2])
        AptInfo_Airport=airport.iat[cnt1, 3].replace('"','')
        AptInfo_City=airport.iat[cnt1, 4].replace('"','')
        AptInfo_Country=airport.iat[cnt1, 5].replace('"','')
        dummyAerodromes=Aerodromes[Aerodromes["ICAO Code"]==AptInfo_Code]
        dummyAerodromes = dummyAerodromes.fillna("NIL")

        if len(dummyAerodromes)==0:
            Apt_comment=""
        else:
            
            Apt_comment="""

            SERA Aerodrome Comments (Review Date: """ + today + """ )  


            Category: """ + str(dummyAerodromes.iat[0,6]) + """
                       
            ___Flight OPS Comment___: 
            """ +str(dummyAerodromes.iat[0,12]) + """

            
            ___Service Info & Ground Ops___: 
            """ + str(dummyAerodromes.iat[0,13]) + """ 

            
            ___Ops Control___: 
            """ + str(dummyAerodromes.iat[0,14]) + """\n


            ___Security___:  
            """ +str(dummyAerodromes.iat[0,16]) + """


            ___Commercial___:  
            """ +str(dummyAerodromes.iat[0,17]) +  """


            ___SMS/FDM___:   
            """+str(dummyAerodromes.iat[0,18])  + """

            
            ___Training___: 
            """+str(dummyAerodromes.iat[0,19]) + """


            ___Insurance___:  
            """+str(dummyAerodromes.iat[0,20])
        
        


        c1=conn.cursor()
        c1.execute("""INSERT INTO AptInfo (Code, IATA, Name, City, Country, Elevation, RwyMeasType, DistUnit, HtUnit, ObDistRef, ObHtRef, CrDate, UpdDate, Comment, MagVar, Lat, Long)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [AptInfo_Code, AptInfo_IATA, AptInfo_Airport, AptInfo_City, AptInfo_Country, AptInfo_Elevation, "1","0", "1","1", "1", now, now, Apt_comment, "0", "" , ""])

    conn.commit()

def DBconst():
    now=datetime.strftime(datetime.now(), "%Y-%m-%d %H:%M:%S")
    
    c3=conn.cursor()
    c3.execute("UPDATE databaseConstants SET airportversionID=?, assembleDate=?, StartDate=?, EndDate=?, StartTime=?, EndTime=?", [todayasvers, now, now, now, now, now])
    conn.commit()

def EOproc():

    c31=conn.cursor()
    c31.execute("DELETE FROM EOProced")
    conn.commit()
    for cnt2 in range(len(eoproc.index)):
        try:
            EOproc_Code=eoproc.iat[cnt2, 0].replace('"','')
            EOproc_str=eoproc.iat[cnt2, 2]
            if EOproc_str.startswith('STD.'):
                EOproc_str='After takeoff, maintain runway heading, contact ATC.'
            EOproc_ID=eoproc.iat[cnt2, 1]
            c35=conn.cursor()
            c35.execute("SELECT Width FROM RwyInfo WHERE Code=? AND ID=?", [EOproc_Code, EOproc_ID])
            ftc35=c35.fetchone()

            if int(ftc35[0])<45:
                EOproc_EOproc="This Runway is narrow ("+ str(ftc35[0])+ "m). "+ EOproc_str
            else:
                EOproc_EOproc=EOproc_str
            
            c4=conn.cursor()
            c4.execute("""INSERT INTO EOProced (Code, ID, ProcedureID, EOProc, AcType, FlapConfig, MinV2, MaxV2, SpeedUnits)
                    VALUES(?,?,?,?,?,?,?,?,?)""",[EOproc_Code, EOproc_ID, "", EOproc_EOproc, "","", "0", "0", "KTAS"])
        except:
            pass
    conn.commit()


    

    
def GAproc():
    c41=conn.cursor()
    c41.execute("DELETE FROM GAProced")
    conn.commit()
    
    for cnt3 in range(len(eoproc.index)):
        EOproc_Code=eoproc.iat[cnt3, 0].replace('"','')
        EOproc_ID=eoproc.iat[cnt3, 1]
        
        
        subrunwayGA=runway[(runway.iloc[:,0]==eoproc.iat[cnt3, 0]) & (runway.iloc[:,2]== eoproc.iat[cnt3, 1])]
        
        c5=conn.cursor()
        c5.execute("""INSERT INTO GAProced (Code, ID, ProcedureID, gaProc, gaGradient, DecisionHt, DeltaHt, AcType, FlapConfig)
                VALUES(?,?,?,?,?,?,?,?,?)""",[EOproc_Code, EOproc_ID, "", "", subrunwayGA.iat[0,12],"0", "0", "", ""])
    conn.commit()

def IntInfo():
    c51=conn.cursor()
    c51.execute("DELETE FROM IntersectInfo")
    conn.commit()
    
    for cnt4 in range(len(inter.index)):
        try:
            Int_Code=inter.iat[cnt4,0].replace('"', '')
            
            Int_ID=inter.iat[cnt4,2].replace('"', '')
            Int_Name=inter.iat[cnt4, 3]
            Int_deltaFL=int(inter.iat[cnt4,17])
            Int_Lineup=int(inter.iat[cnt4,15])
            Int_slope=float(inter.iat[cnt4,8])

            c6=conn.cursor()
            c6.execute("""INSERT INTO IntersectInfo (Code, ID, Name, deltaFL, deltaRef, elevStartTORA,LatStartTORA, LongStartTORA, lineupAngle,  slopeTOD, slopeASD, Comment)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",[Int_Code, Int_ID, Int_Name,Int_deltaFL, "0", "0", "","", Int_Lineup,Int_slope,Int_slope, ""])
           
        
        except Exception as Eint:
            try:
                Int_deltaFL=int(inter.iat[cnt4,16])
                Int_Lineup=int(inter.iat[cnt4,14])
                Int_slope=float(inter.iat[cnt4,7])
                c6=conn.cursor()
                c6.execute("""INSERT INTO IntersectInfo (Code, ID, Name, deltaFL, deltaRef, elevStartTORA,LatStartTORA, LongStartTORA, lineupAngle,  slopeTOD, slopeASD, Comment)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",[Int_Code, Int_ID, Int_Name,Int_deltaFL, "0", "0", "","", Int_Lineup,Int_slope,Int_slope, ""])
            except:
                Int_deltaFL=0
                Int_Lineup=0
                Int_slope=0
                c6=conn.cursor()
                c6.execute("""INSERT INTO IntersectInfo (Code, ID, Name, deltaFL, deltaRef, elevStartTORA,LatStartTORA, LongStartTORA, lineupAngle,  slopeTOD, slopeASD, Comment)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",[Int_Code, Int_ID, Int_Name,Int_deltaFL, "0", "0", "","", Int_Lineup,Int_slope,Int_slope, ""])
                print(Int_Code, Int_ID,"INT HAS ERRORS")
    conn.commit()
    
def NotamInfo():
    c61=conn.cursor()
    c61.execute("DELETE FROM NOTAMinfo")
    conn.commit()







def ObsInfo():

    c71=conn.cursor()
    c71.execute("DELETE FROM ObstInfo")
    conn.commit()
    
    obsunique = obs.drop_duplicates([0,1])
    
    obsunique.reset_index()
    
    for cnt6 in range(len(obsunique.index)):
        namepivot=obsunique.iat[cnt6, 0].replace('"', '')
        IDpivot=obsunique.iat[cnt6, 1]
        dfobs2wrt=obs[(obs[0]=='"'+str(namepivot)+'"') & (obs[1]==IDpivot)]
        
        
        
        c8=conn.cursor()
        c8.execute("SELECT * from ObstInfo WHERE Code=? AND ID=?", [namepivot, IDpivot])
        ftc8=c8.fetchone()

        if ftc8:
            
            c8.execute("DELETE from ObstInfo WHERE Code=? AND ID=?", [namepivot, IDpivot])
            for cnt7 in range(len(dfobs2wrt.index)):
                checktrue=True
                while checktrue==True:
                    rand_int=random.randint(150000000, 300000000)
                    c75=conn.cursor()
                    c75.execute("SELECT * from ObstInfo WHERE Number=?", [rand_int])
                    ftc75=c75.fetchone()
                    if ftc75:
                        checktrue=True
                        
                    else:
                        checktrue=False
                        
                Obs_Code=dfobs2wrt.iat[cnt7, 0].replace('"', '')
                Obs_ID=dfobs2wrt.iat[cnt7, 1]
                Obs_dist=int(dfobs2wrt.iat[cnt7, 2])
                Obs_ht=int(dfobs2wrt.iat[cnt7, 3])
                c8.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID ,Number, Dist, Ht, LatOffset)
                VALUES(?,?,?,?,?,?,?)""",[Obs_Code, Obs_ID, "" ,rand_int ,Obs_dist, Obs_ht, int(0)])
        else:
            for cnt7 in range(len(dfobs2wrt.index)):
                checktrue=True
                while checktrue==True:
                    rand_int=random.randint(150000000, 300000000)
                    c75=conn.cursor()
                    c75.execute("SELECT * from ObstInfo WHERE Number=?", [rand_int])
                    ftc75=c75.fetchone()
                    if ftc75:
                        checktrue=True
                    else:
                        checktrue=False
                Obs_Code=dfobs2wrt.iat[cnt7, 0].replace('"', '')
                Obs_ID=dfobs2wrt.iat[cnt7, 1]
                Obs_dist=int(dfobs2wrt.iat[cnt7, 2])
                Obs_ht=int(dfobs2wrt.iat[cnt7, 3]) 
                c8.execute("""INSERT INTO ObstInfo (Code, ID, ProcedureID, Number, Dist, Ht, LatOffset)
                VALUES(?,?,?,?,?,?,?)""",[Obs_Code, Obs_ID, "", rand_int,Obs_dist, Obs_ht, "0"])
               

    conn.commit()

def PCN(pcnrunway):
    try:
        PCNasSTR=""
        tolGW800=""
        tolGWMK=""
   
        GW800=""
        GWMK=""
        
        dummyPCN=pcnrunway.replace("/","")
        dummyPCN=dummyPCN.replace(" ","")
        dummyPCN=dummyPCN.replace("PCN","")
     
        if dummyPCN[-1]=="T" or dummyPCN[-1]=="U":
            PCNno=int(''.join(filter(str.isdigit, dummyPCN)))
            restdummyPCN=dummyPCN.replace(str(PCNno), "")
            pcnlist=[str(PCNno), restdummyPCN[0], restdummyPCN[1], restdummyPCN[2], restdummyPCN[3]]
            PCNasSTR='/'.join(pcnlist)
            

            if restdummyPCN[0]=="F":
                if restdummyPCN[1]=="A":
                    GW800=79015-(43-PCNno)*36600/(43-20)
                    GWMK=82190-(45-PCNno)*36750/(45-24)
                elif restdummyPCN[1]=="B":
                    GW800=79015-(45-PCNno)*36600/(45-21)
                    GWMK=82190-(48-PCNno)*36750/(48-25)
                elif restdummyPCN[1]=="C":
                    GW800=79015-(50-PCNno)*36600/(50-22)
                    GWMK=82190-(53-PCNno)*36750/(53-27)
                elif restdummyPCN[1]=="D":
                    GW800=79015-(55-PCNno)*36600/(55-26)
                    GWMK=82190-(58-PCNno)*36750/(58-31)
                else:
                    GW800=""
                    GWMK=""
                
                if GW800*1.1>79015:
                    tolGW800=str(79015)
                else:
                    tolGW800=str(int(GW800*1.1))
            
                if GWMK*1.1>82190:
                    tolGWMK=str(82190)
                else:
                    tolGWMK=str(int(GWMK*1.1))

                
                                 
            elif restdummyPCN[0]=="R":
                if restdummyPCN[1]=="A":
                    GW800=79015-(49-PCNno)*36600/(49-23)
                    GWMK=82190-(52-PCNno)*36750/(52-27)
                elif restdummyPCN[1]=="B":
                    GW800=79015-(52-PCNno)*36600/(52-24)
                    GWMK=82190-(54-PCNno)*36750/(54-28)
                elif restdummyPCN[1]=="C":
                    GW800=79015-(54-PCNno)*36600/(54-25)
                    GWMK=82190-(57-PCNno)*36750/(57-30)
                elif restdummyPCN[1]=="D":
                    GW800=79015-(56-PCNno)*36600/(56-27)
                    GWMK=82190-(59-PCNno)*36750/(59-31)
                else:
                    GW800=""
                    GWMK=""
               
                if GW800*1.05>79015:
                    tolGW800=str(79015)
                else:
                    tolGW800=str(int(GW800*1.05))
            
                if GWMK*1.05>82190:
                    tolGWMK=str(82190)
                else:
                    tolGWMK=str(int(GWMK*1.05))
    
              
    except Exception as Etol:
        print("pcn", Etol)
        PCNasSTR=""
        tolGW800=""
        tolGWMK=""
            

    return PCNasSTR, tolGW800, tolGWMK
        

def RWYinfo():

    pcndf=pd.DataFrame([])
    

    c91=conn.cursor()
    
   
    c91.execute("DELETE FROM RwyInfo")
    conn.commit()
        
    outlook = win32com.client.Dispatch('outlook.application')
    pcndf=pd.read_csv("silmePCN/pcn.csv")
    
    for cnt8 in range(len(runway.index)):
        try:
            runway_Code=runway.iat[cnt8,0].replace('"', '')
            runway_ID=runway.iat[cnt8,2]
            runway_head=int(runway.iat[cnt8,3])
            runway_TORA=int(runway.iat[cnt8,5])
            runway_LRA=int(runway.iat[cnt8,6])
            runway_slope=float(runway.iat[cnt8,7])
            runway_cwy=int(runway.iat[cnt8,8])
            runway_stpwy=int(runway.iat[cnt8,9])
            runway_aal=int(runway.iat[cnt8,11])
            runway_groove=int(runway.iat[cnt8,13])
            runway_lineup=int(runway.iat[cnt8,14])
            runway_width=int(runway.iat[cnt8,15].replace('.0\n',''))
            runway_PCN=runway.iat[cnt8, 16]
            PCNvalues=PCN(runway_PCN)


            
            try:
                strPCN=str(PCNvalues[0])
                PCN800n=str(PCNvalues[1])
                PCN8n=str(PCNvalues[2])
            except:
                strPCN=""
                PCN800n=""
                PCN8n=""

            
            TireLimit=""
            if len(strPCN)>=4:
                if strPCN[-3]=="Y" or strPCN[-3]=="Z":
                    TireLimit="OPERATION IS FORBIDDEN DUE TO MAX TIRE PRESSURE!"
            
            
            try:
                PCNstring="""\n
                __PCN Information__\n
                \n
                For 737-800:\n
                """+ PCNvalues[0] + " = " + PCNvalues[1] + " KG\n"+ """
                \n
                \n
                For 737-MAX8:\n
                """+ PCNvalues[0] + " = " + PCNvalues[2] + " KG\n" + """
                \n
                \n
                """ + TireLimit + "\n" +"""
                \n
                PCN INFORMATION GIVEN ABOVE IS INTENDED TO BE USED BY DISPATCHER.\n
                FLIGHT CREW SHOULD NOT TAKE INTO ACCOUNT THE GIVEN INFORMATION DURING PERFORMANCE CALCULATIONS.\n
                FOR PCN ISSUES, FLIGHT CREW MAY ADVISE THE OCC.\n""" 
                
            except:
                PCNstring=""
                

         
            

            

            
            subpcndf=pcndf[(pcndf.iloc[:,0]==runway_Code) & (pcndf.iloc[:,1]==runway_ID)]
            
            if len(subpcndf)==1:
                

                if pd.isna(subpcndf.iat[0,2]):
                    
                    subpcndf.iat[0,2]=""

                if not subpcndf.iat[0,2]==strPCN:
                    
                    
                    mail = outlook.CreateItem(0)
                    mail.To = 'odemirhan@corendon-airlines.com'#;  navigation@corendon-airlines.com; occ@corendon-airlines.com'
                    mail.Subject = 'Attention! PCN Information'
                    try:
                        kg800=str(int(subpcndf.iat[0,3]))
                        kg8=str(int(subpcndf.iat[0,4]))
                    except:
                        kg800=''
                        kg8=''
                    mail.Body = """
                    PCN Change Info:


                    Airport: """+ subpcndf.iat[0,0]+"""
                    Runway: """ + subpcndf.iat[0,1]+"""



                    ------------------------------------------------------------
                    __Old PCN__: """ + """

                    For 737-800:
                    """+ str(subpcndf.iat[0,2]) + " = " + kg800 + " KG"+ """
            
                    For 737-MAX8:
                    """+ str(subpcndf.iat[0,2]) + " = " + kg8 + " KG" + """
                    
                    

                    __New PCN__: """ + """

                    For 737-800:
                    """+ strPCN + " = " + PCN800n + " KG"+ """
            
                    For 737-MAX8:
                    """+ strPCN + " = " + PCN8n + " KG" +"""
                    ------------------------------------------------------------

                    PLEASE SEE THE PAST ANALYSES FOR THIS AIRPORT!

                    
                    mail: foe@corendon-airlines.com
                    
                    """
                    mail.Send()
                    
                    pcndf.loc[(pcndf.iloc[:,0]==runway_Code) & (pcndf.iloc[:,1]==runway_ID), "2"]=strPCN
                    pcndf.loc[(pcndf.iloc[:,0]==runway_Code) & (pcndf.iloc[:,1]==runway_ID), "3"]=PCN800n
                    pcndf.loc[(pcndf.iloc[:,0]==runway_Code) & (pcndf.iloc[:,1]==runway_ID), "4"]=PCN8n

                    
                
                    
                    

            elif len(subpcndf)==0:
                
                PCNappendline=pd.DataFrame([[runway_Code, runway_ID, strPCN, PCN800n, PCN8n]], columns=['0', '1', '2', '3', '4'])
                pcndf=pd.concat([pcndf, PCNappendline])
                
                
            
         
            
            if runway_TORA==0:
                useforTO=0
            else:
                useforTO=1

            
            if runway_LRA==0:
                useforLD=0
            else:
                useforLD=1    
            
            filterednotam=notam[(notam[0]==runway_Code) & (notam[2]==runway_ID)]
            notamComment="""\n
            __NOTAM Information__\n
            \n
            """
            for cnt9 in range(len(filterednotam)):
                notamComment=notamComment + str(filterednotam.iat[cnt9, 5]) + "; "
            if notamComment=="""\n
            __NOTAM Information__\n
            \n
            """:
                notamComment="""\n
                __NOTAM Information__\n
                \n
                NIL\n
                """
            Comment= notamComment +"\n"+"""
            \n
            \n
            """ + PCNstring
            
            c9=conn.cursor()
            c9.execute("""INSERT INTO RwyInfo (Code, ID, TORA, TODA, ASDA, XLDA, Width, Surface,
            SlopeTOD, SlopeASD, SlopeLDA, CrDate, UpdDate, CrTime, UpTime, Comment, eoLabel, lineupAngle,
            elevStartTORA, elevEndTORA, MagHdg, gaLabel, mfrh,  useForTakeoff, useForLanding, mfrhType)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", [runway_Code, runway_ID, runway_TORA, runway_cwy, runway_stpwy, runway_LRA, runway_width, runway_groove*2 ,
                                                                                  runway_slope ,runway_slope, runway_slope,now , now , now, now, Comment, "Engine Failure Procedure", runway_lineup,
                                                                                  "0", "0", runway_head, "Go-Around Procedure", "0",  useforTO, useforLD, "0"])
        except:
            print(runway_Code, runway_ID)
    conn.commit()

    for cntpcn01 in range(len(pcndf)):
        try:
            pcndf.iat[cntpcn01,3]=str(int(pcndf.iat[cntpcn01,3]))
            pcndf.iat[cntpcn01,4]=str(int(pcndf.iat[cntpcn01,4]))
        except:
            pcndf.iloc[cntpcn01,3]=""
            pcndf.iloc[cntpcn01,4]=""
    pcndf=pcndf.astype('str')
    pcndf.to_csv("silmePCN/pcn.csv", index=False)



########################################################################################################################
#################################################   GENERAL DEFINITIONS   ########################################################    



def choosepackage():
    
    global airportfilepath
    airportfilepath=(os.path.join(odpath,  "airport", today))
    if os.path.exists(airportfilepath):
        shutil.rmtree(airportfilepath)
    os.mkdir(airportfilepath)
            
        
    
    chsn_file = filedialog.askopenfilename()
    shutil.move(chsn_file, os.path.join(airportfilepath , "package.zip"))
    patoolib.extract_archive(os.path.join(airportfilepath , "package.zip"), outdir = airportfilepath)



def outlookdown():
    global airportfilepath
    cnt6=0
    while cnt6<5:
        try:
            
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application").GetNamespace("MAPI")
          
            #root_folder = outlook.Folders.Item(1)
            
            sfolder = outlook.GetDefaultFolder(6).Folders.Item("Interim Perf")
         
            messages = sfolder.Items
            messages.Sort("[ReceivedTime]", False)
            message = messages.GetLast()
            airportfilepath=(os.path.join(odpath,  "airport", today))
            
            if os.path.exists(airportfilepath):
                shutil.rmtree(airportfilepath)
            os.mkdir(airportfilepath)
            
            
            
            attachments = message.Attachments
            num_attach = len([cor for cor in attachments])
            
            for cor in range(1, num_attach+1):
                attachment = attachments.Item(cor)
                attachment.SaveAsFile(airportfilepath +'\\'+ attachment.FileName)
           
            fname=attachment.FileName
            
            patoolib.extract_archive(os.path.join(airportfilepath , fname), outdir = airportfilepath)
            break
        
        except Exception as E:
            os.system("taskkill /f /im outlook.exe")
            print(E)
            cnt6+=1
                        

def ParseTxt():

    global runway
    global airport
    
    global inter
    global eoproc
    global obs
    navtechlist=glob.glob(airportfilepath + "/*.txt") #testten sonra düzetl
    
    for cnt7 in range(len(navtechlist)):
        if not "Change" in navtechlist:
            navtechpath=navtechlist[cnt7]
            
    #print(navtechpath)
    #navtechpath=os.path.join(odpath, "airport", "transfer", "navtech.txt") #bu kýrmýzý olacak
    navtech=open(navtechpath, "r")
    ntlines=navtech.readlines()
    a=0
    airport150=pd.DataFrame([])
    runway200=pd.DataFrame([])
    inter200=pd.DataFrame([])
    eoproc216=pd.DataFrame([])
    obs220=pd.DataFrame([])
    
    for i in range(len(ntlines)): #200koy kýsa deneme icin
        ntline=ntlines[i]
        ntliner=ntline.replace(str(i)+" ","")
        
        if ntliner[0:3]=="150":
            ntliner=ntliner.split(' ')
            ntliner = list(filter(None, ntliner))
            #print(len(ntliner))
            cnt1=len(ntliner)-6
            airportname=""
            for ncnt1 in range(cnt1):
                airportname=airportname+str(ntliner[ncnt1+4])+" " 
            aname=airportname.split('"', 2)
            reconliner=[ntliner[1],ntliner[2], ntliner[3], '"'+aname[1]+'"', ntliner[-2],ntliner[-1]]
                        
            arptline=(pd.DataFrame(reconliner)).transpose()
            
            airport150=airport150.append(arptline)
                
        
            k=i+1
    #print(airport150)
            
            while True:
                try: 
                    
                    ntlinea=ntlines[k].replace(str(k)+" ", "")
                    
                    
                    if ntlinea[0:3]=="200":
                        ntlinea=ntlinea.split('"')
                        PCN=str(ntlinea[3])
                        ntlinea=str(ntlinea[0])+str(ntlinea[1])+str(ntlinea[2])+str(ntlinea[4])
                        ntlinea=ntlinea.split(' ')
                        ntlinea = list(filter(None, ntlinea))
                        
                        ntlinea.insert(0, arptline.at[0,0])
                        
                        l=k+1
                        while True:
                            try:
                                
                                #print(ntlines)
                                ntlineb=ntlines[l].replace(str(l)+" ", "")
                                
                                
                                if ntlineb[0:3]=="201":
                                    ntlineb=ntlineb.replace('"',"")
                                    ntlineb=ntlineb.split(' ')
                                    ntlineb = list(filter(None, ntlineb))
                                    aligntype=ntlineb[1]
                                    l+=1
                                elif ntlineb[0:3]=="202":
                                    ntlineb=ntlineb.replace('"',"")
                                    ntlineb=ntlineb.split(' ')
                                    ntlineb = list(filter(None, ntlineb))
                                    rwywidth=ntlineb[1]
                                    
                                    l+=1 #rwy width
                                elif ntlineb[0:3]=="203":
                                    ntlineb=ntlineb.replace('"',"")
                                    ntlineb=ntlineb.split(' ')
                                    ntlineb = list(filter(None, ntlineb))
                                    TOshift=ntlineb[2]
                                    if ntlineb[1]=="1":

                                        ntlinea.insert(-1,  rwywidth)
                                        ntlinea.insert(-2, aligntype)
                                        ntlinea.insert(-1, PCN)
                                        
                                        rwyline=(pd.DataFrame(ntlinea)).transpose()
                                        runway200=runway200.append(rwyline)
                                        cnt3=1+l
                                        while True:
                                            try: 
                                                ntlinec=ntlines[cnt3].replace(str(cnt3)+" ", "")
                                                if ntlinec[0:3]=="216":                              
                                                    ntlinec=ntlinec.split('"')
                                                    ntlinec = list(filter(None, ntlinec))
                                                    EOproc=ntlinec[1]
                                                    EOlist=[ntlinea[0],ntlinea[2],EOproc]
                                                    EOline=(pd.DataFrame(EOlist)).transpose()
                                                    eoproc216=eoproc216.append(EOline)
                                                    cnt3+=1
                                                elif ntlinec[0:3]=="220":
                                                    ntlinec=ntlinec.replace('"',"")
                                                    ntlinec=ntlinec.split(' ')
                                                    ntlinec = list(filter(None, ntlinec))
                                                    obslist=[ntlinea[0],ntlinea[2], ntlinec[2], ntlinec[1]] 
                                                    obsline=(pd.DataFrame(obslist)).transpose()
                                                    obs220=obs220.append(obsline)
                                                    cnt3 +=1
                                                    
                                                elif ntlinec[0:3]=="200":
                                                    break
                                                else:
                                                    cnt3 +=1
                                            except IndexError:
                                                break
                                                
                                                                                 
                                            
                                    else:
                                        ttt=0
                                        while ttt<50:
                                            ntlineb=ntlines[l-ttt].replace(str(l-ttt)+" ", "")
                                            if ntlineb[0:3]=="203":
                                                ntlineb=ntlineb.split(' ')
                                                ntlineb = list(filter(None, ntlineb))
                                                if ntlineb[1]=="1":
                                                    rwylineforint=ntlines[l-ttt-3].replace(str(l-ttt-3)+" ", "")
                                                    rwylistforint=rwylineforint.split(' ')
                                                    rwylistforint = list(filter(None, rwylistforint))
                                                    rwyname=rwylistforint[1]
                                                    break
                                                else:
                                                    ttt+=1
                                            else:
                                                ttt+=1
                                    
                                        try:
                                            ntlinea.insert(2, rwyname)
                                        except:
                                            ntlinea.insert(2, "Null")

                                        ntlinea.insert(-1, TOshift)
                                        ntlinea.insert(-2, rwywidth)
                                        ntlinea.insert(-3, aligntype)
                                        
                                        
                                        intline=(pd.DataFrame(ntlinea)).transpose()
                                        inter200=inter200.append(intline)
                                    l+=1
                                
                                elif ntlineb[0:3]=="200":
                                    break
                                else:
                                    l+=1

                            except IndexError:
                                break
                        
                        k=k+1
                        
                    else:
                        if ntlinea[0:3]=="150":
                                                           
                            break
                        else:
                            k=k+1
                except IndexError:
                    break


    runway=runway200
    
    airport=airport150
    
    inter=inter200
    
    eoproc=eoproc216
   
    obs=obs220


def pypdf():
    global notam
    notamlist=glob.glob(airportfilepath + "/*.pdf")
    
    for cnt8 in range(len(notamlist)):
        if not "Change" in notamlist:
            notampath=notamlist[cnt8]
            
    
    rawText = parser.from_file(notampath)
    rawList = rawText['content'].splitlines()
    rawList = list(filter(None, rawList))
    rawList = list(filter(lambda k: 'Tempo' not in k, rawList))
    rawList = list(filter(lambda k: 'CORENDON' not in k, rawList))
    rawList = list(filter(lambda k: 'ICAO' not in k, rawList))
    rawList = list(filter(lambda k: 'Page' not in k, rawList))
 

    for i in range(len(rawList)):
        matchh=re.search('\d{4}-\d{2}-\d{2}', rawList[i])
        if matchh==None:
            if not len(rawList[i])==8:
                k=i-1
                while True:
                    try:
                        rawList[k]=rawList[k]+ rawList[i]
                        rawList[i]=None
                        break
                    except TypeError:
                        k=k-1
                
         
    rawList=list(filter(None, rawList))
    
    List=[]
    for l in range(len(rawList)):
        rawListA=rawList[l].split(";")
        List.extend(rawListA)

    
    for m in range(len(List)):
        matchhh=re.search('\d{4}-\d{2}-\d{2}', List[m])
        if matchhh==None:
            if not len(List[m])==8:
                n=m-1
                while True:
                    match0=re.search('\d{4}-\d{2}-\d{2}', List[n])
                    if not match0==None:
                        Item=" ".join(List[n].split(" ", 3)[:3])
                        List[m]=Item+List[m]
                        break
                    else:
                        n=n-1
    for kk in range(len(List)):
        if len(List[kk])==8:
            dummyarpt=List[kk]
            kk=kk+1
            try:
                while True:
                    if not len(List[kk])==8:
                        List[kk]=dummyarpt+" "+List[kk]
                        kk=kk+1
                    else:
                        break
            except IndexError:
                break

    ll=0
    while True:
        try:
            if len(List[ll])==8:
                del List[ll]
            else:
                List[ll]=List[ll].replace(" ",",",5)
                ll+=1
        except IndexError:
                break
    
    List=list(filter(None, List))
    notamDF=pd.DataFrame([])
    for cnt5 in range(len(List)):
        splittedline=List[cnt5].split(",")
        
        dfline=(pd.DataFrame(splittedline)).transpose()
        
        notamDF=notamDF.append(dfline)
   
    
        
    
    #obstacleDF=pd.DataFrame(List)
    #print(notamDF)
    #notamDF.to_excel(odpath+ 'airport\\transfer\csv\notams.xlsx', index = False, header=True)
    notam=notamDF


###########################################Choose Method###########################################
    ################################################################################



MsgBox=tk.messagebox.askquestion("Choose a Method to parse...", "Do you want to download the package from Outlook? 'No' will make you to choose the package yourself!") 

if MsgBox=='yes':
    outlookdown()
else:
    choosepackage()

ParseTxt()
pypdf()


AptInfo()
DBconst()
RWYinfo()
EOproc()
GAproc()
IntInfo()
NotamInfo()
ObsInfo()

Exceptions()

conn.close()

shutil.copy(os.path.join(odpath ,  "airport", "Current Airport DB",  today + "_airport.sdb"), os.path.join(odpath ,  "airport","Current Airport DB" , "airports_archieve", today + "_airport.sdb"))
shutil.copy(os.path.join(odpath ,  "airport", "Current Airport DB",  today + "_airport.sdb"), os.path.join(odpath ,  "airport", today, today + "_airport.sdb"))
shutil.move(os.path.join(odpath ,  "airport", "Current Airport DB",  today + "_airport.sdb"), os.path.join(odpath ,  "airport",  "Current Airport DB",   "airport.sdb"))


###### TANKERING CHECK ###########

conn2=sqlite3.connect(os.path.join(odpath ,  "airport",  "Current Airport DB",   "airport.sdb"))
tkairportDF=pd.read_sql("SELECT Code from AptInfo", conn2)
tkrunwayDF=pd.read_sql("SELECT Code, ID, TORA, TODA, ASDA, XLDA from RwyInfo", conn2)
tankeringDF=pd.read_excel("silmePCN/tankering.xlsx")
outlook = win32com.client.Dispatch('outlook.application')



for cntaprt in range(len(tkairportDF)):
    dummytkrunwayDF=tkrunwayDF[tkrunwayDF["Code"]==tkairportDF.iat[cntaprt,0]]
    if len(dummytkrunwayDF)<=2:
        dummytkrunwayDF["TORAPCLW"]=dummytkrunwayDF["TORA"]+dummytkrunwayDF["TODA"]
        dummytkrunwayDF["TORAPSTW"]=dummytkrunwayDF["TORA"]+dummytkrunwayDF["ASDA"]
        dummytkrunwayDF["TORAPCLW"]=dummytkrunwayDF[["TORAPCLW","TORAPSTW"]].min(axis=1)
        dummytkrunwayDF["LDRAPSPW"]=dummytkrunwayDF["XLDA"]
        dummytkrunwayDF["TORAPCLW"]=dummytkrunwayDF.loc[dummytkrunwayDF["TORAPCLW"] > 100, "TORAPCLW"]
        dummytkrunwayDF["LDRAPSPW"]=dummytkrunwayDF.loc[dummytkrunwayDF["LDRAPSPW"] > 100, "LDRAPSPW"]
        mindistanceTODA=dummytkrunwayDF["TORAPCLW"].min()
        mindistanceLDA=dummytkrunwayDF["LDRAPSPW"].min()

        if mindistanceTODA<2000 or mindistanceLDA<2000:
            dummytankeringDF=tankeringDF[tankeringDF["AIRPORT"]==tkairportDF.iat[cntaprt,0]]
            if len(dummytankeringDF)==0:
                tankeringDF=tankeringDF.append(pd.DataFrame(data={"AIRPORT":[tkairportDF.iat[cntaprt,0]],"Min of TODA":[mindistanceTODA], "Min of LDA":[mindistanceLDA]}))
                
                body1="TODA/LDA Info: \n" + str(tkairportDF.iat[cntaprt,0]) + " has been added to the restricted tankering database.\n" + "Min of TODA: " + str(mindistanceTODA) + "m \n"+   "Min of LDA: " + str(mindistanceLDA) +" m  \n"
                mail = outlook.CreateItem(0)
                mail.To = 'odemirhan@corendon-airlines.com'#;  navigation@corendon-airlines.com; occ@corendon-airlines.com'
                mail.Subject = 'Attention! TODA/LDA Information'
                mail.Body=body1
                mail.Send()    
                    
                
                
            else:
                if dummytankeringDF.iat[0,1]==mindistanceTODA:
                    if not dummytankeringDF.iat[0,2]==mindistanceLDA:
                        tankeringDF.loc[tankeringDF.AIRPORT == tkairportDF.iat[cntaprt,0] ,'Min of LDA'] = mindistanceLDA
                        body2="TODA/LDA Info: \n"+ str(tkairportDF.iat[cntaprt,0]) + " has been updated in the restricted tankering database. \n" +  "OLD Min of TODA and LDA: "+ str(dummytankeringDF.iat[0,1]) +  " m (TODA)   "+ str(dummytankeringDF.iat[0,2])+ " m (LDA)\n"+"NEW Min of TODA, LDA: " + str(mindistanceTODA) + " m (TODA)      "+ str(mindistanceLDA)+ " m (LDA)"
                        mail = outlook.CreateItem(0)
                        mail.To = 'odemirhan@corendon-airlines.com'#;  navigation@corendon-airlines.com; occ@corendon-airlines.com'
                        mail.Subject = 'Attention! TODA/LDA Information'
                        mail.Body=body2
                        mail.Send()
                else:
                    if not dummytankeringDF.iat[0,2]==mindistanceLDA:
                        tankeringDF.loc[tankeringDF.AIRPORT == tkairportDF.iat[cntaprt,0] ,'Min of LDA'] = mindistanceLDA
                    
                    tankeringDF.loc[tankeringDF.AIRPORT == tkairportDF.iat[cntaprt,0] ,'Min of TODA'] = mindistanceTODA    
                    body3="TODA/LDA Info: \n"+ str(tkairportDF.iat[cntaprt,0]) + " has been updated in the restricted tankering database. \n" +  "OLD Min of TODA and LDA: "+ str(dummytankeringDF.iat[0,1]) +  " m (TODA)   "+ str(dummytankeringDF.iat[0,2])+ " m (LDA)\n"+"NEW Min of TODA, LDA: " + str(mindistanceTODA) + " m (TODA)      "+ str(mindistanceLDA)+ " m (LDA)"
                    mail = outlook.CreateItem(0)
                    mail.To = 'odemirhan@corendon-airlines.com'#;  navigation@corendon-airlines.com; occ@corendon-airlines.com'
                    mail.Subject = 'Attention! TODA/LDA Information'
                    mail.Body=body3
                    mail.Send()    
        else:
            dummytankeringDF=tankeringDF[tankeringDF["AIRPORT"]==tkairportDF.iat[cntaprt,0]]
            if len(dummytankeringDF)>0:
                body4="TODA/LDA Info: \n"+ str(tkairportDF.iat[cntaprt,0]) + " has been removed from the restricted tankering database. \n" +  "OLD Min of TODA and LDA: "+ str(dummytankeringDF.iat[0,1]) +  " m (TODA)   "+ str(dummytankeringDF.iat[0,2])+ " m (LDA)\n"+"NEW Min of TODA, LDA: " + str(mindistanceTODA) + " m (TODA)      "+ str(mindistanceLDA)+ " m (LDA)"
                mail = outlook.CreateItem(0)
                mail.To = 'odemirhan@corendon-airlines.com'#;  navigation@corendon-airlines.com; occ@corendon-airlines.com'
                mail.Subject = 'Attention! TODA/LDA Information'
                mail.Body=body4
                mail.Send()   
            
            tankeringDF=tankeringDF[tankeringDF["AIRPORT"]!=tkairportDF.iat[cntaprt,0]]
            
            
        

tankeringDF.to_excel("silmePCN/tankering.xlsx", index=None)            

