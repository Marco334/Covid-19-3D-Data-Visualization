#import teradata
#import teradatasql
import pandas as pd
import bpy
import math
import numpy as np

#-----------------------
# CORONA VISUS 19 - 3D data visualization 
# DATA SOURCE : https://www.worldometers.info/coronavirus/
# Author      : Marco Genca 
#-----------------------
#TO DO LIST:
#- WEB  screping
#- DWH  Connection
#- ...
#-----------------------
SOURCE_PTH   = str('C:\\......')
SOURCE_FL_NM = '20201001_CORONA.xlsx'
PROP_FACT    = 1 # proportional factor to be able to visualize to unbilanced numbers
sheet_NM     = 0
X_SPOT       = 0
Y_SPOT       = 0
FT_P         = 0.5 #fattore di distanziamento tra oggetti 
ORAN  = bpy.data.materials.new(name="Orange_T" )
ORAN.diffuse_color = (0.98,0.225,0.01,1)
GLASS = bpy.data.materials.new(name="Glass_V"  )
GLASS.diffuse_color = (0.1, 0.5, 0.7,0)
ROSSO = bpy.data.materials.new(name="Red_S"    )
ROSSO.diffuse_color = (1, 0, 0,0)
VERDE = bpy.data.materials.new(name="Green_S"  )
VERDE.diffuse_color = (0, 1, 0,0)

def CLEAN_DF():
    del CL_df
    return()

def EXCEL_SOURCE_LOAD(SOURCE_PTH,SOURCE_FL_NM,sheet_NM):
   '''SOURCE EXCEL IN DATAFRAME'''
   #CL_df = pd.read_csv(str(SOURCE_PTH + SOURCE_FL_NM ),sheet_name = sheet_nm )
   CL_df = pd.read_excel(str(SOURCE_PTH + SOURCE_FL_NM ),sheet_name = sheet_NM )
   return(CL_df)
   
def GENERAL_LBL():  
   '''EXPLANATION LABEL MANAGEMENT'''  
   OBJ_NE = "TEST per Milion"
   font_curve = bpy.data.curves.new(type="FONT",name=OBJ_NE)
   font_curve.body = OBJ_NE.strip()
   str_len = len(font_curve.body.strip())
   font_obj = bpy.data.objects.new( OBJ_NE, font_curve)
   font_obj.location = (- 3 , -1  , 0 )
   font_obj.scale = ( 1, 1, 1)  
   font_obj.rotation_euler.z = radians(270)
   #ob.context.active_object.data.materials.append(VERDE)
   bpy.data.collections['DESCRIPTION_LAB'].objects.link(font_obj)
   #-----------------------
   OBJ_NE = "TOTAL DEATH ( Volume)"
   font_curve = bpy.data.curves.new(type="FONT",name=OBJ_NE)
   font_curve.body = OBJ_NE.strip()
   str_len = len(font_curve.body.strip())
   font_obj = bpy.data.objects.new( OBJ_NE, font_curve)
   font_obj.location = (0 , -1  , 0 )
   font_obj.scale = ( 1, 1, 1)  
   font_obj.rotation_euler.z = radians(270)
   #ob.context.active_object.data.materials.append(ROSSO)
   bpy.data.collections['DESCRIPTION_LAB'].objects.link(font_obj)
   #-----------------------
   OBJ_NE = "TOTAL TESTED POSITIVE"
   font_curve = bpy.data.curves.new(type="FONT",name=OBJ_NE)
   font_curve.body = OBJ_NE.strip()
   str_len = len(font_curve.body.strip())
   font_obj = bpy.data.objects.new( OBJ_NE, font_curve)
   font_obj.location = (2 , -1  , 0 )
   font_obj.scale = ( 1, 1, 1)  
   font_obj.rotation_euler.z = radians(270)
   #ob.context.active_object.data.materials.append(ROSSO)
   bpy.data.collections['DESCRIPTION_LAB'].objects.link(font_obj)
   
   return()
   
def LBL_MNG_1(OBJ_N,X_SPOT,Y_SPOT,Z_SPOT,C_SIZE) :
   '''NATIONS LABEL MANAGEMENT''' 
   #myFont= Text3d.Font.Load('C:\WINDOWS\Fonts\impact.ttf') 
   font_curve = bpy.data.curves.new(type="FONT",name=OBJ_N)
   font_curve.body = OBJ_N.strip()
   str_len = len(font_curve.body.strip())
   font_obj = bpy.data.objects.new( OBJ_N, font_curve)
   font_obj.location = (- X_SPOT - 1.8 , Y_SPOT-(C_SIZE/2), Z_SPOT )
   font_obj.scale = ( C_SIZE, C_SIZE, 1)  
   #ob.context.active_object.data.materials.append(VERDE)
   bpy.data.collections['NATIONS_LAB'].objects.link(font_obj)
   return()
   
def LBL_MNG_2(OBJ_N,X_SPOT,Y_SPOT,Z_SPOT,C_SIZE) :
   '''NATIONS LABEL MANAGEMENT''' 
   #myFont= Text3d.Font.Load('C:\WINDOWS\Fonts\impact.ttf') 
   font_curve = bpy.data.curves.new(type="FONT",name=OBJ_N)
   font_curve.body = OBJ_N.strip()
   str_len = len(font_curve.body.strip())
   font_obj = bpy.data.objects.new( OBJ_N, font_curve)
   font_obj.location = (Z_SPOT, Y_SPOT-(C_SIZE/2), Z_SPOT - X_SPOT - 1.8 )
   font_obj.scale = ( C_SIZE, C_SIZE, 1)  
   font_obj.rotation_euler.y = radians(270)
   #ob.context.active_object.data.materials.append(VERDE)
   bpy.data.collections['NATIONS_LAB'].objects.link(font_obj)
   return()
   
def CUBE_GEN(M_SIZE ,COLOR_V,OBJ_N,TOT_TPM_NM):
   ''' CUBE GENERATOR ''' 
   print(str( OBJ_N) +'   '+  str( M_SIZE )+'   '+ str( pow( M_SIZE, 1/3 )/100 ) ) 
   global Y_SPOT  
   global FT_P  
   C_SIZE = pow( M_SIZE, 1/3 )/100   #radice 2 per calcolare il volume conversione a CM
   X_SPOT = C_SIZE/2 
   Y_SPOT = Y_SPOT + ( C_SIZE/2)
   Z_SPOT = C_SIZE/2
   bpy.ops.mesh.primitive_cube_add( size = C_SIZE, location=(X_SPOT , Y_SPOT, Z_SPOT ))
   cube = bpy.context.selected_objects[0]
   bpy.context.active_object.name = str( OBJ_N )
   obj = bpy.context.active_object
   bpy.data.collections['DTH_CB'].objects.link(obj)
   bpy.context.scene.collection.objects.unlink(obj)
   bpy.context.active_object.data.materials.append(COLOR_V)
   LBL_MNG_1(OBJ_N,X_SPOT,Y_SPOT,Z_SPOT,C_SIZE) # Per la generazione delle label nazione
   LBL_MNG_2(OBJ_N,X_SPOT,Y_SPOT,Z_SPOT,C_SIZE) # Per la generazione delle label nazione
   CUBE_GEN_3(OBJ_N,X_SPOT,Y_SPOT,Z_SPOT,TOT_TPM_NM,C_SIZE)
   Y_SPOT = Y_SPOT + ( C_SIZE/2) + (C_SIZE*FT_P) # per creare saio tra un cubo e il successivo
   return()
   
def CUBE_GEN_2(M_SIZE ,COLOR_V,OBJ_N,X_PLUS,CONTAG):
   ''' CUBE GENERATOR '''
   global Y_SPOT 
   global FT_P   
   C_SIZE = pow( M_SIZE, 1/3 ) /100   #radice 2 per calcolare il volume conversione a CM
   X_SPOT = (math.sqrt( CONTAG/C_SIZE )/100 )/2 + ((pow( X_PLUS, 1/3 ) /100)*1.8) 
   Y_SPOT = Y_SPOT   + (C_SIZE/2  )
   Z_SPOT = (math.sqrt( CONTAG/C_SIZE )/100)/2
   #----------------------------------
   X_SIZE = math.sqrt( CONTAG/C_SIZE )/100 
   Y_SIZE = C_SIZE
   Z_SIZE = math.sqrt( CONTAG/C_SIZE )/100 
   bpy.ops.mesh.primitive_cube_add( size = 1, location=(X_SPOT , Y_SPOT, Z_SPOT ))
   #bpy.ops.mesh.primitive_cube_add( size = C_SIZE, location=(X_SPOT , Y_SPOT, Z_SPOT ))
   bpy.ops.transform.resize(value=( X_SIZE, Y_SIZE, Z_SIZE) )
   cube = bpy.context.selected_objects[0]
   bpy.context.active_object.name = str( OBJ_N )
   obj = bpy.context.active_object
   bpy.data.collections['DTH_CB'].objects.link(obj)
   bpy.context.scene.collection.objects.unlink(obj)
   bpy.context.active_object.data.materials.append(COLOR_V)
   Y_SPOT = Y_SPOT + ( C_SIZE/2) + (C_SIZE*FT_P ) # per creare saio tra un cubo e il successivo
   return() 
    
def CUBE_GEN_3(OBJ_N ,X_SPOT,Y_SPOT,Z_SPOT,TOT_TPM_NM,C_SIZE):
   ''' CUBE GENERATOR PEr test over milion''' 
   OBJ_Nh = str(OBJ_N + "_TPM")
   X_SPOT =  - 2.5 - TOT_TPM_NM/2000000 
   Y_SPOT = Y_SPOT 
   Z_SPOT = TOT_TPM_NM/2000000 
   #---------------------------------- 
   X_SIZE = TOT_TPM_NM/1000000 
   #Y_SIZE = C_SIZE
   Y_SIZE = 0.1
   Z_SIZE = TOT_TPM_NM/1000000 
   bpy.ops.mesh.primitive_cube_add( size = 1, location=(X_SPOT , Y_SPOT, Z_SPOT ))
   #bpy.ops.mesh.primitive_cube_add( size = C_SIZE, location=(X_SPOT , Y_SPOT, Z_SPOT ))
   bpy.ops.transform.resize(value=( X_SIZE, Y_SIZE, Z_SIZE) )
   cube = bpy.context.selected_objects[0]
   bpy.context.active_object.name = str( OBJ_Nh )
   obj = bpy.context.active_object
   bpy.data.collections['DTH_CB'].objects.link(obj)
   bpy.context.scene.collection.objects.unlink(obj)
   bpy.context.active_object.data.materials.append(VERDE)
   return()
    
def CLLCT_GEN():
   '''COLLECTION MANAGEMENT'''
   collection_T = bpy.data.collections.new('DESCRIPTION_LAB')
   bpy.context.scene.collection.children.link(collection_T)
   collection_T = bpy.data.collections.new('NATIONS_LAB')
   bpy.context.scene.collection.children.link(collection_T)
   collection_T = bpy.data.collections.new('NATIONS_TAB')
   bpy.context.scene.collection.children.link(collection_T)
   collection_V = bpy.data.collections.new('DTH_CB')
   bpy.context.scene.collection.children.link(collection_V)
   collection_X = bpy.data.collections.new('POP_CB')
   bpy.context.scene.collection.children.link(collection_X)
   collection_S = bpy.data.collections.new('CSS_CB')
   bpy.context.scene.collection.children.link(collection_S)
   return(1)

#--------- COLOR MANAGEMENT
#The four values are represented as: [Red, Green, Blue, Alpha]
ORAN  = bpy.data.materials.new(name="Orange_T" )
ORAN.diffuse_color = (0.98,0.225,0.01,1)
GLASS = bpy.data.materials.new(name="Glass_V"  )
GLASS.diffuse_color = (0.1, 0.5, 0.7,0)
ROSSO = bpy.data.materials.new(name="Red_S"    )
ROSSO.diffuse_color = (1, 0, 0,0)
VERDE = bpy.data.materials.new(name="Green_S"  )
VERDE.diffuse_color = (0, 1, 0,0)

SOURCE_PTH   = str('C:\\Users\\evand\\OneDrive\\Documenti\\GRAFICA\\BLEDEL\\RANKING_PROJECT\\01_SOURCES\\')
SOURCE_FL_NM = '20201001_CORONA.xlsx'
sheet_NM     = 0

CL_DF = EXCEL_SOURCE_LOAD(SOURCE_PTH,SOURCE_FL_NM,sheet_NM)
CLLCT_GEN()    # BL Collections generation
CL_DF.sort_values(by=['Total_Deaths'], ascending=False) #ascending Dataframe ORDER BY / SORT

for index, row in CL_DF.iterrows():
 if index < 65:
  OBJ_N      = str(row['Country'        ])
  TOT_DTH_NM = float(row['Total_Deaths' ])
  TOT_CS_NM  = int(row['Total_CASES'    ])
  TOT_POP_NM = int(row['Population'     ])
  TOT_TPM_NM = int(row['TestsP1MIL'     ])
  #print(str(index)  + "   " + OBJ_N + "   " + str(TOT_DTH_NM) )
  CUBE_GEN(TOT_DTH_NM, ROSSO, OBJ_N,TOT_TPM_NM) #CREAZIONE CUBI

MAX_SZ    = CL_DF['Total_Deaths'].max()
#print(MAX_SZ)
X_SPOT    = 0
PROP_FACT = 1 
index     = 0
Y_SPOT    = 0
for index, row in CL_DF.iterrows():
 if index < 65:
  OBJ_N      = str(row['Country'        ])
  TOT_DTH_NM = float(row['Total_Deaths' ])
  TOT_CS_NM  = int(row['Total_CASES'    ]) 
  Y_PLUS = TOT_DTH_NM
  #print(str(index)  + "   " + OBJ_N + "   " + str(TOT_CS_NM) )
  CUBE_GEN_2(TOT_DTH_NM , ORAN, OBJ_N + "_C" , MAX_SZ,TOT_CS_NM) #CREAZIONE CUBI
  GENERAL_LBL()
