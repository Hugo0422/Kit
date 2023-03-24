import pandas as pd 
import tkinter as tk 
import re
from tkinter import filedialog
import openpyxl
import math




def DealDATA():



    f_path = filedialog.askopenfilename()
    print(f_path)

    file_type = f_path.split(".")[1]

    #print(file_type)

    if file_type in ["xlsx","xls"]:
        data = pd.read_excel(f_path)
    
    elif file_type in ["csv"]:
        data = pd.read_csv(f_path)


    #行驶模式run mode        
    M_index = 197       
    MM = data.iloc[:,M_index]
    #MM = MM.apply(lambda x: x.replace('运行模式:纯电','1').replace('运行模式:混动','2') )
    #MM = pd.DataFrame(MM.str.split('、').tolist())
    MM.columns = pd.Index(['run_mode'])



    #电压、电流、SOC、里程
    V_index = 198
    C_index = 199
    SOC_index = 200
    #R_index = 79
    F_index = 11

    CC = data.iloc[:,C_index]  
    #CC = CC.str.replace("A","").str.replace(":","").str.replace("总电流","")
    #CC = pd.DataFrame(CC.str.split('、').tolist())
    CC.columns = pd.Index(['current'])

    VV = data.iloc[:,V_index]  
    #VV = VV.str.replace("V","").str.replace(":","").str.replace("总电压","")
    #VV = pd.DataFrame(VV.str.split('、').tolist())
    VV.columns = pd.Index(['totalVoltage'])

    SS = data.iloc[:,SOC_index]  
    #SS = SS.str.replace("%","").str.replace(":","").str.replace("SOC","")
    #SS = pd.DataFrame(SS.str.split('、').tolist())
    SS.columns = pd.Index(['SOC'])



    FF = data.iloc[:,F_index]  
    #FF = FF.str.replace("km","").str.replace(":","").str.replace("总里程","")
    #FF = pd.DataFrame(FF.str.split('、').tolist())
    FF.columns = pd.Index(['total_KM'])




    #时间
    col_Time_index = 3
    col_Time = data.iloc[:,col_Time_index] 
    #col_Time = pd.DataFrame(col_Time.str.split('、').tolist())
    col_Time.columns = pd.Index(['time'])
    #print(col_Time)  

    #单体温度Temp    
    col_T_index = 285
    col_T = data.iloc[:,col_T_index]
    col_T = col_T.str.replace("[","").str.replace("]","")
    T_3 = pd.DataFrame(col_T.str.split(',').tolist())
    T_3.columns = pd.Index(['Cell_' + str(col+1) for col in T_3.columns]) 
    #print(T_3)

    #单体电压Cell

    col_V_index = 282
    col_V = data.iloc[:,col_V_index]
    col_V = col_V.str.replace("[","").str.replace("]","")
    V_3 = pd.DataFrame(col_V.str.split(',').tolist())
    V_3.columns = pd.Index(['Cell_' + str(col+1) for col in V_3.columns])

    #print(V_3)

    Final = pd.concat([col_Time,VV,CC,SS,FF,MM,V_3,T_3], axis=1)


    save_path = filedialog.asksaveasfilename(defaultextension="output.xlsx", filetypes=[("Excel Files", "*.xlsx")])

    writer = pd.ExcelWriter(save_path)

    Final.to_excel(writer, index=False)

    writer.save()
    writer.close()

    return True

def on_close(self):
        self.root.destroy()
        
def main():
        key = False
        root = tk.Tk()
        root.geometry('350x80')
        root.title('Car Data processing')
        label = tk.Label(root, text='运行中')
        label.pack()
        key = DealDATA()
        if key == True:
            label.config(text='运行结束\n数据已输出至桌面\n需要在excel中进行修正')
        root.mainloop()
    
    
    
    
    
    
if __name__ == '__main__':
    main()