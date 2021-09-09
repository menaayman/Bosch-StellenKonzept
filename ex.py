import pandas as pd

class sap:
    
    
    def data_frame(path):
            df= pd.read_excel(path)
            return df

    def step_1(IdM_Auszug,Stellenkonzept,x):
        IdM_Auszug = IdM_Auszug[['Organization','User ID','Full Name','Application','IT-Role']]
        IdM_Auszug = IdM_Auszug.rename(columns={"Organization":"Abteilung",'User ID':'Benutzer-ID','Full Name':'Vollständiger Name','IT-Role':'Rolle'})
        IdM_Auszug = IdM_Auszug[IdM_Auszug['Application'] == x] 

        IdM_Auszug.columns = IdM_Auszug.columns.str.strip()
        Stellenkonzept.columns = Stellenkonzept.columns.str.strip()

        IdM_Auszug['Benutzer-ID'] = IdM_Auszug['Benutzer-ID'].str.upper()
        Stellenkonzept['Benutzer-ID'] = Stellenkonzept['Benutzer-ID'].str.upper()



        df1 = IdM_Auszug.copy()
        df2 = Stellenkonzept.copy()
        
        Stellentabele= IdM_Auszug['Benutzer-ID'].isin(Stellenkonzept['Benutzer-ID'])
        df1.drop(df1[Stellentabele].index, inplace=True)

        IDM= Stellenkonzept['Benutzer-ID'].isin(IdM_Auszug['Benutzer-ID'])
        df2.drop(df2[IDM].index, inplace=True)

        # df1= df1.drop(['Unnamed: 6', 'Rollen P81','Stelle P81','Bemerkung.1'], axis = 1)
        # df2= df2.drop(['Kostenstelle', 'Stelle P81','Bemerkung.1','Stelle POE','P72'], axis = 1)

        return df1,df2

    def step_2(IdM_Auszug,Stellenkonzept,Rollenkonzept,x):
        IdM_Auszug = IdM_Auszug[['Organization','User ID','Full Name','Application','IT-Role']]
        IdM_Auszug = IdM_Auszug.rename(columns={"Organization":"Abteilung",'User ID':'Benutzer-ID','Full Name':'Vollständiger Name','IT-Role':'Rolle'})
        IdM_Auszug = IdM_Auszug[IdM_Auszug['Application'] == x] 

        IdM_Auszug.columns = IdM_Auszug.columns.str.strip()
        Stellenkonzept.columns = Stellenkonzept.columns.str.strip()
        Rollenkonzept.columns = Rollenkonzept.columns.str.strip()

        IdM_Auszug['Benutzer-ID'] = IdM_Auszug['Benutzer-ID'].str.upper()
        Stellenkonzept['Benutzer-ID'] = Stellenkonzept['Benutzer-ID'].str.upper()


        Rollenkonzept = Rollenkonzept.drop(['Beschreibung'], axis = 1)
        df3 = pd.melt(Rollenkonzept, id_vars='Rolle', value_vars=None, var_name=None, value_name='value', col_level=None)
        df3 = df3.loc[df3['value'] == 'X']
        df3 = df3.rename(columns={"variable":"Stelle P99"})

        #Stellenkonzept = Stellenkonzept.drop(['Kostenstelle','Bemerkung','Stelle P81','Bemerkung','Stelle POE','P72'], axis = 1)

        arr = Stellenkonzept['Benutzer-ID'].unique()
        ex1=pd.DataFrame()
        ex2=pd.DataFrame()

        for i in arr:
            df4 = Stellenkonzept.loc[Stellenkonzept['Benutzer-ID'] == i]
            df5 = pd.merge(df3,df4, how ='right', on = 'Stelle P99')
            # df5 = df5.rename(columns={"Rollen":"Rolle"})
            
            df6 = IdM_Auszug[IdM_Auszug['Benutzer-ID'] == i]
            Nicht_in_IDM = df5[~df5['Rolle'].isin(df6['Rolle'])]
            Nicht_in_Rollen = df6[~df6['Rolle'].isin(df5['Rolle'])]
            
            ex1=pd.concat([Nicht_in_IDM,ex1])
            ex2=pd.concat([Nicht_in_Rollen,ex2])
        
        ex1= ex1.drop(['value'], axis = 1)
        #ex2= ex2.drop(['Bemerkung','Unnamed: 6','Rollen P81','Stelle P81','Bemerkung.1'], axis = 1)

        return ex1,ex2
    
    def step_3(IdM_Auszug,Stellenkonzept,x):
        IdM_Auszug = IdM_Auszug[['Organization','User ID','Full Name','Application','IT-Role']]
        IdM_Auszug = IdM_Auszug.rename(columns={"Organization":"Abteilung",'User ID':'Benutzer-ID','Full Name':'Vollständiger Name','IT-Role':'Rolle'})
        IdM_Auszug = IdM_Auszug[IdM_Auszug['Application'] == x] 

        IdM_Auszug.columns = IdM_Auszug.columns.str.strip()
        Stellenkonzept.columns = Stellenkonzept.columns.str.strip()

        IdM_Auszug['Benutzer-ID'] = IdM_Auszug['Benutzer-ID'].str.upper()
        Stellenkonzept['Benutzer-ID'] = Stellenkonzept['Benutzer-ID'].str.upper()

        IdM_Auszug = IdM_Auszug[['Abteilung','Benutzer-ID','Vollständiger Name','Rolle']]
        Stellenkonzept = Stellenkonzept[['Benutzer-ID','Stelle P99']]
        df7 = IdM_Auszug.merge(Stellenkonzept, on='Benutzer-ID',how='left')
        
        return df7
