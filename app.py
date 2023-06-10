import pandas as pd
import streamlit as st
from dateutil import parser
import streamlit as st
from annotated_text import annotated_text
import datetime as dt
import numpy as np
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import xlsxwriter 
from yfinance import download

def do_stuff_on_page_load():
    st.set_page_config(layout="wide")
"streamlit run appname.py"
def a():
    st.sidebar.subheader('MarkUp(%)')
    a = st.sidebar.text_input("Use ',' como separador", key = 7, placeholder = "Exemplo(45,7%)")
    return a

def b():
    st.sidebar.subheader('Câmbio')
    b = st.sidebar.text_input('O câmbio a ser utilizado', key = 4)
    return b

def c():
    st.sidebar.subheader('Unidades')
    c = st.sidebar.text_input("Número de unidades", key = 8)
    return c

st.sidebar.title("Dados")

with st.container():

    col1, col2 = st.columns((3,3))

    options = st.sidebar.selectbox(
        'Opção',
        ['Selecione a opção', 'Cálculo', 'Consulta de múltiplos produtos', 'Cálculo de múltiplos produtos'])

    ok = st.sidebar.button('OK')
    

    if options == 'Cálculo':
        with col1:
            st.subheader('Planta')
            planta = st.text_input('Planta responsável pela produção', key = 5)

        with col2:
            st.subheader('Produto')
            produto = st.text_input('Código do material', key = 6)

# # Limpeza e organização
# CMM
@st.cache
def fcmm():
     
    filecmm = pd.read_excel(r'C:\Users\R6435260\Documents\Calculo_ATP\CMM_SGB.xlsx', sheet_name='Query' )
    df_cmm = pd.DataFrame(filecmm)
    df_cmm = df_cmm.iloc[:, [2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,]]

    df_cmm.drop(index = [0,1,2], inplace=True)

    new_header = df_cmm.iloc[0] 
    df_cmm = df_cmm[1:]          
    df_cmm.columns = new_header 

    df_cmm.columns.values[0] = "Code"
    df_cmm.columns.values[3] = "Type"

    df_cmm2 = df_cmm
    df_cmm3 = df_cmm

    df_cmm2 = df_cmm.iloc[:,[5,6,7,8,9,9,10,11,12,13,14,15,16,17,18,19]]

    df_cmm3 = df_cmm.iloc[:,[0,1,2,3,4]]

    df_cmm2.dropna(thresh = 1, inplace = True) #exclui somente as linhas com NaN
    df_cmm3.dropna(thresh = 1, inplace = True) #exclui somente as linhas com NaN
    df_cmm2.fillna(method='ffill', axis = 1, inplace= True)
    df_cmm4 = df_cmm3.join(df_cmm2, how = 'right', lsuffix='_left', rsuffix='_right')
    dum = df_cmm4.melt(id_vars=['Code','Plant','Material Plant View','Type','UOM']
                ,var_name=['JAN 2021'], value_name='Valor').sort_values(['JAN 2021'])

    dum['Período'] = pd.to_datetime(dum['JAN 2021'], infer_datetime_format=True)
    dum['Code'] = dum['Code'].astype(str)
    dum = dum[['Code','Plant','Material Plant View','Type','Valor','UOM','Período']]
    dum['Code'] = dum['Material Plant View']
    dum['CMM'] = dum['Valor']
    dum['Key'] = dum['Code'] + dum['Plant']
    dum = dum[['Key','CMM','Período']]
    dum = dum.sort_values(by = 'Período')
    df_cmm = dum
    return df_cmm

# # ATP
@st.cache
def fatp():
    columns = ['Code1','Material1','1','2', "3",'Code2','Code','Type','Plant','Plant2',
                  'Profit Center','ID: Object Deleted','Deletion flag Material Generic',
                  'Deletion flag Material Plant','Deletion flag Material Sales',
                  'Validity Per. Start','Validity period end','Customer','Customer2','Currency',
                  'Base Currency (SOrg)','Scale Unit','Price Unit (KONP)','Scale Qty (KONM)',
                  'ATP','Price Unit (ZPRA)','ZPRA per 1','ano','mercado'] 
    use_cols = ('G:K, P:AA')
    file = pd.read_excel(r'C:\Users\R6435260\Documents\Calculo_ATP\ATP_SGB.xlsx', sheet_name='Table'
                        , header = 14, names = columns
                        , dtype = {'Plant':str,'Code':str}
                        , usecols = use_cols)
    df_atp = pd.DataFrame(file)
    df_atp.dropna(how='any', thresh = 8, inplace=True)#exclui somente as linhas com NaN
    pd.set_option('display.max_columns', None)
    df_atp['Key'] = df_atp['Code'] + df_atp['Plant']
    df_atp['ATP'] = df_atp['ATP'].astype(str).radd('$')
    return df_atp

# # Standard
@st.cache
def fstd():
    filestd = pd.read_excel(r'C:\Users\R6435260\Documents\Calculo_ATP\2022 - STD Query - SGB (1).xlsx', sheet_name='Planilha1')
    df_std = pd.DataFrame(filestd)
    df_std.columns = ['Profit Center', 'Plant', 'Plant Name', 'Status', 'Status Code', 'Code', 'Material Name', 'Base', 'STD']
    df_std = df_std[df_std['Profit Center'].str.contains('#')==False ]
    df_std['Key'] = df_std["Code"] + df_std["Plant"]
    df_std.dropna(inplace=True)
    df_std = df_std[['Key','STD', 'Plant Name', 'Material Name']]
     
    return df_std

if options == 'Cálculo':

    df_atp = fatp()
    df_std = fstd()
    df_cmm = fcmm()

    MarkUp = a()
    cambio = b()
    st.sidebar.subheader('Moeda')
    moeda = st.sidebar.selectbox(
        'Opção',
        ['Selecione a opção', 'Dólar', 'Euro', 'Yuan'])

    Unidades = c()

    chave = produto + planta

    atp = df_atp[df_atp['Key'] == chave] 
    cmm = df_cmm[df_cmm['Key'] == chave]
    std = df_std[df_std['Key'] == chave]['STD']

    cmm['Período'] = cmm['Período'].dt.strftime('%Y-%m-%d').apply(str)
        
    dem_cmm = cmm[['CMM', 'Período']].tail(6)
    dem_cmm = dem_cmm.set_index('Período')

    st.sidebar.write(dem_cmm)

    options2 = st.sidebar.multiselect(label='Select desired month:',options=cmm['Período'].tail(6).unique(), default = cmm['Período'].tail(6).unique())

    df_avg = pd.DataFrame(options2,columns=['Período'])

    cmm = df_avg.merge(cmm, on='Período')

    cmm2 = cmm
    cmm2 = cmm2['CMM']
    cmm2 = cmm2.astype(float)
    cmm2 = cmm2.mean()

    st.sidebar.caption('Média:')
    st.sidebar.write(str(round(cmm2, 2)))

    calcular = st.sidebar.button('Calcular')

    if calcular:
        nome_planta = df_std.loc[df_std['Key'] == chave, 'Plant Name'].values[0]

        nome_material = df_std.loc[df_std['Key'] == chave, 'Material Name'].values[0]

        col5, col6, col7 = st.columns((2,2,2))

        cmm = cmm.tail(6)
        cmm = cmm['CMM']   
        cmm = cmm.astype(float)
        cmm = cmm.mean()

        MarkUp = MarkUp.replace(',', '.').strip('%')
        MarkUp = float(MarkUp)/100
        atp_pelocmm = round(float(Unidades) * (cmm*(1+MarkUp)/float(cambio)), 2)

        std = std.mean()

        with col1:

            st.subheader('Produto:')
            st.caption('Código do material')
            annotated_text((produto, '', '#8ef'))            
            annotated_text((nome_material, '', '#8ef'))
        
            st.subheader('Planta:')
            st.caption('Planta produzida')
            annotated_text((planta, '', '#8ef'))
            annotated_text((nome_planta, '', '#8ef'))

        with col2:

            st.subheader('Standard:')
            st.caption('Valor padrão em reais para 1 unidade:')
            annotated_text((str(round(std, 2)), '', '#8ef'))
            annotated_text('', '', )

            st.subheader('ATP calculado')
            st.caption('Cálculo final em ' + moeda + ' para ' + Unidades + ' unidades' ':')
            annotated_text((str(atp_pelocmm), '', "#faa"))

# # Multiplos

elif options == 'Consulta de múltiplos produtos':

    t = st.text_area("Enter multiline text")

    df_atp = fatp()
    df_std = fstd()
    df_cmm = fcmm()

    if t is not None:
        textsplit = t.splitlines()
        st.write(textsplit)

    df = pd.DataFrame()
    i = 0 
    con_but = st.button('Consultar')

    if con_but:

        st.header('Consulta')
        
        while i < len(textsplit):
            std = df_std.loc[df_std['Key'] == textsplit[i], 'STD']
            cmm = df_cmm.loc[df_cmm['Key'] == textsplit[i], 'CMM'].tail(6).reset_index().drop(columns = "index").astype(float)
            cmm = pd.DataFrame(cmm)
            cmm1 = cmm.loc[[0]]
            cmm1.columns = ['CMM1'] 
            cmm2 = cmm.loc[[1]]
            cmm2.columns = ['CMM2'] 
            cmm3 = cmm.loc[[2]]
            cmm3.columns = ['CMM3'] 
            cmm4 = cmm.loc[[3]]
            cmm4.columns = ['CMM4'] 
            cmm5 = cmm.loc[[4]]
            cmm5.columns = ['CMM5'] 
            cmm6 = cmm.loc[[5]]
            cmm6.columns = ['CMM6'] 
            atp = df_atp.loc[df_atp['Key'] == textsplit[i], 'ATP'].head(1)
            plant = df_std.loc[df_std['Key'] == textsplit[i], 'Plant Name']
            mat = df_std.loc[df_std['Key'] == textsplit[i], 'Material Name']

            std = pd.DataFrame(std)
            atp = pd.DataFrame(atp)
            plant = pd.DataFrame(plant)
            mat = pd.DataFrame(mat)
            
            pdList = [df, mat, plant, std, atp, cmm1, cmm2, cmm3, cmm4, cmm5, cmm6]
            df = pd.concat(pdList, ignore_index=True)
            i = i + 1
                
        df = df.apply(lambda x: pd.Series(x.dropna().values))
        st.dataframe(df)

        def to_excel(df):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            format1 = workbook.add_format({'num_format': '0.00'}) 
            worksheet.set_column('A:A', None, format1)  
            writer.save()
            processed_data = output.getvalue()
            return processed_data
        df_xlsx = to_excel(df)
        st.download_button(label='📥 Download Current Result',
                                    data=df_xlsx ,
                                    file_name= 'df_test.xlsx')

elif options == 'Cálculo de múltiplos produtos':

    t = st.text_area("Enter multiline text")

    df_atp = fatp()
    df_std = fstd()
    df_cmm = fcmm()

    if t is not None:
        textsplit = t.splitlines()

    MarkUp = a()
    cambio = b()
    Unidades = c()

    df = pd.DataFrame()
    i = 0 
    
    selected_code = st.sidebar.selectbox(
        'Selecione a chave',
        textsplit)

    if selected_code:

        cmm = df_cmm[df_cmm['Key'] == selected_code]
        cmm['Período'] = cmm['Período'].dt.strftime('%Y-%m-%d').apply(str)
         
        dem_cmm = cmm[['CMM', 'Período']].tail(6)
        dem_cmm = dem_cmm.set_index('Período')

        st.sidebar.write(dem_cmm)

        options2 = st.sidebar.multiselect(label='Selecione o mês:',options=cmm['Período'].tail(6).unique(), default = cmm['Período'].tail(6).unique())

        df_avg = pd.DataFrame(options2,columns=['Período'])

        cmm = df_avg.merge(cmm, on='Período')

        cmm2 = cmm
        cmm2 = cmm2['CMM']
        cmm2 = cmm2.astype(float)
        cmm2 = cmm2.mean()

        st.sidebar.caption('Média:')
        st.sidebar.write(str(round(cmm2, 2)))

        ok3 = st.sidebar.button('Calcular')

        if ok3:

            st.sidebar.header('Cálculo')
            st.sidebar.write(MarkUp)
            MarkUp = float(MarkUp.replace(',', '.').strip('%'))/100
            Unidades = float(Unidades)
            cambio = float(cambio)

            while i < len(textsplit):
                std = df_std.loc[df_std['Key'] == textsplit[i], 'STD']
                cmm = df_cmm.loc[df_cmm['Key'] == textsplit[i], 'CMM'].tail(6).reset_index().drop(columns = "index").astype(float)
                cmm_mean = round(cmm.mean(), 2)
                cmm_sd = round(cmm.std(), 2 )
                cmm_min = cmm.sort_values('CMM', ascending = True).head(1).astype(float)
                cmm_min.columns = ['CMM mínimo']
                cmm_max = cmm.sort_values('CMM',ascending = True).tail(1).astype(float)
                dist = (cmm_max-cmm_min)/6
                dist = dist.divide(6)
                cmm_max.columns = ['CMM máximo'] 
            
                atp_cal = Unidades * (cmm_mean*(1+MarkUp)/cambio) ## IF PARA CMM NULO, CALCULAR COM STANDARD
                #atp_cal = "%.2f" % atp_cal.iloc[0]
                plant = df_std.loc[df_std['Key'] == textsplit[i], 'Plant Name']
                mat = df_std.loc[df_std['Key'] == textsplit[i], 'Material Name']

                std = pd.DataFrame(std)
                atp = pd.DataFrame(atp_cal, columns = ['ATP'])

                plant = pd.DataFrame(plant)
                mat = pd.DataFrame(mat)
                cmm_mean = pd.DataFrame(cmm_mean, columns = ['CMM']).round(2)
                cmm_sd = pd.DataFrame(cmm_sd, columns = ['CMM SD'])
                #cmm_dist = pd.DataFrame(dist, columns = ['Distância'])
            
                #key = pd.DataFrame(textsplit[i], columns = ['Código'])

                pdList = [df, mat, plant, std, atp, cmm_mean, cmm_sd]
                df = pd.concat(pdList, ignore_index=True)
                i = i + 1

            df = df.apply(lambda x: pd.Series(x.dropna().values))
            st.dataframe(df)

            def to_excel(df):
                output = BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                format1 = workbook.add_format({'num_format': '0.00'}) 
                worksheet.set_column('A:A', None, format1)  
                writer.save()
                processed_data = output.getvalue()
                return processed_data
            df_xlsx = to_excel(df)
            st.download_button(label='📥 Download Current Result',
                                        data=df_xlsx ,
                                        file_name= 'df_test.xlsx')
                            
                            ### PROBLEMAS: NÃO CONSIGO DEIXAR APENAS DUAS CASAS
                            ####           PENSAR NA ESTÉTICA (INCLUSIVE UNIDADES E MOEDAS)
                            ###            ESTUDAR POSSIBILIDADES DE ALTERAÇÃO DA MÉDIA CMM PARA MÚLTIPLOS PRODUTOS
