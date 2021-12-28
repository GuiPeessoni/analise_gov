import tkinter as tk
from tkinter import messagebox
from tkcalendar import DateEntry
import time
from selenium import webdriver
import os
from datetime import  datetime,timedelta
import pandas as pd
import numpy as np
import getpass


user = getpass.getuser()
data = datetime.today().date()
ano = data.year
hoje = data.strftime('%d%m%Y')
selecionadas = []


def checkbox():
    selecionadas.clear()
    if (check_01.get()) == True:
        selecionadas.append('EMP_01')
    if (check_02.get()) == True:
        selecionadas.append('EMP_02')
    if (check_03.get()) == True:
        selecionadas.append('EMP_03')
    if (check_04.get()) == True:
        selecionadas.append('EMP_04')
    campos()

def campos():
    if len(login.get()) == 0 or len(senha.get()) == 0 or len(data_inicio.get()) == 0 or len(data_fim.get()) == 0:
        messagebox.showinfo("Nibbler", """✦✦  Preencha os Campos da forma Correta  ✦✦""")
    elif len(selecionadas) == 0 :
        messagebox.showinfo("Nibbler", """✦✦  Preencha os Campos da forma Correta  ✦✦""")
    else:
        scrap_gov()

def scrap_gov():
    try:
        caminho_driver = r'C:/Nibbler/chromedriver.exe'
        driver = webdriver.Chrome(r'C:/Nibbler/chromedriver.exe')
        driver.quit()

    except:
        messagebox.showinfo("Nibbler", """
       ✦✦  Verifique o ChromeDriver  ✦✦""")

    else:

        print(selecionadas)
        driver = webdriver.Chrome(caminho_driver)
        driver.get(r'https://www.consumidor.gov.br/')
        time.sleep(2)
        driver.get(r'https://www.consumidor.gov.br/pages/administrativo/login')
        driver.maximize_window()
        time.sleep(2)

        driver.find_element_by_id("login").send_keys(login.get())
        driver.find_element_by_id("senha").send_keys(senha.get())
        driver.find_element_by_id("btnLoginPageForm").click()
        time.sleep(2)
        caminho_download = r'C:/Users/{}/Downloads/'.format(user)

        empresas = {
            'EMP_01': '//*[@id="menu_sel_fornecedor"]/option[2]',
            'EMP_02': '//*[@id="menu_sel_fornecedor"]/option[3]',
            'EMP_03': '//*[@id="menu_sel_fornecedor"]/option[4]',
            'EMP_04': '//*[@id="menu_sel_fornecedor"]/option[5]',
        }


        for sel in selecionadas:
            a = empresas[sel]
            print(a)
            driver.find_element_by_id('menu_sel_fornecedor').click()
            time.sleep(0.2)
            driver.find_element_by_xpath(a).click()
            driver.get(r'https://www.consumidor.gov.br/pages/exportacao-dados/novo')
            driver.find_element_by_id("dataIniPeriodo").send_keys(data_inicio.get())
            time.sleep(0.2)
            driver.find_element_by_id("dataFimPeriodo").send_keys(data_fim.get())
            time.sleep(0.2)
            driver.find_element_by_id("colunasExportadas1").click()
            driver.find_element_by_id("btnExportar").click()
            time.sleep(5)
            os.rename(caminho_download + 'consumidor-web-exportar-dados.csv', caminho_download + sel + '.csv')
            time.sleep(1)

        driver.quit()
        messagebox.showinfo("Nibbler", """✦✦ O Nibbler baixou os Arquivos ✦✦
    ✦✦ Agora vai Compilar eles ✦✦""")

        leitura()


def leitura():
    cols = ['Gestor','Protocolo','Canal de Origem','Consumidor','CPF',
            'Região','UF','Cidade','Sexo','Faixa Etária','Ano Abertura',
            'Mês Abertura','Data Abertura','Hora Abertura','Data Resposta',
            'Hora Resposta','Data Análise','Hora Análise','Data Recusa',
            'Hora Recusa','Data Finalização','Hora Finalização','Prazo Resposta',
            'Prazo Analise Gestor','Tempo Resposta','Grupo Econômico',
            'Nome Fantasia','Segmento de Mercado','Área','Assunto',
            'Grupo Problema','Problema','Como Comprou Contratou','Procurou Empresa',
            'Técnico do Último Trâmite','Respondida','Situação','Avaliação Reclamação',
            'Nota do Consumidor','Análise da Recusa','Edição de Conteúdo',
            'Interação do Gestor','Avaliação Positiva']

    df_vazio = pd.DataFrame(columns=cols)

    schema = {
        'CPF': str,
        'Protocolo': str,}

    datatype = ['Data Abertura', 'Data Resposta', 'Data Análise',
                'Data Recusa', 'Data Finalização', 'Prazo Resposta']

    caminho_download = r'C:/Users/{}/Downloads/'.format(user)
    emp_01 = 'EMP_01.csv'
    emp_02 = 'EMP_02.csv'
    emp_03 = 'EMP_03.csv'
    emp_04 = 'EMP_04.csv'

    writer = pd.ExcelWriter(caminho_download + 'consumidor_gov' + hoje + '.xlsx', engine='xlsxwriter')

    try:
        emp_01 = pd.read_csv(caminho_download + emp_01, sep=';',index_col=False,
                              dtype=schema)
    except:
        df_vazio.to_csv(caminho_download + emp_01, sep=';',index=False)
    try:
        emp_02 = pd.read_csv(caminho_download + emp_02, sep=';',index_col=False,
                              dtype=schema)
    except:
        df_vazio.to_csv(caminho_download + emp_02, sep=';',index=False)
    try:
        emp_03 = pd.read_csv(caminho_download + emp_03,    sep=';',index_col=False,
                                  dtype=schema)
    except:
        df_vazio.to_csv(caminho_download + emp_03, sep=';',index=False)
    try:
        emp_04 = pd.read_csv(caminho_download + emp_04,  sep=';',index_col=False,
                              dtype=schema)
    except:
        df_vazio.to_csv(caminho_download + emp_04, sep=';',index=False)


    emp_01 = pd.read_csv(caminho_download + emp_01, sep=';',index_col=False,
                              dtype=schema,encoding='ansi')
    emp_02 = pd.read_csv(caminho_download + emp_02, sep=';',index_col=False,
                              dtype=schema,encoding='ansi')
    emp_03 = pd.read_csv(caminho_download + emp_03,    sep=';',index_col=False,
                                  dtype=schema,encoding='ansi')
    emp_04 = pd.read_csv(caminho_download + emp_04,  sep=';',index_col=False,
                              dtype=schema,encoding='ansi')


    arquivao_unico = pd.concat([emp_01, emp_02, emp_03, emp_04])

    condicao = [
        (arquivo_unico['Avaliação Reclamação'] == ''),
        (arquivo_unico['Avaliação Reclamação'] == 'Não Avaliada') | (
                arquivo_unico['Avaliação Reclamação'] == 'Resolvida'),
        (arquivo_unico['Avaliação Reclamação'] <= 'Não Resolvida')]
    retorno = [0, 1, 0]

    arquivo_unico['Avaliação Positiva'] = np.select(condicao, retorno, ).astype(int)


    arquivo_unico['Data Abertura'] = pd.to_datetime(arquivo_unico['Data Abertura'], format='%d/%m/%Y')
    arquivo_unico['Data Resposta'] = pd.to_datetime(arquivo_unico['Data Resposta'], format='%d/%m/%Y')
    arquivo_unico['Data Análise'] = pd.to_datetime(arquivo_unico['Data Análise'], format='%d/%m/%Y')
    arquivo_unico['Data Recusa'] = pd.to_datetime(arquivo_unico['Data Recusa'], format='%d/%m/%Y')
    arquivo_unico['Data Finalização'] = pd.to_datetime(arquivo_unico['Data Finalização'], format='%d/%m/%Y')
    arquivo_unico['Prazo Resposta'] = pd.to_datetime(arquivo_unico['Prazo Resposta'], format='%d/%m/%Y')

    arquivo_unico.to_csv(caminho_download + 'unico.csv',sep=';',index=False,)
    arquivo_unico = pd.read_csv(caminho_download + 'unico.csv', sep=';',index_col=False,
                                  dtype=schema,dayfirst=True)

    analiticos (arquivo_unico,emp_01,emp_02,emp_03,emp_04,writer)



def analiticos(a,b,c,d,e,z):
    arquivo_unico = a
    emp_01 = b
    emp_02 = c
    emp_03 = d
    emp_04 = e
    writer = z





#seta as colunas necessárias para o Relatório de Performance
    col_resp = ['Protocolo', 'Data Resposta', 'Nome Fantasia', 'Técnico do Último Trâmite', 'Respondida']

#Performance Individual (Geral)
#Parâmetros
    calculo_respostas = arquivo_unico[col_resp]
    calculo_respostas = calculo_respostas.sort_values(by=['Data Resposta'], ascending=False, na_position='last')

#Relatorio
    query_resp = calculo_respostas.query('Respondida == "S"')
    group_resp = query_resp.groupby(["Data Resposta", "Técnico do Último Trâmite"], dropna=True, sort=False)[
        "Protocolo"].count()
    group_resp = group_resp.to_frame()
    group_resp.head()


#indice Corte
    colunas = ['Data Abertura', 'Protocolo', 'Data Finalização', 'Data Resposta', 'Nome Fantasia', 'Técnico do Último Trâmite',
               'Avaliação Reclamação', 'Avaliação Positiva']
    calculo_indices = arquivo_unico[colunas]
    data_corte_query = datetime.strptime(dt_corte.get(), '%d/%m/%Y')
    data_corte_query = data_corte_query - timedelta(days=1)

    data_corte_max = datetime.strptime(data_fim.get(), '%d/%m/%Y')
    data_corte_max = data_corte_max + timedelta(days=1)

    indice_corte = calculo_indices
    indice_corte = indice_corte[(pd.to_datetime(indice_corte['Data Finalização']) > data_corte_query)]
    indice_corte = indice_corte[(pd.to_datetime(indice_corte['Data Finalização']) < data_corte_max)]

    agente_quebrado = indice_corte.groupby(["Técnico do Último Trâmite"], dropna=True)[
        "Avaliação Reclamação"].count()
    agente_quebrado2 = indice_corte.groupby(["Técnico do Último Trâmite"], dropna=True)[
        "Avaliação Positiva"].sum()
    agente_quebrado = pd.merge(agente_quebrado, agente_quebrado2, how='inner',
                         on=['Técnico do Último Trâmite'])
    agente_quebrado['% Solução'] = ((agente_quebrado['Avaliação Positiva'] / agente_quebrado['Avaliação Reclamação']) * 100)
    agente_quebrado['% Solução'] = agente_quebrado['% Solução'].map('{:.2f}%'.format)



    agente_pos_empresa = indice_corte.groupby(["Técnico do Último Trâmite", "Nome Fantasia"], dropna=True)[
        "Avaliação Reclamação"].count()
    agente_pos_empresa2 = indice_corte.groupby(["Técnico do Último Trâmite", "Nome Fantasia"], dropna=True)[
        "Avaliação Positiva"].sum()
    agente_pos_empresa = pd.merge(agente_pos_empresa, agente_pos_empresa2, how='inner',
                               on=['Técnico do Último Trâmite',"Nome Fantasia"])
    agente_pos_empresa['% Solução'] = (
                (agente_pos_empresa['Avaliação Positiva'] / agente_pos_empresa['Avaliação Reclamação']) * 100)
    agente_pos_empresa['% Solução'] = agente_pos_empresa['% Solução'].map('{:.2f}%'.format)



#Indice 30d
    calculo_indices2 = calculo_indices.sort_values(by=['Data Abertura'], ascending=False, na_position='last')
    max_data = calculo_indices2['Data Abertura'].iloc[0]
    max_data = pd.to_datetime(max_data, format='%Y-%m-%d')
    calculo_indices2['Dif_Dias'] = (max_data - pd.to_datetime(calculo_indices['Data Finalização'])).dt.days

    calc_30d = calculo_indices2
    calc_30d = calc_30d.query('Dif_Dias < 30 & Dif_Dias >= 0 ')

    agente30d = calc_30d.groupby(["Técnico do Último Trâmite"], dropna=True)[
        "Avaliação Reclamação"].count()
    agente30d_tt = calc_30d.groupby(["Técnico do Último Trâmite"], dropna=True)[
        "Avaliação Positiva"].sum()
    agente30d = pd.merge(agente30d, agente30d_tt, how='inner',
                             on=['Técnico do Último Trâmite'])
    agente30d['% Solução'] = ((agente30d['Avaliação Positiva'] / agente30d['Avaliação Reclamação']) * 100)
    agente30d['% Solução'] = agente30d['% Solução'].map('{:.2f}%'.format)



    agente30d_emp = calc_30d.groupby(["Técnico do Último Trâmite", "Nome Fantasia"], dropna=True)[
        "Avaliação Reclamação"].count()
    agente30d_emp_tt = calc_30d.groupby(["Técnico do Último Trâmite", "Nome Fantasia"], dropna=True)[
        "Avaliação Positiva"].sum()
    agente30d_emp = pd.merge(agente30d_emp, agente30d_emp_tt, how='inner',
                              on=['Técnico do Último Trâmite', 'Nome Fantasia'])
    agente30d_emp['% Solução'] = ((agente30d_emp['Avaliação Positiva'] / agente30d_emp['Avaliação Reclamação']) * 100)
    agente30d_emp['% Solução'] = agente30d_emp['% Solução'].map('{:.2f}%'.format)


#Indice Dia
    colunas_dia = ['Data Finalização', 'Avaliação Reclamação', 'Avaliação Positiva', "Nome Fantasia"]
    base_calc_dia = arquivo_unico[colunas_dia]
    calc_dia = base_calc_dia.groupby(['Data Finalização'], dropna=True)[
        "Avaliação Positiva"].sum()
    calc_dia_tt = base_calc_dia.groupby(['Data Finalização'], dropna=True)[
        "Avaliação Reclamação"].count()
    calc_dia = pd.merge(calc_dia, calc_dia_tt, how='inner',
                             on=['Data Finalização'])
    calc_dia['% Solução'] = ((calc_dia['Avaliação Positiva'] / calc_dia['Avaliação Reclamação']) * 100)
    calc_dia['% Solução'] = calc_dia['% Solução'].map('{:.2f}%'.format)

# Indice Dia
    calc_dia_emp = base_calc_dia.groupby(['Data Finalização', "Nome Fantasia"], dropna=True)[
        'Avaliação Positiva'].sum()
    calc_dia_emp_tt = base_calc_dia.groupby(['Data Finalização', "Nome Fantasia"], dropna=True)[
        "Avaliação Reclamação"].count()
    calc_dia_emp = pd.merge(calc_dia_emp, calc_dia_emp_tt, how='inner',
                        on=['Data Finalização', "Nome Fantasia"])
    calc_dia_emp['% Solução'] = ((calc_dia_emp['Avaliação Positiva'] / calc_dia_emp['Avaliação Reclamação']) * 100)
    calc_dia_emp['% Solução'] = calc_dia_emp['% Solução'].map('{:.2f}%'.format)

    col = ['Data Finalização', 'Avaliação Reclamação']
    base_pivot = arquivo_unico[col]
    base_pivotada = pd.pivot_table(base_pivot, values=['Avaliação Reclamação'], index='Data Finalização',
                           columns=['Avaliação Reclamação'], aggfunc=np.count_nonzero)

    salvar_bases(arquivo_unico, emp_01, emp_02, emp_03, emp_04, group_resp,
                 agente_quebrado,agente_pos_empresa,agente30d,agente30d_emp,calc_dia,calc_dia_emp, base_pivotada,writer)

def salvar_bases(a, b, c, d, e, g, i,k,l,m,n,o,p, z):
    a.to_excel(z, sheet_name='Analítico', index=False)
    b.to_excel(z, sheet_name='emp_01', index=False)
    c.to_excel(z, sheet_name='emp_02', index=False)
    d.to_excel(z, sheet_name='emp_03', index=False)
    e.to_excel(z, sheet_name='emp_04', index=False)

    g.to_excel(z, sheet_name='Respostas',index = True)
    i.to_excel(z, sheet_name='Primeiro Andamento',index = True)
    k.to_excel(z, sheet_name='InGeral Empresa data corte',index = True)
    l.to_excel(z, sheet_name='Indice 30 Dias',index = True)
    m.to_excel(z, sheet_name='Indice 30 Dias Por Empresa',index = True)
    n.to_excel(z, sheet_name='Indice por Dia', index=True)
    o.to_excel(z, sheet_name='Indice por Dia Por Empresa', index=True)
    p.to_excel(z, sheet_name='Status de Avaliação', index=True)

    z.save()
    excluir()

def excluir():
    exclusao = ['EMP_01.csv', 'EMP_02.csv', 'EMP_03.csv', 'EMP_04.csv','unico.csv']
    caminho_download = r'C:/Users/{}/Downloads/'.format(user)
    for excluir in exclusao:
        os.remove('{}{}'.format(caminho_download, excluir))

    messagebox.showinfo("Nibbler", """✦✦ O Nibbler executou a tarefa com sucesso ✦✦\n✦✦ Até a Próxima ✦✦""")

janela = tk.Tk()
janela.title("Nibbler")
janela.rowconfigure(0,weight=1)
janela.columnconfigure([0,1],weight=1)


#Dim da Janela
Larg = 270
Alt = 350

#Resol. da Tela
Larg_tela = janela.winfo_screenwidth()
Alt_tela = janela.winfo_screenheight()

#Posição Widget
posx = Larg_tela/2 - Larg/2
posy = Alt_tela/2 - Alt/2


# Define Posição // Não deixa redimensionar // Coloca o Ícone do App // Cor de Fundo
janela.geometry('%dx%d+%d+%d' % (Larg,Alt,posx,posy))
janela.resizable(False,False)
janela.iconbitmap(r"C:\Nibbler\Nibbler.ico")
janela['bg'] = '#F0DFAD'


#Titulos
boas_vindas = tk.Label(text = """Olá, Bem Vindo ao Nibbler\n \n   O Robô de Análises do Consumidor.gov   """,fg='#73bd35',bg = '#1D2666')
login_plataforma = tk.Label(text = "Login .Gov""",bg = '#F0DFAD')
senha_plataforma = tk.Label(text = "Senha .Gov",bg = '#F0DFAD')
dt_inicio = tk.Label(text = "Data Inicio",bg = '#F0DFAD')
dt_fim = tk.Label(text = "Data Fim",bg = '#F0DFAD')
data_corte = tk.Label(text = """  Data 
de Corte""",bg = '#F0DFAD')
empresas = tk.Label(text = "Empresas",bg = '#F0DFAD',font="-weight bold -size 10")


#Botões
iniciar = tk.Button(text='Iniciar',command=checkbox,activebackground = '#72BF44',activeforeground = 'white',bg = '#273593',fg = 'White')
ajuda = tk.Button(text='Ajuda',activebackground = '#72BF44',activeforeground = 'white',bg = '#273593',fg = 'White')

#Campos Editaveis
login = tk.Entry()
senha = tk.Entry(show="*")
data_inicio = DateEntry(year=ano,
                     date_pattern="dd/mm/yyyy",
                     firstweekday="sunday",
                     background="#1D2666",
                     foreground="#73bd35",
                     disabledforeground="#F0DFAD",
                     normalbackground="#F0DFAD",
                     weekendbackground="#F0DFAD",
                     disabledbackground="F0DFAD",)
data_fim = DateEntry(year=ano,
                     date_pattern="dd/mm/yyyy",
                     firstweekday="sunday",
                     background="#1D2666",
                     foreground="#73bd35",
                     disabledforeground="#F0DFAD",
                     normalbackground="#F0DFAD",
                     weekendbackground="#F0DFAD",
                     disabledbackground="F0DFAD",)
dt_corte = DateEntry(year=ano,
                     date_pattern="dd/mm/yyyy",
                     firstweekday="sunday",
                     background="#1D2666",
                     foreground="#73bd35",
                     disabledforeground="#F0DFAD",
                     normalbackground="#F0DFAD",
                     weekendbackground="#F0DFAD",
                     disabledbackground="F0DFAD",)

#Espaços
espaco_1 = tk.Label(text = "",bg = '#F0DFAD')
espaco_2 = tk.Label(text = "",bg = '#F0DFAD')
espaco_3 = tk.Label(text = "",bg = '#F0DFAD')
espaco_4 = tk.Label(text = "",bg = '#F0DFAD')
espaco_5 = tk.Label(text = "",bg = '#F0DFAD')

#Checkbox "Empresa"
check_01 = tk.BooleanVar()
check_01_text = tk.Checkbutton(janela, text='EMPRESA_1', var=check_01,bg='#F0DFAD')
check_02 = tk.BooleanVar()
check_02_text = tk.Checkbutton(janela, text='EMPRESA_2', var=check_02,bg='#F0DFAD')
check_03 = tk.BooleanVar()
check_03_text = tk.Checkbutton(janela, text='EMPRESA_3', var=check_03,bg='#F0DFAD')
check_04 = tk.BooleanVar()
check_04_text = tk.Checkbutton(janela, text='EMPRESA_4', var=check_04,bg='#F0DFAD')


#Pack ("Print" dos Campos acima)
boas_vindas.grid(row = 1, column = 0,columnspan = 2)
login.grid(row = 3, column = 1)
login_plataforma.grid(row = 3, column = 0)
senha.grid(row = 6, column = 1)
senha_plataforma.grid(row = 6, column = 0)
dt_inicio.grid(row = 8, column = 0)
dt_fim.grid(row = 8, column = 1)
data_inicio.grid(row = 9, column = 0)
data_fim.grid(row = 9, column = 1)
data_corte.grid(row = 11, column = 0)
dt_corte.grid(row = 11, column = 1)
empresas.grid(row = 14,columnspan = 2, sticky = 'NSEW')
iniciar.grid(row = 18, column = 0,columnspan = 1)
ajuda.grid(row = 18, column = 1,columnspan = 2)

#Pack Espaços
espaco_1.grid(row = 2, column = 1)
espaco_2.grid(row = 7, column = 1,columnspan = 2)
espaco_3.grid(row = 16, column = 1,columnspan = 2)
espaco_4.grid(row = 19, column = 1,columnspan = 2)
espaco_5.grid(row = 17, column = 1,columnspan = 2)

#Pack Checkbox
check_01_text.grid(row=15,column=1,columnspan = 1 ,sticky = 'W' )
check_02_text.grid(row=15,column=0,columnspan = 1 ,sticky = 'E' )
check_03_text.grid(row=16,column=0,columnspan = 1 ,sticky = 'E' )
check_04_text.grid(row=16,column=1,columnspan = 1 ,sticky = 'W' )


janela.mainloop()
