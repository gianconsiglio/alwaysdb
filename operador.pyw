#IMPORTAÇÃO BIBLIOTECAS
import pyodbc as sqlserver
import PySimpleGUI as sg
import time
from threading import Thread
import os
import wmi
import webbrowser as wb
import sys
#IMPORTAÇÃO BIBLIOTECAS
#versão 1.6.7
#DESENVOLVIDO POR GIANPIETRO CONSIGLIO
def verifica_discos():
    onde_esta_meu_disco = []
    discos = ['A:\\','B:\\','C:\\','D:\\','E:\\','F:\\','G:\\','H:\\','I:\\','J:\\','K:\\','L:\\','M:\\','N:\\',
    'O:\\','P:\\','Q:\\','R:\\','S:\\','T:\\','U:\\','V:\\','W:\\','X:\\','Y:\\','Z:\\']
    for x in discos:
        try:
            os.chdir(x)
        except:
            pass
        else:
            onde_esta_meu_disco.append(x)
            break
    return onde_esta_meu_disco


def criador_de_pastas():
    x = verifica_discos()

    try:
        os.mkdir(f'{x[0]}alwaysdb')
    except:
        pass

    try:
        os.mkdir(f'{x[0]}alwaysdb\\cfgs')
    except:
        pass

    try:
        os.mkdir(f'{x[0]}alwaysdb\\erros do sistema')
    except:
        pass

def Tela_Principal():

    sg.theme('Reddit')    
    layout = [
        [sg.Button('CL',border_width=1),sg.Button('->',border_width=1),sg.Input(key='cfg_load',size=(10,10),border_width=2)],
        [sg.Text('STATUS:'),sg.Text('',key='time_out')],
        [sg.Text('Banco de dados[1]:'),sg.Push(),sg.Input(key='database1',size=(15,15),border_width=2,disabled=False)],
        [sg.Text('Banco de dados[2]:'),sg.Push(),sg.Input(key='database2',size=(15,15),border_width=2,disabled=False)],
        [sg.Text('Banco de dados[3]:'),sg.Push(),sg.Input(key='database3',size=(15,15),border_width=2,disabled=False)],
        [sg.Text('Banco de dados[4]:'),sg.Push(),sg.Input(key='database4',size=(15,15),border_width=2,disabled=False)],
        [sg.Text('Banco de dados[5]:'),sg.Push(),sg.Input(key='database5',size=(15,15),border_width=2,disabled=False)],
        [sg.Text('Banco de dados[6]:'),sg.Push(),sg.Input(key='database6',size=(15,15),border_width=2,disabled=False)],
        [sg.FolderBrowse('',button_color='red'), sg.Text('Diretório(DIÁRIO):'),sg.Push(),sg.Input(key='path',size=(15,15),border_width=2,disabled=False)],
        [sg.FolderBrowse('',button_color='red'),sg.Text('Diretório(NUVEM):'),sg.Push(),sg.Input(key='path2',size=(15,15),border_width=2,disabled=False)],
        [sg.FolderBrowse('',button_color='red'),sg.Text('Diretório(SEMANAL):'),sg.Push(),sg.Input(key='path3',size=(15,15),border_width=2,disabled=False)],
        [sg.Text('Servidor:'),sg.Push(),sg.Input(key='server',size=(15,15),border_width=2,disabled=False)],
        [sg.Text('Usuário:'),sg.Push(),sg.Input(key='username',size=(15,15),border_width=2,disabled=False)],
        [sg.Text('Senha:'),sg.Push(),sg.Input(key='password',password_char='*',size=(15,15),border_width=2,disabled=False)],
        [sg.Text('Cfg:'),sg.Push(),sg.Input(key='name_cfg',size=(15,15),border_width=2,disabled=False)],
        [sg.Text('Horário:'),sg.Push(),sg.Combo(['00','01','02', '03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23'],default_value='',key='hour',readonly=True,disabled=False),sg.Text(':'),sg.Combo(['00','05','10', '15','20','25','30','35','40','45','50','55'],default_value='',key='minute',readonly=True,disabled=False)],
        [sg.Text('Semana:'),sg.Push(),sg.Combo(['Segunda-Feira','Terça-Feira','Quarta-Feira','Quinta-Feira','Sexta-Feira','Sábado','Domingo','Nunca'],default_value='',key='weekend',readonly=True,disabled=False)],
        [sg.Button('SALVAR',border_width=1,disabled=False),sg.Push(),sg.Button('BACKUP',border_width=1,disabled=True),sg.Button('SUPORTE',border_width=1,disabled=False)],
        [sg.Button('VERIFICAR alwaysdb',border_width=1),sg.Button('FINALIZAR alwaysdb',border_width=1)],
    ]
    return sg.Window('alwaysdb operador 1.6.7',layout=layout,finalize=True,font='Verdana',icon='database.ico',text_justification='c')  
def tela_loading():
    sg.theme('Reddit')
    layout = [
        [sg.Text('AGUARDE...')]
    ]
    return sg.Window('alwaysdb operador 1.6.7',layout=layout,finalize=True,font='Verdana',icon='database.ico',text_justification='c',no_titlebar=True)

def send_to_txt(msg):
    msg = str(msg)
    x = verifica_discos()
    try:
        os.chdir(f"{x[0]}:\\alwaysdb\\erros do sistema\\") 

    except:
        sg.popup('Erro não foi enviado pois diretório erros do sistema não foi encontrado!')
        
    else:
        day = str(time.strftime('%d'))
        month = str(time.strftime('%m'))
        year = str(time.strftime('%Y'))
        hour = str(time.strftime('%H'))
        minute = str(time.strftime('%M'))
        seconds = str(time.strftime('%S'))
        exception = f'{day}-{month}-{year}_{hour}-{minute}-{seconds}'
        exceptions = open(f'{exception}.txt','w')
        exceptions.write(msg)
        exceptions.close()
        time.sleep(1)
        del day,month,year,hour,minute,seconds,exception,exceptions


def backup_diario(tempo,backup):
    try:
        window['time_out'].update(F'CONECTANDO {values[backup]} DIÁRIO')
        conexao = sqlserver.connect('DRIVER={SQL Server};SERVER='+values['server']+';DATABASE='+values[f'{backup}']+';UID='+values['username']+';PWD='+ values['password'])
        cursor = conexao.cursor() 

    except Exception as erro:
        window['time_out'].update(F'FALHA AO CONECTAR {values[backup]} DIÁRIO')
        send_to_txt(erro)
        time.sleep(tempo)
        pass

    else:
        window['time_out'].update(F'{values[backup]} DIÁRIO CONECTADO')
        time.sleep(tempo)
        day = time.strftime('%d')
        month = time.strftime('%m')
        year = time.strftime('%Y')
        folder = f'{year}{month}{day} alwaysdb' 

        try:
            os.chdir(values['path']) 

        except Exception as erro:
            send_to_txt(erro)
            window['time_out'].update(f'{values[backup]} DIÁRIO FALHOU')
            pass

        else:

            try:
                os.mkdir(folder)

            except:
                pass
               
            try:
                window['time_out'].update(f'FAZENDO {values[backup]} DIÁRIO')
                conexao.autocommit = True
                cursor.execute(F"BACKUP DATABASE [{values[f'{backup}']}] TO DISK = '{values['path'] + folder}//{values[f'{backup}']}.bak' WITH INIT")  
                window['time_out'].update(f'{values[backup]} DIÁRIO OK')       
                time.sleep(x)     
                
            except Exception as erro:
                send_to_txt(erro)
                window['time_out'].update(f'{values[backup]} DIÁRIO FALHOU')
                time.sleep(tempo)
                pass

           


def backup_nuvem(tempo,backup):
    try:
        window['time_out'].update(F'CONECTANDO {values[backup]} NUVEM')
        conexao = sqlserver.connect('DRIVER={SQL Server};SERVER='+values['server']+';DATABASE='+values[f'{backup}']+';UID='+values['username']+';PWD='+ values['password'])
        cursor = conexao.cursor() 

    except Exception as erro:
        window['time_out'].update(F'FALHA AO CONECTAR {values[backup]} NUVEM')
        send_to_txt(erro)
        time.sleep(tempo)
        pass

    else:
        window['time_out'].update(F'BANCO {values[backup]} NUVEM CONECTADO')
        time.sleep(tempo)

        try:
            window['time_out'].update(f'FAZENDO {values[backup]} NUVEM')
            conexao.autocommit = True
            cursor.execute(F"BACKUP DATABASE [{values[f'{backup}']}] TO DISK = '{values['path2']}{values[f'{backup}']}.bak' WITH INIT")  
            window['time_out'].update(f'{values[backup]} NUVEM OK')
            time.sleep(x)
           
        except Exception as erro:
            send_to_txt(erro)
            window['time_out'].update(f'{values[backup]} NUVEM FALHOU')
            time.sleep(tempo)
            pass



def backup_semanal(tempo,backup):
    try:
        window['time_out'].update(F'CONECTANDO {values[backup]} SEMANAL')
        conexao = sqlserver.connect('DRIVER={SQL Server};SERVER='+values['server']+';DATABASE='+values[f'{backup}']+';UID='+values['username']+';PWD='+ values['password'])
        cursor = conexao.cursor() 

    except Exception as erro:
        window['time_out'].update(F'FALHA AO CONECTAR {values[backup]} SEMANAL')
        send_to_txt(erro)
        time.sleep(tempo)
        pass

    else:
        window['time_out'].update(F'BANCO {values[backup]} SEMANAL CONECTADO ')
        time.sleep(tempo)
        day = time.strftime('%d')
        month = time.strftime('%m')
        year = time.strftime('%Y')
        folder = f'{year}{month}{day} alwaysdb' 

        try:
            os.chdir(values['path3']) 

        except Exception as erro:
            send_to_txt(erro)
            pass

        else:

            try:
                os.mkdir(folder)

            except:
                pass
               
            try:
                window['time_out'].update(f'FAZENDO {values[backup]} SEMANAL')
                conexao.autocommit = True
                cursor.execute(F"BACKUP DATABASE [{values[f'{backup}']}] TO DISK = '{values['path3'] + folder}//{values[f'{backup}']}.bak' WITH INIT")  
                window['time_out'].update(f'{values[backup]} SEMANAL OK')
                time.sleep(x)
            
            except Exception as erro:
               send_to_txt(erro)
               window['time_out'].update(f'{values[backup]} SEMANAL FALHOU')
               time.sleep(tempo)
               pass


def limpar():
    lista = ['username','password','server','weekend','path','path2','path3','hour','minute','name_cfg','database1','database2','database3','database4','database5','database6']
    for x in lista:    
        window[x].update('')
    del lista   
     
        
def backup_manual():
    ative_disabled(True)
    pos = 0
    loading = 0
    lista = ['database1','database2','database3','database4','database5','database6']        
    for x in lista: 
        pos += 1
        if values[f'{x}'] == 'Vazio':
            window['time_out'].update(f'DATABASE{[pos]} DESCONSIDERADO!')
            time.sleep(1)
            pass
        else:
            backup_diario(1,x)
            backup_nuvem(1,x)
            backup_semanal(1,x)

    limpar()    
    del loading
    window['time_out'].update('TAREFA FINALIZADA')
    ative_disabled(False)


def ative_disabled(info):
    bool(info)
    lista = ['username','password','server','weekend','path','path2','path3','hour','minute','name_cfg','database1','database2','database3','database4','database5','database6','BACKUP','SALVAR','SUPORTE','CL','->']
    for x in lista:
        window[x].update(disabled=info)
janela1 = tela_loading() 
janela1.refresh()
c = wmi.WMI()
qtd = 0
for process in c.Win32_Process ():
    if process.Name == 'operador.exe':
        qtd += 1

if qtd > 2:
    sg.popup('Programa já está sendo executado!')
    sys.exit()

criador_de_pastas()
janela1.close()
janela = Tela_Principal()

while True:
    window,event,values = sg.read_all_windows()
    if event == 'BACKUP':
        pos = 0
        lista = ['username','password','server','weekend','path','path2','path3','hour','minute','name_cfg','database1','database2','database3','database4','database5','database6']
        for x in lista:
            pos += 1
            if values[x] == '':
                campo = values[x]
                sg.popup(f'Campo {campo} está vazio!')
                break
        if pos >=16:
            Thread(target=backup_manual).start()  
        
    elif event == '->':
        x = verifica_discos()
        try:
            os.chdir(f'{x[0]}alwaysdb\\cfgs')
        except:
            sg.popup('Diretório cfgs não foi encontrado!')
            continue

        try:
            cfg = open(values['cfg_load'] + '.txt')
        except:
            sg.popup('Arquivo não existe!')
            continue
        else:
            content2 = []
            content = cfg.readlines()
            for x in content:
                content2.append(x)

        lista = ['username','password','server','weekend','path','path2','path3','hour','minute','name_cfg','database1','database2','database3','database4','database5','database6']
        pos = -1
        
        for x in lista:
            pos += 1 

            try:
                window[x].update(str(content2[pos].strip()))

            except Exception as erro:
                send_to_txt(erro)
                sg.popup('Erro ao carregar arquivo txt!')    
                limpar()
                break

        else:
            window['cfg_load'].update('')
            window['time_out'].update('')
            window['BACKUP'].update(disabled=False)
            del lista

    elif event == 'SALVAR':

        if len(values['name_cfg']) <= 0:
            sg.popup('Campo nome cfg vazio!')
            continue
        x = verifica_discos()    
        try:
            os.chdir(f'{x[0]}alwaysdb\\cfgs\\')

        except Exception as erro:
            sg.popup('Não foi possível encontrar diretório cfgs')
            sg.popup(erro)
            continue

        else:
            cfg = open(values['name_cfg'] + '.txt','w')

        lista = ['username','password','server','weekend','path','path2','path3','hour','minute','name_cfg','database1','database2','database3','database4','database5','database6']
        pos = 0

        try:

            for x in lista:

                if len(values[x]) <= 0:
                    nova_palavra = 'Vazio'
                    cfg.write(nova_palavra + '\n')
                     
                else:
                    cfg.write(values[x] + '\n')
                                         
        except Exception as erro:
            send_to_txt(erro)
            sg.popup('Não foi possível salvar CFG')  
            limpar()    
            continue  

        else:
            sg.popup('CFG salva com sucesso')
            cfg.close() 
            limpar()
            del lista

    elif event == 'CL':
        lista = ['username','password','server','weekend','path','path2','path3','hour','minute','name_cfg','database1','database2','database3','cfg_load','time_out','database4','database5','database6']
        for x in lista:
            window[x].update('')

    elif event == 'SUPORTE':
        wb.open('https://gianpietro-consiglio.github.io/site-alwaysdb/')

    elif event == 'VERIFICAR alwaysdb':
        c = wmi.WMI()
        lista_processos = []
        for process in c.Win32_Process ():
            lista_processos.append(process.Name)

        for x in lista_processos:
            if x == 'alwaysdb.exe':
                x = True
                break

        if x == True:
            sg.popup('CONECTADO!')
        else:
            sg.popup('NÃO CONECTADO!')      

    elif event == 'FINALIZAR alwaysdb':
        x = os.system("taskkill /f /im  alwaysdb.exe")
        if x == 128:
            sg.popup('PROCESSO NÃO ENCONTRADO!')

        elif x == 0:
            sg.popup('PROCESSO ENCERRADO!')   

    elif event == sg.WIN_CLOSED:
        break 