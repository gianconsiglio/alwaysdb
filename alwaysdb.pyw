import pyodbc as sqlserver
import time
import os
import os.path
from datetime import datetime
import PySimpleGUI as sg
import wmi
import sys
#versão 1.2.1
#DESENVOLVIDO POR GIANPIETRO CONSIGLIO
qtd = 0
c = wmi.WMI()
lista_processos = []
for process in c.Win32_Process ():
    if process.Name == 'alwaysdb.exe':
        qtd += 1

if qtd > 2:
    sg.popup('Programa já está sendo executado!')
    sys.exit()


def tela_inicial():
    sg.theme('Reddit')
    layout = [
        [sg.Text('alwaysdb versão 1.2.1')],
        [sg.Text('lite and smart!')],
    ]
    return sg.Window('',layout = layout, finalize=True,font='Verdana',text_justification='c',no_titlebar=True)

janela = tela_inicial()    
janela.refresh()

 
def backup_diario(backup):  

    try:
        conexao = sqlserver.connect('DRIVER={SQL Server};SERVER='+servidor+';DATABASE='+backup+';UID='+usuario+';PWD='+ senha)
        cursor = conexao.cursor()

    except Exception as erro:
        send_to_txt(erro)
        pass

    else:
        day = time.strftime('%d')
        month = time.strftime('%m')
        year = time.strftime('%Y')
        folder = f'{year}{month}{day} alwaysdb' 
        try:
            os.chdir(f'{diario}') 
        except Exception as erro:
            send_to_txt(erro)
            pass
        else:
            try:
                os.mkdir(folder)
            except:
                pass

            try:
                conexao.autocommit = True
                cursor.execute(F"BACKUP DATABASE {backup} TO DISK = '{diario}{folder}\\{backup}.bak' WITH INIT")  
  
            except Exception as erro:
                send_to_txt(erro)
                pass

      

def backup_semanal(backup):
  
    try:
        conexao = sqlserver.connect('DRIVER={SQL Server};SERVER='+servidor+';DATABASE='+backup+';UID='+usuario+';PWD='+ senha)
        cursor = conexao.cursor()

    except Exception as erro:
        send_to_txt(erro) 
        pass

    else:
        day = time.strftime('%d')
        month = time.strftime('%m')
        year = time.strftime('%Y')
        folder = f'{year}{month}{day} alwaysdb' 
        try:
            os.chdir(f'{semanal}') 
        except Exception as erro:
            send_to_txt(erro)
            pass
        else:
            try:
                os.mkdir(folder)
            except:
                pass

            try:
                conexao.autocommit = True
                cursor.execute(F"BACKUP DATABASE {backup} TO DISK = '{semanal}{folder}\\{backup}.bak' WITH INIT")   
  
            except Exception as erro:
                send_to_txt(erro)
                pass

def backup_nuvem(backup):
    try:
        conexao = sqlserver.connect('DRIVER={SQL Server};SERVER='+servidor+';DATABASE='+backup+';UID='+usuario+';PWD='+ senha)
        cursor = conexao.cursor()
        
    except Exception as erro:
        send_to_txt(erro)
        pass

    else:

        try:
            conexao.autocommit = True
            cursor.execute(F"BACKUP DATABASE {backup} TO DISK = '{nuvem}{backup}.bak' WITH INIT")  
  
        except Exception as erro:
            send_to_txt(erro)
            pass


def main():
    global database1,database2,database3,database4,database5,database6,diario,usuario,senha,nuvem,semanal,servidor,hora,minuto,semana,cfg
    while True: 
        x = verifica_discos()
        try:
            os.chdir(f'{x[0]}:\\alwaysdb\\cfgs\\')

        except Exception as erro:
            send_to_txt(erro)
            time.sleep(5)
            continue

        try:
            cfg = open('backup.txt')

        except Exception as erro:    
            send_to_txt(erro)
            time.sleep(5)
            continue

        else:
            conteudo = []
            try:
                itens = cfg.readlines()
                for x in itens:
                    conteudo.append(x)
               
                usuario = conteudo[0]
                usuario = str(usuario).strip() 
                senha = conteudo[1]   
                senha = str(senha).strip()
                servidor = conteudo[2]   
                servidor = str(servidor).strip()
                semana = conteudo[3] 
                semana = str(semana).strip()
                diario = conteudo[4] 
                diario = str(diario).strip()  
                nuvem = conteudo[5]   
                nuvem = str(nuvem).strip()
                semanal = conteudo[6]   
                semanal = str(semanal).strip()
                hora = conteudo[7]  
                hora = str(hora).strip()
                minuto = conteudo[8]   
                minuto = str(minuto).strip()
                cfg = conteudo[9]   
                cfg = str(cfg).strip()
                database1 = conteudo[10] 
                database1 = str(database1).strip()
                database2 = conteudo[11] 
                database2 = str(database2).strip()  
                database3 = conteudo[12]  
                database3 = str(database3).strip()
                database4 = conteudo[13] 
                database4 = str(database4).strip()
                database5 = conteudo[14] 
                database5 = str(database5).strip()
                database6 = conteudo[15] 
                database6 = str(database6).strip()

            except Exception as erro:
                send_to_txt(erro)
                time.sleep(5)
                continue
        time.sleep(5)
        if hora == time.strftime('%H') and minuto == time.strftime('%M'):  
            dia_semana = datetime.today().strftime(('%A'))
            lista = [database1,database2,database3,database4,database5,database6]
            os.system('cls' if os.name == 'nt' else 'clear')
            for x in lista:
                if x == 'Vazio':
                    pass
                else:
                    backup_diario(x)
                    backup_nuvem(x)
                    if dia_semana == 'Monday':
                        dia_semana = 'Segunda-Feira'
                    elif dia_semana == 'Tuesday':
                        dia_semana = 'Terça-Feira'
                    elif dia_semana == 'Wednesday':
                        dia_semana = 'Quarta-Feira'
                    elif dia_semana == 'Thursday':
                        dia_semana = 'Quinta-Feira' 
                    elif dia_semana == 'Friday':
                        dia_semana = 'Sexta-Feira'
                    elif dia_semana == 'Saturday':
                        dia_semana = 'Sábado'
                    elif dia_semana == 'Sunday':
                        dia_semana = 'Domingo' 

                    if semana == dia_semana:
                        backup_semanal(x) 

            time.sleep(60)     

def verifica_discos():
    global discos_presentes
    discos_totais = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N',
    'O','P','Q','R','S','T','U','V','W','X','Y','Z']
    discos_presentes = []
    pos = -1
    for x in discos_totais:
        pos += 1
        try:
            os.chdir(f'{discos_totais[pos]}:\\')

        except:
            pass

        else:
            discos_presentes.append(discos_totais[pos])
            break

    return discos_presentes 


def send_to_txt(msg):
    msg = str(msg)
    try:
        os.chdir(f'{discos_presentes[0]}:\\alwaysdb\\erros do sistema\\')
    except:
        pass
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

time.sleep(2)
janela.close()
main()