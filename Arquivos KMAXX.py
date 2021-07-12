from PySimpleGUI.PySimpleGUI import WINDOW_CLOSED
import pyexcel as p
import shutil as st
from datetime import date
import PySimpleGUI as sg
from time import sleep

dia = date.today().day
mes = date.today().month
ano = date.today().year

listerror = []


def txtGen(txt):
    layoutGen = [
        [sg.Text(f'{txt}')],
        [sg.Ok()]
    ]
    janelaGen = sg.Window('ATENÇÃO', layout=layoutGen)
    janelaGen.read()


def errArq(arquivo):
    layoutArq = [
        [sg.Text(f'ERRO NO ARQUIVO {arquivo}')],
        [sg.Ok()]
    ]
    janelaArq = sg.Window('ERRO ARQUIVO', layout=layoutArq)
    janelaArq.read()


def updtBar(x):
    progress_bar.UpdateBar(x, 6)
    sleep(.5)


layoutBarp = [[sg.Text('Progresso')],
              [sg.ProgressBar(1, orientation='h', size=(20, 20), key='progress')],
              ]

layout = [
    [sg.Text('Conversão dos arquivos em EXCEL.XLS para EXCEL.XLSX')],
    [sg.Button('Ok', pad=(110, 0)), sg.Button('Quit', pad=(110, 0))]
]

janela1 = sg.Window('ARQUIVOS EXCEL KMAXX', layout=layout)

resposta, values = janela1.read()

if resposta == 'Ok':
    janela1.close()
    unidade = 1
    windowBar = sg.Window('PROGRESSO CONVERSAO', layout=layoutBarp).Finalize()
    progress_bar = windowBar.FindElement('progress')

    try:
        updtBar(0)
        p.save_book_as(file_name=f'Z:\Kmaxx\{unidade}.xls',
                       dest_file_name=f'Z:\Kmaxx\SP Leste Matriz\{ano}\XLSX\matriz_{dia}{mes}{ano}.xlsx')
        st.move(f'Z:\Kmaxx\{unidade}.xls', f'Z:\Kmaxx\SP Leste Matriz\{ano}\XLS\matriz_{dia}{mes}{ano}.xls')
    except:
        listerror.append('MATRIZ.xls')
    updtBar(1)

    unidade = 2
    try:
        p.save_book_as(file_name=f'Z:\Kmaxx\{unidade}.xls',
                       dest_file_name=f'Z:\Kmaxx\PR\{ano}\XLSX\pr_{dia}{mes}{ano}.xlsx')
        st.move(f'Z:\Kmaxx\{unidade}.xls', f'Z:\Kmaxx\PR\{ano}\XLS\pr_{dia}{mes}{ano}.xls')
    except:
        listerror.append('PR.xls')
    updtBar(2)
    unidade = 4
    try:
        p.save_book_as(file_name=f'Z:\Kmaxx\{unidade}.xls',
                       dest_file_name=f'Z:\Kmaxx\Sorocaba\{ano}\XLSX\sorocaba_{dia}{mes}{ano}.xlsx')
        st.move(f'Z:\Kmaxx\{unidade}.xls', f'Z:\Kmaxx\Sorocaba\{ano}\XLS\sorocaba_{dia}{mes}{ano}.xls')
    except:
        listerror.append('SOROCABA.xls')
    updtBar(3)
    unidade = 5
    try:
        p.save_book_as(file_name=f'Z:\Kmaxx\{unidade}.xls',
                       dest_file_name=f'Z:\Kmaxx\SJC\{ano}\XLSX\sjc_{dia}{mes}{ano}.xlsx')
        st.move(f'Z:\Kmaxx\{unidade}.xls', f'Z:\Kmaxx\SJC\{ano}\XLS\sjc_{dia}{mes}{ano}.xls')
    except:
        listerror.append('SJC.xls')
    updtBar(4)
    unidade = 7
    try:
        p.save_book_as(file_name=f'Z:\Kmaxx\{unidade}.xls',
                       dest_file_name='Z:\Kmaxx\SP Sul\{}\XLSX\sul_{}{}{}.xlsx'.format(ano, dia, mes, ano))
        st.move(f'Z:\Kmaxx\{unidade}.xls', f'Z:\Kmaxx\SP Sul\{ano}\XLS\sul_{dia}{mes}{ano}.xls')
    except:
        listerror.append('SUL.xls')
    updtBar(5)
    unidade = 8
    try:
        p.save_book_as(file_name=f'Z:\Kmaxx\{unidade}.xls',
                       dest_file_name=f'Z:\Kmaxx\SJRP\{ano}\XLSX\sjrp_{dia}{mes}{ano}.xlsx')
        st.move(f'Z:\Kmaxx\{unidade}.xls', f'Z:\Kmaxx\SJRP\{ano}\XLS\sjrp_{dia}{mes}{ano}.xls')
    except:
        listerror.append('SJRP')
    updtBar(6)

    windowBar.close()

    if len(listerror) >= 1:
        txtGen(f'ERRO NO(S) ITEM(NS): {listerror}\nFavor verificar se o arquivo está no caminho Z:\Kmaxx\ NUMERO_UNIDADE.XLS')

    else:
        txtGen('TAREFA CONCLUÍDA COM ÊXITO EM TODOS OS ARQUIVOS.')

elif resposta == WINDOW_CLOSED or resposta == 'Quit':
    janela1.close()
