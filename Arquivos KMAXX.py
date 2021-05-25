from PySimpleGUI.PySimpleGUI import WINDOW_CLOSED
import pyexcel as p
import shutil as st
import datetime
import PySimpleGUI as sg

layout = [
    [sg.Text('Conversão dos arquivos em EXCEL.XLS para EXCEL.XLSX')],
    [sg.Button('Ok', pad=(110, 0)), sg.Button('Quit', pad=(110, 0))]
]

janela1 = sg.Window('ARQUIVOS EXCEL KMAXX', layout=layout)

self, values = janela1.read()

if self == 'Ok':

    try:

        dia = datetime.date.today().day
        mes = datetime.date.today().month
        ano = datetime.date.today().year


        #LESTE
        p.save_book_as(file_name='Z:/Kmaxx/1- SP Leste Matriz/matriz.xls',
                    dest_file_name='Z:/Kmaxx/1- SP Leste Matriz/{}/XLSX/matriz_{}{}{}.xlsx'.format(ano,dia,mes,ano))



        #PR
        p.save_book_as(file_name='Z:/Kmaxx/2- PR/pr.xls',
                    dest_file_name='Z:/Kmaxx/2- PR/{}/XLSX/pr_{}{}{}.xlsx'.format(ano,dia,mes,ano))


        #SOR
        p.save_book_as(file_name='Z:/Kmaxx/4- Sorocaba/sorocaba.xls',
                    dest_file_name='Z:/Kmaxx/4- Sorocaba/{}/XLSX/sorocaba_{}{}{}.xlsx'.format(ano,dia,mes,ano))


        #SJC
        p.save_book_as(file_name='Z:/Kmaxx/5- SJC/sjc.xls',
                    dest_file_name='Z:/Kmaxx/5- SJC/{}/XLSX/sjc_{}{}{}.xlsx'.format(ano,dia,mes,ano))


        #SUL
        p.save_book_as(file_name='Z:/Kmaxx/7- SP Sul/sul.xls',
                    dest_file_name='Z:/Kmaxx/7- SP Sul/{}/XLSX/sul_{}{}{}.xlsx'.format(ano,dia,mes,ano))


        #SJRP
        p.save_book_as(file_name='Z:/Kmaxx/8- SJRP/sjrp.xls',
                    dest_file_name='Z:/Kmaxx/8- SJRP/{}/XLSX/sjrp_{}{}{}.xlsx'.format(ano,dia,mes,ano))

        st.move('Z:/Kmaxx/1- SP Leste Matriz/matriz.xls', 'Z:/Kmaxx/1- SP Leste Matriz/{}/XLS/matriz_{}{}{}.xls'.format(ano,dia,mes,ano))
        st.move('Z:/Kmaxx/2- PR/pr.xls', 'Z:/Kmaxx/2- PR/{}/XLS/pr_{}{}{}.xls'.format(ano,dia,mes,ano))
        st.move('Z:/Kmaxx/4- Sorocaba/sorocaba.xls', 'Z:/Kmaxx/4- Sorocaba/{}/XLS/sorocaba_{}{}{}.xls'.format(ano,dia,mes,ano))
        st.move('Z:/Kmaxx/5- SJC/sjc.xls', 'Z:/Kmaxx/5- SJC/{}/XLS/sjc_{}{}{}.xls'.format(ano,dia,mes,ano))
        st.move('Z:/Kmaxx/7- SP Sul/sul.xls', 'Z:/Kmaxx/7- SP Sul/{}/XLS/sul_{}{}{}.xls'.format(ano,dia,mes,ano))
        st.move('Z:/Kmaxx/8- SJRP/sjrp.xls', 'Z:/Kmaxx/8- SJRP/{}/XLS/sjrp_{}{}{}.xls'.format(ano,dia,mes,ano))

        janela1.close()
        
    except:
        layoutErr = [
            [sg.Text('Um erro ocorreu. Verifique o nome dos arquivos e tente novamente.')],
            [sg.Text('matriz.xls/pr.xls/sorocaba.xls/sjc.xls/sul.xls/sjrp.xls')],
            [sg.Button('Ok')]
        ]

        janelaErr = sg.Window('ERRO NOS ARQUIVOS', layout=layoutErr)

        self, values = janelaErr.read()

        janelaErr.close()

    else:
        layoutSuc = [
            [sg.Text('Os arquivos foram convertidos com êxito!')],
            [sg.Button('Ok')]
        ]

        janelaSuc = sg.Window('Arquivos KMAXX', layout=layoutSuc)

        self, values = janelaSuc.read()

        janelaSuc.close()

elif self == WINDOW_CLOSED or self == 'Quit':
        janela1.close()
        


