import xlwings as xw  # para chamar o excel
import tkinter as tk  # para criar a janela
import tkinter.messagebox as tkm  # para criar a janela de alerta
import keyboard as kb  # comando ESC para fechar a janela


# ---------------------------------------------------------------------------------------------------------------------
def salva():
    try:
        bd.save('bd.xlsx')  # salva a planilha
    except Exception as er:
        print(er)


# ---------------------------------------------------------------------------------------------------------------------
def teste_vazio():  # testa se os campos estão preenchidos
    if nome_var.get() != '' and cel_var.get() != '':
        return False
    else:
        return True


# ---------------------------------------------------------------------------------------------------------------------
def teste_num():  # testa se o número é int
    if cel_var.get().isdigit():
        return False
    else:
        return True


# ---------------------------------------------------------------------------------------------------------------------
def busca_registro():  # busca o registro
    for x in range(1, bd.sheets(1).used_range.last_cell.row):  # laço de repetição
        if nome_var.get().upper() == bd.sheets(1).range('A' + str(x + 1)).value:  # teste se exite
            return x + 1  # retornos da posição caso já haja registro


# ---------------------------------------------------------------------------------------------------------------------
def mostrar_dados():  # mostra os dados na caixa de alerta
    acumulado = ''
    for x in range(1, bd.sheets(1).used_range.last_cell.row):  # laço de repetição
        if bd.sheets(1).range('A' + str(x + 1)).value:  # testa se registro está vázio
            nome = str(bd.sheets(1).range('A' + str(x + 1)).value)  # recebe o nome
            celular = int(bd.sheets(1).range('B' + str(x + 1)).value)  # recebe o celular
            acumulado += nome + ', ' + str(celular) + '\n'  # acumula nome e celular

    if acumulado == '':  # testa se há registros no banco
        tkm.showinfo('Dados', 'Vazio!')
    else:
        tkm.showinfo('Dados', 'Nome: Celular: \n\n' + acumulado)


# ---------------------------------------------------------------------------------------------------------------------
def adicionar():
    if busca_registro():  # testa se registro já existe
        tkm.showinfo('Duplicado', 'Nome já cadastrado!')
    else:
        if teste_vazio():  # testa se as celulas estão preenchidos
            tkm.showinfo('Vázio', 'Preencha todos os campos!')
        else:
            if teste_num():  # testa se a celula do celular é int
                tkm.showinfo('Números', 'Campo \'Celular\' devem ter apenas números!')
            else:
                # adiciona registros na última linha do Excel
                bd.sheets(1).range('A' + str(bd.sheets(1).used_range.last_cell.row + 1)).value = nome_var.get().upper()
                bd.sheets(1).range('B' + str(bd.sheets(1).used_range.last_cell.row)).value = cel_var.get()

                nome_var.set('')  # esvazia as celula no programa
                cel_var.set('')  # esvazia as celula no programa
                nome_in.focus()  # foca na primeira celula
                salva()   # salva a planilha


# ---------------------------------------------------------------------------------------------------------------------
def deletar():
    if busca_registro():  # busca a posição do registro caso exista
        bd.sheets(1).range('A' + str(busca_registro()) + ':B' + str(busca_registro())).value = ''  # esvazia a linha

        nome_var.set('')  # esvazia as celula no programa
        cel_var.set('')  # esvazia as celula no programa
        nome_in.focus()  # foca na primeira celula
        salva()  # salva a planilha

    else:
        tkm.showinfo('Vázio', 'Registro não encontrado!')


# ---------------------------------------------------------------------------------------------------------------------
try:  # testando se o arquivo existe
    bd = xw.Book('bd.xlsx')  # abrindo arquivo já criando
except FileNotFoundError:
    bd = xw.Book()  # criando novo arquivo

    # adicionando nome das colunas
    bd.sheets(1).range("A1").value = 'Nome'
    bd.sheets(1).range("B1").value = 'Celular'
# ---------------------------------------------------------------------------------------------------------------------
# criação e configuração da janela
janela = tk.Tk()
janela.resizable(width=False, height=False)
janela.title('Excel')
# tamanho: 325X250, posição: calula e posiciona
janela.geometry("%dx%d%d%d" % (325, 200, float(325 / 2 - janela.winfo_screenwidth() / 2),
                               float(200 / 2 - janela.winfo_screenheight() / 2)))
# ---------------------------------------------------------------------------------------------------------------------
url_texto = tk.Label(janela, text='Agenda', font=('', 12))  # título da janela
url_texto.place(x=10, y=10)  # posição do objeto
# ---------------------------------------------------------------------------------------------------------------------
nome_txt = tk.Label(janela, text='Nome', font=('', 10))  # objeto do tkinter
nome_txt.place(x=10, y=42)  # posição do objeto

nome_var = tk.StringVar()  # objeto string do tkinter
nome_in = tk.Entry(janela, textvariable=nome_var)  # celúla de entrada do tkinter
nome_in.focus()  # foca na celula no início do programa
nome_in.place(x=100, y=44, width=200)  # posição do objeto
# ---------------------------------------------------------------------------------------------------------------------
cel_txt = tk.Label(janela, text='Celular', font=('', 10))  # objeto do tkinter
cel_txt.place(x=10, y=72)  # posição do objeto

cel_var = tk.StringVar()  # objeto string do tkinter
cel_in = tk.Entry(janela, textvariable=cel_var)  # celúla de entrada do tkinter
cel_in.place(x=100, y=74, width=200)  # posição do objeto
# ---------------------------------------------------------------------------------------------------------------------
mostrar_btn = tk.Button(janela, text='Mostrar', command=lambda: mostrar_dados())
mostrar_btn.place(x=10, y=104, width=300)
# ---------------------------------------------------------------------------------------------------------------------
deletar_btn = tk.Button(janela, text='Deletar', command=lambda: deletar())
deletar_btn.place(x=10, y=134, width=300)
# ---------------------------------------------------------------------------------------------------------------------
adicionar_btn = tk.Button(janela, text='Adicionar', command=lambda: adicionar())
adicionar_btn.place(x=10, y=164, width=300)
# ---------------------------------------------------------------------------------------------------------------------
kb.on_press_key('ESC', lambda _: janela.destroy())  # comando para fechar a janela
# ---------------------------------------------------------------------------------------------------------------------
janela.mainloop()  # mantem a janela aberta
# ---------------------------------------------------------------------------------------------------------------------
try:  # caso o arquivo do Excel seja fechado antes do programa
    bd.close()
except Exception as e:
    print(e)
