import customtkinter as c
import sys
from Código.Utils import *

class Tela_Principal(c.CTk):
    def __init__(self):
        super().__init__()

        c.set_appearance_mode('dark')
        self.title('Bot Whtasapp')
        self.centralizar_tela_na_direita()
        self.resizable(False, False)
        self.grid_propagate(False)
        
        self.frame_principal()
        self.area_explorador_planilha()
        self.area_explorador_img()
        self.area_mensagem()
        self.botoes_acoes()
        self.area_log()

        # Funções
        self.funcao_botoes_pesquisar = Funcao_Botoes_Pesquisar(self, self.texto_caminho_planilha, self.texto_caminho_img)
        
        self.funcao_botao_acoes = Funcao_Botoes_Acoes(self, self.texto_caminho_planilha, self.texto_caminho_img)

    def frame_principal(self):
        self.frame = c.CTkFrame(
            self,
            height=700,
            width=610,
            corner_radius=0,
            fg_color='#363f4e'
        )
        self.frame.grid(row=0, column=0, sticky='nswe')
        self.frame.grid_propagate(False)

        self.frame.grid_columnconfigure(0, weight=0)
        self.frame.grid_columnconfigure(1, weight=0)
        self.frame.grid_columnconfigure(2, weight=0)
        
    def area_explorador_planilha(self):
        label_titulo_planilha = c.CTkLabel(
            self.frame,
            text='Selecione o caminho da planilha de base de clientes:',
            font=('Arial', 16),
            anchor='center'
        )
        label_titulo_planilha.grid(row=0, column=0, sticky='w', padx=15, pady=10, columnspan=2)


        frame_planilha = c.CTkFrame(
            self.frame,
            height=40,
            width=500,
            corner_radius=10,
            border_width=2,
            border_color='black',
            fg_color='#363f4e',
            bg_color='#363f4e'
        )
        frame_planilha.grid(row=1, column=0, sticky='w', padx=10, columnspan=2)
        frame_planilha.grid_propagate(False)


        self.texto_caminho_planilha = c.CTkEntry(
            frame_planilha,
            placeholder_text='Caminho planilha...',
            width=480,
            font=('Arial', 12),
            border_color='#363f4e',
            fg_color='#363f4e',
            bg_color='#363f4e'
        )
        self.texto_caminho_planilha.grid(row=0, column=0, sticky='w', padx=10, pady=6)
        self.texto_caminho_planilha.configure(state='disable')


        self.botao_procurar_planilha = c.CTkButton(
            self.frame,
            text='PESQUISAR',
            font=('Arial', 12, 'bold'),
            height=37,
            width=30,
            anchor='center',
            cursor='hand2',

            command=lambda: self.funcao_botoes_pesquisar.abrir_explorador_planilha()
        )
        self.botao_procurar_planilha.grid(row=1, column=2, sticky='nswe')

    def area_explorador_img(self):
        label_titulo_img = c.CTkLabel(
            self.frame,
            text='Selecione o caminho da imagem:',
            font=('Arial', 16),
            anchor='center'
        )
        label_titulo_img.grid(row=3, column=0, sticky='w', padx=15, pady=10, columnspan=2)


        frame_img = c.CTkFrame(
            self.frame,
            height=40,
            width=500,
            corner_radius=10,
            border_width=2,
            border_color='black',
            fg_color='#363f4e',
            bg_color='#363f4e'
        )
        frame_img.grid(row=4, column=0, sticky='w', padx=10, columnspan=2)
        frame_img.grid_propagate(False)


        self.texto_caminho_img = c.CTkEntry(
            frame_img,
            placeholder_text='Caminho imagem...',
            width=480,
            font=('Arial', 12),
            border_color='#363f4e',
            fg_color='#363f4e',
            bg_color='#363f4e'
        )
        self.texto_caminho_img.grid(row=0, column=0, sticky='w', padx=10, pady=6)
        self.texto_caminho_img.configure(state='disable')


        self.botao_procurar_img = c.CTkButton(
            self.frame,
            text='PESQUISAR',
            font=('Arial', 12, 'bold'),
            height=37,
            width=30,
            anchor='center',
            cursor='hand2',

            command=lambda: self.funcao_botoes_pesquisar.abrir_explorador_img()
        )
        self.botao_procurar_img.grid(row=4, column=2, sticky='nswe')

    def area_mensagem(self):
        label_titulo_msg = c.CTkLabel(
            self.frame,
            text='Digite a mensagem desejada:',
            font=('Arial', 16),
            anchor='center'
        )
        label_titulo_msg.grid(row=5, column=0, sticky='w', padx=15, pady=10, columnspan=2)

        self.mensagem_digitada = c.CTkTextbox(
            self.frame,
            height=250,
            width=590,
            font=('Arial', 14)
        )
        self.mensagem_digitada.grid(row=6, column=0, sticky='n', padx=10, columnspan=4)

    def botoes_acoes(self):
        self.botao_iniciar = c.CTkButton(
            self.frame,
            text='INICIAR ENVIO',
            font=('Arial', 14, 'bold'),
            height=40,
            width=100,
            fg_color='#2E8B57',
            hover_color='#006400',
            cursor='hand2',

            command=lambda: self.funcao_botao_acoes.botao_iniciar_envio()
        )
        self.botao_iniciar.grid(row=7, column=0, sticky='w', padx=10, pady=10, columnspan=1)
        
        
        self.botao_interromper = c.CTkButton(
            self.frame,
            text='INTERROMPER',
            font=('Arial', 14, 'bold'),
            height=40,
            width=100,
            fg_color='#A52A2A',
            hover_color='#800000',
            cursor='hand2',

            command=lambda: self.funcao_botao_acoes.botao_interromper() 
        )
        self.botao_interromper.grid(row=7, column=0, sticky='w', padx=135, pady=10, columnspan=1)
        self.botao_interromper.configure(state='disable')

        self.checar = c.StringVar(value='off')
        self.opcao_img = c.CTkCheckBox(
            self.frame,
            text='Enviar mensagem com a imagem',
            font=('Arial', 14),
            variable= self.checar,
            onvalue='on',
            offvalue='off'
        )
        self.opcao_img.grid(row=7, column=0, sticky='e',padx=50, columnspan=3)

    def area_log(self):
        label_titulo_msg = c.CTkLabel(
            self.frame,
            text='Log do sistema:',
            font=('Arial', 16),
            anchor='center'
        )
        label_titulo_msg.grid(row=8, column=0, sticky='w', padx=15, pady=10, columnspan=2)


        self.logs = c.CTkTextbox(
            self.frame,
            height=110,
            width=590,
            font=('Arial', 14)
        )
        self.logs.grid(row=9, column=0, sticky='n', padx=10, columnspan=4)
        self.logs.configure(state='disable')

    def centralizar_tela_na_direita(self):
        largura_app = 610
        altura_app = 700

        largura_monitor = self.winfo_screenwidth()
        altura_monitor = self.winfo_screenheight()

        pos_x = largura_monitor - largura_app
        pos_y = (altura_monitor - altura_app) // 2

        self.geometry(f'{largura_app}x{altura_app}+{pos_x-30}+{pos_y}')

    def texto_area_log(self, texto=''):
        self.logs.configure(state='normal')
        self.logs.insert('end', f'-> {texto}\n\n')
        self.logs.see('end')
        self.logs.configure(state='disable')


if __name__=='__main__':
    if hasattr(sys, '_MEIPASS'):
        from customtkinter import set_appearance_mode

    app = Tela_Principal()
    app.mainloop()