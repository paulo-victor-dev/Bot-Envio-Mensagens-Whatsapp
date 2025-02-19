import customtkinter as c

class Tela_Auxiliar(c.CTkToplevel):
    def __init__(self, master, btn_sim_nao=False, funcao_btn_s=None, texto='', titulo='ATENÇÃO'):
        super().__init__(master)
        self.texto = texto
        self.funcao_btn_s = funcao_btn_s
        self.titulo = titulo

        c.set_appearance_mode('dark')
        self.title('AVISO')
        self.centralizar_tela(master, 400, 300)
        self.resizable(False, False)
        self.grab_set()

        if btn_sim_nao:
            self.frame_janela_btn_sim_nao()
            self.janela_btn_sim_nao()
        else:
            self.frame_janela_aviso()
            self.janela_de_aviso()
            
    def frame_janela_aviso(self):
        self.frame_principal_janela_aviso = c.CTkFrame(
            self,
            height=300,
            width=400,
            fg_color='#1d2027'
        )
        self.frame_principal_janela_aviso.grid(row=0, column=0)
        self.frame_principal_janela_aviso.grid_propagate(False)

        self.frame_principal_janela_aviso.grid_columnconfigure(0, weight=1)
        self.frame_principal_janela_aviso.grid_columnconfigure(1, weight=0)
        self.frame_principal_janela_aviso.grid_columnconfigure(2, weight=1)

        self.frame_principal_janela_aviso.grid_rowconfigure(0, weight=0)
        self.frame_principal_janela_aviso.grid_rowconfigure(1, weight=0)
        self.frame_principal_janela_aviso.grid_rowconfigure(2, weight=0)

    def janela_de_aviso(self):
        titulo = c.CTkLabel(
            self.frame_principal_janela_aviso,
            text=self.titulo,
            font=('Arial', 20, 'bold')
        )
        titulo.grid(row=0, column=1, sticky='n', pady=50)


        self.texto_janela_aviso = c.CTkLabel(
            self.frame_principal_janela_aviso,
            text=self.texto,
            font=('Arial', 14),
            anchor='center'
        )
        self.texto_janela_aviso.grid(row=1, column=1, sticky='n', padx=10)


        botão_ok = c.CTkButton(
            self.frame_principal_janela_aviso,
            text='OK',
            height=40,
            width=60,
            font=('Arial', 14),
            bg_color='#1d2027',
            cursor='hand2',

            command=lambda: self.destroy()
        )
        botão_ok.grid(row=2, column=1, pady=50)


    def frame_janela_btn_sim_nao(self):
        self.frame_principal_janela_btn_sim_nao = c.CTkFrame(
            self,
            height=300,
            width=400,
            fg_color='#1d2027'
        )
        self.frame_principal_janela_btn_sim_nao.grid(row=0, column=0)
        self.frame_principal_janela_btn_sim_nao.grid_propagate(False)

        self.frame_principal_janela_btn_sim_nao.grid_columnconfigure(0, weight=1)
        self.frame_principal_janela_btn_sim_nao.grid_columnconfigure(1, weight=0)
        self.frame_principal_janela_btn_sim_nao.grid_columnconfigure(2, weight=0)
        self.frame_principal_janela_btn_sim_nao.grid_columnconfigure(3, weight=0)
        self.frame_principal_janela_btn_sim_nao.grid_columnconfigure(4, weight=1)

        self.frame_principal_janela_btn_sim_nao.grid_rowconfigure(0, weight=0)
        self.frame_principal_janela_btn_sim_nao.grid_rowconfigure(1, weight=0)
        self.frame_principal_janela_btn_sim_nao.grid_rowconfigure(2, weight=0)
        
    def janela_btn_sim_nao(self):
        titulo = c.CTkLabel(
            self.frame_principal_janela_btn_sim_nao,
            text=self.titulo,
            font=('Arial', 20, 'bold')
        )
        titulo.grid(row=0, column=2, sticky='n', pady=50)


        self.texto_btn_s_n = c.CTkLabel(
            self.frame_principal_janela_btn_sim_nao,
            text=self.texto,
            font=('Arial', 14),
            anchor='center'
        )
        self.texto_btn_s_n.grid(row=1, column=0, sticky='n', columnspan=5)


        self.botão_sim = c.CTkButton(
            self.frame_principal_janela_btn_sim_nao,
            text='SIM',
            height=40,
            width=60,
            font=('Arial', 14),
            bg_color='#1d2027',
            anchor='center',
            cursor='hand2',

            command=lambda: self.execucao_btn_s()
        )
        self.botão_sim.grid(row=2, column=1, sticky='e', pady=50, columnspan=1)
        
        
        self.botão_nao = c.CTkButton(
            self.frame_principal_janela_btn_sim_nao,
            text='NÃO',
            height=40,
            width=60,
            font=('Arial', 14),
            bg_color='#1d2027',
            anchor='center',
            cursor='hand2',

            command=lambda: self.destroy()
        )
        self.botão_nao.grid(row=2, column=3, sticky='w', pady=50, columnspan=1)


    def centralizar_tela(self, janela_pai, largura_janela_filha, altura_janela_filha):
        largura_janela_pai = janela_pai.winfo_width()
        altura_janela_pai = janela_pai.winfo_height()
        pos_x_janela_pai = janela_pai.winfo_x()
        pos_y_janela_pai = janela_pai.winfo_y()

        pos_x_janela_filha = pos_x_janela_pai + (largura_janela_pai - largura_janela_filha) // 2
        pos_y_janela_filha = pos_y_janela_pai + (altura_janela_pai - altura_janela_filha) // 2

        self.geometry(f'{largura_janela_filha}x{altura_janela_filha}+{pos_x_janela_filha}+{pos_y_janela_filha}')

    def execucao_btn_s(self):
        self.destroy()

        self.funcao_btn_s()
