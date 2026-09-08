import os
import shutil
import json
import threading
import pandas as pd
import xlwings as xw
from openpyxl import load_workbook

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# =========================================================
# CONFIGURAÇÃO DE CORES DA INTERFACE
# =========================================================
COR_FUNDO = "#424549"          # Fundo principal
COR_CARD = "#1e2124"           # Destaque de fundo (Cards / Painéis)
COR_DESTAQUE = "#7289da"       # Cor de destaque (Botões / Ações)
COR_TEXTO = "#ffffff"          # Cor da escrita

# =========================================================
# FUNÇÕES DE CONFIGURAÇÃO JSON
# =========================================================
def escrevendoArquivo_Jason(caminho):
    with open("conf.json", 'w', encoding='utf-8') as arquivo:
        json.dump(caminho, arquivo, ensure_ascii=False, indent=4)

def LendoArquivo_Jason():
    if os.path.exists("conf.json"):
        with open("conf.json", "r", encoding='utf-8') as arquivo:
            try:
                return json.load(arquivo)
            except json.JSONDecodeError:
                return ""
    return ""

# =========================================================
# FUNÇÕES PRINCIPAIS DE PROCESSAMENTO
# =========================================================
def prenchendoPlanilha_ACC(caminho_arquivo, PN_Cliente, REV_Acc, RFQ_Acc, Projeto_Acc, ClientePlanta_Acc, VolumeAnual_Acc, DataEntrada_Acc, DataResposta_Acc, TipoItem_Acc):
    Plan_Saida = load_workbook(caminho_arquivo)
    Pagina_Saida = Plan_Saida['Plan1']
    
    Pagina_Saida['F6']  = PN_Cliente
    Pagina_Saida['Q6']  = REV_Acc
    Pagina_Saida['S3']  = RFQ_Acc
    Pagina_Saida['AD3'] = Projeto_Acc
    Pagina_Saida['AK3'] = ClientePlanta_Acc
    Pagina_Saida['F9']  = ", ".join(map(str, VolumeAnual_Acc)) if isinstance(VolumeAnual_Acc, list) else VolumeAnual_Acc
    Pagina_Saida['H12'] = DataEntrada_Acc
    Pagina_Saida['T12'] = DataResposta_Acc
    
    if TipoItem_Acc == "Novo":
        Pagina_Saida['S9'] = "X"
    elif TipoItem_Acc == "Protótipo":
        Pagina_Saida['S10'] = "X"
    elif TipoItem_Acc == "Modificação":
        Pagina_Saida['AA9'] = "X"
    else:
        Pagina_Saida['V10'] = TipoItem_Acc
        Pagina_Saida['AA10'] = "X"
                
    Plan_Saida.save(caminho_arquivo)

def PegandoDados_PQPMain(APQPSelecionada):
    Plan_Entrada = xw.Book(f'{APQPSelecionada}.xlsx')
    Pagina_Entrada = Plan_Entrada.sheets['APQP']
    
    realizado     = Pagina_Entrada.range('B4:B1000').value     
    Observacao    = Pagina_Entrada.range('C4:C1000').value    
    Cliente       = Pagina_Entrada.range('E1:E1000').value  
    ClientePlanta = Pagina_Entrada.range('F4:F1000').value  
    tipo          = Pagina_Entrada.range('G4:G1000').value
    RFQ           = Pagina_Entrada.range('H4:H1000').value    
    PNs           = Pagina_Entrada.range('I4:I1000').value   
    REV           = Pagina_Entrada.range('J4:J1000').value
    descricao     = Pagina_Entrada.range('K4:K1000').value
    comprador     = Pagina_Entrada.range('L4:L1000').value
    Desenhos3D    = Pagina_Entrada.range('M4:M1000').value
    Desenhos2D    = Pagina_Entrada.range("N4:N1000").value
    VolumeAnual   = Pagina_Entrada.range('O4:P1000').value   
    PesoPeca      = Pagina_Entrada.range('P4:P1000').value   
    QTD           = Pagina_Entrada.range('R4:R1000').value   
    Componentes   = Pagina_Entrada.range('S4:S1000').value   
    Projeto       = Pagina_Entrada.range('T4:T1000').value   
    DataEntrada   = Pagina_Entrada.range('U4:U1000').value
    DataResposta  = Pagina_Entrada.range('W4:W1000').value
    Reponsavel    = Pagina_Entrada.range('X4:X1000').value
    
    def PegarInf(ColumSelecionada):
        listaLimpa = []
        if ColumSelecionada is None:
            return []
        for Dado in ColumSelecionada:  
            if Dado is not None and str(Dado).strip() != '':                                                 
                listaLimpa.append(Dado)         
        return listaLimpa
    
    ObservacaoList    = PegarInf(Observacao)
    PNsList           = PegarInf(PNs)
    REVList           = PegarInf(REV)
    RFQList           = PegarInf(RFQ)
    ProjetoList       = PegarInf(Projeto)
    ClientePlantaList = PegarInf(ClientePlanta)
    Desenhos2DList    = PegarInf(Desenhos2D)
    Desenhos3DList    = PegarInf(Desenhos3D)
    VolumeAnualList   = PegarInf(VolumeAnual)
    DataEntradaList   = PegarInf(DataEntrada)
    DataRespostaList  = PegarInf(DataResposta)
    QTDLista          = PegarInf(QTD)
    ComponentesLista  = PegarInf(Componentes)
    ReponsavelLista   = PegarInf(Reponsavel)
    realizadoLista     = PegarInf(realizado)
    ClienteLista        = PegarInf(Cliente)
    tipoLista          = PegarInf(tipo)
    descricaoLista      = PegarInf(descricao)
    compradorLista      = PegarInf(comprador)
    PesoPecaLista       = PegarInf(PesoPeca)

    return ObservacaoList, PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, Desenhos2DList, Desenhos3DList, VolumeAnualList, DataEntradaList, DataRespostaList, QTDLista, ComponentesLista, ReponsavelLista, realizadoLista, ClienteLista, tipoLista, descricaoLista, compradorLista, PesoPecaLista

def CriandoPastas(PNs_Ciar, rev_Craidas, RFQ_criar, Projeto_criar, ClientePlanta_criar, VolumeAnual_criar, DataEntrada_criar, DataResposta_criar, TipoItem_criar, log_func):
    pastas_Segudarias = ["1-RFQ","2-Desenhos","3-Analise Critica","4-Planilha de Custo","5-CBD","6-Carta Comercial","8-Terceiros","9-Pedidos","10-FII","11-Emails","12-Custo Logistico","13-Obsoleto","14-Modificações"]
    
    if len(PNs_Ciar) > 0:
        for i in range(len(PNs_Ciar)):
            pasta_principal = str(PNs_Ciar[i])
            os.makedirs(pasta_principal, exist_ok=True)
            
            for nome_subpasta in pastas_Segudarias:
                caminho_Completo = os.path.join(pasta_principal, nome_subpasta)
                os.makedirs(caminho_Completo, exist_ok=True)

            novo_nome_excel = f'{pasta_principal}_REV.{rev_Craidas[i]}_ACC.xlsx'
            caminho_excel_copiado = os.path.join(pasta_principal, "3-Analise Critica", novo_nome_excel)
            
            shutil.copy2('ACC.xlsx', caminho_excel_copiado)
            prenchendoPlanilha_ACC(caminho_excel_copiado, PNs_Ciar[i], rev_Craidas[i], RFQ_criar[i], Projeto_criar[i], ClientePlanta_criar[i], VolumeAnual_criar[i], DataEntrada_criar[i], DataResposta_criar[i], TipoItem_criar[i])
            
            log_func(f"✓ Sucesso: Pasta '{pasta_principal}' criada e planilha preenchida!")
    else:
        log_func("⚠️ Ops! A lista de PNs está vazia. Nenhum dado encontrado.")

def PreenchendoAPQP_Exec(APQP_Input, celula_inicio, log_func):
    ObservacaoList, PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, Desenhos2DList, Desenhos3DList, VolumeAnualList, DataEntradaList, DataRespostaList, QTDList, ComponentesList, ReponsavelList, RealizadoList, ClienteList, tipoList, descricaoList, compradorList, PesoPecaList = PegandoDados_PQPMain('MainAPQP')
    
    Tamanho_list = len(PNsList)
    i = 0
    
    Plan_Entrada = xw.Book(str(APQP_Input))
    Pagina_Entrada = Plan_Entrada.sheets['APQP']
    
    while (Tamanho_list > i):    
        NumeroCelula = int(celula_inicio) + i
        
        Pagina_Entrada.range(f'B{NumeroCelula}').value = RealizadoList[i]
        Pagina_Entrada.range(f'C{NumeroCelula}').value = ObservacaoList[i]
        Pagina_Entrada.range(f'F{NumeroCelula}').value = ClienteList[i]
        Pagina_Entrada.range(f'G{NumeroCelula}').value = ClientePlantaList[i]
        Pagina_Entrada.range(f'H{NumeroCelula}').value = tipoList[i]
        Pagina_Entrada.range(f'I{NumeroCelula}').value = RFQList[i]
        Pagina_Entrada.range(f'J{NumeroCelula}').value = PNsList[i]
        Pagina_Entrada.range(f'K{NumeroCelula}').value = REVList[i]
        Pagina_Entrada.range(f'L{NumeroCelula}').value = descricaoList[i]
        Pagina_Entrada.range(f'M{NumeroCelula}').value = compradorList[i]
        Pagina_Entrada.range(f'N{NumeroCelula}').value = Desenhos2DList[i]
        Pagina_Entrada.range(f'O{NumeroCelula}').value = Desenhos3DList[i]
        Pagina_Entrada.range(f'P{NumeroCelula}').value = VolumeAnualList[i]
        Pagina_Entrada.range(f'Q{NumeroCelula}').value = PesoPecaList[i]
        Pagina_Entrada.range(f'S{NumeroCelula}').value = QTDList[i]
        Pagina_Entrada.range(f'T{NumeroCelula}').value = ComponentesList[i]
        Pagina_Entrada.range(f'Y{NumeroCelula}').value = DataEntradaList[i]
        Pagina_Entrada.range(f'Z{NumeroCelula}').value = ReponsavelList[i]
        Pagina_Entrada.range(f'AA{NumeroCelula}').value = DataRespostaList[i]
        
        i += 1
        
    Plan_Entrada.save()
    log_func("✓ Planilha APQP preenchida e salva com sucesso!")

# =========================================================
# CLASSES DA INTERFACE GRÁFICA
# =========================================================
class AppAPQP(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Sistema de Automação APQP")
        self.geometry("780x620")
        self.configure(bg=COR_FUNDO)

        # Configurar estilos de widgets TTK
        self.style = ttk.Style()
        self.style.theme_use('default')
        
        # Frame Estilizado (Card)
        self.style.configure('Card.TFrame', background=COR_CARD)
        
        # Rótulos
        self.style.configure('TLabel', background=COR_CARD, foreground=COR_TEXTO, font=('Segoe UI', 10))
        self.style.configure('Header.TLabel', background=COR_FUNDO, foreground=COR_TEXTO, font=('Segoe UI', 14, 'bold'))

        self.criar_layout()
        self.carregar_config_inicial()

    def criar_layout(self):
        # Título da Aplicação
        lbl_titulo = ttk.Label(self, text="Gerenciador e Automatizador APQP", style='Header.TLabel')
        lbl_titulo.pack(pady=(15, 10), padx=20, anchor='w')

        # Painel Configuração de Arquivo (Card)
        card_arquivo = ttk.Frame(self, style='Card.TFrame', padding=15)
        card_arquivo.pack(fill='x', padx=20, pady=5)

        lbl_arq = ttk.Label(card_arquivo, text="Arquivo Alvo APQP (Nome ou Caminho Completo):")
        lbl_arq.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 5))

        self.ent_caminho = tk.Entry(
            card_arquivo, bg=COR_FUNDO, fg=COR_TEXTO, insertbackground=COR_TEXTO,
            font=('Segoe UI', 10), relief='flat', bd=5
        )
        self.ent_caminho.grid(row=1, column=0, sticky='ew', padx=(0, 10))
        card_arquivo.columnconfigure(0, weight=1)

        btn_procurar = tk.Button(
            card_arquivo, text="Buscar Arquivo", bg=COR_DESTAQUE, fg=COR_TEXTO,
            activebackground="#5b6eae", activeforeground=COR_TEXTO,
            font=('Segoe UI', 9, 'bold'), relief='flat', cursor='hand2',
            command=self.selecionar_arquivo
        )
        btn_procurar.grid(row=1, column=1)

        # Painel Execução e Configurações de Linha (Card)
        card_exec = ttk.Frame(self, style='Card.TFrame', padding=15)
        card_exec.pack(fill='x', padx=20, pady=10)

        lbl_celula = ttk.Label(card_exec, text="Linha Inicial para Inserção no APQP (ex: 15):")
        lbl_celula.grid(row=0, column=0, sticky='w', pady=(0, 5))

        self.ent_celula = tk.Entry(
            card_exec, bg=COR_FUNDO, fg=COR_TEXTO, insertbackground=COR_TEXTO,
            font=('Segoe UI', 10), relief='flat', bd=5, width=15
        )
        self.ent_celula.insert(0, "15")
        self.ent_celula.grid(row=1, column=0, sticky='w', pady=(0, 10))

        # Botões de Ação
        btn_frame = tk.Frame(card_exec, bg=COR_CARD)
        btn_frame.grid(row=2, column=0, sticky='ew')

        btn_executar_tudo = tk.Button(
            btn_frame, text="🚀 Executar Processo Completo", bg=COR_DESTAQUE, fg=COR_TEXTO,
            activebackground="#5b6eae", activeforeground=COR_TEXTO,
            font=('Segoe UI', 10, 'bold'), relief='flat', cursor='hand2', 
            padx=12, pady=6, # <--- Substituído 'padding=8' por 'padx' e 'pady'
            command=self.iniciar_processo_thread
        )
        btn_executar_tudo.pack(side='left', padx=(0, 10))

        # Console de Saída / Logs
        card_log = ttk.Frame(self, style='Card.TFrame', padding=10)
        card_log.pack(fill='both', expand=True, padx=20, pady=(0, 20))

        lbl_log = ttk.Label(card_log, text="Console de Status / Logs:")
        lbl_log.pack(anchor='w', pady=(0, 5))

        self.txt_log = scrolledtext.ScrolledText(
            card_log, bg=COR_CARD, fg=COR_TEXTO, insertbackground=COR_TEXTO,
            font=('Consolas', 9), relief='flat', bd=0
        )
        self.txt_log.pack(fill='both', expand=True)

    def log(self, mensagem):
        """Escreve mensagens na área de log de forma segura."""
        self.txt_log.insert(tk.END, f"{mensagem}\n")
        self.txt_log.see(tk.END)

    def carregar_config_inicial(self):
        caminho_salvo = LendoArquivo_Jason()
        if caminho_salvo:
            self.ent_caminho.delete(0, tk.END)
            self.ent_caminho.insert(0, caminho_salvo)
            self.log(f"Configuração carregada. Arquivo alvo: '{caminho_salvo}'")

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione o arquivo da APQP",
            filetypes=[("Arquivos Excel", "*.xlsx *.xlsm")]
        )
        if caminho:
            self.ent_caminho.delete(0, tk.END)
            self.ent_caminho.insert(0, caminho)
            escrevendoArquivo_Jason(caminho)
            self.log(f"Novo caminho salvo em conf.json: {caminho}")

    def iniciar_processo_thread(self):
        """Inicia o processamento em uma thread separada para não travar a GUI."""
        caminho = self.ent_caminho.get().strip()
        celula = self.ent_celula.get().strip()

        if not caminho:
            messagebox.showwarning("Aviso", "Por favor, especifique o nome ou caminho do arquivo APQP.")
            return

        if not celula.isdigit():
            messagebox.showwarning("Aviso", "Informe um número válido para a linha inicial.")
            return

        escrevendoArquivo_Jason(caminho)
        
        # Desabilita o botão enquanto roda
        thread = threading.Thread(target=self.executar_rotina, args=(caminho, celula))
        thread.daemon = True
        thread.start()

    def executar_rotina(self, caminho_alvo, celula_inicio):
        self.log("\n==========================================")
        self.log("Iniciando busca do arquivo e processamento...")

        # 1. Busca do Arquivo se apenas o nome tiver sido fornecido
        caminhoAPQP = caminho_alvo
        if not os.path.exists(caminhoAPQP):
            self.log(f"Procurando por '{caminho_alvo}' no sistema...")
            pasta_busca = os.path.expanduser("~")
            encontrado = False
            for root, dirs, files in os.walk(pasta_busca):
                if caminho_alvo in files or os.path.basename(caminho_alvo) in files:
                    caminhoAPQP = os.path.join(root, caminho_alvo)
                    self.log(f"✓ Arquivo localizado: {caminhoAPQP}")
                    encontrado = True
                    break
            
            if not encontrado:
                self.log(f"❌ Erro: O arquivo '{caminho_alvo}' não foi localizado no seu computador.")
                return

        try:
            # 2. Leitura dos dados da planilha MainAPQP
            self.log("Lendo dados do arquivo 'MainAPQP.xlsx'...")
            ObservacaoList, PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, Desenhos2DList, Desenhos3DList, VolumeAnualList, DataEntradaList, DataRespostaList, QTDLista, ComponentesLista, ReponsavelLista, realizadoLista, ClienteLista, tipoLista, descricaoLista, compradorLista, PesoPecaLista = PegandoDados_PQPMain('MainAPQP')

            # 3. Criando pastas e gerando arquivos ACC
            self.log("Criando pastas e preenchendo arquivos ACC.xlsx...")
            CriandoPastas(PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoLista, self.log)

            # 4. Preenchendo APQP final via xlwings
            self.log(f"Preenchendo APQP a partir da linha {celula_inicio}...")
            PreenchendoAPQP_Exec(caminhoAPQP, celula_inicio, self.log)

            self.log("✨ Processo concluído com sucesso!")
            messagebox.showinfo("Sucesso", "Todo o fluxo de automação foi concluído!")

        except Exception as e:
            self.log(f"❌ Ocorreu um erro durante a execução: {str(e)}")
            messagebox.showerror("Erro", f"Erro durante o processamento:\n{str(e)}")


if __name__ == "__main__":
    app = AppAPQP()
    app.mainloop()