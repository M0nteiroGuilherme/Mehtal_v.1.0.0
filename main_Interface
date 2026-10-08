import re
import os
import sys
import shutil
import json
import time
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from openpyxl import load_workbook
import xlwings as xw

# =========================================================
# CONFIGURAÇÃO DE PALETA DE CORES E TEMA (DARK MODE)
# =========================================================
BG_DARK = "#0f172a"        # Fundo Principal (Slate 900)
CARD_BG = "#1e293b"        # Fundo dos Cards/Painéis (Slate 800)
CARD_BORDER = "#334155"    # Borda dos Cards (Slate 700)
TEXT_MAIN = "#f8fafc"      # Texto Principal (Slate 50)
TEXT_MUTED = "#94a3b8"     # Texto Secundário (Slate 400)
ACCENT_BLUE = "#2563eb"    # Azul Principal
ACCENT_BLUE_HOVER = "#1d4ed8"
ACCENT_GREEN = "#059669"   # Verde Destaque
ACCENT_GREEN_HOVER = "#047857"
CONSOLE_BG = "#020617"     # Fundo do Terminal (Slate 950)
CONSOLE_FG = "#38bdf8"     # Texto do Terminal (Sky 400)
INPUT_BG = "#0f172a"       # Fundo dos Campos de Entrada
INPUT_FG = "#f8fafc"       # Texto dos Campos de Entrada

# =========================================================
# CORREÇÃO PARA FUNCIONAR EM EXE (PyInstaller)
# =========================================================
def get_app_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).resolve().parent

def get_resource_path(nome_arquivo):
    if getattr(sys, 'frozen', False):
        base = Path(sys._MEIPASS)  # type: ignore
    else:
        base = Path(__file__).resolve().parent
    return base / nome_arquivo

APP_DIR = get_app_dir()
CONF_PATH = APP_DIR / "conf.json"

MODELO_ACC = get_resource_path("ACC.xlsx")
MODELO_CUSTO = get_resource_path("Planilha_Custo.xlsm")
MODELO_FIIN = get_resource_path("FIIN.xlsx")

# =========================================================
# FUNÇÕES DE FORMATAÇÃO E UTILITÁRIOS
# =========================================================
def format_texto(texto):
    if texto is None:
        return ""
    return re.sub(r'[\\/*?:"<>|]', '_', str(texto)).strip()

def formatar_volume(valor):
    if valor is None or str(valor).strip() == "":
        return ""
    if isinstance(valor, list):
        if len(valor) == 0:
            return ""
        flat = []
        for v in valor:
            if isinstance(v, list):
                flat.extend(v)
            else:
                flat.append(v)
        flat = [x for x in flat if x not in (None, "")]
        if len(flat) == 0:
            return ""
        if len(flat) == 1:
            valor = flat[0]
        else:
            partes = []
            for v in flat:
                try:
                    partes.append(f"{int(float(v)):,}")
                except:
                    partes.append(str(v))
            return ", ".join(partes)
    try:
        num = int(float(valor))
        return f"{num:,}"
    except (ValueError, TypeError):
        return str(valor).strip()

# =========================================================
# OPERAÇÕES DE EXCEL
# =========================================================
def prenchendoPlanilha_ACC(caminho_arquivo, PN_MTH_ACC, PN_Cliente_ACC, REV_Acc, RFQ_Acc, Projeto_Acc, ClientePlanta_Acc, VolumeAnual_Acc, DataEntrada_Acc, DataResposta_Acc, TipoItem_Acc, Nome_Peca_Acc, Comprador_Acc):
    Plan_Saida = load_workbook(caminho_arquivo)
    Pagina_Saida = Plan_Saida['Plan1']
    Pagina_Saida['F3'] = str(PN_MTH_ACC) if PN_MTH_ACC else ""
    Pagina_Saida['F6'] = str(PN_Cliente_ACC) if PN_Cliente_ACC else ""
    Pagina_Saida['Q6'] = str(REV_Acc) if REV_Acc else ""
    Pagina_Saida['S3'] = str(RFQ_Acc) if RFQ_Acc else ""
    Pagina_Saida['AD3'] = str(Projeto_Acc) if Projeto_Acc else ""
    Pagina_Saida['AK3'] = str(ClientePlanta_Acc) if ClientePlanta_Acc else ""
    Pagina_Saida['Z6'] = str(Nome_Peca_Acc) if Nome_Peca_Acc else ""
    Pagina_Saida['AJ6'] = str(Comprador_Acc) if Comprador_Acc else ""
    Pagina_Saida['H12'] = str(DataEntrada_Acc).split(' ')[0] if DataEntrada_Acc else ""
    Pagina_Saida['T12'] = str(DataResposta_Acc).split(' ')[0] if DataResposta_Acc else ""
    
    vol = VolumeAnual_Acc
    if isinstance(vol, list):
        Pagina_Saida['F9'].value = ", ".join([formatar_volume(v) for v in vol])
    else:
        Pagina_Saida['F9'].value = formatar_volume(vol)
            
    if TipoItem_Acc == "Novo":
        Pagina_Saida['S9'].value = "X"
    elif TipoItem_Acc == "Protótipo":
        Pagina_Saida['S10'].value = "X"    
    elif TipoItem_Acc == "Modificação":
        Pagina_Saida['AA9'].value = "X"
    else:
        if TipoItem_Acc:
            Pagina_Saida['V10'].value = f"{TipoItem_Acc}"
            Pagina_Saida['AA10'].value = "X"
    Plan_Saida.save(caminho_arquivo)


def prenchendoPlanilha_Custo(caminho_arquivo, PN_cliente, PN_MethalLsit, Revisao, Volume, Cliente, Tipo, Peso):
    time.sleep(0.1)
    Plan_saida = load_workbook(caminho_arquivo, keep_vba=True)
    pagina_saida = Plan_saida['Planilha de Custo']
    pagina_saida['C6'].value = str(PN_cliente) if PN_cliente else ""
    pagina_saida['C7'].value = str(PN_MethalLsit) if PN_MethalLsit else ""
    pagina_saida['C8'].value = str(Revisao) if Revisao else ""
    pagina_saida['F9'].value = str(Cliente) if Cliente else ""
    pagina_saida['G8'].value = str(Tipo) if Tipo else ""
    pagina_saida['M8'].value = str(Peso) if Peso else ""
    
    if isinstance(Volume, list):
        pagina_saida['F9'].value = ", ".join([formatar_volume(v) for v in Volume])
    else:
        pagina_saida['F9'].value = formatar_volume(Volume)
    
    Plan_saida.save(caminho_arquivo)


def prenchendoPlanilha_FINN(caminho_Arquivo, PNsList, Cliente, Planta_clinete, Local_clinete, Clinete_Comprador, Projeto, Descricao, Tipo, Rev, PNsListInterno, Volume):
    Plan_saida = load_workbook(caminho_Arquivo)
    
    if 'FOR.ENG.03' in Plan_saida.sheetnames:
        pagina_saida = Plan_saida['FOR.ENG.03']
    else:
        pagina_saida = Plan_saida.active
    
    pagina_saida['C6'].value = str(Cliente) if Cliente else ""
    pagina_saida['H6'].value = str(Planta_clinete) if Planta_clinete else ""
    pagina_saida['K6'].value = str(Local_clinete) if Local_clinete else ""
    pagina_saida['C7'].value = str(Clinete_Comprador) if Clinete_Comprador else ""
    pagina_saida['H7'].value = str(Projeto) if Projeto else ""
    pagina_saida['H8'].value = str(Descricao) if Descricao else ""
    pagina_saida['H9'].value = str(Rev) if Rev else ""
    pagina_saida['E9'].value = str(PNsList) if PNsList else ""
    pagina_saida['E8'].value = str(PNsListInterno) if PNsListInterno else ""
    
    try:
        vol_str = formatar_volume(Volume)
        pagina_saida['D15'].value = str(vol_str)
    except Exception:
        pagina_saida['D15'].value = str(Volume)
    
    if Tipo == "Novo":
        pagina_saida['C5'].value = "X"
    elif Tipo == "Modificação":
        pagina_saida['F5'].value = "X"
    elif Tipo == "Substituição":
        pagina_saida['I5'].value = "X"
    else:
        if Tipo:
            pagina_saida['H5'].value = str(Tipo)
            pagina_saida['I5'].value = "X"
        
    Plan_saida.save(caminho_Arquivo)


def PegandoDados_APQP(caminho_apqp: Path, linha_inicio: str, linha_fim: str):
    if not caminho_apqp or not caminho_apqp.exists():
        raise FileNotFoundError(f"Ficheiro APQP não encontrado: {caminho_apqp}")
    
    print(f"📂 A ler APQP: {caminho_apqp.name} (Linhas {linha_inicio} a {linha_fim})")
    
    app = xw.App(visible=False)
    try:
        Plan_Entrada = app.books.open(str(caminho_apqp))
        Pagina_Entrada = Plan_Entrada.sheets['APQP']
        
        PN_Methal     = Pagina_Entrada.range(f'A{linha_inicio}:A{linha_fim}').value
        Cliente       = Pagina_Entrada.range(f'F{linha_inicio}:F{linha_fim}').value  
        ClientePlanta = Pagina_Entrada.range(f'G{linha_inicio}:G{linha_fim}').value  
        tipo          = Pagina_Entrada.range(f'H{linha_inicio}:H{linha_fim}').value
        RFQ           = Pagina_Entrada.range(f'I{linha_inicio}:I{linha_fim}').value    
        PNs           = Pagina_Entrada.range(f'J{linha_inicio}:J{linha_fim}').value   
        REV           = Pagina_Entrada.range(f'K{linha_inicio}:K{linha_fim}').value
        descricao     = Pagina_Entrada.range(f'L{linha_inicio}:L{linha_fim}').value
        comprador     = Pagina_Entrada.range(f'M{linha_inicio}:M{linha_fim}').value
        VolumeAnual   = Pagina_Entrada.range(f'P{linha_inicio}:P{linha_fim}').value   
        Peso          = Pagina_Entrada.range(f'Q{linha_inicio}:Q{linha_fim}').value   
        Declinar      = Pagina_Entrada.range(f'V{linha_inicio}:V{linha_fim}').value
        
        celulas_entrada = Pagina_Entrada.range(f'Y{linha_inicio}:Y{linha_fim}') 
        DataEntrada = [cell.api.Text for cell in celulas_entrada]
        
        celulas_resposta = Pagina_Entrada.range(f'AA{linha_inicio}:AA{linha_fim}')
        DataResposta = [cell.api.Text for cell in celulas_resposta]
        
        Plan_Entrada.close()
    finally:
        app.quit()

    def PegarInf(ColumSelecionada):
        if ColumSelecionada is None:
            return []
        if not isinstance(ColumSelecionada, list):
            ColumSelecionada = [ColumSelecionada]
        return [str(Dado).strip() if Dado is not None else "" for Dado in ColumSelecionada]

    PN_MethalLsit       = PegarInf(PN_Methal)
    PNsList_Formatar    = PegarInf(PNs)
    REVList_Formatar    = PegarInf(REV)
    RFQList             = PegarInf(RFQ)
    ClientePlantaList   = PegarInf(ClientePlanta)
    VolumeAnualList     = PegarInf(VolumeAnual)
    DataEntradaList     = PegarInf(DataEntrada)
    DataRespostaList    = PegarInf(DataResposta)
    ClienteList         = PegarInf(Cliente)
    tipoList            = PegarInf(tipo)
    descricaoLista      = PegarInf(descricao)
    compradorList       = PegarInf(comprador)
    PesoList            = PegarInf(Peso)

    if Declinar is None:
        DeclinarList = [0] * len(PNsList_Formatar)
    elif not isinstance(Declinar, list):
        DeclinarList = [Declinar]
    else:
        DeclinarList = Declinar

    ListaFormanta_CompradorACC = []
    PNsList = [format_texto(pn) for pn in PNsList_Formatar]
    REVList = [format_texto(rev) for rev in REVList_Formatar]
    
    for i in range(len(PNsList)):
        comp = compradorList[i] if i < len(compradorList) else ""
        cli  = ClienteList[i]   if i < len(ClienteList)   else ""
        if comp and cli:
            ListaFormanta_CompradorACC.append(f"{comp}   {cli}")
        else:
            ListaFormanta_CompradorACC.append(f"{comp}{cli}")
    
    return PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, compradorList, ClienteList, PesoList, DeclinarList


def CriandoPastas(PN_MethalLsit_Criar, PNs_Ciar, rev_Craidas, RFQ_criar, Projeto_criar, ClientePlanta_criar, VolumeAnual_criar, DataEntrada_criar, DataResposta_criar, TipoItem_criar, Nome_Peca_Criar, Comprador_Criar, ClienteList_cria, compradorList_criar, peso_criar, Declinar_crirar):
    
    pastas_Secundarias = [
        "1-RFQ", "2-Desenhos", "3-Analise Critica", "4-Planilha de Custo",
        "5-CBD", "6-Carta Comercial", "8-Terceiros", "9-Pedidos",
        "10-FIIN", "11-Emails", "12-Custo Logistico", "13-Obsoleto", "14-Modificações"
    ] 
    
    if len(PNs_Ciar) == 0:
        print("⚠️ A lista de PNs está vazia.")
        return

    for i in range(len(PNs_Ciar)):
        if i < len(Declinar_crirar) and Declinar_crirar[i] == 1:
            print(f"⏩ A ignorar {PNs_Ciar[i]} (declinado)")
            continue
        pn_limpo = format_texto(PNs_Ciar[i])
        if not pn_limpo:
            continue
        rfq_limpo = format_texto(RFQ_criar[i])
        pasta_principal = pn_limpo
        caminho_Base = os.path.join(rfq_limpo, pasta_principal) if rfq_limpo else pasta_principal
        
        os.makedirs(caminho_Base, exist_ok=True)
        for nome_subpasta in pastas_Secundarias:
            os.makedirs(os.path.join(caminho_Base, nome_subpasta), exist_ok=True)

        novo_nome_excel = f'{pasta_principal}_REV.{rev_Craidas[i]}_ACC.xlsx'
        caminho_excel_copiado = os.path.join(caminho_Base, "3-Analise Critica", novo_nome_excel)
        if not MODELO_ACC.exists():
            print(f"❌ Modelo ACC não encontrado em {MODELO_ACC}")
            continue
        shutil.copy2(MODELO_ACC, caminho_excel_copiado)
        prenchendoPlanilha_ACC(
            caminho_excel_copiado, PN_MethalLsit_Criar[i], PNs_Ciar[i], rev_Craidas[i], RFQ_criar[i], Projeto_criar[i], ClientePlanta_criar[i], VolumeAnual_criar[i], DataEntrada_criar[i], DataResposta_criar[i], TipoItem_criar[i], Nome_Peca_Criar[i], Comprador_Criar[i]
        )
        
        Novo_PlanilhaCusto = f'{pasta_principal}_Planilha_Custo.xlsm'
        PlanilhaCusto_copiada = os.path.join(caminho_Base, "4-Planilha de Custo", Novo_PlanilhaCusto)
        if not MODELO_CUSTO.exists():
            print("❌ Modelo Custo não encontrado")
            continue
        shutil.copy2(MODELO_CUSTO, PlanilhaCusto_copiada)
        prenchendoPlanilha_Custo(PlanilhaCusto_copiada, PNs_Ciar[i], PN_MethalLsit_Criar[i], rev_Craidas[i], VolumeAnual_criar[i], ClienteList_cria[i], TipoItem_criar[i], peso_criar[i])
        print(f"✅ Sucesso: Pasta '{pasta_principal}' criada com ACC e Custo!")


def CriandoFIIN(PN_MethalLsit_Criar, PNs_Ciar, rev_Craidas, RFQ_criar, ClientePlanta_criar, VolumeAnual_criar, TipoItem_criar, Nome_Peca_Criar, ClienteList_cria, compradorList_criar, Declinar_crirar):

    if len(PNs_Ciar) == 0:
        print("⚠️ A lista de PNs está vazia.")
        return

    for i in range(len(PNs_Ciar)):
        if i < len(Declinar_crirar) and Declinar_crirar[i] == 1:
            continue
        pn_limpo = format_texto(PNs_Ciar[i])
        if not pn_limpo:
            continue
        pasta_principal = pn_limpo
        caminho_destino_pasta = os.path.join(pasta_principal, "10-FIIN") if os.path.exists(pasta_principal) else pasta_principal
        os.makedirs(caminho_destino_pasta, exist_ok=True)

        Novo_Fiin = f"{pasta_principal}_FIIN.xlsx"
        Fiin_copiada = os.path.join(caminho_destino_pasta, Novo_Fiin)

        if not MODELO_FIIN.exists():
            print(f"❌ Modelo FIIN não encontrado em {MODELO_FIIN}")
            continue

        shutil.copy2(MODELO_FIIN, Fiin_copiada)
                                
        prenchendoPlanilha_FINN(
            Fiin_copiada, 
            PNs_Ciar[i], 
            ClienteList_cria[i], 
            ClientePlanta_criar[i], 
            ClientePlanta_criar[i], 
            compradorList_criar[i], 
            RFQ_criar[i], 
            Nome_Peca_Criar[i], 
            TipoItem_criar[i], 
            rev_Craidas[i], 
            PN_MethalLsit_Criar[i],
            VolumeAnual_criar[i]
        )
        print(f"✅ FIIN para '{pasta_principal}' criada!")

# =========================================================
# REDIRECIONADOR DE CONSOLA PARA TKINTER
# =========================================================
class ConsoleRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, str_val):
        self.text_widget.insert(tk.END, str_val)
        self.text_widget.see(tk.END)

    def flush(self):
        pass

# =========================================================
# INTERFACE GRÁFICA MODERNA (TKINTER COM TEMA PERSONALIZADO)
# =========================================================
class APQPAppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Comercial - Automação APQP / ACC / FIIN")
        self.root.geometry("860x700")
        self.root.minsize(750, 600)
        self.root.configure(bg=BG_DARK)

        self.apqp_path = None
        
        self.aplicar_estilos()
        self.setup_ui()
        self.carregar_configuracao_inicial()
        
        # Redireciona o stdout (print) para a caixa de texto
        sys.stdout = ConsoleRedirector(self.log_text)

    def aplicar_estilos(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configurações Gerais
        style.configure('.', background=BG_DARK, foreground=TEXT_MAIN, font=('Segoe UI', 9))
        
        # Frame de Card Personalizado
        style.configure('Card.TFrame', background=CARD_BG, relief='flat', borderwidth=1)
        
        # Labels
        style.configure('Title.TLabel', background=BG_DARK, foreground=TEXT_MAIN, font=('Segoe UI', 14, 'bold'))
        style.configure('SubTitle.TLabel', background=BG_DARK, foreground=TEXT_MUTED, font=('Segoe UI', 9))
        style.configure('CardHeader.TLabel', background=CARD_BG, foreground=TEXT_MAIN, font=('Segoe UI', 10, 'bold'))
        style.configure('CardBody.TLabel', background=CARD_BG, foreground=TEXT_MUTED, font=('Segoe UI', 9))
        style.configure('Path.TLabel', background=CARD_BG, foreground='#38bdf8', font=('Consolas', 9))

        # Entradas de Texto
        style.configure('Custom.TEntry', fieldbackground=INPUT_BG, foreground=INPUT_FG, insertcolor=TEXT_MAIN, borderwidth=1)

        # Botões Principais (Azul)
        style.configure('Primary.TButton', 
                        background=ACCENT_BLUE, 
                        foreground='#ffffff', 
                        font=('Segoe UI', 9, 'bold'), 
                        borderwidth=0, 
                        focusthickness=0, 
                        padding=(12, 8))
        style.map('Primary.TButton', 
                  background=[('active', ACCENT_BLUE_HOVER), ('disabled', '#334155')],
                  foreground=[('disabled', '#64748b')])

        # Botões Secundários/Ação (Verde)
        style.configure('Success.TButton', 
                        background=ACCENT_GREEN, 
                        foreground='#ffffff', 
                        font=('Segoe UI', 9, 'bold'), 
                        borderwidth=0, 
                        focusthickness=0, 
                        padding=(12, 8))
        style.map('Success.TButton', 
                  background=[('active', ACCENT_GREEN_HOVER), ('disabled', '#334155')],
                  foreground=[('disabled', '#64748b')])

        # Botão neutro (Cinza escuro)
        style.configure('Secondary.TButton', 
                        background='#334155', 
                        foreground='#ffffff', 
                        font=('Segoe UI', 9), 
                        borderwidth=0, 
                        padding=(10, 6))
        style.map('Secondary.TButton', background=[('active', '#475569')])

    def criar_card(self, parent, titulo):
        # Cria um container em formato de cartão com cabeçalho
        outer_frame = tk.Frame(parent, bg=CARD_BORDER, bd=1, relief='flat')
        card_content = tk.Frame(outer_frame, bg=CARD_BG, padx=15, pady=12)
        card_content.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        lbl_title = ttk.Label(card_content, text=titulo, style='CardHeader.TLabel')
        lbl_title.pack(anchor=tk.W, pady=(0, 8))

        return outer_frame, card_content

    def setup_ui(self):
        main_container = tk.Frame(self.root, bg=BG_DARK, padx=20, pady=20)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Header Principal
        header_frame = tk.Frame(main_container, bg=BG_DARK)
        header_frame.pack(fill=tk.X, pady=(0, 15))

        lbl_title = ttk.Label(header_frame, text="⚙️ Automação de Comercial", style='Title.TLabel')
        lbl_title.pack(anchor=tk.W)
        lbl_sub = ttk.Label(header_frame, text="Gestão de APQP, Geração Automática de ACC, Custo e FIIN", style='SubTitle.TLabel')
        lbl_sub.pack(anchor=tk.W)

        # 1. Card de Seleção do Ficheiro APQP
        card_apqp_border, card_apqp = self.criar_card(main_container, "📁 Ficheiro APQP Ativo")
        card_apqp_border.pack(fill=tk.X, pady=(0, 12))

        apqp_info_frame = tk.Frame(card_apqp, bg=CARD_BG)
        apqp_info_frame.pack(fill=tk.X)

        self.lbl_apqp_path = ttk.Label(apqp_info_frame, text="Nenhum ficheiro selecionado...", style='Path.TLabel', wraplength=580)
        self.lbl_apqp_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        btn_selecionar = ttk.Button(apqp_info_frame, text="📂 Selecionar APQP", style='Secondary.TButton', command=self.selecionar_apqp)
        btn_selecionar.pack(side=tk.RIGHT)

        # 2. Card de Intervalo de Linhas
        card_linhas_border, card_linhas = self.criar_card(main_container, "📊 Intervalo de Leitura (Linhas no Excel)")
        card_linhas_border.pack(fill=tk.X, pady=(0, 12))

        grid_linhas = tk.Frame(card_linhas, bg=CARD_BG)
        grid_linhas.pack(fill=tk.X)

        lbl_inicio = ttk.Label(grid_linhas, text="Linha Inicial:", style='CardBody.TLabel')
        lbl_inicio.grid(row=0, column=0, padx=(0, 8), pady=5, sticky=tk.W)
        
        self.entry_inicio = ttk.Entry(grid_linhas, width=10, style='Custom.TEntry')
        self.entry_inicio.grid(row=0, column=1, padx=(0, 20), pady=5)
        self.entry_inicio.insert(0, "2")

        lbl_fim = ttk.Label(grid_linhas, text="Linha Final:", style='CardBody.TLabel')
        lbl_fim.grid(row=0, column=2, padx=(0, 8), pady=5, sticky=tk.W)
        
        self.entry_fim = ttk.Entry(grid_linhas, width=10, style='Custom.TEntry')
        self.entry_fim.grid(row=0, column=3, padx=(0, 10), pady=5)
        self.entry_fim.insert(0, "10")

        # 3. Card de Operações/Ações
        card_acoes_border, card_acoes = self.criar_card(main_container, "🚀 Execução das Automações")
        card_acoes_border.pack(fill=tk.X, pady=(0, 12))

        btns_frame = tk.Frame(card_acoes, bg=CARD_BG)
        btns_frame.pack(fill=tk.X, pady=5)

        self.btn_acc = ttk.Button(btns_frame, text="⚡ Criar ACC & Estrutura de Pastas", style='Primary.TButton', command=lambda: self.executar_em_thread(self.acao_criar_acc))
        self.btn_acc.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self.btn_fiin = ttk.Button(btns_frame, text="📄 Gerar Ficheiros FIIN", style='Success.TButton', command=lambda: self.executar_em_thread(self.acao_criar_fiin))
        self.btn_fiin.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        # 4. Card do Terminal / Console
        card_log_border, card_log = self.criar_card(main_container, "🖥️ Console de Processamento em Tempo Real")
        card_log_border.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            card_log, 
            height=10, 
            font=("Consolas", 9), 
            bg=CONSOLE_BG, 
            fg=CONSOLE_FG, 
            insertbackground="#ffffff",
            relief='flat',
            bd=0
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

    def carregar_configuracao_inicial(self):
        if CONF_PATH.exists():
            try:
                with open(CONF_PATH, "r", encoding='utf-8') as arquivo:
                    conteudo = json.load(arquivo)
                    caminho_str = conteudo.get("apqp_path") if isinstance(conteudo, dict) else conteudo
                    if caminho_str:
                        caminho = Path(caminho_str)
                        if caminho.exists():
                            self.apqp_path = caminho
                            self.lbl_apqp_path.config(text=str(caminho))
            except Exception as e:
                print(f"⚠️ Erro ao carregar conf.json: {e}")

    def selecionar_apqp(self):
        ficheiro = filedialog.askopenfilename(
            title="Selecionar Ficheiro APQP",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("Todos os ficheiros", "*.*")]
        )
        if ficheiro:
            caminho = Path(ficheiro)
            self.apqp_path = caminho
            self.lbl_apqp_path.config(text=str(caminho))
            try:
                with open(CONF_PATH, 'w', encoding='utf-8') as a:
                    json.dump({"apqp_path": str(caminho)}, a, indent=4, ensure_ascii=False)
                print(f"💾 Caminho APQP guardado: {caminho.name}")
            except Exception as e:
                print(f"⚠️ Erro ao guardar configuração: {e}")

    def validar_entradas(self):
        if not self.apqp_path or not self.apqp_path.exists():
            messagebox.showerror("Atenção", "Por favor, selecione um ficheiro APQP válido primeiro.")
            return None, None
        
        inicio = self.entry_inicio.get().strip()
        fim = self.entry_fim.get().strip()
        
        if not inicio.isdigit() or not fim.isdigit():
            messagebox.showerror("Erro de Entrada", "As linhas inicial e final devem conter apenas números.")
            return None, None
            
        if int(inicio) > int(fim):
            messagebox.showerror("Erro de Intervalo", "A linha inicial não pode ser superior à linha final.")
            return None, None
            
        return inicio, fim

    def definir_estado_botoes(self, ativo: bool):
        estado = tk.NORMAL if ativo else tk.DISABLED
        self.btn_acc.config(state=estado)
        self.btn_fiin.config(state=estado)

    def executar_em_thread(self, funcao):
        self.definir_estado_botoes(False)
        def wrapper():
            try:
                funcao()
            except Exception as e:
                print(f"\n❌ Erro na Execução: {e}")
            finally:
                self.definir_estado_botoes(True)
        
        t = threading.Thread(target=wrapper, daemon=True)
        t.start()

    def acao_criar_acc(self):
        inicio, fim = self.validar_entradas()
        if not inicio:
            return
        
        print("\n" + "="*55)
        print("🚀 A Iniciar Automação de Pastas, ACC e Custo...")
        
        PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, compradorList, ClienteList, PesoList, Declinar = PegandoDados_APQP(self.apqp_path, inicio, fim)
        
        CriandoPastas(
            PN_MethalLsit, PNsList, REVList, RFQList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, ClienteList, compradorList, PesoList, Declinar
        )
        print("✨ Processo finalizado com sucesso!")

    def acao_criar_fiin(self):
        inicio, fim = self.validar_entradas()
        if not inicio:
            return
        
        print("\n" + "="*55)
        print("📄 A Iniciar Geração de Ficheiros FIIN...")
        
        PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, compradorList, ClienteList, PesoList, Declinar = PegandoDados_APQP(self.apqp_path, inicio, fim)
        
        CriandoFIIN(
            PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, tipoList, descricaoLista, ClienteList, compradorList, Declinar
        )
        print("✨ Processo finalizado com sucesso!")

# =========================================================
# PONTO DE ENTRADA DO PROGRAMA
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = APQPAppGUI(root)
    root.mainloop()