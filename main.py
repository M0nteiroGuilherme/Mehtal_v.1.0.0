import re
import os
import sys
import shutil
import json
import time
from pathlib import Path
from openpyxl import load_workbook
import xlwings as xw

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


def format(texto):
    if texto is None:
        return ""
    return re.sub(r'[\\/*?:"<>|]', '_', str(texto)).strip()

def formatar_volume(valor):
    if valor is None or str(valor).strip() == "":
        return ""
    # Se for lista, pega primeiro ou junta
    if isinstance(valor, list):
        if len(valor) == 0:
            return ""
        # Se for lista de listas (caso do xlwings)
        flat = []
        for v in valor:
            if isinstance(v, list):
                flat.extend(v)
            else:
                flat.append(v)
        # filtra vazios
        flat = [x for x in flat if x not in (None, "")]
        if len(flat) == 0:
            return ""
        if len(flat) == 1:
            valor = flat[0]
        else:
            # junta vários volumes
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
# BUSCA INTELIGENTE DA APQP PELO NOME
# =========================================================
def encontrar_apqp_por_nome(nome_apqp: str) -> Path | None:
    nome_limpo = nome_apqp.strip()
    if not nome_limpo:
        return None
    p_teste = Path(nome_limpo)
    if p_teste.exists() and p_teste.is_file():
        return p_teste
    if Path(str(p_teste) + ".xlsx").exists():
        return Path(str(p_teste) + ".xlsx")
    nome_arquivo = nome_limpo if nome_limpo.lower().endswith('.xlsx') else f"{nome_limpo}.xlsx"
    print(f"\n🔍 Procurando por '{nome_arquivo}'...")

    pastas_busca = [APP_DIR]
    try:
        home = Path.home()
        for sub in ["Documents", "Documentos", "Desktop", "Área de Trabalho", "Downloads"]:
            p = home / sub
            if p.exists():
                pastas_busca.append(p)
        if (home / "OneDrive").exists():
            pastas_busca.append(home / "OneDrive")
    except:
        pass

    IGNORAR = {'Windows', 'Program Files', 'Program Files (x86)', 'ProgramData',
               '$Recycle.Bin', 'AppData', '.git', '__pycache__', 'node_modules'}

    def busca_rapida(pasta_base: Path):
        try:
            for arquivo in pasta_base.rglob(nome_arquivo):
                if any(ign in arquivo.parts for ign in IGNORAR):
                    continue
                if arquivo.is_file():
                    return arquivo
        except:
            pass
        return None

    for pasta in pastas_busca:
        print(f"  > Verificando em: {pasta}")
        resultado = busca_rapida(pasta)
        if resultado:
            print(f"✅ Encontrado em: {resultado}")
            return resultado

    print("  > Varrendo C: e D:...")
    for letra in ["C:", "D:"]:
        drive = Path(letra + "/")
        if not drive.exists():
            continue
        try:
            for root, dirs, files in os.walk(drive, topdown=True):
                dirs[:] = [d for d in dirs if d not in IGNORAR and not d.startswith('$')]
                lower_files = {f.lower(): f for f in files}
                if nome_arquivo.lower() in lower_files:
                    achado = Path(root) / lower_files[nome_arquivo.lower()]
                    print(f"✅ Encontrado em: {achado}")
                    return achado
        except:
            continue
    print(f"❌ '{nome_arquivo}' não encontrado.")
    return None


def escrevendoArquivo_Jason():
    Dados_Entrada = input("Coloque o NOME da planilha APQP (ex: APQP_Projeto_X): ").strip()
    nome_limpo = Path(Dados_Entrada).stem if Dados_Entrada.lower().endswith('.xlsx') else Path(Dados_Entrada).name
    if "\\" in Dados_Entrada or "/" in Dados_Entrada:
        nome_limpo = Path(Dados_Entrada).stem
    dados = {"apqp_nome": nome_limpo}
    with open(CONF_PATH, 'w', encoding='utf-8') as arquivo:
        json.dump(dados, arquivo, indent=4, ensure_ascii=False)
    print(f"💾 Salvo em {CONF_PATH}: {dados}")


def LendoArquivo_Jason():
    if not CONF_PATH.exists():
        return None
    try:
        with open(CONF_PATH, "r", encoding='utf-8') as arquivo:
            conteudo = json.load(arquivo)
            if isinstance(conteudo, str):
                return Path(conteudo).stem
            if isinstance(conteudo, dict):
                return conteudo.get("apqp_nome")
            return conteudo
    except Exception as e:
        print(f"Erro ao ler conf.json: {e}")
        return None


# ---------------------------------------------------------
# 1. FUNÇÕES QUE EDITAM AS PLANILHAS (CORRIGIDAS)
# ---------------------------------------------------------
def prenchendoPlanilha_ACC(caminho_arquivo, PN_MTH_ACC, PN_Cliente_ACC, REV_Acc, RFQ_Acc, Projeto_Acc, ClientePlanta_Acc, VolumeAnual_Acc, DataEntrada_Acc, DataResposta_Acc, TipoItem_Acc, Nome_Peça_Acc, Comprador_Acc):
    Plan_Saida = load_workbook(caminho_arquivo)
    Pagina_Saida = Plan_Saida['Plan1']
    Pagina_Saida['F3'] = str(PN_MTH_ACC) if PN_MTH_ACC else ""
    Pagina_Saida['F6'] = str(PN_Cliente_ACC) if PN_Cliente_ACC else ""
    Pagina_Saida['Q6'] = str(REV_Acc) if REV_Acc else ""
    Pagina_Saida['S3'] = str(RFQ_Acc) if RFQ_Acc else ""
    Pagina_Saida['AD3'] = str(Projeto_Acc) if Projeto_Acc else ""
    Pagina_Saida['AK3'] = str(ClientePlanta_Acc) if ClientePlanta_Acc else ""
    Pagina_Saida['Z6'] = str(Nome_Peça_Acc) if Nome_Peça_Acc else ""
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
    time.sleep(0.2)
    Plan_saida = load_workbook(caminho_arquivo, keep_vba=True)
    pagina_saida = Plan_saida['Planilha de Custo']
    pagina_saida['C6'].value = str(PN_cliente) if PN_cliente else ""
    pagina_saida['C7'].value = str(PN_MethalLsit) if PN_MethalLsit else ""
    pagina_saida['C8'].value = str(Revisao) if Revisao else ""
    pagina_saida['F9'].value = str(Cliente) if Cliente else ""
    pagina_saida['G8'].value = str(Tipo) if Tipo else ""
    pagina_saida['M8'].value = str(Peso) if Peso else ""
    
    # Volume já formatado
    if isinstance(Volume, list):
        pagina_saida['F9'].value = ", ".join([formatar_volume(v) for v in Volume])
    else:
        # Se já for string formatada, mantém
        pagina_saida['F9'].value = formatar_volume(Volume)
    
    Plan_saida.save(caminho_arquivo)


def prenchendoPlanilha_FINN(caminho_Arquivo, PNsList, Cliente, Planta_clinete, Local_clinete, Clinete_Comprador, Projeto, Descricao, Tipo, Rev, PNsListInterno, Volume):
    # CORREÇÃO DO ERRO: tratamento robusto
    Plan_saida = load_workbook(caminho_Arquivo)
    
    # Verifica se aba existe
    if 'FOR.ENG.03' in Plan_saida.sheetnames:
        pagina_saida = Plan_saida['FOR.ENG.03']
    else:
        print(f"⚠️ Aba FOR.ENG.03 não encontrada. Abas disponíveis: {Plan_saida.sheetnames}. Usando aba ativa.")
        pagina_saida = Plan_saida.active
    
    # Escreve sempre como string para evitar erro do openpyxl
    pagina_saida['C6'].value = str(Cliente) if Cliente else ""
    pagina_saida['H6'].value = str(Planta_clinete) if Planta_clinete else ""
    pagina_saida['K6'].value = str(Local_clinete) if Local_clinete else ""
    pagina_saida['C7'].value = str(Clinete_Comprador) if Clinete_Comprador else ""
    pagina_saida['H7'].value = str(Projeto) if Projeto else ""
    pagina_saida['H8'].value = str(Descricao) if Descricao else ""
    pagina_saida['H9'].value = str(Rev) if Rev else ""
    pagina_saida['E9'].value = str(PNsList) if PNsList else ""
    pagina_saida['E8'].value = str(PNsListInterno) if PNsListInterno else ""
    
    # VOLUME - AQUI ESTAVA O ERRO
    try:
        vol_str = formatar_volume(Volume)
        pagina_saida['D15'].value = str(vol_str)
    except Exception as e:
        print(f"⚠️ Erro ao escrever volume {Volume}: {e}. Escrevendo como texto.")
        pagina_saida['D15'].value = str(Volume)
    
    if Tipo == "Novo":
        pagina_saida['C5'].value = "X"
    elif Tipo == "Modificação":
        pagina_saida['F5'].value = "X"
    elif Tipo == "Substitução":
        pagina_saida['I5'].value = "X"
    else:
        if Tipo:
            pagina_saida['H5'].value = str(Tipo)
            pagina_saida['I5'].value = "X"
        
    Plan_saida.save(caminho_Arquivo)


# ---------------------------------------------------------
# 2. LÊ DADOS DA APQP
# ---------------------------------------------------------
def PegandoDados_APQP(APQPSelecionada):
    caminho_apqp = encontrar_apqp_por_nome(APQPSelecionada)
    if not caminho_apqp or not caminho_apqp.exists():
        raise FileNotFoundError(f"Não achei a APQP '{APQPSelecionada}.xlsx' no PC.")
    
    print(f"📂 Usando APQP: {caminho_apqp}")
    Plan_Entrada = xw.Book(str(caminho_apqp))
    Pagina_Entrada = Plan_Entrada.sheets['APQP']
    
    Inicio_ACC_Criar = input("Coloque a linha que iniciará as ACC's: ").strip()
    Fim_ACC_Criar = input("Coloque a última linha: ").strip()
    
    PN_Methal     = Pagina_Entrada.range(f'A{Inicio_ACC_Criar}:A{Fim_ACC_Criar}').value
    Cliente       = Pagina_Entrada.range(f'F{Inicio_ACC_Criar}:F{Fim_ACC_Criar}').value  
    ClientePlanta = Pagina_Entrada.range(f'G{Inicio_ACC_Criar}:G{Fim_ACC_Criar}').value  
    tipo          = Pagina_Entrada.range(f'H{Inicio_ACC_Criar}:H{Fim_ACC_Criar}').value
    RFQ           = Pagina_Entrada.range(f'I{Inicio_ACC_Criar}:I{Fim_ACC_Criar}').value    
    PNs           = Pagina_Entrada.range(f'J{Inicio_ACC_Criar}:J{Fim_ACC_Criar}').value   
    REV           = Pagina_Entrada.range(f'K{Inicio_ACC_Criar}:K{Fim_ACC_Criar}').value
    descricao     = Pagina_Entrada.range(f'L{Inicio_ACC_Criar}:L{Fim_ACC_Criar}').value
    comprador     = Pagina_Entrada.range(f'M{Inicio_ACC_Criar}:M{Fim_ACC_Criar}').value
    VolumeAnual   = Pagina_Entrada.range(f'P{Inicio_ACC_Criar}:P{Fim_ACC_Criar}').value   
    Peso          = Pagina_Entrada.range(f'Q{Inicio_ACC_Criar}:Q{Fim_ACC_Criar}').value   
    Declinar      = Pagina_Entrada.range(f'V{Inicio_ACC_Criar}:V{Fim_ACC_Criar}').value
    
    celulas_entrada = Pagina_Entrada.range(f'Y{Inicio_ACC_Criar}:Y{Fim_ACC_Criar}') 
    DataEntrada = [cell.api.Text for cell in celulas_entrada]
    
    celulas_resposta = Pagina_Entrada.range(f'AA{Inicio_ACC_Criar}:AA{Fim_ACC_Criar}')
    DataResposta = [cell.api.Text for cell in celulas_resposta]

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
    PNsList = [format(pn) for pn in PNsList_Formatar]
    REVList = [format(rev) for rev in REVList_Formatar]
    
    for i in range(len(PNsList)):
        comp = compradorList[i] if i < len(compradorList) else ""
        cli  = ClienteList[i]   if i < len(ClienteList)   else ""
        if comp and cli:
            ListaFormanta_CompradorACC.append(f"{comp}   {cli}")
        else:
            ListaFormanta_CompradorACC.append(f"{comp}{cli}")
    
    return PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, compradorList, ClienteList, PesoList, DeclinarList


# ---------------------------------------------------------
# 3. CRIA PASTAS
# ---------------------------------------------------------
def CriandoPastas(PN_MethalLsit_Criar, PNs_Ciar, rev_Craidas, RFQ_criar, Projeto_criar, ClientePlanta_criar, VolumeAnual_criar, DataEntrada_criar, DataResposta_criar, TipoItem_criar, Nome_Peça_Criar, Comprador_Criar, ClienteList_cria, compradorList_criar, peso_criar, Declinar_crirar):
    
    pastas_Secundarias = [
        "1-RFQ", "2-Desenhos", "3-Analise Critica", "4-Planilha de Custo",
        "5-CBD", "6-Carta Comercial", "8-Terceiros", "9-Pedidos",
        "10-FIIN", "11-Emails", "12-Custo Logistico", "13-Obsoleto", "14-Modificações"
    ] 
    
    if len(PNs_Ciar) == 0:
        print("Ops! A lista de PNs está vazia.")
        return

    for i in range(len(PNs_Ciar)):
        if i < len(Declinar_crirar) and Declinar_crirar[i] == 1:
            print(f"⏩ Pulando {PNs_Ciar[i]} - declinado")
            continue
        pn_limpo = format(PNs_Ciar[i])
        if not pn_limpo:
            continue
        rfq_limpo = format(RFQ_criar[i])
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
            caminho_excel_copiado, PN_MethalLsit_Criar[i], PNs_Ciar[i], rev_Craidas[i], RFQ_criar[i], Projeto_criar[i], ClientePlanta_criar[i], VolumeAnual_criar[i], DataEntrada_criar[i], DataResposta_criar[i], TipoItem_criar[i], Nome_Peça_Criar[i], Comprador_Criar[i]
        )
        
        Novo_PlanilhaCusto = f'{pasta_principal}_Planilha_Custo.xlsm'
        PlanilhaCusto_copiada = os.path.join(caminho_Base, "4-Planilha de Custo", Novo_PlanilhaCusto)
        if not MODELO_CUSTO.exists():
            print(f"❌ Modelo Custo não encontrado")
            continue
        shutil.copy2(MODELO_CUSTO, PlanilhaCusto_copiada)
        prenchendoPlanilha_Custo(PlanilhaCusto_copiada, PNs_Ciar[i], PN_MethalLsit_Criar[i], rev_Craidas[i], VolumeAnual_criar[i], ClienteList_cria[i], TipoItem_criar[i], peso_criar[i])
        print(f"✅ Sucesso: Pasta '{pasta_principal}' criada!")


def CriandoFIIN(PN_MethalLsit_Criar, PNs_Ciar, rev_Craidas, RFQ_criar, ClientePlanta_criar, VolumeAnual_criar, TipoItem_criar, Nome_Peça_Criar, ClienteList_cria, compradorList_criar, Declinar_crirar):

    if len(PNs_Ciar) == 0:
        print("Ops! Lista vazia")
        return

    for i in range(len(PNs_Ciar)):
        if i < len(Declinar_crirar) and Declinar_crirar[i] == 1:
            continue
        pn_limpo = format(PNs_Ciar[i])
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
            Nome_Peça_Criar[i], 
            TipoItem_criar[i], 
            rev_Craidas[i], 
            PN_MethalLsit_Criar[i],
            VolumeAnual_criar[i]
        )
        print(f"✅ FIIN para '{pasta_principal}' criada!")


# =========================================================
# MENU PRINCIPAL
# =========================================================
Play = 0 

while Play == 0:
    print("\nO que gostaria?")
    print("[1] Mudar APQP")
    print("[2] Criar ACC")
    print("[3] Criar FIIN")
    print("[X] SAIR")
    
    opcao_menu = input("_> ").strip().upper()
    
    if opcao_menu == "1":
        escrevendoArquivo_Jason()
        
    elif opcao_menu == "2":
        Caminho = LendoArquivo_Jason()
        if not Caminho:
            print("⚠️ Erro: Nenhum caminho configurado no conf.json. Use a opção [1] primeiro.")
            continue
        try:
            PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, compradorList, ClienteList, PesoList, Declinar = PegandoDados_APQP(Caminho)
            CriandoPastas(
                PN_MethalLsit, PNsList, REVList, RFQList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, ClienteList, compradorList, PesoList, Declinar
            )
        except Exception as e:
            import traceback
            print(f"❌ Erro: {e}")
            traceback.print_exc()

    elif opcao_menu == "3":
        Caminho = LendoArquivo_Jason()
        if not Caminho:
            print("⚠️ Erro: Nenhum caminho configurado no conf.json. Use a opção [1] primeiro.")
            continue
        try:
            PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, compradorList, ClienteList, PesoList, Declinar = PegandoDados_APQP(Caminho)
            CriandoFIIN(
                PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, tipoList, descricaoLista, ClienteList, compradorList, Declinar
            )
        except Exception as e:
            import traceback
            print(f"❌ Erro: {e}")
            traceback.print_exc()
        
    elif opcao_menu == "X":
        Play = 1
        print("Programa encerrado com sucesso!")
