import re
import os
import shutil
import pandas as pd
import xlwings as xw
import json
import time
from pathlib import Path
from openpyxl import load_workbook

# Caminho base do diretório onde o script Python está localizado
BASE_DIR = Path(__file__).resolve().parent

# Caminhos absolutos dos arquivos modelo
MODELO_ACC = BASE_DIR / "ACC.xlsx"
MODELO_CUSTO = BASE_DIR / "Planilha_Custo.xlsm"
MODELO_FIIN = BASE_DIR / "FIIN.xlsx"


def format(texto):
    if texto is None:
        return ""
    return re.sub(r'[\\/*?:"<>|]', '_', str(texto)).strip()


def formatar_volume(valor):
    """Converte float/int em texto com separador de milhar (ex: 1111.0 -> 1,111)."""
    if valor is None or str(valor).strip() == "":
        return ""
    try:
        num = int(float(valor))
        return f"{num:,}"  # Se preferir o padrão brasileiro (1.111), use: f"{num:,}".replace(',', '.')
    except (ValueError, TypeError):
        return str(valor)


def escrevendoArquivo_Jason():
    Dados_Entrada = input("Coloque o caminho da planilha APQP (sem a extensão .xlsx): ").strip()

    with open(BASE_DIR / "conf.json", 'w') as arquivo:
        json.dump(Dados_Entrada, arquivo)


def LendoArquivo_Jason():
    caminho_conf = BASE_DIR / "conf.json"
    if not caminho_conf.exists():
        return None

    with open(caminho_conf, "r") as arquivo:
        return json.load(arquivo)


# ---------------------------------------------------------
# 1. FUNÇÕES QUE EDITAM AS PLANILHAS COPIADAS
# ---------------------------------------------------------
def prenchendoPlanilha_ACC(caminho_arquivo, PN_MTH_ACC, PN_Cliente_ACC, REV_Acc, RFQ_Acc, Projeto_Acc, ClientePlanta_Acc, VolumeAnual_Acc, DataEntrada_Acc, DataResposta_Acc, TipoItem_Acc, Nome_Peça_Acc, Comprador_Acc):
    Plan_Saida = load_workbook(caminho_arquivo)
    Pagina_Saida = Plan_Saida['Plan1']
    
    Pagina_Saida['F3']  = PN_MTH_ACC       # PN Methal
    Pagina_Saida['F6']  = PN_Cliente_ACC   # PN Cliente
    Pagina_Saida['Q6']  = REV_Acc          # REV
    Pagina_Saida['S3']  = RFQ_Acc          # RFQ
    Pagina_Saida['AD3'] = Projeto_Acc      # Projeto
    Pagina_Saida['AK3'] = ClientePlanta_Acc # Cliente Planta
    Pagina_Saida['Z6']  = Nome_Peça_Acc     # Descrição peça
    Pagina_Saida['AJ6'] = Comprador_Acc     # Comprador
    Pagina_Saida['H12'] = str(DataEntrada_Acc).split(' ')[0] # Data de Entrada
    Pagina_Saida['T12'] = str(DataResposta_Acc).split(' ')[0] # Data de Resposta
    
    if isinstance(VolumeAnual_Acc, list):
        Pagina_Saida['F9'] = ", ".join([formatar_volume(v) for v in VolumeAnual_Acc])
    else:
        Pagina_Saida['F9'] = formatar_volume(VolumeAnual_Acc)
            
    if TipoItem_Acc == "Novo":
        Pagina_Saida['S9'] = "X"
    elif TipoItem_Acc == "Protótipo":
        Pagina_Saida['S10'] = "X"    
    elif TipoItem_Acc == "Modificação":
        Pagina_Saida['AA9'] = "X"
    else:
        Pagina_Saida['V10'] = f"{TipoItem_Acc}"
        Pagina_Saida['AA10'] = "X"
        
    Plan_Saida.save(caminho_arquivo)


def prenchendoPlanilha_Custo(caminho_arquivo, PN_cliente, PN_MethalLsit, Revisao, Volume, Cliente, Tipo, Peso):
    time.sleep(0.5)  # Pausa para sincronização do arquivo
    
    Plan_saida = load_workbook(caminho_arquivo, keep_vba=True)
    pagina_saida = Plan_saida['Planilha de Custo']
    
    pagina_saida['C6'] = PN_cliente
    pagina_saida['C7'] = PN_MethalLsit
    pagina_saida['C8'] = Revisao
    pagina_saida['F9'] = Cliente
    pagina_saida['G8'] = Tipo
    pagina_saida['M8'] = f"{Peso}"
    
    if isinstance(Volume, list):
        pagina_saida['F9'] = ", ".join([formatar_volume(v) for v in Volume])
    else:
        pagina_saida['F9'] = formatar_volume(Volume)
    
    Plan_saida.save(caminho_arquivo)


def prenchendoPlanilha_FINN(caminho_Arquivo, PNsList, Cliente, Planta_clinete, Local_clinete, Clinete_Comprador, Projeto, Descricao, Tipo, Rev, PNsListInterno, Volume):
    Plan_saida = load_workbook(caminho_Arquivo)
    pagina_saida = Plan_saida['FOR.ENG.03']
    
    pagina_saida['C6'] = Cliente
    pagina_saida['H6'] = Planta_clinete
    pagina_saida['K6'] = Local_clinete
    pagina_saida['C7'] = Clinete_Comprador
    pagina_saida['H7'] = Projeto
    pagina_saida['H8'] = Descricao
    pagina_saida['H9'] = Rev
    pagina_saida['E9'] = PNsList          # Código do Cliente
    pagina_saida['E8'] = PNsListInterno   # Código do Methal
    
    if isinstance(Volume, list):
        pagina_saida['C15'] = ", ".join([formatar_volume(v) for v in Volume])
    else:
        pagina_saida['C15'] = formatar_volume(Volume)
    
    if Tipo == "Novo":
        pagina_saida['C5'] = "X"
    elif Tipo == "Modificação":
        pagina_saida['F5'] = "X"
    elif Tipo == "Substitução":
        pagina_saida['I5'] = "X"
    else:
        pagina_saida['H5'] = f'{Tipo}'
        pagina_saida['I5'] = "X"
        
    Plan_saida.save(caminho_Arquivo)


# ---------------------------------------------------------
# 2. FUNÇÃO QUE LÊ OS DADOS DO MAIN APQP
# ---------------------------------------------------------
def PegandoDados_APQP(APQPSelecionada):
    caminho_apqp = BASE_DIR / f'{APQPSelecionada}.xlsx' if not str(APQPSelecionada).endswith('.xlsx') else BASE_DIR / str(APQPSelecionada)
    
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
# 3. FUNÇÕES QUE CRIAM PASTAS E GERAM OS ARQUIVOS
# ---------------------------------------------------------
def CriandoPastas(PN_MethalLsit_Criar, PNs_Ciar, rev_Craidas, RFQ_criar, Projeto_criar, ClientePlanta_criar, VolumeAnual_criar, DataEntrada_criar, DataResposta_criar, TipoItem_criar, Nome_Peça_Criar, Comprador_Criar, ClienteList_cria, compradorList_criar, peso_criar, Declinar_crirar):
    
    pastas_Secundarias = [
        "1-RFQ", "2-Desenhos", "3-Analise Critica", "4-Planilha de Custo",
        "5-CBD", "6-Carta Comercial", "8-Terceiros", "9-Pedidos",
        "10-FIIN", "11-Emails", "12-Custo Logistico", "13-Obsoleto", "14-Modificações"
    ] 
    
    if len(PNs_Ciar) > 0:
        for i in range(len(PNs_Ciar)):
            
            if i < len(Declinar_crirar) and Declinar_crirar[i] == 1:
                continue
                
            pn_limpo = format(PNs_Ciar[i])
            if not pn_limpo:
                continue
                
            rfq_limpo = format(RFQ_criar[i])
            pasta_principal = pn_limpo
            
            if rfq_limpo:
                caminho_Base = os.path.join(rfq_limpo, pasta_principal)
            else:
                caminho_Base = pasta_principal
            
            # Cria a estrutura de pastas
            os.makedirs(caminho_Base, exist_ok=True)
            for nome_subpasta in pastas_Secundarias:
                caminho_Completo = os.path.join(caminho_Base, nome_subpasta)
                os.makedirs(caminho_Completo, exist_ok=True)

            # 1. Copia e preenche a ACC
            novo_nome_excel = f'{pasta_principal}_REV.{rev_Craidas[i]}_ACC.xlsx'
            caminho_excel_copiado = os.path.join(caminho_Base, "3-Analise Critica", novo_nome_excel)
            
            shutil.copy2(MODELO_ACC, caminho_excel_copiado)
            
            prenchendoPlanilha_ACC(
                caminho_excel_copiado, PN_MethalLsit_Criar[i], PNs_Ciar[i], rev_Craidas[i], RFQ_criar[i], Projeto_criar[i], ClientePlanta_criar[i], VolumeAnual_criar[i], DataEntrada_criar[i], DataResposta_criar[i], TipoItem_criar[i], Nome_Peça_Criar[i], Comprador_Criar[i]
            )
            
            # 2. Copia e preenche a Planilha de Custo (CORRIGIDO: Sem espaço no .xlsm)
            Novo_PlanilhaCusto = f'{pasta_principal}_Planilha_Custo.xlsm'
            PlanilhaCusto_copiada = os.path.join(caminho_Base, "4-Planilha de Custo", Novo_PlanilhaCusto)
            
            shutil.copy2(MODELO_CUSTO, PlanilhaCusto_copiada)
            
            prenchendoPlanilha_Custo(PlanilhaCusto_copiada, PNs_Ciar[i], PN_MethalLsit_Criar[i], rev_Craidas[i], VolumeAnual_criar[i], ClienteList_cria[i], TipoItem_criar[i], peso_criar[i])
            
            print(f"Sucesso: Pasta '{pasta_principal}' criada e planilhas preenchidas!")
    else:
        print("Ops! A lista de PNs está vazia. Nenhum dado encontrado.")


def CriandoFIIN(PN_MethalLsit_Criar, PNs_Ciar, rev_Craidas, RFQ_criar, ClientePlanta_criar, VolumeAnual_criar, TipoItem_criar, Nome_Peça_Criar, ClienteList_cria, compradorList_criar, Declinar_crirar):

    if len(PNs_Ciar) > 0:
        for i in range(len(PNs_Ciar)):
            
            if i < len(Declinar_crirar) and Declinar_crirar[i] == 1:
                continue

            pn_limpo = format(PNs_Ciar[i])
            if not pn_limpo:
                continue

            pasta_principal = pn_limpo
            
            # Coloca dentro da subpasta 10-FIIN se ela existir, senão salva na pasta raiz
            caminho_destino_pasta = os.path.join(pasta_principal, "10-FIIN") if os.path.exists(pasta_principal) else pasta_principal
            os.makedirs(caminho_destino_pasta, exist_ok=True)

            Novo_Fiin = f"{pasta_principal}_FIIN.xlsx"
            Fiin_copiada = os.path.join(caminho_destino_pasta, Novo_Fiin)

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
            print(f"Sucesso: FIIN para '{pasta_principal}' criada com sucesso!")


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
            
        PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, compradorList, ClienteList, PesoList, Declinar = PegandoDados_APQP(Caminho)

        CriandoPastas(
            PN_MethalLsit,               # PN_MethalLsit_Criar
            PNsList,                     # PNs_Ciar
            REVList,                     # rev_Craidas
            RFQList,                     # RFQ_criar
            RFQList,                     # Projeto_criar
            ClientePlantaList,           # ClientePlanta_criar
            VolumeAnualList,             # VolumeAnual_criar
            DataEntradaList,             # DataEntrada_criar
            DataRespostaList,            # DataResposta_criar
            tipoList,                    # TipoItem_criar
            descricaoLista,              # Nome_Peça_Criar
            ListaFormanta_CompradorACC,  # Comprador_Criar
            ClienteList,                 # ClienteList_cria
            compradorList,               # compradorList_criar
            PesoList,                    # peso_criar
            Declinar                     # Declinar_crirar
        )

    elif opcao_menu == "3":
        Caminho = LendoArquivo_Jason()
        
        if not Caminho:
            print("⚠️ Erro: Nenhum caminho configurado no conf.json. Use a opção [1] primeiro.")
            continue
        
        PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, compradorList, ClienteList, PesoList, Declinar = PegandoDados_APQP(Caminho)

        CriandoFIIN(
            PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, tipoList, descricaoLista, ClienteList, compradorList, Declinar
        )
        
    elif opcao_menu == "X":
        Play = 1
        print("Programa encerrado com sucesso!")