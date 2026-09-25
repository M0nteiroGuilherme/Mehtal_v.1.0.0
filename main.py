import re
import os
import shutil
import pandas as pd
import xlwings as xw
import json
from openpyxl import load_workbook

def format(texto):
    if texto is None:
        return " "
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
    Dados_Entrada ={}
    Dados_Entrada = input("Coloque o caminho: ")

    with open("conf.json", 'w')as arquivo:
        json.dump(Dados_Entrada,arquivo)


def LendoArquivo_Jason():
    with open("conf.json", "r") as arquivo:
        Dados_Saida = json.load(arquivo)

        return Dados_Saida


# ---------------------------------------------------------
# 1. FUNÇÃO QUE EDITA O EXCEL COPIADO
# ---------------------------------------------------------
def prenchendoPlanilha_ACC(caminho_arquivo, PN_MTH_ACC, PN_Cliente_ACC, REV_Acc, RFQ_Acc, Projeto_Acc, ClientePlanta_Acc, VolumeAnual_Acc, DataEntrada_Acc, DataResposta_Acc, TipoItem_Acc, Nome_Peça_Acc, Comprador_Acc):
    
    # Abre o arquivo exato que acabou de ser criado na pasta "3-Analise Critica"
    Plan_Saida = load_workbook(caminho_arquivo)
    Pagina_Saida = Plan_Saida['Plan1']
    
    # Preenche as células (Sem precisar de laço While!)
    Pagina_Saida['F3']  = PN_MTH_ACC       # PN Cliente
    Pagina_Saida['F6']  = PN_Cliente_ACC      # PN Cliente
    Pagina_Saida['Q6']  = REV_Acc         # REV
    Pagina_Saida['S3']  = RFQ_Acc         # RFQ
    Pagina_Saida['AD3'] = Projeto_Acc     # Projeto
    Pagina_Saida['AK3'] = ClientePlanta_Acc # Cliente Planta
    Pagina_Saida['Z6'] = Nome_Peça_Acc # Descrição peça
    Pagina_Saida['AJ6'] = Comprador_Acc # Comprador
    Pagina_Saida['H12'] = str(DataEntrada_Acc).split(' ')[0] # Data de Entrada
    Pagina_Saida['T12'] = str(DataResposta_Acc).split(' ')[0] # Data de Resposta
    
    if isinstance(VolumeAnual_Acc, list):
        Pagina_Saida['F9'] = ", ".join([formatar_volume(v) for v in VolumeAnual_Acc])
    else:
        Pagina_Saida['F9'] = formatar_volume(VolumeAnual_Acc)
           
    
    if(TipoItem_Acc == "Novo"):
        
        Pagina_Saida['S9'] = "X"
        
    elif(TipoItem_Acc == "Protótipo"):
        
            Pagina_Saida['S10'] = "X"    
            
    elif(TipoItem_Acc == "Modificação"):
        
                Pagina_Saida['AA9'] = "X"
    else:
            Pagina_Saida['V10'] = f"{TipoItem_Acc}"
            Pagina_Saida['AA10'] = "X"
                
                        
    
    Plan_Saida.save(caminho_arquivo)


def prenchendoPlanilha_Custo(caminho_arquivo, PN_cliente, PN_MethalLsit, Revisao, Volume, Cliente, Tipo, Peso):
    
    Plan_saida = load_workbook(caminho_arquivo)
    pagina_saida = Plan_saida['Planilha de Custo']
    
    pagina_saida['C6'] = PN_cliente
    pagina_saida['C7'] = PN_MethalLsit
    pagina_saida['C8'] = Revisao
    #pagina_saida['C9'] = DataRevisao
    #pagina_saida['E7'] = DataCotacao
    pagina_saida['F9'] = Cliente
    pagina_saida['G8'] = Tipo
    pagina_saida['M8'] = f"{Peso}"
    
    if isinstance(Volume, list):
        pagina_saida['F9'] = ", ".join([formatar_volume(v) for v in Volume])
    else:
       pagina_saida['F9'] = formatar_volume(Volume)
    
    Plan_saida.save(caminho_arquivo)
    
    
def prenchendoPlanilha_FINN(caminho_Arquivo, PNsList ,Cliente, Planta_clinete, Local_clinete, Clinete_Comprador, Projeto, Descricao, Tipo, Rev, PNsListInterno):
    
    Plan_saida = load_workbook(caminho_Arquivo)
    pagina_saida = Plan_saida ['FOR.ENG.03']
    
    pagina_saida['C6'] = Cliente         # CORRIGIDO: C6 é a célula inicial mesclada de Cliente (era D6)
    pagina_saida['H6'] = Planta_clinete  # Célula ao lado do rótulo "Planta:" (ou H6 dependendo da mesclagem)
    pagina_saida['K6'] = Local_clinete   # Célula ao lado de "Localização:"
    
    pagina_saida['C7'] = Clinete_Comprador  # Célula mesclada do Comprador
    pagina_saida['H7'] = Projeto           # CORRIGIDO: H7 é onde fica o campo Projeto (onde o cursor verde está)
    
    pagina_saida['H8'] = Descricao
    pagina_saida['H9'] = Rev
    pagina_saida['E9'] = PNsList          # Código do Cliente
    pagina_saida['E8'] = PNsListInterno   # Código do Methal
    
    if (Tipo == "Novo"):
        
        pagina_saida['C5'] = "X"
        
    elif(Tipo == "Modificação"):
        
        pagina_saida['F5'] = "X"
        
    elif(Tipo == "Substitução"):
        
        pagina_saida['I5'] = "X"
        
    else:

        pagina_saida['H5'] = f'{Tipo}'
        pagina_saida['I5'] = "X"
        
    Plan_saida.save(caminho_Arquivo)
    
    

# ---------------------------------------------------------
# 2. FUNÇÃO QUE LÊ OS DADOS DO MAIN APQP
# ---------------------------------------------------------

def PegandoDados_APQP(APQPSelecionada):
    
    Plan_Entrada = xw.Book(f'{APQPSelecionada}.xlsx')
    Pagina_Entrada = Plan_Entrada.sheets['APQP']
    
    Inicio_ACC_Criar = input("coloque a linha que iniciara As ACC's: ")
           
    Fim_ACC_Criar = input("Colque a ultima linha: ")
    
    
        
    PN_Methal     = Pagina_Entrada.range (f'A{Inicio_ACC_Criar}:A{Fim_ACC_Criar}').value
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
        if not isinstance(ColumSelecionada,list):
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


    ListaFormanta_CompradorACC = []
    
    PNsList = [format(pn) for pn in PNsList_Formatar]
    REVList = [format(rev) for rev in REVList_Formatar]
    
    for i in range(len(PNsList)):
        comp = compradorList[i] if i < len(compradorList) else ""
        cli  = ClienteList[i]   if i < len(ClienteList)   else ""
    
        # Formata como texto (ex: "João / Empresa X" ou só "João")
        if comp and cli:
            ListaFormanta_CompradorACC.append(f"{comp}   {cli}")
        else:
            ListaFormanta_CompradorACC.append(f"{comp}{cli}")
    
    
    # Removido o .save() pois você só está lendo dados, não editou nada!
    
    return PN_MethalLsit, PNsList, REVList, RFQList,  ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC , compradorList, ClienteList, PesoList, Declinar

# ---------------------------------------------------------
# 3. FUNÇÃO QUE CRIA PASTAS E DISTRIBUI ARQUIVOS
# ---------------------------------------------------------
def CriandoPastas(PN_MethalLsit_Criar, PNs_Ciar, rev_Craidas, RFQ_criar, Projeto_criar, ClientePlanta_criar, VolumeAnual_criar, DataEntrada_criar, DataResposta_criar, TipoItem_criar, Nome_Peça_Criar, Comprador_Criar, ClienteList_cria, compradorList_criar, peso_criar, Declinar_crirar):
    
    pastas_Segudarias = [
        "1-RFQ", "2-Desenhos", "3-Analise Critica", "4-Planilha de Custo",
        "5-CBD", "6-Carta Comercial", "8-Terceiros", "9-Pedidos",
        "10-FIIN", "11-Emails", "12-Custo Logistico", "13-Obsoleto", "14-Modificações"
    ] 
    
    
    # 1. Verifica se há itens na lista (número > 0)
    if len(PNs_Ciar) > 0:
                    
            for i in range(len(PNs_Ciar)):
                    
                if  Declinar_crirar[i] != 1:
                    
                    # Se o PN estiver em branco, pula para o próximo
                    if not str(PNs_Ciar[i]).strip():
                        continue
                        
                    RFQ_CraiarPasta = str(RFQ_criar[i]).strip()
                    pasta_principal = str(PNs_Ciar[i]).strip()
                    
                    # Monta o caminho base (com ou sem pasta RFQ)
                    if RFQ_CraiarPasta:
                        caminho_Base = os.path.join(RFQ_CraiarPasta, pasta_principal)
                    else:
                        caminho_Base = pasta_principal
                    
                    # 2. Cria a pasta do PN e as subpastas
                    os.makedirs(caminho_Base, exist_ok=True)
                    for nome_subpasta in pastas_Segudarias:
                        caminho_Completo = os.path.join(caminho_Base, nome_subpasta)
                        os.makedirs(caminho_Completo, exist_ok=True)

                    # 3. Copia e preenche a ACC UMA ÚNICA VEZ (fora do loop das subpastas)
                    novo_nome_excel = f'{pasta_principal}_REV.{rev_Craidas[i]}_ACC.xlsx'
                    caminho_excel_copiado = os.path.join(caminho_Base, "3-Analise Critica", novo_nome_excel)
                    
                    shutil.copy2('ACC.xlsx', caminho_excel_copiado)
                            
                    prenchendoPlanilha_ACC(
                        caminho_excel_copiado, PN_MethalLsit_Criar[i], PNs_Ciar[i], rev_Craidas[i], RFQ_criar[i], Projeto_criar[i], ClientePlanta_criar[i], VolumeAnual_criar[i], DataEntrada_criar[i], DataResposta_criar[i], TipoItem_criar[i], Nome_Peça_Criar[i], Comprador_Criar[i]
                    )
                                        
                    #Criando Fiin
                    Novo_Fiin = f'{pasta_principal}_FIIN.xlsx'
                    Fiin_copiada = os.path.join(caminho_Base,"10-FIIN", Novo_Fiin)
                    
                    shutil.copy2("FIIN.xlsx", Fiin_copiada)
                    
                    prenchendoPlanilha_FINN(Fiin_copiada, PNs_Ciar[i], ClienteList_cria[i], ClientePlanta_criar[i], ClientePlanta_criar[i], compradorList_criar[i], RFQ_criar[i], Nome_Peça_Criar[i], TipoItem_criar[i], rev_Craidas[i], PN_MethalLsit_Criar[i])
                    
                    Novo_PlanilhaCusto = f'{pasta_principal}_Planilha_Custo.xlsx'
                    PlanilhaCusto_copiada = os.path.join(caminho_Base, "4-Planilha de Custo", Novo_PlanilhaCusto)
                    
                    shutil.copy2("Planilha_Custo.xlsm", PlanilhaCusto_copiada)
                    
                    prenchendoPlanilha_Custo(PlanilhaCusto_copiada, PNs_Ciar[i], PN_MethalLsit_Criar[i], rev_Craidas[i], VolumeAnual_criar[i], compradorList_criar[i], TipoItem_criar[i], peso_criar[i])
                    
                    print(f"Sucesso: Pasta '{pasta_principal}' criada e planilha preenchida!")
                    
    else:
            print("Ops! A lista de PNs está vazia. Nenhum dado encontrado.")

# =========================================================
# O SEU CÓDIGO PRINCIPAL QUE RODA TUDO:
# =========================================================



# =========================================================
# O SEU CÓDIGO PRINCIPAL QUE RODA TUDO:
# =========================================================




Play = 0 

while Play == 0:
    print("\nO que gostaria?")
    print("[1] Mudar APQP")
    print("[2] Criar ACC")
    print("[X] SAIR")
    
    opcao_menu = input("_> ").strip().upper()  # Trata espaços e converte para maiúsculo
    
    if opcao_menu == "1":
        escrevendoArquivo_Jason()
        
    elif opcao_menu == "2":
        Caminho = LendoArquivo_Jason()
        
        if not Caminho:
            print("⚠️ Erro: Nenhum caminho configurado no conf.json. Use a opção [1] primeiro.")
            continue
            
        PN_MethalLsit, PNsList, REVList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, compradorList, ClienteList, PesoList, Declinar = PegandoDados_APQP(Caminho)

        CriandoPastas(
            PN_MethalLsit, PNsList, REVList, RFQList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC, compradorList, ClienteList, PesoList, Declinar
        )

    elif opcao_menu == "X":
        Play = 1
        print("Programa encerrado com sucesso!")