import os
import shutil
import pandas as pd
import xlwings as xw
import json
from openpyxl import load_workbook

def escrevendoArquivo_Jason():
    Dados_Entrada ={}
    Dados_Entrada = input("Coloque o caminho: ")

    with open("conf.json", 'w')as arquivo:
        json.dump(Dados_Entrada,arquivo)


def LendoArquivo_Jason():
    with open("conf.json", "r") as arquivo:
        Dados_Saida = json.load(arquivo)

        return Dados_Saida

escrevendoArquivo_Jason()
Caminho = LendoArquivo_Jason()

print(Caminho)

pasta_busca = os.path.expanduser("~")
nome_arquivo_alvo = Caminho

for root, dirs, files in os.walk(pasta_busca):
    if nome_arquivo_alvo in files:
        caminhoAPQP = os.path.join(root, nome_arquivo_alvo)
        print(f"Arquivo localizado em: {caminhoAPQP}")
        break  # Para a busca assim que encontrar o primeiro arquivo


# ---------------------------------------------------------
# 1. FUNÇÃO QUE EDITA O EXCEL COPIADO
# ---------------------------------------------------------
def prenchendoPlanilha_ACC(caminho_arquivo, PN_Cliente, REV_Acc, RFQ_Acc, Projeto_Acc, ClientePlanta_Acc, VolumeAnual_Acc, DataEntrada_Acc, DataResposta_Acc, TipoItem_Acc):
    
    # Abre o arquivo exato que acabou de ser criado na pasta "3-Analise Critica"
    Plan_Saida = load_workbook(caminho_arquivo)
    Pagina_Saida = Plan_Saida['Plan1']
    
    # Preenche as células (Sem precisar de laço While!)
    Pagina_Saida['F6']  = PN_Cliente      # PN Cliente
    Pagina_Saida['Q6']  = REV_Acc         # REV
    Pagina_Saida['S3']  = RFQ_Acc         # RFQ
    Pagina_Saida['AD3'] = Projeto_Acc     # Projeto
    Pagina_Saida['AK3'] = ClientePlanta_Acc # Cliente Planta
    Pagina_Saida['F9']  = ", ".join(map(str, VolumeAnual_Acc))if isinstance( VolumeAnual_Acc, list) else VolumeAnual_Acc # Volume Anual
    Pagina_Saida['H12'] = DataEntrada_Acc # Data de Entrada
    Pagina_Saida['T12'] = DataResposta_Acc# Data de Resposta
    
    if(TipoItem_Acc == "Novo"):
        
        Pagina_Saida['S9'] = "X"
        
    elif(TipoItem_Acc == "Protótipo"):
        
            Pagina_Saida['S10'] = "X"    
            
    elif(TipoItem_Acc == "Modificação"):
        
                Pagina_Saida['AA9'] = "X"
    else:
            Pagina_Saida['V10'] = TipoItem_Acc
            Pagina_Saida['AA10'] = "X"
                
                        
    
    Plan_Saida.save(caminho_arquivo)


# ---------------------------------------------------------
# 2. FUNÇÃO QUE LÊ OS DADOS DO MAIN APQP
# ---------------------------------------------------------
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
    DataEntrada   = Pagina_Entrada.range('U4:U1000').value # Corrigido para U
    DataResposta  = Pagina_Entrada.range('W4:W1000').value
    Reponsavel    = Pagina_Entrada.range('X4:X1000').value # Corrigido espaço
    
    def PegarInf(ColumSelecionada):

        listaLimpa = []
        
        if ColumSelecionada is None:
            return []
            
        for Dado in ColumSelecionada:  
            if Dado is not None and str(Dado).strip() != '':                                                 
                listaLimpa.append(Dado)         
        return listaLimpa
    
    # PEGANDO AS VARIÁVEIS USANDO O "_" PARA IGNORAR A EXTRAS
    ObservacaoList = PegarInf(Observacao)
    PNsList            = PegarInf(PNs)
    REVList            = PegarInf(REV)
    RFQList            = PegarInf(RFQ)
    ProjetoList        = PegarInf(Projeto)
    ClientePlantaList  = PegarInf(ClientePlanta)
    Desenhos2DList     = PegarInf(Desenhos2D)
    Desenhos3DList     = PegarInf(Desenhos3D)
    VolumeAnualList    = PegarInf(VolumeAnual)
    DataEntradaList    = PegarInf(DataEntrada)
    DataRespostaList   = PegarInf(DataResposta)
    QTDLista           = PegarInf(QTD)
    ComponentesLista   = PegarInf(Componentes)
    ReponsavelLista    = PegarInf(Reponsavel)
    realizadoLista      = PegarInf(realizado)
    ClienteLista         = PegarInf(Cliente)
    tipoLista           = PegarInf(tipo)
    descricaoLista       = PegarInf(descricao)
    compradorLista       = PegarInf(comprador)
    PesoPecaLista        = PegarInf(PesoPeca)

    # Removido o .save() pois você só está lendo dados, não editou nada!
    
    return ObservacaoList, PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, Desenhos2DList, Desenhos3DList, VolumeAnualList, DataEntradaList, DataRespostaList, QTDLista, ComponentesLista, ReponsavelLista,realizadoLista, ClienteLista, tipoLista, descricaoLista, compradorLista, PesoPecaLista


# ---------------------------------------------------------
# 3. FUNÇÃO QUE CRIA PASTAS E DISTRIBUI ARQUIVOS
# ---------------------------------------------------------
def CriandoPastas(PNs_Ciar, rev_Craidas, RFQ_criar, Projeto_criar, ClientePlanta_criar, VolumeAnual_criar, DataEntrada_criar, DataResposta_criar, TipoItem_criar):
    
    pastas_Segudarias = ["1-RFQ","2-Desenhos","3-Analise Critica","4-Planilha de Custo","5-CBD","6-Carta Comercial","8-Terceiros","9-Pedidos","10-FII","11-Emails","12-Custo Logistico","13-Obsoleto","14-Modificações"]
    
    if len(PNs_Ciar) > 0:
        
        # Como os dados já são Listas, o laço 'for' navega direto por eles.
        for i in range(len(PNs_Ciar)):
            
            # Pega os dados exatos desta iteração do laço
            pasta_principal = str(PNs_Ciar[i])
            
            # Cria a pasta raiz
            os.makedirs(pasta_principal, exist_ok=True)
            
            # Cria as subpastas
            for nome_subpasta in pastas_Segudarias:
                caminho_Completo = os.path.join(pasta_principal, nome_subpasta)
                os.makedirs(caminho_Completo, exist_ok=True)

            # Define o nome exato de onde o arquivo deve parar
            novo_nome_excel = f'{pasta_principal}_REV.{rev_Craidas[i]}_ACC.xlsx'
            caminho_excel_copiado = os.path.join(pasta_principal, "3-Analise Critica", novo_nome_excel)
            
            # Copia o modelo para a pasta correta
            shutil.copy2('ACC.xlsx', caminho_excel_copiado)
                      
            # Edita o Excel QUE ACABOU DE SER COPIADO usando os dados da lista
            prenchendoPlanilha_ACC(caminho_excel_copiado, PNs_Ciar[i], rev_Craidas[i], RFQ_criar[i], Projeto_criar[i], ClientePlanta_criar[i], VolumeAnual_criar[i], DataEntrada_criar[i], DataResposta_criar[i], TipoItem_criar[i])
            
            print(f"Sucesso: Pasta '{pasta_principal}' criada e planilha preenchida!")
            
    else:
        print("Ops! A lista de PNs está vazia. Nenhum dado encontrado.") 

# =========================================================
# O SEU CÓDIGO PRINCIPAL QUE RODA TUDO:
# =========================================================

# 1. Puxa os dados do Excel "Ao vivo"
ObservacaoList, PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, Desenhos2DList, Desenhos3DList, VolumeAnualList, DataEntradaList, DataRespostaList, QTDLista, ComponentesLista, ReponsavelLista,realizadoLista, ClienteLista, tipoLista, descricaoLista, compradorLista, PesoPecaLista = PegandoDados_PQPMain('MainAPQP')

# 2. Cria as pastas e edita os arquivos copiados usando as listas
CriandoPastas(PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoLista)

def PreenchendoAPQP (APQP_Input):
    print('celula para adidicinoar: ')
    CelulaInico = input()
    #pede a celula para começar a adição ex b15 Ele soma com o I e adicoina apartir dai usando o "PegandoDados_PQPMain"
    
    ObservacaoList, PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, Desenhos2DList, Desenhos3DList, VolumeAnualList, DataEntradaList, DataRespostaList, QTDList, ComponentesList, ReponsavelList,RealizadoList, ClienteList, tipoList, descricaoList, compradorList, PesoPecaList = PegandoDados_PQPMain('MainAPQP')
    
    Tamanho_list = len(PNsList)
    i = 0
    
    Plan_Entrada = xw.Book(str(APQP_Input))
    Pagina_Entrada = Plan_Entrada.sheets['APQP']
    
    while (Tamanho_list > i):    
        
        NumeroCelula = int(CelulaInico) + i
        
        Pagina_Entrada.range(f'B{NumeroCelula}').value = RealizadoList   [i]
        Pagina_Entrada.range(f'C{NumeroCelula}').value = ObservacaoList  [i]
        Pagina_Entrada.range(f'F{NumeroCelula}').value = ClienteList [i]
        Pagina_Entrada.range(f'G{NumeroCelula}').value = ClientePlantaList   [i]
        Pagina_Entrada.range(f'H{NumeroCelula}').value = tipoList    [i]
        Pagina_Entrada.range(f'I{NumeroCelula}').value = RFQList [i]
        Pagina_Entrada.range(f'J{NumeroCelula}').value = PNsList [i]
        Pagina_Entrada.range(f'K{NumeroCelula}').value = REVList [i]
        Pagina_Entrada.range(f'L{NumeroCelula}').value = descricaoList   [i]
        Pagina_Entrada.range(f'M{NumeroCelula}').value = compradorList   [i]
        Pagina_Entrada.range(f'N{NumeroCelula}').value = Desenhos2DList  [i]
        Pagina_Entrada.range(f'O{NumeroCelula}').value = Desenhos3DList  [i]
        Pagina_Entrada.range(f'P{NumeroCelula}').value = VolumeAnualList [i]
        Pagina_Entrada.range(f'Q{NumeroCelula}').value = PesoPecaList    [i]
        Pagina_Entrada.range(f'S{NumeroCelula}').value = QTDList [i]
        Pagina_Entrada.range(f'T{NumeroCelula}').value = ComponentesList [i]
        Pagina_Entrada.range(f'Y{NumeroCelula}').value = DataEntradaList [i]
        Pagina_Entrada.range(f'Z{NumeroCelula}').value = ReponsavelList  [i]
        Pagina_Entrada.range(f'AA{NumeroCelula}').value  = DataRespostaList   [i]
        
        i += 1
    
PreenchendoAPQP(caminhoAPQP)