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
def prenchendoPlanilha_ACC(caminho_arquivo, PN_MTH_ACC, PN_Cliente_ACC, REV_Acc, RFQ_Acc, Projeto_Acc, ClientePlanta_Acc, VolumeAnual_Acc, DataEntrada_Acc, DataResposta_Acc, TipoItem_Acc, Nome_Peça_Acc, Comprador_Acc):
    
    # Abre o arquivo exato que acabou de ser criado na pasta "3-Analise Critica"
    Plan_Saida = load_workbook(caminho_arquivo)
    Pagina_Saida = Plan_Saida['Plan1']
    
    # Preenche as células (Sem precisar de laço While!)
    Pagina_Saida['B3']  = PN_MTH_ACC       # PN Cliente
    Pagina_Saida['F6']  = PN_Cliente_ACC      # PN Cliente
    Pagina_Saida['Q6']  = REV_Acc         # REV
    Pagina_Saida['S3']  = RFQ_Acc         # RFQ
    Pagina_Saida['AD3'] = Projeto_Acc     # Projeto
    Pagina_Saida['AK3'] = ClientePlanta_Acc # Cliente Planta
    Pagina_Saida['Z6'] = Nome_Peça_Acc # Descrição peça
    Pagina_Saida['AJ6'] = Comprador_Acc # Comprador
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
def PegandoDados_APQP(APQPSelecionada):
    
    Plan_Entrada = xw.Book(f'{APQPSelecionada}.xlsx')
    Pagina_Entrada = Plan_Entrada.sheets['APQP']
    
    print("coloque a coluna que iniciara As ACC's: ")
    
    Inicio_ACC_Criar = input("")
    
    print("Colque a ultima colum")
    Fim_ACC_Criar = input("")
    
        
    PN_Methal     =Pagina_Entrada.range (f'A{Inicio_ACC_Criar}A').value
    #realizado     = Pagina_Entrada.range(f'B{Inicio_ACC_Criar}:B{Fim_ACC_Criar}').value     
    #Observacao    = Pagina_Entrada.range(f'C{Inicio_ACC_Criar}:C{Fim_ACC_Criar}').value  
    Cliente       = Pagina_Entrada.range(f'F{Inicio_ACC_Criar}:F{Fim_ACC_Criar}').value  
    ClientePlanta = Pagina_Entrada.range(f'G{Inicio_ACC_Criar}:G{Fim_ACC_Criar}').value  
    tipo          = Pagina_Entrada.range(f'H{Inicio_ACC_Criar}:H{Fim_ACC_Criar}').value
    RFQ           = Pagina_Entrada.range(f'I{Inicio_ACC_Criar}:I{Fim_ACC_Criar}').value    
    PNs           = Pagina_Entrada.range(f'J{Inicio_ACC_Criar}:J{Fim_ACC_Criar}').value   
    REV           = Pagina_Entrada.range(f'K{Inicio_ACC_Criar}:K{Fim_ACC_Criar}').value
    descricao     = Pagina_Entrada.range(f'L{Inicio_ACC_Criar}:l{Fim_ACC_Criar}').value
    comprador     = Pagina_Entrada.range(f'M{Inicio_ACC_Criar}:m{Fim_ACC_Criar}').value
    #Desenhos3D    = Pagina_Entrada.range(f'M{Inicio_ACC_Criar}:M{Fim_ACC_Criar}').value
    #Desenhos2D    = Pagina_Entrada.range(f"N{Inicio_ACC_Criar}:N{Fim_ACC_Criar}").value
    VolumeAnual   = Pagina_Entrada.range(f'P{Inicio_ACC_Criar}:P{Fim_ACC_Criar}').value   
    #PesoPeca      = Pagina_Entrada.range(f'P{Inicio_ACC_Criar}:P{Fim_ACC_Criar}').value   
    #QTD           = Pagina_Entrada.range(f'R{Inicio_ACC_Criar}:R{Fim_ACC_Criar}').value   
    #Componentes   = Pagina_Entrada.range(f'S{Inicio_ACC_Criar}:S{Fim_ACC_Criar}').value   
    DataEntrada   = Pagina_Entrada.range(f'Y{Inicio_ACC_Criar}:Y{Fim_ACC_Criar}').value # Corrigido para U
    DataResposta  = Pagina_Entrada.range(f'AA{Inicio_ACC_Criar}:AA{Fim_ACC_Criar}').value
    #Reponsavel    = Pagina_Entrada.range(f'X{Inicio_ACC_Criar}:X{Fim_ACC_Criar}').value # Corrigido espaço
    
    def PegarInf(ColumSelecionada):

        listaLimpa = []
        
        if ColumSelecionada is None:
            return []
            
        for Dado in ColumSelecionada:  
            if Dado is not None and str(Dado).strip() != '':                                                 
                listaLimpa.append(Dado)         
        return listaLimpa
    
    # PEGANDO AS VARIÁVEIS USANDO O "_" PARA IGNORAR A EXTRAS
    PN_MethalLsit = PegarInf(PN_Methal)
    #ObservacaoList = PegarInf(Observacao)
    PNsList            = PegarInf(PNs)
    REVList            = PegarInf(REV)
    RFQList            = PegarInf(RFQ)
    #ProjetoList        = PegarInf(Projeto)
    ClientePlantaList  = PegarInf(ClientePlanta)
    #Desenhos2DList     = PegarInf(Desenhos2D)
    #Desenhos3DList     = PegarInf(Desenhos3D)
    VolumeAnualList    = PegarInf(VolumeAnual)
    DataEntradaList    = PegarInf(DataEntrada)
    DataRespostaList   = PegarInf(DataResposta)
    #TDLista           = PegarInf(QTD)
    #ComponentesLista   = PegarInf(Componentes)
    #ReponsavelLista    = PegarInf(Reponsavel)
    #realizadoLista      = PegarInf(realizado)
    ClienteList         = PegarInf(Cliente)
    tipoList            = PegarInf(tipo)
    descricaoLista       = PegarInf(descricao)
    compradorList       = PegarInf(comprador)
    #PesoPecaLista        = PegarInf(PesoPeca)

    ListaFormanta_CompradorACC =list(zip(compradorList,ClienteList))
    
    
    # Removido o .save() pois você só está lendo dados, não editou nada!
    
    return PN_MethalLsit, PNsList, REVList, RFQList,  ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC 

# ---------------------------------------------------------
# 3. FUNÇÃO QUE CRIA PASTAS E DISTRIBUI ARQUIVOS
# ---------------------------------------------------------
def CriandoPastas(PN_MethalLsit_Criar, PNs_Ciar, rev_Craidas, RFQ_criar, Projeto_criar, ClientePlanta_criar, VolumeAnual_criar, DataEntrada_criar, DataResposta_criar, TipoItem_criar, Nome_Peça_Criar, Comprador_Criar):
    
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
            prenchendoPlanilha_ACC(caminho_excel_copiado, PN_MethalLsit_Criar[i], PNs_Ciar[i], rev_Craidas[i], RFQ_criar[i], Projeto_criar[i], ClientePlanta_criar[i], VolumeAnual_criar[i], DataEntrada_criar[i], DataResposta_criar[i], TipoItem_criar[i], Nome_Peça_Criar[i], Comprador_Criar[i])
            
            print(f"Sucesso: Pasta '{pasta_principal}' criada e planilha preenchida!")
            
    else:
        print("Ops! A lista de PNs está vazia. Nenhum dado encontrado.") 

# =========================================================
# O SEU CÓDIGO PRINCIPAL QUE RODA TUDO:
# =========================================================

PN_MethalLsit, PNsList, REVList, RFQList,  ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC = PegandoDados_APQP()

CriandoPastas(PN_MethalLsit, PNsList, REVList, RFQList, RFQList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList, tipoList, descricaoLista, ListaFormanta_CompradorACC)