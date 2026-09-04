import os
import shutil
import pandas as pd
import xlwings as xw
import json
from openpyxl import load_workbook

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
    Pagina_Saida['F9']  = VolumeAnual_Acc # Volume Anual
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
        Linhas = 0
        listaLimpa = []
        
        if ColumSelecionada is None:
            return [], 0
            
        for Dado in ColumSelecionada:  
            if Dado is not None and str(Dado).strip() != '':                                                 
                listaLimpa.append(Dado)        
                Linhas += 1      
        return listaLimpa, Linhas
    
    # PEGANDO AS VARIÁVEIS USANDO O "_" PARA IGNORAR AS LINHAS EXTRAS
    ObservacaoList, Linhas = PegarInf(Observacao)
    PNsList, _             = PegarInf(PNs)
    REVList, _             = PegarInf(REV)
    RFQList, _             = PegarInf(RFQ)
    ProjetoList, _         = PegarInf(Projeto)
    ClientePlantaList, _   = PegarInf(ClientePlanta)
    Desenhos2DList, _      = PegarInf(Desenhos2D)
    Desenhos3DList, _      = PegarInf(Desenhos3D)
    VolumeAnualList, _     = PegarInf(VolumeAnual)
    DataEntradaList, _     = PegarInf(DataEntrada)
    DataRespostaList, _    = PegarInf(DataResposta)
    QTDLista, _            = PegarInf(QTD)
    ComponentesLista, _    = PegarInf(Componentes)
    ReponsavelLista, _     = PegarInf(Reponsavel)
    realizadoLista,_       = PegarInf(realizado)
    ClienteLista,_         = PegarInf(Cliente)
    tipoLista,_            = PegarInf(tipo)
    descricaoLista,_       = PegarInf(descricao)
    compradorLista,_       = PegarInf(comprador)
    PesoPecaLista,_        = PegarInf(PesoPeca)

    # Removido o .save() pois você só está lendo dados, não editou nada!
    
    return ObservacaoList, Linhas, PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, Desenhos2DList, Desenhos3DList, VolumeAnualList, DataEntradaList, DataRespostaList, QTDLista, ComponentesLista, ReponsavelLista,realizadoLista, ClienteLista, tipoLista, descricaoLista, compradorLista, PesoPecaLista


# ---------------------------------------------------------
# 3. FUNÇÃO QUE CRIA PASTAS E DISTRIBUI ARQUIVOS
# ---------------------------------------------------------
def CriandoPastas(PNs_Ciar, rev_Craidas, RFQ_criar, Projeto_criar, ClientePlanta_criar, VolumeAnual_criar, DataEntrada_criar, DataResposta_criar):
    
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
            prenchendoPlanilha_ACC(caminho_excel_copiado, PNs_Ciar[i], rev_Craidas[i], RFQ_criar[i], Projeto_criar[i], ClientePlanta_criar[i], VolumeAnual_criar[i], DataEntrada_criar[i], DataResposta_criar[i])
            
            print(f"Sucesso: Pasta '{pasta_principal}' criada e planilha preenchida!")
            
    else:
        print("Ops! A lista de PNs está vazia. Nenhum dado encontrado.") 

# =========================================================
# O SEU CÓDIGO PRINCIPAL QUE RODA TUDO:
# =========================================================

# 1. Puxa os dados do Excel "Ao vivo"
ObservacaoList, Linhas, PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, Desenhos2DList, Desenhos3DList, VolumeAnualList, DataEntradaList, DataRespostaList, QTDLista, ComponentesLista, ReponsavelLista,realizadoLista, ClienteLista, tipoLista, descricaoLista, compradorLista, PesoPecaLista = PegandoDados_PQPMain('MainAPQP')

# 2. Cria as pastas e edita os arquivos copiados usando as listas
CriandoPastas(PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, VolumeAnualList, DataEntradaList, DataRespostaList)

def PreenchendoAPQP (APQP_Input):
    print('celula para adidicinoar: ')
    CelulaInico = input()
    #pede a celula para começar a adição ex b15 Ele soma com o I e adicoina apartir dai usando o "PegandoDados_PQPMain"
    
    ObservacaoList, Linhas, PNsList, REVList, RFQList, ProjetoList, ClientePlantaList, Desenhos2DList, Desenhos3DList, VolumeAnualList, DataEntradaList, DataRespostaList, QTDLista, ComponentesList, ReponsavelList,RealizadoList, ClienteList, tipoList, descricaoList, compradorList, PesoPecaList = PegandoDados_PQPMain('MainAPQP')
    
    Tamanho_list = len(PNsList)
    i = 0
    
    Plan_Entrada = xw.Book(f'{APQP_Input}.xlsx')
    Pagina_Entrada = Plan_Entrada.sheets['APQP']
        
    
    while (Tamanho_list > i):    
    
        Pagina_Entrada.cells['B'] = RealizadoList
        Pagina_Entrada.cells['C'] = ObservacaoList
        Pagina_Entrada.cells['F'] = ClienteList
        Pagina_Entrada.cells['G'] = ClientePlantaList
        Pagina_Entrada.cells['H'] = tipoList
        Pagina_Entrada.cells['I'] = RFQList
        Pagina_Entrada.cells['J'] = ItemList
        Pagina_Entrada.cells['K'] = REVList
        Pagina_Entrada.cells['L'] = descricaoList
        Pagina_Entrada.cells['M'] = compradorList
        Pagina_Entrada.cells['N'] = Desenhos2DList
        Pagina_Entrada.cells['O'] = Desenhos3DList
        Pagina_Entrada.cells['P'] = VolumeAnualList
        Pagina_Entrada.cells['P'] = PesoPecaList
        Pagina_Entrada.cells['Y'] = DataEntradaList
        Pagina_Entrada.cells['Z'] = AberturaList
        Pagina_Entrada.cells['AA'] = DataRespostaList

    
        
    
    