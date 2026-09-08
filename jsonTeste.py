import json 

def escrevendoArquivo_Jason():
    Dados_Entrada ={}
    Dados_Entrada = input("Coloque o caminho: ")

    with open("conf.json", 'w')as arquivo:
        json.dump(Dados_Entrada,arquivo)


def LendoArquivo_Jason():
    with open("conf.json", "r") as arquivo:
        Dados_Saida = json.load(arquivo)

        return Dados_Saida

Caminho = LendoArquivo_Jason()

print(Caminho)