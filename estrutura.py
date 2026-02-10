import os

def gerar_estrutura_txt(diretorio_raiz, arquivo_saida):
    with open(arquivo_saida, "w", encoding="utf-8") as f:
        for raiz, pastas, arquivos in os.walk(diretorio_raiz):
            nivel = raiz.replace(diretorio_raiz, "").count(os.sep)
            indentacao = "    " * nivel
            f.write(f"{indentacao}{os.path.basename(raiz)}/\n")

            sub_indentacao = "    " * (nivel + 1)
            for arquivo in arquivos:
                f.write(f"{sub_indentacao}{arquivo}\n")

if __name__ == "__main__":
    diretorio = r"C:/Users/42706671840/Documents/Automacao_Atestados"   # ajuste o caminho
    saida_txt = "estrutura_diretorio.txt"
    gerar_estrutura_txt(diretorio, saida_txt)
    print("Arquivo TXT gerado com sucesso.")
