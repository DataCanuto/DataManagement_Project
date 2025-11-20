import os
import shutil
#Caminho absoluto onde será criado o diretório raiz
root_dir = r"D:\Documentos\JOBS\DATA\RelatórioAndiappToledo\Plano de Ações\Processos Anddiap"

#Criando referência para arquivo modelo (Planilha_Automatizada)
modelo = r"D:\Documentos\JOBS\DATA\RelatórioAndiappToledo\Plano de Ações\Planilha_Automatizada.xlsx"

#Criar diretório raiz se não existir
os.makedirs(root_dir, exist_ok=True)

#Criar 100 pastas de clientes e copiar a planilha modelo para dentro delas
for i in range (1,101):
    cliente_dir = os.path.join(root_dir, f'Cliente{i}')
    os.makedirs(cliente_dir, exist_ok=True) #Cria uma pasta para cada cliente se ela não existir

    #caminho de destino da Planilha_Automatizada
    destino = os.path.join(cliente_dir, "Planilha_Automatizada.xlsx")

    #copiar planilha modelo para a pasta de cada cliente sem sobrescrever caso exista
    if not os.path.exists(destino):
        shutil.copy(modelo, destino)

print(f'Diretório criado com sucesso em: {root_dir}')

