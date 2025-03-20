# ConsultaMEI-ME
Sistema de automação para consulta de optantes MEI/ME através do CNPJ


O sistema acessa o site https://www8.receita.fazenda.gov.br/simplesnacional/aplicacoes.aspx?id=21 e verifica uma lista de CNPJ's fornecida pelo usuário, gerando uma planilha que contém: Data da Consulta, CNPJ, Situacao atual simples, Situacao atual Simei, Periodos Anteriores Simples e Periodos Anteriores Simei

# Como usar (executável):
Coloque os CNPJ's no arquivo "lista.txt" dentro da pasta dist/automacao, um cnpj por linha, e rode o arquivo automacao.exe

# Como usar (Arquivo .py)
Garanta que tenha "Python" instalado, coloque os CNPJ's no arquivo "lista.txt" na pasta raiz, abra o terminal na pasta raiz e digite "python automacao.py"
