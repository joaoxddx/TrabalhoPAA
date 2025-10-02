# TrabalhoPAA
Trabalho de Projeto de analise de algoritimo

Bibliotecas nessesarias:
    pandas: ler a tabela de execel
    tkinter: para abrir o popup de escolha de arquivo

precisa ser usado a partir de um arquivo XLS ou XLSX, e devolve um arquivo em mesmo formato, JÁ FORMATADO

ACOMPANHA UM ARQUIVO DE TESTES (TODOS OS DADOS DAQUELE ARQUIVO SAO RETIRADOS DO GEMINI, FIZ A TABELA USANDO A IA DO GOOGLE SHEETS PARA TESTES DO CODIGO)

Complexidade
    O(NlogN)

Explicacao do codigo
    Bloco 1
        O objetivo é fazer com que o usuario escolha o arquivo que será usado na compilacao
    Bloco 2
        O objetivo é fazer com que o usuario escolha qual a coluna da tabela será ordenada
    Bloco 3
        O objetivo é fazer com que o usuario escolha qual o tipo do dado será ultilizado para a organizacao, Int ou String(o de int ignora todos os caracteres nao numerais)
    Bloco 4
        seleciona o merge a ser ultilizado corforme o bloco 3
    Bloco 5
        le a tabela e monta uma lista encadeada(de acordo com a escolha feita no bloco 3)  e monta uma lista de indice q é o tamanho da lista de valores.
    Bloco 6: Merge Sort
        Bloco 6.1
            Divide ambas as listas passada ao meio 
        Bloco 6.2
            junta sublistas já organizadas, comparando os elementos que ha compoem 
        Bloco 6.3
            Trabalha com os elementos restantes a esquerda
        Bloco 6.4
            Trabalha com os elementos restantes a direita
    Bloco 7
        realiza a devolucao do arquivo ordenado