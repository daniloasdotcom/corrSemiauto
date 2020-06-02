
import openpyxl
import os

#1 Insira o caminho para seu banco de dados
os.chdir("C:\\Users\\Usuário\\Desktop")

#2 Diga o nome do seu banco de dados, o nome da planilha e o total de variáveis que seu arquivo de dados possui
wb = openpyxl.load_workbook('example.xlsx')
sheet = wb['Planilha1']
totalVar = 8


listStocks = []


for i in range(1, (totalVar + 1), 1):
    nameStock = sheet.cell(row=1, column=i).value
    nameStock = str(nameStock)
    listStocks.append(nameStock)


texto = str(
    '---\n'
    #3 Nas três próximas linhas você pode editar o titulo do seu documento, nome do autor e data
    'title: "Correlação"\n'
    'author: "Danilo Andrade"\n'
    'date: "27 de maio de 2020"\n'
    'output: word_document\n'
    '---\n'
    '\n'
    '```{r setup, include=FALSE}\n'
    'knitr::opts_chunk$set(echo = TRUE)\n'
    '```\n'
    '\n'
    '## Abaixo, a tabela de dados\n'
    '\n'
    '```{r echo=FALSE}\n'
    'library(xlsx)\n'
    'library(corrplot)\n'
    'getwd()\n'
    
    #4 O endereço abaixo indicará para o RSudio onde está o seu arquivo
    # Você deve editar ele também
    'setwd("C:/Users/Usuário/Desktop")\n'
    
    #5 Diga o nome do seu banco de dados em formato .csv
    'dados<-read.xlsx("example.xlsx", sheetName = "Planilha1")\n'
    'dados\n'
    'str(dados)\n'
    '```\n'
    '\n'
    '## Tabela de correlação\n'
    '\n'
    '```{r echo=FALSE}\n'
    'cor(dados, use = "complete.obs")\n'
    'tabela<- cor(dados, use = "complete.obs")\n'
    'write.xlsx(tabela, file = "tabela.xlsx")\n'
    '```\n'
    '\n'
    '```{r echo=FALSE}\n'
    'corrplot(tabela)\n'
    '```\n'
    '\n')

j = len(listStocks)

i = j
n = 1
m = 0
y = 1

arquivo = open('scriptRmarkdown.md', 'w')
arquivo.write(texto)

while i > 1:
    n = m + 1
    while n != (len(listStocks)):
        arquivo.write("**Correlação nº" + str(y) + " - " + listStocks[m] + " x " + listStocks[n] + "**")
        arquivo.write('\n')
        arquivo.write('\n')
        arquivo.write("**Coeficiente de correlação:**")
        arquivo.write('\n')
        arquivo.write('```{r echo=FALSE}')
        arquivo.write('\n')
        arquivo.write("corr = cor.test(dados$" + listStocks[m] + ",dados$" + listStocks[n] + ")" + '\n')
        arquivo.write("corr$estimate" + '\n')
        arquivo.write('\n')
        arquivo.write('```')
        arquivo.write('\n')
        arquivo.write("**Significância:**")
        arquivo.write('\n')
        arquivo.write('\n')
        arquivo.write('```{r echo=FALSE}')
        arquivo.write('\n')
        arquivo.write("corr = cor.test(dados$" + listStocks[m] + ",dados$" +listStocks[n] +")" + '\n')
        arquivo.write("corr$p.value" + '\n')
        arquivo.write('\n')
        arquivo.write('```')
        arquivo.write('\n')

        n += 1
        y = y + 1

    i = i - 1
    m = m + 1