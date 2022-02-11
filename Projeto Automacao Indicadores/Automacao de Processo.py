#!/usr/bin/env python
# coding: utf-8

# ### Passo 1 - Importar Arquivos e Bibliotecas

# In[ ]:


# Importando as bibliotecas

import pandas as pd
import pathlib
import win32com.client as win32


# In[ ]:


# Importando e tratando as bases de dados

emails = pd.read_excel(r'C:\Users\Marcelo Desktop\Downloads\Projeto AutomacaoIndicadores\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx')
emails = emails.drop(['Unnamed: 3','Unnamed: 4'], axis=1)
vendas = pd.read_excel(r'C:\Users\Marcelo Desktop\Downloads\Projeto AutomacaoIndicadores\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx')
lojas = pd.read_csv(r'C:\Users\Marcelo Desktop\Downloads\Projeto AutomacaoIndicadores\Projeto AutomacaoIndicadores\Bases de Dados\Lojas.csv', sep=';', encoding='latin1')

vendas = vendas.merge(lojas, on='ID Loja')


# ### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

# In[ ]:


# Criando um dicionário com todas as tabelas por loja

dicionario_lojas = {}

for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja,:]

# Obtendo o dia que serão analisados os indicadores
    
dia_indicador = vendas['Data'].max()


# ### Passo 3 - Salvar a planilha na pasta de backup

# In[ ]:


# Criando o caminho, pastas e salvando as tabelas de cada loja como backup

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    
    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    
    dicionario_lojas[loja].to_excel(local_arquivo)


# ### Passo 4 - Calcular o indicador para 1 loja

# In[ ]:


# Definição das metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500 
meta_ticketmedio_ano = 500


# In[ ]:


for loja in dicionario_lojas:
    
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador,:]

    # Calculo do faturamento

    faturamento_ano = vendas_loja['Valor Final'].sum() 
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()

    # Calculo da diversidade de produtos

    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())

    # Calculo do valor do ticket médio

    valor_venda = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    # Enviando os emails
    
    nome = emails.loc[emails['Loja']==loja,'Gerente'].values[0]

    outlook = win32.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja']==loja,'E-mail'].values[0]
    mail.Subject = 'OnePage Dia {}/{} - Loja {}'.format(dia_indicador.day, dia_indicador.month, loja)

    if faturamento_dia >= meta_faturamento_dia:
            cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLBody = f'''

    <p>Bom dia, {nome}</p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>loja {loja}</strong> foi:</p>

    <table>
          <tr>
            <th>Indicador</th>
            <th>Valor Dia</th>
            <th>Meta Dia</th>
            <th>Cenário Dia</th>
          </tr>
          <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamento_dia:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
          </tr>
          <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtde_produtos_dia}</td>
            <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
            <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
          </tr>
          <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
            <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
          </tr>
        </table>
        <br>
        <table>
          <tr>
            <th>Indicador</th>
            <th>Valor Ano</th>
            <th>Meta Ano</th>
            <th>Cenário Ano</th>
          </tr>
          <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamento_ano:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
          </tr>
          <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtde_produtos_ano}</td>
            <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
            <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
          </tr>
          <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
            <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
          </tr>
        </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou a disposição.</p>

    <p>Att., Marcelo</p>

    '''

    attachment = pathlib.Path.cwd() / caminho_backup / loja / '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    mail.Attachments.Add(str(attachment))
    mail.Send()
    
    print('E-mail da loja {} enviado.'.format(loja))
    


# ### Passo 5 - Enviar por e-mail para o gerente

# ### Passo 6 - Automatizar todas as lojas

# ### Passo 7 - Criar ranking para diretoria

# In[27]:


faturamento_lojas = vendas.groupby('Loja')[['Loja','Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))

vendas_dia = vendas.loc[vendas['Data']==dia_indicador,:]
vendas_dia = vendas_dia.groupby('Loja')[['Loja','Valor Final']].sum()
faturamento_lojas_dia = vendas_dia.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking Diário.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


# ### Passo 8 - Enviar e-mail para diretoria

# In[30]:


outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja']=='Diretoria','E-mail'].values[0]
mail.Subject = 'Ranking Dia {}/{}'.format(dia_indicador.day, dia_indicador.month)
mail.Body = f'''

Prezados, bom dia

Melhor loja do dia em faturamento: Loja {faturamento_lojas_dia.index[0]}, com faturamento de R${faturamento_lojas_dia.iloc[0,0]:.2f}.
Pior loja do dia em faturamento: Loja {faturamento_lojas_dia.index[-1]}, com faturamento de R${faturamento_lojas_dia.iloc[-1,0]:.2f}.

Melhor loja do ano em faturamento: Loja {faturamento_lojas_ano.index[0]}, com faturamento de R${faturamento_lojas_ano.iloc[0,0]:.2f}.
Pior loja do ano em faturamento: Loja {faturamento_lojas_ano.index[-1]}, com faturamento de R${faturamento_lojas_ano.iloc[-1,0]:.2f}.

Segue em anexo os rankings do ano e do dia com todas as lojas.

Qualquer dúvida estou a disposição.

Att.,
Marcelo

'''

attachment = pathlib.Path.cwd() / caminho_backup / '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
mail.Attachments.Add(str(attachment))

attachment = pathlib.Path.cwd() / caminho_backup / '{}_{}_Ranking Diário.xlsx'.format(dia_indicador.month, dia_indicador.day)
mail.Attachments.Add(str(attachment))

mail.Send()

print('E-mail da diretoria enviado.')


# In[ ]:




