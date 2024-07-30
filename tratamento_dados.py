import pandas as pd
import matplotlib.pyplot as plt

try:
    df = pd.read_excel(r'C:\Users\jbern\Downloads\Projeto\Big Data em Python\Vendas.xlsx')
    print("Arquivo lido com sucesso!")
except FileNotFoundError:
    print("Erro: Arquivo não encontrado. Verifique o caminho.")
except Exception as e:
    print(f"Erro ao ler o arquivo: {e}")

print(df.head())

print(df.isnull().sum())

df['Data da Venda'].fillna(method='ffill', inplace=True)

df['Data da Venda'] = pd.to_datetime(df['Data da Venda'], format='%d/%m/%Y')
df['Hora da Venda'] = pd.to_datetime(df['Hora da Venda'], format='%H:%M').dt.time

vendas_por_produto = df['Produto'].value_counts()
print(vendas_por_produto)

vendas_por_dia = df.groupby('Dia').size()
print(vendas_por_dia)

plt.figure(figsize=(10, 6))
vendas_por_produto.plot(kind='bar', color='skyblue')
plt.title('Vendas por Produto')
plt.xlabel('Produto')
plt.ylabel('Quantidade de Vendas')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(r'C:\Users\jbern\Downloads\Projeto\Big Data em Python\Vendas_por_Produto.png')
plt.show()

plt.figure(figsize=(10, 6))
vendas_por_dia.plot(kind='bar', color='lightgreen')
plt.title('Vendas por Dia')
plt.xlabel('Dia')
plt.ylabel('Quantidade de Vendas')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(r'C:\Users\jbern\Downloads\Projeto\Big Data em Python\Vendas_por_Dia.png')
plt.show()

vendas_por_dia_produto = df.groupby(['Dia', 'Produto']).size().unstack().fillna(0)
print(vendas_por_dia_produto)

vendas_por_dia_produto.plot(kind='bar', stacked=True, figsize=(12, 7), colormap='viridis')
plt.title('Quantidade de Produtos Vendidos por Dia da Semana')
plt.xlabel('Dia da Semana')
plt.ylabel('Quantidade de Vendas')
plt.xticks(rotation=45)
plt.legend(title='Produto', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.tight_layout()
plt.savefig(r'C:\Users\jbern\Downloads\Projeto\Big Data em Python\Vendas_por_Produto_Por_Dia.png')
plt.show()

try:
    df.to_excel(r'C:\Users\jbern\Downloads\Projeto\Big Data em Python\Tratamento_Dados.xlsx', index=False)
    print("Arquivo salvo com sucesso!")
except Exception as e:
    print(f"Erro ao salvar o arquivo: {e}")