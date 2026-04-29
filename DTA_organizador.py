#Teste para o DTA -> programação orientada a definição/transformação/ação
#este programa lê arquivos, converte-os e pode retorna-los em outro formato, além de mostrar dados filtrados e
#poder modifica-los, leia a descrição no final do código
import plotly.express as px
import pandas as pd
import os
import ast
from tkinter import simpledialog, messagebox, filedialog, scrolledtext, ttk
import tkinter as tk

root = tk.Tk()#esconde uma janela fantasma mas útil do tkinter
root.withdraw()
#Ddadosglobais:
graficospossiveis = ['linhas','barras','dispersao','pizza','histograma','caixa']
graficosDadoscriadosAut = []
#Ddadosglobaisfim-

class DgraficoDados():
    def __init__(self):#assim cria por objeto, e não para toda a classe
        self.eixoXvalores = []#preferencia de numeros
        self.eixoYvalores = []
        self.arquivoX = None #o arquivo csv ou txt precisa vir com uma coluna de dados com um por linha
        self.arquivoY = None
        self.arquivoGeral = None #arquivo que vem com várias colunas
    def AVerArquivos(self):
        if self.arquivoX is not None:
            print(self.arquivoX, " - Fim arquivo do eixo X")
        if self.arquivoY is not None:
            print(self.arquivoY, " - Fim arquivo do eixo Y")
        if self.arquivoGeral is not None:
            print(self.arquivoGeral, " - Fim arquivo geral")
    def AVerEixos(self):
        print(self.eixoXvalores)
        print(self.eixoYvalores)

#txt, csv e outros (uma coluna por arquivo)

def T_tratarDados(defqualdataframe, qualeixo):
    verificarSeSaoNumericos = pd.to_numeric(defqualdataframe[qualeixo], errors='coerce')
    if not verificarSeSaoNumericos.isna().all():#se forem numeros, conseguirá converter, se não, irão ficar nulos
        defqualdataframe[qualeixo] = pd.to_numeric(defqualdataframe[qualeixo], errors='coerce')#vira numerico
        defqualdataframe[qualeixo] = defqualdataframe[qualeixo].fillna(0)#NULL vira 0
        defqualdataframe = defqualdataframe.sort_values(by = qualeixo, ascending=False)#ordena
    return defqualdataframe #retorna numeros se correto e como veio se não

def D_arquivotxt_arquivografico(nomedoarquivo):
    try:
        with open(f'Python/{nomedoarquivo}', 'r') as t:
            return t.read()
    except FileNotFoundError:
        return f"Erro T_arquivo_dado: O arquivo {nomedoarquivo} não foi encontrado."

def D_arquivotxt_eixosgrafico(nomedoarquivo):#nome completo, só entra coluna unica (exemplo.txt)
    try:
        with open(f'Python/{nomedoarquivo}', 'r') as t:
            linhas = t.read().splitlines()
            try:
                return [int(valor) for valor in linhas if valor.strip()]
            except ValueError:
                return [valor for valor in linhas if valor.strip()]#se não for int, vira strings
    except FileNotFoundError:
        print(f"Erro: O arquivo {nomedoarquivo} não foi encontrado.")
        return []
    except ValueError:
        print(f"Erro: O arquivo {nomedoarquivo} teve problema com os dados.")
        return []   

def T_adicionarLinha_arquivotxt(nomedoarquivo):
    def D_nova_linha():
        valor = input(f"Digite o valor para adicionar ao seu arquivo").strip()
        if not valor or valor.upper() == "NULL":
            return None
        try:
            if "." in valor:
                return float(valor)
            else:
                return int(valor)
        except ValueError:
            return valor
    nova_linha = D_nova_linha()
    try:
        nomedoarquivo.append(nova_linha) #concatena
        return nomedoarquivo
    except Exception as e:
        print(f"Erro ao adicionar linha: {e}")
        return nomedoarquivo      

#com Pandas

def D_arquivoExcel_arquivografico_comPandas(nomedocaminhodoarquivo):
    try:
        df = pd.read_excel(nomedocaminhodoarquivo) #se for sem cabeçario -> , header=None 
        return df
    except FileNotFoundError:
        print(f"Erro: O arquivo {nomedocaminhodoarquivo} não foi encontrado.")
        return None

def D_arquivotxt_arquivografico_comPandas(nomedocaminhodoarquivo):
    try:
        df = pd.read_csv(nomedocaminhodoarquivo, sep=None, engine='python') #se for sem cabeçario -> , header=None 
        return df
    except FileNotFoundError:
        print(f"Erro: O arquivo {nomedocaminhodoarquivo} não foi encontrado.")
        return None
    
def T_SubstituirDadosDeColuna_arquivograficoGeral_comPandas(qualgraficoInstancia, colunabusca , valorbusca, valornovo):
    df = qualgraficoInstancia.arquivoGeral
    if colunabusca not in df.columns:
        print(f"Erro: A coluna '{colunabusca}' não existe no arquivo.")
        return df
    tipo_da_coluna = df[colunabusca].dtype
    if pd.api.types.is_datetime64_any_dtype(tipo_da_coluna):#ve se é uma data
        try:#se for, transforma os valores realmente para data
            valorbusca = pd.to_datetime(valorbusca)
            valornovo = pd.to_datetime(valornovo)
        except:
            print("Erro: Formato de data inválido!")
            return df
    tipo_busca = pd.Series([valorbusca]).dtype #pd.series é 1 coluna individual
    tipo_novo = pd.Series([valornovo]).dtype
    if tipo_busca == tipo_da_coluna and tipo_da_coluna == tipo_novo:
        df.loc[df[colunabusca] == valorbusca, colunabusca] = valornovo
        return df
    else:
        print(f"Erro: Você inseriu tipos incompatíveis com a coluna escolhida!")
        return df
    
def T_adicionarLinha_arquivoExcel_comPandas_input(df):#insira o arquivo onde está salvo, cuidado que não verifica chaves primarias
    def D_dados_nova_linha():
        nova_linha_dict = {}
        for coluna in df.columns:#para cada coluna
            valor = input(f"Digite o valor para a coluna [{coluna}]: ")
            try:
                if "." in valor:#tenta converter
                    valor = float(valor)
                else:
                    valor = int(valor)
            except ValueError:
                try:
                    valor_final = pd.to_datetime(valor, dayfirst=True, errors='coerce')
                    if not pd.isna(valor_final):#verifica se deu certo
                        valor = valor_final
                except:
                    pass # Mantém como string se não deu certo
            nova_linha_dict[coluna] = valor
        return nova_linha_dict
    
    dados_nova_linha = D_dados_nova_linha()
    try:
        nova_linha = pd.DataFrame([dados_nova_linha])
        df_atualizado = pd.concat([df, nova_linha], ignore_index=True)#concatena
        return df_atualizado
    except Exception as e:
        print(f"Erro ao adicionar linha: {e}")
        return df
    
def T_adicionarLinha_arquivoExcel_comPandas(df, **dic_dados_nova_linha):#recebe na entrada (ex:Produto="bola" , Preço=20.0)    
    colunas_extras = set(dic_dados_nova_linha.keys()) - set(df.columns)
    if colunas_extras:
        print(f"Erro: A(s) coluna(s) {colunas_extras} não existe(m) no arquivo! (Verifique se escreveu corretamente)")
        return df
    try:
        colunas_faltantes = set(df.columns) - set(dic_dados_nova_linha.keys())
        if colunas_faltantes:
            print(f"As colunas {colunas_faltantes} ficarão vazias (NaN) por não terem dados para inserir.")
        nova_linha = pd.DataFrame([dic_dados_nova_linha])
        df_atualizado = pd.concat([df, nova_linha], ignore_index=True)#concatena
        return df_atualizado
    except Exception as e:
        print(f"Erro ao adicionar linha: {e}")
        return df
    
def T_arquivografico_eixosgrafico_comPandas(nomedoarquivo):
    #if not nomedoarquivo or nomedoarquivo == "": return [] #se estiver vazio
    if nomedoarquivo is None or (isinstance(nomedoarquivo, pd.DataFrame) and nomedoarquivo.empty):
        return []
    try:#vê se veio do open
        return [int(v) for v in nomedoarquivo.splitlines()]#num
    except ValueError:
        return [v for v in nomedoarquivo.splitlines()]#string
    except AttributeError: #o arquivo compilado veio do pandas, a maneira que ele armazena é diferente do open
        coluna = nomedoarquivo.iloc[:, 0]#transforma a primeira (e única) coluna em uma lista
        coluna_numerica = pd.to_numeric(coluna, errors='coerce')#verifica se é de números
        if not coluna_numerica.isna().all(): #se tudo não ficou nulo, deu certo
            return coluna_numerica.fillna(0).tolist()
        else:
            return coluna.tolist()#string
        
def A_mostrarOuAlterarArquivoLinhasComFiltro_comPandas(qualgraficoInstancia, colunabusca , valorbusca, mostrarqntdelinhas=False, seQuiserRetornarAListaComFiltroBoteTrue=False):
    #para apenas buscar, chame a função com 3 argumentos, para alterar o arquivo, peça para o arquivo receber o 
    #valor desta função, com mais dois argumentos no final, o primeiro sendo false e o segundo sendo True (o primeiro é para retornar a quantidade de linhas, o segundo é para retornar o arquivo mostrado)
    #receber o valor sem argumento deixará o arquivo vazio!
    df = qualgraficoInstancia.arquivoGeral
    if colunabusca not in df.columns:
        print(f"Erro: A coluna '{colunabusca}' não existe no arquivo.")
        # df == valorbusca retorna uma tabela de True/False
        # .any(axis=1) retorna True se houver pelo menos um True na linha
        filtro = (df == valorbusca).any(axis=1)
        return df.loc[filtro]
    if pd.api.types.is_datetime64_any_dtype(df[colunabusca].dtype):#se for data
        valorbusca = pd.to_datetime(valorbusca)
    filtro = df[colunabusca] == valorbusca#retorna true nas linhas com a busca
    resultado = df.loc[filtro] #retorna só aonde é true
    if resultado.empty:
        print(f"Nenhum registro encontrado para '{valorbusca}' na coluna '{colunabusca}'.")
    else:
        if mostrarqntdelinhas:
            print(resultado)
            messagebox.showinfo("Linhas retornadas",f"Foram encontradas {len(resultado)} linhas!")
        if seQuiserRetornarAListaComFiltroBoteTrue != False:
            return resultado

def A_arquivotxt_arquivoExcel_comPandas(nomedoarquivo):#insira o nome completo(exemplo.txt)
    try:
        df = pd.read_csv(f'Python/{nomedoarquivo}', sep=None, engine='python') #se for sem cabeçario -> , header=None  
        nome_saida = nomedoarquivo.replace('.txt', '').replace('.csv', '') + '.xlsx' #ajuste de nome
        try:#ve se o arquivo já existe
            versejatem = pd.read_excel(f'Python/{nome_saida}')
            resp = input(f'Já existe um arquivo com esse nome, o existente será sobrescrito, tem certeza? (s ou n)').lower()
            if resp == 's':
                raise FileNotFoundError#finge que não existe, mas vai sobrescrever
        except FileNotFoundError: #se o arquivo resultado não existe, já faz direto
            caminho_saida = f'Python/{nome_saida}'
            df.to_excel(caminho_saida, index=False, engine='openpyxl')
            print(f"Arquivo exportado como: {caminho_saida}")
            return df
    except FileNotFoundError:#caso não encontre o arquivo para converter
        print(f"Erro: O arquivo {nomedoarquivo} não foi encontrado.")
        return None

def A_arquivoExcel_arquivotxt_comPandas(nomedoarquivoexcel, novonome):
    if not ".xlsx" in nomedoarquivoexcel:
        nomedoarquivoexcel = nomedoarquivoexcel + ".xlsx"
    try:
        tabela = pd.read_excel(f'Python/{nomedoarquivoexcel}')
        nome_final = novonome + '.txt'
        try:#ve se o arquivo já existe
            versejatem = pd.read_csv(f'Python/{nome_final}')
            resp = input(f'Já existe um arquivo com esse nome, o existente será sobrescrito, tem certeza? (s ou n)').lower()
            if resp == 's':
                raise FileNotFoundError#finge que não existe, mas vai sobrescrever
        except FileNotFoundError: #se o arquivo resultado não existe, já faz direto
            tabela.to_csv(f'Python/{nome_final}', sep=";", index=False)#separar por TAB use sep="\t"
            print(f"Arquivo exportado para txt como {nome_final}")        
    except FileNotFoundError:#caso não encontre o arquivo para converter
        print(f"Erro: O arquivo {nomedoarquivoexcel} não foi encontrado.")
        return None
    
def A_arquivografico_arquivoExcel_comPandas(qualarquivo, nomeaserdado, nomecaminho):
    df = qualarquivo
    if df is not None: 
        nome_saida = nomeaserdado + '.xlsx' #ajuste de nome
        try:#ve se o arquivo já existe
            versejatem = pd.read_excel(f'{nomecaminho}/{nome_saida}')#vê se vai
            resp = messagebox.askyesno("Confirmação", "Já existe um arquivo com esse nome, o existente será sobrescrito, tem certeza?")
            if resp:
                raise FileNotFoundError#finge que não existe, mas vai sobrescrever
        except FileNotFoundError: #se o arquivo resultado não existe, já faz direto
            caminho_saida = f'{nomecaminho}/{nome_saida}'
            try:
                df.to_excel(caminho_saida, index=False, engine='openpyxl')
                print(f"Arquivo exportado como: {caminho_saida}")
                return True
            except AttributeError:#transforma string em uma lista, se necessario
                lista_real = ast.literal_eval(qualarquivo)
                df = pd.DataFrame(lista_real, columns=['Valores'])
                df.to_excel(caminho_saida, index=False, engine='openpyxl')
                print(f"Arquivo exportado como: {caminho_saida}")
                return True #antigamente retornava df nos trues, vê se pode, mudei para a opção 4 na automação
    else:
        print(f"Erro: O arquivo do gráfico está vazio!")
        return None

def A_arquivografico_arquivotxt_comPandas(qualarquivo, nomeaserdado, nomecaminho):
    df = qualarquivo
    if df is not None: 
        nome_saida = nomeaserdado + '.txt' #ajuste de nome
        caminho_saida = os.path.join(nomecaminho, nome_saida)
        if os.path.exists(caminho_saida):
            resp = messagebox.askyesno("Confirmação", f"Já existe um arquivo {nome_saida}, o existente será sobrescrito, tem certeza?")
            if resp:
                return False          
        try:
            if isinstance(df, pd.DataFrame):
                df.to_csv(caminho_saida, index=False, sep='\t', encoding='utf-8')
            else:
                # Se não for DF, tenta converter a string/lista
                lista_real = ast.literal_eval(str(qualarquivo))
                df_novo = pd.DataFrame(lista_real, columns=['Valores'])
                df_novo.to_csv(caminho_saida, index=False, sep='\t', encoding='utf-8')#não aceita sep=None
            print(f"Arquivo exportado com sucesso em: {caminho_saida}")
            return True
        except Exception as e:
            print(f"Erro ao exportar: {e}")
            return False
    else:
        print(f"Erro: O arquivo do gráfico está vazio!")
        return None

def A_verGrafico(qualgraficoInstancia, qualgraficotipo):
    data = {'Eixo X': qualgraficoInstancia.eixoXvalores, 'Eixo Y': qualgraficoInstancia.eixoYvalores} #necessária essa separação com as variáveis
    df = pd.DataFrame(data) 
    df = T_tratarDados(df,'Eixo X')#trata cada coluna separadamente
    df = T_tratarDados(df,'Eixo Y')
    match qualgraficotipo:
        case 'linhas':
            fig = px.line(df, x='Eixo X', y='Eixo Y', title="Gráfico de Linhas")
            fig.show()
        case 'barras':
            fig = px.bar(df, x='Eixo X', y='Eixo Y', title="Gráfico de Barras")
            fig.show()
        case _:
            print("Gráfico não encontrado ou digitado incorretamente, tente um desses -->", *graficospossiveis, sep=" ")# o * trata cada argumento separadamente

def A_verGraficodoGeral(qualgraficoInstancia, qualgraficotipo, eixoX, eixoY):#bote por nome de cada parte desejada
    data = {'Eixo X': qualgraficoInstancia.arquivoGeral[f'{eixoX}'], 'Eixo Y': qualgraficoInstancia.arquivoGeral[f'{eixoY}']} #necessária essa separação com as variáveis
    df = pd.DataFrame(data) 
    df = T_tratarDados(df,'Eixo X')#trata cada coluna separadamente
    df = T_tratarDados(df,'Eixo Y')
    match qualgraficotipo:
        case 'linhas':
            fig = px.line(df, x='Eixo X', y='Eixo Y', title="Gráfico de Linhas")
            fig.show()
        case 'barras':
            fig = px.bar(df, x='Eixo X', y='Eixo Y', title="Gráfico de Barras")
            fig.show()
        case 'dispersao':
            fig = px.scatter(
            df, 
            x="Eixo X", 
            y="Eixo Y", 
            size="Eixo X", #O tamanho da bola varia de acordo com os valores de X
            title="Relação entre X e Y")
            fig.show()
        case 'pizza':
            fig = px.pie(df, values='Eixo X', names='Eixo Y', 
            title='Gráfico de pizza',
            hole=0)
            fig.show()
        case 'histograma':#conta valores repetidos
            fig = px.histogram(df, x='Eixo X', 
                   nbins=30,#largura das barras 
                   marginal="rug", # Adiciona uma 'franja' para ver densidade
                   title="Gráfico de histograma")
            fig.show()
        case 'caixa':
            fig = px.box(df, x= 'Eixo X', y='Eixo Y',title="Gráfico de caixa")
            fig.update_traces(line_width=10)
            fig.show()
        case _:
            print("Gráfico não encontrado ou digitado incorretamente, tente um desses -->", *graficospossiveis, sep=" ")# o * trata cada argumento separadamente

def A_Automatizacao():#arquivo é o "DgraficoDados"
    
    def Janelacomscroll(titulo, dataframe):
        janela_dados = tk.Toplevel()
        janela_dados.title(titulo)
        janela_dados.geometry("600x400")
        # wrap=tk.NONE impede que as linhas quebrem sozinhas, mantendo a estrutura da tabela
        area_texto = scrolledtext.ScrolledText(janela_dados, undo=True, wrap=tk.NONE)
        area_texto.insert(tk.INSERT, dataframe)
        area_texto.configure(state='disabled')#só leitura
        area_texto.pack(expand=True, fill='both')
        janela_dados.wait_window()

    def Janeladebotoes(botao1, botao2, botao3=""):
        popup = tk.Toplevel(root)
        popup.title("Escolha um")
        popup.geometry("250x140+500+250")#Largura x Altura + X + Y
        popup.grab_set()
        escolha = tk.StringVar(value="Nenhuma")
        def clicar(valor):
            escolha.set(valor)
            popup.destroy()
        tk.Label(popup, text="Selecione o destino:", pady=10).pack()
        tk.Button(popup, text=botao1, command=lambda: clicar(botao1), width=25).pack(pady=2)
        tk.Button(popup, text=botao2, command=lambda: clicar(botao2), width=25).pack(pady=2)
        if(botao3 != ""):
            tk.Button(popup, text=botao3, command=lambda: clicar(botao3), width=25).pack(pady=2)
        #se for adicionar mais botões, ajuste o tamanho da janela no poput.geometry lá em cima
        root.wait_window(popup)
        return escolha.get()#retorna a opção escolhida

    def janelaDeCamposAInserir(titulo,lista_campos):
        respostas = {}#dicionario para as respostas que serão inseridas
        janela = tk.Toplevel()
        janela.title(titulo)

        largura_janela = 300
        altura_janela = len(lista_campos) * 40 + 80
        largura_tela = janela.winfo_screenwidth()
        altura_tela = janela.winfo_screenheight()
        pos_x = (largura_tela // 2) - (largura_janela // 2)
        pos_y = (altura_tela // 2) - (altura_janela // 2)
        janela.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")#toda essa parte centraliza

        janela.resizable(False, False)
        janela.grab_set()# Faz a janela ficar na frente e travar o código
        entradas = []
        for i, campo in enumerate(lista_campos): #i é o indice e campo será o nome em cada conjunto de lista_campos
            tk.Label(janela, text=campo).pack(pady=(5, 0))#pady é espaçamento
            ent = tk.Entry(janela, width=30)#caixa de texto criada
            ent.pack(pady=(0, 5))
            entradas.append((campo, ent))#salva os dados em uma lista
        def confirmar():
            for campo, ent in entradas:
                valor = ent.get()#transfere os dados para o dicionario
                try: #tenta converter
                    if "." in valor:
                        respostas[campo] = float(valor)
                    else:
                        respostas[campo] = int(valor)
                except ValueError:
                   respostas[campo] = valor
            janela.destroy()
        tk.Button(janela, text="Confirmar", command=confirmar, width=15).pack(pady=10)#command chama a função sem () para executar só ao clicar o botão
        janela.wait_window()# Espera a janela ser fechada para continuar a execução do script
        return respostas if respostas else None #volta um dicionario com as resp, acesse com respostas[nomedocampo]
    
    def anexar_documento_retornacaminho(qualtiporetornar):#janela de escolher arquivo
        if qualtiporetornar == "txt":
            tipos_arquivos = [
            ("Arquivos de Texto", "*.txt"),
            ("Arquivos CSV", "*.csv"),
            ("Todos os arquivos suportados", "*.txt *.csv")
            ]#aparece no canto inferior direito para selecionar
        elif qualtiporetornar == "excel":
            tipos_arquivos = [
            ("Planilhas Excel", "*.xlsx *.xls"),
            ("Todos os arquivos suportados", "*.xlsx *.xls")
            ]
        else:
            tipos_arquivos = [(("Todos os arquivos suportados", "*.*"))]
        caminho_arquivo = filedialog.askopenfilename(
            initialdir="/",
            title="Selecione um documento",
            filetypes=tipos_arquivos )
        if caminho_arquivo:
            print(f"O usuário escolheu o arquivo: {caminho_arquivo}")
            return caminho_arquivo
        else:
            print("Nenhum arquivo foi selecionado pelo usuário")
            return None

    def escolherdiretorio_retornacaminho():#janela de escolher local para salvar, falta fazer
        diretorio_escolhido = filedialog.askdirectory(
            initialdir="/",
            title="Selecione a pasta onde deseja salvar os arquivos"
        )
        if diretorio_escolhido:
            print(f"Diretório selecionado: {diretorio_escolhido}")
            return diretorio_escolhido
        else:
            messagebox.showerror("Erro!","Nenhum diretório selecionado!")
            return None

    def Janeladeopçoes(titulo, lista_opcoes):
        resultado = {"valor": None}#vai salvar aqui, apesar de retornar a resposta
        janela = tk.Toplevel()
        janela.title("Selecione uma Opção")

        largura_janela = 300
        altura_janela = 150
        largura_tela = janela.winfo_screenwidth()
        altura_tela = janela.winfo_screenheight()
        pos_x = (largura_tela // 2) - (largura_janela // 2)
        pos_y = (altura_tela // 2) - (altura_janela // 2)
        janela.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")
        
        janela.grab_set()
        label = tk.Label(janela, text=titulo, pady=10)
        label.pack()
        selecao = ttk.Combobox(janela, values=lista_opcoes, state="readonly")#mostra a lista de opcoes
        selecao.pack(pady=5)
        if lista_opcoes:
            selecao.current(0)# Define a primeira opção como padrão
        def confirmar():
            resultado["valor"] = selecao.get()#salva a resposta no resultado["valor"]
            janela.destroy()
        botao = tk.Button(janela, text="Confirmar", command=confirmar)
        botao.pack(pady=10)
        janela.wait_window()
        return resultado["valor"]

    def janelaFiltroDinamico(titulo, lista_colunas, campos_extras):#caixa de opção e entrada normal, feito exclusivamente para a opção 3
        #lista colunas recebe colunas de um arquivo
        resultado = {}
        janela = tk.Toplevel()
        janela.title(titulo)
        janela.grab_set()

        largura, altura = 350, 200 + (len(campos_extras) * 40)
        pos_x = (janela.winfo_screenwidth() // 2) - (largura // 2)
        pos_y = (janela.winfo_screenheight() // 2) - (altura // 2)
        janela.geometry(f"{largura}x{altura}+{pos_x}+{pos_y}")

        tk.Label(janela, text="Selecione a Coluna:", font=("Arial", 10, "bold")).pack(pady=5)
        selecao_coluna = ttk.Combobox(janela, values=lista_colunas, state="readonly", width=30)
        selecao_coluna.pack(pady=5)
        if lista_colunas: selecao_coluna.current(0)

        # Campos Adicionais (Valor a procurar, Valor novo, etc)- insira o valor na chamada como lista -> ["valores"]
        entradas = {}
        for campo in campos_extras:
            tk.Label(janela, text=f"{campo.replace('_', ' ')}:").pack(pady=2)
            en = tk.Entry(janela, width=33)
            en.pack(pady=2)
            entradas[campo] = en

        def confirmar():
            resultado["Nome_da_coluna"] = selecao_coluna.get()
            for campo, widget in entradas.items():
                resultado[campo] = widget.get()
            janela.destroy()

        tk.Button(janela, text="Confirmar", command=confirmar, bg="lightgray").pack(pady=15)
        janela.wait_window()
        return resultado

    def AcessarGraficoDados(qualindice):
        numresp = simpledialog.askinteger("Int. Ass.","O que você quer fazer?\n1)Importar dados para arquivo\n2)Ver dados do arquivo" \
        "\n3)Modificar ou filtrar dados\n4)Exportar dados do arquivo\n5)Fazer gráficos com os dados\n6)Escolher outro arquivo\n0)Sair do programa")
        if(numresp == 1):
            resp = Janeladebotoes("Arquivotxt/csv","Excel")
            if resp == "Arquivotxt/csv":
                caminho = anexar_documento_retornacaminho("txt")
                if(caminho != None):
                    graficosDadoscriadosAut[qualindice].arquivoGeral = D_arquivotxt_arquivografico_comPandas(caminho)
                    messagebox.showinfo("Sucesso!", "Operação realizada com sucesso!")
                    #funcionou, se não achar ele dá erro na própria função, se der ruim é só botar um try aqui
            elif resp == "Excel":
                caminho = anexar_documento_retornacaminho("excel")
                if(caminho != None):
                    graficosDadoscriadosAut[qualindice].arquivoGeral = D_arquivoExcel_arquivografico_comPandas(caminho)
                    messagebox.showinfo("Sucesso!", "Operação realizada com sucesso!")
            AcessarGraficoDados(qualindice)
            return
        elif(numresp == 2):
            try:
                conteudo = graficosDadoscriadosAut[qualindice].arquivoGeral.to_string()
                Janelacomscroll("Informações:", conteudo)
            except AttributeError:
                messagebox.showerror("Erro!","Arquivo sem informações!")
            AcessarGraficoDados(qualindice)
            return
        elif(numresp == 3):
            try:
                if graficosDadoscriadosAut[qualindice].arquivoGeral.to_string() != None:
                    resp = Janeladebotoes("Reduzir linhas com filtro","Apenas ver com filtro","Substituir valores")
                    if resp != "Substituir valores":
                        respostas = janelaFiltroDinamico("Insira", graficosDadoscriadosAut[qualindice].arquivoGeral.columns.to_list(), ["Valor_a_procurar"])
                        try:
                            conteudo = A_mostrarOuAlterarArquivoLinhasComFiltro_comPandas(graficosDadoscriadosAut[qualindice], respostas["Nome_da_coluna"] , respostas["Valor_a_procurar"],True,True).to_string()
                            Janelacomscroll("Informações:", conteudo)#aqui ele só mostra com filtro
                            if resp == "Reduzir linhas com filtro":#aqui ele pergunta se quer aplicar o filtro
                                simounao = messagebox.askyesno("Confirmação", "Deseja salvar as alterações no arquivo?")
                                if simounao: # True para 'Sim'
                                    graficosDadoscriadosAut[qualindice].arquivoGeral = A_mostrarOuAlterarArquivoLinhasComFiltro_comPandas(graficosDadoscriadosAut[qualindice], respostas["Nome_da_coluna"] , respostas["Valor_a_procurar"],False,True)
                        except AttributeError or conteudo == None:#nao sei se esse or realmente funciona
                            messagebox.showinfo("Atenção!","Sem registros com esse valor!")
                    else:
                        respostas = janelaFiltroDinamico("Insira", graficosDadoscriadosAut[qualindice].arquivoGeral.columns.to_list(), ["Valor_a_procurar", "Valor_novo"])
                        try:
                            conteudo = T_SubstituirDadosDeColuna_arquivograficoGeral_comPandas(graficosDadoscriadosAut[qualindice], respostas["Nome_da_coluna"] , respostas["Valor_a_procurar"],respostas["Valor_novo"]).to_string()
                            Janelacomscroll("Informações:", conteudo)
                            simounao = messagebox.askyesno("Confirmação", "Deseja salvar as alterações no arquivo?")
                            if simounao: # True para 'Sim'
                                graficosDadoscriadosAut[qualindice].arquivoGeral = T_SubstituirDadosDeColuna_arquivograficoGeral_comPandas(graficosDadoscriadosAut[qualindice], respostas["Nome_da_coluna"] , respostas["Valor_a_procurar"],respostas["Valor_novo"])
                        except AttributeError or conteudo == None:#nao sei se esse or realmente funciona
                            messagebox.showinfo("Atenção!","Sem registros com esse valor!")
            except AttributeError:
                messagebox.showerror("Erro!","Arquivo sem informações para serem usadas!")
            AcessarGraficoDados(qualindice)
            return
        elif(numresp == 4):
            try:
                if graficosDadoscriadosAut[qualindice].arquivoGeral.to_string() != None:
                    caminho = escolherdiretorio_retornacaminho()
                    if caminho != None:
                        resposta = Janeladebotoes("Arquivotxt/csv","Excel")
                        resp = simpledialog.askstring("Insira:", "Digite o nome a ser designado ao arquivo exportado:")
                        if resposta == "Excel":
                            if A_arquivografico_arquivoExcel_comPandas(graficosDadoscriadosAut[qualindice].arquivoGeral, resp, caminho):
                                messagebox.showinfo("Sucesso!", f"Arquivo {resp} em excel exportado no diretório {caminho} !")
                            else:
                                messagebox.showerror("Erro!","Não foi possível exportar!")
                        elif resposta == "Arquivotxt/csv":
                            if A_arquivografico_arquivotxt_comPandas(graficosDadoscriadosAut[qualindice].arquivoGeral, resp, caminho):
                                messagebox.showinfo("Sucesso!", f"Arquivo {resp} em txt exportado no diretório {caminho} !")
                            else:
                                messagebox.showerror("Erro!","Não foi possível exportar!")
            except AttributeError:
                messagebox.showerror("Erro!","Arquivo sem informações para serem usadas!")
            AcessarGraficoDados(qualindice)
            return
        elif(numresp == 5):
            try:
                respostaX = Janeladeopçoes("Qual coluna servirá para o eixo horizontal?",graficosDadoscriadosAut[qualindice].arquivoGeral.columns.tolist())
                colunas_restantes = [c for c in graficosDadoscriadosAut[qualindice].arquivoGeral.columns.tolist() if c != respostaX]
                respostaY = Janeladeopçoes("Qual coluna servirá para o eixo vertical?",colunas_restantes)
                respgrafico = Janeladeopçoes("Qual gráfico fazer?",graficospossiveis)
                A_verGraficodoGeral(graficosDadoscriadosAut[qualindice],respgrafico,respostaX,respostaY)
            except AttributeError:
                messagebox.showerror("Erro!","Arquivo sem informações para serem usadas!")
            AcessarGraficoDados(qualindice)
            return
        elif(numresp == 6):
            saberQualArquivo()

    def saberQualArquivo():
        numgrafico = simpledialog.askstring("Int. Ass.", "Qual o número do seu arquivo? (se for um novo digite novo)")
        if(numgrafico != "novo"):
            try:
                numgrafico = int(numgrafico)
                if(len(graficosDadoscriadosAut) >= numgrafico and numgrafico > 0):#vê se existe e tá nos limites/ numgrafico, oquefazer, dadosextras, **dicdadosextras
                    messagebox.showinfo("Arquivo encontrado!", f"Acessando arquivo {numgrafico}!")
                    AcessarGraficoDados(numgrafico - 1)
                else:
                    resposta = messagebox.askyesno("Aviso:", "Arquivo não encontrado, deseja criar um novo?")
                    if resposta:
                        numgrafico = str(numgrafico)
                        numgrafico = "novo"
                    else:
                        saberQualArquivo()
            except ValueError:
                messagebox.showerror("Erro!", "Valor inválido para encontrar o arquivo!")
        if(numgrafico == "novo"):#sem ser else mesmo
            grafico1 = DgraficoDados()
            graficosDadoscriadosAut.append(grafico1)
            messagebox.showinfo("Novo arquivo", f"Criando novo arquivo... seu número será {len(graficosDadoscriadosAut)}!")
            AcessarGraficoDados(len(graficosDadoscriadosAut) - 1)#sempre criado no final
    
    #INICIO--------------------------
    messagebox.showinfo("A_Automatizacao", "Acessando interface de assistência, guardamos seu arquivo por números!")
    saberQualArquivo()

if True:#códigos
    A_Automatizacao()

"""detalhes/anotações: 
os arquivos precisam ser uma coluna apenas com seus dados por linhas
os arquivos devem ser txt, csv ou excel 
Se for ler o arquivo com pandas, a primeira linha deve conter o cabeçario, mas isso pode ser mudado na linha 100
para ver gráficos, o eixo x deve ser o numérico
graficos txt serão descontinuados, é melhor usar o pandas 
geralmente as funções possuem o nome do seu tipo entre D T A, o que ela recebe _ o que ela retornará ou fará
A_arquivografico_arquivoTXT_comPandas ainda não fiz, mas deve ser só ler o arquivogeral com .to_string

combo de exemplo: 
    grafico1 = DgraficoDados()
    grafico1.arquivoGeral = D_arquivoExcel_arquivografico_comPandas("dadosexcel.xlsx") #você precisa criar um arquivo com esse nome!
    grafico1.arquivoGeral = T_SubstituirDadosDeColuna_arquivograficoGeral_comPandas(grafico1, 'Status' , "Pendente", "Entregue")
    A_arquivografico_arquivoExcel_comPandas(grafico1.arquivoGeral ,"graficoex1")

    grafico1 = DgraficoDados()
    grafico1.arquivoGeral = D_arquivoExcel_arquivografico_comPandas("dadosexcel.xlsx")
    A_verGraficodoGeral(grafico1,"pizza",'Total','Status')
    grafico1.arquivoGeral = T_SubstituirDadosDeColuna_arquivograficoGeral_comPandas(grafico1, 'Status' , "Pendente", "Entregue")
    A_verGraficodoGeral(grafico1,"pizza",'Total','Status')

    grafico1 = DgraficoDados()
    grafico1.arquivoGeral = D_arquivoExcel_arquivografico_comPandas("dadosexcel.xlsx")
    grafico1.arquivoGeral = T_SubstituirDadosDeColuna_arquivograficoGeral_comPandas(grafico1, 'Status' , "Pendente", "Entregue")
    grafico1.AVerArquivos()
    A_arquivoExcel_arquivotxt_comPandas("dadosexcel","excelinho")
    A_arquivotxt_arquivoExcel_comPandas("excelinho.txt")

feito por Daniel automatização feita, a janela principal não consegue minimiza-la

def A_arquivografico_arquivoExcel_comPandas(qualarquivo, nomeaserdado, nomecaminho):
    df = qualarquivo
    if df is not None: 
        nome_saida = nomeaserdado + '.xlsx' #ajuste de nome
        try:#ve se o arquivo já existe
            versejatem = pd.read_excel(f'{nomecaminho}/{nome_saida}')#vê se vai
            resp = input(f'Já existe um arquivo com esse nome, o existente será sobrescrito, tem certeza? (s ou n)').lower()
            if resp == 's':
                raise FileNotFoundError#finge que não existe, mas vai sobrescrever
        except FileNotFoundError: #se o arquivo resultado não existe, já faz direto
            caminho_saida = f'{nomecaminho}/{nome_saida}'
            try:
                df.to_excel(caminho_saida, index=False, engine='openpyxl')
                print(f"Arquivo exportado como: {caminho_saida}")
                return True
            except AttributeError:#transforma string em uma lista, se necessario
                lista_real = ast.literal_eval(qualarquivo)
                df = pd.DataFrame(lista_real, columns=['Valores'])
                df.to_excel(caminho_saida, index=False, engine='openpyxl')
                print(f"Arquivo exportado como: {caminho_saida}")
                return True #antigamente retornava df nos trues, vê se pode, mudei para a opção 4 na automação
    else:
        print(f"Erro: O arquivo do gráfico está vazio!")
        return None

  fique atento nos caminhos utilizados, caso você queira rodar o código!        
"""
