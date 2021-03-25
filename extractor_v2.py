import tabula
import os
import pandas as pd

def getPath(pasta):
    lista = []
    for nome in os.listdir(pasta):
        lista.append({"caminho": os.path.join(pasta, nome), "nome":nome.lower().replace('.pdf', ".xlsx")})
    return lista

def processDataFrame(df):
    lista = []
    for i in range(0, len(df)):
        try:
            df[i] = df[i].iloc[1:, ]
            df[i] = df[i].dropna(how='all', axis=1)
            print(df[i])
            if len(df[i].columns) == 4:
                df[i].columns = [0, 1, 2, 3]
                for a in df[i].index:
                    temp = df[i][0][a]
                    index0 = temp.rfind(' ')
                    index00 = temp.find(' ')
                    numero = temp[:index00].strip().replace('-', '')
                    plano = temp[index00:index0].strip()
                    valor = temp[index0:].strip()
                    dic1 = {'numero': numero,
                            'plano': plano, 'valor': valor}
                    lista.append(dic1)
                    if type(df[i][1][a]) is str:
                        numero = df[i][1][a].strip().replace('-', '')
                        plano = df[i][2][a].strip()
                        valor = df[i][3][a].strip()
                        dic1 = {'numero': numero,
                                'plano': plano, 'valor': valor}
                        lista.append(dic1)
            elif len(df[i].columns) == 3:
                df[i].columns = [0, 1, 2]
                for a in df[i].index:
                    temp = df[i][0][a]
                    if 'Número' not in temp:
                        index0 = temp.rfind(' ')
                        index00 = temp.find(' ')
                        numero = temp[:index00].strip().replace('-', '')
                        index1 = temp[:index0].rfind(' ')
                        plano = temp[index00:index1].strip()
                        valor = temp[index1:index0].strip()
                        numero2 = temp[index0:].strip().replace('-', '')
                        dic1 = {'numero': numero,
                                'plano': plano, 'valor': valor}
                        lista.append(dic1)
                        plano = df[i][1][a]
                        valor = df[i][2][a]
                        dic1 = {'numero': numero2,
                                'plano': plano, 'valor': valor}
                        lista.append(dic1)
            else:
                df[i].columns = [0, 1]
                for a in df[i].index:
                    temp = df[i][0][a]
                    if 'Número' not in temp:
                        listatemp = temp.split(' ')
                        numero = listatemp[0].strip().replace('-', '')
                        index = 1
                        for b in range(1, len(listatemp)): 
                            if ',' in listatemp[b]: 
                                index = b
                                break
                        plano = ' '.join(listatemp[1:index])
                        valor = listatemp[index]
                        dic1 = {'numero': numero,
                                'plano': plano, 'valor': valor}
                        lista.append(dic1)
                        if index+1 < len(listatemp):
                            numero2 = listatemp[index+1].strip().replace('-', '')
                            plano2 = ' '.join(listatemp[index+2:])
                            valor2 = df[i][1][a]
                            dic1 = {'numero': numero2,
                                    'plano': plano2, 'valor': valor2}
                            lista.append(dic1)
        except Exception as e:
            print(e)
            print('errou')
    return lista

def write(lista, name):
    df = pd.DataFrame(lista)
    writer = pd.ExcelWriter('XLSX/'+name, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    writer.save()

def main():
    folder = 'PDF'
    paths = getPath(folder)
    for path in paths:
        tables = tabula.read_pdf(path['caminho'], pages = [2,3,4,5], multiple_tables = True)
        df = []
        for x in tables:
            if 'VEJA OS NÚMEROS VIVO E PLANOS QUE COMPÕEM A SUA CONTA' in x:
                df.append(x)
            elif 'VEJA OS NÚMEROS VIVO E PLANOS QUE COMPÕEM A SUA CONTA - Continuação' in x:
                df.append(x)
            elif 'DETALHAMENTO TOTAL DA CONTA' in x:
                if 'VEJA OS NÚMEROS VIVO E PLANOS QUE COMPÕEM A SUA CONTA - Continuação' == x['DETALHAMENTO TOTAL DA CONTA'][0]:
                    df.append(x)
                elif 'VEJA OS NÚMEROS VIVO E PLANOS QUE COMPÕEM A SUA CONTA' == x['DETALHAMENTO TOTAL DA CONTA'][0]:
                    df.append(x)
        lista = processDataFrame(df)
        write(lista, path['nome'])
main()