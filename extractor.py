import tabula
import os
import pandas as pd


class linha:
    def __init__(self, numero, plano, valor):
        self.numero = numero
        self.plano = plano
        self.valor = valor


def getPath(pasta):
    return [os.path.join(pasta, nome) for nome in os.listdir(pasta)]


def read(pastas, pags):
    lista = []
    for pasta in pastas:
        df = tabula.read_pdf(pasta, pages=pags, multiple_tables=True)
        if len(pags) != 2:
            for i in range(0, len(df)):
                try:
                    df[i] = df[i].iloc[1:, ]

                    print(df[i])
                    if i % 2 == 0:
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
                    else:
                        df[i].columns = [0, 1, 2]
                        for a in df[i].index:
                            temp = df[i][0][a]
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
                except:
                    print('errou')
        else:
            for i in range(0, len(df)):
                try:
                    df[i] = df[i].iloc[2:, ]
                    print(df[i])
                    if i % 2 != 0:
                        for a in df[i].index:
                            print(type(df[i][1][a]))
                            if type(df[i][1][a]) is not str and type(df[i][1][a]) is not float:
                                temp = df[i][0][a]
                                index0 = temp.rfind(' 54-')
                                index00 = temp.find(' ')
                                numero = temp[:index00].strip().replace(
                                    '-', '')
                                temp2 = temp[index00:index0].strip()
                                indexlast = temp2.rfind(' ')
                                plano = temp2[:indexlast].strip()
                                valor = temp2[indexlast:].strip()
                                dic1 = {'numero': numero,
                                        'plano': plano, 'valor': valor}
                                temp3 = temp[index0:].strip()
                                indexfirst = temp.find(' ')
                                lista.append(dic1)
                                numero = temp3[:indexfirst].strip().replace(
                                    '-', '')
                                plano = temp3[indexfirst:].strip()
                                valor = df[i][2][a].strip()
                                dic1 = {'numero': numero,
                                        'plano': plano, 'valor': valor}
                                lista.append(dic1)
                            else:
                                temp = df[i][0][a]
                                index0 = temp.rfind(' ')
                                index00 = temp.find(' ')
                                numero = temp[:index00].strip().replace(
                                    '-', '')
                                plano = temp[index00:index0].strip()
                                valor = temp[index0:].strip()
                                dic1 = {'numero': numero,
                                        'plano': plano, 'valor': valor}
                                lista.append(dic1)
                                if type(df[i][1][a]) is str:
                                    numero = df[i][1][a].strip().replace(
                                        '-', '')
                                    plano = df[i][2][a].strip()
                                    valor = df[i][3][a].strip()
                                    dic1 = {'numero': numero,
                                            'plano': plano, 'valor': valor}
                                    lista.append(dic1)
                    else:
                        for a in df[i].index:
                            temp = df[i][0][a]
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
                            plano = df[i][1][a].strip()
                            valor = df[i][2][a].strip()
                            dic1 = {'numero': numero2,
                                    'plano': plano, 'valor': valor}
                            lista.append(dic1)
                except:
                    print('errou')
    return lista


def write(lista, name):
    df = pd.DataFrame(lista)
    writer = pd.ExcelWriter(name+'.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1')
    writer.save()


caminhos = getPath('PDF')
paginas = [4]
lista = read(caminhos, paginas)
write(lista, 'gracioli111092020')
