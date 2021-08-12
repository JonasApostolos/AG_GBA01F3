import win32com.client
import py_dss_interface
from win32com.client import makepy
from pylab import *
from operator import itemgetter
import matplotlib.pyplot as plt
import random
import os
import csv
import numpy
import statistics

class DSS(object):  # Classe DSS
    def __init__(self, dssFileName):
        self.dss = py_dss_interface.DSSDLL()
        self.dssFileName = dssFileName

    def compile_DSS(self):
        self.dss.dss_clearall()
        self.dss.text("compile [{}]".format(self.dssFileName))

        # OpenDSS folder
        self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)

    def solve(self, solucao, kWRatedList, barras, porcentagem_prosumidores, kwhrated=50000, kwhstored=30000):
        # self.compile_DSS()
        self.results_path = self.OpenDSS_folder_path + "/results_Main"
        self.dss.text("set DataPath=" + self.results_path)

        # # Monitores
        # for load in self.dss.loads_allnames():
        #     self.dss.text("New Monitor." + load + " Element=Load." + load + " mode=32 terminal=1")

        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]
        Loadshape = [LoadshapePointsList[ctd] for ctd in solucao[2:]]
        Loadshape = self.LoadshapeToMediaMovel(Loadshape)

        # kWhstored = 0.6*kwHRatedList[solucao[1]]
        # print(Loadshape)
        # print(str(barras[solucao[1]]))

        self.dss.text("Redirect PVSystems_" + str(porcentagem_prosumidores) + ".dss")

        self.dss.text("Loadshape.Loadshape1.mult=" + str(Loadshape))
        self.dss.text("Storage.storage.Bus1=" + str(barras[solucao[1]]))
        self.dss.text("Storage.storage.kWrated=" + str(kWRatedList[solucao[0]]))
        self.dss.text("Storage.storage.kva=" + str(kWRatedList[solucao[0]]))
        self.dss.text("Storage.storage.kw=" + str(kWRatedList[solucao[0]]))

        self.dss.text("Storage.storage.kWhrated=" + str(kwhrated))
        self.dss.text("Storage.storage.kWhstored=" + str(kwhstored))

        self.dss.text("Storage.storage.enabled=yes")

        self.dss.text("Solve")

        # self.dss.text("export eventlog")
        self.dss.text("export meters")
        for monitor in self.dss.monitors_allnames():
            self.dss.text("export monitor " + monitor)
        # self.dss.text("export monitor Potencia_Feeder")
        # self.dss.text("export monitor Storage")
        #
        # for load in self.dss.loads_allnames():
        #     self.dss.text("export monitor " + load)

    def funcaoCusto(self, solucao, kWRatedList, barras, porcentagem_prosumidores):
        self.compile_DSS()
        self.solve(solucao, kWRatedList, barras, porcentagem_prosumidores)

        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]
        Loadshape = [LoadshapePointsList[ctd] for ctd in solucao[2:]]
        Loadshape = self.LoadshapeToMediaMovel(Loadshape)

        # Punicao para maximar a amplitude da loadshape
        maximo = max([abs(min(Loadshape)), max(Loadshape)])
        if maximo >= 0.95:
            PunicaoMaxLoadshape = 0
        elif maximo >= 0.875 and maximo < 0.95:
            PunicaoMaxLoadshape = 5
        elif maximo >= 0.8 and maximo < 0.875:
            PunicaoMaxLoadshape = 10
        elif maximo < 0.8:
            PunicaoMaxLoadshape = 30

        # Inclinaçoes
        Inclinacao = 0
        ListaInclinacoes = self.InclinacoesLoadshape(solucao)

        for i in ListaInclinacoes:
            if numpy.abs(i) > 40:
                Inclinacao += numpy.abs(i)

        # Punição Niveis de Tensão
        if self.BarrasTensaoVioladas() > self.BarrasTensaoVioladasOriginal:
            PunicaoTensao = 9999999999
        else:
            PunicaoTensao = 0

        # PESOS
        a = 0.5  # Perdas
        b = 0.5  # DP do Carregamento do trafo

        # CICLO DE CARGA DA BATERIA
        # É preciso garantir que ao final das 48h o nível de carregamento da bateria seja o mesmo do inicio da simulacao
        Carregamento48h, PunicaoCicloCarga = self.PunicaoCiclodeCarga()

        # PERDAS
        ### Acessando arquivo CSV Potência
        dataEnergymeterCSV = {}
        self.dataperda = {}

        fname = "C:\\Users\\jonas\\PycharmProjects\\AG_GBA01F3\\ARB_GBA01F3_2019\\results_Main\\ARBGBA_EXP_METERS.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataEnergymeterCSV[row] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                    rowdata = row[ndata].replace(" ", "").replace('"',"")
                    if rowdata == "FEEDER" or rowdata == "":
                        dataEnergymeterCSV[name_col[ndata]].append(rowdata)
                    else:
                        dataEnergymeterCSV[name_col[ndata]].append(float(rowdata))

        self.dataperda['Perdas %'] = (dataEnergymeterCSV[' "Zone Losses kWh"'][0]/dataEnergymeterCSV[' "Zone kWh"'][0])*100
        os.remove(fname)

        # DESVIO PADRÃO DO CARREGAMENTO DO TRAFO
        ### Acessando arquivo CSV Potência
        dataFeederMmonitorCSV = {}

        fname = "C:\\Users\\jonas\\PycharmProjects\\AG_GBA01F3\\ARB_GBA01F3_2019\\results_Main\\ARBGBA_Mon_potencia_feeder_1.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataFeederMmonitorCSV[row] = []

            dataFeederMmonitorCSV['PTotal'] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                Pt = 0
                for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                    rowdata = row[ndata].replace(" ", "").replace('"',"")
                    dataFeederMmonitorCSV[name_col[ndata]].append(float(rowdata))
                    if name_col[ndata] == ' P1 (kW)' or name_col[ndata] == ' P2 (kW)' or name_col[ndata] == ' P3 (kW)':
                        Pt += float(rowdata)

                dataFeederMmonitorCSV['PTotal'].append(Pt)
        Desvio = statistics.pstdev(dataFeederMmonitorCSV['PTotal'])
        Perdas_sem_Pv_Stor = 2.316

        # Custo = a/(Perdas_sem_Pv_Stor/100-self.dataperda['Perdas %']/100) + Desvio + Inclinacao + PunicaoTensao + PunicaoCicloCarga
        Custo = self.dataperda['Perdas %'] + Desvio + Inclinacao + PunicaoTensao + PunicaoCicloCarga + PunicaoMaxLoadshape
        return Custo

    def mutacao(self, dominio, passo, solucao):
        i = random.randint(0, len(dominio) - 1)
        mutante = solucao

        if random.random() < 0.5:
            if solucao[i] != dominio[i][0] and solucao[i] >= (dominio[i][0] + passo):
                mutante = solucao[0:i] + [solucao[i] - passo] + solucao[i + 1:]
        else:
            if solucao[i] != dominio[i][1] and solucao[i] <= (dominio[i][1] - passo):
                mutante = solucao[0:i] + [solucao[i] + passo] + solucao[i + 1:]

        return mutante

    def cruzamento(self, dominio, individuo1, individuo2):
        i = random.randint(1, len(dominio) - 2)
        return individuo1[0:i] + individuo2[i:]

    def genetico(self,porcentagem_prosumidores, kWRatedList, barras, dominio, tamanho_populacao=80,  passo=1,
                 probabilidade_mutacao=0.2, elitismo=0.2):
        start2 = time.time()
        # self.Cenario(porcentagem_prosumidores) # cria o cenario

        self.BarrasTensaoVioladasOriginal = self.CalculaCustosOriginal(porcentagem_prosumidores)

        populacao = []

        # Cria a primeira geração
        for i in range(tamanho_populacao):
            # Solucao para todos os valores Random
            MaxLoadshapeFlag = 0 # Flag que marca se na loadshape tem algum ponto 1 ou -1
            while MaxLoadshapeFlag != 1:
                solucao = []
                for ctd in range(len(dominio)):
                    if ctd <= 2:
                        solucao.append(random.randint(dominio[ctd][0], dominio[ctd][1]))
                    else:
                        a = [dominio[ctd][0], solucao[-1] - 14]
                        a = max(a)
                        b = [dominio[ctd][1], solucao[-1] + 14]
                        b = min(b)
                        solucao.append(random.randint(a, b))
                # print(solucao)
                if min(solucao[2:]) <= 2 or max(solucao[2:]) >= 38:
                    MaxLoadshapeFlag = 1

            populacao.append(solucao)
        # print(populacao)
        numero_elitismo = int(elitismo * tamanho_populacao)
        geracao = 1
        stop = False
        melhor_solucao = [] # Lista de menor custo por geracao

        while stop == False:
            start = time.time()
            custos = [(self.funcaoCusto(individuo, kWRatedList, barras, porcentagem_prosumidores), individuo) for individuo in populacao]
            custos.sort()
            melhor_solucao.append(custos[0][0])
            if melhor_solucao.count(custos[0][0]) == int(0.2*tamanho_populacao): # Criterio de parada
                stop = True
            # custos_traduzidos = [(ctd[0], kWRatedList[ctd[1][0]], [LoadshapePointsList[i] for i in ctd[1][1:]]) for ctd in custos]
            # custos_traduzidos = [(ctd[0], kWRatedList[ctd[1][0]], kwHRatedList[ctd[1][1]]) for ctd in custos]
            custos_traduzidos = [(ctd[0], kWRatedList[ctd[1][0]], barras[ctd[1][1]]) for ctd in custos]
            print("Geração::", geracao,  custos_traduzidos)
            # self.CalculaCustos(custos[0][1], kWRatedList, barras, porcentagem_prosumidores)
            print("Melhores Resultados", melhor_solucao)
            geracao += 1
            individuos_ordenados = [individuo for (custo, individuo) in custos]
            populacao = individuos_ordenados[0:numero_elitismo]
            lista_rank = [(individuo, (tamanho_populacao - individuos_ordenados.index(individuo))/(tamanho_populacao*(tamanho_populacao-1))) for individuo in individuos_ordenados]
            lista_rank.reverse()
            soma=0
            for ctd in lista_rank:
                soma += ctd[1]

            # Cruzamento e Mutacao dos individuos
            while len(populacao) < tamanho_populacao:
                if random.random() < probabilidade_mutacao:
                    m = random.randint(0, numero_elitismo)
                    populacao.append(self.mutacao(dominio, passo, individuos_ordenados[m]))
                else:
                    aleatorio = random.uniform(0, soma)
                    # print('aleatorio', aleatorio)
                    s = 0
                    for j in lista_rank:
                        s += j[1]
                        if aleatorio < s:
                            c1 = j[0]
                            # print('c1', c1)
                            break
                    aleatorio = random.uniform(0, soma)
                    s = 0
                    for j in lista_rank:
                        s += j[1]
                        if aleatorio < s:
                            c2 = j[0]
                            # print('c2', c2)
                            break
                    populacao.append(self.cruzamento(dominio, c1, c2))
                    # c1 = random.randint(0, numero_elitismo)
                    # c2 = random.randint(0, numero_elitismo)
                    # populacao.append(self.cruzamento(dominio, individuos_ordenados[c1], individuos_ordenados[c2]))

            end = time.time()
            print("Tempo da geração:", end - start)

        Loadshape, Perda, Carregamento, Inclinacao, Tensao, Desvio, kWhRated, Demanda, kWhstored = self.CalculaCustos(custos[0][1], kWRatedList, barras, porcentagem_prosumidores)

        end2 = time.time()

        results_file = open("Resultados.txt", "a")
        results_file.write(f"{geracao}, {(end2 - start2)/3600}, {custos_traduzidos[0][0]}, {custos_traduzidos[0][1]}, {custos_traduzidos[0][2]}, 50000, {Loadshape}, {Perda}, {Carregamento}, {Inclinacao}, {Tensao}, {Desvio}, {kWhRated}, {kWhstored} \n{Demanda} \n{melhor_solucao} \n")
        results_file.close()

        print("tempo total:", end2 - start2)
        return custos[0][1]

    def LoadshapeToMediaMovel(self, loadshape):
        medias_moveis = []
        num_media = 2
        i = 0
        while i < (len(loadshape) - num_media + 1):
            grupo = loadshape[i: i + num_media]
            media_grupo = round(sum(grupo) / num_media, 3)
            medias_moveis.append(media_grupo)
            i += 1
        # medias_moveis.insert(0, medias_moveis[0])
        return medias_moveis

    def InclinacoesLoadshape(self, solucao):
        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]
        Loadshape = [LoadshapePointsList[i] for i in solucao[2:]]
        Inclinacoes = []

        for i in range((len(Loadshape))):
            if i == 24:
                x = Loadshape[0] - Loadshape[24]
            else:
                x = Loadshape[i+1] - Loadshape[i]
            Inclinacoes.append(numpy.arctan(x)*180/pi)
        return Inclinacoes

    def CalculaCustosOriginal(self, porcentagem_prosumidores):
        self.compile_DSS()

        self.results_path = self.OpenDSS_folder_path + "/results_Main"
        self.dss.text("set DataPath=" + self.results_path)

        self.dss.text("Storage.storage.enabled=no")
        self.dss.text("Redirect PVSystems_" + str(porcentagem_prosumidores) + ".dss")

        self.dss.text("Solve")
        # self.dss.text("Plot monitor object= potencia_feeder channels=(1 3 5 )")
        print("Solved")

        for monitor in self.dss.monitors_allnames():
            self.dss.text("export monitor " + monitor)
        self.dss.text("export meters")
        # self.dss.text("export monitor Potencia_Feeder")
        # for load in self.dss.loads_allnames():
        #     self.dss.text("export monitor " + load)

        ### Acessando arquivo CSV Potência
        dataEnergymeterCSV = {}
        self.dataperda = {}

        fname = "C:\\Users\\jonas\\PycharmProjects\\AG_GBA01F3\\ARB_GBA01F3_2019\\results_Main\\ARBGBA_EXP_METERS.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataEnergymeterCSV[row] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                    rowdata = row[ndata].replace(" ", "").replace('"',"")
                    if rowdata == "FEEDER" or rowdata == "":
                        dataEnergymeterCSV[name_col[ndata]].append(rowdata)
                    else:
                        dataEnergymeterCSV[name_col[ndata]].append(float(rowdata))

        self.dataperda['Perdas %'] = (dataEnergymeterCSV[' "Zone Losses kWh"'][0]/dataEnergymeterCSV[' "Zone kWh"'][0])*100
        os.remove(fname)

        ### Acessando arquivo CSV Potência
        dataFeederMmonitorCSV = {}

        fname = "C:\\Users\\jonas\\PycharmProjects\\AG_GBA01F3\\ARB_GBA01F3_2019\\results_Main\\ARBGBA_Mon_potencia_feeder_1.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataFeederMmonitorCSV[row] = []

            dataFeederMmonitorCSV['PTotal'] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                Pt = 0
                for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                    rowdata = row[ndata].replace(" ", "").replace('"', "")
                    dataFeederMmonitorCSV[name_col[ndata]].append(float(rowdata))
                    if name_col[ndata] == ' P1 (kW)' or name_col[ndata] == ' P2 (kW)' or name_col[ndata] == ' P3 (kW)':
                        Pt += float(rowdata)

                dataFeederMmonitorCSV['PTotal'].append(Pt)

        barrasVioladas = self.BarrasTensaoVioladas()
        print('Custos Sistema Original (Somente GD-PV)')
        print('Perdas:', self.dataperda['Perdas %'], 'Violações de Tensao:', barrasVioladas, 'PTotal:', dataFeederMmonitorCSV['PTotal'], '\n')
        return barrasVioladas

    def CalculaCustos(self, solucao, kWRatedList, barras, porcentagem_prosumidores):
        self.compile_DSS()
        self.solve(solucao, kWRatedList, barras, porcentagem_prosumidores)

        ### Acessando CSV Storage
        dataStorageMmonitorCSV = {}

        fname = "C:\\Users\\jonas\\PycharmProjects\\AG_GBA01F3\\ARB_GBA01F3_2019\\results_Main\\ARBGBA_Mon_storage_1.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataStorageMmonitorCSV[row] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                for ndata in range(0, len(name_col)-1):  ## Varendo todas as colunas
                    if row != ['ÿÿÿÿ']:
                        rowdata = row[ndata].replace(" ", "").replace('"', "")
                        dataStorageMmonitorCSV[name_col[ndata]].append(float(rowdata))

        maxkWh = max(dataStorageMmonitorCSV[' kWh'])
        minkWh = min(dataStorageMmonitorCSV[' kWh'])
        kWhRated = (maxkWh-minkWh)/0.7
        kwhstored = 30000-minkWh+0.25*kWhRated

        self.compile_DSS()
        self.solve(solucao, kWRatedList, barras, porcentagem_prosumidores, kWhRated, kwhstored)

        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]
        Loadshape = [LoadshapePointsList[ctd] for ctd in solucao[2:]]
        Loadshape = self.LoadshapeToMediaMovel(Loadshape)

        # Punicao para maximar a amplitude da loadshape
        maximo = max([abs(min(Loadshape)), max(Loadshape)])
        if maximo >= 0.95:
            PunicaoMaxLoadshape = 0
        elif maximo >= 0.875 and maximo < 0.95:
            PunicaoMaxLoadshape = 5
        elif maximo >= 0.8 and maximo < 0.875:
            PunicaoMaxLoadshape = 10
        elif maximo < 0.8:
            PunicaoMaxLoadshape = 30

        # CICLO DE CARGA DA BATERIA
        # É preciso garantir que ao final das 48h o nível de carregamento da bateria seja o mesmo do inicio da simulacao
        Carregamento48h, PunicaoCicloCarga = self.PunicaoCiclodeCarga()

        # Inclinaçoes
        Inclinacao = 0
        ListaInclinacoes = self.InclinacoesLoadshape(solucao)

        for i in ListaInclinacoes:
            if numpy.abs(i) > 40:
                Inclinacao += numpy.abs(i)

        ### Acessando arquivo CSV Perdas
        dataEnergymeterCSV = {}
        self.dataperda = {}

        fname = "C:\\Users\\jonas\\PycharmProjects\\AG_GBA01F3\\ARB_GBA01F3_2019\\results_Main\\ARBGBA_EXP_METERS.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataEnergymeterCSV[row] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                    rowdata = row[ndata].replace(" ", "").replace('"',"")
                    if rowdata == "FEEDER" or rowdata == "":
                        dataEnergymeterCSV[name_col[ndata]].append(rowdata)
                    else:
                        dataEnergymeterCSV[name_col[ndata]].append(float(rowdata))

        self.dataperda['Perdas %'] = (dataEnergymeterCSV[' "Zone Losses kWh"'][0]/dataEnergymeterCSV[' "Zone kWh"'][0])*100
        os.remove(fname)

        ### Acessando arquivo CSV Potência
        dataFeederMmonitorCSV = {}

        fname = "C:\\Users\\jonas\\PycharmProjects\\AG_GBA01F3\\ARB_GBA01F3_2019\\results_Main\\ARBGBA_Mon_potencia_feeder_1.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)

            for row in name_col:
                dataFeederMmonitorCSV[row] = []

            dataFeederMmonitorCSV['PTotal'] = []

            for row in csv_reader_object:  ##Varendo todas as linhas
                Pt = 0
                for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                    rowdata = row[ndata].replace(" ", "").replace('"', "")
                    dataFeederMmonitorCSV[name_col[ndata]].append(float(rowdata))
                    if name_col[ndata] == ' P1 (kW)' or name_col[ndata] == ' P2 (kW)' or name_col[ndata] == ' P3 (kW)':
                        Pt += float(rowdata)

                dataFeederMmonitorCSV['PTotal'].append(Pt)

            Desvio = statistics.pstdev(dataFeederMmonitorCSV['PTotal'])

        # print('Perdas:', self.dataperda['Perdas %'], 'kWh 48h:', Carregamento48h, 'Inclinação:', Inclinacao, 'Barras_Violada:', self.BarrasTensaoVioladas(), 'Desvio:', Desvio, 'PTotal:', dataFeederMmonitorCSV['PTotal'])
        # print(self.LoadshapeToMediaMovel(Loadshape),",", self.dataperda['Perdas %'],",", Carregamento48h,",", Inclinacao,",", self.BarrasTensaoVioladas(),",", Desvio, ",", kWhRated)
        # print(dataFeederMmonitorCSV['PTotal'])
        # print('Loadshape:', self.LoadshapeToMediaMovel(Loadshape))
        return str(Loadshape).replace("[", "").replace("]", ""), self.dataperda['Perdas %'], Carregamento48h, Inclinacao, self.BarrasTensaoVioladas(), Desvio, kWhRated, dataFeederMmonitorCSV['PTotal'], kwhstored

    def BarrasTensaoVioladas(self):
        BarrasVioladas = 0

        for trafo in self.dss.transformers_allNames():
            self.dss.transformers_write_name(trafo)
            if self.dss.transformers_read_kv() != 13.8:
                dataMonitorCargas = {}
                fname = "C:\\Users\\jonas\\PycharmProjects\\AG_GBA01F3\\ARB_GBA01F3_2019\\results_Main\\ARBGBA_Mon_" + trafo + "_1.csv"

                with open(str(fname), 'r', newline='') as file:
                    csv_reader_object = csv.reader(file)
                    name_col = next(csv_reader_object)

                    for row in name_col:
                        dataMonitorCargas[row] = []

                    for row in csv_reader_object:  ##Varendo todas as linhas
                        for ndata in range(0, len(name_col)):  ## Varendo todas as colunas
                            rowdata = row[ndata].replace(" ", "").replace('"',"")
                            if name_col[ndata] == ' |V|1 (volts)' or name_col[ndata] == ' |V|2 (volts)' or name_col[ndata] == ' |V|3 (volts)':
                                dataMonitorCargas[name_col[ndata]].append(float(rowdata)/127)

                TensaoPUFasesBarras = dataMonitorCargas[' |V|1 (volts)'] + dataMonitorCargas[' |V|2 (volts)']
                # print(TensaoPUFasesBarras)
                for ctd in TensaoPUFasesBarras:
                    if ctd > 1.03 or ctd < 0.95:
                        BarrasVioladas += 1

        # TensaoPUFasesBarras = d.dssCircuit.AllNodeVmagPUByPhase(1) + d.dssCircuit.AllNodeVmagPUByPhase(2) + d.dssCircuit.AllNodeVmagPUByPhase(3)
        # for i in TensaoPUFasesBarras:
        #     if i > 1.03 or i < 0.97:
        #         BarrasVioladas += 1
        return BarrasVioladas

    def PunicaoCiclodeCarga(self, kwhstored=30000):

        dataMonitorStorage = {}

        fname = "C:\\Users\\jonas\\PycharmProjects\\AG_GBA01F3\\ARB_GBA01F3_2019\\results_Main\\ARBGBA_Mon_storage_1.csv"

        with open(str(fname), 'r', newline='') as file:
            csv_reader_object = csv.reader(file)
            name_col = next(csv_reader_object)
            for row in name_col:
                dataMonitorStorage[row] = []
            for row in csv_reader_object:  ##Varendo todas as linhas
                for ndata in range(0, len(name_col)-3):  ## Varendo todas as colunas
                    if row != ['ÿÿÿÿ']:
                        rowdata = row[ndata].replace(" ", "").replace('"',"")
                        dataMonitorStorage[name_col[ndata]].append(float(rowdata))

        Carregamento48h = dataMonitorStorage[' kWh'][-1]
        PunicaoCicloCarga = abs((kwhstored-Carregamento48h)/100)
        return Carregamento48h, PunicaoCicloCarga

if __name__ == '__main__':
    d = DSS(r"C:\Users\jonas\PycharmProjects\AG_GBA01F3\ARB_GBA01F3_2019\GBA01F3_2019.dss")

    # LISTAS DO AG
    d.compile_DSS()
    barras = []
    for bus in list(d.dss.circuit_allbusnames()):
        d.dss.circuit_setactivebus(bus)
        if len(d.dss.bus_nodes()) == 3 and ('gba' not in bus):
            barras.append(d.dss.bus_name())
    print(barras)

    # results_file = open("Monitores.txt", "a")
    # Monitores = []
    # for trafo in d.dss.transformers_allNames():
    #     d.dss.transformers_write_name(trafo)
    #     if d.dss.transformers_read_kv() != 13.8:
    #         Monitores.append("New Monitor." + trafo + " Element=Transformer." + trafo + " mode=32 terminal=2\n")
    #
    # results_file.writelines(Monitores)
    # results_file.close()

    kWRatedList = list(range(100, 5000, 200))
    kwHRatedList = list(range(1000, 35000, 500))
    dominio = [(0, len(kWRatedList) - 1), (0, len(barras) - 1), (0, 40) , (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40), (0, 40)]
    # dominio = [(0, len(kWRatedList) - 1), (0, 40) , (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40)]
    # dominio = [(0, len(kWRatedList) - 1), (0, len(barras) - 1), (0, 40) , (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40)]
    # dominio = [(0, len(kWRatedList) - 1), (0, len(kwHRatedList) - 1), (0, 40), (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40)]
    # dominio = [(0, 40) , (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40)]

    # ALGORITMO GENÉTICO
    porcentagem_prosumidores = 0.1
    for i in list(range(1,2,1)):
        d.genetico(porcentagem_prosumidores, kWRatedList, barras, dominio)



