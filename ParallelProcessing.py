import time
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

    def funcaoCusto(self, populacao, kWRatedList, barras, pv_percentage, kwhrated=50000, kwhstored=30000):
        cpu1 = populacao[0:len(populacao)//3]
        cpu2 = populacao[len(populacao)//3:2*len(populacao)//3]
        cpu3 = populacao[2*len(populacao)//3:len(populacao)]

        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]

        for trafo in self.dss.transformers_allNames():
            self.dss.transformers_write_name(trafo)
            if self.dss.transformers_read_kv() != 13.8:
                self.dss.text("New Monitor." + trafo + " Element=Transformer." + trafo + " mode=32 terminal=2")

        for ctd in range(0, len(cpu1)):
            print(ctd)
            solucao1 = cpu1[ctd]
            solucao2 = cpu2[ctd]
            solucao3 = cpu3[ctd]

            self.dss.dss_clearall()
            self.dss.text("ClearAll")
            self.dss.parallel_write_activeparallel(0)
            self.dss.text("compile [{}]".format(self.dssFileName))
            self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)
            self.dss.parallel_write_actorcpu(0)
            self.dss.solution_solve()

            self.dss.text("clone 2")

            ################## CPU 0
            self.dss.parallel_write_activeactor(1)
            # print("active actor", self.dss.parallel_read_activeactor())
            self.dss.parallel_write_actorcpu(0)
            self.dss.text("set mode=Daily stepsize=15m number=97")
            self.dss.text("Redirect PVSystems_" + str(pv_percentage) + ".dss")

            Loadshape = [LoadshapePointsList[ctd] for ctd in solucao1[2:]]
            Loadshape = self.LoadshapeToMediaMovel(Loadshape)
            self.dss.text("Loadshape.Loadshape1.mult=" + str(Loadshape))
            self.dss.text("Storage.storage.Bus1=" + str(barras[solucao1[1]]))
            self.dss.text("Storage.storage.kWrated=" + str(kWRatedList[solucao1[0]]))
            self.dss.text("Storage.storage.kva=" + str(kWRatedList[solucao1[0]]))
            self.dss.text("Storage.storage.kw=" + str(kWRatedList[solucao1[0]]))
            self.dss.text("Storage.storage.kWhrated=" + str(kwhrated))
            self.dss.text("Storage.storage.kWhstored=" + str(kwhstored))
            self.dss.text("Storage.storage.enabled=yes")

            ################## CPU 2
            self.dss.parallel_write_activeactor(2)
            # print("active actor", self.dss.parallel_read_activeactor())
            self.dss.parallel_write_actorcpu(2)
            self.dss.text("set mode=Daily stepsize=15m number=97")
            self.dss.text("Redirect PVSystems_" + str(pv_percentage) + ".dss")

            Loadshape = [LoadshapePointsList[ctd] for ctd in solucao2[2:]]
            Loadshape = self.LoadshapeToMediaMovel(Loadshape)
            self.dss.text("Loadshape.Loadshape1.mult=" + str(Loadshape))
            self.dss.text("Storage.storage.Bus1=" + str(barras[solucao2[1]]))
            self.dss.text("Storage.storage.kWrated=" + str(kWRatedList[solucao1[0]]))
            self.dss.text("Storage.storage.kva=" + str(kWRatedList[solucao2[0]]))
            self.dss.text("Storage.storage.kw=" + str(kWRatedList[solucao2[0]]))
            self.dss.text("Storage.storage.kWhrated=" + str(kwhrated))
            self.dss.text("Storage.storage.kWhstored=" + str(kwhstored))
            self.dss.text("Storage.storage.enabled=yes")

            ################## CPU 3
            self.dss.parallel_write_activeactor(3)
            # print("active actor", self.dss.parallel_read_activeactor())
            self.dss.parallel_write_actorcpu(3)
            self.dss.text("set mode=Daily stepsize=15m number=97")
            self.dss.text("Redirect PVSystems_" + str(pv_percentage) + ".dss")

            Loadshape = [LoadshapePointsList[ctd] for ctd in solucao3[2:]]
            # Loadshape = self.LoadshapeToMediaMovel(Loadshape)
            self.dss.text("Loadshape.Loadshape1.mult=" + str(Loadshape))
            self.dss.text("Storage.storage.Bus1=" + str(barras[solucao3[1]]))
            self.dss.text("Storage.storage.kWrated=" + str(kWRatedList[solucao3[0]]))
            self.dss.text("Storage.storage.kva=" + str(kWRatedList[solucao3[0]]))
            self.dss.text("Storage.storage.kw=" + str(kWRatedList[solucao3[0]]))
            self.dss.text("Storage.storage.kWhrated=" + str(kwhrated))
            self.dss.text("Storage.storage.kWhstored=" + str(kwhstored))
            self.dss.text("Storage.storage.enabled=yes")

            self.dss.parallel_write_activeactor(1)
            self.dss.parallel_write_activeparallel(1)
            self.dss.solution_solveall()

            boolstatus = 0
            while boolstatus == 0:
                if self.dss.parallel_actorstatus() == (1, 1, 1):
                    boolstatus = 1

            for i in range(1, self.dss.parallel_numactors() + 1):
                self.dss.parallel_write_activeactor(i)
                self.dss.text("export meters")
                for monitor in self.dss.monitors_allnames():
                    self.dss.text("export monitor " + monitor)





    def genetico(self,pv_percentage, kWRatedList, barras, dominio, tamanho_populacao=81,  passo=1,
                 probabilidade_mutacao=0.2, elitismo=0.2):
        start2 = time.time()

        self.BarrasTensaoVioladasOriginal = self.CalculaCustosOriginal(pv_percentage)

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

        numero_elitismo = int(elitismo * tamanho_populacao)
        geracao = 1
        stop = False
        melhor_solucao = [] # Lista de menor custo por geracao

        while stop == False:
            start = time.time()
            custos = self.funcaoCusto(populacao, kWRatedList, barras, pv_percentage)

    def CalculaCustosOriginal(self, porcentagem_prosumidores):
        self.dss.dss_clearall()
        self.dss.text("compile [{}]".format(self.dssFileName))

        # OpenDSS folder
        self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)

        self.results_path = self.OpenDSS_folder_path + "/results_Main"
        self.dss.text("set DataPath=" + self.results_path)

        # Monitores
        for trafo in self.dss.transformers_allNames():
            self.dss.transformers_write_name(trafo)
            if self.dss.transformers_read_kv() != 13.8:
                self.dss.text("New Monitor." + trafo + " Element=Transformer." + trafo + " mode=32 terminal=2")

        self.dss.text("Storage.storage.enabled=no")
        self.dss.text("Redirect PVSystems_" + str(porcentagem_prosumidores) + ".dss")

        self.dss.parallel_write_activeparallel(0)
        self.dss.text("Solve")
        # self.dss.text("Plot monitor object= potencia_feeder channels=(1 3 5 )")

        self.dss.text("export meters")
        for monitor in self.dss.monitors_allnames():
            self.dss.text("export monitor " + monitor)

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

    def BarrasTensaoVioladas(self):
        BarrasVioladas = 0

        for load in self.dss.loads_allnames():
            dataMonitorCargas = {}
            fname = "C:\\Users\\jonas\\PycharmProjects\\AG_GBA01F3\\ARB_GBA01F3_2019\\results_Main\\ARBGBA_Mon_" + load + "_1.csv"

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


if __name__ == '__main__':
    d = DSS(r"C:\Users\jonas\PycharmProjects\AG_GBA01F3\ARB_GBA01F3_2019\GBA01F3_2019.dss")

    d.dss.dss_clearall()
    d.dss.text("compile [{}]".format(d.dssFileName))

    barras = []
    for bus in list(d.dss.circuit_allbusnames()):
        d.dss.circuit_setactivebus(bus)
        if len(d.dss.bus_nodes()) == 3 and ('gba' not in bus):
            barras.append(d.dss.bus_name())
    # print(barras)

    pv_percentage = 0.2
    kWRatedList = list(range(100, 5000, 200))
    kwHRatedList = list(range(1000, 35000, 500))
    dominio = [(0, len(kWRatedList) - 1), (0, len(barras) - 1), (0, 40) , (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40), (0, 40)]

    d.genetico(pv_percentage, kWRatedList, barras, dominio)
