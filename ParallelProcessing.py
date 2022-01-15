import time
import py_dss_interface
from pylab import *
import random
import os
import numpy
import statistics

class DSS(object):  # Classe DSS
    def __init__(self, dssFileName):
        self.dss = py_dss_interface.DSSDLL()
        self.dssFileName = dssFileName

    def funcaoCusto(self, populacao, kWRatedList, barras, pv_percentage, kwhrated=50000, kwhstored=30000):
        start_funcaocusto = time.time()
        custos = []

        cpu1 = populacao[0:len(populacao)//3]
        cpu2 = populacao[len(populacao)//3:2*len(populacao)//3]
        cpu3 = populacao[2*len(populacao)//3:len(populacao)]
        # print("cpu1", cpu1)

        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]

        for ctd in range(0, len(cpu1)):
            # print("Solucao", ctd)


            solucoesCPU = [cpu1[ctd], cpu2[ctd], cpu3[ctd]]

            self.dss.text("ClearAll")
            # self.dss.text("Set Parallel=No")
            # self.dss.text("compile [{}]".format(self.dssFileName))
            # self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)
            # self.dss.text("Redirect PVSystems_" + str(pv_percentage) + ".dss")
            # self.dss.parallel_write_actorcpu(0)
            # self.dss.text("clone 2")
            self.dss.text("Set Parallel=Yes")

            ################## CPU 0
            self.dss.text("set ActiveActor=1")
            self.dss.parallel_write_actorcpu(0)

            self.dss.text("compile [{}]".format(self.dssFileName))
            self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)
            self.dss.text("set DataPath=" + self.results_path)
            self.dss.text("Redirect PVSystems_" + str(pv_percentage) + ".dss")

            Loadshape1 = [LoadshapePointsList[ctd] for ctd in solucoesCPU[0][2:]]
            Loadshape1 = self.LoadshapeToMediaMovel(Loadshape1)
            # print("active actor", self.dss.parallel_read_activeactor(), kWRatedList[solucoesCPU[0][0]], barras[solucoesCPU[0][1]], Loadshape1)
            self.dss.text("Loadshape.Loadshape1.mult=" + str(Loadshape1))
            self.dss.text("Storage.storage.Bus1=" + str(barras[solucoesCPU[0][1]]))
            self.dss.text("Storage.storage.kWrated=" + str(kWRatedList[solucoesCPU[0][0]]))
            self.dss.text("Storage.storage.kva=" + str(kWRatedList[solucoesCPU[0][0]]))
            self.dss.text("Storage.storage.kw=" + str(kWRatedList[solucoesCPU[0][0]]))
            self.dss.text("Storage.storage.kWhrated=" + str(kwhrated))
            self.dss.text("Storage.storage.kWhstored=" + str(kwhstored))
            self.dss.text("Storage.storage.enabled=yes")
            self.dss.solution_buildymatrix()

            ################## CPU 2
            self.dss.parallel_createactor()
            self.dss.text("set ActiveActor=2")
            self.dss.parallel_write_actorcpu(2)

            self.dss.text("compile [{}]".format(self.dssFileName))
            self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)
            self.dss.text("set DataPath=" + self.results_path)
            self.dss.text("Redirect PVSystems_" + str(pv_percentage) + ".dss")

            Loadshape2 = [LoadshapePointsList[ctd] for ctd in solucoesCPU[1][2:]]
            Loadshape2 = self.LoadshapeToMediaMovel(Loadshape2)
            # print("active actor", self.dss.parallel_read_activeactor(), kWRatedList[solucoesCPU[1][0]], barras[solucoesCPU[1][1]], Loadshape2)
            self.dss.text("Loadshape.Loadshape1.mult=" + str(Loadshape2))
            self.dss.text("Storage.storage.Bus1=" + str(barras[solucoesCPU[1][1]]))
            self.dss.text("Storage.storage.kWrated=" + str(kWRatedList[solucoesCPU[1][0]]))
            self.dss.text("Storage.storage.kva=" + str(kWRatedList[solucoesCPU[1][0]]))
            self.dss.text("Storage.storage.kw=" + str(kWRatedList[solucoesCPU[1][0]]))
            self.dss.text("Storage.storage.kWhrated=" + str(kwhrated))
            self.dss.text("Storage.storage.kWhstored=" + str(kwhstored))
            self.dss.text("Storage.storage.enabled=yes")
            self.dss.solution_buildymatrix()

            ################## CPU 3
            self.dss.parallel_createactor()
            self.dss.text("set ActiveActor=3")
            self.dss.parallel_write_actorcpu(3)

            self.dss.text("compile [{}]".format(self.dssFileName))
            self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)
            self.dss.text("set DataPath=" + self.results_path)
            self.dss.text("Redirect PVSystems_" + str(pv_percentage) + ".dss")

            Loadshape3 = [LoadshapePointsList[ctd] for ctd in solucoesCPU[2][2:]]
            Loadshape3 = self.LoadshapeToMediaMovel(Loadshape3)
            # print("active actor", self.dss.parallel_read_activeactor(), kWRatedList[solucoesCPU[2][0]], barras[solucoesCPU[2][1]], Loadshape3)
            self.dss.text("Loadshape.Loadshape1.mult=" + str(Loadshape3))
            self.dss.text("Storage.storage.Bus1=" + str(barras[solucoesCPU[2][1]]))
            self.dss.text("Storage.storage.kWrated=" + str(kWRatedList[solucoesCPU[2][0]]))
            self.dss.text("Storage.storage.kva=" + str(kWRatedList[solucoesCPU[2][0]]))
            self.dss.text("Storage.storage.kw=" + str(kWRatedList[solucoesCPU[2][0]]))
            self.dss.text("Storage.storage.kWhrated=" + str(kwhrated))
            self.dss.text("Storage.storage.kWhstored=" + str(kwhstored))
            self.dss.text("Storage.storage.enabled=yes")
            self.dss.solution_buildymatrix()

            #######
            self.dss.text("set ActiveActor=*")
            self.dss.text("set mode=Daily stepsize=15m number=97")
            self.dss.text("set ActiveActor=1")
            self.dss.text("Set Parallel=Yes")
            self.dss.solution_solveall()

            boolstatus = 0
            while boolstatus == 0:
                if self.dss.parallel_actorstatus() == (1, 1, 1):
                    boolstatus = 1

            loadshapes = [Loadshape1, Loadshape2, Loadshape3]
            for i in range(1, self.dss.parallel_numactors() + 1):
                self.dss.text("set ActiveActor="+str(i))
                # print('actor2', self.dss.parallel_read_activeactor())
                # self.dss.text("export meters")
                # for monitor in self.dss.monitors_allnames():
                #     self.dss.text("export monitor " + monitor)

                # Punicao para maximar a amplitude da loadshape
                maximo = max([abs(min(loadshapes[i-1])), max(loadshapes[i-1])])
                if maximo >= 0.95:
                    PunicaoMaxLoadshape = 0
                elif maximo >= 0.875 and maximo < 0.95:
                    PunicaoMaxLoadshape = 5
                elif maximo >= 0.8 and maximo < 0.875:
                    PunicaoMaxLoadshape = 10
                elif maximo < 0.8:
                    PunicaoMaxLoadshape = 30
                else:
                    PunicaoMaxLoadshape = 0

                # Inclinaçoes
                Inclinacao = 0
                ListaInclinacoes = self.InclinacoesLoadshape(solucoesCPU[i-1])

                for inclinacao in ListaInclinacoes:
                    if numpy.abs(inclinacao) > 40:
                        Inclinacao += numpy.abs(inclinacao)

                # Punição Niveis de Tensão
                if self.BarrasTensaoVioladas() > self.BarrasTensaoVioladasOriginal:
                    PunicaoTensao = 999
                else:
                    PunicaoTensao = 0

                # CICLO DE CARGA DA BATERIA
                # É preciso garantir que ao final das 48h o nível de carregamento da bateria seja o mesmo do inicio da simulacao
                Carregamento48h, PunicaoCicloCarga = self.PunicaoCiclodeCarga()

                ### PERDAS
                self.dss.meters_write_name("Feeder")
                ZonekWh = self.dss.meters_registervalues()[4]
                ZoneLosseskWh = self.dss.meters_registervalues()[12]
                Losses = ZoneLosseskWh / ZonekWh * 100
                # print(self.dataperda['Perdas %'], Losses)

                # DESVIO PADRÃO DO CARREGAMENTO DO TRAFO
                self.dss.monitors_write_name("potencia_feeder")
                DemandaTotal = []
                Pt1 = self.dss.monitors_channel(1)
                Pt2 = self.dss.monitors_channel(3)
                Pt3 = self.dss.monitors_channel(5)
                for iteracao in range(0, len(self.dss.monitors_channel(1))):
                    Pt = Pt1[iteracao] + Pt2[iteracao] + Pt3[iteracao]
                    DemandaTotal.append(Pt)
                Desvio = statistics.pstdev(DemandaTotal)

                # print(self.dss.parallel_read_activeactor())
                Custo = round(Losses + Desvio + Inclinacao + PunicaoCicloCarga + PunicaoMaxLoadshape + PunicaoTensao, 5)
                custos.append((Custo, solucoesCPU[i-1]))

                # custos.append((Custo, solucoesCPU[i-1], self.dss.parallel_read_activeactor(), Desvio, Losses))
                # if:
                #     print("cpu", Custo, solucoesCPU[0]), barras[solucoesCPU[0][1]], kWRatedList[solucoesCPU[0][0]], Loadshape1)

        end_funcaocusto = time.time()
        # custos.sort()
        print('Tempo da geração: ', end_funcaocusto-start_funcaocusto, custos)
        return custos

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

        # for cu in range(1, 82):
        #     populacao.append([2, 263, 8, 14, 3, 0, 0, 11, 12, 17, 4, 3, 11, 16, 26, 20, 30, 29, 22, 33, 32, 23, 37, 35, 24, 26, 15])
        #
        numero_elitismo = int(elitismo * tamanho_populacao)
        geracao = 1
        stop = False
        melhor_solucao = [] # Lista de menor custo por geracao

        while stop == False:
            # print("pop0", populacao)
            custos = self.funcaoCusto(populacao, kWRatedList, barras, pv_percentage)
            custos.sort()
            melhor_solucao.append(custos[0][0])
            if len(melhor_solucao) > int(0.2 * tamanho_populacao): # Criterio de parada
                if abs(melhor_solucao[-16] - melhor_solucao[-1]) < 1:
                    stop=True
            # if melhor_solucao.count(custos[0][0]) == int(0.2 * tamanho_populacao):  # Criterio de parada
            #     stop = True
            custos_traduzidos = [(ctd[0], kWRatedList[ctd[1][0]], barras[ctd[1][1]]) for ctd in custos]
            print("Geração::", geracao,  custos_traduzidos)
            print("Melhores Resultados", melhor_solucao)
            geracao += 1
            individuos_ordenados = [individuo for (custo, individuo) in custos]
            populacao = individuos_ordenados[0:numero_elitismo]
            lista_rank = [(individuo, (tamanho_populacao - individuos_ordenados.index(individuo))/(tamanho_populacao*(tamanho_populacao-1))) for individuo in individuos_ordenados]
            lista_rank.reverse()
            soma=0
            for ctd in lista_rank:
                soma += ctd[1]

            # print("pop", populacao)
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

        Loadshape, Perda, Carregamento, Inclinacao, Tensao, Desvio, kWhRated, Demanda, kWhstored = self.CalculaCustos(custos[0][1], kWRatedList, barras, pv_percentage)

        end2 = time.time()

        results_file = open("Resultados.txt", "a")
        results_file.write(f"{geracao}, {(end2 - start2)/3600}, {custos_traduzidos[0][0]}, {custos_traduzidos[0][1]}, {custos_traduzidos[0][2]}, 50000, {Loadshape}, {Perda}, {Carregamento}, {Inclinacao}, {Tensao}, {Desvio}, {kWhRated}, {kWhstored} \n{Demanda} \n{melhor_solucao} \n")
        results_file.close()

        print("tempo total:", end2 - start2)
        return custos[0][1]

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

    def CalculaCustos(self, solucao, kWRatedList, barras, pv_percentage):
        self.dss.text("ClearAll")
        self.dss.text("compile [{}]".format(self.dssFileName))
        self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)
        self.results_path = self.OpenDSS_folder_path + "/results_Main"
        self.dss.text("set DataPath=" + self.results_path)

        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]
        Loadshape = [LoadshapePointsList[ctd] for ctd in solucao[2:]]
        Loadshape = self.LoadshapeToMediaMovel(Loadshape)

        self.dss.text("Redirect PVSystems_" + str(pv_percentage) + ".dss")
        self.dss.text("Loadshape.Loadshape1.mult=" + str(Loadshape))
        self.dss.text("Storage.storage.Bus1=" + str(barras[solucao[1]]))
        self.dss.text("Storage.storage.kWrated=" + str(kWRatedList[solucao[0]]))
        self.dss.text("Storage.storage.kva=" + str(kWRatedList[solucao[0]]))
        self.dss.text("Storage.storage.kw=" + str(kWRatedList[solucao[0]]))
        self.dss.text("Storage.storage.kWhrated=50000")
        self.dss.text("Storage.storage.kWhstored=30000")
        self.dss.text("Storage.storage.enabled=yes")

        self.dss.text("set mode=Daily stepsize=15m number=97")
        self.dss.parallel_write_activeparallel(0)
        self.dss.text("Solve")

        ### DEFININDO CAPACIDADE DE ARMAZENAMENTO
        self.dss.monitors_write_name("storage")
        kWh = self.dss.monitors_channel(1)
        maxkWh = max(kWh)
        minkWh = min(kWh)
        kWhRated = (maxkWh-minkWh)/0.7
        kwhstored = 30000-minkWh+0.25*kWhRated

        ### SIMULACAO COM A NOVA CAP ARM
        self.dss.text("ClearAll")
        self.dss.text("compile [{}]".format(self.dssFileName))
        self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)
        self.results_path = self.OpenDSS_folder_path + "/results_Main"
        self.dss.text("set DataPath=" + self.results_path)

        self.dss.text("Redirect PVSystems_" + str(pv_percentage) + ".dss")
        self.dss.text("Loadshape.Loadshape1.mult=" + str(Loadshape))
        self.dss.text("Storage.storage.Bus1=" + str(barras[solucao[0]]))
        self.dss.text("Storage.storage.kWrated=" + str(kWRatedList[solucao[0]]))
        self.dss.text("Storage.storage.kva=" + str(kWRatedList[solucao[0]]))
        self.dss.text("Storage.storage.kw=" + str(kWRatedList[solucao[0]]))
        self.dss.text("Storage.storage.kWhrated=" + str(kWhRated))
        self.dss.text("Storage.storage.kWhstored=" + str(kwhstored))
        self.dss.text("Storage.storage.enabled=yes")

        self.dss.text("set mode=Daily stepsize=15m number=97")
        self.dss.parallel_write_activeparallel(0)
        self.dss.text("Solve")

        # CICLO DE CARGA DA BATERIA
        # É preciso garantir que ao final das 48h o nível de carregamento da bateria seja o mesmo do inicio da simulacao
        Carregamento48h, PunicaoCicloCarga = self.PunicaoCiclodeCarga()

        # Inclinaçoes
        Inclinacao = 0
        ListaInclinacoes = self.InclinacoesLoadshape(solucao)

        for i in ListaInclinacoes:
            if numpy.abs(i) > 40:
                Inclinacao += numpy.abs(i)

        ### PERDAS
        self.dss.meters_write_name("Feeder")
        ZonekWh = self.dss.meters_registervalues()[4]
        ZoneLosseskWh = self.dss.meters_registervalues()[12]
        Losses = ZoneLosseskWh / ZonekWh * 100

        # DESVIO PADRÃO DO CARREGAMENTO DO TRAFO
        self.dss.monitors_write_name("potencia_feeder")
        DemandaTotal = []
        Pt1 = self.dss.monitors_channel(1)
        Pt2 = self.dss.monitors_channel(3)
        Pt3 = self.dss.monitors_channel(5)
        for iteracao in range(0, len(self.dss.monitors_channel(1))):
            Pt = Pt1[iteracao] + Pt2[iteracao] + Pt3[iteracao]
            DemandaTotal.append(Pt)
        Desvio = statistics.pstdev(DemandaTotal)

        # print('Perdas:', self.dataperda['Perdas %'], 'kWh 48h:', Carregamento48h, 'Inclinação:', Inclinacao, 'Barras_Violada:', self.BarrasTensaoVioladas(), 'Desvio:', Desvio, 'PTotal:', dataFeederMmonitorCSV['PTotal'])
        # print(self.LoadshapeToMediaMovel(Loadshape),",", self.dataperda['Perdas %'],",", Carregamento48h,",", Inclinacao,",", self.BarrasTensaoVioladas(),",", Desvio, ",", kWhRated)
        # print(dataFeederMmonitorCSV['PTotal'])
        # print('Loadshape:', self.LoadshapeToMediaMovel(Loadshape))
        return str(Loadshape).replace("[", "").replace("]", ""), Losses, Carregamento48h, Inclinacao, self.BarrasTensaoVioladas(), Desvio, kWhRated, DemandaTotal, kwhstored

    def CalculaCustosOriginal(self, porcentagem_prosumidores):
        self.dss.text("ClearAll")
        self.dss.text("compile [{}]".format(self.dssFileName))
        self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)
        self.results_path = self.OpenDSS_folder_path + "/results_Main"
        self.dss.text("set DataPath=" + self.results_path)

        self.dss.text("Storage.storage.enabled=no")
        self.dss.text("Redirect PVSystems_" + str(porcentagem_prosumidores) + ".dss")
        self.dss.text("set mode=Daily stepsize=15m number=97")

        self.dss.parallel_write_activeparallel(0)
        self.dss.text("Solve")

        ### PERDAS
        self.dss.meters_write_name("Feeder")
        ZonekWh = self.dss.meters_registervalues()[4]
        ZoneLosseskWh = self.dss.meters_registervalues()[12]
        Losses = ZoneLosseskWh/ZonekWh*100
        # print(self.dataperda['Perdas %'], Losses)

        ### DEMANDA DO TRAFO AT MT
        self.dss.monitors_write_name("potencia_feeder")
        DemandaTotal = []
        Pt1 = self.dss.monitors_channel(1)
        Pt2 = self.dss.monitors_channel(3)
        Pt3 = self.dss.monitors_channel(5)
        for iteracao in range(0, len(self.dss.monitors_channel(1))):
            Pt = Pt1[iteracao] + Pt2[iteracao] + Pt3[iteracao]
            DemandaTotal.append(Pt)

        barrasVioladas = self.BarrasTensaoVioladas()

        print('Custos Sistema Original (Somente GD-PV)')
        print('Perdas:', Losses, 'Violações de Tensao:', barrasVioladas, 'PTotal:', DemandaTotal, '\n')
        return barrasVioladas

    def BarrasTensaoVioladas(self, actor=1):
        BarrasVioladas = 0

        for trafo in self.dss.transformers_allNames():
            self.dss.transformers_write_name(trafo)
            if self.dss.transformers_read_kv() != 13.8:
                self.dss.monitors_write_name(str(trafo))
                V1 = self.dss.monitors_channel(1)
                V2 = self.dss.monitors_channel(2)
                V3 = self.dss.monitors_channel(3)
                Vtotal = [x/127 for x in list(V1 + V2 + V3)]
                # Vtotal.sort()
                # print(Vtotal)
                for ctd in Vtotal:
                    if ctd > 1.03 or ctd < 0.95:
                        BarrasVioladas += 1
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

    def InclinacoesLoadshape(self, solucao):
        LoadshapePointsList = [round(ctd, 2) for ctd in list(numpy.arange(-1.0, 1.05, 0.05))]
        Loadshape = [LoadshapePointsList[i] for i in solucao[2:]]
        # Loadshape = [LoadshapePointsList[i] for i in solucao[1:]]
        Inclinacoes = []

        for i in range((len(Loadshape))):
            if i == 24:
                x = Loadshape[0] - Loadshape[24]
            else:
                x = Loadshape[i+1] - Loadshape[i]
            Inclinacoes.append(numpy.arctan(x)*180/pi)
        return Inclinacoes

    def PunicaoCiclodeCarga(self, kwhstored=30000):
        self.dss.monitors_write_name("storage")
        Carregamento48h = self.dss.monitors_channel(1)[-1]
        PunicaoCicloCarga = abs((kwhstored-Carregamento48h)/100)

        return Carregamento48h, PunicaoCicloCarga


if __name__ == '__main__':
    d = DSS(r"C:\Users\jonas\PycharmProjects\AG_GBA01F3\ARB_GBA01F3_2019\GBA01F3_2019.dss")

    d.dss.dss_clearall()
    d.dss.text("compile [{}]".format(d.dssFileName))

    barras = []
    for bus in list(d.dss.circuit_allbusnames()):
        d.dss.circuit_setactivebus(bus)
        if len(d.dss.bus_nodes()) == 3 and ('gba' not in bus) and ('m784' in bus):
            barras.append(d.dss.bus_name())
    # print(len(barras))

    pv_percentage = 0.2
    kWRatedList = list(range(100, 5000, 200))
    kwHRatedList = list(range(1000, 35000, 500))
    dominio = [(0, len(kWRatedList) - 1), (0, len(barras) - 1), (0, 40) , (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40),  (0, 40), (0, 40)]
    # dominio = [(0, len(kWRatedList) - 1), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40),(0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40), (0, 40)]

    d.genetico(pv_percentage, kWRatedList, barras, dominio)
