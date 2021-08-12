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

    def Cenario(self, porcentagem_prosumidores):
        self.compile_DSS()
        # self.dss.text("Storage.storage.enabled=no")

        self.dss.text("Solve")

        pv_file = open("PVSystems_" + str(porcentagem_prosumidores) + ".dss", "w")
        # Cargas e Barras
        loadlist = []
        loaddict = {}
        somaload=0
        somaPmpp=0
        for load in self.dss.loads_allnames():
            self.dss.loads_write_name(load)
            kvbase = self.dss.loads_read_kv()
            numphases = self.dss.cktelement_numphases()
            bus = str(self.dss.cktelement_read_busnames()).replace("'", "").replace('(', "").replace(')', "").replace(',',
                                                                                                                      "")
            curva = self.dss.loads_read_daily()
            self.dss.loadshapes_write_name(curva)
            Epv = 7.89 * 0.97 ** 2  # capacidade de geracao
            Ec = 0  # consumo diario medio
            for i in self.dss.loadshapes_read_pmult():
                Ec += i * self.dss.loads_read_kw() *0.25
            # print(load, self.dss.loads_read_kw(), len(self.dss.loadshapes_read_pmult()))
            pmpp = round(Ec / Epv, 2)
            loaddict[load] = [numphases, bus, kvbase, pmpp]
            somaload+=self.dss.loads_read_kw()
            somaPmpp+=pmpp

            loadlist.append((Ec, load))
        loadlist.sort(reverse=True)
        # print('loadlist', loadlist)
        print("Media carga", somaload/len(self.dss.loads_allnames()))
        print(somaPmpp)

        # Seleção por Roleta dos Prossumidores
        fim = round(len(loadlist) * porcentagem_prosumidores)
        # print('fim', fim)
        print(loadlist)
        prossumidores = []
        while len(prossumidores) < fim:
            soma = 0
            for ctd in loadlist:
                soma += ctd[0]
            # print(soma)

            aleatorio = random.uniform(0, soma)
            # print('aleatorio', aleatorio)
            s = 0
            for j in loadlist:
                s += j[0]
                if aleatorio < s:
                    prossumidor = j
                    # print('prossumidor', prossumidor)
                    break
            prossumidores.append(prossumidor[1])
            loadlist.remove(prossumidor)

        print('Prossumidores', prossumidores)

        # Inserindo os PVsystems
        ctd = 0
        for load in prossumidores:
            pv_file.write(
                f"New PVSystem.PV{ctd} phases={loaddict[load][0]} Bus1={loaddict[load][1]} kV={loaddict[load][2]} kVA={loaddict[load][3]} Pmpp={loaddict[load][3]} conn=wye PF = 1 %cutin = 0.00005 %cutout = 0.00005 effcurve = Myeff P-TCurve = MyPvsT Daily = MyIrrad TDaily = Mytemp \n")
            ctd += 1
        pv_file.close()


if __name__ == '__main__':
    d = DSS(r"C:\Users\jonas\PycharmProjects\AG_GBA01F3\ARB GBA01F3 2019\GBA01F3 2019.dss")

    # CRIAÇÃO DOS CENARIOS
    for i in range(10,100,10):
        i = i/100
        d.Cenario(i)