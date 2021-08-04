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

    def compile_DSS(self):
        self.dss.text("ClearAll")
        self.dss.text("set parallel=no")
        self.dss.text("compile [{}]".format(self.dssFileName))
        self.dss.parallel_write_actorcpu(0)
        self.dss.text("Solve")
        # OpenDSS folder
        self.OpenDSS_folder_path = os.path.dirname(self.dssFileName)

    def solve(self, pv_percentages):
        self.dss.text("clone 1")
        print("Numactors", self.dss.parallel_numactors())

        self.dss.parallel_write_activeactor(1)
        # self.dss.text("set activeactor=1")
        print("active actor", self.dss.parallel_read_activeactor())
        self.dss.parallel_write_actorcpu(1)
        self.dss.text("set mode=Daily stepsize=15m number=97")
        self.dss.text("Redirect PVSystems_" + str(pv_percentages[0]) + ".dss")
        self.dss.text("Storage.storage.enabled=no")

        self.dss.parallel_write_activeactor(2)
        # self.dss.text("set activeactor=2")
        print("active actor", self.dss.parallel_read_activeactor())
        self.dss.parallel_write_actorcpu(2)
        self.dss.text("set mode=Daily stepsize=15m number=97")
        self.dss.text("Redirect PVSystems_" + str(pv_percentages[1]) + ".dss")
        self.dss.text("Storage.storage.enabled=no")

        # self.dss.parallel_write_activeactor(3)
        # # self.dss.text("set activeactor=3")
        # print("active actor", self.dss.parallel_read_activeactor())
        # self.dss.parallel_write_actorcpu(3)
        # self.dss.text("set mode=Daily stepsize=15m number=9700")
        # # self.dss.text("Redirect PVSystems_" + str(pv_percentages[2]) + ".dss")
        # self.dss.text("Storage.storage.enabled=no")

        self.dss.parallel_write_activeactor(1)
        self.dss.parallel_write_activeparallel(1)
        self.dss.solution_solveall()
        boolstatus = 0
        print(self.dss.parallel_actorstatus())
        while boolstatus == 0:
            if self.dss.parallel_actorstatus() == (1, 1):
                boolstatus = 1
        print(self.dss.parallel_actorstatus())

        for i in range(1, self.dss.parallel_numactors()+1):
            print(i)
            self.dss.parallel_write_activeactor(i)
            # self.dss.text("set activeactor=" + str(i))
            print("active actor", self.dss.parallel_read_activeactor())
            self.dss.text("export monitor Potencia_Feeder")

            self.dss.text("Plot monitor object=Potencia_Feeder")

        print("solved")


if __name__ == '__main__':
    d = DSS(r"C:\Users\jonas\PycharmProjects\AG_GBA01F3\ARB_GBA01F3_2019\GBA01F3_2019.dss")

    pv_percentages = [0.1, 0.2, 0.3]

    d.compile_DSS()
    d.solve(pv_percentages)