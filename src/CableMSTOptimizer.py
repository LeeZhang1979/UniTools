# coding:utf-8
# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import math
import os
import time
import json
from shutil import copyfile
import wx

FIRST_COL_LABEL = 45
FIRST_COL_TEXT = 115

SECOND_COL_LABEL = 230
SECOND_COL_TEXT = 320

THIRD_COL_LABEL = 430
THIRD_COL_TEXT = 510

FOURTH_COL_LABEL = 620
FOURTH_COL_TEXT = 710

LABEL_WIDTH = 70
TEXT_WIDTH = 90

LABEL_TEXT_HEIGHT = 20
FIRST_ELE_X = 10
FIRST_ELE_Y = 10

LINE_HEIGHT = 3
LINE_WIDTH = 840

input_n_wt = 45  # 风机数量
input_rated_power = 4  # 单个风机容量
input_open_year = 25  # 预计运行年限
input_wt_max_cable = 2  # 机组最多输出电缆数
input_u_col = 35  # 额定电压
input_u_trans = 220  # 并网电压
input_pf = 0.9  # 并网点功率因数
input_price_cin_intl = 30  # 集电系统安装费用（万元/km）
input_price_trn_intl = 138  # 输电系统安装费用（万元/km）
input_price_sea = 0.7  # 海域占用费用（元/mm2）
input_with_cable = 20  # 海域宽度（m）


def CableOptData(event):
    settings = {
        "id": "testWF",
        # cable data
        "cable_name": ['HYJQ41-3x70', 'HYJQ41-3x120', 'HYJQ41-3x185', 'HYJQ41-3x240',
                       'HYJQ41-3x300', 'HYJQ41-3x400', 'HYJQ41-3x800'],
        "cable_voltage": [35, 35, 35, 35, 35, 35, 220],
        "cable_cross_section": [70, 120, 185, 240, 300, 400, 800],
        "cable_price": [1180000, 1370000, 1930000, 2170000, 2350000, 2950000, 1880000],
        "cable_ampacity": [260, 349, 433, 500, 545, 580, 800],
        "cable_conductor_resistance": [0.31, 0.159, 0.123, 0.098, 0.074, 0.061, 0.05],
        "cable_inductive_reactance": [0.138, 0.122, 0.119, 0.113, 0.107, 0.912, 0.06],

        # wind farm data
        "n_wt": int(input_n_wt),
        "rated_power": int(input_rated_power),
        "oper_year": int(input_open_year),
        "utilization_hour": 2927,
        "wt_max_cable": int(input_wt_max_cable),
        "d_f": 0.05,
        # loss cost
        "loss_cost": 1,
        "Ce": 850,
        "N_A": 0.7,
        "u_col": int(input_u_col),
        # transimission cost
        "transmission_cost": 1,
        "u_trans": int(input_u_trans),
        "pf": float(input_pf),
        "fc": 1,
        "k_length": 1,
        # installation cost
        "installation_cost": 460000,
        "priceClnIntl": int(input_price_cin_intl),
        "priceTrnIntl": int(input_price_trn_intl),
        # sea cost
        "sea_cost": 1,
        "priceSea": float(input_price_sea),
        "wthCable": int(input_with_cable),
        "original_layout_x": [627706.1687, 627510.0258, 624592.7677, 624787.2221, 624398.3133, 627315.5714, 629843.9208,
                              629649.4664, 626926.6627, 627121.1170, 626732.2083, 624981.6764, 621821.1291, 623693.0843,
                              622015.5835, 621626.6747, 622210.0378, 623498.6300, 625370.5852, 625176.1308, 622915.2669,
                              623304.1756, 623109.7212, 629455.012, 633247.7573, 620750.5136, 633247.7573, 633247.7573,
                              633247.7573, 620750.5136, 620750.5136, 633247.7573, 620750.5136, 620750.5136, 620750.5136,
                              632226.6506, 623887.5387, 631326.9672, 629066.1033, 629260.5577, 628871.6489, 631132.5129,
                              632421.105, 632032.1963, 630549.1498],
        "original_layout_y": [2384625.9858, 2385099.5166, 2386916.1487, 2386446.6944, 2387385.6031, 2385568.9710,
                              2384691.2475, 2385160.7019, 2386507.8797, 2386038.4253, 2386977.3340, 2385977.2400,
                              2385768.0984, 2385168.4877, 2385298.6441, 2386237.5528, 2384829.1897, 2385637.9420,
                              2385038.3313, 2385507.7856, 2387046.3051, 2386107.3964, 2386576.8508, 2385630.1563,
                              2386376.1628, 2384599.3779, 2385826.1628, 2384726.1628, 2385276.1628, 2385178.2215,
                              2387378.2215, 2386926.1628, 2386828.2215, 2385728.2215, 2386278.2215, 2386778.2066,
                              2384699.0333, 2385030.5455, 2386569.0650, 2386099.6106, 2387038.5194, 2385499.9999,
                              2386308.7522, 2387247.6609, 2386908.3630],
        "original_layout_no": ['WT' + str(i) for i in list(np.arange(int(input_n_wt)) + 1)],
        "OffStn_id": ['OSS1'],
        "OffStn_x": [627977.984],
        "OffStn_y": [2387103.422],
        "OnStn_id": [],
        "MST_iter": 45,
        "cable_PSO_iter": 0,
        "cable_swarmsize": 10
    }
    CableOpt = CableMSTOptimizer(settings)
    CableOpt.MSTCableOpt()
    CableOpt.output()


class CableMSTOptimizer(object):
    def __init__(self, settings):
        self.start = time.time()
        self.id = settings["id"]
        self.save_path = './outputs/%s/' % self.id
        self.save_path = os.path.realpath(self.save_path)
        if not os.path.isdir(os.path.join(self.save_path, 'monitoring')):
            os.makedirs(os.path.join(self.save_path, 'monitoring'))
        cablespec = pd.DataFrame({"name": settings["cable_name"],
                                  "voltage": [float(v) for v in settings["cable_voltage"]],
                                  "cross_section": [float(c) for c in settings["cable_cross_section"]],
                                  "price": [float(p) for p in settings["cable_price"]],
                                  "ampacity": [float(a) for a in settings["cable_ampacity"]],
                                  "resistance": [float(r) for r in settings["cable_conductor_resistance"]],
                                  "X": [float(x) for x in settings["cable_inductive_reactance"]]})
        self.cablespec = cablespec.sort_values(by=["cross_section"])
        Prated = settings["rated_power"]  # [MW],          Rated Power of Wind Turbines
        self.Nt = int(settings["n_wt"])  # wind turbine number
        self.Ny = settings["oper_year"]  # [year],        Lifetime of the WF operation for energy losses
        self.Nh = settings["utilization_hour"]  # [hour/year],   Utilization factor of WF’s installed capacity
        # self.Ntin = settings["wt_max_cable"]  # [pcs],         Maximum number of WT's cables
        d_f = settings["d_f"]
        d_ff = 0
        for i in range(self.Ny):
            d_ff = 1 / (1 + d_f) ** i + d_ff
        self.d_ff = d_ff
        if settings["loss_cost"]:
            self.loss_cost = 1
            self.Ce = settings["Ce"]  # [rmb/MWh],     Electricity price (China 2015, 388.25rmb/MWh)
            self.N_A = settings["N_A"]  # []             Current correction factor
        else:
            self.loss_cost = 0
            # self.Ce = None
            # self.N_A = None
        self.U_col = settings["u_col"]  # [kV],          Voltage of the collection nets
        if settings["transmission_cost"]:  # [kV],          Voltage of the transmission nets
            self.transmission_cost = 1
            self.U_trans = settings["u_trans"]
        else:
            self.transmission_cost = 0
            # self.U_trans = None
        self.PF = settings["pf"]  # [],            Power factor
        self.fc = settings["fc"]  # [],            Cable current carrying capacity correction factor
        self.k_length = settings["k_length"]  # [],            Ratio of cable length and distance
        if settings["installation_cost"]:
            self.installation_cost = 1
            self.Ct = settings["priceClnIntl"] * 10000
        else:
            self.installation_cost = 0
            # self.priceClnIntl = None
        # self.Ct = settings["priceTrnIntl"]  # [万元/km],      Price of Collection cable installation
        if settings["sea_cost"]:
            self.sea_cost = 1
            self.priceSea = settings["priceSea"]  # [RMB/mm2]      Price of the sea
            self.wthCable = settings["wthCable"]  # [m]            Width of Cable
        else:
            self.sea_cost = 0

        self.In = Prated / np.sqrt(3) / self.U_col / self.PF * self.fc  # [kA], nominal current of single turbine
        self.Nb_max = math.floor(np.max(self.cablespec[self.cablespec["voltage"] < 100][
                                            "ampacity"].values) / 1e3 / self.In)  # maximum allowable per unit power, 14 for case for paper
        Cs = self.cablespec["voltage"].values < 220  # cable types that
        self.Ic = self.cablespec["ampacity"][
                      Cs] / 1e3  # [kA], cable current carrying capacity at different cross-section
        Ncs = np.size(self.Ic)

        wtx = np.array([float(x) for x in settings["original_layout_x"]]) / 1e3
        wty = np.array([float(y) for y in settings["original_layout_y"]]) / 1e3
        no = settings["original_layout_no"]
        self.layout = pd.DataFrame({'X': wtx, "Y": wty, "no": no})
        WT = self.layout.loc[:, "X":"Y"]
        # WT = pd.DataFrame({"X": self.wtx, "Y": self.wty})
        if len(settings["OffStn_id"]) > 0:
            OS = pd.DataFrame({"X": np.array([float(x) for x in settings["OffStn_x"]]) / 1e3,
                               "Y": np.array([float(x) for x in settings["OffStn_y"]]) / 1e3})
        else:
            OS = pd.DataFrame({"X": np.mean(wtx), "Y": np.mean(wty)})
        self.Gn = pd.concat([OS, WT])  # calculate in km
        self.num_OS = OS.shape[0]
        self.OS = self.Gn.iloc[0, :].values  # select the first OS
        OS["no"] = settings["OffStn_id"]
        OS["index"] = np.arange(len(settings["OffStn_id"])) + 1
        self.WT = self.Gn.iloc[self.num_OS:, :].values

        # %% inner array cable model with cost
        price_idx = np.zeros([3, Ncs + 1])
        price_idx[0, :-1] = [math.floor(x) for x in self.cablespec["ampacity"][Cs] / self.In / 1e3]
        # math.floor(cablespec[:,3][Cs]/In/1e3)
        # [print(a) for a in range(1,6)]
        price_idx[0, -1] = self.Nt + 1
        price_idx[1, :] = np.array(self.cablespec["price"].values)  # cable cost [Rmb/km]
        price_idx[2, :] = np.array(self.cablespec["cross_section"].values)
        self.price_idx = price_idx

        if len(settings["OnStn_id"]) > 0:
            # self.OnStn = [settings["OnStn_x"][0], settings["OnStn_y"][0]]
            sub_on_x = settings["OnStn_x"][0] / 1e3
            sub_on_y = settings["OnStn_y"][0] / 1e3
        else:
            # %% offshore substation coordinate
            # Sub_x = OS[0, 0]
            # Sub_y = OS[0, 1]

            # manual design onshore substation
            sub_on_x = self.OS[0] - 14
            sub_on_y = self.OS[1] - 20

            # %% main programme
        self.STS = np.sqrt(pow((sub_on_x - self.OS[0]), 2) + pow((sub_on_y - self.OS[1]), 2))
        index = list(np.arange(self.Nt) + len(settings["OffStn_id"]) + 1)
        self.layout["index"] = index
        self.layout = pd.concat([OS, self.layout])
        self.MST_iter = settings["MST_iter"]
        self.PSO_iter = 0  # settings["cable_PSO_iter"]
        self.PSO_swarm = settings["cable_swarmsize"]
        self.structure_hist = []
        self.cost_hist = []
        self.index = []

    def zoneDivision(self, zone_rot):
        """if layout is not None:
            self.WT = layout / 1e3
            self.Gn = pd.concat(self.OS, self.WT) / 1e3
        if OS is not None:
            self.OS = OS / 1e3
            self.Gn = pd.concat(self.OS, self.WT)"""
        k = np.zeros(self.Nt)
        theta = np.zeros(self.Nt)
        for i in range(self.Nt):
            k[i] = (self.WT[i, 1] - self.OS[1]) / (self.WT[i, 0] - self.OS[0])
            theta[i] = math.atan(k[i]) * 180 / math.pi
            if self.WT[i, 0] < self.OS[0] and self.WT[i, 1] > self.OS[1]:
                theta[i] = 180 - abs(theta[i])
            elif self.WT[i, 0] < self.OS[0] and self.WT[i, 1] < self.OS[1]:
                theta[i] = theta[i] + 180
            elif self.WT[i, 0] > self.OS[0] and self.WT[i, 1] < self.OS[1]:
                theta[i] = 360 - abs(theta[i])
        k_angle_org = np.sort(theta)
        idx_angle_org = np.argsort(theta)
        k_angle_list = k_angle_org.tolist()
        idx_angle_list = idx_angle_org.tolist()

        k_angle_upd = k_angle_list[zone_rot:] + k_angle_list[:zone_rot]
        idx_angle_upd = idx_angle_list[zone_rot:] + idx_angle_list[:zone_rot]

        k_angle = np.array(k_angle_upd)
        idx_angle = np.array(idx_angle_upd)

        # % manual design
        # N_gr = 8  # number of zones
        # if N_gr < math.ceil(Nt / self.Nb_max):
        N_gr = math.ceil(self.Nt / self.Nb_max)  # N_gr确定
        N_last = self.Nt - (N_gr - 1) * self.Nb_max
        if N_last == 0:
            N_gr = N_gr - 1

        # end
        Nb_max = math.floor(self.Nt / N_gr)
        Nb_reg = (self.Nt % N_gr)
        N_last = self.Nt - (N_gr - 1) * Nb_max
        if N_last == 0:
            N_gr = N_gr - 1

        G = []
        idx_gr = []
        for i in range(N_gr):
            if i <= (Nb_reg - 1):
                G.append(k_angle[(Nb_max + 1) * (i):(Nb_max + 1) * (i + 1)])
                idx_gr.append(idx_angle[(Nb_max + 1) * (i):(Nb_max + 1) * (i + 1)])
            else:
                G.append(k_angle[((Nb_max + 1) * Nb_reg + Nb_max * (i - Nb_reg)):(Nb_max + 1) * Nb_reg + Nb_max * (
                        i - Nb_reg + 1)])
                idx_gr.append(idx_angle[(Nb_max + 1) * Nb_reg + Nb_max * (i - Nb_reg):(Nb_max + 1) * Nb_reg + Nb_max * (
                        i - Nb_reg + 1)])
        return [N_gr, idx_gr]

    def MSTFormulation(self, PoSWt, idx):
        # Pos
        OS = self.OS
        Pos_gr = PoSWt  # position gcrosroup

        Dis_oSWt = np.zeros(np.size(PoSWt, 0))
        for ii in range(np.size(PoSWt, 0)):  # i=0-3
            Dis_oSWt[ii] = np.sqrt(
                np.square(Pos_gr[ii, 0] - OS[0]) + np.square(Pos_gr[ii, 1] - OS[1]))  # Distance between oS and Wt

        Dis_wt_idx = np.argsort(Dis_oSWt)
        Dis_wt_idx_upd = idx[Dis_wt_idx]
        Pos_gr_upd = Pos_gr[Dis_wt_idx, :]  # Updated position of group

        # the cost optimization and OS selection programme
        size_col = np.size(Pos_gr_upd[:, 0])
        Dis_col = np.zeros([size_col, size_col])
        for i in range(size_col):
            for j in range(size_col):
                Dis_col[i, j] = np.sqrt(
                    np.square(Pos_gr_upd[i, 0] - Pos_gr_upd[j, 0]) + np.square(Pos_gr_upd[i, 1] - Pos_gr_upd[j, 1]))

        connect_gr = np.array([-1, 1, np.min(Dis_oSWt), 1, 10])
        node_C = np.zeros(1)  # Counted node
        node_U = np.arange(1, np.size(Dis_col, 0))  # Uncounted node
        tmp = 1
        cnt = 1
        S_p = 0
        while np.size(connect_gr, 0) != np.size(Dis_col, 0):
            if node_C.size == 1:
                DisMT = Dis_col[0, node_U]
            else:
                node_C_list = list(node_C)
                node_U_list = list(node_U)
                DisMT_upd = Dis_col[node_C_list]
                DisMT = DisMT_upd[:, node_U_list]
            if node_C.size == 1:
                node_C_idx = 0
                node_U_idx = np.where(DisMT == min(DisMT))
                node_U_idx = node_U_idx[0]
            elif node_U.size == 1:
                node_C_idx = np.where(DisMT == min(DisMT))
                node_U_idx = 0
                node_C_idx = node_C_idx[0]
            else:
                [node_C_idx, node_U_idx] = np.where(DisMT == np.min(DisMT))
                node_C_idx = node_C_idx[0]
                node_U_idx = node_U_idx[0]
            node_U = np.array(node_U)
            if node_C.size == 1:
                idx_C = node_C
            else:
                idx_C = node_C[node_C_idx]

            if node_U.size == 1:
                idx_U = node_U
            else:
                idx_U = node_U[node_U_idx]

            connect_gr = np.vstack(
                (connect_gr, np.array([int(idx_C + 1), int(idx_U + 1), np.min(DisMT), 1, self.price_idx[2, 0]])))
            # this should be changed for different voltage cable

            tmp = idx_C + 1
            while tmp != -1:
                idx_out = np.where(connect_gr[:, 1] == tmp)
                idx_out = idx_out[0]
                idx_in = np.where(connect_gr[:, 0] == tmp)
                idx_in = idx_in[0]
                tmp = connect_gr[idx_out, 0]
                tmp = int(tmp)
                connect_gr[idx_out, 3] = connect_gr[idx_out, 3] + 1
                S_area = np.argwhere(self.price_idx[0, :] >= connect_gr[idx_out, 3])[0]
                # S_area=0
                connect_gr[idx_out, 4] = self.price_idx[2, S_area]

            node_C = np.append(node_C, idx_U)
            node_C = node_C.astype(int)
            node_U = np.delete(node_U, node_U_idx, axis=0)
            cnt = cnt + 1

        #
        num_col1 = np.size(connect_gr, 0)  # result_2 5*4  num_col1 4
        #

        # # calculate the total price of collection cables
        Col_price = []
        for ii in range(num_col1):
            idx_Cco = np.where(self.price_idx[0, :] >= (connect_gr[ii, 3]))
            idx_Cco = idx_Cco[0][0]

            Col_price_ii = self.price_idx[1, idx_Cco] * connect_gr[ii, 2]
            Col_price.append(Col_price_ii)
        #
        cost_col = np.sum(Col_price)

        d_idx = np.where(self.price_idx[0, :] >= (connect_gr[0, 3]))
        d_idx = d_idx[0][0]

        Co_OStoWT_1st = np.sqrt(np.square(Pos_gr_upd[0, 0] - OS[0]) + np.square(Pos_gr_upd[0, 1] - OS[1])) * \
                        self.cablespec["price"][
                            self.cablespec["cross_section"] == self.cablespec["cross_section"][d_idx]].values
        cost_col_all = cost_col + Co_OStoWT_1st

        # # total cost of trench
        if self.installation_cost:
            cost_trench = self.Ct * np.sum(connect_gr[1:, 2])
        else:
            cost_trench = self.installation_cost
        # # total energy loss on the cables

        Ic = self.cablespec[
            "ampacity"]  # cablespec[:, 3] / 1e3  # [kA], cable current carrying capacity at different cross-section
        Rc = self.cablespec["resistance"]  # [ohm/km], cable resistance at different cross-section
        Xc = self.cablespec["X"]  # [ohm/km], cable reactance at different cross-section
        In = self.In  # [kA], nominal current of single turbine
        # # energyloss1 = In*In*a;
        energyloss = []
        for k in range(np.size(connect_gr[:, 3])):
            energyloss_k = Rc[self.cablespec["cross_section"] == connect_gr[k, 4]].values[0] * connect_gr[
                k, 3] * In * In * np.square(
                connect_gr[k, 3])
            energyloss.append(energyloss_k)

        if self.loss_cost:
            energyloss_OStoWT1st = Rc[self.cablespec["cross_section"] == self.cablespec.iloc[
                d_idx, 2]].values[0] * In * In * np.square(connect_gr[0, 3])
            # energyloss_OStoWT1st = Rc[d_idx] * In * In * np.square(connect_gr[0, 3])
            # # NPV of energy cost
            cost_energyloss1 = self.Nh * self.Ce * 3

            #
            cost_energyloss = (np.sum(energyloss) + energyloss_OStoWT1st) * cost_energyloss1 * self.d_ff
        else:
            cost_energyloss = self.loss_cost
        # # total cost
        if self.sea_cost:
            sea_cost_col = self.priceSea * np.sum(connect_gr[1:, 2]) * self.wthCable

        else:
            sea_cost_col = 0
            sea_cost_trans = 0
        totalcost = cost_trench + cost_col_all + cost_energyloss + sea_cost_col
        connect_gr[0, 0] = 1
        connect_gr[1:, 0] = Dis_wt_idx_upd[(connect_gr[1:, 0] - 1).astype(int)] + 2
        connect_gr[:, 1] = Dis_wt_idx_upd[(connect_gr[:, 1] - 1).astype(int)] + 2
        connect = connect_gr

        return [totalcost, cost_trench, cost_col_all, cost_energyloss, connect, sea_cost_col]

    def MSTCableOpt(self, layout=None, OS=None, STS=None):
        cost_col, Cost_col = [], []
        # Gbest_cost_col = 1e13
        # self.cost_hist.append(1e13, 1e13)
        iter = min(self.Nt, self.MST_iter)
        evol_hist_file = os.path.join(self.save_path, 'evol_hist.txt')
        f = open(evol_hist_file, 'w+')
        f.close()
        min_cost_col = float("inf")
        min_connect = []
        min_k = 0
        min_N_gr = 0
        min_totalcost = 0
        for k in range(iter):

            N_gr, idx_gr = self.zoneDivision(k)
            TC, CTren, Cco, Cen, Connect = [], [], [], [], []
            cost_sea_col, cost_sea_trans = [], []
            for i in range(N_gr):
                totalcost, cost_trench, cost_col_all, cost_energyloss, connect, sea_cost_col = self.MSTFormulation(
                    self.WT[idx_gr[i]],
                    idx_gr[i])
                TC.append(totalcost)
                CTren.append(cost_trench)
                Cco.append(cost_col_all)
                Cen.append(cost_energyloss)
                Connect.append(connect)
                cost_sea_col.append(sea_cost_col)
                # cost_sea_trans.append(sea_cost_trans)

            cost_col = np.sum(Cco[:])
            if cost_col < min_cost_col:
                min_cost_col = cost_col
                min_connect = Connect
                min_k = k
                min_N_gr = N_gr
                min_totalcost = totalcost
            Cost_col.append(cost_col)
            cost_sea_trans = self.priceSea * self.wthCable * self.STS
            cost_transmission = self.STS * self.price_idx[1, -1]
            cost_trench = np.sum(CTren[:])
            cost_trench_transmission = 0

            cost_sea_col = np.sum(cost_sea_trans)

            totalcost = cost_trench + cost_col + cost_transmission + cost_sea_col + cost_sea_trans
            print('totalcost:', totalcost,
                  'cost_trench', cost_trench,
                  'cost_trasmission', cost_transmission,
                  'cost_col_all', cost_col,
                  'energy loss', np.sum(Cen),
                  'cost_sea_trans', cost_sea_trans,
                  'cost_sea_col', cost_sea_col)

            monitor_file = os.path.join(self.save_path, 'monitoring', '%d.csv' % (k + 1))
            outcon = list()
            for i in range(len(Connect)):
                outcon += list(Connect[i])
            outcon = np.array(outcon)
            cable_name = [self.cablespec["name"][self.cablespec["cross_section"] == s].values[0] for s in
                          np.array(outcon)[:, -1]]
            voltage = [self.cablespec["voltage"][self.cablespec["cross_section"] == s].values[0] for s in
                       np.array(outcon)[:, -1]]
            no_in = [self.layout["no"][self.layout["index"] == s].values[0] for s in np.array(outcon)[:, 0]]
            x_in = [self.layout["X"][self.layout["index"] == s].values[0] * 1000 for s in np.array(outcon)[:, 0]]
            y_in = [self.layout["Y"][self.layout["index"] == s].values[0] * 1000 for s in np.array(outcon)[:, 0]]
            no_out = [self.layout["no"][self.layout["index"] == s].values[0] for s in np.array(outcon)[:, 1]]
            x_out = [self.layout["X"][self.layout["index"] == s].values[0] * 1000 for s in np.array(outcon)[:, 1]]
            y_out = [self.layout["Y"][self.layout["index"] == s].values[0] * 1000 for s in np.array(outcon)[:, 1]]
            monitor_df = pd.DataFrame({"no_in": no_in, "x_in": x_in, "y_in": y_in,
                                       "no_out": no_out, "x_out": x_out, "y_out": y_out,
                                       "cable_type": cable_name, "voltage": voltage,
                                       "cable_length": [l for l in outcon[:, 2]],
                                       "cable_wt": [l for l in outcon[:, 3]]})

            # outcon2 = [list(o) for o in outcon]
            # for l in range(len(outcon)):
            # outcon2[l] = [no_in[l], x_in[l], y_in[l], no_out[l], x_out[l], y_out[l], cable_name[l],
            #              voltage[l]] + list(outcon[l][2:4])
            # outcon2[l].append(voltage[l])
            monitor_df.to_csv(monitor_file)
            with open(monitor_file, 'a') as f:
                f.write('Total cost,' + str(totalcost / 10000) + '\n')
                f.write('Collection cable cost,' + str(cost_col / 10000) + '\n')
                f.write('Transmission cable cost,' + str(cost_transmission / 10000) + '\n')
                f.write('Trench cost of collection system,' + str(cost_trench / 10000) + '\n')
                f.write('Trench cost of transmission system,' + str(cost_trench_transmission / 10000) + '\n')
                f.write('Energy loss cost in operation period,' + str(np.sum(Cen) / 10000) + '\n')
                f.write('Sea cost of collection system,' + str(cost_sea_col / 10000) + '\n')
                f.write('Sea cost of transmission system,' + str(cost_sea_trans / 10000) + '\n')
                # f.write
            # self.plot_func(Connect, N_gr, k, cost_col, totalcost)

            if len(self.cost_hist) == 0 or self.cost_hist[-1][1] > cost_col:
                self.structure_hist.append(monitor_df)
                self.cost_hist.append(
                    [totalcost,  # 线缆总成本TC
                     cost_col,  # 集电线缆成本CSC
                     cost_transmission,  # 输电线缆成本TSC
                     cost_trench,  # 集电线缆安装成本CSIC
                     cost_trench_transmission,  # 输电线缆安装成本TSIC，目前置0
                     sum(Cen),  # 损耗成本 energy loss
                     cost_sea_col,  # 集电占海SAC_CS
                     cost_sea_trans])  # 输电占海SAC_TS
                self.index.append(k)

                print(k + 1, totalcost)
                with open(evol_hist_file, 'a') as f:
                    f.write('%d %.2f\n' % (k + 1, totalcost / 10000))
        self.plot_func(min_connect, min_N_gr, min_k, min_cost_col, min_totalcost)

    def plot_func(self, connec, N_gr, k, cost_col, totalcost):
        # layout plotting
        # plot cable connectionlayout
        # figure = plt.figure()
        Gn = np.array(self.Gn)
        plt.scatter(Gn[:, 0], Gn[:, 1])
        # for i in range(len(self.Gn)):
        #     plt.annotate(self.layout["no"][self.layout["index"] == (i + 1)].values[0], xy=(Gn[i, 0], Gn[i, 1]))
        for i in range(N_gr):
            con = connec[i]
            for j in range(np.size(connec[i], 0)):
                # Gn = np.array(self.Gn)

                x = [Gn[con[j, 0].astype(int) - 1, 0], Gn[con[j, 1].astype(int) - 1, 0]]
                y = [Gn[con[j, 0].astype(int) - 1, 1], Gn[con[j, 1].astype(int) - 1, 1]]
                plt.plot(x, y)
        # if  cost_col_min<cost_col:
        # cost_col_min=cost_col_min
        # else
        # cost_col_min=cost_col:

        plt.xlabel('x [m]')
        plt.ylabel('y [m]')
        plt.title('Cost Collection=%.2fRMB' % (cost_col))
        plt.savefig(os.path.join(self.save_path, 'monitoring', 'Cable routing %d.png' % (k + 1)))
        plt.draw()
        # plt.pause(1)
        # plt.close()
        plt.show()
        # plt.close()

    # plt.xlabel('x [m]')
    # plt.ylabel('y [m]')
    # plt.title('Cost Collection=%.2fRMB' % (cost_col_min))
    # plt.savefig(os.path.join(self.save_path, 'monitoring', 'Cable routing %d.png' % (k + 1)))
    # plt.draw()

    def output(self):
        best_cost = self.cost_hist[-1]
        best_cost = np.array(best_cost) / 10000
        best_structure = self.structure_hist[-1]
        calc_time = time.time() - self.start
        step = min(self.Nt, self.MST_iter) + self.PSO_iter * self.PSO_swarm
        cost_file = os.path.join(self.save_path, 'cost.json')
        cost_dict = {
            "CSC": best_cost[1],
            "TSC": best_cost[2],
            "TC": (best_cost[1] + best_cost[2]),
            "CSIC": best_cost[3],
            "TSIC": best_cost[4],
            "IC": (best_cost[3] + best_cost[4]),
            "SAC_CS": best_cost[6],
            "SAC_TS": best_cost[7],
            "SAC": (best_cost[6] + best_cost[7])
        }
        cost_dict["OC"] = cost_dict["TC"] + cost_dict["IC"] + cost_dict["SAC"] + best_cost[5]
        cost_dict["TL"] = self.STS
        cost_dict["CL"] = np.sum(best_structure["cable_length"].values)
        cost_dict["OL"] = cost_dict['TL'] + cost_dict["CL"]
        with open(cost_file, 'w') as f:
            json.dump(cost_dict, f)
        information_file = os.path.join(self.save_path, 'information.json')
        # print(information_file)
        information = {"calc_time": calc_time,
                       "step": step,
                       "CSC": cost_dict["CSC"],
                       "CL": cost_dict["CL"],
                       "TL": cost_dict["TL"],
                       "OL": cost_dict["OL"],
                       "OC": cost_dict["OC"]}
        with open(information_file, 'w') as f:
            json.dump(information, f)
        structure_file = os.path.join(self.save_path, 'connection.csv')
        best_structure.to_csv(structure_file, index=None)
        summary_file = os.path.join(self.save_path, 'summary.txt')
        with open(summary_file, 'w') as f:
            f.write('Calculation time: %.2f\n' % calc_time)
            f.write('Total step: %d\n' % step)
            for k, v in cost_dict.items():
                f.write('%s of best routing: %.2f\n' % (k, v))
        layout_file = os.path.join(self.save_path, 'best layout.csv')
        self.layout["X"] = self.layout["X"] * 1000
        self.layout["Y"] = self.layout["Y"] * 1000
        self.layout.loc[:, ["no", "X", "Y"]].to_csv(layout_file, index=None)
        source_figure = os.path.join(self.save_path, 'monitoring', 'Cable routing %d.png' % (self.index[-1] + 1))
        destination_figure = os.path.join(self.save_path, 'Best structure.png')
        copyfile(source_figure, destination_figure)


def Dataload(cablespec, OS, WT, infoWF):
    Prated = infoWF[0, 1]  # [MW],          Rated Power of Wind Turbines
    Nt = int(infoWF[1, 1])  # [pcs],         Number of WTs in the WF
    Ny = int(infoWF[2, 1])  # [year],        Lifetime of the WF operation for energy losses
    Nh = int(infoWF[3, 1])  # [hour/year],   Utilization factor of WF’s installed capacity
    Ntin = int(infoWF[4, 1])  # [pcs],         Maximum number of WT's cables
    d_f = infoWF[5, 1]  # [],            Discount rate
    Ce = infoWF[6, 1]  # [rmb/MWh],     Electricity price (China 2015, 388.25rmb/MWh)
    U_col = int(infoWF[7, 1])  # [kV],          Voltage of the collection nets
    U_trans = int(infoWF[8, 1])  # [kV],          Voltage of the transmission nets
    PF = infoWF[9, 1]  # [],            Power factor
    fc = infoWF[10, 1]  # [],            Cable current carrying capacity correction factor
    k_length = infoWF[11, 1]  # [],            Ratio of cable length and distance
    priceClnIntl = infoWF[12, 1]  # [万元/km],      Price of Transmission cable installation
    Ct = infoWF[13, 1]  # [万元/km],      Price of Collection cable installation
    priceSea = infoWF[14, 1]  # [RMB/mm2]      Price of the sea
    wthCable = infoWF[15, 1]  # [m]            Width of Cable
    N_A = infoWF[16, 1]  # []             Current correction factor

    In = Prated / np.sqrt(3) / U_col / PF * fc  # [kA], nominal current of single turbine
    # % ----- Cables parameters -----
    # print(cablespec.head())
    """Nb_max = math.floor(max(cablespec[0: -1, 3] / 1e3 / In))  # maximum allowable per unit power, 14 for case for paper
    Cs = cablespec[:, 3] / 1e3 <= In * (Nb_max + 1)  # cable types that
    Ic = cablespec[:, 3][Cs] / 1e3  # [kA], cable current carrying capacity at different cross-section
    Ncs = np.size(Ic)"""
    Nb_max = 1
    Cs = 1
    Ic = 1

    # %% inner array cable model with cost
    price_idx = []
    """price_idx = np.zeros([3, Ncs + 1])
    price_idx[0, :-1] = [math.floor(x) for x in cablespec[:, 3][Cs] / In / 1e3]
    # math.floor(cablespec[:,3][Cs]/In/1e3)
    # [print(a) for a in range(1,6)]
    price_idx[0, -1] = Nt + 1
    price_idx[1, :] = cablespec[:, 0].T  # cable cost [Rmb/km]
    price_idx[2, :] = cablespec[:, 4].T"""

    # %% offshore substation coordinate
    # Sub_x = OS[0]
    # Sub_y = OS[1]

    # manual design onshore substation
    # sub_on_x = OS[0] - 14
    # sub_on_y = OS[1] - 20

    # %% main programme
    # STS = np.sqrt(pow((sub_on_x - Sub_x), 2) + pow((sub_on_y - Sub_y), 2))
    STS = 0
    # %% scan algorithm for partition zones
    return [Prated, Nt, Nh, U_col, PF, fc, Ce, Ct, d_f, Ny, price_idx, Nb_max, STS]


class MyPanel1(wx.Panel):
    def __init__(self, parent):
        super(MyPanel1, self).__init__(parent)
        # 第一行
        m_static_nwt = wx.StaticText(self, wx.ID_ANY, u"风机数量", pos=(FIRST_COL_LABEL, LABEL_TEXT_HEIGHT),
                                     size=(60, 20))
        self.m_text_nwt = wx.TextCtrl(self, wx.ID_ANY, u"45", pos=(FIRST_COL_TEXT, LABEL_TEXT_HEIGHT),
                                      size=(TEXT_WIDTH, 20), style=wx.TE_PROCESS_ENTER)
        self.m_text_nwt.Bind(wx.EVT_TEXT, self.EndTextK23)

        m_static_rated_power = wx.StaticText(self, wx.ID_ANY, u"单个风机容量（MW）",
                                             pos=(SECOND_COL_LABEL, LABEL_TEXT_HEIGHT),
                                             size=(120, 20))
        self.text_rated_power = wx.TextCtrl(self, wx.ID_ANY, u"4", pos=(SECOND_COL_TEXT+40, LABEL_TEXT_HEIGHT),
                                            size=(TEXT_WIDTH-40, 20), style=wx.TE_PROCESS_ENTER)
        self.text_rated_power.Bind(wx.EVT_TEXT, self.EndTextK20)

        m_open_year = wx.StaticText(self, wx.ID_ANY, u"预计运行年限", pos=(THIRD_COL_LABEL, LABEL_TEXT_HEIGHT),
                                    size=(80, 20))
        self.m_text_open_year = wx.TextCtrl(self, wx.ID_ANY, u"25", pos=(THIRD_COL_TEXT+20, LABEL_TEXT_HEIGHT),
                                            size=(TEXT_WIDTH-20, LABEL_TEXT_HEIGHT), style=wx.TE_PROCESS_ENTER)
        self.m_text_open_year.Bind(wx.EVT_TEXT, self.EndTextC3)

        m_wt_max_cable = wx.StaticText(self, wx.ID_ANY, u"机组最多输出电缆数",
                                       pos=(FIRST_COL_LABEL, LABEL_TEXT_HEIGHT + 55),
                                       size=(LABEL_WIDTH + 50, LABEL_TEXT_HEIGHT))
        self.m_text_wt_max_cable = wx.TextCtrl(self, wx.ID_ANY, u"2", pos=(FIRST_COL_TEXT + 50, LABEL_TEXT_HEIGHT + 55),
                                               size=(TEXT_WIDTH - 40, LABEL_TEXT_HEIGHT), style=wx.TE_PROCESS_ENTER)
        self.m_text_wt_max_cable.Bind(wx.EVT_TEXT, self.EndTextB8)

        m_u_col = wx.StaticText(self, wx.ID_ANY, u"额定电压（kV）", pos=(SECOND_COL_LABEL, LABEL_TEXT_HEIGHT + 55),
                                size=(LABEL_WIDTH+20, LABEL_TEXT_HEIGHT))
        self.m_text_u_col = wx.TextCtrl(self, wx.ID_ANY, u"35", pos=(SECOND_COL_TEXT+20, LABEL_TEXT_HEIGHT + 55),
                                        size=(TEXT_WIDTH-20, LABEL_TEXT_HEIGHT), style=wx.TE_PROCESS_ENTER)
        self.m_text_u_col.Bind(wx.EVT_TEXT, self.EndTextK8)

        m_u_trans = wx.StaticText(self, wx.ID_ANY, u"并网电压（kV）", pos=(THIRD_COL_LABEL, LABEL_TEXT_HEIGHT + 55),
                                  size=(LABEL_WIDTH+20, LABEL_TEXT_HEIGHT))
        self.m_text_u_trans = wx.TextCtrl(self, wx.ID_ANY, u"220", pos=(THIRD_COL_TEXT+20, LABEL_TEXT_HEIGHT + 55),
                                          size=(TEXT_WIDTH-20, LABEL_TEXT_HEIGHT), style=wx.TE_PROCESS_ENTER)
        self.m_text_u_trans.Bind(wx.EVT_TEXT, self.EndTextO17)

        m_static_PF = wx.StaticText(self, wx.ID_ANY, u"并网点功率因数", pos=(FIRST_COL_LABEL, LABEL_TEXT_HEIGHT + 110),
                                    size=(LABEL_WIDTH + 20, LABEL_TEXT_HEIGHT))
        self.text_pf = wx.TextCtrl(self, wx.ID_ANY, u"0.9", pos=(FIRST_COL_TEXT + 20, LABEL_TEXT_HEIGHT + 110),
                                   size=(TEXT_WIDTH, 20), style=wx.TE_PROCESS_ENTER)
        self.text_pf.Bind(wx.EVT_TEXT, self.EndTextB28)

        m_staticTextI10 = wx.StaticText(self, wx.ID_ANY, u"集电系统安装费用（万元/km）",
                                        pos=(SECOND_COL_LABEL, LABEL_TEXT_HEIGHT + 110),
                                        size=(LABEL_WIDTH + 100, LABEL_TEXT_HEIGHT))
        self.price_cin_intl = wx.TextCtrl(self, wx.ID_ANY, u"30", pos=(SECOND_COL_TEXT + 80, LABEL_TEXT_HEIGHT + 110),
                                          size=(TEXT_WIDTH - 20, 20), style=wx.TE_PROCESS_ENTER)
        self.price_cin_intl.Bind(wx.EVT_TEXT, self.EndTextD10)

        m_staticTextJ10 = wx.StaticText(self, wx.ID_ANY, u"输电系统安装费用（万元/km）",
                                        pos=(THIRD_COL_LABEL + 70, LABEL_TEXT_HEIGHT + 110),
                                        size=(LABEL_WIDTH + 100, LABEL_TEXT_HEIGHT))
        self.price_trn_intl = wx.TextCtrl(self, wx.ID_ANY, u"138", pos=(THIRD_COL_TEXT + 160, LABEL_TEXT_HEIGHT + 110),
                                          size=(TEXT_WIDTH, 20), style=wx.TE_PROCESS_ENTER)
        self.price_trn_intl.Bind(wx.EVT_TEXT, self.EndTextC10)

        m_staticText36 = wx.StaticText(self, wx.ID_ANY, u"海域占用费用（元/mm2）",
                                       pos=(FIRST_COL_LABEL, LABEL_TEXT_HEIGHT + 165),
                                       size=(LABEL_WIDTH + 80, LABEL_TEXT_HEIGHT))
        self.m_price_sea = wx.TextCtrl(self, wx.ID_ANY, u"0.7", pos=(FIRST_COL_TEXT + 80, LABEL_TEXT_HEIGHT + 165),
                                       size=(TEXT_WIDTH, LABEL_TEXT_HEIGHT), style=wx.TE_PROCESS_ENTER)
        self.m_price_sea.Bind(wx.EVT_TEXT, self.EndTextD14)

        m_staticText37 = wx.StaticText(self, wx.ID_ANY, u"海域宽度（m）",
                                       pos=(SECOND_COL_LABEL + 80, LABEL_TEXT_HEIGHT + 165),
                                       size=(LABEL_WIDTH + 20, LABEL_TEXT_HEIGHT))
        self.m_wth_cable = wx.TextCtrl(self, wx.ID_ANY, u"20", pos=(SECOND_COL_TEXT + 90, LABEL_TEXT_HEIGHT + 165),
                                       size=(TEXT_WIDTH - 10, LABEL_TEXT_HEIGHT), style=wx.TE_PROCESS_ENTER)
        self.m_wth_cable.Bind(wx.EVT_TEXT, self.EndTextC14)

        button = wx.Button(self, wx.ID_ANY, '计算结果', pos=(FIRST_COL_LABEL, LABEL_TEXT_HEIGHT + 230),
                           size=(TEXT_WIDTH, LABEL_TEXT_HEIGHT+20))
        button.Bind(wx.EVT_BUTTON, CableOptData)

    def __del__(self):
        pass

    def EndTextK23(self, event):
        global input_n_wt
        input_n_wt = self.m_text_nwt.GetValue()
        event.Skip()

    def EndTextK20(self, event):
        global input_rated_power
        input_rated_power = self.text_rated_power.GetValue()
        event.Skip()

    def EndTextC3(self, event):
        global input_open_year
        input_open_year = self.m_text_open_year.GetValue()
        event.Skip()

    def EndTextB8(self, event):
        global input_wt_max_cable
        input_wt_max_cable = self.m_text_wt_max_cable.GetValue()
        event.Skip()

    def EndTextK8(self, event):
        global input_u_col
        input_u_col = self.m_text_u_col.GetValue()
        event.Skip()

    def EndTextO17(self, event):
        global input_u_trans
        input_u_trans = self.m_text_u_trans.GetValue()
        event.Skip()

    def EndTextB28(self, event):
        global input_pf
        input_pf = self.text_pf.GetValue()
        event.Skip()

    def EndTextD10(self, event):
        global input_price_cin_intl
        input_price_cin_intl = self.price_cin_intl.GetValue()
        event.Skip()

    def EndTextC10(self, event):
        global input_price_trn_intl
        input_price_trn_intl = self.price_trn_intl.GetValue()
        event.Skip()

    def EndTextD14(self, event):
        global input_price_sea
        input_price_sea = self.m_price_sea.GetValue()
        event.Skip()

    def EndTextC14(self, event):
        global input_with_cable
        input_with_cable = self.m_wth_cable.GetValue()
        event.Skip()


app = wx.App()
frame = wx.Frame(None, title="电缆程序", pos=(200, 200), size=(840, 400))
nb = wx.Notebook(frame)
p1 = MyPanel1(nb)

nb.AddPage(p1, "电缆计算")
frame.Show()
app.MainLoop()
